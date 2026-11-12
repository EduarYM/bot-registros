"""
Bot de Telegram · Notas de egreso -> Excel con reparto · Subcentro de Vinces

Mandas las fotos de las notas de egreso y el bot:
  1. Las lee con Google Gemini (gratis): producto, lote, caducidad y cantidad.
  2. Las guarda en UN solo Excel, con una hoja por día.
  3. Reparte cada cantidad entre los lugares de distribución.
  4. Te manda el Excel y lo deja fijado en el chat.

Variables de entorno (Render > Environment, NUNCA dentro del código):
  TELEGRAM_TOKEN       Token de @BotFather                                 (obligatorio)
  GEMINI_API_KEY       Clave de Google AI Studio                           (obligatorio)
  USUARIOS_PERMITIDOS  IDs de Telegram permitidos, con comas                (opcional)
  LUGARES              Lugares de reparto; con pesos: Urbano:3, Palenque:1  (opcional)
  ESPERA_EXCEL         Segundos de espera antes de mandar el Excel          (opcional, 15)
  FOTOS_A_LA_VEZ       Fotos que se leen en paralelo                        (opcional, 3)
  NOMBRE_EXCEL         Nombre del archivo                                   (opcional)
  GEMINI_MODEL         Modelo de Gemini                                     (opcional)
  ZONA_HORARIA         Por defecto America/Guayaquil
"""

import asyncio
import base64
import io
import json
import logging
import os
import re
import secrets
from datetime import date, datetime
from zoneinfo import ZoneInfo

import httpx
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# ============================================================================
# Configuración
# ============================================================================
logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("bot")

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
USUARIOS_PERMITIDOS = {u.strip() for u in os.environ.get("USUARIOS_PERMITIDOS", "").split(",") if u.strip()}
LUGARES_TXT = os.environ.get(
    "LUGARES", "Urbano, Palenque, Nicaragua, Antonio, Balzar, Juaneche, Bagatela, Clariza, Matecito"
)
ESPERA_EXCEL = int(os.environ.get("ESPERA_EXCEL", "15"))
FOTOS_A_LA_VEZ = max(1, int(os.environ.get("FOTOS_A_LA_VEZ", "3")))
NOMBRE_EXCEL = os.environ.get("NOMBRE_EXCEL", "Egresos_Subcentro_Vinces.xlsx").strip()
if not NOMBRE_EXCEL.lower().endswith(".xlsx"):
    NOMBRE_EXCEL += ".xlsx"
ZONA = ZoneInfo(os.environ.get("ZONA_HORARIA", "America/Guayaquil"))
MODELOS = [
    m
    for m in dict.fromkeys(
        [os.environ.get("GEMINI_MODEL", "").strip(), "gemini-flash-latest", "gemini-2.5-flash", "gemini-flash-lite-latest"]
    )
    if m
]
GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta"

URL_PUBLICA = os.environ.get("WEBHOOK_URL") or os.environ.get("RENDER_EXTERNAL_URL")
PUERTO = int(os.environ.get("PORT", "10000"))


def leer_lugares(texto: str) -> list[tuple[str, float]]:
    lugares = []
    for parte in texto.split(","):
        parte = parte.strip()
        if not parte:
            continue
        if ":" in parte:
            nombre, peso = parte.rsplit(":", 1)
            try:
                lugares.append((nombre.strip(), float(peso.replace(",", "."))))
                continue
            except ValueError:
                pass
        lugares.append((parte, 1.0))
    return lugares


LUGARES = leer_lugares(LUGARES_TXT)
NOMBRES_LUGARES = [n for n, _ in LUGARES]

COL_FIJAS = ["N°", "Producto", "Lote", "Caducidad", "Cantidad"]
COL_FINALES = ["Total repartido", "Imagen"]
ENCABEZADOS = COL_FIJAS + NOMBRES_LUGARES + COL_FINALES

# ============================================================================
# Estado en memoria por chat
# ============================================================================
DATOS: dict[int, dict] = {}


def nuevo_lote() -> dict:
    return {"total": 0, "listas": 0, "productos": 0, "fallidas": [], "aviso": "", "msg": None, "dias": []}


def datos_chat(chat_id: int) -> dict:
    if chat_id not in DATOS:
        DATOS[chat_id] = {
            "pendientes": [],      # productos leídos que aún no están en el Excel
            "procesando": 0,       # fotos en lectura
            "version": 0,
            "enviar_ya": False,    # el usuario pidió /excel mientras se leía
            "dia": None,           # día elegido con /fecha (None = hoy)
            "base": None,          # bytes del último Excel
            "lote": nuevo_lote(),  # mensaje de estado del grupo de fotos actual
            "excel_lock": asyncio.Lock(),
            "estado_lock": asyncio.Lock(),
            "semaforo": asyncio.Semaphore(FOTOS_A_LA_VEZ),
        }
    return DATOS[chat_id]


def permitido(update: Update) -> bool:
    return not USUARIOS_PERMITIDOS or str(update.effective_user.id) in USUARIOS_PERMITIDOS


def hoy() -> date:
    return datetime.now(ZONA).date()


DIAS_SEMANA = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]


def dia_activo(d: dict) -> date:
    return d["dia"] or hoy()


def corto(dia: date) -> str:
    return f"{DIAS_SEMANA[dia.weekday()]} {dia:%d/%m}"


def leer_fecha_texto(texto: str) -> date | None:
    """Entiende: hoy, ayer, anteayer, 10/11, 10-11-2026, 10/11/26."""
    t = (texto or "").strip().lower()
    base = hoy()
    if t in ("hoy",):
        return base
    if t in ("ayer",):
        return base.fromordinal(base.toordinal() - 1)
    if t in ("anteayer", "antier"):
        return base.fromordinal(base.toordinal() - 2)
    m = re.search(r"\b(\d{1,2})[/\-.](\d{1,2})(?:[/\-.](\d{2,4}))?\b", t)
    if not m:
        return None
    dd, mm, aa = int(m.group(1)), int(m.group(2)), m.group(3)
    anio = int(aa) + (2000 if aa and len(aa) == 2 else 0) if aa else base.year
    try:
        f = date(anio, mm, dd)
    except ValueError:
        return None
    if not aa and f > base:  # "28/12" escrito en enero = año pasado
        f = date(anio - 1, mm, dd)
    return f


# ============================================================================
# Gemini (API REST; acepta claves "AQ." y "AIza")
# ============================================================================
PROMPT = """Imagen: NOTA DE EGRESO de medicamentos e insumos médicos (columnas N°, PRODUCTO,
LOTE, CADUCIDAD y una de CANTIDAD cuyo encabezado puede ser un código como "12OT05").
Transcribe TODAS las filas de productos. Responde SOLO este JSON:
{"productos":[{"producto":"...","lote":"...","caducidad":"AAAA-MM-DD","cantidad":0}]}
- producto: nombre completo tal cual (concentración, forma, presentación).
- lote: exacto (ceros al inicio, puntos, guiones, letras).
- cantidad: entero sin separadores de miles.
- Si algo no se lee claro, tu mejor lectura + "(?)". No inventes. Sin encabezados ni totales.
- Si no hay productos: {"productos":[]}"""

_modo_auth = {"actual": None}
MODOS_AUTH = ["header", "query", "bearer"]
MODO_SIMPLE: set[str] = set()  # modelos que no aceptan las opciones extra (esquema / razonamiento)
_cliente: httpx.AsyncClient | None = None


def cliente() -> httpx.AsyncClient:
    global _cliente
    if _cliente is None:
        _cliente = httpx.AsyncClient(timeout=httpx.Timeout(150, connect=15))
    return _cliente


def _auth(modo: str) -> tuple[dict, dict]:
    if modo == "header":
        return {"x-goog-api-key": GEMINI_API_KEY}, {}
    if modo == "query":
        return {}, {"key": GEMINI_API_KEY}
    return {"Authorization": f"Bearer {GEMINI_API_KEY}"}, {}


def _es_error_auth(r: httpx.Response) -> bool:
    return r.status_code in (400, 401, 403) and any(
        t in r.text for t in ("API_KEY", "API key", "UNAUTHENTICATED", "ACCESS_TOKEN", "credential", "PERMISSION_DENIED")
    )


async def comprobar_gemini() -> str:
    """Al arrancar: detecta cómo enviar la clave (se ve en los Logs de Render)."""
    ultimo = ""
    for modo in MODOS_AUTH:
        h, q = _auth(modo)
        try:
            r = await cliente().get(f"{GEMINI_URL}/models", headers=h, params={**q, "pageSize": 50})
        except httpx.HTTPError as e:
            return f"sin conexión con Gemini ({e})"
        if r.status_code == 200:
            _modo_auth["actual"] = modo
            return f"OK (modo {modo})"
        ultimo = f"{r.status_code}: {r.text[:200]}"
    return f"la clave NO funcionó -> {ultimo}"


ESQUEMA = {
    "type": "OBJECT",
    "properties": {
        "productos": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "producto": {"type": "STRING"},
                    "lote": {"type": "STRING"},
                    "caducidad": {"type": "STRING"},
                    "cantidad": {"type": "INTEGER"},
                },
                "required": ["producto", "lote", "caducidad", "cantidad"],
            },
        }
    },
    "required": ["productos"],
}


def _config(modelo: str) -> dict:
    """JSON con formato fijo y razonamiento bajo: más rápido y sin respuestas raras."""
    cfg = {"responseMimeType": "application/json", "temperature": 0}
    if modelo not in MODO_SIMPLE:
        cfg["responseSchema"] = ESQUEMA
        if "2.5" in modelo:
            cfg["thinkingConfig"] = {"thinkingBudget": 0}
        else:
            cfg["thinkingConfig"] = {"thinkingLevel": "low"}
    return cfg


async def llamar_gemini(imagen: bytes, mime: str, avisar=None) -> dict:
    partes = [
        {"inline_data": {"mime_type": mime, "data": base64.b64encode(imagen).decode()}},
        {"text": PROMPT},
    ]
    loop = asyncio.get_running_loop()
    limite = loop.time() + 280
    ultimo_error = "sin respuesta"
    modos = [_modo_auth["actual"]] if _modo_auth["actual"] else MODOS_AUTH

    for modelo in MODELOS:
        intento = 0
        while intento < 3 and loop.time() < limite:
            intento += 1
            cuerpo = {"contents": [{"parts": partes}], "generationConfig": _config(modelo)}
            r = None
            for modo in modos:
                h, q = _auth(modo)
                t0 = loop.time()
                try:
                    r = await cliente().post(
                        f"{GEMINI_URL}/models/{modelo}:generateContent", headers=h, params=q, json=cuerpo
                    )
                except httpx.TimeoutException:
                    ultimo_error = "Gemini tardó demasiado"
                    log.warning("Gemini %s: timeout tras %.0fs", modelo, loop.time() - t0)
                    r = "timeout"
                    break
                except httpx.HTTPError as e:
                    ultimo_error = f"conexión: {type(e).__name__}"
                    r = None
                    break
                log.info("Gemini %s -> %s en %.1fs", modelo, r.status_code, loop.time() - t0)
                if _es_error_auth(r) and not _modo_auth["actual"]:
                    ultimo_error = f"{r.status_code} {r.text[:200]}"
                    r = None
                    continue
                if r.status_code == 200 and not _modo_auth["actual"]:
                    _modo_auth["actual"] = modo
                break
            if r == "timeout":
                break  # modelo lento: pruebo el siguiente
            if r is None:
                await asyncio.sleep(2)
                continue

            if r.status_code == 200:
                try:
                    cand = r.json()["candidates"][0]
                    p = cand["content"]["parts"]
                    return interpretar_json("".join(x.get("text", "") for x in p if not x.get("thought")))
                except (KeyError, IndexError, ValueError) as e:
                    ultimo_error = f"respuesta incompleta ({type(e).__name__})"
                    log.warning("Gemini %s respuesta rara: %s", modelo, r.text[:500])
                    continue

            ultimo_error = f"{r.status_code} {r.text[:300]}"
            log.warning("Gemini %s: %s", modelo, ultimo_error)
            if r.status_code == 400 and modelo not in MODO_SIMPLE:
                MODO_SIMPLE.add(modelo)  # no acepta las opciones extra: reintento sin ellas
                intento -= 1
                continue
            if r.status_code == 429:
                if "limit: 0" in r.text or "PerDay" in r.text or intento >= 2:
                    break  # sin cupo en este modelo: pruebo el siguiente
                if avisar:
                    await avisar("Gemini con mucha demanda, reintentando…")
                await asyncio.sleep(10)
                continue
            if r.status_code in (500, 502, 503, 504):
                if avisar:
                    await avisar("Gemini ocupado, reintentando…")
                await asyncio.sleep(5)
                continue
            break  # otro error: siguiente modelo
    raise RuntimeError(ultimo_error)


def interpretar_json(texto: str) -> dict:
    texto = re.sub(r"^```(?:json)?\s*|\s*```$", "", texto.strip())
    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        i, f = texto.find("{"), texto.rfind("}")
        if i != -1 and f > i:
            return json.loads(texto[i : f + 1])
        raise


# ============================================================================
# Limpieza de datos y reparto
# ============================================================================
def a_fecha(valor):
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    s = str(valor or "").strip()
    for patron, orden in (
        (r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", "amd"),
        (r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$", "dma"),
    ):
        m = re.match(patron, s)
        if m:
            a, b, c = (int(x) for x in m.groups())
            try:
                return date(a, b, c) if orden == "amd" else date(c, b, a)
            except ValueError:
                return s
    return s


def a_entero(valor):
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)):
        return int(round(valor))
    digitos = re.sub(r"[^\d]", "", str(valor))
    return int(digitos) if digitos else None


def limpiar_productos(resultado: dict) -> list[dict]:
    salida = []
    if isinstance(resultado, list):
        lista = resultado
    elif isinstance(resultado, dict):
        lista = resultado.get("productos")
        if lista is None:  # otra clave: toma la primera lista que encuentre
            lista = next((v for v in resultado.values() if isinstance(v, list)), [])
    else:
        lista = []
    for p in lista or []:
        if not isinstance(p, dict):
            continue
        producto = str(p.get("producto") or "").strip()
        if not producto or producto.lower() in {"producto", "total"}:
            continue
        salida.append(
            {
                "producto": producto,
                "lote": str(p.get("lote") or "").strip(),
                "caducidad": a_fecha(p.get("caducidad")),
                "cantidad": a_entero(p.get("cantidad")),
            }
        )
    return salida


def repartir(cantidad, lugares=LUGARES) -> list:
    """Reparte en enteros según los pesos (método del mayor residuo).
    Lo que sobra se da en el orden de la lista (Urbano primero)."""
    if cantidad is None:
        return [None] * len(lugares)
    total_peso = sum(p for _, p in lugares) or 1
    exactos = [cantidad * p / total_peso for _, p in lugares]
    partes = [int(x) for x in exactos]
    faltan = cantidad - sum(partes)
    orden = sorted(range(len(lugares)), key=lambda i: (-(exactos[i] - partes[i]), i))
    for i in orden[:faltan]:
        partes[i] += 1
    return partes


# ============================================================================
# Excel
# ============================================================================
AZUL = PatternFill("solid", fgColor="1F4E78")
GRIS = PatternFill("solid", fgColor="D9E1F2")
ROJO = PatternFill("solid", fgColor="F8CBAD")
LINEA = Side(style="thin", color="A6A6A6")
BORDE = Border(left=LINEA, right=LINEA, top=LINEA, bottom=LINEA)


def nombre_hoja(d: date) -> str:
    return d.strftime("%d-%m-%Y")


def leer_hoja(ws) -> list[dict]:
    """Lee los productos de una hoja creada por el bot (para no perderlos al agregar más)."""
    fila_enc = None
    for r in range(1, 8):
        valores = [str(ws.cell(r, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)]
        if "producto" in valores:
            fila_enc = r
            mapa = {v: i + 1 for i, v in enumerate(valores) if v}
            break
    if not fila_enc:
        return []
    filas = []
    for r in range(fila_enc + 1, ws.max_row + 1):
        producto = ws.cell(r, mapa["producto"]).value
        if producto is None or str(producto).strip().upper() == "TOTAL":
            continue

        def val(nombre):
            c = mapa.get(nombre.lower())
            return ws.cell(r, c).value if c else None

        cantidad = a_entero(val("Cantidad"))
        reparto = []
        for lugar in NOMBRES_LUGARES:
            v = val(lugar)
            reparto.append(a_entero(v) if v is not None and not str(v).startswith("=") else None)
        if all(v is None for v in reparto):
            reparto = repartir(cantidad)
        filas.append(
            {
                "producto": str(producto).strip(),
                "lote": str(val("Lote") or ""),
                "caducidad": a_fecha(val("Caducidad")),
                "cantidad": cantidad,
                "reparto": reparto,
                "imagen": val("Imagen") or "",
            }
        )
    return filas


def escribir_hoja(ws, dia: date, filas: list[dict]) -> None:
    n_cols = len(ENCABEZADOS)
    ultima = get_column_letter(n_cols)
    c_cant = COL_FIJAS.index("Cantidad") + 1
    c_l1 = len(COL_FIJAS) + 1
    c_ln = c_l1 + len(NOMBRES_LUGARES) - 1
    c_tot = c_ln + 1

    # Encabezados (fila 1)
    for c, nombre in enumerate(ENCABEZADOS, start=1):
        celda = ws.cell(1, c, nombre)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = AZUL
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = BORDE
    ws.row_dimensions[1].height = 30

    # Datos
    r = 2
    for n, f in enumerate(filas, start=1):
        valores = [n, f["producto"], f["lote"], f["caducidad"], f["cantidad"]] + list(f["reparto"])
        for c, v in enumerate(valores, start=1):
            ws.cell(r, c, v)
        ws.cell(r, c_tot, f"=SUM({get_column_letter(c_l1)}{r}:{get_column_letter(c_ln)}{r})")
        ws.cell(r, c_tot + 1, f["imagen"])
        for c in range(1, n_cols + 1):
            celda = ws.cell(r, c)
            celda.border = BORDE
            celda.alignment = Alignment(
                vertical="center", wrap_text=(c == 2),
                horizontal="left" if c in (2, 3) else "center",
            )
        ws.cell(r, c_tot).font = Font(bold=True)
        ws.cell(r, 3).number_format = "@"
        if isinstance(f["caducidad"], date):
            ws.cell(r, 4).number_format = "DD/MM/YYYY"
        r += 1
    fin = r - 1

    # Si "Total repartido" no coincide con "Cantidad", la celda se pinta de rojo
    if fin >= 2:
        col_t, col_c = get_column_letter(c_tot), get_column_letter(c_cant)
        ws.conditional_formatting.add(
            f"{col_t}2:{col_t}{fin}", FormulaRule(formula=[f"${col_t}2<>${col_c}2"], fill=ROJO)
        )

    # Anchos y vista
    anchos = {1: 5, 2: 60, 3: 16, 4: 12, 5: 10, n_cols: 16}
    for c in range(1, n_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = anchos.get(c, max(10, len(ENCABEZADOS[c - 1]) + 2))
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{ultima}{max(fin, 1)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"


def agregar_al_excel(base: bytes | None, dia: date, nuevos: list[dict]) -> tuple[bytes, int]:
    """Agrega los productos nuevos a la hoja del día. Devuelve (excel, total_en_la_hoja)."""
    if base:
        wb = load_workbook(io.BytesIO(base))
    else:
        wb = Workbook()
        wb.remove(wb.active)

    nombre = nombre_hoja(dia)
    anteriores = []
    posicion = None
    if nombre in wb.sheetnames:
        ws_viejo = wb[nombre]
        anteriores = leer_hoja(ws_viejo)
        posicion = wb.sheetnames.index(nombre)
        wb.remove(ws_viejo)

    filas = anteriores + [dict(p, reparto=repartir(p["cantidad"])) for p in nuevos]
    ws = wb.create_sheet(nombre, posicion)
    escribir_hoja(ws, dia, filas)

    # Ordenar hojas por fecha (las que no son fechas quedan al inicio)
    def clave(h):
        try:
            return (1, datetime.strptime(h.title, "%d-%m-%Y"))
        except ValueError:
            return (0, datetime.min)

    wb._sheets.sort(key=clave)
    wb.active = wb.sheetnames.index(nombre)
    for h in wb.worksheets:
        h.sheet_view.tabSelected = h.title == nombre

    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue(), len(filas)


# ============================================================================
# Mensaje de estado (uno solo por grupo de fotos, se va actualizando)
# ============================================================================
def texto_estado(lote: dict, d: dict) -> str:
    t = lote["total"]
    if lote["listas"] < t:
        txt = f"⏳ Leyendo fotos… {lote['listas']}/{t}\n📦 {lote['productos']} productos"
        if lote["aviso"]:
            txt += f"\n{lote['aviso']}"
    else:
        txt = f"✅ {t} foto{'s' if t != 1 else ''} · {lote['productos']} productos"
        if d["pendientes"]:
            txt += "\n📊 Preparando Excel…"
    if lote["dias"]:
        txt += "\n📅 Hoja: " + ", ".join(nombre_hoja(x) for x in lote["dias"])
    if lote["fallidas"]:
        for n, motivo in lote["fallidas"]:
            txt += f"\n⚠️ Foto {n} no leída ({motivo}). Reenvíala."
    return txt


def motivo_corto(e: Exception) -> str:
    t = str(e)
    if "sin productos" in t:
        return "no encontré productos"
    if "tardó" in t:
        return "Gemini tardó demasiado"
    if t.startswith("429"):
        return "cupo de Gemini agotado, espera unos minutos"
    if t[:3].isdigit():
        return f"error {t[:3]} de Gemini"
    return t[:60] or type(e).__name__


async def mostrar_estado(bot, chat_id: int, d: dict) -> None:
    async with d["estado_lock"]:
        lote = d["lote"]
        texto = texto_estado(lote, d)
        try:
            if lote["msg"] is None:
                lote["msg"] = await bot.send_message(chat_id, texto)
            elif lote.get("ultimo") != texto:
                await lote["msg"].edit_text(texto)
            lote["ultimo"] = texto
        except Exception:  # noqa: BLE001  (p. ej. "message is not modified")
            pass


# ============================================================================
# Envío del Excel (queda fijado en el chat: ahí se guarda el historial)
# ============================================================================
async def obtener_base(bot, chat_id: int, d: dict) -> bytes | None:
    if d["base"]:
        return d["base"]
    try:
        fijado = (await bot.get_chat(chat_id)).pinned_message
        if fijado and fijado.document and (fijado.document.file_name or "").lower().endswith(".xlsx"):
            archivo = await fijado.document.get_file()
            d["base"] = bytes(await archivo.download_as_bytearray())
    except Exception:  # noqa: BLE001
        log.exception("No pude leer el Excel fijado")
    return d["base"]


async def generar_y_enviar(bot, chat_id: int) -> None:
    d = datos_chat(chat_id)
    async with d["excel_lock"]:
        if not d["pendientes"]:
            return
        nuevos = sorted(d["pendientes"], key=lambda p: p["orden"])
        excel = await obtener_base(bot, chat_id, d)
        lineas = []
        for dia in sorted({p["dia"] for p in nuevos}):
            del_dia = [p for p in nuevos if p["dia"] == dia]
            excel, total_dia = agregar_al_excel(excel, dia, del_dia)
            lineas.append(f"📅 {nombre_hoja(dia)} · +{len(del_dia)} productos ({total_dia} en el día)")
        caption = "\n".join(lineas)
        dudosos = sum("(?)" in (p["producto"] + p["lote"]) or p["cantidad"] is None for p in nuevos)
        if dudosos:
            caption += f"\n🔎 {dudosos} fila(s) marcadas con (?) para revisar"
        msg = await bot.send_document(chat_id, io.BytesIO(excel), filename=NOMBRE_EXCEL, caption=caption)
        d["base"] = excel
        d["pendientes"] = []
        d["enviar_ya"] = False
        if d["procesando"] == 0:
            lote = d["lote"]
            if lote["msg"] is not None:
                try:
                    await lote["msg"].edit_text(texto_estado(lote, d))  # quita "Preparando Excel…"
                except Exception:  # noqa: BLE001
                    pass
            d["lote"] = nuevo_lote()
        try:
            await bot.unpin_all_chat_messages(chat_id)
            await bot.pin_chat_message(chat_id, msg.message_id, disable_notification=True)
        except Exception:  # noqa: BLE001
            log.warning("No pude fijar el Excel")


async def envio_automatico(bot, chat_id: int, version: int) -> None:
    d = datos_chat(chat_id)
    if not d["enviar_ya"]:
        await asyncio.sleep(ESPERA_EXCEL)
    if d["version"] == version and d["procesando"] == 0 and d["pendientes"]:
        try:
            await generar_y_enviar(bot, chat_id)
        except Exception:  # noqa: BLE001
            log.exception("Error enviando el Excel")


# ============================================================================
# Telegram
# ============================================================================
AYUDA = (
    "📷 Mándame las fotos de las notas de egreso.\n"
    "📊 Te devuelvo el Excel: una hoja por día y el reparto por lugar.\n"
    "📌 El Excel queda fijado arriba; cada día se agrega una hoja.\n\n"
    "📅 Las fotos van al día de HOY. Para otro día, escribe la fecha "
    "(ej. «ayer» o «10/11») o usa /fecha, y luego manda las fotos.\n\n"
    "/fecha – elegir el día\n"
    "/hoy – volver a hoy\n"
    "/excel – recibir el Excel ya\n"
    "/nuevo – descartar lo leído sin guardar"
)


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(AYUDA + f"\n\nTu ID: {update.effective_user.id}")


async def cmd_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    chat_id = update.effective_chat.id
    d = datos_chat(chat_id)
    if d["procesando"]:
        d["enviar_ya"] = True
        await update.message.reply_text("👌 Te lo mando apenas termine de leer.")
        return
    if d["pendientes"]:
        d["version"] += 1
        await generar_y_enviar(context.bot, chat_id)
        return
    base = await obtener_base(context.bot, chat_id, d)
    if base:
        await update.message.reply_document(io.BytesIO(base), filename=NOMBRE_EXCEL, caption="📊 Excel actual")
    else:
        await update.message.reply_text("Aún no hay registros. Mándame las fotos 📷")


async def cmd_nuevo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    d = datos_chat(update.effective_chat.id)
    d["pendientes"], d["enviar_ya"] = [], False
    d["version"] += 1
    if d["procesando"] == 0:
        d["lote"] = nuevo_lote()
    await update.message.reply_text("🗑 Descartado. El Excel fijado no cambió.")


async def cmd_lugares(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    ejemplo = repartir(100)
    await update.message.reply_text(
        "Reparto de 100 unidades:\n" + "\n".join(f"• {n}: {v}" for n, v in zip(NOMBRES_LUGARES, ejemplo))
    )


def teclado_fechas() -> InlineKeyboardMarkup:
    base = hoy()
    dias = [base.fromordinal(base.toordinal() - i) for i in range(8)]
    botones = [
        InlineKeyboardButton(("Hoy · " if i == 0 else "Ayer · " if i == 1 else "") + corto(x),
                             callback_data=f"dia:{x.isoformat()}")
        for i, x in enumerate(dias)
    ]
    return InlineKeyboardMarkup([botones[i : i + 2] for i in range(0, len(botones), 2)])


def texto_dia_elegido(dia: date) -> str:
    if dia == hoy():
        return f"📅 Listo: las fotos van a HOY ({nombre_hoja(dia)})."
    return f"📅 Listo: las próximas fotos van a la hoja {nombre_hoja(dia)} ({corto(dia)}).\n/hoy para volver a hoy."


def elegir_dia(d: dict, dia: date) -> None:
    d["dia"] = None if dia == hoy() else dia


async def cmd_fecha(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    d = datos_chat(update.effective_chat.id)
    if context.args:
        dia = leer_fecha_texto(" ".join(context.args))
        if dia:
            elegir_dia(d, dia)
            await update.message.reply_text(texto_dia_elegido(dia))
            return
    await update.message.reply_text(
        f"¿De qué día son las fotos? (ahora: {nombre_hoja(dia_activo(d))})", reply_markup=teclado_fechas()
    )


async def cmd_hoy(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    d = datos_chat(update.effective_chat.id)
    d["dia"] = None
    await update.message.reply_text(texto_dia_elegido(hoy()))


async def boton_fecha(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    if not permitido(update):
        return
    dia = date.fromisoformat(q.data.split(":", 1)[1])
    elegir_dia(datos_chat(update.effective_chat.id), dia)
    await q.edit_message_text(texto_dia_elegido(dia))


async def recibir_texto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Si escriben una fecha ("ayer", "10/11") cambia el día; si no, muestra la ayuda."""
    dia = leer_fecha_texto(update.message.text) if permitido(update) else None
    if dia and len(update.message.text) <= 20:
        elegir_dia(datos_chat(update.effective_chat.id), dia)
        await update.message.reply_text(texto_dia_elegido(dia))
    else:
        await cmd_start(update, context)


async def recibir_imagen(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        await update.message.reply_text("No tienes permiso para usar este bot.")
        return
    msg = update.message
    chat_id = update.effective_chat.id
    bot = context.bot
    d = datos_chat(chat_id)
    pie = leer_fecha_texto(msg.caption) if msg.caption else None
    if pie:  # fecha escrita en el pie de la foto: cambia el día
        elegir_dia(d, pie)
    dia = dia_activo(d)
    d["version"] += 1
    d["procesando"] += 1
    lote = d["lote"]
    if dia not in lote["dias"]:
        lote["dias"].append(dia)
    lote["total"] += 1
    numero = lote["total"]
    await mostrar_estado(bot, chat_id, d)

    async def avisar(texto):
        lote["aviso"] = f"⏳ {texto}"
        await mostrar_estado(bot, chat_id, d)

    try:
        async with d["semaforo"]:
            media = msg.photo[-1] if msg.photo else msg.document
            mime = "image/jpeg" if msg.photo else (msg.document.mime_type or "image/jpeg")
            imagen = bytes(await (await media.get_file()).download_as_bytearray())
            productos = limpiar_productos(await llamar_gemini(imagen, mime, avisar))
        if not productos:
            raise ValueError("sin productos")
        etiqueta = f"Foto {numero} ({datetime.now(ZONA):%H:%M})"
        for i, p in enumerate(productos):
            p["imagen"] = etiqueta
            p["orden"] = (msg.message_id, i)
            p["dia"] = dia
        d["pendientes"].extend(productos)
        lote["productos"] += len(productos)
    except Exception as e:  # noqa: BLE001
        log.warning("Foto %s no leída: %s", numero, e)
        lote["fallidas"].append((numero, motivo_corto(e)))
    finally:
        lote["listas"] += 1
        lote["aviso"] = ""
        d["procesando"] -= 1

    await mostrar_estado(bot, chat_id, d)
    context.application.create_task(envio_automatico(bot, chat_id, d["version"]))


async def recibir_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    d = datos_chat(update.effective_chat.id)
    try:
        contenido = bytes(await (await update.message.document.get_file()).download_as_bytearray())
        hojas = load_workbook(io.BytesIO(contenido)).sheetnames
    except Exception:  # noqa: BLE001
        await update.message.reply_text("⚠️ No pude abrir ese Excel.")
        return
    d["base"] = contenido
    await update.message.reply_text(f"📎 Usaré este Excel como base ({len(hojas)} hojas).")


async def manejar_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.error("Error no controlado", exc_info=context.error)


async def al_iniciar(app: Application) -> None:
    log.info("Lugares: %s", ", ".join(NOMBRES_LUGARES))
    log.info("Comprobando clave de Gemini... %s", await comprobar_gemini())


def main() -> None:
    if not TELEGRAM_TOKEN or not GEMINI_API_KEY:
        raise SystemExit("Faltan las variables TELEGRAM_TOKEN y/o GEMINI_API_KEY")

    app = (
        Application.builder()
        .token(TELEGRAM_TOKEN)
        .concurrent_updates(True)  # lee varias fotos y responde comandos a la vez
        .post_init(al_iniciar)
        .build()
    )
    app.add_handler(CommandHandler(["start", "ayuda", "help"], cmd_start))
    app.add_handler(CommandHandler("excel", cmd_excel))
    app.add_handler(CommandHandler(["nuevo", "borrar"], cmd_nuevo))
    app.add_handler(CommandHandler("lugares", cmd_lugares))
    app.add_handler(CommandHandler("fecha", cmd_fecha))
    app.add_handler(CommandHandler("hoy", cmd_hoy))
    app.add_handler(CallbackQueryHandler(boton_fecha, pattern=r"^dia:"))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, recibir_imagen))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), recibir_excel))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_texto))
    app.add_error_handler(manejar_error)

    if URL_PUBLICA:
        log.info("Modo webhook: %s/telegram", URL_PUBLICA)
        app.run_webhook(
            listen="0.0.0.0",
            port=PUERTO,
            url_path="telegram",
            webhook_url=f"{URL_PUBLICA.rstrip('/')}/telegram",
            secret_token=secrets.token_urlsafe(24),
            allowed_updates=Update.ALL_TYPES,
        )
    else:
        log.info("Modo polling (prueba local)")
        app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

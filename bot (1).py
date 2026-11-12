"""
Bot de Telegram - Notas de egreso -> Excel con reparto por lugares
Subcentro de Vinces

Envías las fotos de las notas de egreso y el bot:
  1. Lee con la IA gratuita de Google Gemini: producto, lote, caducidad y cantidad.
  2. Guarda todo en UN solo Excel, con una hoja por día (la fecha en que mandaste las fotos).
  3. Reparte la cantidad de cada producto entre los lugares de distribución.
  4. Te devuelve el Excel y lo deja fijado en el chat, para seguir agregando días.

Variables de entorno (se ponen en Render > Environment, NO dentro del código):
  TELEGRAM_TOKEN       Token que te dio @BotFather                         (obligatorio)
  GEMINI_API_KEY       Clave de Google AI Studio                           (obligatorio)
  USUARIOS_PERMITIDOS  IDs de Telegram que pueden usar el bot, con comas    (opcional)
  LUGARES              Lugares de reparto, con comas. Para repartir en      (opcional)
                       proporciones distintas usa Nombre:peso
                       Ej: Urbano:3, Palenque:1, Nicaragua:1 ...
  ESPERA_EXCEL         Segundos sin fotos nuevas antes de mandar el Excel    (opcional, 90)
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
from telegram import Update
from telegram.ext import (
    Application,
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
USUARIOS_PERMITIDOS = {
    u.strip() for u in os.environ.get("USUARIOS_PERMITIDOS", "").split(",") if u.strip()
}
LUGARES_TXT = os.environ.get(
    "LUGARES",
    "Urbano, Palenque, Nicaragua, Antonio, Balzar, Juaneche, Bagatela, Clariza, Matecito",
)
ESPERA_EXCEL = int(os.environ.get("ESPERA_EXCEL", "90"))
NOMBRE_EXCEL = os.environ.get("NOMBRE_EXCEL", "Egresos_Subcentro_Vinces.xlsx").strip()
if not NOMBRE_EXCEL.lower().endswith(".xlsx"):
    NOMBRE_EXCEL += ".xlsx"
ZONA = ZoneInfo(os.environ.get("ZONA_HORARIA", "America/Guayaquil"))
MODELOS = [
    m
    for m in dict.fromkeys(
        [
            os.environ.get("GEMINI_MODEL", "").strip(),
            "gemini-flash-latest",
            "gemini-2.5-flash",
            "gemini-flash-lite-latest",
        ]
    )
    if m
]
GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta"

# Render define RENDER_EXTERNAL_URL solo. Sin ella, el bot usa "polling" (prueba en PC).
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

# Columnas de cada hoja
COL_FIJAS = ["N°", "Producto", "Lote", "Caducidad", "Cantidad"]
COL_FINALES = ["Total repartido", "Imagen"]
ENCABEZADOS = COL_FIJAS + NOMBRES_LUGARES + COL_FINALES

# ============================================================================
# Estado en memoria (se pierde si Render duerme el bot; el Excel queda fijado
# en el chat de Telegram, así que el historial no se pierde)
# ============================================================================
DATOS: dict[int, dict] = {}
CANDADOS: dict[int, asyncio.Lock] = {}


def datos_chat(chat_id: int) -> dict:
    if chat_id not in DATOS:
        DATOS[chat_id] = {
            "pendientes": [],  # productos leídos que aún no están en el Excel
            "imagenes": 0,
            "procesando": 0,
            "version": 0,
            "base": None,  # bytes del último Excel conocido
        }
    return DATOS[chat_id]


def candado(chat_id: int) -> asyncio.Lock:
    if chat_id not in CANDADOS:
        CANDADOS[chat_id] = asyncio.Lock()
    return CANDADOS[chat_id]


def permitido(update: Update) -> bool:
    return not USUARIOS_PERMITIDOS or str(update.effective_user.id) in USUARIOS_PERMITIDOS


def hoy() -> date:
    return datetime.now(ZONA).date()


# ============================================================================
# Gemini (API REST; acepta claves nuevas "AQ." y antiguas "AIza")
# ============================================================================
PROMPT = """Esta imagen es una NOTA DE EGRESO de medicamentos e insumos médicos.
La tabla tiene columnas como: N°, PRODUCTO, LOTE, CADUCIDAD y una columna con la
CANTIDAD (su encabezado puede ser un código, por ejemplo "12OT05").

Extrae TODAS las filas de productos, sin omitir ninguna, y responde SOLO con JSON así:
{"fecha_nota": "AAAA-MM-DD", "productos": [
  {"producto": "...", "lote": "...", "caducidad": "AAAA-MM-DD", "cantidad": 300}
]}

Reglas:
- "producto": el nombre COMPLETO tal como aparece (concentración, forma, presentación).
- "lote": exactamente igual, respetando ceros al inicio, puntos, guiones y letras.
- "caducidad": en formato AAAA-MM-DD.
- "cantidad": número entero, sin puntos ni comas de miles.
- "fecha_nota": la fecha de la nota si aparece; si no, "".
- Si algo no se lee claro, escribe tu mejor lectura seguida de "(?)"; nunca inventes datos.
- No incluyas la fila de encabezados ni filas de totales.
- Si la imagen no es una nota de egreso o no tiene productos, responde {"fecha_nota": "", "productos": []}."""

_modo_auth = {"actual": None}
MODOS_AUTH = ["header", "query", "bearer"]


def _auth(modo: str) -> tuple[dict, dict]:
    if modo == "header":
        return {"x-goog-api-key": GEMINI_API_KEY}, {}
    if modo == "query":
        return {}, {"key": GEMINI_API_KEY}
    return {"Authorization": f"Bearer {GEMINI_API_KEY}"}, {}


def _modos_a_probar() -> list[str]:
    actual = _modo_auth["actual"]
    return [actual] if actual else MODOS_AUTH


def _es_error_auth(r: httpx.Response) -> bool:
    return r.status_code in (400, 401, 403) and any(
        t in r.text
        for t in ("API_KEY", "API key", "UNAUTHENTICATED", "ACCESS_TOKEN", "credential", "PERMISSION_DENIED")
    )


async def comprobar_gemini() -> str:
    """Revisa al arrancar que la clave funcione (se ve en los Logs de Render)."""
    async with httpx.AsyncClient(timeout=30) as cli:
        ultimo = ""
        for modo in MODOS_AUTH:
            h, q = _auth(modo)
            try:
                r = await cli.get(f"{GEMINI_URL}/models", headers=h, params={**q, "pageSize": 50})
            except httpx.HTTPError as e:
                return f"sin conexión con Gemini ({e})"
            if r.status_code == 200:
                _modo_auth["actual"] = modo
                return f"OK (modo {modo})"
            ultimo = f"{r.status_code}: {r.text[:200]}"
        return f"la clave NO funcionó -> {ultimo}"


async def llamar_gemini(imagen: bytes, mime: str) -> dict:
    cuerpo = {
        "contents": [
            {
                "parts": [
                    {"inline_data": {"mime_type": mime, "data": base64.b64encode(imagen).decode()}},
                    {"text": PROMPT},
                ]
            }
        ],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0},
    }
    ultimo_error = "sin respuesta"
    async with httpx.AsyncClient(timeout=180) as cli:
        for modelo in MODELOS:
            intento = 0
            while intento < 3:
                intento += 1
                auth_ok = False
                for modo in _modos_a_probar():
                    h, q = _auth(modo)
                    try:
                        r = await cli.post(
                            f"{GEMINI_URL}/models/{modelo}:generateContent",
                            headers=h, params=q, json=cuerpo,
                        )
                    except httpx.HTTPError as e:
                        ultimo_error = f"conexión: {e}"
                        break
                    if _es_error_auth(r) and not _modo_auth["actual"]:
                        ultimo_error = f"{r.status_code} {r.text[:200]}"
                        continue  # prueba otra forma de enviar la clave
                    auth_ok = True
                    break
                else:
                    raise RuntimeError(f"La clave de Gemini no fue aceptada ({ultimo_error})")
                if not auth_ok:
                    await asyncio.sleep(3)
                    continue

                if r.status_code == 200:
                    _modo_auth["actual"] = _modo_auth["actual"] or modo
                    datos = r.json()
                    try:
                        partes = datos["candidates"][0]["content"]["parts"]
                        texto = "".join(p.get("text", "") for p in partes if not p.get("thought"))
                        return interpretar_json(texto)
                    except (KeyError, IndexError, json.JSONDecodeError) as e:
                        ultimo_error = f"respuesta rara: {e}"
                        await asyncio.sleep(2)
                        continue
                ultimo_error = f"{r.status_code} {r.text[:200]}"
                if r.status_code in (429, 500, 503):
                    espera = 20 * intento
                    log.warning("Gemini ocupado/límite (%s). Espero %ss", r.status_code, espera)
                    await asyncio.sleep(espera)
                    continue
                if r.status_code == 404:
                    log.warning("Modelo %s no disponible, pruebo otro", modelo)
                    break
                log.warning("Error Gemini %s", ultimo_error)
                break
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
    for p in resultado.get("productos", []) or []:
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

    # Título
    ws.merge_cells(f"A1:{ultima}1")
    ws["A1"] = f"NOTAS DE EGRESO — registradas el {dia.strftime('%d/%m/%Y')}"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    # Encabezados (fila 2)
    for c, nombre in enumerate(ENCABEZADOS, start=1):
        celda = ws.cell(2, c, nombre)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = AZUL
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = BORDE
    ws.row_dimensions[2].height = 30

    # Datos
    r = 3
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
        ws.cell(r, 3).number_format = "@"
        if isinstance(f["caducidad"], date):
            ws.cell(r, 4).number_format = "DD/MM/YYYY"
        r += 1

    # Totales
    fin = r - 1
    ws.cell(r, 2, "TOTAL")
    for c in [c_cant] + list(range(c_l1, c_tot + 1)):
        col = get_column_letter(c)
        ws.cell(r, c, f"=SUM({col}3:{col}{fin})" if fin >= 3 else 0)
    for c in range(1, n_cols + 1):
        celda = ws.cell(r, c)
        celda.font = Font(bold=True)
        celda.fill = GRIS
        celda.border = BORDE
        celda.alignment = Alignment(horizontal="center", vertical="center")

    # Si "Total repartido" no coincide con "Cantidad", la celda se pinta de rojo
    if fin >= 3:
        rango = f"{get_column_letter(c_tot)}3:{get_column_letter(c_tot)}{fin}"
        ws.conditional_formatting.add(
            rango,
            FormulaRule(
                formula=[f"${get_column_letter(c_tot)}3<>${get_column_letter(c_cant)}3"],
                fill=ROJO,
            ),
        )

    # Anchos
    anchos = {1: 5, 2: 60, 3: 16, 4: 12, 5: 10}
    for c in range(1, n_cols + 1):
        ancho = anchos.get(c, max(10, len(ENCABEZADOS[c - 1]) + 2))
        ws.column_dimensions[get_column_letter(c)].width = ancho
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{ultima}{max(fin, 2)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "2:2"


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
# Envío del Excel (y lo fija en el chat para no perder el historial)
# ============================================================================
async def obtener_base(bot, chat_id: int, d: dict) -> bytes | None:
    if d["base"]:
        return d["base"]
    try:
        chat = await bot.get_chat(chat_id)
        fijado = chat.pinned_message
        if fijado and fijado.document and (fijado.document.file_name or "").lower().endswith(".xlsx"):
            archivo = await fijado.document.get_file()
            d["base"] = bytes(await archivo.download_as_bytearray())
            log.info("Excel recuperado del mensaje fijado")
    except Exception:  # noqa: BLE001
        log.exception("No pude leer el mensaje fijado")
    return d["base"]


async def generar_y_enviar(bot, chat_id: int, automatico: bool = False) -> None:
    async with candado(chat_id):
        d = datos_chat(chat_id)
        if not d["pendientes"]:
            return
        dia = hoy()
        base = await obtener_base(bot, chat_id, d)
        excel, total_hoja = agregar_al_excel(base, dia, d["pendientes"])
        nuevos = len(d["pendientes"])
        texto = (
            f"📊 Hoja {nombre_hoja(dia)}: +{nuevos} producto(s), {total_hoja} en total ese día.\n"
            f"Repartido entre: {', '.join(NOMBRES_LUGARES)}."
        )
        if automatico:
            texto += "\n(Enviado automáticamente; si mandas más fotos, te llega actualizado.)"
        msg = await bot.send_document(
            chat_id=chat_id, document=io.BytesIO(excel), filename=NOMBRE_EXCEL, caption=texto
        )
        d["base"] = excel
        d["pendientes"] = []
        d["imagenes"] = 0
        try:
            await bot.unpin_all_chat_messages(chat_id)
            await bot.pin_chat_message(chat_id, msg.message_id, disable_notification=True)
        except Exception:  # noqa: BLE001
            log.warning("No pude fijar el mensaje del Excel")


async def envio_automatico(bot, chat_id: int, version: int) -> None:
    await asyncio.sleep(ESPERA_EXCEL)
    d = DATOS.get(chat_id)
    if d and d["version"] == version and d["procesando"] == 0 and d["pendientes"]:
        try:
            await generar_y_enviar(bot, chat_id, automatico=True)
        except Exception:  # noqa: BLE001
            log.exception("Error en envío automático")


# ============================================================================
# Comandos de Telegram
# ============================================================================
def texto_ayuda() -> str:
    return (
        "Hola 👋 Mándame las fotos de las notas de egreso (una o varias a la vez).\n\n"
        "📷 Por cada foto te digo cuántos productos leí.\n"
        f"📊 Cuando pasan {ESPERA_EXCEL} s sin fotos nuevas te mando el Excel solo "
        "(o escribe /excel para recibirlo ya).\n"
        "📅 Todo se guarda en el MISMO Excel, en una hoja con la fecha de hoy.\n"
        f"📦 Cada cantidad se reparte entre: {', '.join(NOMBRES_LUGARES)}.\n"
        "📌 El Excel queda fijado arriba del chat; ahí se siguen sumando los días.\n\n"
        "Otros comandos:\n"
        "/ver — lo que llevo leído\n"
        "/nuevo — descartar lo leído que aún no está en el Excel\n"
        "/lugares — ver cómo se reparte\n"
        "Si me mandas un Excel, lo uso como base para seguir agregando hojas."
    )


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        texto_ayuda() + f"\n\nTu ID de Telegram: {update.effective_user.id}"
    )


async def cmd_ver(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    d = datos_chat(update.effective_chat.id)
    await update.message.reply_text(
        f"Fotos leídas (aún sin Excel): {d['imagenes']}\n"
        f"Productos pendientes: {len(d['pendientes'])}\n"
        f"Hoja de hoy: {nombre_hoja(hoy())}"
    )


async def cmd_lugares(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    total = sum(p for _, p in LUGARES)
    lineas = [f"• {n}: {p / total * 100:.1f} %" for n, p in LUGARES]
    ejemplo = repartir(100)
    await update.message.reply_text(
        "Reparto actual:\n" + "\n".join(lineas)
        + "\n\nEjemplo con 100 unidades: "
        + ", ".join(f"{n} {v}" for n, v in zip(NOMBRES_LUGARES, ejemplo))
    )


async def cmd_nuevo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    d = datos_chat(update.effective_chat.id)
    d["pendientes"], d["imagenes"] = [], 0
    d["version"] += 1
    await update.message.reply_text("🗑 Descarté lo leído. El Excel fijado no cambió.")


async def cmd_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    chat_id = update.effective_chat.id
    d = datos_chat(chat_id)
    d["version"] += 1  # cancela el envío automático pendiente
    if d["pendientes"]:
        await generar_y_enviar(context.bot, chat_id)
        return
    base = await obtener_base(context.bot, chat_id, d)
    if base:
        await update.message.reply_document(
            document=io.BytesIO(base), filename=NOMBRE_EXCEL,
            caption="No hay fotos nuevas; este es el Excel actual.",
        )
    else:
        await update.message.reply_text("Todavía no hay registros. Mándame las fotos 📷")


async def recibir_imagen(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        await update.message.reply_text("No tienes permiso para usar este bot.")
        return
    msg = update.message
    chat_id = update.effective_chat.id
    d = datos_chat(chat_id)
    d["version"] += 1
    d["procesando"] += 1
    numero = d["imagenes"] + 1
    aviso = await msg.reply_text(f"⏳ Leyendo foto {numero}...")
    try:
        if msg.photo:
            archivo = await msg.photo[-1].get_file()
            mime = "image/jpeg"
        else:
            archivo = await msg.document.get_file()
            mime = msg.document.mime_type or "image/jpeg"
        imagen = bytes(await archivo.download_as_bytearray())
        productos = limpiar_productos(await llamar_gemini(imagen, mime))
        if not productos:
            await aviso.edit_text("🤔 No encontré productos en esa foto. ¿Es una nota de egreso?")
        else:
            d["imagenes"] = numero
            etiqueta = f"Foto {numero} ({datetime.now(ZONA):%H:%M})"
            for p in productos:
                p["imagen"] = etiqueta
            d["pendientes"].extend(productos)
            dudosos = sum("(?)" in p["producto"] + p["lote"] for p in productos)
            sin_cant = sum(p["cantidad"] is None for p in productos)
            extra = ""
            if dudosos:
                extra += f"\n⚠️ {dudosos} dato(s) dudoso(s) marcados con (?)."
            if sin_cant:
                extra += f"\n⚠️ {sin_cant} producto(s) sin cantidad legible."
            await aviso.edit_text(
                f"✅ Foto {numero}: {len(productos)} producto(s). "
                f"Llevo {len(d['pendientes'])}.{extra}\n"
                "Manda más fotos o escribe /excel"
            )
    except Exception as e:  # noqa: BLE001
        log.exception("Error leyendo imagen")
        await aviso.edit_text(
            "⚠️ No pude leer esa foto. Prueba de nuevo en un minuto o con una foto más clara."
            f"\n({str(e)[:180]})"
        )
    finally:
        d["procesando"] -= 1
    context.application.create_task(envio_automatico(context.bot, chat_id, d["version"]))


async def recibir_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not permitido(update):
        return
    d = datos_chat(update.effective_chat.id)
    try:
        archivo = await update.message.document.get_file()
        contenido = bytes(await archivo.download_as_bytearray())
        hojas = load_workbook(io.BytesIO(contenido)).sheetnames
    except Exception:  # noqa: BLE001
        await update.message.reply_text("⚠️ No pude abrir ese Excel.")
        return
    d["base"] = contenido
    await update.message.reply_text(
        f"📎 Listo, usaré este Excel como base ({len(hojas)} hoja(s)). "
        "Las fotos que mandes se agregarán en la hoja de hoy."
    )


async def manejar_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.error("Error no controlado", exc_info=context.error)


async def al_iniciar(app: Application) -> None:
    log.info("Lugares de reparto: %s", ", ".join(NOMBRES_LUGARES))
    log.info("Comprobando clave de Gemini... %s", await comprobar_gemini())


# ============================================================================
# Arranque
# ============================================================================
def main() -> None:
    if not TELEGRAM_TOKEN:
        raise SystemExit("Falta la variable TELEGRAM_TOKEN")
    if not GEMINI_API_KEY:
        raise SystemExit("Falta la variable GEMINI_API_KEY")

    app = Application.builder().token(TELEGRAM_TOKEN).post_init(al_iniciar).build()
    app.add_handler(CommandHandler(["start", "ayuda", "help"], cmd_start))
    app.add_handler(CommandHandler("excel", cmd_excel))
    app.add_handler(CommandHandler("ver", cmd_ver))
    app.add_handler(CommandHandler("lugares", cmd_lugares))
    app.add_handler(CommandHandler(["nuevo", "borrar"], cmd_nuevo))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, recibir_imagen))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), recibir_excel))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, cmd_start))
    app.add_error_handler(manejar_error)

    if URL_PUBLICA:
        ruta = "telegram"
        log.info("Modo webhook: %s/%s (puerto %s)", URL_PUBLICA, ruta, PUERTO)
        app.run_webhook(
            listen="0.0.0.0",
            port=PUERTO,
            url_path=ruta,
            webhook_url=f"{URL_PUBLICA.rstrip('/')}/{ruta}",
            secret_token=secrets.token_urlsafe(24),
            allowed_updates=Update.ALL_TYPES,
        )
    else:
        log.info("Modo polling (prueba local)")
        app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

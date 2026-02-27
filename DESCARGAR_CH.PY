# DESCARGAR_CH_COMPLETO_FINAL_V3.py
# ------------------------------------------------------------------------------
# Automatización Gestiona Salud - Clínica de Heridas
# Correcciones:
# 1. Interfaz Tkinter (Error -width solucionado).
# 2. Playwright (Error path solucionado: detecta si es .py o .exe).
# 3. Lógica Multi-Herida (Llave compuesta).
# AJUSTES SOLICITADOS (SIN TOCAR LO QUE YA FUNCIONA):
# 4. Desactivar opción "Modo Headless" (siempre navegador visible).
# 5. Agregar botones Pausar/Reanudar y Detener (control cooperativo).
# 6. Generar REPORTE en Excel al final (misma carpeta DESCARGAS) con estructura
#    basada en los campos de lectura utilizados (B/H/I-L/BC/BG/BH/BI/BJ) + estado.
# 7. Después de introducir contraseña, esperar X segundos antes de "Ingresar"
#    (configurable desde la interfaz; default = 5s).
# ------------------------------------------------------------------------------

import os
import sys
import re
import json
import csv
import queue
import threading
import calendar as pycal
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import Optional, List, Tuple, Dict, Any

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

# Intentar importar openpyxl
try:
    from openpyxl import load_workbook
    from openpyxl.utils import column_index_from_string
except Exception:
    load_workbook = None
    column_index_from_string = None

# ==============================================================================
# 0) CONFIGURACIÓN Y RUTAS (CORREGIDO PARA .PY y .EXE)
# ==============================================================================
def configurar_entorno():
    # Detectar si es ejecutable compilado (.exe) o script (.py)
    es_frozen = getattr(sys, "frozen", False)

    if es_frozen:
        base_path = os.path.dirname(sys.executable)
        # Si es .exe, forzamos la ruta local 'empaquetar'
        ruta_browsers = os.path.join(base_path, "empaquetar")
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = ruta_browsers
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
        ruta_browsers = ""
        # Si es .py, NO forzamos la variable de entorno para que use la instalación global
        if "PLAYWRIGHT_BROWSERS_PATH" in os.environ:
            del os.environ["PLAYWRIGHT_BROWSERS_PATH"]

    return base_path, ruta_browsers


BASE_DIR, RUTA_NAVEGADORES = configurar_entorno()
APP_CONFIG_PATH = os.path.join(BASE_DIR, "app_config.json")

# URLs fijas
DEFAULT_URL_LOGIN = "https://saludgestiona.com/business/signin/MTAwNQ=="
DEFAULT_URL_REPORTES = "https://saludgestiona.com/business/reports"

# Parámetros Playwright
DEFAULT_HEADLESS = False
DEFAULT_SLOWMO_MS = 0
DEFAULT_TIMEOUT_MS = 120000

# Selectores Login
SEL_DOC = "input[name='document'], input[placeholder*='Documento']"
SEL_PASS = "input[name='password'], input[placeholder*='Contraseña'], input[type='password']"

# Selectores Reportes
SEL_BTN_GENERAR = "#button_pdf"
SEL_CAT_REGISTROS = "a.category_for_reports[data-category='2'], a.category_for_reports:has-text('Registros')"

# Valores Fijos Formulario
TIPO_REPORTE_VALUE = "10"  # 2. Evoluciones (PDF)
PLANTILLA_VALUE = "13"     # EV008 - CLINICA DE HERIDAS
PERS1_VALUE = "755"        # UBICACION GENERAL
PERS2_VALUE = "756"        # ZONA
PERS3_VALUE = "757"        # ZONA ESPECIFICA

# ==============================================================================
# 1) HELPERS GENERALES
# ==============================================================================
def _read_json_file(path: str) -> dict:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}

def _write_json_file(path: str, data: dict) -> None:
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except Exception:
        pass

def _extract_spreadsheet_id(text: str) -> str:
    s = (text or "").strip()
    if not s: return ""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", s)
    if m: return m.group(1)
    s = s.split("?")[0].split("#")[0].strip()
    if "/edit" in s: s = s.split("/edit")[0].strip()
    return s

def _norm_text(s: str) -> str:
    s = (s or "").strip().upper()
    s = re.sub(r"\s+", " ", s)
    return s

def _norm_doc(s) -> str:
    s = "" if s is None else str(s)
    s = s.strip()
    s = s.replace(" ", "").replace(".", "").replace(",", "")
    return s

def _parse_date(value) -> Optional[datetime]:
    if value is None or value == "": return None
    if isinstance(value, datetime): return value
    if isinstance(value, date): return datetime(value.year, value.month, value.day)

    s = str(value).strip()
    if not s: return None

    fmts = [
        "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y", "%d-%m-%y",
        "%Y%m%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
    ]
    for f in fmts:
        try:
            return datetime.strptime(s, f)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None

def _parse_date_ui_start(s: str) -> Optional[datetime]:
    d = _parse_date(s)
    if d is None: return None
    return datetime(d.year, d.month, d.day, 0, 0, 0)

def _parse_date_ui_end(s: str) -> Optional[datetime]:
    d = _parse_date(s)
    if d is None: return None
    return datetime(d.year, d.month, d.day, 23, 59, 59)

def _in_range(dt: datetime, desde: Optional[datetime], hasta: Optional[datetime]) -> bool:
    if desde and dt < desde: return False
    if hasta and dt > hasta: return False
    return True

def _safe_filename(s: str, max_len: int = 180) -> str:
    s = (s or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) > max_len: s = s[:max_len].rstrip()
    return s

def _join_name(parts) -> str:
    clean = []
    for p in parts:
        p = (p or "").strip()
        if p: clean.append(p)
    return _norm_text(" ".join(clean))

def _ym_now() -> str:
    n = datetime.now()
    return f"{n.year:04d}-{n.month:02d}"

def _month_range(ym: str) -> Tuple[str, str]:
    try:
        y, m = ym.split("-")
        y = int(y)
        m = int(m)
        first = date(y, m, 1)
        last_day = pycal.monthrange(y, m)[1]
        last = date(y, m, last_day)
        today = date.today()
        last_use = today if (y == today.year and m == today.month) else last
        return first.strftime("%Y-%m-%d"), last_use.strftime("%Y-%m-%d")
    except Exception:
        t = date.today().strftime("%Y-%m-%d")
        return t, t

def _detectar_encoding_csv(path: str) -> str:
    candidatos = ["utf-8-sig", "utf-16", "utf-16le", "utf-16be", "cp1252", "latin-1"]
    with open(path, "rb") as f:
        raw = f.read(8192)
    for enc in candidatos:
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return "latin-1"

def _now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# ==============================================================================
# 1.1) REPORTE EXCEL FINAL
# ==============================================================================
def escribir_reporte_excel(output_dir: str, cfg, resultados: List[Dict[str, Any]], log):
    """
    Crea un Excel en la misma carpeta de DESCARGAS con columnas basadas en la
    estructura de lectura usada (B/H/BC/BG/BH/BI/BJ + nombre) y campos de ejecución.
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
    except Exception:
        raise RuntimeError("Para generar el reporte en Excel necesitas openpyxl. Instala: pip install openpyxl")

    os.makedirs(output_dir, exist_ok=True)

    origen = "ARCHIVO" if cfg.mode == "archivo" else "SHEETS"
    base = "BASE"
    try:
        if cfg.mode == "archivo" and cfg.file_path:
            base = os.path.splitext(os.path.basename(cfg.file_path))[0]
        elif cfg.mode == "sheets":
            base = f"{(cfg.sheets_tab or 'TAB')}".strip() or "TAB"
    except Exception:
        base = "BASE"

    fname = f"REPORTE_DESCARGAS_CH_{origen}_{_safe_filename(base, 40)}_{_now_stamp()}.xlsx"
    out_path = os.path.join(output_dir, fname)

    wb = Workbook()
    ws = wb.active
    ws.title = "REPORTE"

    headers = [
        "N", "ESTADO", "ERROR",
        "ARCHIVO_PDF", "RUTA_PDF",
        "B_FECHA_EVOLUCION",
        "H_DOCUMENTO",
        "NOMBRE (I-J-K-L)",
        "BC_ETIOLOGIA",
        "BG_UBICACION_GENERAL (Valor 1)",
        "BH_ZONA (Valor 2)",
        "BI_ZONA_ESPECIFICA (Valor 3)",
        "BJ_CLASIFICACION",
        "FILA_ORIGEN",
        "ORIGEN",
        "MODO",
        "ARCHIVO_ORIGEN",
        "SHEETS_ID",
        "SHEETS_TAB",
        "FECHA_DESDE_FILTRO",
        "FECHA_HASTA_FILTRO",
        "ESPERA_LOGIN_SEGUNDOS",
        "FECHA_EJECUCION"
    ]

    ws.append(headers)

    bold = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    fejec = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for r in resultados:
        ws.append([
            r.get("n", ""),
            r.get("estado", ""),
            r.get("error", ""),
            r.get("archivo_pdf", ""),
            r.get("ruta_pdf", ""),
            r.get("fecha_evolucion", ""),
            r.get("doc", ""),
            r.get("nombre", ""),
            r.get("etiologia", ""),
            r.get("bg", ""),
            r.get("bh", ""),
            r.get("bi", ""),
            r.get("clasificacion", ""),
            r.get("fila_origen", ""),
            origen,
            cfg.mode,
            cfg.file_path if cfg.mode == "archivo" else "",
            cfg.sheets_id if cfg.mode == "sheets" else "",
            cfg.sheets_tab if cfg.mode == "sheets" else "",
            cfg.fecha_desde.strftime("%Y-%m-%d"),
            cfg.fecha_hasta.strftime("%Y-%m-%d"),
            getattr(cfg, "login_wait_seconds", 5),
            fejec
        ])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 6, "B": 10, "C": 45, "D": 35, "E": 60, "F": 18, "G": 18, "H": 30,
        "I": 24, "J": 28, "K": 22, "L": 26, "M": 20, "N": 12, "O": 10, "P": 10,
        "Q": 40, "R": 25, "S": 18, "T": 16, "U": 16, "V": 18, "W": 19,
    }
    for col, w in widths.items():
        try:
            ws.column_dimensions[col].width = w
        except Exception:
            pass

    wb.save(out_path)
    log(f"[REPORTE] Excel generado: {out_path}")
    return out_path


# ==============================================================================
# 2) MODELO DE DATOS Y LECTURA
# ==============================================================================
@dataclass
class Task:
    doc: str
    nombre: str
    etiologia: str
    bg: str
    bh: str
    bi: str
    clasificacion: str
    fecha_evolucion: datetime
    row: int

# --- LECTURA EXCEL ---
def leer_excel_y_deduplicar(path_xlsx: str, sheet_name: Optional[str],
                           fecha_desde: Optional[datetime], fecha_hasta: Optional[datetime], log) -> List[Task]:
    if load_workbook is None:
        raise RuntimeError("openpyxl no está disponible. Instala: pip install openpyxl")

    wb = load_workbook(path_xlsx, data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

    col_B = column_index_from_string("B")
    col_H = column_index_from_string("H")
    col_I = column_index_from_string("I")
    col_J = column_index_from_string("J")
    col_K = column_index_from_string("K")
    col_L = column_index_from_string("L")
    col_BC = column_index_from_string("BC")
    col_BG = column_index_from_string("BG")
    col_BH = column_index_from_string("BH")
    col_BI = column_index_from_string("BI")
    col_BJ = column_index_from_string("BJ")

    best = {}

    for r in range(2, ws.max_row + 1):
        fecha = _parse_date(ws.cell(r, col_B).value)
        if not fecha: continue
        if not _in_range(fecha, fecha_desde, fecha_hasta): continue

        doc = _norm_doc(ws.cell(r, col_H).value)
        if not doc: continue

        nombre = _join_name([ws.cell(r, col_I).value, ws.cell(r, col_J).value,
                             ws.cell(r, col_K).value, ws.cell(r, col_L).value])
        etiologia = _norm_text(ws.cell(r, col_BC).value)
        bg = _norm_text(ws.cell(r, col_BG).value)
        bh = _norm_text(ws.cell(r, col_BH).value)
        bi = _norm_text(ws.cell(r, col_BI).value)
        clasificacion = _norm_text(ws.cell(r, col_BJ).value)

        if not bg or not bh or not bi: continue

        key = (doc, bg, bh, bi, etiologia, clasificacion)

        t = Task(doc, nombre or "SIN_NOMBRE", etiologia or "SIN_ETIOLOGIA",
                 bg, bh, bi, clasificacion or "SIN_CLASIFICACION", fecha, r)

        prev = best.get(key)
        if prev is None:
            best[key] = (fecha, r, t)
        else:
            prev_fecha, prev_row, _ = prev
            if fecha > prev_fecha or (fecha == prev_fecha and r > prev_row):
                best[key] = (fecha, r, t)

    tasks = [v[2] for v in best.values()]
    log(f"[EXCEL] Total tareas (Multi-Herida): {len(tasks)}")
    return tasks

# --- LECTURA CSV ---
def leer_csv_y_deduplicar(path_csv: str, fecha_desde: Optional[datetime], fecha_hasta: Optional[datetime], log) -> List[Task]:
    idx_B = 1
    idx_H = 7
    idx_I = 8
    idx_J = 9
    idx_K = 10
    idx_L = 11
    idx_BC = 54
    idx_BG = 58
    idx_BH = 59
    idx_BI = 60
    idx_BJ = 61

    enc = _detectar_encoding_csv(path_csv)
    log(f"[CSV] Encoding: {enc}")

    with open(path_csv, "r", encoding=enc, errors="replace", newline="") as f:
        head = [f.readline() for _ in range(10)]
    best_delim = ","
    try:
        import csv as _csv
        s = _csv.Sniffer()
        best_delim = s.sniff("".join(head), delimiters=";,|\t").delimiter
    except:
        pass
    log(f"[CSV] Delimitador: {repr(best_delim)}")

    best = {}

    with open(path_csv, "r", encoding=enc, errors="replace", newline="") as f:
        reader = csv.reader(f, delimiter=best_delim)
        for rownum, row in enumerate(reader, start=1):
            if not row or len(row) <= idx_BJ: continue

            if rownum == 1 and not _parse_date(row[idx_B]): continue

            fecha = _parse_date(row[idx_B])
            if not fecha or not _in_range(fecha, fecha_desde, fecha_hasta): continue

            doc = _norm_doc(row[idx_H])
            if not doc: continue

            nombre = _join_name([row[idx_I], row[idx_J], row[idx_K], row[idx_L]])
            etiologia = _norm_text(row[idx_BC])
            bg = _norm_text(row[idx_BG])
            bh = _norm_text(row[idx_BH])
            bi = _norm_text(row[idx_BI])
            clasificacion = _norm_text(row[idx_BJ])

            if not bg or not bh or not bi: continue

            key = (doc, bg, bh, bi, etiologia, clasificacion)

            t = Task(doc, nombre or "SIN_NOMBRE", etiologia or "SIN_ETIOLOGIA",
                     bg, bh, bi, clasificacion or "SIN_CLASIFICACION", fecha, rownum)

            prev = best.get(key)
            if prev is None:
                best[key] = (fecha, rownum, t)
            else:
                prev_fecha, prev_row, _ = prev
                if fecha > prev_fecha or (fecha == prev_fecha and rownum > prev_row):
                    best[key] = (fecha, rownum, t)

    tasks = [v[2] for v in best.values()]
    log(f"[CSV] Total tareas (Multi-Herida): {len(tasks)}")
    return tasks

# --- LECTURA GOOGLE SHEETS ---
def get_sheets_service(credentials_path: str, token_path: str):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

    creds = None
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        except:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return build("sheets", "v4", credentials=creds)

def leer_sheets_y_deduplicar(spreadsheet_id: str, tab_name: str,
                            credentials_path: str, token_path: str,
                            fecha_desde: Optional[datetime], fecha_hasta: Optional[datetime], log) -> List[Task]:
    svc = get_sheets_service(credentials_path, token_path)
    rng = f"{tab_name}!A:BK"
    resp = svc.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=rng).execute()
    values = resp.get("values", [])

    if not values or len(values) < 2:
        log("[SHEETS] No hay filas para procesar.")
        return []

    idx_B, idx_H = 1, 7
    idx_I, idx_J, idx_K, idx_L = 8, 9, 10, 11
    idx_BC, idx_BG, idx_BH, idx_BI, idx_BJ = 54, 58, 59, 60, 61

    best = {}

    for row_idx, row in enumerate(values[1:], start=2):
        def get(i): return row[i] if i < len(row) else ""

        fecha = _parse_date(get(idx_B))
        if not fecha or not _in_range(fecha, fecha_desde, fecha_hasta): continue

        doc = _norm_doc(get(idx_H))
        if not doc: continue

        nombre = _join_name([get(idx_I), get(idx_J), get(idx_K), get(idx_L)])
        etiologia = _norm_text(get(idx_BC))
        bg = _norm_text(get(idx_BG))
        bh = _norm_text(get(idx_BH))
        bi = _norm_text(get(idx_BI))
        clasificacion = _norm_text(get(idx_BJ))

        if not bg or not bh or not bi: continue

        key = (doc, bg, bh, bi, etiologia, clasificacion)

        t = Task(
            doc,
            nombre or "SIN_NOMBRE",
            etiologia or "SIN_ETIOLOGIA",
            bg, bh, bi,
            clasificacion or "SIN_CLASIFICACION",
            fecha,
            row_idx
        )

        prev = best.get(key)
        if prev is None:
            best[key] = (fecha, row_idx, t)
        else:
            prev_fecha, prev_row, _ = prev
            if fecha > prev_fecha or (fecha == prev_fecha and row_idx > prev_row):
                best[key] = (fecha, row_idx, t)

    tasks = [v[2] for v in best.values()]
    log(f"[SHEETS] Total tareas (Multi-Herida): {len(tasks)}")
    return tasks


# ==============================================================================
# 3) LÓGICA PLAYWRIGHT
# ==============================================================================
def _login_ok(page) -> bool:
    try:
        if page.locator(SEL_PASS).first.is_visible(timeout=800): return False
    except:
        pass
    url = (page.url or "").lower()
    if "signin" in url or "sign-in" in url: return False
    return True

def submit_login_enter_retry(page, attempts: int = 2, wait_seconds: int = 3):
    for _ in range(attempts):
        page.press(SEL_PASS, "Enter")
        page.wait_for_timeout(wait_seconds * 1000)
        try:
            page.wait_for_load_state("domcontentloaded", timeout=10000)
        except:
            pass
        if _login_ok(page): return
    raise RuntimeError("Login falló: Credenciales incorrectas o error de carga.")

def do_login(page, user, password, log, pre_enter_wait_seconds: int = 5):
    log("[LOGIN] Navegando...")
    page.goto(DEFAULT_URL_LOGIN, wait_until="domcontentloaded", timeout=DEFAULT_TIMEOUT_MS)
    page.wait_for_timeout(500)
    page.fill(SEL_DOC, user)
    page.fill(SEL_PASS, password)

    try:
        s = int(pre_enter_wait_seconds)
    except Exception:
        s = 5
    if s < 0:
        s = 0
    log(f"[LOGIN] Esperando {s}s antes de ingresar...")
    page.wait_for_timeout(s * 1000)

    log("[LOGIN] Enviando credenciales...")
    submit_login_enter_retry(page, attempts=2, wait_seconds=3)
    log(f"[LOGIN] OK. URL actual: {page.url}")

def _xpath_select_following(label_contains: str) -> str:
    t = (label_contains or "").replace("'", "\\'")
    return f"xpath=//*[self::label or self::div or self::span or self::p or self::h1 or self::h2 or self::h3][contains(normalize-space(),'{t}')]/following::select[1]"

def _wait_select_enabled(page, selector: str, timeout_ms: int = 12000):
    page.wait_for_selector(selector, timeout=timeout_ms)
    loc = page.locator(selector).first
    try:
        h = loc.element_handle(timeout=timeout_ms)
        page.wait_for_function("el => el && !el.disabled", h, timeout=timeout_ms)
    except:
        pass
    return loc

def _select_value_fast(page, select_selector: str, value: str, log_prefix: str, timeout_ms: int = 2500):
    if not value: return
    loc = page.locator(select_selector).first
    if loc.count() == 0:
        loc = _wait_select_enabled(page, select_selector, timeout_ms=12000)
    try:
        cur = loc.input_value(timeout=800)
        if cur == value: return
    except:
        pass

    last_err = None
    for _ in range(2):
        try:
            page.select_option(select_selector, value=value, timeout=timeout_ms)
            try:
                if loc.input_value(timeout=800) == value: return
            except:
                return
        except Exception as e:
            last_err = e
        page.wait_for_timeout(150)
    raise RuntimeError(f"{log_prefix}: error value={value} ({last_err})")

def _find_option_value_contains(select_locator, needle_upper: str) -> str:
    try:
        return select_locator.evaluate(
            """(s, t) => {
                const norm = (x) => (x || '').replace(/\s+/g,' ').trim().toUpperCase();
                const opts = Array.from(s.options || []);
                const hit = opts.find(o => norm(o.textContent).includes(t));
                return hit ? hit.value : '';
            }""", needle_upper)
    except:
        return ""

def _select_by_label_contains(page, select_selector: str, label_text: str, timeout_ms: int = 8000):
    if not label_text: return
    loc = _wait_select_enabled(page, select_selector, timeout_ms=timeout_ms)
    needle = _norm_text(label_text)

    try:
        cur = loc.locator("option:checked").first.text_content() or ""
        if needle in _norm_text(cur): return
    except:
        pass

    val = _find_option_value_contains(loc, needle)
    if not val:
        raise RuntimeError(f"Opción no encontrada: '{label_text}' en {select_selector}")
    page.select_option(select_selector, value=val, timeout=4000)

def asegurar_fijos_formulario(page, log):
    tipo_sel = _xpath_select_following("Tipo Reporte")
    _select_value_fast(page, tipo_sel, TIPO_REPORTE_VALUE, "Tipo Reporte", 2000)

    plant_sel = _xpath_select_following("Plantilla")
    try:
        _select_value_fast(page, plant_sel, PLANTILLA_VALUE, "Plantilla", 2500)
    except:
        page.wait_for_selector("select:has(option[value='13'])", timeout=12000)
        _select_value_fast(page, "select:has(option[value='13'])", PLANTILLA_VALUE, "Plantilla", 2500)

    _select_value_fast(page, _xpath_select_following("Personalizado 1"), PERS1_VALUE, "P1", 2200)
    _select_value_fast(page, _xpath_select_following("Personalizado 2"), PERS2_VALUE, "P2", 2200)
    _select_value_fast(page, _xpath_select_following("Personalizado 3"), PERS3_VALUE, "P3", 2200)

def asegurar_reportes_listo(page, log, force_refresh: bool = False):
    if force_refresh:
        log("[NAV] Refrescando página (limpieza)...")
        page.goto(DEFAULT_URL_REPORTES, wait_until="domcontentloaded", timeout=DEFAULT_TIMEOUT_MS)
        page.wait_for_timeout(1500)
    else:
        try:
            if page.locator(SEL_BTN_GENERAR).first.is_visible(timeout=700):
                asegurar_fijos_formulario(page, log)
                return
        except:
            pass
        page.goto(DEFAULT_URL_REPORTES, wait_until="domcontentloaded", timeout=DEFAULT_TIMEOUT_MS)

    try:
        loc_reg = page.locator(SEL_CAT_REGISTROS).first
        if loc_reg.is_visible(timeout=8000):
            loc_reg.scroll_into_view_if_needed()
            loc_reg.click()
            page.wait_for_timeout(1000)
    except:
        pass

    page.wait_for_selector("text=Tipo Reporte", timeout=20000)
    asegurar_fijos_formulario(page, log)
    page.wait_for_selector(SEL_BTN_GENERAR, state="visible", timeout=20000)

def set_documento_paciente_or_fail(page, doc: str, log):
    doc = _norm_doc(doc)
    if not doc: raise RuntimeError("Documento vacío")

    candidates = ["select[name='filter']", "select#filter", _xpath_select_following("Paciente")]

    for sel in candidates:
        try:
            if page.locator(sel).first.count() > 0:
                ok = page.evaluate("""(args) => {
                    let s = document.querySelector(args.sel);
                    if (!s) return false;
                    let opts = Array.from(s.options);
                    let hit = opts.find(o => o.text.includes(args.doc) || o.value.includes(args.doc));
                    if(hit){
                        s.value = hit.value;
                        s.dispatchEvent(new Event('change', {bubbles:true}));
                        return true;
                    }
                    return false;
                }""", {"sel": sel, "doc": doc})
                if ok:
                    log(f"[FORM] Paciente {doc} seleccionado (JS).")
                    return
        except:
            pass

    try:
        s2 = page.locator("xpath=//*[contains(normalize-space(),'Paciente')]/following::span[contains(@class,'select2-selection')][1]").first
        if s2.count() > 0:
            s2.click()
            page.wait_for_timeout(500)
            search = page.locator("input.select2-search__field").first
            search.fill(doc)
            page.wait_for_timeout(1500)
            search.press("Enter")
            page.wait_for_timeout(1000)
            log(f"[FORM] Paciente {doc} seleccionado (Select2).")
            return
    except:
        pass

    raise RuntimeError(f"No se pudo seleccionar el paciente {doc}")

def generar_y_descargar_pdf(page, task: Task, output_dir: str, log) -> str:
    asegurar_fijos_formulario(page, log)
    set_documento_paciente_or_fail(page, task.doc, log)
    page.wait_for_timeout(500)

    _select_by_label_contains(page, _xpath_select_following("Valor 1"), task.bg, 12000)
    page.wait_for_timeout(200)
    _select_by_label_contains(page, _xpath_select_following("Valor 2"), task.bh, 12000)
    page.wait_for_timeout(200)
    _select_by_label_contains(page, _xpath_select_following("Valor 3"), task.bi, 12000)
    page.wait_for_timeout(500)

    fname = f"{task.nombre}-{task.doc}-{task.etiologia}-{task.bi}-{task.bh}-{task.clasificacion}.pdf"
    fname = _safe_filename(fname)

    out_path = os.path.join(output_dir, fname)
    c = 1
    root, ext = os.path.splitext(fname)
    while os.path.exists(out_path):
        out_path = os.path.join(output_dir, f"{root}_{c}{ext}")
        c += 1

    log(f"[PDF] Descargando: {os.path.basename(out_path)}")
    with page.expect_download(timeout=DEFAULT_TIMEOUT_MS) as dl_info:
        page.click(SEL_BTN_GENERAR)
    dl = dl_info.value
    dl.save_as(out_path)
    log("[OK] Descarga completada.")
    return out_path


# ==============================================================================
# 4) INTERFAZ + WORKER
# ==============================================================================
def detect_default_login_json(base_dir):
    p1 = os.path.join(base_dir, "CREDENCIALES", "gestiona_login.json")
    if os.path.exists(p1): return p1
    p2 = os.path.join(base_dir, "gestiona_login.json")
    if os.path.exists(p2): return p2
    return ""

def load_login_creds(path):
    d = _read_json_file(path)
    return d.get("user", ""), d.get("pass", "")

@dataclass
class RunConfig:
    mode: str
    file_path: str
    excel_sheet: str
    sheets_id: str
    sheets_tab: str
    credentials_path: str
    token_path: str
    output_dir: str
    fecha_desde: datetime
    fecha_hasta: datetime
    login_user: str
    login_pass: str
    headless: bool
    slowmo_ms: int
    login_wait_seconds: int

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Descargar Evoluciones CH - V.Final Fix")
        self.geometry("1000x850")

        self.style = ttk.Style()
        self.style.theme_use('clam')

        bg_color = "#f4f4f4"
        self.configure(bg=bg_color)

        AZUL_OSCURO = "#0056b3"
        BLANCO = "#ffffff"

        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabelframe", background=bg_color, relief="groove")
        self.style.configure("TLabelframe.Label", background=bg_color, foreground=AZUL_OSCURO, font=('Helvetica', 10, 'bold'))
        self.style.configure("TLabel", background=bg_color, font=('Helvetica', 9))
        self.style.configure("TRadiobutton", background=bg_color)
        self.style.configure("TCheckbutton", background=bg_color)

        self.style.configure("Blue.TButton",
                             font=('Helvetica', 10, 'bold'),
                             background=AZUL_OSCURO,
                             foreground=BLANCO,
                             borderwidth=1)
        self.style.map("Blue.TButton",
                       background=[('active', '#004494'), ('disabled', '#cccccc')],
                       foreground=[('disabled', '#666666')])

        self.cfg_mem = _read_json_file(APP_CONFIG_PATH)
        self.uiq = queue.Queue()
        self.worker = None

        self.stop_event = threading.Event()
        self.pause_event = threading.Event()

        self.mode = tk.StringVar(value=self.cfg_mem.get("mode", "archivo"))
        self.file_path = tk.StringVar(value="")
        self.excel_sheet = tk.StringVar(value=self.cfg_mem.get("excel_sheet", ""))

        self.sheets_id = tk.StringVar(value=self.cfg_mem.get("sheets_id", ""))
        self.sheets_tab = tk.StringVar(value=self.cfg_mem.get("sheets_tab", ""))
        self.credentials_path = tk.StringVar(value=self.cfg_mem.get("credentials_path", str(Path(BASE_DIR)/"CREDENCIALES"/"credentials.json")))
        self.token_path = tk.StringVar(value=self.cfg_mem.get("token_path", str(Path(BASE_DIR)/"CREDENCIALES"/"token_formulacion.json")))

        self.output_dir = tk.StringVar(value=self.cfg_mem.get("output_dir", str(Path(BASE_DIR)/"DESCARGAS")))
        self.last_dir = self.cfg_mem.get("last_dir", str(Path(BASE_DIR)))

        self.month_ym = tk.StringVar(value=_ym_now())
        d0, d1 = _month_range(self.month_ym.get())
        self.fecha_desde_str = tk.StringVar(value=d0)
        self.fecha_hasta_str = tk.StringVar(value=d1)

        def_json = detect_default_login_json(BASE_DIR)
        u, p = load_login_creds(def_json)
        self.login_user = tk.StringVar(value=self.cfg_mem.get("login_user", u))
        self.login_pass = tk.StringVar(value=self.cfg_mem.get("login_pass", p))

        self.headless = tk.BooleanVar(value=False)
        self.slowmo = tk.IntVar(value=0)

        try:
            default_wait = int(self.cfg_mem.get("login_wait_seconds", 5))
        except Exception:
            default_wait = 5
        if default_wait < 0:
            default_wait = 0
        self.login_wait_seconds = tk.IntVar(value=default_wait)

        self._build_ui()
        self.after(100, self._poll_ui_queue)

    def _build_ui(self):
        main = ttk.Frame(self, padding=15)
        main.pack(fill="both", expand=True)

        lf_origen = ttk.LabelFrame(main, text=" 1. Origen de Datos ", padding=10)
        lf_origen.pack(fill="x", pady=5)

        row = ttk.Frame(lf_origen)
        row.pack(fill="x")
        ttk.Radiobutton(row, text="Archivo Local (Excel/CSV)", variable=self.mode, value="archivo").pack(side="left", padx=10)
        ttk.Radiobutton(row, text="Google Sheets", variable=self.mode, value="sheets").pack(side="left", padx=10)

        row_file = ttk.Frame(lf_origen)
        row_file.pack(fill="x", pady=5)
        ttk.Label(row_file, text="Archivo:", width=10).pack(side="left")
        ttk.Entry(row_file, textvariable=self.file_path).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row_file, text="Examinar", style="Blue.TButton", command=self._pick_file).pack(side="left")

        row_sheet = ttk.Frame(lf_origen)
        row_sheet.pack(fill="x")
        ttk.Label(row_sheet, text="Hoja (Opc):", width=10).pack(side="left")
        ttk.Entry(row_sheet, textvariable=self.excel_sheet, width=20).pack(side="left", padx=5)

        row_gs = ttk.Frame(lf_origen)
        row_gs.pack(fill="x", pady=5)
        ttk.Label(row_gs, text="ID Sheets:", width=10).pack(side="left")
        ttk.Entry(row_gs, textvariable=self.sheets_id).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Label(row_gs, text="Tab:").pack(side="left")
        ttk.Entry(row_gs, textvariable=self.sheets_tab, width=15).pack(side="left", padx=5)

        row_cred = ttk.Frame(lf_origen)
        row_cred.pack(fill="x")
        ttk.Label(row_cred, text="Creds JSON:", width=10).pack(side="left")
        ttk.Entry(row_cred, textvariable=self.credentials_path).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row_cred, text="...", command=self._pick_creds).pack(side="left")

        lf_fechas = ttk.LabelFrame(main, text=" 2. Filtro de Fechas (Columna B) ", padding=10)
        lf_fechas.pack(fill="x", pady=5)

        row_f = ttk.Frame(lf_fechas)
        row_f.pack(fill="x")

        vals = []
        hoy = datetime.now()
        y, m = hoy.year, hoy.month
        for _ in range(12):
            vals.append(f"{y:04d}-{m:02d}")
            m -= 1
            if m == 0:
                m = 12
                y -= 1

        ttk.Label(row_f, text="Mes Rápido:").pack(side="left")
        cb = ttk.Combobox(row_f, textvariable=self.month_ym, values=vals, width=10, state="readonly")
        cb.pack(side="left", padx=5)
        cb.bind("<<ComboboxSelected>>", lambda e: self._apply_month_defaults())
        ttk.Button(row_f, text="Aplicar", command=self._apply_month_defaults).pack(side="left", padx=5)

        ttk.Separator(row_f, orient="vertical").pack(side="left", fill="y", padx=10)

        ttk.Label(row_f, text="Desde:").pack(side="left")
        ttk.Entry(row_f, textvariable=self.fecha_desde_str, width=12).pack(side="left", padx=2)
        ttk.Button(row_f, text="📅", width=3, command=lambda: self._open_calendar(self.fecha_desde_str)).pack(side="left")

        ttk.Label(row_f, text="Hasta:").pack(side="left", padx=5)
        ttk.Entry(row_f, textvariable=self.fecha_hasta_str, width=12).pack(side="left", padx=2)
        ttk.Button(row_f, text="📅", width=3, command=lambda: self._open_calendar(self.fecha_hasta_str)).pack(side="left")

        lf_run = ttk.LabelFrame(main, text=" 3. Ejecución ", padding=10)
        lf_run.pack(fill="x", pady=5)

        row_u = ttk.Frame(lf_run)
        row_u.pack(fill="x", pady=2)
        ttk.Label(row_u, text="Usuario:", width=10).pack(side="left")
        ttk.Entry(row_u, textvariable=self.login_user, width=25).pack(side="left", padx=5)
        ttk.Label(row_u, text="Pass:").pack(side="left")
        ttk.Entry(row_u, textvariable=self.login_pass, show="*", width=25).pack(side="left", padx=5)

        ttk.Label(row_u, text="Espera login (s):").pack(side="left", padx=(10, 2))
        sp = ttk.Spinbox(row_u, from_=0, to=60, width=5, textvariable=self.login_wait_seconds)
        sp.pack(side="left")

        row_d = ttk.Frame(lf_run)
        row_d.pack(fill="x", pady=5)
        ttk.Label(row_d, text="Guardar en:", width=10).pack(side="left")
        ttk.Entry(row_d, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row_d, text="...", command=self._pick_dir).pack(side="left")

        row_opt = ttk.Frame(lf_run)
        row_opt.pack(fill="x")

        self.chk_headless = ttk.Checkbutton(
            row_opt,
            text="Modo Headless (Navegador oculto)",
            variable=self.headless
        )
        self.chk_headless.pack(side="left")
        self.headless.set(False)
        self.chk_headless.state(["disabled"])

        self.btn_start = ttk.Button(main, text="INICIAR PROCESO", style="Blue.TButton", command=self._start)
        self.btn_start.pack(fill="x", pady=10, ipady=8)

        ctrl = ttk.Frame(main)
        ctrl.pack(fill="x", pady=5)

        self.btn_pause = ttk.Button(ctrl, text="PAUSAR", style="Blue.TButton", command=self._toggle_pause, state="disabled")
        self.btn_pause.pack(side="left", fill="x", expand=True, padx=(0, 5), ipady=6)

        self.btn_stop = ttk.Button(ctrl, text="DETENER", style="Blue.TButton", command=self._stop_run, state="disabled")
        self.btn_stop.pack(side="left", fill="x", expand=True, padx=(5, 0), ipady=6)

        self.log_widget = ScrolledText(main, height=12)
        self.log_widget.pack(fill="both", expand=True)

    def _pick_file(self):
        p = filedialog.askopenfilename(initialdir=self.last_dir, filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv")])
        if p:
            self.file_path.set(p)
            self.last_dir = os.path.dirname(p)

    def _pick_dir(self):
        p = filedialog.askdirectory(initialdir=self.output_dir.get())
        if p: self.output_dir.set(p)

    def _pick_creds(self):
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if p: self.credentials_path.set(p)

    def _apply_month_defaults(self):
        d0, d1 = _month_range(self.month_ym.get())
        self.fecha_desde_str.set(d0)
        self.fecha_hasta_str.set(d1)

    def _toggle_pause(self):
        if not self.worker or not self.worker.is_alive():
            return
        if not self.pause_event.is_set():
            self.pause_event.set()
            self.btn_pause.config(text="REANUDAR")
            self._log_msg("[PAUSA] Pausado por el usuario.")
        else:
            self.pause_event.clear()
            self.btn_pause.config(text="PAUSAR")
            self._log_msg("[PAUSA] Reanudado.")

    def _stop_run(self):
        if not self.worker or not self.worker.is_alive():
            return
        self.stop_event.set()
        self.pause_event.clear()
        self.btn_pause.config(text="PAUSAR")
        self._log_msg("[STOP] Solicitud de detención recibida. Se detendrá al terminar el paso actual.")

    def _open_calendar(self, target_var):
        top = tk.Toplevel(self)
        top.title("Seleccionar Fecha")
        top.resizable(False, False)

        try:
            curr = _parse_date(target_var.get()) or datetime.now()
            year, month = curr.year, curr.month
        except:
            now = datetime.now()
            year, month = now.year, now.month

        y_var = tk.IntVar(value=year)
        m_var = tk.IntVar(value=month)

        frm_head = ttk.Frame(top)
        frm_head.pack()
        ttk.Spinbox(frm_head, from_=2000, to=2100, textvariable=y_var, width=5).pack(side="left")
        ttk.Spinbox(frm_head, from_=1, to=12, textvariable=m_var, width=3).pack(side="left")

        frm_cal = ttk.Frame(top)
        frm_cal.pack()

        def refresh_cal(*args):
            for w in frm_cal.winfo_children(): w.destroy()
            cal = pycal.monthcalendar(y_var.get(), m_var.get())
            days = ["Lu","Ma","Mi","Ju","Vi","Sa","Do"]
            for c, d in enumerate(days):
                ttk.Label(frm_cal, text=d, font="TkFixedFont").grid(row=0, column=c)

            for r, week in enumerate(cal):
                for c, day in enumerate(week):
                    if day == 0: continue
                    btn = ttk.Button(frm_cal, text=str(day), width=3, command=lambda d=day: select_day(d))
                    btn.grid(row=r+1, column=c)

        def select_day(d):
            target_var.set(f"{y_var.get():04d}-{m_var.get():02d}-{d:02d}")
            top.destroy()

        y_var.trace_add("write", lambda *_: refresh_cal())
        m_var.trace_add("write", lambda *_: refresh_cal())
        refresh_cal()

        top.update_idletasks()
        x = self.winfo_rootx() + 50
        y = self.winfo_rooty() + 50
        top.geometry(f"+{x}+{y}")

    def _log_msg(self, msg):
        self.uiq.put(("log", msg))

    def _poll_ui_queue(self):
        try:
            while True:
                kind, data = self.uiq.get_nowait()
                if kind == "log":
                    self.log_widget.insert("end", data + "\n")
                    self.log_widget.see("end")
                elif kind == "done":
                    messagebox.showinfo("Finalizado", "Proceso completado.")
                    self.btn_start.config(state="normal")
                    self.btn_pause.config(state="disabled", text="PAUSAR")
                    self.btn_stop.config(state="disabled")
                elif kind == "error":
                    messagebox.showerror("Error", data)
                    self.btn_start.config(state="normal")
                    self.btn_pause.config(state="disabled", text="PAUSAR")
                    self.btn_stop.config(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_ui_queue)

    def _start(self):
        if self.worker and self.worker.is_alive(): return

        data = {
            "mode": self.mode.get(),
            "excel_sheet": self.excel_sheet.get(),
            "sheets_id": self.sheets_id.get(),
            "sheets_tab": self.sheets_tab.get(),
            "credentials_path": self.credentials_path.get(),
            "token_path": self.token_path.get(),
            "output_dir": self.output_dir.get(),
            "login_user": self.login_user.get(),
            "login_pass": self.login_pass.get(),
            "login_wait_seconds": int(self.login_wait_seconds.get() or 0),
            "last_dir": self.last_dir,
            "headless": False
        }
        _write_json_file(APP_CONFIG_PATH, data)

        f_start = _parse_date_ui_start(self.fecha_desde_str.get())
        f_end = _parse_date_ui_end(self.fecha_hasta_str.get())
        if not f_start or not f_end:
            messagebox.showerror("Error", "Fechas inválidas.")
            return

        try:
            wait_s = int(self.login_wait_seconds.get())
        except Exception:
            wait_s = 5
        if wait_s < 0:
            wait_s = 0
        if wait_s > 600:
            wait_s = 600

        cfg = RunConfig(
            mode=self.mode.get(),
            file_path=self.file_path.get(),
            excel_sheet=self.excel_sheet.get(),
            sheets_id=_extract_spreadsheet_id(self.sheets_id.get()),
            sheets_tab=self.sheets_tab.get(),
            credentials_path=self.credentials_path.get(),
            token_path=self.token_path.get(),
            output_dir=self.output_dir.get(),
            fecha_desde=f_start,
            fecha_hasta=f_end,
            login_user=self.login_user.get(),
            login_pass=self.login_pass.get(),
            headless=False,
            slowmo_ms=self.slowmo.get(),
            login_wait_seconds=wait_s
        )

        if not cfg.login_user or not cfg.login_pass:
            messagebox.showerror("Error", "Faltan credenciales.")
            return

        self.stop_event.clear()
        self.pause_event.clear()

        self.btn_start.config(state="disabled")
        self.btn_pause.config(state="normal", text="PAUSAR")
        self.btn_stop.config(state="normal")

        self.log_widget.delete("1.0", "end")
        self._log_msg("--- INICIANDO ---")

        self.worker = threading.Thread(target=self._worker_run, args=(cfg,), daemon=True)
        self.worker.start()

    def _worker_run(self, cfg: RunConfig):
        try:
            tasks = []
            if cfg.mode == "archivo":
                if cfg.file_path.lower().endswith(".csv"):
                    tasks = leer_csv_y_deduplicar(cfg.file_path, cfg.fecha_desde, cfg.fecha_hasta, self._log_msg)
                else:
                    tasks = leer_excel_y_deduplicar(cfg.file_path, cfg.excel_sheet, cfg.fecha_desde, cfg.fecha_hasta, self._log_msg)
            else:
                if not cfg.sheets_id: raise ValueError("Falta ID de Sheets")
                tasks = leer_sheets_y_deduplicar(cfg.sheets_id, cfg.sheets_tab, cfg.credentials_path, cfg.token_path,
                                                cfg.fecha_desde, cfg.fecha_hasta, self._log_msg)

            if not tasks:
                self._log_msg("No se encontraron registros válidos.")
                self.uiq.put(("done", None))
                return

            self._run_playwright(cfg, tasks)
            self.uiq.put(("done", None))

        except Exception as e:
            self.uiq.put(("error", str(e)))

    def _run_playwright(self, cfg: RunConfig, tasks: List[Task]):
        self._log_msg(f"[PW] Lanzando navegador (Headless={cfg.headless})...")

        try:
            os.makedirs(cfg.output_dir, exist_ok=True)
        except:
            pass

        resultados: List[Dict[str, Any]] = []

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=cfg.headless, slow_mo=cfg.slowmo_ms)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            try:
                do_login(page, cfg.login_user, cfg.login_pass, self._log_msg, pre_enter_wait_seconds=cfg.login_wait_seconds)
                asegurar_reportes_listo(page, self._log_msg, force_refresh=True)

                ok_count = 0
                fail_count = 0
                total = len(tasks)

                for i, t in enumerate(tasks, start=1):
                    if self.stop_event.is_set():
                        self._log_msg("[STOP] Proceso detenido por el usuario.")
                        break

                    paused_logged = False
                    while self.pause_event.is_set():
                        if not paused_logged:
                            self._log_msg("[PAUSA] En pausa... (esperando reanudar)")
                            paused_logged = True
                        if self.stop_event.is_set():
                            self._log_msg("[STOP] Proceso detenido por el usuario (desde pausa).")
                            break
                        page.wait_for_timeout(300)

                    if self.stop_event.is_set():
                        break

                    self._log_msg(f"\n--- Tarea {i}/{total} | Paciente: {t.doc} ---")
                    self._log_msg(f"    Ubicación: {t.bg} - {t.bh} - {t.bi}")
                    self._log_msg(f"    Fecha Evol: {t.fecha_evolucion}")

                    try:
                        self._log_msg("[CICLO] Refrescando página...")
                        asegurar_reportes_listo(page, self._log_msg, force_refresh=True)

                        pdf_path = generar_y_descargar_pdf(page, t, cfg.output_dir, self._log_msg)
                        ok_count += 1

                        resultados.append({
                            "n": i,
                            "estado": "OK",
                            "error": "",
                            "archivo_pdf": os.path.basename(pdf_path) if pdf_path else "",
                            "ruta_pdf": pdf_path or "",
                            "fecha_evolucion": t.fecha_evolucion.strftime("%Y-%m-%d %H:%M:%S") if t.fecha_evolucion else "",
                            "doc": t.doc,
                            "nombre": t.nombre,
                            "etiologia": t.etiologia,
                            "bg": t.bg,
                            "bh": t.bh,
                            "bi": t.bi,
                            "clasificacion": t.clasificacion,
                            "fila_origen": t.row,
                        })

                    except Exception as e:
                        fail_count += 1
                        self._log_msg(f"[ERROR] {e}")

                        resultados.append({
                            "n": i,
                            "estado": "ERROR",
                            "error": str(e),
                            "archivo_pdf": "",
                            "ruta_pdf": "",
                            "fecha_evolucion": t.fecha_evolucion.strftime("%Y-%m-%d %H:%M:%S") if t.fecha_evolucion else "",
                            "doc": t.doc,
                            "nombre": t.nombre,
                            "etiologia": t.etiologia,
                            "bg": t.bg,
                            "bh": t.bh,
                            "bi": t.bi,
                            "clasificacion": t.clasificacion,
                            "fila_origen": t.row,
                        })

                        try:
                            page.reload()
                        except:
                            pass

                self._log_msg(f"\n[RESUMEN] Exitosos: {ok_count} | Fallidos: {fail_count}")

                try:
                    escribir_reporte_excel(cfg.output_dir, cfg, resultados, self._log_msg)
                except Exception as rep_e:
                    self._log_msg(f"[REPORTE][ERROR] {rep_e}")

            finally:
                context.close()
                browser.close()

if __name__ == "__main__":
    App().mainloop()

import pyautogui
import cv2
import numpy as np
import pandas as pd
import time
import os
import json
import re
import unicodedata
from datetime import datetime
import pygetwindow as gw
import keyboard
import warnings

# Suprimir warnings de PyTorch/torchvision
warnings.filterwarnings('ignore', category=UserWarning, module='torch.utils.data.dataloader')
warnings.filterwarnings('ignore', category=UserWarning, module='torchvision')
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'  # Suprimir logs de tensorflow si estÃ¡ instalado

# Usar exclusivamente EasyOCR (requisito: no usar Tesseract)
try:
    import easyocr
    USE_EASYOCR = True
except Exception:
    USE_EASYOCR = False

log_callback = None

EXCEL_PATH = "precios_market.xlsx"
GOOGLE_SHEET_URL = os.getenv(
    "GOOGLE_SHEET_URL",
    "https://docs.google.com/spreadsheets/d/1P-n6hKBvJ18YsaHEfEzVxKfU3IMVpfWzcsIKzSjUmtQ/edit?gid=1152689078#gid=1152689078",
)
GOOGLE_SHEET_GID = os.getenv("GOOGLE_SHEET_GID", "1152689078")
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "google_service_account.json")
GOOGLE_SYNC_ENABLED = os.getenv("GOOGLE_SYNC_ENABLED", "1").strip().lower() not in ("0", "false", "no", "off")

# Flags for optional OCR engines (Paddle/pytesseract not used by default)
USE_PADDLEOCR = False
USE_PYTESSERACT = False
paddle_ocr = None
reader = None

try:
    import gspread
    from gspread.cell import Cell
    GOOGLE_SHEETS_AVAILABLE = True
except Exception:
    gspread = None
    Cell = None
    GOOGLE_SHEETS_AVAILABLE = False

def set_log_callback(callback):
    """Establece el callback para logs (para GUI)"""
    global log_callback
    log_callback = callback

def log_message(msg):
    """Registra un mensaje. Usa callback si estÃ¡ disponible, sino usa print"""
    if log_callback:
        log_callback(msg)
    else:
        print(msg)

def normalize_text(value):
    text = str(value or "").strip().casefold()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return text


def normalize_item_name(item_name):
    text = normalize_text(item_name)
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_item_spacing(item_text):
    text = str(item_text or "").strip()
    if not text:
        return ""
    text = text.replace("\u2019", "'").replace("`", "'").replace("\u00B4", "'")
    text = re.sub(r"(?<=\w)\s*'\s*(?=\w)", "'", text)
    text = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    # Correccion OCR: "Trolls Spiritstone" -> "Troll's Spiritstone".
    text = re.sub(r"\bTrolls\s+Spiritstone\b", "Troll's Spiritstone", text, flags=re.IGNORECASE)
    return text
def bbox_iou(b1, b2):
    x1_min, y1_min = float(b1[0][0]), float(b1[0][1])
    x1_max, y1_max = float(b1[2][0]), float(b1[2][1])
    x2_min, y2_min = float(b2[0][0]), float(b2[0][1])
    x2_max, y2_max = float(b2[2][0]), float(b2[2][1])

    inter_x_min = max(x1_min, x2_min)
    inter_y_min = max(y1_min, y2_min)
    inter_x_max = min(x1_max, x2_max)
    inter_y_max = min(y1_max, y2_max)
    inter_w = max(0.0, inter_x_max - inter_x_min)
    inter_h = max(0.0, inter_y_max - inter_y_min)
    inter_area = inter_w * inter_h
    if inter_area <= 0:
        return 0.0

    a1 = max(0.0, (x1_max - x1_min) * (y1_max - y1_min))
    a2 = max(0.0, (x2_max - x2_min) * (y2_max - y2_min))
    union = a1 + a2 - inter_area
    if union <= 0:
        return 0.0
    return inter_area / union


def item_candidate_score(text, conf):
    score = float(conf)
    if re.search(r"[A-Za-z]['\u2019\u00B4`][A-Za-z]", text):
        score += 0.10
    if re.search(r"\d", text):
        score -= 0.12
    if re.search(r"[?]", text):
        score -= 0.08
    if len(text.strip()) < 4:
        score -= 0.05
    return score


def normalize_price_text(price_text):
    raw = str(price_text or "").strip()
    number = parse_number(raw)
    if number is None:
        return raw
    text = f"{number:.8f}".rstrip("0").rstrip(".")
    return text if text else "0"


def normalize_number_for_sheet(value):
    number = parse_number(value)
    if number is None:
        return ""
    if float(number).is_integer():
        return str(int(number))
    return str(number)


def parse_number(value):
    raw = str(value or "").strip()
    if not raw:
        return None

    # Mantener solo digitos y separadores comunes.
    candidate = re.sub(r"[^0-9,.\-]", "", raw)
    if not candidate:
        return None

    # Si hay ambos separadores, usar el ultimo como decimal.
    if "," in candidate and "." in candidate:
        last_comma = candidate.rfind(",")
        last_dot = candidate.rfind(".")
        decimal_sep = "," if last_comma > last_dot else "."
        thousand_sep = "." if decimal_sep == "," else ","
        normalized = candidate.replace(thousand_sep, "")
        normalized = normalized.replace(decimal_sep, ".")
    elif "," in candidate:
        # Unica coma: si tiene 1-2 decimales, tratarla como decimal; si no, miles.
        if candidate.count(",") == 1 and len(candidate.split(",")[1]) <= 2:
            normalized = candidate.replace(",", ".")
        else:
            normalized = candidate.replace(",", "")
    elif "." in candidate:
        # Unico punto: si tiene 1-2 decimales, decimal; con varios puntos, miles.
        if candidate.count(".") > 1:
            normalized = candidate.replace(".", "")
        else:
            normalized = candidate
    else:
        normalized = candidate

    try:
        return float(normalized)
    except Exception:
        digits = "".join(ch for ch in candidate if ch.isdigit())
        if digits:
            try:
                return float(digits)
            except Exception:
                return None
    return None


def parse_sheet_id_from_url(url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", str(url or ""))
    return match.group(1) if match else ""


def get_column_letter(col_idx):
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result or "A"


def detect_header_map(headers):
    aliases = {
        "item": {"item", "items", "nombre", "producto", "articulo"},
        "price_wemix": {"price wemix", "precio wemix", "wemix", "precio wemix market", "price_wemix", "precio_wemix"},
        "price_diamantes": {
            "price diamantes",
            "precio diamantes",
            "diamantes",
            "diamonds",
            "price diamonds",
            "precio_diamantes",
            "price_diamantes",
        },
        "sales": {"sales", "venta", "ventas", "cantidad sales", "total sales", "sales_total"},
        "last_update": {"last update", "updated", "timestamp", "fecha", "ultima actualizacion", "last_update"},
    }
    mapped = {}
    for idx, header in enumerate(headers):
        normalized = normalize_text(header)
        for key, valid in aliases.items():
            if normalized in valid and key not in mapped:
                mapped[key] = idx
    return mapped


def rows_to_price_map(items):
    price_map = {}
    skipped_zero = 0
    for row in items or []:
        if not isinstance(row, dict):
            continue
        item_name = str(row.get("item", "")).strip()
        price = str(row.get("price", "")).strip()
        sales = str(row.get("sales", "")).strip()
        if not item_name or not price:
            continue

        price_number = parse_number(price)
        sales_number = parse_number(sales)
        if price_number == 0 or sales_number == 0:
            skipped_zero += 1
            continue

        key = normalize_item_name(item_name)
        if not key:
            continue
        price_map[key] = {
            "item": item_name,
            "price": normalize_price_text(price),
            "sales": sales,
        }
    if skipped_zero:
        log_message(f"[INFO] Filtrados por precio/sales en 0: {skipped_zero}")
    return price_map


def get_worksheet_from_gid(spreadsheet, gid):
    try:
        gid_int = int(str(gid).strip())
    except Exception:
        gid_int = None
    if gid_int is None:
        return spreadsheet.sheet1
    for ws in spreadsheet.worksheets():
        if ws.id == gid_int:
            return ws
    return spreadsheet.sheet1


def ensure_last_update_next_to_sales(spreadsheet, worksheet, headers):
    header_map = detect_header_map(headers)
    if "sales" not in header_map or "last_update" not in header_map:
        return headers

    sales_idx = header_map["sales"]
    update_idx = header_map["last_update"]
    if update_idx == sales_idx + 1:
        return headers

    destination_index = sales_idx + 1 if update_idx > sales_idx else sales_idx
    spreadsheet.batch_update(
        {
            "requests": [
                {
                    "moveDimension": {
                        "source": {
                            "sheetId": worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": update_idx,
                            "endIndex": update_idx + 1,
                        },
                        "destinationIndex": destination_index,
                    }
                }
            ]
        }
    )
    log_message("[OK] Columna 'last_update' movida junto a 'sales'")
    return worksheet.row_values(1)


def sync_to_google_sheet(items_wemix, items_diamantes):
    if not GOOGLE_SYNC_ENABLED:
        log_message("[INFO] Sincronizacion Google Sheets desactivada (GOOGLE_SYNC_ENABLED=0)")
        return False
    if not GOOGLE_SHEETS_AVAILABLE:
        log_message("[WARN] gspread/google-auth no disponibles. Ejecuta: pip install gspread google-auth")
        return False
    if not os.path.exists(GOOGLE_SERVICE_ACCOUNT_FILE):
        log_message(f"[WARN] No existe credencial Google: {GOOGLE_SERVICE_ACCOUNT_FILE}")
        return False

    sheet_id = parse_sheet_id_from_url(GOOGLE_SHEET_URL)
    if not sheet_id:
        log_message("[WARN] GOOGLE_SHEET_URL no es valido")
        return False

    wemix_map = rows_to_price_map(items_wemix)
    diamantes_map = rows_to_price_map(items_diamantes)
    merged = {}
    for key, data in wemix_map.items():
        merged[key] = {
            "item": data["item"],
            "price_wemix": data["price"],
            "price_diamantes": "",
            "sales": parse_number(data.get("sales")),
        }
    for key, data in diamantes_map.items():
        entry = merged.setdefault(
            key,
            {"item": data["item"], "price_wemix": "", "price_diamantes": "", "sales": 0.0},
        )
        if not entry.get("item"):
            entry["item"] = data["item"]
        entry["price_diamantes"] = data["price"]
        current_sales = entry.get("sales")
        if current_sales is None:
            current_sales = 0.0
        entry["sales"] = float(current_sales) + float(parse_number(data.get("sales")) or 0.0)

    if not merged:
        log_message("[WARN] No hay datos para sincronizar en Google Sheets")
        return False

    try:
        gc = gspread.service_account(filename=GOOGLE_SERVICE_ACCOUNT_FILE)
        spreadsheet = gc.open_by_key(sheet_id)
        ws = get_worksheet_from_gid(spreadsheet, GOOGLE_SHEET_GID)

        values = ws.get_all_values()
        headers = values[0] if values else []
        default_headers = ["item", "price_wemix", "price_diamantes", "sales", "last_update"]

        if not headers:
            headers = default_headers[:]
            ws.update("A1:E1", [headers], value_input_option="USER_ENTERED")

        header_map = detect_header_map(headers)
        changed_headers = False
        for required_col in default_headers:
            if required_col not in header_map:
                headers.append(required_col)
                header_map[required_col] = len(headers) - 1
                changed_headers = True

        if changed_headers:
            end_col = get_column_letter(len(headers))
            ws.update(f"A1:{end_col}1", [headers], value_input_option="USER_ENTERED")
        headers = ws.row_values(1)
        headers = ensure_last_update_next_to_sales(spreadsheet, ws, headers)
        header_map = detect_header_map(headers)

        all_values = ws.get_all_values()
        existing_rows = all_values[1:] if len(all_values) > 1 else []

        item_col = header_map["item"]
        wemix_col = header_map["price_wemix"]
        diam_col = header_map["price_diamantes"]
        sales_col = header_map["sales"]
        update_col = header_map["last_update"]

        existing_by_item = {}
        empty_item_rows = []
        for idx, row in enumerate(existing_rows, start=2):
            current_item = row[item_col] if item_col < len(row) else ""
            key = normalize_item_name(current_item)
            if key:
                existing_by_item[key] = (idx, row)
            else:
                empty_item_rows.append(idx)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cells_to_update = []
        rows_to_append = []

        for key, data in merged.items():
            if key in existing_by_item:
                row_number, current_row = existing_by_item[key]
                updates = []
                if str(data.get("price_wemix", "")).strip():
                    updates.append((wemix_col, data.get("price_wemix", "")))
                if str(data.get("price_diamantes", "")).strip():
                    updates.append((diam_col, data.get("price_diamantes", "")))
                updates.append((sales_col, normalize_number_for_sheet(data.get("sales"))))
                updates.append((update_col, now))
                for col_idx, new_value in updates:
                    current_value = current_row[col_idx] if col_idx < len(current_row) else ""
                    if str(current_value).strip() != str(new_value).strip():
                        cells_to_update.append(Cell(row=row_number, col=col_idx + 1, value=str(new_value)))
            else:
                if empty_item_rows:
                    row_number = empty_item_rows.pop(0)
                    inserts = [
                        (item_col, data.get("item", "")),
                        (wemix_col, data.get("price_wemix", "")),
                        (diam_col, data.get("price_diamantes", "")),
                        (sales_col, normalize_number_for_sheet(data.get("sales"))),
                        (update_col, now),
                    ]
                    for col_idx, new_value in inserts:
                        cells_to_update.append(Cell(row=row_number, col=col_idx + 1, value=str(new_value)))
                else:
                    new_row = [""] * len(headers)
                    new_row[item_col] = data.get("item", "")
                    new_row[wemix_col] = data.get("price_wemix", "")
                    new_row[diam_col] = data.get("price_diamantes", "")
                    new_row[sales_col] = normalize_number_for_sheet(data.get("sales"))
                    new_row[update_col] = now
                    rows_to_append.append(new_row)

        if cells_to_update:
            ws.update_cells(cells_to_update, value_input_option="USER_ENTERED")
        if rows_to_append:
            ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")

        log_message(
            f"[OK] Google Sheets sincronizado. Actualizados: {len(cells_to_update)} celdas, Nuevos items: {len(rows_to_append)}"
        )
        return True
    except Exception as e:
        log_message(f"[ERROR] Fallo sincronizando Google Sheets: {e}")
        return False


def build_rows_for_tipo(items, tipo, timestamp=None):
    """Construye filas tabulares para persistencia en Excel."""
    now = timestamp or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    if not items:
        return rows

    if isinstance(items[0], dict):
        for row in items:
            formatted = {"tipo": tipo, "timestamp": now}
            formatted.update(row)
            rows.append(formatted)
    else:
        for row in items:
            rows.append(
                {
                    "tipo": tipo,
                    "contenido": row,
                    "timestamp": now,
                }
            )
    return rows


def append_capture_to_excel(items, tipo, timestamp=None, excel_path=EXCEL_PATH):
    """Agrega una captura procesada al Excel sin sobrescribir historial."""
    new_rows = build_rows_for_tipo(items, tipo, timestamp=timestamp)
    if not new_rows:
        log_message(f"[WARN] No hay filas para guardar en Excel ({tipo})")
        return 0

    new_df = pd.DataFrame(new_rows)
    if os.path.exists(excel_path):
        try:
            current_df = pd.read_excel(excel_path)
            renamed_items = 0
            if "item" in current_df.columns and "item" in new_df.columns:
                latest_item_by_key = {}
                for raw_item in new_df["item"].tolist():
                    item_name = str(raw_item or "").strip()
                    key = normalize_item_name(item_name)
                    if key and item_name:
                        latest_item_by_key[key] = item_name

                if latest_item_by_key:
                    for idx in current_df.index:
                        current_item = str(current_df.at[idx, "item"] or "").strip()
                        key = normalize_item_name(current_item)
                        if not key or key not in latest_item_by_key:
                            continue
                        updated_item = latest_item_by_key[key]
                        if current_item != updated_item:
                            current_df.at[idx, "item"] = updated_item
                            renamed_items += 1

            if renamed_items:
                log_message(f"[INFO] Items renombrados en Excel por mayusculas/minusculas: {renamed_items}")
            output_df = pd.concat([current_df, new_df], ignore_index=True)
        except Exception as exc:
            log_message(f"[WARN] No se pudo leer {excel_path}. Se re-crea. Error: {exc}")
            output_df = new_df
    else:
        output_df = new_df

    output_df.to_excel(excel_path, index=False)
    log_message(f"[OK] Captura {tipo} guardada en Excel. Filas agregadas: {len(new_rows)}")
    return len(new_rows)


def esperar_f12(mensaje):
    """Espera a que el usuario presione F12"""
    log_message("")
    log_message(mensaje)
    log_message("[Esperando F12...]")
    keyboard.wait('f12')
    log_message("[OK] Preparando captura...")
    time.sleep(1)  # Esperar a que se estabilice
    log_message("[OK] Capturando...")

def obtener_ventana_juego():
    """Detecta la ventana del juego 'YmirGL' (elige la ventana Ymir mÃ¡s grande)."""
    try:
        wins = [w for w in gw.getAllWindows() if w and w.title]
    except Exception:
        return None

    # Filtrar ventanas que contienen 'ymir' o 'ymirgl' en el tÃ­tulo (case-insensitive)
    candidatas = [w for w in wins if ('ymir' in w.title.lower() or 'ymirgl' in w.title.lower())]
    if not candidatas:
        return None

    # Preferir ventanas que contengan 'ymirgl' exacto
    exact_ymirgl = [w for w in candidatas if 'ymirgl' in w.title.lower()]
    if exact_ymirgl:
        candidatas = exact_ymirgl

    # Excluir ventanas con tÃ­tulos de editores o utilidades comunes
    exclude_terms = ['visual studio', 'visual studio code', 'vscode', 'code', 'notepad', 'explorer', 'onenote', 'word', 'excel', 'powerpoint', 'sublime', 'pycharm']
    filtered = [w for w in candidatas if not any(t in w.title.lower() for t in exclude_terms)]
    if filtered:
        candidatas = filtered

    # Elegir la ventana con mayor Ã¡rea (ancho*alto)
    def area(win):
        try:
            return (getattr(win, 'width', 0) or 0) * (getattr(win, 'height', 0) or 0)
        except Exception:
            return 0

    ventana = max(candidatas, key=area)
    return ventana


def preprocesar(img, for_numeric=False):
    """Preprocesado para OCR.
    Devuelve (imagen_preprocesada, factor_de_escala).
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    if for_numeric:
        # En columnas numericas, subir contraste + escalar mejora digitos aislados.
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        gray = clahe.apply(gray)
        gray = cv2.resize(gray, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)
        return gray, 4.0
    # Para items: resize + CLAHE + sharpen.
    gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    gray = clahe.apply(gray)
    kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
    sharp = cv2.filter2D(gray, -1, kernel)
    return sharp, 2.0


def load_rois_config(cfg_path='rois.json'):
    """Carga configuracion de ROIs desde rois.json con defaults basicos."""
    rois_config = {
        'item': {'x': [0.00, 0.65], 'y': [0.0, 1.0]},
        'price': {'x': [0.65, 0.90], 'y': [0.0, 1.0]},
        'sales': {'x': [0.90, 1.0], 'y': [0.0, 1.0]},
    }
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            for col in ['item', 'price', 'sales']:
                if col in cfg:
                    col_cfg = cfg[col]
                    if isinstance(col_cfg, dict) and 'x' in col_cfg:
                        rois_config[col] = col_cfg
                    elif isinstance(col_cfg, (list, tuple)) and len(col_cfg) >= 2:
                        rois_config[col]['x'] = list(col_cfg)
            for col, col_cfg in cfg.items():
                if col in rois_config:
                    continue
                if isinstance(col_cfg, dict) and 'x' in col_cfg:
                    rois_config[col] = col_cfg
                elif isinstance(col_cfg, (list, tuple)) and len(col_cfg) >= 2:
                    rois_config[col] = {'x': list(col_cfg), 'y': [0.0, 1.0]}
        except Exception as e:
            log_message(f'[WARN] Error cargando rois.json: {e}')
    return rois_config


def _roi_rect_from_rel(img_shape, roi_cfg):
    h_full, w_full = img_shape[:2]
    x0_rel, x1_rel = roi_cfg['x'][0], roi_cfg['x'][1]
    y0_rel, y1_rel = roi_cfg['y'][0], roi_cfg['y'][1]
    x0 = max(0, min(w_full - 1, int(x0_rel * w_full)))
    x1 = max(1, min(w_full, int(x1_rel * w_full)))
    y0 = max(0, min(h_full - 1, int(y0_rel * h_full)))
    y1 = max(1, min(h_full, int(y1_rel * h_full)))
    if x1 <= x0:
        x1 = min(w_full, x0 + 1)
    if y1 <= y0:
        y1 = min(h_full, y0 + 1)
    return x0, y0, x1, y1


def detectar_tipo_market(img):
    """Detecta si el tab activo es WEMIX o DIAMANTES via color del switch."""
    rois_config = load_rois_config('rois.json')
    switch_cfg = rois_config.get('tab_switch')
    if not switch_cfg:
        switch_cfg = {'x': [0.72, 0.99], 'y': [0.00, 0.16]}

    x0, y0, x1, y1 = _roi_rect_from_rel(img.shape, switch_cfg)
    roi = img[y0:y1, x0:x1]
    if roi.size == 0:
        log_message('[WARN] ROI tab_switch vacia; no se pudo detectar tab activo')
        return 'UNKNOWN'

    hsv = cv2.cvtColor(roi, cv2.COLOR_BGR2HSV)
    sat_mask = (hsv[:, :, 1] >= 45) & (hsv[:, :, 2] >= 35)
    valid_pixels = int(np.count_nonzero(sat_mask))
    if valid_pixels <= 0:
        log_message('[WARN] ROI tab_switch sin color suficiente; no se pudo detectar tab activo')
        return 'UNKNOWN'

    hue = hsv[:, :, 0]
    diam_mask = sat_mask & (hue >= 80) & (hue <= 120)
    wemix_mask = sat_mask & (hue >= 125) & (hue <= 170)

    diam_ratio = float(np.count_nonzero(diam_mask)) / float(valid_pixels)
    wemix_ratio = float(np.count_nonzero(wemix_mask)) / float(valid_pixels)

    detected = 'UNKNOWN'
    if diam_ratio > wemix_ratio * 1.15 and diam_ratio >= 0.08:
        detected = 'DIAMANTES'
    elif wemix_ratio > diam_ratio * 1.15 and wemix_ratio >= 0.08:
        detected = 'WEMIX'
    else:
        hue_values = hue[sat_mask]
        mean_h = float(np.mean(hue_values)) if hue_values.size else 0.0
        if 80 <= mean_h <= 120:
            detected = 'DIAMANTES'
        elif 125 <= mean_h <= 170:
            detected = 'WEMIX'

    log_message(
        f"[DEBUG] Tab ROI ({x0},{y0})-({x1},{y1}) | ratio_diam={diam_ratio:.3f} ratio_wemix={wemix_ratio:.3f} => {detected}"
    )
    return detected


def init_ocr():
    """Inicializa motores OCR (EasyOCR) si estÃ¡ disponible. Idempotente."""
    global reader
    if reader is None and USE_EASYOCR:
        try:
            log_message("[*] Inicializando EasyOCR (primera vez puede tardar)...")
            reader = easyocr.Reader(['en'], gpu=False)
            log_message("[OK] EasyOCR listo")
        except Exception as e:
            log_message(f"[WARN] fallo inicializando EasyOCR: {e}")
            reader = None


def capturar_ventana():
    """Restaura/activa la ventana del juego y captura la regiÃ³n de la tabla.

    Devuelve una imagen BGR (numpy) o None si falla.
    """
    win = obtener_ventana_juego()
    if not win:
        return None

    try:
        # Restaurar si estÃ¡ minimizada
        if getattr(win, 'isMinimized', None) and win.isMinimized:
            try:
                win.restore()
            except Exception:
                pass
        # Intentar activar/traer al frente
        try:
            win.activate()
        except Exception:
            try:
                win.minimize()
                time.sleep(0.05)
                win.restore()
            except Exception:
                pass

        time.sleep(0.4)

        # Try to use Win32 APIs to get the client rect (visible content) if available
        left = None
        top = None
        width = None
        height = None
        hwnd = None
        try:
            import win32gui
            import win32con
            # try to obtain hwnd from pygetwindow Window object
            for attr in ('_hWnd', 'hWnd', 'hwnd'):
                if hasattr(win, attr):
                    hwnd = getattr(win, attr)
                    break
            # fallback: find by title substring
            if not hwnd:
                title = win.title if getattr(win, 'title', None) else ''
                def enum_proc(h, ctx):
                    txt = win32gui.GetWindowText(h)
                    if txt and title.lower() in txt.lower():
                        ctx.append(h)
                matches = []
                try:
                    win32gui.EnumWindows(lambda h, ctx=matches: enum_proc(h, ctx), None)
                except Exception:
                    pass
                if matches:
                    hwnd = matches[0]

            if hwnd:
                # get client rect (content) and convert to screen coords
                try:
                    l, t, r, b = win32gui.GetClientRect(hwnd)
                    # client rect is relative to window; convert to screen coords
                    pt = win32gui.ClientToScreen(hwnd, (l, t))
                    left = int(pt[0])
                    top = int(pt[1])
                    width = int(r - l)
                    height = int(b - t)
                except Exception:
                    # fallback to GetWindowRect
                    try:
                        l, t, r, b = win32gui.GetWindowRect(hwnd)
                        left, top = int(l), int(t)
                        width = int(r - l)
                        height = int(b - t)
                    except Exception:
                        left = None
        except Exception:
            hwnd = None

        # final fallback to pygetwindow attributes
        if left is None:
            left = getattr(win, 'left', None) or (getattr(win, 'topleft', (0, 0))[0])
            top = getattr(win, 'top', None) or (getattr(win, 'topleft', (0, 0))[1])
            try:
                width = int(getattr(win, 'width'))
            except Exception:
                try:
                    width = int(getattr(win, 'right') - left)
                except Exception:
                    width = 800
            try:
                height = int(getattr(win, 'height'))
            except Exception:
                try:
                    height = int(getattr(win, 'bottom') - top)
                except Exception:
                    height = 600

        # Capture the full window client area (no cropping)
        rx = int(left)
        ry = int(top)
        rw = int(width)
        rh = int(height)

        log_message(f"[DEBUG] Captura completa ventana: x={rx}, y={ry}, w={rw}, h={rh} (hwnd={hwnd})")

        im = pyautogui.screenshot(region=(rx, ry, rw, rh))
        im = cv2.cvtColor(np.array(im), cv2.COLOR_RGB2BGR)
        return im
    except Exception as e:
        log_message(f"[ERROR] capturar_ventana fallo: {e}")
        return None
def extraer_precios(img, debug_tag=None):
    """Extrae columnas usando OCR por ROI.
    Devuelve lista de dicts con claves de columnas configuradas.
    """
    init_ocr()
    if reader is None:
        log_message('[ERROR] EasyOCR no inicializado')
        return []
    rois_config = load_rois_config('rois.json')

    base_cols = ['item', 'price', 'sales']
    control_rois = {'tab_switch'}
    column_order = base_cols + [c for c in rois_config.keys() if c not in base_cols and c not in control_rois]
    rois = {}
    for col_name in column_order:
        cfg = rois_config[col_name]
        x0, y0, x1, y1 = _roi_rect_from_rel(img.shape, cfg)
        rois[col_name] = (x0, y0, x1 - x0, y1 - y0)
    detections_by_col = {col_name: [] for col_name in column_order}
    os.makedirs('debug_cols', exist_ok=True)
    for col_name, (rx, ry, rw, rh) in rois.items():
        roi_color = img[ry:ry+rh, rx:rx+rw]
        is_numeric = col_name in ['price', 'sales']
        proc, scale_factor = preprocesar(roi_color, for_numeric=is_numeric)
        variants = [proc]
        if is_numeric:
            _, inv_otsu = cv2.threshold(proc, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            variants.append(inv_otsu)
        allowlist = "0123456789.," if is_numeric else "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -_'\u2019\u00B4"
        raw_results = []
        for variant in variants:
            try:
                raw_results.extend(reader.readtext(variant, detail=1, paragraph=False, allowlist=allowlist))
            except Exception as e:
                log_message(f"[WARN] EasyOCR fallo en columna {col_name}: {e}")
        dedup = []
        for (bbox, text, conf) in raw_results:
            text = text.strip()
            if not text:
                continue
            y_center = int((bbox[0][1] + bbox[2][1]) / 2)
            duplicate_idx = None
            for i, prev in enumerate(dedup):
                prev_bbox, prev_text, prev_conf = prev
                prev_y = int((prev_bbox[0][1] + prev_bbox[2][1]) / 2)
                iou = bbox_iou(prev_bbox, bbox)
                is_item_overlap_dup = (col_name == 'item' and abs(prev_y - y_center) <= 14 and iou >= 0.35)
                is_exact_dup = (prev_text == text and abs(prev_y - y_center) <= 10)
                if is_exact_dup or is_item_overlap_dup:
                    duplicate_idx = i
                    if col_name == 'item':
                        if item_candidate_score(text, conf) > item_candidate_score(prev_text, prev_conf):
                            dedup[i] = (bbox, text, conf)
                    elif conf > prev_conf:
                        dedup[i] = (bbox, text, conf)
                    break
            if duplicate_idx is None:
                dedup.append((bbox, text, conf))
        vis_base = proc if len(proc.shape) == 3 else cv2.cvtColor(proc, cv2.COLOR_GRAY2BGR)
        vis = cv2.cvtColor(vis_base, cv2.COLOR_BGR2RGB)
        for (bbox, text, conf) in dedup:
            x0, y0 = int(bbox[0][0]), int(bbox[0][1])
            x2, y2 = int(bbox[2][0]), int(bbox[2][1])
            cv2.rectangle(vis, (x0, y0), (x2, y2), (0, 255, 0), 1)
            cv2.putText(vis, f"{text} {conf:.2f}", (x0, max(0, y0 - 6)), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (0, 255, 0), 1)
        try:
            # Usar nombre fijo por columna para sobrescribir el debug anterior.
            debug_name = f'col_{col_name}.png'
            debug_path = os.path.join('debug_cols', debug_name)
            cv2.imwrite(debug_path, cv2.cvtColor(vis, cv2.COLOR_RGB2BGR))
        except Exception:
            pass
        conf_threshold = 0.10 if is_numeric else 0.20
        for (bbox, text, conf) in dedup:
            if conf < conf_threshold:
                continue
            # Convertir coordenadas del frame escalado al sistema original.
            x_center = int(((bbox[0][0] + bbox[2][0]) / 2) / scale_factor) + rx
            y_center = int(((bbox[0][1] + bbox[2][1]) / 2) / scale_factor) + ry
            detections_by_col[col_name].append({'x': x_center, 'y': y_center, 'text': text, 'conf': conf})
    for col_name in detections_by_col:
        detections_by_col[col_name].sort(key=lambda d: (d['y'], d['x']))

    # Estimar altura de fila usando columnas numÃ©ricas (mÃ¡s estables que item).
    numeric_ys = sorted([d['y'] for c in ['price', 'sales'] for d in detections_by_col.get(c, [])])
    row_step = None
    if len(numeric_ys) >= 2:
        diffs = [numeric_ys[i] - numeric_ys[i - 1] for i in range(1, len(numeric_ys))]
        filtered_diffs = [d for d in diffs if d > 4]
        if filtered_diffs:
            row_step = float(np.median(filtered_diffs))
    if not row_step:
        all_ys = sorted([d['y'] for col in detections_by_col.values() for d in col])
        if len(all_ys) >= 2:
            diffs = [all_ys[i] - all_ys[i - 1] for i in range(1, len(all_ys))]
            filtered_diffs = [d for d in diffs if d > 6]
            if filtered_diffs:
                row_step = float(np.median(filtered_diffs))

    row_tolerance = 14
    if row_step and row_step > 0:
        row_tolerance = max(10, min(24, int(row_step * 0.45)))

    # Construir filas agrupando detecciones por Y, para evitar explosiÃ³n de filas
    # cuando EasyOCR separa palabras del nombre del item.
    all_detections = []
    for col_name, detections in detections_by_col.items():
        for det in detections:
            all_detections.append({'col': col_name, **det})
    if not all_detections:
        return []

    all_detections.sort(key=lambda d: d['y'])
    row_targets = []
    current_cluster = [all_detections[0]['y']]
    for det in all_detections[1:]:
        cluster_center = int(np.median(current_cluster))
        if abs(det['y'] - cluster_center) <= row_tolerance:
            current_cluster.append(det['y'])
        else:
            row_targets.append(int(np.median(current_cluster)))
            current_cluster = [det['y']]
    row_targets.append(int(np.median(current_cluster)))

    row_buckets = [{col: [] for col in column_order} for _ in row_targets]
    for det in all_detections:
        best_idx = None
        best_dist = None
        for idx, target_y in enumerate(row_targets):
            dist = abs(det['y'] - target_y)
            if best_idx is None or dist < best_dist:
                best_idx = idx
                best_dist = dist
        if best_idx is not None and best_dist <= row_tolerance:
            row_buckets[best_idx][det['col']].append(det)

    rows = []
    incomplete_rows = []
    numeric_cols = [c for c in ['price', 'sales'] if c in column_order]
    for i, bucket in enumerate(row_buckets):
        row = {col: '' for col in column_order}

        item_parts = sorted(bucket.get('item', []), key=lambda d: d['x'])
        if item_parts:
            raw_item = ' '.join(part['text'] for part in item_parts).strip()
            row['item'] = normalize_item_spacing(raw_item)

        for col in [c for c in column_order if c != 'item']:
            candidates = bucket.get(col, [])
            if not candidates:
                continue
            best = max(candidates, key=lambda d: (d['conf'], -abs(d['y'] - row_targets[i])))
            row[col] = best['text']

        # Filtrar ruido: mantener filas con al menos un campo numÃ©rico.
        if numeric_cols and not any(row.get(col) for col in numeric_cols):
            continue

        missing_fields = [col for col in column_order if not row.get(col)]
        if missing_fields:
            incomplete_rows.append((len(rows) + 1, missing_fields))
        rows.append(row)
    if incomplete_rows:
        row_list = ', '.join(str(r[0]) for r in incomplete_rows)
        log_message(f"[WARN] Una fila no pudo ser capturada completamente. Filas: {row_list}")
        for row_num, fields in incomplete_rows:
            log_message(f"[WARN] Fila {row_num} incompleta: falto {', '.join(fields)}")
    return rows

# =========================
# MAIN
# =========================

def main():
    log_message("========================================")
    log_message(" Legend of Ymir - OCR Market Auto")
    log_message("========================================")
    log_message("")

    # Inicializar EasyOCR si estÃ¡ disponible
    init_ocr()
    log_message("")

    # Detectar ventana
    log_message("[*] Buscando ventana del juego...")
    ventana = obtener_ventana_juego()
    if not ventana:
        log_message("[ERROR] No se encontrÃ³ la ventana del juego")
        return False
    log_message(f"[OK] Ventana encontrada: {ventana.title}")
    log_message("[*] Se haran 2 capturas y el programa detecta automaticamente si estas en WEMIX o DIAMANTES")

    items_por_tipo = {'WEMIX': [], 'DIAMANTES': []}

    for idx in range(1, 3):
        log_message("")
        esperar_f12(f"[>] Presiona F12 para captura #{idx} (puede ser WEMIX o DIAMANTES)...")

        log_message(f"[*] Capturando market (captura #{idx})...")
        img = capturar_ventana()
        if img is None:
            log_message("[ERROR] No se pudo capturar la ventana")
            return False

        tipo_detectado = detectar_tipo_market(img)
        if tipo_detectado not in items_por_tipo:
            log_message("[WARN] No se pudo detectar tab activo; se usara WEMIX por defecto para esta captura")
            tipo_detectado = 'WEMIX'

        log_message(f"[*] Tab detectado en captura #{idx}: {tipo_detectado}")
        log_message(f"[*] Extrayendo items y precios de {tipo_detectado}...")
        items = extraer_precios(img, debug_tag=tipo_detectado.lower())
        log_message(f"[OK] Filas encontradas: {len(items)}")

        if items and isinstance(items[0], dict):
            for i, row in enumerate(items[:5]):
                resumen = ", ".join([f"{k}='{str(v)[:40]}'" for k, v in row.items()])
                log_message(f"    {i+1}: {resumen}")
        elif items:
            for i, item in enumerate(items[:3]):
                log_message(f"    {i+1}: {item[:120]}...")

        if items_por_tipo[tipo_detectado]:
            log_message(f"[WARN] Ya habia una captura para {tipo_detectado}; se reemplaza por la mas reciente")
        items_por_tipo[tipo_detectado] = items

    items_wemix = items_por_tipo['WEMIX']
    items_diamantes = items_por_tipo['DIAMANTES']

    # Guardar en Excel
    log_message("")
    log_message("[*] Guardando en Excel...")

    # Construir DataFrame estructurado si OCR devolviÃ³ dicts
    rows = []
    rows.extend(build_rows_for_tipo(items_wemix, "WEMIX"))
    rows.extend(build_rows_for_tipo(items_diamantes, "DIAMANTES"))
    pd.DataFrame(rows).to_excel(EXCEL_PATH, index=False)
    sync_to_google_sheet(items_wemix, items_diamantes)

    log_message("")
    log_message("========================================")
    log_message("[OK] Proceso finalizado")
    log_message(f"[OK] Excel generado: {EXCEL_PATH}")
    log_message(f"[OK] Total filas WEMIX: {len(items_wemix)}")
    log_message(f"[OK] Total filas DIAMANTES: {len(items_diamantes)}")
    log_message("========================================")
    return True


if __name__ == '__main__':
    main()






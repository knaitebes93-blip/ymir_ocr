import pyautogui
import cv2
import numpy as np
import pandas as pd
import time
import os
import json
from datetime import datetime
import pygetwindow as gw
import keyboard
import warnings

# Suprimir warnings de PyTorch/torchvision
warnings.filterwarnings('ignore', category=UserWarning, module='torch.utils.data.dataloader')
warnings.filterwarnings('ignore', category=UserWarning, module='torchvision')
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'  # Suprimir logs de tensorflow si está instalado

# Usar exclusivamente EasyOCR (requisito: no usar Tesseract)
try:
    import easyocr
    USE_EASYOCR = True
except Exception:
    USE_EASYOCR = False
    print('[ERROR] EasyOCR no disponible. Instalar easyocr y dependencias.')

reader = None

EXCEL_PATH = "precios_market.xlsx"

# Flags for optional OCR engines (Paddle/pytesseract not used by default)
USE_PADDLEOCR = False
USE_PYTESSERACT = False
paddle_ocr = None

def esperar_f12(mensaje):
    """Espera a que el usuario presione F12"""
    print("")
    print(mensaje)
    print("[Esperando F12...]")
    keyboard.wait('f12')
    print("[OK] Preparando captura...")
    time.sleep(1)  # Esperar a que se estabilice
    print("[OK] Capturando...")

def obtener_ventana_juego():
    """Detecta la ventana del juego 'YmirGL' (elige la ventana Ymir más grande)."""
    try:
        wins = [w for w in gw.getAllWindows() if w and w.title]
    except Exception:
        return None

    # Filtrar ventanas que contienen 'ymir' o 'ymirgl' en el título (case-insensitive)
    candidatas = [w for w in wins if ('ymir' in w.title.lower() or 'ymirgl' in w.title.lower())]
    if not candidatas:
        return None

    # Preferir ventanas que contengan 'ymirgl' exacto
    exact_ymirgl = [w for w in candidatas if 'ymirgl' in w.title.lower()]
    if exact_ymirgl:
        candidatas = exact_ymirgl

    # Excluir ventanas con títulos de editores o utilidades comunes
    exclude_terms = ['visual studio', 'visual studio code', 'vscode', 'code', 'notepad', 'explorer', 'onenote', 'word', 'excel', 'powerpoint', 'sublime', 'pycharm']
    filtered = [w for w in candidatas if not any(t in w.title.lower() for t in exclude_terms)]
    if filtered:
        candidatas = filtered

    # Elegir la ventana con mayor área (ancho*alto)
    def area(win):
        try:
            return (getattr(win, 'width', 0) or 0) * (getattr(win, 'height', 0) or 0)
        except Exception:
            return 0

    ventana = max(candidatas, key=area)
    return ventana


def preprocesar(img, for_numeric=False):
    """Preprocesado para OCR.
    
    for_numeric=True: resize 3x para números (mejor detección de dígitos pequeños)
    for_numeric=False: resize 2x + CLAHE+sharpen para texto (ítems)
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    if for_numeric:
        # Para columnas numéricas: resize 3x (mejor para dígitos pequeños)
        gray = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_LINEAR)
        return gray
    else:
        # Para items: resize 2x + CLAHE + sharpen
        gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        gray = clahe.apply(gray)
        kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
        sharp = cv2.filter2D(gray, -1, kernel)
        return sharp


def init_ocr():
    """Inicializa motores OCR (EasyOCR) si está disponible. Idempotente."""
    global reader
    if reader is None and USE_EASYOCR:
        try:
            print("[*] Inicializando EasyOCR (primera vez puede tardar)...")
            reader = easyocr.Reader(['en'], gpu=False)
            print("[OK] EasyOCR listo")
        except Exception as e:
            print(f"[WARN] fallo inicializando EasyOCR: {e}")
            reader = None


def capturar_ventana():
    """Restaura/activa la ventana del juego y captura la región de la tabla.

    Devuelve una imagen BGR (numpy) o None si falla.
    """
    win = obtener_ventana_juego()
    if not win:
        return None

    try:
        # Restaurar si está minimizada
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

        print(f"[DEBUG] Captura completa ventana: x={rx}, y={ry}, w={rw}, h={rh} (hwnd={hwnd})")

        im = pyautogui.screenshot(region=(rx, ry, rw, rh))
        im = cv2.cvtColor(np.array(im), cv2.COLOR_RGB2BGR)
        return im
    except Exception as e:
        print(f"[ERROR] capturar_ventana fallo: {e}")
        return None
def extraer_precios(img):
    """Extrae columnas separadas: item, price, sales usando OCR por ROI.

    Devuelve lista de dicts {'item','price','sales'}.
    """
    init_ocr()
    if reader is None:
        print('[ERROR] EasyOCR no inicializado')
        return []

    h_full, w_full = img.shape[:2]

    # Si existe configuración de ROIs guardada, usarla (valores relativos 0.0-1.0)
    cfg_path = 'rois.json'
    rois_config = {
        'item': {'x': [0.00, 0.65], 'y': [0.0, 1.0]},
        'price': {'x': [0.65, 0.90], 'y': [0.0, 1.0]},
        'sales': {'x': [0.90, 1.0], 'y': [0.0, 1.0]},
    }
    
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            # Soportar nuevo formato (dict con 'x' y 'y') y formato antiguo (array)
            for col in ['item', 'price', 'sales']:
                if col in cfg:
                    col_cfg = cfg[col]
                    if isinstance(col_cfg, dict) and 'x' in col_cfg:
                        # Nuevo formato
                        rois_config[col] = col_cfg
                    elif isinstance(col_cfg, (list, tuple)) and len(col_cfg) >= 2:
                        # Formato antiguo: solo X
                        rois_config[col]['x'] = col_cfg
        except Exception as e:
            print(f'[WARN] Error cargando rois.json: {e}')

    rois = {}
    for col_name in ['item', 'price', 'sales']:
        cfg = rois_config[col_name]
        x0_rel, x1_rel = cfg['x'][0], cfg['x'][1]
        y0_rel, y1_rel = cfg['y'][0], cfg['y'][1]
        
        x0 = int(x0_rel * w_full)
        x1 = int(x1_rel * w_full)
        y0 = int(y0_rel * h_full)
        y1 = int(y1_rel * h_full)
        
        # (x, y, width, height)
        rois[col_name] = (x0, y0, x1 - x0, y1 - y0)

    detections_by_col = {'item': [], 'price': [], 'sales': []}

    # Crear carpeta debug
    os.makedirs('debug_cols', exist_ok=True)

    for col_name, (rx, ry, rw, rh) in rois.items():
        roi_color = img[ry:ry+rh, rx:rx+rw]
        # Use different preprocessing for numeric vs text columns
        is_numeric = col_name != 'item'
        proc = preprocesar(roi_color, for_numeric=is_numeric)
        try:
            results = reader.readtext(proc, detail=1, paragraph=False,
                                      allowlist=( '0123456789.,' if col_name != 'item' else 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -_'))
        except Exception as e:
            print(f"[WARN] EasyOCR fallo en columna {col_name}: {e}")
            results = []

        # Save debug image annotated
        vis = cv2.cvtColor(proc if len(proc.shape)==3 else cv2.cvtColor(proc, cv2.COLOR_GRAY2BGR), cv2.COLOR_BGR2RGB)
        for (bbox, text, conf) in results:
            x0, y0 = int(bbox[0][0]), int(bbox[0][1])
            x2, y2 = int(bbox[2][0]), int(bbox[2][1])
            cv2.rectangle(vis, (x0, y0), (x2, y2), (0,255,0), 1)
            cv2.putText(vis, f"{text} {conf:.2f}", (x0, max(0,y0-6)), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (0,255,0), 1)
        try:
            cv2.imwrite(os.path.join('debug_cols', f'col_{col_name}.png'), cv2.cvtColor(vis, cv2.COLOR_RGB2BGR))
        except Exception:
            pass

        # Lower confidence threshold for numeric columns (numbers can have legitimate lower confidence)
        conf_threshold = 0.10 if col_name != 'item' else 0.20
        for (bbox, text, conf) in results:
            if conf < conf_threshold:
                continue
            x_center = int((bbox[0][0] + bbox[2][0]) / 2) + rx
            y_center = int((bbox[0][1] + bbox[2][1]) / 2) + ry
            detections_by_col[col_name].append({'x': x_center, 'y': y_center, 'text': text.strip(), 'conf': conf})

    # Combine Y positions and cluster into rows
    # Use items as reference points for row centers
    item_ys = [d['y'] for d in detections_by_col['item']]
    if not item_ys:
        return []
    
    item_ys.sort()

    # Sort detections by Y within each column
    for col_name in detections_by_col:
        detections_by_col[col_name].sort(key=lambda d: d['y'])

    # Create rows by simple index matching (order of detections)
    # Row 0 = item[0] + price[0] + sales[0]
    # Row 1 = item[1] + price[1] + sales[1], etc.
    rows = []
    row_count = len(item_ys)
    
    for i in range(row_count):
        row = {'item': '', 'price': '', 'sales': ''}
        
        # Item
        if i < len(detections_by_col['item']):
            row['item'] = detections_by_col['item'][i]['text']
        
        # Price
        if i < len(detections_by_col['price']):
            row['price'] = detections_by_col['price'][i]['text']
        
        # Sales
        if i < len(detections_by_col['sales']):
            row['sales'] = detections_by_col['sales'][i]['text']
        
        rows.append(row)

    return rows

def detectar_tipo_market(img):
    """Detecta si es WEMIX o Diamantes mirando el tab activo"""
    # Buscar texto "WEMIX" en la esquina superior derecha
    altura, ancho = img.shape[:2]
    region_tab = img[0:100, max(0, ancho-400):]
    
    gray = cv2.cvtColor(region_tab, cv2.COLOR_BGR2GRAY)
    txt = pytesseract.image_to_string(gray, config="--psm 6")
    
    if "WEMIX" in txt.upper():
        return "WEMIX"
    else:
        return "DIAMANTES"

# =========================
# MAIN
# =========================

def main():
    print("========================================")
    print(" Legend of Ymir - OCR Market Auto")
    print("========================================")
    print()

    # Inicializar EasyOCR si está disponible
    init_ocr()
    print()

    # Detectar ventana
    print("[*] Buscando ventana del juego...")
    ventana = obtener_ventana_juego()
    if not ventana:
        print("[ERROR] No se encontró la ventana del juego")
        input("Presioná ENTER para salir")
        exit(1)

    print(f"[OK] Ventana encontrada: {ventana.title}")
    print("[*] Asegurate de estar en el tab WEMIX")
    esperar_f12("[>] Presioná F12 cuando estés listo...")

    # Capturar WEMIX
    print("[*] Capturando tab WEMIX...")
    img_wemix = capturar_ventana()
    if img_wemix is None:
        print("[ERROR] No se pudo capturar la ventana")
        input("Presioná ENTER para salir")
        exit(1)

    print("[*] Extrayendo items y precios de WEMIX...")
    items_wemix = extraer_precios(img_wemix)
    print(f"[OK] Filas encontradas: {len(items_wemix)}")
    if items_wemix:
        # si el resultado es una lista de dicts (structured), mostrar columnas
        if isinstance(items_wemix[0], dict):
            for i, row in enumerate(items_wemix[:5]):
                print(f"    {i+1}: item='{row.get('item','')[:40]}', price='{row.get('price','')}', sales='{row.get('sales','')}'")
        else:
            for i, item in enumerate(items_wemix[:3]):
                print(f"    {i+1}: {item[:120]}...")  # Mostrar los primeros 3

    # Cambiar a Diamantes
    print()
    print("[*] Cambia MANUALMENTE al tab DIAMANTES")
    esperar_f12("[>] Presioná F12 cuando estés listo...")

    # Capturar DIAMANTES
    print("[*] Capturando tab DIAMANTES...")
    img_diamantes = capturar_ventana()
    if img_diamantes is None:
        print("[ERROR] No se pudo capturar la ventana")
        input("Presioná ENTER para salir")
        exit(1)

    print("[*] Extrayendo items y precios de DIAMANTES...")
    items_diamantes = extraer_precios(img_diamantes)
    print(f"[OK] Filas encontradas: {len(items_diamantes)}")
    if items_diamantes:
        if isinstance(items_diamantes[0], dict):
            for i, row in enumerate(items_diamantes[:5]):
                print(f"    {i+1}: item='{row.get('item','')[:40]}', price='{row.get('price','')}', sales='{row.get('sales','')}'")
        else:
            for i, item in enumerate(items_diamantes[:3]):
                print(f"    {i+1}: {item[:120]}...")

    # Guardar en Excel
    print()
    print("[*] Guardando en Excel...")

    # Construir DataFrame estructurado si OCR devolvió dicts
    rows = []
    def append_rows(items, tipo):
        if not items:
            return
        if isinstance(items[0], dict):
            for r in items:
                rows.append({
                    'tipo': tipo,
                    'item': r.get('item',''),
                    'price': r.get('price',''),
                    'sales': r.get('sales',''),
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
        else:
            for r in items:
                rows.append({
                    'tipo': tipo,
                    'contenido': r,
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })

    append_rows(items_wemix, 'WEMIX')
    append_rows(items_diamantes, 'DIAMANTES')

    df = pd.DataFrame(rows)
    df.to_excel(EXCEL_PATH, index=False)

    print()
    print("========================================")
    print("[OK] Proceso finalizado")
    print(f"[OK] Excel generado: {EXCEL_PATH}")
    print(f"[OK] Total filas WEMIX: {len(items_wemix)}")
    print(f"[OK] Total filas DIAMANTES: {len(items_diamantes)}")
    print("========================================")
    esperar_f12("[>] Presioná F12 para salir...")


if __name__ == '__main__':
    main()


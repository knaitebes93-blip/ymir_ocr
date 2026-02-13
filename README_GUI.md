# Legend of Ymir - OCR Market Extractor

## 🎮 Descripción

Este programa automatiza la extracción de precios del Market del juego "Legend of Ymir" usando OCR (reconocimiento óptico de caracteres) con una interfaz gráfica visual.

**NUEVA VERSIÓN: Interfaz gráfica mejorada** ✨

Ahora el programa muestra todos los mensajes y acciones directamente en una interfaz visual, sin necesidad de usar la consola.

---

## 📋 Requisitos

- **Python 3.8+** instalado
- **Legend of Ymir** abierto en tu computadora
- Archivo `rois.json` correctamente configurado (ver sección más abajo)

### Dependencias automáticas

El programa instala automáticamente:
- OpenCV (cv2)
- EasyOCR
- NumPy
- Pandas
- OpenPyXL
- PyAutoGUI
- PyGetWindow
- Keyboard
- Pillow

---

## 🚀 Cómo usar

### Opción 1: Interfaz Gráfica (RECOMENDADO)

**Doble-clic en `ejecutar_gui.bat`** para abrir la interfaz gráfica con todos los logs visuales.

### Opción 2: Línea de comandos

```bash
python gui_main.py
```

---

## 🛠️ Configuración inicial (ROIs)

Antes de ejecutar por primera vez, necesitas configurar las regiones de interés (ROIs) del Market:

### Pasos:

1. Abre el juego y ve al Market (WEMIX o Diamantes, da igual)
2. En la consola, ejecuta:
   ```bash
   python select_rois_tk.py
   ```
3. Aparecerá una ventana mostrando la captura del Market
4. **Selecciona las regiones** donde están:
   - **Item** (nombre del producto)
   - **Price** (precio)
   - **Sales** (volumen vendido)

5. Dibuja rectángulos alrededor de cada columna
6. El archivo `rois.json` se genera automáticamente

### Ejemplo de `rois.json`:

```json
{
  "item": {
    "x": [0.0, 0.65],
    "y": [0.0, 1.0]
  },
  "price": {
    "x": [0.65, 0.90],
    "y": [0.0, 1.0]
  },
  "sales": {
    "x": [0.90, 1.0],
    "y": [0.0, 1.0]
  }
}
```

---

## 📖 Flujo de uso (Interfaz Gráfica)

### 1. **Abra `ejecutar_gui.bat`**
   Una ventana se abrirá mostrando la interfaz del programa

### 2. **Presione "Iniciar Extracción"**
   El programa:
   - Detecta la ventana del juego
   - Muestra mensajes en la ventana de logs ✓

### 3. **Siga las instrucciones en la GUI**
   - Asegúrese estar en el tab **WEMIX**
   - Cuando vea el mensaje "Presioná F12 cuando estés listo"
   - **Presione F12** en el juego
   - El programa capturará y extraerá precios automáticamente

### 4. **Cambio a DIAMANTES**
   - Cambie manualmente al tab DIAMANTES en el juego
   - Presione F12 nuevamente cuando esté listo
   - El programa extraerá los precios de diamantes

### 5. **Resultado**
   - Se genera un archivo Excel: `precios_market.xlsx`
   - Contiene todas las filas de ambos markets
   - Incluye timestamps de extracción

---

## 📊 Archivos generados

### `precios_market.xlsx`
Archivo Excel con columnas:
- **tipo**: WEMIX o DIAMANTES
- **item**: Nombre del producto
- **price**: Precio extraído
- **sales**: Volumen de ventas
- **timestamp**: Fecha y hora de extracción

### `debug_cols/`
Carpeta con imágenes de debug que muestran:
- `col_item.png` - Detecciones de items
- `col_price.png` - Detecciones de precios
- `col_sales.png` - Detecciones de volumen

---

## 🎯 Características principales

✅ **Interfaz gráfica visual**
- Todos los logs y mensajes en una ventana bonita
- Sin necesidad de consola
- Colores para diferentes tipos de mensajes

✅ **Automático**
- Detecta automáticamente la ventana del juego
- Captura regiones específicas via ROIs
- Genera Excel automáticamente

✅ **Confiable**
- Usa EasyOCR (mejor que Tesseract)
- Preprocesamiento optimizado para números y texto
- Manejo de errores robusto

✅ **Rápido**
- Extracción de múltiples precios en segundos
- Threading para no bloquear la GUI
- Inicialización de OCR una sola vez

---

## ⚠️ Solución de problemas

### "No se encontró la ventana del juego"
- Asegúrate que Legend of Ymir está abierto
- Verifica que la ventana es visible en la pantalla

### "Error al cargar rois.json"
- Ejecuta `python select_rois_tk.py` nuevamente
- Asegúrate de dibujar rectángulos en todas las columnas

### "EasyOCR fallo"
- Primera ejecución puede tardar mientras descarga modelos
- Requiere conexión a internet la primera vez
- Modelos se guardan en caché después

### "No se capturó la ventana correctamente"
- Verifica que el Market está completamente visible
- Intenta cambiar la resolución o posición de la ventana

---

## 🔧 Archivos principales

| Archivo | Descripción |
|---------|------------|
| `gui_main.py` | Interfaz gráfica principal |
| `ocr_market.py` | Motor OCR y lógica principal |
| `select_rois_tk.py` | Herramienta para configurar ROIs |
| `ejecutar_gui.bat` | Ejecutor de la GUI (Windows) |
| `ejecutar.bat` | Ejecutor de línea de comandos (Legacy) |
| `rois.json` | Configuración de regiones de interés |
| `precios_market.xlsx` | Archivo Excel con resultados |

---

## 📝 Notas técnicas

- **OCR Engine**: EasyOCR (mejor para texto pequeño que Tesseract)
- **Preprocesamiento**: 
  - Items: 2x resize + CLAHE + Sharpen
  - Números: 3x resize (mejor para dígitos)
- **Lenguaje**: Inglés (puede adaptarse)
- **Thread-safe**: La GUI no se bloquea durante extracción

---

## 🎨 Interfaz Visual

La nueva interfaz incluye:

```
┌─ Legend of Ymir - OCR Market Extractor ─────────────────────────┐
│                                                                   │
│  📋 Registro de actividades:                                     │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │ [14:32:45] ========================================           │ │
│  │ [14:32:45]   BIENVENIDO AL EXTRACTOR OCR DE YMIR MARKET    │ │
│  │ [14:32:45] ========================================           │ │
│  │ [14:32:46] ▸ Asegúrese de que:                              │ │
│  │ [14:32:46]   1. El juego Legend of Ymir está abierto       │ │
│  │ [14:32:46]   2. El archivo rois.json está configurado      │ │
│  │ [14:32:46] ✓ Proceso completado exitosamente               │ │
│  └────────────────────────────────────────────────────────────┘ │
│                                                                   │
│  Progreso: [████████████░░░░░░░░░░░░░░░]                       │
│                                                                   │
│  [Estado: Listo]                                                │
│                                                                   │
│  [▶ Iniciar] [🗑️ Limpiar] .......................... [❌ Salir]  │
└─────────────────────────────────────────────────────────────────┘
```

---

## 📈 Próximas mejoras posibles

- [ ] Multi-idioma
- [ ] Exportar a CSV/JSON además de Excel
- [ ] Gráficos de comparación de precios
- [ ] Historial de extracciones
- [ ] Alertas de cambios de precio

---

## 📞 Soporte

Si tienes problemas:

1. Revisa el archivo de logs en la GUI
2. Ejecuta nuevamente la configuración de ROIs
3. Verifica que all dependencias están instaladas:
   ```bash
   python -m pip list | findstr "opencv easyocr pandas pyautogui"
   ```

---

## ⚖️ Disclaimer

Este programa usa OCR para extraer datos visibles en la pantalla. No intercepta tráfico de red ni accede a memoria del juego. Úsalo responsablemente.

**Legend of Ymir es una marca registrada de Smilegate (SEA).**

---

¡Disfrutá del programa! 🎮✨

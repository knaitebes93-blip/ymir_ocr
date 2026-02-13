# 🎮 NUEVA INTERFAZ GRÁFICA - ACTUALIZACIÓN

## ✨ ¿QUÉ HAY DE NUEVO?

Tu programa **Legend of Ymir OCR Market Extractor** ahora tiene una **interfaz gráfica visual profesional** en lugar de mensajes en consola.

### Cambios principales:

| Aspecto | Antes | Ahora |
|---------|-------|-------|
| **Interface** | Consola negra (CMD) | Ventana gráfica bonita |
| **Logs** | Texto plano | Texto con colores (verde/rojo/naranja) |
| **Timestamps** | No había | Sí, cada mensaje con hora |
| **Progreso** | No visible | Barra de progreso animada |
| **Estado** | Confuso | Barra de estado clara |
| **Botones** | No había | Iniciar, Limpiar, Salir |

---

## 🚀 CÓMO EMPEZAR

### Opción 1: La más fácil (Recomendado) ⭐

**Haz doble-clic en:** `ejecutar_gui.bat`

Eso es todo. La interfaz gráfica se abrirá automáticamente.

### Opción 2: Desde PowerShell/CMD

```bash
python gui_main.py
```

---

## 📱 LAYOUT DE LA INTERFAZ

```
┌─────────────────────────────────────────────────────────────────┐
│   Legend of Ymir - OCR Market Extractor                         │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  📋 Registro de actividades:                                    │
│  ┌───────────────────────────────────────────────────────────┐ │
│  │ [14:32:45] [*] Buscando ventana del juego...              │ │
│  │ [14:32:46] [OK] Ventana encontrada: YmirGL                │ │
│  │ [14:32:47] [*] Asegurate de estar en el tab WEMIX        │ │
│  │ [14:32:48] [>] Presioná F12 cuando estés listo...        │ │
│  │ [14:32:52] [OK] Extrayendo items y precios de WEMIX...   │ │
│  │ [14:32:55] [OK] Filas encontradas: 25                    │ │
│  │ [14:33:01] [OK] Proceso completado exitosamente ✓        │ │
│  │                                                            │ │
│  └───────────────────────────────────────────────────────────┘ │
│                                                                 │
│  Progreso: [████████████░░░░░░░░░░░░░░░]                      │
│                                                                 │
│  [Estado: Completado]                                          │
│                                                                 │
│  [▶ Iniciar] [🗑️ Limpiar] .............. [❌ Salir]            │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🎯 FLUJO RÁPIDO

1. **Abre** `ejecutar_gui.bat`
2. **Haz clic** en "▶ Iniciar Extracción"
3. **Sigue** las instrucciones que aparecen en los logs
4. **Presiona** F12 en el juego cuando se te pida
5. **Espera** a que se complete
6. **Listo** - Excel generado: `precios_market.xlsx`

---

## 🎨 ENTENDER LOS COLORES

Los logs tienen colores diferentes según el tipo de mensaje:

| Color | Tipo | Ejemplo |
|-------|------|---------|
| 🟢 Verde | Info/Éxito | `[*] Buscando...` o `[OK] Listo` |
| 🔴 Rojo | Error | `[ERROR] No se encontró ventana` |
| 🟠 Naranja | Advertencia | `[WARN] Error cargando archivo` |
| 🔵 Azul | Debug | `[DEBUG] Captura: x=100, y=50...` |

---

## 📁 ARCHIVOS NUEVOS

Se han agregado estos archivos al proyecto:

```
├── gui_main.py              ← Interfaz gráfica (nuevo)
├── ejecutar_gui.bat         ← Ejecutor de la GUI (nuevo)
├── README_GUI.md            ← Documentación completa (nuevo)
└── GUIA_INTERFAZ_GRAFICA.txt ← Guía paso a paso (nuevo)
```

Los archivos antiguos siguen funcionando:
- `ocr_market.py` (modificado para soportar GUI)
- `select_rois_tk.py` (sin cambios)
- `ejecutar.bat` (sigue funcionando pero es legacy)

---

## 🔧 CAMBIOS EN EL CÓDIGO

### Se agregó soporte para callbacks de logging

Antes (consola):
```python
print("[OK] Proceso completado")
```

Ahora (con GUI):
```python
log_message("[OK] Proceso completado")
```

El sistema automáticamente usa el callback de la GUI si está disponible, o usa `print()` si se ejecuta desde línea de comandos.

---

## 💡 CARACTERÍSTICAS DE LA INTERFAZ

### ▶ Botón "Iniciar Extracción"
- Inicia el proceso OCR
- Se deshabilita durante ejecución
- Se habilita cuando termina

### 🗑️ Botón "Limpiar Logs"
- Borra todos los mensajes
- Útil después de procesos largos

### ❌ Botón "Salir"
- Cierra la aplicación
- Si hay proceso en ejecución, te pide confirmación

### 📊 Barra de Progreso
- Se anima durante ejecución
- Indica que algo está pasando
- Se detiene cuando termina

### 📋 Área de Logs
- Muestra todos los mensajes con timestamps
- Auto-scroll hacia el final
- Colores según tipo de mensaje
- Copy-paste habilitado

### 📌 Barra de Estado
- Muestra el estado actual
- Cambia de color (azul → naranja → verde/rojo)
- Indica "Listo", "En ejecución...", "Completado", etc.

---

## 🐛 SOLUCIÓN DE PROBLEMAS

### La GUI no abre
```bash
# Verifica Python está instalado
python --version

# Si no funciona, instala Python desde:
# https://www.python.org/downloads/
```

### "ModuleNotFoundError: No module named 'tkinter'"
```bash
# En Windows, tkinter viene con Python
# Reinstala Python marcando "tcl/tk and IDLE" en la instalación

# En Linux:
sudo apt-get install python3-tk

# En macOS:
# Ya viene incluido con Python
```

### Los logs no aparecen
- Asegúrate que el juego está abierto
- Verifica que `rois.json` existe
- Mira el archivo `log.txt` para errores

### F12 no funciona
- Asegúrate de presionar F12 en la ventana del juego
- No en la interfaz gráfica
- Usa el teclado físico, no pantalla táctil

---

## 📊 EJEMPLO DE SALIDA

```
[14:30:00] ========================================
[14:30:00]  Legend of Ymir - OCR Market Auto
[14:30:00] ========================================
[14:30:00] 
[14:30:01] [*] Inicializando EasyOCR (primera vez puede tardar)...
[14:30:05] [OK] EasyOCR listo
[14:30:05] [*] Buscando ventana del juego...
[14:30:05] [OK] Ventana encontrada: YmirGL
[14:30:05] [*] Asegurate de estar en el tab WEMIX
[14:30:05] [>] Presioná F12 cuando estés listo...
[14:30:10] [OK] Preparando captura...
[14:30:11] [OK] Capturando...
[14:30:12] [DEBUG] Captura: x=100, y=50, w=1280, h=720
[14:30:13] [*] Capturando tab WEMIX...
[14:30:14] [*] Extrayendo items y precios de WEMIX...
[14:30:17] [OK] Filas encontradas: 25
[14:30:17]     1: item='Potion of Strength', price='100', sales='250'
[14:30:17]     2: item='Potion of Health', price='80', sales='350'
[14:30:18] [*] Cambia MANUALMENTE al tab DIAMANTES
[14:30:18] [>] Presioná F12 cuando estés listo...
[14:30:30] [OK] Preparando captura...
[14:30:35] [OK] Filas encontradas: 18
[14:30:36] [*] Guardando en Excel...
[14:30:36] 
[14:30:36] ========================================
[14:30:36] [OK] Proceso finalizado
[14:30:36] [OK] Excel generado: precios_market.xlsx
[14:30:36] [OK] Total filas WEMIX: 25
[14:30:36] [OK] Total filas DIAMANTES: 18
[14:30:36] ========================================
```

---

## 📚 DOCUMENTACIÓN COMPLETA

Para documentación más detallada:
- **Guía paso a paso**: `GUIA_INTERFAZ_GRAFICA.txt`
- **README completo**: `README_GUI.md`

---

## 🎉 RESUMEN

| Punto | Detalle |
|-------|---------|
| **Ejecutar** | Doble-clic en `ejecutar_gui.bat` |
| **Interfaz** | Ventana gráfica bonita y profesional |
| **Logs** | Con colores, timestamps y scroll automático |
| **Estado** | Barra indicadora clara |
| **Botones** | Iniciar, Limpiar, Salir |
| **Compatibilidad** | 100% compatible con código anterior |
| **Dependencias** | Se instalan automáticamente |

---

## 🚀 ¡LISTO!

**Disfruta tu nuevo programa con interfaz gráfica profesional!**

```
╔════════════════════════════════════════════════╗
║     Haz doble-clic en: ejecutar_gui.bat        ║
║                                                ║
║  ¡Interfaz gráfica automáticamente lista!      ║
╚════════════════════════════════════════════════╝
```

---

**Versión**: 2.0 (Con interfaz gráfica)  
**Fecha**: Febrero 2025  
**Cambios**: [Ver arriba]

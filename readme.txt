Legend of Ymir – OCR Market Price Extractor
===========================================

Este proyecto automatiza la extracción de precios del Market del juego
"Legend of Ymir" usando OCR (reconocimiento óptico de caracteres),
evitando la carga manual de datos.

El objetivo principal es capturar precios de:
- Market WEMIX
- Market Diamantes

y volcarlos a un archivo Excel para luego analizarlos (arbitraje,
comparaciones, spreads, etc.).


--------------------------------------------
1. CONTEXTO Y MOTIVACIÓN
--------------------------------------------

El Market del juego no expone APIs ni endpoints web accesibles.
Los precios solo están disponibles dentro del cliente del juego.

Por ese motivo:
- No se usa scraping web
- No se intercepta tráfico de red
- No se hookea memoria del juego

La solución elegida es OCR controlado sobre la interfaz gráfica,
minimizando riesgos y complejidad.


--------------------------------------------
2. ENFOQUE TÉCNICO
--------------------------------------------

El flujo del script es:

1) El usuario abre el juego y entra al Market
2) Ejecuta el script
3) El script guía al usuario paso a paso:
   - El usuario se asegura de estar en el tab correcto
   - Marca con el mouse la región donde está el precio
4) El script:
   - Toma un screenshot de esa región
   - Preprocesa la imagen (blanco/negro)
   - Aplica OCR SOLO para números
5) Se repite el proceso para el otro tab
6) Se genera un archivo Excel con los precios detectados


--------------------------------------------
3. CARACTERÍSTICAS IMPORTANTES
--------------------------------------------

- El cambio de tabs es MANUAL (no automatizado)
- El script espera confirmación del usuario (ENTER)
- No depende de delays fijos
- No toca el PATH del sistema
- No modifica archivos del juego
- Funciona con cualquier resolución (previa calibración)


--------------------------------------------
4. ESTRUCTURA DEL PROYECTO
--------------------------------------------

ymir_ocr/
│
├─ ejecutar.bat
│   Script principal para usuarios no técnicos.
│   Verifica entorno, instala dependencias y ejecuta el OCR.
│
├─ ocr_market.py
│   Script Python principal.
│   Maneja:
│   - interacción con el usuario
│   - captura de regiones
│   - OCR
│   - generación del Excel
│
├─ requirements.txt
│   Lista de dependencias Python.
│
└─ precios_market.xlsx
    Archivo generado automáticamente con los precios.


--------------------------------------------
5. REQUISITOS
--------------------------------------------

- Windows 10 / 11
- Python 3.10 o superior (recomendado)
  IMPORTANTE: marcar "Add Python to PATH" al instalar
- Tesseract OCR (UB Mannheim build)
- Resolución de pantalla constante
- Ventana del juego sin mover durante la ejecución


--------------------------------------------
6. DEPENDENCIAS PYTHON
--------------------------------------------

- opencv-python
- pytesseract
- numpy
- pandas
- openpyxl
- pyautogui


--------------------------------------------
7. USO BÁSICO
--------------------------------------------

1) Abrir Legend of Ymir
2) Entrar al Market
3) Ejecutar ejecutar.bat
4) Seguir las instrucciones en pantalla:
   - Asegurarse de estar en el tab correcto
   - Marcar con el mouse la región del precio
5) Al finalizar, revisar precios_market.xlsx


--------------------------------------------
8. LIMITACIONES CONOCIDAS
--------------------------------------------

- El OCR puede fallar si:
  - El texto tiene animaciones
  - Hay blur o sombras
  - El contraste es bajo
- Cambiar resolución o mover la ventana invalida la calibración
- El script actual captura un solo item por ejecución


--------------------------------------------
9. MEJORAS FUTURAS (ROADMAP)
--------------------------------------------

- Guardar coordenadas para no recalibrar cada vez
- Loop automático para múltiples ítems
- Scroll controlado del market
- OCR más robusto (adaptive threshold)
- Escritura directa en un Excel de arbitraje existente
- Empaquetado en un solo .exe (PyInstaller)
- Logs más detallados por item


--------------------------------------------
10. NOTAS DE SEGURIDAD
--------------------------------------------

Este proyecto:
- No modifica el juego
- No inyecta DLLs
- No lee memoria
- No intercepta red

Se limita a leer lo que el usuario ve en pantalla,
de forma equivalente a una captura manual.


--------------------------------------------
11. AUTOR / USO
--------------------------------------------

Proyecto experimental / personal.
Uso bajo responsabilidad del usuario.
Pensado para análisis de mercado y arbitraje in-game.


--------------------------------------------
FIN
--------------------------------------------

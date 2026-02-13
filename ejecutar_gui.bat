@echo off
REM Script mejorado para ejecutar el OCR Market con interfaz gráfica
REM Legend of Ymir - OCR Market Extractor

echo.
echo =========================================
echo Legend of Ymir - OCR Market Extractor
echo =========================================
echo.

REM Verificar si Python está instalado
python --version > nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no está instalado o no está en el PATH
    echo Por favor, instale Python 3.8+ desde https://www.python.org/
    pause
    exit /b 1
)

echo [OK] Python encontrado
echo.

REM Instalar/verificar dependencias
echo [*] Verificando dependencias...
python -m pip install --upgrade pip > nul 2>&1

REM Instalar las dependencias requeridas silenciosamente
python -m pip install -q opencv-python easyocr numpy pandas openpyxl pyautogui pygetwindow keyboard pillow > nul 2>&1

if errorlevel 1 (
    echo.
    echo ERROR: No se pudieron instalar las dependencias
    echo Intentando instalar manualmente...
    python -m pip install opencv-python easyocr numpy pandas openpyxl pyautogui pygetwindow keyboard pillow
    if errorlevel 1 (
        echo.
        echo ERROR: La instalación de dependencias falló
        pause
        exit /b 1
    )
)

echo [OK] Dependencias verificadas
echo.

REM Ejecutar la interfaz gráfica
echo [*] Iniciando interfaz gráfica...
echo.

python gui_main.py

if errorlevel 1 (
    echo.
    echo ERROR: La aplicación no se pudo ejecutar
    echo Verifique que gui_main.py esté en el mismo directorio
    pause
    exit /b 1
)

exit /b 0

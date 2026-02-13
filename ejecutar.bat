@echo off
setlocal
chcp 65001 >nul

title Legend of Ymir - OCR Market
color 0A

set LOGFILE=log.txt
set SETUP_FLAG=.setup_done

echo ========================================= > "%LOGFILE%"
echo Legend of Ymir - OCR Market Tool >> "%LOGFILE%"
echo Fecha: %DATE% %TIME% >> "%LOGFILE%"
echo ========================================= >> "%LOGFILE%"
echo. >> "%LOGFILE%"

echo Iniciando...
echo Iniciando... >> "%LOGFILE%"

REM ------------------------------------------------
REM VERIFICAR SI YA SE HIZO SETUP
REM ------------------------------------------------
if exist "%SETUP_FLAG%" (
    echo Setup ya completado, saltando verificaciones...
    echo Setup previo detectado >> "%LOGFILE%"
    goto :EJECUTAR
)

echo Primer uso - verificando configuracion...
echo. >> "%LOGFILE%"

REM ------------------------------------------------
REM VERIFICAR PYTHON
REM ------------------------------------------------
echo Verificando Python...
echo Verificando Python... >> "%LOGFILE%"

python --version >> "%LOGFILE%" 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] Python no detectado
    echo [ERROR] Python no detectado >> "%LOGFILE%"
    echo.
    echo ERROR: Python no esta instalado o no esta en PATH
    echo Descargar desde https://www.python.org/downloads/windows/
    echo Marcar "Add Python to PATH"
    pause
    exit /b 1
)

echo Python OK >> "%LOGFILE%"
echo Python detectado correctamente

REM ------------------------------------------------
REM DETECTAR TESSERACT (SIN PATH)
REM ------------------------------------------------
echo Buscando Tesseract...
echo Buscando Tesseract... >> "%LOGFILE%"

set "TESS_EXE="

IF EXIST "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    set "TESS_EXE=C:\Program Files\Tesseract-OCR\tesseract.exe"
)

IF EXIST "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" (
    set "TESS_EXE=C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
)

IF "%TESS_EXE%"=="" (
    echo [ERROR] Tesseract no encontrado >> "%LOGFILE%"
    echo.
    echo ERROR: Tesseract OCR no esta instalado
    echo Instalar desde:
    echo https://github.com/UB-Mannheim/tesseract/wiki
    echo.
    pause
    exit /b 1
)

echo Tesseract encontrado en: %TESS_EXE%
echo Tesseract encontrado en %TESS_EXE% >> "%LOGFILE%"

"%TESS_EXE%" --version >> "%LOGFILE%" 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] Tesseract existe pero no se puede ejecutar >> "%LOGFILE%"
    echo.
    echo ERROR: Tesseract no se puede ejecutar
    pause
    exit /b 1
)

echo Tesseract OK >> "%LOGFILE%"
echo Tesseract listo

REM ------------------------------------------------
REM INSTALAR DEPENDENCIAS PYTHON
REM ------------------------------------------------
echo.
echo Instalando dependencias Python...
echo Instalando dependencias Python... >> "%LOGFILE%"

python -m pip install --upgrade pip >> "%LOGFILE%" 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] Fallo al actualizar pip >> "%LOGFILE%"
    echo ERROR actualizando pip
    pause
    exit /b 1
)

python -m pip install -r requirements.txt >> "%LOGFILE%" 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] Fallo al instalar requirements.txt >> "%LOGFILE%"
    echo ERROR instalando dependencias
    pause
    exit /b 1
)

echo Dependencias instaladas correctamente
echo Dependencias OK >> "%LOGFILE%"

REM Crear archivo flag para no verificar de nuevo
echo. > "%SETUP_FLAG%"
echo Flag de setup creado >> "%LOGFILE%"

:EJECUTAR
echo.
echo Ejecutando OCR...
echo Ejecutando OCR... >> "%LOGFILE%"

python ocr_market.py

echo OCR finalizado >> "%LOGFILE%"

REM ------------------------------------------------
REM FIN
REM ------------------------------------------------
echo.
echo =========================================
echo Proceso finalizado
echo Revisar log.txt si algo fallo
echo =========================================

echo ========================================= >> "%LOGFILE%"
echo Fin del proceso >> "%LOGFILE%"
echo ========================================= >> "%LOGFILE%"

pause

@echo off
chcp 65001 >nul
title Instalador OCR to Excel App v2.0
color 0B

echo.
echo ========================================
echo    INSTALADOR OCR TO EXCEL APP
echo ========================================
echo.

:: Verificar privilegios de administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Necesitas ejecutar como Administrador
    echo.
    echo Haz clic derecho en este archivo y selecciona:
    echo "Ejecutar como administrador"
    echo.
    pause
    exit /b 1
)

echo [1/8] Verificando sistema...
echo.

:: Verificar Windows 10 o superior
ver | findstr "10." > nul
if errorlevel 1 (
    ver | findstr "11." > nul
    if errorlevel 1 (
        echo [ADVERTENCIA] Windows 10 o 11 recomendado
        echo.
    )
)

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [2/8] Python no encontrado. Instalando...
    echo.
    
    :: Descargar Python
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe' -OutFile 'python_installer.exe'"
    
    if exist python_installer.exe (
        echo Ejecutando instalador de Python...
        echo IMPORTANTE: Marca "Add Python to PATH" durante la instalacion
        echo.
        start /wait python_installer.exe
        del python_installer.exe
    ) else (
        echo [ERROR] No se pudo descargar Python
        echo Descarga manual desde: https://python.org
        pause
        exit /b 1
    )
)

echo [3/8] Verificando Python instalado...
python --version
echo.

echo [4/8] Instalando dependencias de Python...
echo Esto puede tomar algunos minutos...
echo.

:: Instalar pip si no existe
python -m ensurepip --upgrade

:: Instalar dependencias
pip install --upgrade pip
pip install Pillow pytesseract opencv-python-headless pandas openpyxl numpy

echo.
echo [5/8] Instalando Tesseract OCR...
echo.

:: Verificar si Tesseract ya está instalado
where tesseract >nul 2>&1
if errorlevel 1 (
    echo Descargando Tesseract OCR...
    
    :: Descargar instalador de Tesseract
    powershell -Command "Invoke-WebRequest -Uri 'https://github.com/UB-Mannheim/tesseract/wiki/archive/refs/tags/5.3.3.zip' -OutFile 'tesseract.zip'"
    
    if exist tesseract.zip (
        echo Extrayendo Tesseract...
        powershell -Command "Expand-Archive -Path 'tesseract.zip' -DestinationPath 'C:\Tesseract-OCR' -Force"
        del tesseract.zip
        
        :: Agregar al PATH
        setx PATH "%PATH%;C:\Tesseract-OCR" /M
        set PATH=%PATH%;C:\Tesseract-OCR
    ) else (
        echo [ADVERTENCIA] No se pudo descargar Tesseract
        echo.
        echo Instala manualmente desde:
        echo https://github.com/UB-Mannheim/tesseract/wiki
        echo.
    )
) else (
    echo Tesseract ya está instalado.
)

echo.
echo [6/8] Configurando la aplicación...
echo.

:: Crear estructura de carpetas
if not exist "tessdata" mkdir tessdata
if not exist "config" mkdir config
if not exist "docs" mkdir docs
if not exist "ejemplos" mkdir ejemplos
if not exist "utils" mkdir utils
if not exist "exportados" mkdir exportados

:: Descargar datos de idioma para Tesseract
echo Descargando idiomas para OCR...
echo.

:: Inglés
if not exist "tessdata\eng.traineddata" (
    powershell -Command "Invoke-WebRequest -Uri 'https://github.com/tesseract-ocr/tessdata/raw/main/eng.traineddata' -OutFile 'tessdata\eng.traineddata'"
    echo [OK] Inglés descargado
)

:: Español
if not exist "tessdata\spa.traineddata" (
    powershell -Command "Invoke-WebRequest -Uri 'https://github.com/tesseract-ocr/tessdata/raw/main/spa.traineddata' -OutFile 'tessdata\spa.traineddata'"
    echo [OK] Español descargado
)

echo.
echo [7/8] Creando accesos directos...
echo.

:: Crear acceso directo en escritorio
set SCRIPT_DIR=%~dp0
set DESKTOP=%USERPROFILE%\Desktop

echo [Desktop Shortcut] > "%DESKTOP%\OCR to Excel.lnk.url"
echo URL=file:///%SCRIPT_DIR%OCR_APP.py >> "%DESKTOP%\OCR to Excel.lnk.url"
echo IconIndex=0 >> "%DESKTOP%\OCR to Excel.lnk.url"
echo IconFile=%SCRIPT_DIR%icon.ico >> "%DESKTOP%\OCR to Excel.lnk.url"

:: Crear script de inicio
echo @echo off > "Iniciar OCR.bat"
echo chcp 65001 >> "Iniciar OCR.bat"
echo title OCR to Excel App >> "Iniciar OCR.bat"
echo echo Iniciando aplicacion... >> "Iniciar OCR.bat"
echo python "%~dp0OCR_APP.py" >> "Iniciar OCR.bat"
echo pause >> "Iniciar OCR.bat"

echo.
echo [8/8] Instalación completada!
echo.
echo ========================================
echo    INSTALACIÓN EXITOSA
echo ========================================
echo.
echo Accesos creados:
echo 1. Acceso directo en el Escritorio
echo 2. Archivo "Iniciar OCR.bat" en esta carpeta
echo.
echo Para usar la aplicación:
echo 1. Ejecuta "Iniciar OCR.bat"
echo 2. O haz doble clic en el acceso directo
echo.
echo Presiona cualquier tecla para probar la aplicación...
pause >nul

:: Probar la aplicación
echo.
echo Iniciando aplicación OCR to Excel...
echo.
python "OCR_APP.py"

echo.
echo Para ejecutar nuevamente, usa "Iniciar OCR.bat"
echo.
pause
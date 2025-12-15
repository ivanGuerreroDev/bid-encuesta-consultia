@echo off
REM Script para construir el ejecutable de Windows
REM Ejecutar este archivo en Windows con PyInstaller instalado

echo ========================================
echo Generador de Reportes BID - Builder
echo ========================================
echo.

REM Verificar si PyInstaller está instalado
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller no está instalado. Instalando...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo Error: No se pudo instalar PyInstaller
        pause
        exit /b 1
    )
)

echo Instalando dependencias...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Error: No se pudieron instalar las dependencias
    pause
    exit /b 1
)

echo.
echo Construyendo ejecutable...
pyinstaller build_windows.spec --clean

if errorlevel 1 (
    echo.
    echo Error al construir el ejecutable
    pause
    exit /b 1
)

echo.
echo ========================================
echo Ejecutable creado exitosamente!
echo ========================================
echo.
echo El ejecutable se encuentra en: dist\GeneradorReportesBID.exe
echo.
pause

@echo off
REM ================================================================
REM  Lanzador de la GUI en Windows (doble clic).
REM  - Crea el entorno virtual .venv la primera vez.
REM  - Instala las dependencias si faltan.
REM  - Abre la interfaz grafica (gui.py).
REM  Requiere Python 3.12 instalado y en el PATH (python.org).
REM ================================================================
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo ERROR: no se pudo crear el entorno. Instala Python 3.12 desde python.org
        echo y asegurate de marcar "Add python.exe to PATH".
        pause
        exit /b 1
    )
    echo Instalando dependencias ...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)

echo Abriendo la interfaz grafica ...
".venv\Scripts\python.exe" gui.py

if errorlevel 1 (
    echo.
    echo La aplicacion termino con un error. Revisa el mensaje de arriba.
    pause
)
endlocal

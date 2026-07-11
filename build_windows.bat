@echo off
REM ================================================================
REM  Compila un EJECUTABLE .exe autonomo (sin necesidad de Python).
REM
REM  Ejecuta esto UNA VEZ en una maquina Windows que SI tenga Python
REM  3.12 instalado (por ejemplo, la de un desarrollador). El .exe
REM  resultante se copia luego a los equipos que NO tienen Python.
REM
REM  Salida:  dist\GeneradorEtiquetasQR.exe
REM ================================================================
setlocal
cd /d "%~dp0"

REM 1) Entorno virtual de compilacion
if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: instala Python 3.12 desde python.org y marca "Add python.exe to PATH".
        pause
        exit /b 1
    )
)

REM 2) Dependencias + PyInstaller
echo Instalando dependencias y PyInstaller ...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo ERROR instalando dependencias.
    pause
    exit /b 1
)

REM 3) Compilar a un unico .exe con ventana (sin consola).
REM    Se empaquetan los logos de muestra como respaldo; el usuario
REM    igual puede elegir sus propios archivos desde la interfaz.
echo Compilando el ejecutable ...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --windowed ^
    --name "GeneradorEtiquetasQR" ^
    --add-data "cliente.png;." ^
    --add-data "LOGO_RPCI.jpg;." ^
    gui.py
if errorlevel 1 (
    echo ERROR durante la compilacion.
    pause
    exit /b 1
)

echo.
echo ================================================================
echo  LISTO. El ejecutable esta en:  dist\GeneradorEtiquetasQR.exe
echo.
echo  Copialo a los equipos sin Python. Junto al .exe deja tu
echo  TAGS.csv (y, si quieres tus propios logos, cliente.png y
echo  LOGO_RPCI.jpg). Ahi mismo se generaran la carpeta URLS\ y el
echo  documento Images_Table.docx.
echo ================================================================
pause
endlocal

# CreateSupersetRectangles

Genera etiquetas con código QR (logo del cliente al centro + logo de RPCI + texto del TAG) a partir de un listado de enlaces en CSV, y las inserta automáticamente en una tabla de un documento Word listo para imprimir.

Incluye dos tipos de etiqueta:

| Tipo | Etiqueta | Orientación | Tamaño físico |
|---|---|---|---|
| TCard | tipo tarjeta / credencial (Equipos e Instrumentos) | Portrait | 5.5 cm alto x 8.6 cm ancho |
| Adhesivo | Blowers / CCM | Landscape | 11 cm ancho x 7 cm alto |

## Formas de uso

Hay tres maneras de generar las etiquetas, todas usando el mismo motor (`qr_generator.py`):

1. **Interfaz gráfica (recomendada)** — `python gui.py`: eliges el tipo de etiqueta, seleccionas el CSV fuente y la salida, y generas con un botón.
2. **Línea de comandos** — `python qr_generator.py tcard TAGS.csv` (o `adhesive`).
3. **Scripts originales** — `python main.py` (TCard) y `python main_blowers.py` (Adhesivo), que se conservan por compatibilidad.

## Requisitos

- **Python 3.12** en Windows o Linux/Ubuntu (multiplataforma). La fuente se resuelve automáticamente: Arial Bold en Windows, DejaVu Sans Bold en Linux/Mac.
- Dependencias de Python: ver `requirements.txt` (`qrcode`, `pillow`, `python-docx`).
- Para la interfaz gráfica se usa **Tkinter** (viene con Python). En Ubuntu/Debian, si no está, instálalo con:
  ```bash
  sudo apt install python3-tk
  ```

## Instalación

Windows (nativo, sin WSL):

Instala **Python 3.12** desde [python.org](https://www.python.org/downloads/windows/) marcando *"Add python.exe to PATH"* (Tkinter viene incluido; no hace falta instalar nada aparte). Luego:

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python gui.py
```

> Si PowerShell bloquea `activate`, ejecuta una vez `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, o usa directamente `.venv\Scripts\python.exe gui.py` sin activar.

**Atajo:** haz **doble clic en `run_windows.bat`** — crea el entorno, instala dependencias la primera vez y abre la interfaz.

### Equipos sin Python (ejecutable .exe autónomo)

Para máquinas donde **no se puede instalar Python** (por políticas de la empresa), se genera un **`.exe` independiente** que ya lleva Python y las librerías dentro:

1. En una máquina Windows que **sí** tenga Python 3.12, haz **doble clic en `build_windows.bat`** (o córrelo desde la consola). Al terminar crea **`dist\GeneradorEtiquetasQR.exe`**.
2. **Copia ese `.exe`** a los equipos sin Python. Se ejecuta con doble clic, sin instalar nada.
3. Junto al `.exe`, deja tu **`TAGS.csv`** (y, si quieres tus propios logos, `cliente.png` y `LOGO_RPCI.jpg`). Ahí mismo se generan la carpeta **`URLS\`** y el documento **`Images_Table.docx`**.

El `.exe` se compila con [PyInstaller](https://pyinstaller.org/) (`--onefile --windowed`). Nota: PyInstaller **no** puede generar un `.exe` de Windows desde Linux/WSL; la compilación debe hacerse en Windows.

Linux / Ubuntu:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
sudo apt install python3-tk   # solo si vas a usar la interfaz gráfica
```

## Archivos de entrada necesarios

Deben existir en la raíz del proyecto antes de ejecutar:

- **`TAGS.csv`** — listado de etiquetas a generar, delimitado por `;`, con columnas:
  - `DOMAIN`: dominio/base del sitio (ej. `https://sites.google.com/view/qr-resinas-cajic/`)
  - `SUBSITE`: subcarpeta o sección (ej. `EQUIPOS/`)
  - `TAG`: identificador que se imprime debajo del QR y da nombre al archivo generado (ej. `800-P-01`)
  - `LINK`: URL completa a codificar en el QR
- **`cliente.png`** — logo del cliente, se pega al centro del QR.
- **`LOGO_RPCI.jpg`** — logo de RPCI, se pega a la derecha del QR.
- **Carpeta `URLS/`** — el motor (`qr_generator.py` / `gui.py`) la crea si no existe. Ahí se guarda cada etiqueta como `URLS/<TAG>.png`. (Los scripts originales `main.py`/`main_blowers.py` todavía requieren que exista de antemano.)

## Uso

### Interfaz gráfica (recomendada)

```bash
python gui.py        # Windows
python3 gui.py       # Linux / Mac
```

En la ventana: (1) elige TCard o Adhesivo, (2) selecciona el CSV fuente (y opcionalmente los logos y la carpeta/documento de salida), (3) pulsa **Generar etiquetas**. Una barra de progreso y un registro muestran el avance.

En modo **Adhesivo** se habilita el campo **Ancho (cm)**. **El ancho manda**: la imagen se inserta con exactamente ese ancho y el **Alto (cm)** se calcula solo por proporción y se muestra (no editable), para que el QR y el logo **nunca se deformen**. Por ejemplo, 15 cm de ancho → imagen de 15 × 8.87 cm. (No se pueden fijar ancho y alto a la vez sin deformar la imagen, por eso el alto es automático.)

**Número de columnas del Word (adhesivo):** si el ancho es **≤ 13 cm** la tabla usa 2 columnas (dos adhesivos por fila); si es **> 13 cm** usa 1 sola columna (incluso con la hoja en horizontal). El umbral está en la constante `ADHESIVE_TWO_COLUMN_MAX_CM` de `qr_generator.py`.

### Línea de comandos

```bash
python qr_generator.py tcard TAGS.csv      # TCard
python qr_generator.py adhesive TAGS.csv   # Adhesivo
```

### Scripts originales (compatibilidad)

```bash
python main.py            # TCard (portrait)
python main_blowers.py    # Adhesivo (landscape, Blowers/CCM)
```

Cada ejecución:

1. Lee el CSV y genera un PNG por fila en la carpeta de salida (QR + logo cliente + logo RPCI + texto TAG).
2. Arma el documento Word con todas las imágenes en una tabla de 2 columnas, lista para imprimir/cortar.

## Ajustes según tipo de etiqueta

Las diferencias entre TCard y Adhesivo (orientación del Word, tamaño de imagen, tamaño de fuente, tamaño y posición del logo RPCI) están centralizadas en el diccionario `PRESETS` de `qr_generator.py`. Con la GUI o la CLI **no** hay que editar código: basta elegir el tipo. Los scripts originales conservan los bloques comentados *"Para Equipos e Instrumentos"* vs *"Para Blowers y CCM"* que había que alternar a mano.

## Estructura del repositorio

- `gui.py` — interfaz gráfica de escritorio (Tkinter).
- `qr_generator.py` — motor compartido parametrizado (tcard/adhesive); usable por GUI, CLI o import.
- `main.py`, `main_blowers.py` — scripts originales, conservados por compatibilidad.
- `logos/` — logos de clientes usados en distintos proyectos.
- `TCard/`, `Adhesive/` — CSVs y documentos (Word/PDF) generados históricamente por cliente, agrupados por tipo de etiqueta.
- `URLS/` — salida de las imágenes QR generadas en la última ejecución.
- `Images_Table.docx` — documento Word generado con las etiquetas listas para imprimir.
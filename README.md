# CreateSupersetRectangles

Genera etiquetas con código QR (logo del cliente al centro + logo de RPCI + texto del TAG) a partir de un listado de enlaces en CSV, y las inserta automáticamente en una tabla de un documento Word listo para imprimir.

Incluye dos variantes según el tipo de etiqueta a producir:

| Script | Etiqueta | Orientación | Tamaño físico |
|---|---|---|---|
| `main.py` | TCard (tipo tarjeta / credencial) | Portrait | 5.5 cm alto x 8.6 cm ancho |
| `main_blowers.py` | Adhesivo (Blowers / CCM) | Landscape | 11 cm ancho x 7 cm alto |

## Requisitos

- Windows con Python 3.12 (el script usa la fuente `C:/Windows/Fonts/arialbd.ttf`, por lo que solo corre en Windows tal cual está).
- Dependencias de Python: ver `requirements.txt` (`qrcode`, `pillow`, `python-docx`).

## Instalación

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
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
- **Carpeta `URLS/`** — debe existir de antemano (el script no la crea); ahí se guarda cada etiqueta generada como `URLS/<TAG>.png`.

## Uso

Etiquetas tipo TCard (portrait):

```bash
python main.py
```

Etiquetas tipo Adhesivo (landscape, Blowers/CCM):

```bash
python main_blowers.py
```

Cada ejecución:

1. Lee `TAGS.csv` y genera un PNG por fila en `URLS/` (QR + logo cliente + logo RPCI + texto TAG).
2. Arma `Images_Table.docx` con todas las imágenes de `URLS/` en una tabla de 2 columnas, lista para imprimir/cortar.

## Ajustes según tipo de trabajo

Dentro de ambos scripts hay bloques de código comentados (marcados como *"Para Equipos e Instrumentos"* vs *"Para Blowers y CCM"*) que ajustan tamaños de imagen, márgenes y tamaño de fuente según el tipo de etiqueta. Al cambiar de un tipo de trabajo a otro, hay que comentar/descomentar manualmente el bloque correspondiente antes de ejecutar (por ejemplo `BASE_WIDTH`, los tamaños de imagen en `create_WordDocument`, y el tamaño de fuente de la fila en blanco).

## Estructura del repositorio

- `main.py`, `main_blowers.py` — scripts principales.
- `logos/` — logos de clientes usados en distintos proyectos.
- `TCard/`, `Adhesive/` — CSVs y documentos (Word/PDF) generados históricamente por cliente, agrupados por tipo de etiqueta.
- `URLS/` — salida de las imágenes QR generadas en la última ejecución.
- `Images_Table.docx` — documento Word generado con las etiquetas listas para imprimir.
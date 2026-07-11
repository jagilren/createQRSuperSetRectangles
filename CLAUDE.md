# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

Setup:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Run — three entry points, all producing the same output (PNGs in `URLS/` + a Word doc):

```bash
python gui.py                            # desktop GUI (Tkinter) — pick type, CSV, output
python qr_generator.py tcard TAGS.csv    # CLI, TCard mode
python qr_generator.py adhesive TAGS.csv # CLI, Adhesive mode
python main.py                           # legacy TCard script, portrait, 5.5cm x 8.6cm
python main_blowers.py                   # legacy Adhesive script, landscape, 11cm x 7cm
```

The GUI needs Tkinter; on Ubuntu/Debian install it with `sudo apt install python3-tk` (it ships with Python on Windows).

There is no test suite, linter, or build/packaging config in this repo.

## Architecture

There are now two layers:

- **`qr_generator.py`** — the canonical shared engine. All label logic lives here, parameterized by a `mode` (`"tcard"` / `"adhesive"`) whose differences are collected in the `PRESETS` dict (Word orientation, image size, spacer font pt, RPCI-logo size/offset, initial font size, margins). Nothing runs on import; call `generate(csv_path, mode=..., output_dir=..., docx_path=..., progress=...)`. It creates `output_dir` if missing and resolves the font per-OS via `resolve_font_path()` (Arial Bold on Windows, DejaVu Sans Bold on Linux/Mac). Verified pixel-identical to the original `main.py` output in tcard mode.
- **`gui.py`** — Tkinter desktop UI over the engine. Runs `generate()` in a worker thread and marshals progress back to the UI via a `queue.Queue` polled with `root.after`. No generation logic of its own.

**`main.py` / `main_blowers.py` are the legacy scripts**, kept for compatibility. They are near-duplicate pipelines (same function names, same structure) that differ only in orientation/sizing constants and which of their internally commented-out blocks are active. They do NOT share code with `qr_generator.py` — the engine is a faithful re-implementation, not an import. Prefer editing `qr_generator.py` (single source of truth) for new work; only touch the legacy scripts if a caller still depends on them.

Legacy-script pipeline (both scripts run top-to-bottom, no `if __name__ == "__main__"` guard):

Each script runs top-to-bottom as a script (no `if __name__ == "__main__"` guard) in this order:

1. Reads `TAGS.csv` (delimiter `;`, columns `DOMAIN;SUBSITE;TAG;LINK`) row by row.
2. `create_qr_with_logo_label_and_frame()` — builds the QR code from `LINK` via `qrcode`, pastes the client logo (`cliente.png`) centered on top of it.
3. `create_TagTex_at_Bottom()` — pastes the QR onto a white canvas sized like a credit card (dimensions derived from `BASE_WIDTH` and the 4.30cm/2.54in and 1.69 aspect-ratio constants), then renders the `TAG` text below it, auto-growing the font size in a loop until the text width matches a fraction of `BASE_WIDTH` (thresholds depend on `len(TAG)`).
4. `create_tag_text_logoRPCI()` — pastes the RPCI logo (`LOGO_RPCI.jpg`) to the right of the QR, resized to a size hardcoded per script/label-type.
5. The final image is saved as `URLS/<TAG>.png`.
6. `create_WordDocument()` — after all rows are processed, reads every image back out of `URLS/`, lays them into a 2-column table in a new Word document (`python-docx`), and saves it as `Images_Table.docx`.

Key things to know when editing:

- `URLS/` must already exist before running — the scripts never create it (`os.listdir("URLS")` / `f'URLS/{TAG}.png'` will fail otherwise).
- The text-fitting loop in `create_TagTex_at_Bottom()` and the logo sizing in `create_tag_text_logoRPCI()` are tuned per label type via magic numbers, not parameters. Each script has both an active code path and a commented-out alternate path (labeled `#Para Equipos e Instrumentos` vs `#Para Blowers y CCM`) — switching which physical label type you're generating means manually commenting/uncommenting the matching blocks (`BASE_WIDTH`, image `width`/`height` in `create_WordDocument`, `font_size` of the blank spacer row, and `section.orientation` in `main.py`), not passing a flag.
- Font rendering hardcodes `C:/Windows/Fonts/arialbd.ttf` (Windows-only).
- `TCard/` and `Adhesive/` hold historical per-client outputs (CSVs plus generated Word/PDF) from past runs — not code, useful as reference for expected output but not read by the scripts.

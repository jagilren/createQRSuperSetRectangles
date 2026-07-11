"""Interfaz gráfica de escritorio para generar etiquetas QR (TCard / Adhesivo).

Motor gráfico: Tkinter (viene en la biblioteca estándar de Python, funciona
igual en Windows y en Ubuntu, sin dependencias extra de pip).

En Ubuntu, si Tkinter no está instalado, ejecutar:
    sudo apt install python3-tk

Uso:
    python gui.py        (Windows)
    python3 gui.py       (Linux / Mac)

Toda la lógica de generación vive en qr_generator.py; esta capa es solo la UI.
"""

import os
import queue
import threading
import traceback

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:  # pragma: no cover
    raise SystemExit(
        "Tkinter no está disponible.\n"
        "En Ubuntu/Debian instálalo con:  sudo apt install python3-tk"
    )

import qr_generator as engine


# Directorio del script: usado para resolver los archivos por defecto sin
# depender del directorio de trabajo desde el que se lance la app.
APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _default(path):
    """Ruta por defecto relativa a la carpeta del proyecto."""
    return os.path.join(APP_DIR, path)


class QRLabelApp:
    def __init__(self, root):
        self.root = root
        root.title("Generador de Etiquetas QR — RPCI")
        root.minsize(640, 560)

        # Cola para comunicar el hilo de trabajo con la UI.
        self.msg_queue = queue.Queue()
        self.worker = None

        # Variables de la UI.
        self.mode_var = tk.StringVar(value="tcard")
        self.csv_var = tk.StringVar(value=_default("TAGS.csv"))
        self.client_logo_var = tk.StringVar(value=_default("cliente.png"))
        self.rpci_logo_var = tk.StringVar(value=_default("LOGO_RPCI.jpg"))
        self.out_dir_var = tk.StringVar(value=_default("URLS"))
        self.docx_var = tk.StringVar(value=_default("Images_Table.docx"))
        # Tamaño de la etiqueta Adhesivo (cm). La imagen se ajusta DENTRO de
        # esta caja preservando la proporción, así que nunca se deforma.
        self.width_cm_var = tk.StringVar(value="11")
        self.height_cm_var = tk.StringVar(value="7")

        self._build_ui()
        self.mode_var.trace_add("write", lambda *_: self._sync_size_fields())
        self._sync_size_fields()
        self.root.after(100, self._poll_queue)

    # ------------------------------------------------------------------ UI --
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}
        main = ttk.Frame(self.root, padding=14)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)

        # --- 1. Tipo de etiqueta ---------------------------------------- #
        type_frame = ttk.LabelFrame(main, text="1. Tipo de etiqueta", padding=10)
        type_frame.grid(row=0, column=0, sticky="ew", **pad)
        ttk.Radiobutton(
            type_frame, text=engine.PRESETS["tcard"]["label"],
            variable=self.mode_var, value="tcard",
        ).grid(row=0, column=0, sticky="w", padx=6, pady=2)
        ttk.Radiobutton(
            type_frame, text=engine.PRESETS["adhesive"]["label"],
            variable=self.mode_var, value="adhesive",
        ).grid(row=1, column=0, sticky="w", padx=6, pady=2)

        # Tamaño en cm (solo aplica al Adhesivo). El ANCHO manda: se respeta el
        # ancho pedido y el alto se calcula por proporción (no editable), para
        # que el QR y el logo nunca se deformen.
        self.size_frame = ttk.Frame(type_frame)
        self.size_frame.grid(row=2, column=0, sticky="w", padx=6, pady=(8, 2))
        self.size_hint = ttk.Label(
            self.size_frame,
            text="Tamaño Adhesivo — escribe el ancho; el alto se calcula solo:")
        self.size_hint.grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 4))
        ttk.Label(self.size_frame, text="Ancho (cm):").grid(row=1, column=0, sticky="w")
        self.width_entry = ttk.Entry(self.size_frame, textvariable=self.width_cm_var, width=7)
        self.width_entry.grid(row=1, column=1, padx=(4, 14))
        ttk.Label(self.size_frame, text="Alto (cm):").grid(row=1, column=2, sticky="w")
        # El alto es solo informativo (calculado por proporción): no editable.
        self.height_entry = ttk.Entry(self.size_frame, textvariable=self.height_cm_var,
                                      width=7, state="readonly")
        self.height_entry.grid(row=1, column=3, padx=4)
        # Recalcular el alto cada vez que cambie el ancho.
        self.width_cm_var.trace_add("write", lambda *_: self._update_auto_height())

        # --- 2. Archivos de entrada ------------------------------------- #
        src_frame = ttk.LabelFrame(main, text="2. Archivos de entrada", padding=10)
        src_frame.grid(row=1, column=0, sticky="ew", **pad)
        src_frame.columnconfigure(1, weight=1)
        self._file_row(src_frame, 0, "Archivo CSV (fuente):", self.csv_var,
                       self._pick_csv)
        self._file_row(src_frame, 1, "Logo del cliente (centro del QR):",
                       self.client_logo_var, self._pick_client_logo)
        self._file_row(src_frame, 2, "Logo RPCI (derecha del QR):",
                       self.rpci_logo_var, self._pick_rpci_logo)

        # --- 3. Salida --------------------------------------------------- #
        out_frame = ttk.LabelFrame(main, text="3. Salida", padding=10)
        out_frame.grid(row=2, column=0, sticky="ew", **pad)
        out_frame.columnconfigure(1, weight=1)
        self._file_row(out_frame, 0, "Carpeta de imágenes PNG:", self.out_dir_var,
                       self._pick_out_dir, is_dir=True)
        self._file_row(out_frame, 1, "Documento Word:", self.docx_var,
                       self._pick_docx, save=True)

        # --- 4. Acción --------------------------------------------------- #
        action = ttk.Frame(main)
        action.grid(row=3, column=0, sticky="ew", **pad)
        action.columnconfigure(0, weight=1)
        self.generate_btn = ttk.Button(action, text="Generar etiquetas",
                                       command=self._on_generate)
        self.generate_btn.grid(row=0, column=0, sticky="ew", ipady=4)

        # --- Progreso + log --------------------------------------------- #
        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.grid(row=4, column=0, sticky="ew", **pad)

        self.status_var = tk.StringVar(value="Listo.")
        ttk.Label(main, textvariable=self.status_var).grid(
            row=5, column=0, sticky="w", padx=10)

        log_frame = ttk.LabelFrame(main, text="Registro", padding=6)
        log_frame.grid(row=6, column=0, sticky="nsew", **pad)
        main.rowconfigure(6, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=8, wrap="word", state="disabled")
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

    def _file_row(self, parent, row, label, var, command, is_dir=False, save=False):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="Examinar…", command=command).grid(
            row=row, column=2, padx=6)

    # -------------------------------------------------------------- pickers --
    def _pick_csv(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo CSV",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos", "*.*")],
            initialdir=self._initdir(self.csv_var))
        if path:
            self.csv_var.set(path)

    def _pick_client_logo(self):
        path = filedialog.askopenfilename(
            title="Selecciona el logo del cliente",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg"), ("Todos", "*.*")],
            initialdir=self._initdir(self.client_logo_var))
        if path:
            self.client_logo_var.set(path)

    def _pick_rpci_logo(self):
        path = filedialog.askopenfilename(
            title="Selecciona el logo RPCI",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg"), ("Todos", "*.*")],
            initialdir=self._initdir(self.rpci_logo_var))
        if path:
            self.rpci_logo_var.set(path)

    def _pick_out_dir(self):
        path = filedialog.askdirectory(
            title="Selecciona la carpeta de salida",
            initialdir=self._initdir(self.out_dir_var))
        if path:
            self.out_dir_var.set(path)

    def _pick_docx(self):
        path = filedialog.asksaveasfilename(
            title="Guardar documento Word como",
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")],
            initialdir=self._initdir(self.docx_var))
        if path:
            self.docx_var.set(path)

    @staticmethod
    def _initdir(var):
        d = os.path.dirname(var.get())
        return d if os.path.isdir(d) else APP_DIR

    # -------------------------------------------------------- tamaño (cm) --
    def _sync_size_fields(self):
        """Habilita el ancho solo en modo Adhesivo. El alto siempre es de solo
        lectura (se calcula por proporción)."""
        adhesive = self.mode_var.get() == "adhesive"
        self.width_entry.configure(state="normal" if adhesive else "disabled")
        self.height_entry.configure(state="readonly" if adhesive else "disabled")
        self.size_hint.configure(foreground="black" if adhesive else "gray")
        if adhesive:
            self._update_auto_height()

    def _update_auto_height(self):
        """Calcula el alto proporcional a partir del ancho y lo muestra."""
        try:
            w = float(self.width_cm_var.get().replace(",", "."))
        except ValueError:
            self.height_cm_var.set("—")
            return
        if w <= 0:
            self.height_cm_var.set("—")
            return
        h = w / engine.label_aspect_ratio()
        self.height_cm_var.set(f"{h:.2f}")

    def _read_width_cm(self):
        """Valida y devuelve el ancho en cm, o None si hay error."""
        try:
            w = float(self.width_cm_var.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Error", "El ancho debe ser un número en cm.")
            return None
        if w <= 0:
            messagebox.showerror("Error", "El ancho debe ser mayor que 0.")
            return None
        return w

    # ---------------------------------------------------------- generación --
    def _on_generate(self):
        if self.worker and self.worker.is_alive():
            return

        csv_path = self.csv_var.get().strip()
        if not os.path.exists(csv_path):
            messagebox.showerror("Error", f"No se encuentra el CSV:\n{csv_path}")
            return

        mode = self.mode_var.get()
        width_cm = height_cm = None
        if mode == "adhesive":
            width_cm = self._read_width_cm()
            if width_cm is None:
                return  # error ya mostrado

        self._set_running(True)
        self._log_clear()
        self.progress.configure(value=0, maximum=100)

        params = dict(
            csv_path=csv_path,
            mode=mode,
            output_dir=self.out_dir_var.get().strip(),
            docx_path=self.docx_var.get().strip(),
            client_logo_path=self.client_logo_var.get().strip(),
            rpci_logo_path=self.rpci_logo_var.get().strip(),
            width_cm=width_cm,
            height_cm=height_cm,
        )
        self.worker = threading.Thread(target=self._run_engine, args=(params,), daemon=True)
        self.worker.start()

    def _run_engine(self, params):
        """Se ejecuta en un hilo aparte; comunica progreso por la cola."""
        def progress(done, total, message):
            self.msg_queue.put(("progress", (done, total, message)))

        try:
            result = engine.generate(progress=progress, **params)
            self.msg_queue.put(("done", result))
        except Exception as exc:  # noqa: BLE001 - queremos mostrar cualquier fallo
            self.msg_queue.put(("error", (exc, traceback.format_exc())))

    # ----------------------------------------------------- bucle de la UI --
    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "progress":
                    done, total, message = payload
                    self.progress.configure(maximum=total, value=done)
                    self.status_var.set(message)
                    self._log(message)
                elif kind == "done":
                    self._on_done(payload)
                elif kind == "error":
                    self._on_error(*payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _on_done(self, result):
        self._set_running(False)
        n = len(result["tags"])
        self.status_var.set(f"¡Listo! {n} etiquetas generadas.")
        self._log(f"Documento Word: {result['docx_path']}")
        messagebox.showinfo(
            "Completado",
            f"Se generaron {n} etiquetas.\n\n"
            f"Imágenes: {result['output_dir']}\n"
            f"Word: {result['docx_path']}")

    def _on_error(self, exc, tb):
        self._set_running(False)
        self.status_var.set("Error durante la generación.")
        self._log(tb)
        messagebox.showerror("Error", str(exc))

    # --------------------------------------------------------- utilidades --
    def _set_running(self, running):
        self.generate_btn.configure(
            state="disabled" if running else "normal",
            text="Generando…" if running else "Generar etiquetas")

    def _log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _log_clear(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")


def main():
    root = tk.Tk()
    # Tema nativo más agradable cuando está disponible.
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass
    QRLabelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ──────────────────────────────────────────────────────────────────────────────
#  Mapeo de formatos numéricos (nombre visible → código Excel)
# ──────────────────────────────────────────────────────────────────────────────
NUMBER_FORMAT_MAP = {
    "General":                "General",
    "Número":                 "#,##0.00",
    "Moneda Colombiana":      "[$$-2058]#,##0.00;[RED]-[$$-2058]#,##0.00",
    "Moneda Estadounidense":  "$#,##0.00",
    "Fecha":                  "dd/mm/yyyy",
    "Hora":                   "hh:mm:ss",
    "Porcentaje":             "0.00%",
    "Texto":                  "@",
}

FONTS_AVAILABLE   = ["Arial", "Calibri"]
ALIGN_OPTIONS     = ["left", "center", "right"]
VALIGN_OPTIONS    = ["top", "center", "bottom"]
FONT_SIZE_DEFAULT = 11
FONT_SIZE_MIN     = 8
FONT_SIZE_MAX     = 72


class DataCleanerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("XLSX Data Cleaner")
        self.root.geometry("1150x720")

        # ── Estado interno ──────────────────────────────────────────────────
        self.current_file: str | None = None
        self.dataframe: pd.DataFrame | None = None
        self.edited_cells: dict = {}
        self.cell_formats: dict = {}   # (row_idx, col_idx) → dict de formato
        self.selected_cell: tuple | None = None
        self.row_id_map: dict = {}     # item_id de Treeview → índice de df

        # ── Construcción de la UI ───────────────────────────────────────────
        self._build_ui()

    # ══════════════════════════════════════════════════════════════════════════
    #  CONSTRUCCIÓN DE LA INTERFAZ
    # ══════════════════════════════════════════════════════════════════════════

    def _build_ui(self) -> None:
        """Construye todos los widgets de la ventana principal."""
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self._build_toolbar(main)
        self._build_content(main)

    # ── Barra de herramientas ─────────────────────────────────────────────────

    def _build_toolbar(self, parent: ttk.Frame) -> None:
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, pady=(0, 8))

        self.upload_btn = ttk.Button(bar, text="Cargar archivo",
                                     command=self.upload_file)
        self.upload_btn.pack(side=tk.LEFT, padx=4)

        self.save_btn = ttk.Button(bar, text="Guardar cambios",
                                   command=self.save_file, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=4)

        self.file_label = ttk.Label(bar, text="Ningún archivo cargado")
        self.file_label.pack(side=tk.LEFT, padx=12)

    # ── Área principal (tabla + panel Editar) ─────────────────────────────────

    def _build_content(self, parent: ttk.Frame) -> None:
        content = ttk.Frame(parent)
        content.pack(fill=tk.BOTH, expand=True)
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(0, weight=1)

        self._build_table(content)
        self._build_edit_panel(content)

    # ── Tabla Treeview ────────────────────────────────────────────────────────

    def _build_table(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 8))
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(frame, yscrollcommand=vsb.set,
                                 xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)

        # Eventos
        self.tree.bind("<Button-1>",   self._on_cell_click)
        self.tree.bind("<Double-1>",   self._on_cell_double_click)

    # ── Panel "Editar" ────────────────────────────────────────────────────────

    def _build_edit_panel(self, parent: ttk.Frame) -> None:
        """
        Panel lateral con todas las opciones de edición de celda:
          • Valor de la celda
          • Fuente (Arial / Calibri)
          • Tamaño de fuente
          • Alineación horizontal
          • Alineación vertical
          • Formato de número
        """
        panel = ttk.LabelFrame(parent, text="Editar", padding=12)
        panel.grid(row=0, column=1, sticky=tk.NSEW, padx=4)

        # ── Celda seleccionada ──────────────────────────────────────────────
        self.cell_info_label = ttk.Label(panel, text="Ninguna celda seleccionada",
                                         relief=tk.SUNKEN, anchor=tk.W)
        self.cell_info_label.pack(fill=tk.X, pady=(0, 10))

        # ── Valor ───────────────────────────────────────────────────────────
        self._section_label(panel, "Valor de la celda")
        self.value_entry = ttk.Entry(panel, width=22)
        self.value_entry.pack(fill=tk.X, pady=(0, 10))

        ttk.Separator(panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Fuente ──────────────────────────────────────────────────────────
        self._section_label(panel, "Fuente")
        self.font_var = tk.StringVar(value="Calibri")
        ttk.Combobox(panel, textvariable=self.font_var, width=20,
                     values=FONTS_AVAILABLE, state="readonly").pack(fill=tk.X, pady=(0, 10))

        # ── Tamaño de fuente ─────────────────────────────────────────────────
        self._section_label(panel, "Tamaño de fuente")
        self.font_size_var = tk.StringVar(value=str(FONT_SIZE_DEFAULT))
        ttk.Spinbox(panel, from_=FONT_SIZE_MIN, to=FONT_SIZE_MAX,
                    textvariable=self.font_size_var,
                    width=20).pack(fill=tk.X, pady=(0, 10))

        ttk.Separator(panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Alineación horizontal ────────────────────────────────────────────
        self._section_label(panel, "Alineación horizontal")
        self.h_align_var = tk.StringVar(value="left")
        h_align_frame = ttk.Frame(panel)
        h_align_frame.pack(fill=tk.X, pady=(0, 10))
        for label, value in [("← Izquierda", "left"),
                              ("↔ Centro",    "center"),
                              ("→ Derecha",   "right")]:
            ttk.Radiobutton(h_align_frame, text=label,
                            variable=self.h_align_var, value=value).pack(anchor=tk.W)

        # ── Alineación vertical ──────────────────────────────────────────────
        self._section_label(panel, "Alineación vertical")
        self.v_align_var = tk.StringVar(value="center")
        v_align_frame = ttk.Frame(panel)
        v_align_frame.pack(fill=tk.X, pady=(0, 10))
        for label, value in [("↑ Superior", "top"),
                              ("↕ Medio",   "center"),
                              ("↓ Inferior","bottom")]:
            ttk.Radiobutton(v_align_frame, text=label,
                            variable=self.v_align_var, value=value).pack(anchor=tk.W)

        ttk.Separator(panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Formato de número ────────────────────────────────────────────────
        self._section_label(panel, "Formato de número")
        self.number_format_var = tk.StringVar(value="General")
        ttk.Combobox(panel, textvariable=self.number_format_var, width=20,
                     values=list(NUMBER_FORMAT_MAP.keys()),
                     state="readonly").pack(fill=tk.X, pady=(0, 10))

        ttk.Separator(panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Botón Aplicar ────────────────────────────────────────────────────
        self.apply_btn = ttk.Button(panel, text="✔ Aplicar formato",
                                    command=self._apply_formatting,
                                    state=tk.DISABLED)
        self.apply_btn.pack(fill=tk.X, pady=(4, 0))

    # ── Utilidad: etiqueta de sección ─────────────────────────────────────────

    @staticmethod
    def _section_label(parent: ttk.Frame, text: str) -> None:
        ttk.Label(parent, text=text, font=("TkDefaultFont", 8, "bold")).pack(
            anchor=tk.W, pady=(4, 2))

    # ══════════════════════════════════════════════════════════════════════════
    #  CARGA Y VISUALIZACIÓN DE DATOS
    # ══════════════════════════════════════════════════════════════════════════

    def upload_file(self) -> None:
        """Abre el explorador de archivos para seleccionar un XLSX."""
        path = filedialog.askopenfilename(
            title="Seleccionar archivo XLSX",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if not path:
            return
        try:
            self.current_file = path
            self.dataframe    = pd.read_excel(path)
            self.edited_cells = {}
            self.cell_formats = {}
            self._display_data()
            self.file_label.config(text=f"Cargado: {os.path.basename(path)}")
            self.save_btn.config(state=tk.NORMAL)
            messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{exc}")

    def _display_data(self) -> None:
        """Rellena el Treeview con los datos del DataFrame."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.row_id_map = {}

        if self.dataframe is None:
            return

        self.tree["columns"] = list(self.dataframe.columns)
        self.tree.column("#0", width=50, anchor=tk.CENTER)
        self.tree.heading("#0", text="Fila")

        for col in self.dataframe.columns:
            self.tree.column(col, width=110, anchor=tk.W)
            self.tree.heading(col, text=str(col))

        for idx, row in self.dataframe.iterrows():
            item_id = self.tree.insert("", tk.END, text=str(idx),
                                       values=[str(v) for v in row])
            self.row_id_map[item_id] = idx

    # ══════════════════════════════════════════════════════════════════════════
    #  EVENTOS DE TABLA
    # ══════════════════════════════════════════════════════════════════════════

    def _resolve_cell(self, event) -> tuple | None:
        """Devuelve (row_index, col_index, item_id) o None si no es una celda."""
        region = self.tree.identify_region(event.x, event.y)
        col    = self.tree.identify_column(event.x)
        row    = self.tree.identify_row(event.y)

        if region != "cell" or not row:
            return None

        col_idx = int(col) - 1
        if col_idx < 0 or row not in self.row_id_map:
            return None

        return self.row_id_map[row], col_idx, row

    def _on_cell_click(self, event) -> None:
        """Clic simple: carga los datos de la celda en el panel Editar."""
        result = self._resolve_cell(event)
        if result is None:
            return

        row_index, col_index, item_id = result
        self.selected_cell = (row_index, col_index)

        current_value = self.tree.item(item_id)["values"][col_index]
        col_name      = self.dataframe.columns[col_index]

        self.cell_info_label.config(text=f"Celda: {col_name}  (Fila {row_index})")
        self.value_entry.delete(0, tk.END)
        self.value_entry.insert(0, str(current_value))

        # Cargar formato guardado o restablecer valores por defecto
        fmt = self.cell_formats.get((row_index, col_index), {})
        self.font_var.set(fmt.get("font",          "Calibri"))
        self.font_size_var.set(str(fmt.get("size", FONT_SIZE_DEFAULT)))
        self.h_align_var.set(fmt.get("h_align",    "left"))
        self.v_align_var.set(fmt.get("v_align",    "center"))
        self.number_format_var.set(fmt.get("number_format", "General"))

        self.apply_btn.config(state=tk.NORMAL)

    def _on_cell_double_click(self, event) -> None:
        """Doble clic: abre ventana de edición rápida del valor."""
        result = self._resolve_cell(event)
        if result is None:
            return

        row_index, col_index, item_id = result
        current_value = self.tree.item(item_id)["values"][col_index]
        self._open_edit_window(row_index, col_index, current_value)

    # ══════════════════════════════════════════════════════════════════════════
    #  EDICIÓN
    # ══════════════════════════════════════════════════════════════════════════

    def _open_edit_window(self, row_index: int, col_index: int,
                          current_value) -> None:
        """Ventana emergente para editar el valor de la celda."""
        win = tk.Toplevel(self.root)
        win.title("Editar celda")
        win.geometry("320x110")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        col_name = self.dataframe.columns[col_index]
        ttk.Label(win, text=f"Columna: {col_name}  — Fila: {row_index}").pack(pady=5)

        entry = ttk.Entry(win)
        entry.insert(0, str(current_value))
        entry.pack(pady=4, padx=10, fill=tk.X)
        entry.focus()

        def _save():
            new_val = entry.get()
            self.dataframe.at[row_index, col_name] = new_val
            self.edited_cells[(row_index, col_index)] = new_val
            self._display_data()
            win.destroy()

        btn_row = ttk.Frame(win)
        btn_row.pack(pady=6)
        ttk.Button(btn_row, text="Guardar",   command=_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Cancelar",  command=win.destroy).pack(side=tk.LEFT, padx=5)
        entry.bind("<Return>", lambda _: _save())

    def _apply_formatting(self) -> None:
        """Guarda el formato y el valor del panel Editar en la celda seleccionada."""
        if self.selected_cell is None:
            messagebox.showwarning("Atención", "Selecciona una celda primero.")
            return

        row_index, col_index = self.selected_cell

        # Validar tamaño de fuente
        try:
            size = int(self.font_size_var.get())
            if not (FONT_SIZE_MIN <= size <= FONT_SIZE_MAX):
                raise ValueError
        except ValueError:
            messagebox.showerror("Error",
                f"El tamaño de fuente debe estar entre {FONT_SIZE_MIN} y {FONT_SIZE_MAX}.")
            return

        # Guardar formato
        self.cell_formats[(row_index, col_index)] = {
            "font":          self.font_var.get(),
            "size":          size,
            "h_align":       self.h_align_var.get(),
            "v_align":       self.v_align_var.get(),
            "number_format": self.number_format_var.get(),
        }

        # Actualizar valor en el DataFrame
        col_name  = self.dataframe.columns[col_index]
        new_value = self.value_entry.get()
        self.dataframe.at[row_index, col_name] = new_value
        self.edited_cells[(row_index, col_index)] = new_value

        self._display_data()
        messagebox.showinfo("Aplicado", "Formato aplicado correctamente.")

    # ══════════════════════════════════════════════════════════════════════════
    #  GUARDADO
    # ══════════════════════════════════════════════════════════════════════════

    def save_file(self) -> None:
        """Guarda el DataFrame y aplica los formatos almacenados en el XLSX."""
        if self.current_file is None or self.dataframe is None:
            messagebox.showwarning("Atención", "No hay ningún archivo cargado.")
            return

        try:
            # 1) Volcar datos con pandas
            self.dataframe.to_excel(self.current_file, index=False)

            # 2) Aplicar formatos con openpyxl
            wb = load_workbook(self.current_file)
            ws = wb.active

            for (row_index, col_index), fmt in self.cell_formats.items():
                # openpyxl es 1-based; la fila 1 es el encabezado → +2
                cell = ws.cell(row=row_index + 2, column=col_index + 1)

                # Fuente
                cell.font = Font(
                    name=fmt.get("font", "Calibri"),
                    size=fmt.get("size", FONT_SIZE_DEFAULT),
                )

                # Alineación
                h = fmt.get("h_align", "left")
                v = fmt.get("v_align", "center")
                cell.alignment = Alignment(
                    horizontal=h if h in ALIGN_OPTIONS  else "left",
                    vertical=  v if v in VALIGN_OPTIONS else "center",
                    wrap_text=True,
                )

                # Formato de número
                fmt_key = fmt.get("number_format", "General")
                cell.number_format = NUMBER_FORMAT_MAP.get(fmt_key, "General")

            wb.save(self.current_file)
            messagebox.showinfo("Guardado",
                                "Archivo guardado correctamente con formato.")
            self.edited_cells = {}
            self.cell_formats = {}

        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{exc}")


# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app  = DataCleanerApp(root)
    root.mainloop()
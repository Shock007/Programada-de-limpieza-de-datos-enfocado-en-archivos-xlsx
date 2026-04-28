import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# ──────────────────────────────────────────────────────────────────────────────
#  Constantes globales
# ──────────────────────────────────────────────────────────────────────────────
NUMBER_FORMAT_MAP = {
    "General":               "General",
    "Número":                "#,##0.00",
    "Moneda Colombiana":     "[$$-2058]#,##0.00;[RED]-[$$-2058]#,##0.00",
    "Moneda Estadounidense": "$#,##0.00",
    "Fecha":                 "dd/mm/yyyy",
    "Hora":                  "hh:mm:ss",
    "Porcentaje":            "0.00%",
    "Texto":                 "@",
}

FONTS_AVAILABLE   = ["Arial", "Calibri"]
ALIGN_OPTIONS     = ["left", "center", "right"]
VALIGN_OPTIONS    = ["top", "center", "bottom"]
FONT_SIZE_DEFAULT = 11
FONT_SIZE_MIN     = 8
FONT_SIZE_MAX     = 72

COLOR_CELL_SELECT = "#cce5ff"   # azul claro — celda seleccionada
COLOR_COL_SELECT  = "#d4edda"   # verde claro — columna seleccionada


class DataCleanerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("XLSX Data Cleaner")
        self.root.geometry("1200x740")

        # ── Estado interno ──────────────────────────────────────────────────
        self.current_file:    str | None          = None
        self.dataframe:       pd.DataFrame | None = None
        self.edited_cells:    dict                = {}
        self.cell_formats:    dict                = {}
        self.selected_cell:   tuple | None        = None  # (row_idx, col_idx)
        self.selected_column: int   | None        = None  # col_idx (columna entera)
        self.row_id_map:      dict                = {}    # item_id → df row index

        self._build_ui()

    # ══════════════════════════════════════════════════════════════════════════
    #  CONSTRUCCIÓN DE LA INTERFAZ
    # ══════════════════════════════════════════════════════════════════════════

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._build_toolbar(main)
        self._build_content(main)

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

    def _build_content(self, parent: ttk.Frame) -> None:
        content = ttk.Frame(parent)
        content.pack(fill=tk.BOTH, expand=True)
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(0, weight=1)
        self._build_table(content)
        self._build_edit_panel(content)

    # ── Tabla ─────────────────────────────────────────────────────────────────

    def _build_table(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 8))
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(
            frame, yscrollcommand=vsb.set,
            xscrollcommand=hsb.set, selectmode="browse"
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)

        # Tags de resaltado visual
        self.tree.tag_configure("cell_selected", background=COLOR_CELL_SELECT)
        self.tree.tag_configure("col_selected",  background=COLOR_COL_SELECT)

        # Un solo binding para clic (heading y celda se discriminan dentro)
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<Double-1>", self._on_double_click)

    # ── Panel Editar ──────────────────────────────────────────────────────────

    def _build_edit_panel(self, parent: ttk.Frame) -> None:
        panel = ttk.LabelFrame(parent, text="Editar", padding=12)
        panel.grid(row=0, column=1, sticky=tk.NSEW, padx=4)

        self.cell_info_label = ttk.Label(
            panel, text="Ninguna celda seleccionada",
            relief=tk.SUNKEN, anchor=tk.W)
        self.cell_info_label.pack(fill=tk.X, pady=(0, 10))

        self._section_label(panel, "Valor de la celda")
        self.value_entry = ttk.Entry(panel, width=22)
        self.value_entry.pack(fill=tk.X, pady=(0, 10))

        ttk.Separator(panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        self._section_label(panel, "Fuente")
        self.font_var = tk.StringVar(value="Calibri")
        ttk.Combobox(panel, textvariable=self.font_var, width=20,
                     values=FONTS_AVAILABLE, state="readonly").pack(fill=tk.X, pady=(0, 10))

        self._section_label(panel, "Tamaño de fuente")
        self.font_size_var = tk.StringVar(value=str(FONT_SIZE_DEFAULT))
        ttk.Spinbox(panel, from_=FONT_SIZE_MIN, to=FONT_SIZE_MAX,
                    textvariable=self.font_size_var,
                    width=20).pack(fill=tk.X, pady=(0, 10))

        ttk.Separator(panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        self._section_label(panel, "Alineación horizontal")
        self.h_align_var = tk.StringVar(value="left")
        frm_h = ttk.Frame(panel)
        frm_h.pack(fill=tk.X, pady=(0, 10))
        for lbl, val in [("← Izquierda", "left"),
                          ("↔ Centro",    "center"),
                          ("→ Derecha",   "right")]:
            ttk.Radiobutton(frm_h, text=lbl, variable=self.h_align_var,
                            value=val).pack(anchor=tk.W)

        self._section_label(panel, "Alineación vertical")
        self.v_align_var = tk.StringVar(value="center")
        frm_v = ttk.Frame(panel)
        frm_v.pack(fill=tk.X, pady=(0, 10))
        for lbl, val in [("↑ Superior", "top"),
                          ("↕ Medio",   "center"),
                          ("↓ Inferior","bottom")]:
            ttk.Radiobutton(frm_v, text=lbl, variable=self.v_align_var,
                            value=val).pack(anchor=tk.W)

        ttk.Separator(panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        self._section_label(panel, "Formato de número")
        self.number_format_var = tk.StringVar(value="General")
        ttk.Combobox(panel, textvariable=self.number_format_var, width=20,
                     values=list(NUMBER_FORMAT_MAP.keys()),
                     state="readonly").pack(fill=tk.X, pady=(0, 10))

        ttk.Separator(panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        self.apply_btn = ttk.Button(panel, text="✔ Aplicar formato",
                                    command=self._apply_formatting,
                                    state=tk.DISABLED)
        self.apply_btn.pack(fill=tk.X, pady=(4, 0))

    @staticmethod
    def _section_label(parent: ttk.Frame, text: str) -> None:
        ttk.Label(parent, text=text,
                  font=("TkDefaultFont", 8, "bold")).pack(anchor=tk.W, pady=(4, 2))

    def _cast_value(self, col_index: int, raw: str):
        """
        Convierte el string 'raw' al dtype original de la columna en el DataFrame.
        Si la conversión falla (e.g. el usuario escribió texto en una col numérica),
        retorna el string tal cual para que pandas lo convierta a object.
        """
        col_name = self.dataframe.columns[col_index]
        dtype    = self.dataframe[col_name].dtype

        try:
            if pd.api.types.is_integer_dtype(dtype):
                # Acepta "9", "9.0", "9,0" → int
                return int(float(raw.replace(",", ".")))
            elif pd.api.types.is_float_dtype(dtype):
                return float(raw.replace(",", "."))
            elif pd.api.types.is_bool_dtype(dtype):
                return raw.strip().lower() in ("true", "1", "sí", "si", "yes")
            else:
                return raw          # string, datetime, etc. → sin conversión
        except (ValueError, AttributeError):
            return raw              # fallback: dejar como string

    # ══════════════════════════════════════════════════════════════════════════
    #  CARGA Y VISUALIZACIÓN
    # ══════════════════════════════════════════════════════════════════════════

    def upload_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar archivo XLSX",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if not path:
            return
        try:
            self.current_file    = path
            self.dataframe       = pd.read_excel(path)
            self.edited_cells    = {}
            self.cell_formats    = {}
            self.selected_cell   = None
            self.selected_column = None
            self._display_data()
            self.file_label.config(text=f"Cargado: {os.path.basename(path)}")
            self.save_btn.config(state=tk.NORMAL)
            messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{exc}")

    def _display_data(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.row_id_map = {}

        if self.dataframe is None:
            return

        cols = list(self.dataframe.columns)
        self.tree["columns"] = cols

        self.tree.column("#0", width=50, anchor=tk.CENTER, stretch=False)
        self.tree.heading("#0", text="Fila")

        for i, col in enumerate(cols):
            self.tree.column(col, width=110, anchor=tk.W, minwidth=60)
            # Cada encabezado registra su propio command con índice y nombre capturados
            self.tree.heading(
                col, text=str(col),
                command=lambda idx=i, name=col: self._on_column_header_click(idx, name)
            )

        for idx, row in self.dataframe.iterrows():
            item_id = self.tree.insert("", tk.END, text=str(idx),
                                       values=[str(v) for v in row])
            self.row_id_map[item_id] = idx

    # ══════════════════════════════════════════════════════════════════════════
    #  RESOLUCIÓN DE CLICS
    # ══════════════════════════════════════════════════════════════════════════

    def _parse_col_index(self, event) -> int:
        """
        Convierte '#N' (salida de identify_column) en índice 0-based de datos.

        El BUG original era hacer int(col) sobre el string '#1', '#2', etc.,
        lo que lanza ValueError. La corrección es lstrip('#') antes de int().

        Retorna -1 si el clic fue en la columna de árbol (#0).
        """
        raw = self.tree.identify_column(event.x)  # '#0', '#1', '#2', …
        num = int(raw.lstrip("#"))                 # ← FIX principal
        return num - 1                             # 0-based; -1 para col de árbol

    def _on_click(self, event) -> None:
        """Dispatcher de clic simple: delega a celda o ignora el encabezado."""
        region = self.tree.identify_region(event.x, event.y)

        # Los encabezados ya tienen su command= registrado en heading();
        # solo necesitamos manejar clics en celdas de datos.
        if region == "heading":
            return

        if region != "cell":
            return

        item_id = self.tree.identify_row(event.y)
        if not item_id or item_id not in self.row_id_map:
            return

        col_idx = self._parse_col_index(event)
        if col_idx < 0:
            return   # clic en la columna "#0" (número de fila), ignorar

        self._select_cell(item_id, col_idx)

    def _on_double_click(self, event) -> None:
        """Doble clic: abre editor emergente de valor."""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        item_id = self.tree.identify_row(event.y)
        if not item_id or item_id not in self.row_id_map:
            return

        col_idx = self._parse_col_index(event)
        if col_idx < 0:
            return

        row_index = self.row_id_map[item_id]
        current   = self.tree.item(item_id)["values"][col_idx]
        self._open_edit_window(row_index, col_idx, current)

    # ── Selección de celda ────────────────────────────────────────────────────

    def _select_cell(self, item_id: str, col_idx: int) -> None:
        row_index = self.row_id_map[item_id]
        self.selected_cell   = (row_index, col_idx)
        self.selected_column = None

        self._clear_highlight()
        self.tree.item(item_id, tags=("cell_selected",))

        col_name      = self.dataframe.columns[col_idx]
        current_value = self.tree.item(item_id)["values"][col_idx]

        self.cell_info_label.config(
            text=f"Celda — Columna: {col_name}  |  Fila: {row_index}")

        self.value_entry.config(state=tk.NORMAL)
        self.value_entry.delete(0, tk.END)
        self.value_entry.insert(0, str(current_value))

        fmt = self.cell_formats.get((row_index, col_idx), {})
        self.font_var.set(fmt.get("font",          "Calibri"))
        self.font_size_var.set(str(fmt.get("size", FONT_SIZE_DEFAULT)))
        self.h_align_var.set(fmt.get("h_align",    "left"))
        self.v_align_var.set(fmt.get("v_align",    "center"))
        self.number_format_var.set(fmt.get("number_format", "General"))

        self.apply_btn.config(state=tk.NORMAL)

    # ── Selección de columna ──────────────────────────────────────────────────

    def _on_column_header_click(self, col_idx: int, col_name: str) -> None:
        """Invocado por el command= de cada encabezado de columna."""
        self.selected_cell   = None
        self.selected_column = col_idx

        self._clear_highlight()
        for item_id in self.tree.get_children():
            self.tree.item(item_id, tags=("col_selected",))

        n_rows = len(self.dataframe)
        self.cell_info_label.config(
            text=f"Columna seleccionada: {col_name}  ({n_rows} filas)")

        self.value_entry.delete(0, tk.END)
        self.value_entry.config(state=tk.DISABLED)

        fmt = self.cell_formats.get((None, col_idx), {})
        self.font_var.set(fmt.get("font",          "Calibri"))
        self.font_size_var.set(str(fmt.get("size", FONT_SIZE_DEFAULT)))
        self.h_align_var.set(fmt.get("h_align",    "left"))
        self.v_align_var.set(fmt.get("v_align",    "center"))
        self.number_format_var.set(fmt.get("number_format", "General"))

        self.apply_btn.config(state=tk.NORMAL)

    def _clear_highlight(self) -> None:
        for item_id in self.tree.get_children():
            self.tree.item(item_id, tags=())

    # ══════════════════════════════════════════════════════════════════════════
    #  EDICIÓN
    # ══════════════════════════════════════════════════════════════════════════

    def _open_edit_window(self, row_index: int, col_index: int,
                          current_value) -> None:
        win = tk.Toplevel(self.root)
        win.title("Editar celda")
        win.geometry("320x115")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        col_name = self.dataframe.columns[col_index]
        ttk.Label(win, text=f"Columna: {col_name}  —  Fila: {row_index}").pack(pady=6)

        entry = ttk.Entry(win)
        entry.insert(0, str(current_value))
        entry.pack(pady=4, padx=10, fill=tk.X)
        entry.select_range(0, tk.END)
        entry.focus()

        def _save():
            new_val = self._cast_value(col_index, entry.get())
            self.dataframe.at[row_index, col_name] = new_val
            self.edited_cells[(row_index, col_index)] = new_val
            self._display_data()
            win.destroy()

        btn_row = ttk.Frame(win)
        btn_row.pack(pady=6)
        ttk.Button(btn_row, text="Guardar",  command=_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Cancelar", command=win.destroy).pack(side=tk.LEFT, padx=5)
        entry.bind("<Return>", lambda _: _save())
        entry.bind("<Escape>", lambda _: win.destroy())

    def _apply_formatting(self) -> None:
        try:
            size = int(self.font_size_var.get())
            if not (FONT_SIZE_MIN <= size <= FONT_SIZE_MAX):
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Error",
                f"El tamaño debe estar entre {FONT_SIZE_MIN} y {FONT_SIZE_MAX}.")
            return

        fmt = {
            "font":          self.font_var.get(),
            "size":          size,
            "h_align":       self.h_align_var.get(),
            "v_align":       self.v_align_var.get(),
            "number_format": self.number_format_var.get(),
        }

        if self.selected_cell is not None:
            row_index, col_index = self.selected_cell
            self.cell_formats[(row_index, col_index)] = fmt

            col_name  = self.dataframe.columns[col_index]
            new_value = self._cast_value(col_index, self.value_entry.get())
            self.dataframe.at[row_index, col_name] = new_value
            self.edited_cells[(row_index, col_index)] = new_value

            self._display_data()
            messagebox.showinfo("Aplicado", "Formato aplicado a la celda.")

        elif self.selected_column is not None:
            col_index = self.selected_column
            for idx in self.dataframe.index:
                self.cell_formats[(idx, col_index)] = fmt
            self.cell_formats[(None, col_index)] = fmt   # marca de columna completa

            self._display_data()
            col_name = self.dataframe.columns[col_index]
            messagebox.showinfo(
                "Aplicado",
                f"Formato aplicado a toda la columna '{col_name}'.")
        else:
            messagebox.showwarning("Atención", "Selecciona una celda o columna primero.")

    # ══════════════════════════════════════════════════════════════════════════
    #  GUARDADO
    # ══════════════════════════════════════════════════════════════════════════

    def save_file(self) -> None:
        if self.current_file is None or self.dataframe is None:
            messagebox.showwarning("Atención", "No hay ningún archivo cargado.")
            return

        try:
            self.dataframe.to_excel(self.current_file, index=False)

            wb = load_workbook(self.current_file)
            ws = wb.active

            for (row_index, col_index), fmt in self.cell_formats.items():
                if row_index is None:
                    continue   # clave de columna completa, ya expandida por fila
                cell = ws.cell(row=row_index + 2, column=col_index + 1)

                cell.font = Font(
                    name=fmt.get("font", "Calibri"),
                    size=fmt.get("size", FONT_SIZE_DEFAULT),
                )

                h = fmt.get("h_align", "left")
                v = fmt.get("v_align", "center")
                cell.alignment = Alignment(
                    horizontal=h if h in ALIGN_OPTIONS  else "left",
                    vertical=  v if v in VALIGN_OPTIONS else "center",
                    wrap_text=True,
                )

                fmt_key = fmt.get("number_format", "General")
                cell.number_format = NUMBER_FORMAT_MAP.get(fmt_key, "General")

            wb.save(self.current_file)
            messagebox.showinfo("Guardado", "Archivo guardado con formato correctamente.")
            self.edited_cells = {}
            self.cell_formats = {}

        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{exc}")


# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app  = DataCleanerApp(root)
    root.mainloop()
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
        self._exact_dup_col_pairs: list           = []    # [(orig, copia), …]

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

        # ── Notebook derecho: "Editar" | "Automatización" ─────────────────
        self.notebook = ttk.Notebook(content)
        self.notebook.grid(row=0, column=1, sticky=tk.NSEW, padx=4)

        # Pestaña Editar
        tab_edit = ttk.Frame(self.notebook, padding=0)
        self.notebook.add(tab_edit, text="  Editar  ")
        self._build_edit_panel(tab_edit)

        # Pestaña Automatización
        tab_auto = ttk.Frame(self.notebook, padding=0)
        self.notebook.add(tab_auto, text="  Automatización  ")
        self._build_auto_panel(tab_auto)

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
        panel.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

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

    # ══════════════════════════════════════════════════════════════════════════
    #  PESTAÑA AUTOMATIZACIÓN
    # ══════════════════════════════════════════════════════════════════════════

    def _build_auto_panel(self, parent: ttk.Frame) -> None:
        """
        Panel de Automatización con tres secciones:
          1. Vacíos — detección y eliminación de celdas/columnas vacías
          2. Duplicados — detección y eliminación de filas duplicadas
        """
        # Contenedor con scroll para que todo quepa en cualquier tamaño de ventana
        canvas = tk.Canvas(parent, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor=tk.NW)

        def _on_resize(event):
            canvas.itemconfig(inner_id, width=event.width)
        def _on_frame_configure(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<Configure>", _on_resize)
        inner.bind("<Configure>", _on_frame_configure)

        # ════════════════════════════════════════════════════════════════════
        # SECCIÓN 1 — VACÍOS
        # ════════════════════════════════════════════════════════════════════
        sec_empty = ttk.LabelFrame(inner, text="Celdas y columnas vacías", padding=10)
        sec_empty.pack(fill=tk.X, padx=6, pady=(8, 4))

        # ── Totales ──────────────────────────────────────────────────────────
        self._section_label(sec_empty, "Resumen general")

        self.lbl_total_empty_cells = ttk.Label(
            sec_empty, text="Celdas vacías: —",
            foreground="#c0392b", font=("TkDefaultFont", 9, "bold"))
        self.lbl_total_empty_cells.pack(anchor=tk.W, pady=1)

        self.lbl_total_empty_cols = ttk.Label(
            sec_empty, text="Columnas vacías: —",
            foreground="#8e44ad", font=("TkDefaultFont", 9, "bold"))
        self.lbl_total_empty_cols.pack(anchor=tk.W, pady=1)

        ttk.Separator(sec_empty, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Detalle celdas vacías ─────────────────────────────────────────────
        self._section_label(sec_empty, "Celdas vacías por posición")

        cell_frame = ttk.Frame(sec_empty)
        cell_frame.pack(fill=tk.X, pady=(2, 6))
        cell_frame.grid_rowconfigure(0, weight=1)
        cell_frame.grid_columnconfigure(0, weight=1)

        cell_vsb = ttk.Scrollbar(cell_frame, orient=tk.VERTICAL)
        self.auto_cell_text = tk.Text(
            cell_frame, height=5, wrap=tk.WORD, state=tk.DISABLED,
            relief=tk.FLAT, bg="#f8f9fa", font=("TkDefaultFont", 9),
            yscrollcommand=cell_vsb.set)
        cell_vsb.config(command=self.auto_cell_text.yview)
        self.auto_cell_text.grid(row=0, column=0, sticky=tk.NSEW)
        cell_vsb.grid(row=0, column=1, sticky=tk.NS)

        # Botón eliminar filas con celdas vacías
        self.btn_del_empty_rows = ttk.Button(
            sec_empty, text="🗑 Eliminar filas con celdas vacías",
            command=self._delete_rows_with_empty_cells, state=tk.DISABLED)
        self.btn_del_empty_rows.pack(fill=tk.X, pady=(2, 6))

        ttk.Separator(sec_empty, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Detalle columnas vacías ───────────────────────────────────────────
        self._section_label(sec_empty, "Columnas completamente vacías")

        col_frame = ttk.Frame(sec_empty)
        col_frame.pack(fill=tk.X, pady=(2, 6))
        col_frame.grid_rowconfigure(0, weight=1)
        col_frame.grid_columnconfigure(0, weight=1)

        col_vsb = ttk.Scrollbar(col_frame, orient=tk.VERTICAL)
        self.auto_col_text = tk.Text(
            col_frame, height=4, wrap=tk.WORD, state=tk.DISABLED,
            relief=tk.FLAT, bg="#f8f9fa", font=("TkDefaultFont", 9),
            yscrollcommand=col_vsb.set)
        col_vsb.config(command=self.auto_col_text.yview)
        self.auto_col_text.grid(row=0, column=0, sticky=tk.NSEW)
        col_vsb.grid(row=0, column=1, sticky=tk.NS)

        # Botón eliminar columnas vacías
        self.btn_del_empty_cols = ttk.Button(
            sec_empty, text="🗑 Eliminar columnas vacías",
            command=self._delete_empty_columns, state=tk.DISABLED)
        self.btn_del_empty_cols.pack(fill=tk.X, pady=(2, 6))

        ttk.Separator(sec_empty, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # Botón analizar vacíos
        ttk.Button(
            sec_empty, text="🔍 Analizar vacíos",
            command=self._refresh_auto_panel
        ).pack(fill=tk.X, pady=(0, 2))

        # ════════════════════════════════════════════════════════════════════
        # SECCIÓN 2 — DUPLICADOS
        # ════════════════════════════════════════════════════════════════════
        sec_dup = ttk.LabelFrame(inner, text="Registros duplicados", padding=10)
        sec_dup.pack(fill=tk.X, padx=6, pady=(4, 8))

        # ── Resumen general ──────────────────────────────────────────────────
        self._section_label(sec_dup, "Resumen general")

        self.lbl_total_dups = ttk.Label(
            sec_dup, text="Filas duplicadas: —",
            foreground="#e67e22", font=("TkDefaultFont", 9, "bold"))
        self.lbl_total_dups.pack(anchor=tk.W, pady=1)

        self.lbl_unique_rows = ttk.Label(
            sec_dup, text="Filas únicas: —",
            foreground="#27ae60", font=("TkDefaultFont", 9, "bold"))
        self.lbl_unique_rows.pack(anchor=tk.W, pady=1)

        self.lbl_total_dup_vals = ttk.Label(
            sec_dup, text="Valores duplicados: —",
            foreground="#c0392b", font=("TkDefaultFont", 9, "bold"))
        self.lbl_total_dup_vals.pack(anchor=tk.W, pady=1)

        ttk.Separator(sec_dup, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Filas duplicadas ─────────────────────────────────────────────────
        self._section_label(sec_dup, "Posición de filas duplicadas")

        dup_frame = ttk.Frame(sec_dup)
        dup_frame.pack(fill=tk.X, pady=(2, 4))
        dup_frame.grid_rowconfigure(0, weight=1)
        dup_frame.grid_columnconfigure(0, weight=1)

        dup_vsb = ttk.Scrollbar(dup_frame, orient=tk.VERTICAL)
        self.auto_dup_text = tk.Text(
            dup_frame, height=4, wrap=tk.WORD, state=tk.DISABLED,
            relief=tk.FLAT, bg="#fff8f0", font=("TkDefaultFont", 9),
            yscrollcommand=dup_vsb.set)
        dup_vsb.config(command=self.auto_dup_text.yview)
        self.auto_dup_text.grid(row=0, column=0, sticky=tk.NSEW)
        dup_vsb.grid(row=0, column=1, sticky=tk.NS)

        self.btn_del_dups = ttk.Button(
            sec_dup, text="🗑 Eliminar filas duplicadas",
            command=self._delete_duplicates, state=tk.DISABLED)
        self.btn_del_dups.pack(fill=tk.X, pady=(2, 6))

        ttk.Separator(sec_dup, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Valores duplicados en el archivo ─────────────────────────────────
        self._section_label(sec_dup, "Valores duplicados en el archivo")

        val_dup_frame = ttk.Frame(sec_dup)
        val_dup_frame.pack(fill=tk.X, pady=(2, 6))
        val_dup_frame.grid_rowconfigure(0, weight=1)
        val_dup_frame.grid_columnconfigure(0, weight=1)

        val_dup_vsb = ttk.Scrollbar(val_dup_frame, orient=tk.VERTICAL)
        self.auto_val_dup_text = tk.Text(
            val_dup_frame, height=5, wrap=tk.WORD, state=tk.DISABLED,
            relief=tk.FLAT, bg="#fef9e7", font=("TkDefaultFont", 9),
            yscrollcommand=val_dup_vsb.set)
        val_dup_vsb.config(command=self.auto_val_dup_text.yview)
        self.auto_val_dup_text.grid(row=0, column=0, sticky=tk.NSEW)
        val_dup_vsb.grid(row=0, column=1, sticky=tk.NS)

        ttk.Separator(sec_dup, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Análisis de columnas duplicadas ──────────────────────────────────
        self._section_label(sec_dup, "Análisis de columnas duplicadas")

        col_dup_frame = ttk.Frame(sec_dup)
        col_dup_frame.pack(fill=tk.X, pady=(2, 6))
        col_dup_frame.grid_rowconfigure(0, weight=1)
        col_dup_frame.grid_columnconfigure(0, weight=1)

        col_dup_vsb = ttk.Scrollbar(col_dup_frame, orient=tk.VERTICAL)
        self.auto_col_dup_text = tk.Text(
            col_dup_frame, height=5, wrap=tk.WORD, state=tk.DISABLED,
            relief=tk.FLAT, bg="#eaf4fb", font=("TkDefaultFont", 9),
            yscrollcommand=col_dup_vsb.set)
        col_dup_vsb.config(command=self.auto_col_dup_text.yview)
        self.auto_col_dup_text.grid(row=0, column=0, sticky=tk.NSEW)
        col_dup_vsb.grid(row=0, column=1, sticky=tk.NS)

        ttk.Separator(sec_dup, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Posición de columnas duplicadas ──────────────────────────────────
        self._section_label(sec_dup, "Posición de columnas duplicadas")

        col_pos_dup_frame = ttk.Frame(sec_dup)
        col_pos_dup_frame.pack(fill=tk.X, pady=(2, 4))
        col_pos_dup_frame.grid_rowconfigure(0, weight=1)
        col_pos_dup_frame.grid_columnconfigure(0, weight=1)

        col_pos_dup_vsb = ttk.Scrollbar(col_pos_dup_frame, orient=tk.VERTICAL)
        self.auto_col_pos_dup_text = tk.Text(
            col_pos_dup_frame, height=4, wrap=tk.WORD, state=tk.DISABLED,
            relief=tk.FLAT, bg="#f0f8ff", font=("TkDefaultFont", 9),
            yscrollcommand=col_pos_dup_vsb.set)
        col_pos_dup_vsb.config(command=self.auto_col_pos_dup_text.yview)
        self.auto_col_pos_dup_text.grid(row=0, column=0, sticky=tk.NSEW)
        col_pos_dup_vsb.grid(row=0, column=1, sticky=tk.NS)

        self.btn_del_dup_cols = ttk.Button(
            sec_dup, text="🗑 Eliminar columnas duplicadas",
            command=self._delete_duplicate_columns, state=tk.DISABLED)
        self.btn_del_dup_cols.pack(fill=tk.X, pady=(2, 6))

        ttk.Separator(sec_dup, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        ttk.Button(
            sec_dup, text="🔍 Analizar duplicados",
            command=self._refresh_duplicates
        ).pack(fill=tk.X, pady=(0, 2))

    def _refresh_auto_panel(self) -> None:
        """
        Analiza el DataFrame en busca de celdas y columnas vacías y
        actualiza los widgets del panel de Automatización.
        """
        from openpyxl.utils import get_column_letter

        def _set_text(widget: tk.Text, msg: str) -> None:
            widget.config(state=tk.NORMAL)
            widget.delete("1.0", tk.END)
            widget.insert(tk.END, msg)
            widget.config(state=tk.DISABLED)

        if self.dataframe is None:
            _set_text(self.auto_cell_text, "Carga un archivo para analizar.")
            _set_text(self.auto_col_text,  "Carga un archivo para analizar.")
            self.lbl_total_empty_cells.config(text="Celdas vacías: —")
            self.lbl_total_empty_cols.config(text="Columnas vacías: —")
            self.btn_del_empty_rows.config(state=tk.DISABLED)
            self.btn_del_empty_cols.config(state=tk.DISABLED)
            return

        df   = self.dataframe
        cols = list(df.columns)

        def _is_empty(val) -> bool:
            if pd.isna(val):
                return True
            if isinstance(val, str) and val.strip() == "":
                return True
            return False

        # Columnas completamente vacías
        empty_cols = [col for col in cols if df[col].apply(_is_empty).all()]

        # Celdas vacías individuales
        empty_cells: list[str] = []
        for col_idx, col in enumerate(cols):
            col_letter = get_column_letter(col_idx + 1)
            for df_row_idx in df.index:
                if _is_empty(df.at[df_row_idx, col]):
                    excel_row = int(df_row_idx) + 2
                    empty_cells.append(f"{col_letter}{excel_row}")

        # Filas que tienen AL MENOS una celda vacía
        rows_with_empty = df[
            df.apply(lambda row: row.apply(_is_empty).any(), axis=1)
        ].index.tolist()

        # ── Actualizar totales ───────────────────────────────────────────────
        self.lbl_total_empty_cells.config(text=f"Celdas vacías: {len(empty_cells)}")
        self.lbl_total_empty_cols.config(text=f"Columnas vacías: {len(empty_cols)}")

        # Habilitar / deshabilitar botones de eliminación
        self.btn_del_empty_rows.config(
            state=tk.NORMAL if rows_with_empty else tk.DISABLED)
        self.btn_del_empty_cols.config(
            state=tk.NORMAL if empty_cols else tk.DISABLED)

        # ── Texto de celdas vacías ───────────────────────────────────────────
        if not empty_cells:
            cell_msg = "✔ No se encontraron celdas vacías en el archivo."
        else:
            from collections import defaultdict
            by_col: dict[str, list[str]] = defaultdict(list)
            for ref in empty_cells:
                letter = "".join(c for c in ref if c.isalpha())
                by_col[letter].append(ref)

            lines = []
            for letter, refs in by_col.items():
                col_idx  = ord(letter) - ord("A") if len(letter) == 1 \
                           else (ord(letter[0]) - ord("A") + 1) * 26 + (ord(letter[1]) - ord("A"))
                col_name = cols[col_idx] if col_idx < len(cols) else letter
                positions = ", ".join(refs)
                lines.append(
                    f'• Columna "{col_name}": '
                    f'{"celda vacía en" if len(refs) == 1 else "celdas vacías en"} '
                    f'{positions}.'
                )
            cell_msg = "\n".join(lines)

        _set_text(self.auto_cell_text, cell_msg)

        # ── Texto de columnas vacías ─────────────────────────────────────────
        if not empty_cols:
            col_msg = "✔ Todas las columnas contienen al menos un dato."
        else:
            col_msg = "\n".join(
                f'• La columna "{col}" no contiene ningún dato.' for col in empty_cols
            )

        _set_text(self.auto_col_text, col_msg)

    # ── Eliminación de vacíos ─────────────────────────────────────────────────

    def _delete_empty_columns(self) -> None:
        """Elimina del DataFrame todas las columnas donde TODOS los valores son vacíos."""
        if self.dataframe is None:
            return

        def _is_empty(val) -> bool:
            if pd.isna(val):
                return True
            return isinstance(val, str) and val.strip() == ""

        cols_to_drop = [
            col for col in self.dataframe.columns
            if self.dataframe[col].apply(_is_empty).all()
        ]

        if not cols_to_drop:
            messagebox.showinfo("Sin cambios", "No hay columnas completamente vacías.")
            return

        names = ", ".join(f'"{c}"' for c in cols_to_drop)
        confirm = messagebox.askyesno(
            "Confirmar eliminación",
            f"Se eliminarán {len(cols_to_drop)} columna(s) vacía(s):\n{names}\n\n"
            "¿Deseas continuar?"
        )
        if not confirm:
            return

        self.dataframe.drop(columns=cols_to_drop, inplace=True)
        # Limpiar formatos de celdas huérfanas
        self.cell_formats  = {}
        self.edited_cells  = {}
        self.selected_cell = None
        self.selected_column = None

        self._display_data()
        self._refresh_auto_panel()
        messagebox.showinfo(
            "Eliminadas",
            f"{len(cols_to_drop)} columna(s) vacía(s) eliminadas correctamente.")

    def _delete_rows_with_empty_cells(self) -> None:
        """Elimina todas las filas que contengan AL MENOS una celda vacía."""
        if self.dataframe is None:
            return

        def _is_empty(val) -> bool:
            if pd.isna(val):
                return True
            return isinstance(val, str) and val.strip() == ""

        mask = self.dataframe.apply(lambda row: row.apply(_is_empty).any(), axis=1)
        n_rows = int(mask.sum())

        if n_rows == 0:
            messagebox.showinfo("Sin cambios", "No hay filas con celdas vacías.")
            return

        confirm = messagebox.askyesno(
            "Confirmar eliminación",
            f"Se eliminarán {n_rows} fila(s) que contienen al menos una celda vacía.\n\n"
            "¿Deseas continuar?"
        )
        if not confirm:
            return

        self.dataframe = self.dataframe[~mask].reset_index(drop=True)
        self.cell_formats    = {}
        self.edited_cells    = {}
        self.selected_cell   = None
        self.selected_column = None

        self._display_data()
        self._refresh_auto_panel()
        messagebox.showinfo(
            "Eliminadas",
            f"{n_rows} fila(s) con celdas vacías eliminadas correctamente.")

    # ── Duplicados ────────────────────────────────────────────────────────────

    def _refresh_duplicates(self) -> None:
        """
        Análisis de duplicados en tres niveles:
          1. Filas completas duplicadas (toda la fila idéntica a otra).
          2. Valores duplicados en el archivo (valores que aparecen >1 vez).
          3. Columnas duplicadas o posiblemente duplicadas.
        """
        def _set_text(widget: tk.Text, msg: str) -> None:
            widget.config(state=tk.NORMAL)
            widget.delete("1.0", tk.END)
            widget.insert(tk.END, msg)
            widget.config(state=tk.DISABLED)

        # ── Sin archivo ──────────────────────────────────────────────────────
        if self.dataframe is None:
            for w in (self.auto_dup_text, self.auto_val_dup_text,
                      self.auto_col_dup_text, self.auto_col_pos_dup_text):
                _set_text(w, "Carga un archivo para analizar.")
            self.lbl_total_dups.config(text="Filas duplicadas: —")
            self.lbl_unique_rows.config(text="Filas únicas: —")
            self.lbl_total_dup_vals.config(text="Valores duplicados: —")
            self.btn_del_dups.config(state=tk.DISABLED)
            self.btn_del_dup_cols.config(state=tk.DISABLED)
            return

        df   = self.dataframe
        cols = list(df.columns)

        # ════════════════════════════════════════════════════════════════════
        # NIVEL 1 — Filas completas duplicadas
        # ════════════════════════════════════════════════════════════════════
        dup_mask    = df.duplicated(keep="first")
        dup_indices = df.index[dup_mask].tolist()
        n_dups      = len(dup_indices)
        n_unique    = len(df) - n_dups

        self.lbl_total_dups.config(text=f"Filas duplicadas: {n_dups}")
        self.lbl_unique_rows.config(text=f"Filas únicas: {n_unique}")
        self.btn_del_dups.config(state=tk.NORMAL if n_dups > 0 else tk.DISABLED)

        if n_dups == 0:
            _set_text(self.auto_dup_text,
                      "✔ No se encontraron filas completamente duplicadas.")
        else:
            MAX_ROWS = 150
            refs = [f"Fila {int(i) + 2}" for i in dup_indices[:MAX_ROWS]]
            msg  = f"Las siguientes {n_dups} fila(s) son idénticas a otra anterior:\n"
            msg += ", ".join(refs)
            if n_dups > MAX_ROWS:
                msg += f"\n… y {n_dups - MAX_ROWS} más."
            _set_text(self.auto_dup_text, msg)

        # ════════════════════════════════════════════════════════════════════
        # NIVEL 2 — Valores duplicados en el archivo
        # ════════════════════════════════════════════════════════════════════
        # Aplanar todos los valores en una sola Serie, contar frecuencias
        all_values = pd.Series(
            df.values.flatten()
        ).dropna().astype(str).str.strip()
        # Excluir strings vacíos
        all_values = all_values[all_values != ""]

        freq = all_values.value_counts()
        dup_vals = freq[freq > 1]           # valores que aparecen más de una vez

        self.lbl_total_dup_vals.config(
            text=f"Valores duplicados: {len(dup_vals)}")

        if dup_vals.empty:
            _set_text(self.auto_val_dup_text,
                      "✔ No se encontraron valores duplicados en el archivo.")
        else:
            MAX_VALS = 60
            lines = []
            for val, count in dup_vals.head(MAX_VALS).items():
                lines.append(f'• "{val}"  →  aparece {count} veces')
            msg = "\n".join(lines)
            if len(dup_vals) > MAX_VALS:
                msg += f"\n… y {len(dup_vals) - MAX_VALS} valores más."
            _set_text(self.auto_val_dup_text, msg)

        # ════════════════════════════════════════════════════════════════════
        # NIVEL 3 — Columnas duplicadas o posiblemente duplicadas
        # ════════════════════════════════════════════════════════════════════
        # Paso A: buscar columnas con valores exactamente iguales entre sí
        exact_dup_pairs: list[tuple[str, str]] = []
        already_paired: set[str] = set()

        for i in range(len(cols)):
            for j in range(i + 1, len(cols)):
                col_a, col_b = cols[i], cols[j]
                # Comparar ignorando el tipo (convertir a string para uniformidad)
                series_a = df[col_a].astype(str).reset_index(drop=True)
                series_b = df[col_b].astype(str).reset_index(drop=True)
                if series_a.equals(series_b):
                    exact_dup_pairs.append((col_a, col_b))
                    already_paired.add(col_a)
                    already_paired.add(col_b)

        # Paso B: entre columnas sin pareja exacta, encontrar la que tiene
        # más valores internos duplicados → "Posiblemente duplicada"
        non_paired = [c for c in cols if c not in already_paired]
        possibly_dup_col:  str | None = None
        possibly_dup_count: int       = 0

        for col in non_paired:
            n_internal_dups = int(df[col].astype(str).duplicated(keep="first").sum())
            if n_internal_dups > possibly_dup_count:
                possibly_dup_count = n_internal_dups
                possibly_dup_col   = col

        # ── Construir mensaje de análisis ─────────────────────────────────────
        col_lines: list[str] = []

        if not exact_dup_pairs and possibly_dup_col is None:
            col_lines.append("✔ No se detectaron columnas duplicadas o sospechosas.")
        else:
            if exact_dup_pairs:
                col_lines.append("📋 Columnas exactamente duplicadas:")
                for col_a, col_b in exact_dup_pairs:
                    col_lines.append(
                        f'  • La columna "{col_b}" es una copia exacta de "{col_a}".')

            if possibly_dup_col and possibly_dup_count > 0:
                pct = round(possibly_dup_count / len(df) * 100, 1)
                col_lines.append("")
                col_lines.append("⚠ Columna posiblemente duplicada:")
                col_lines.append(
                    f'  • "{possibly_dup_col}" tiene {possibly_dup_count} valor(es) '
                    f'repetidos internamente ({pct}% de sus filas).')
                col_lines.append(
                    "    (Es la columna con más repeticiones internas entre "
                    "las que no tienen copia exacta.)")

        _set_text(self.auto_col_dup_text, "\n".join(col_lines))

        # ── Posición de columnas duplicadas y botón de eliminación ────────────
        # Guardar los pares exactos en el estado para que el botón pueda usarlos
        self._exact_dup_col_pairs = exact_dup_pairs   # [(original, copia), …]

        if not exact_dup_pairs:
            _set_text(self.auto_col_pos_dup_text,
                      "✔ No hay columnas exactamente duplicadas.")
            self.btn_del_dup_cols.config(state=tk.DISABLED)
        else:
            from openpyxl.utils import get_column_letter
            col_list  = list(df.columns)
            pos_lines: list[str] = []
            cols_to_remove: list[str] = []

            for col_a, col_b in exact_dup_pairs:
                idx_a = col_list.index(col_a) + 1   # 1-based → letra Excel
                idx_b = col_list.index(col_b) + 1
                let_a = get_column_letter(idx_a)
                let_b = get_column_letter(idx_b)
                pos_lines.append(
                    f'• "{col_b}" (columna {let_b})  es copia de  '
                    f'"{col_a}" (columna {let_a})  → se eliminará "{col_b}".')
                cols_to_remove.append(col_b)

            n_copies = len(cols_to_remove)
            pos_lines.insert(0,
                f"Se detectaron {n_copies} columna(s) duplicada(s) "
                f"(se conservará la primera de cada par):\n")
            _set_text(self.auto_col_pos_dup_text, "\n".join(pos_lines))
            self.btn_del_dup_cols.config(state=tk.NORMAL)

    def _delete_duplicates(self) -> None:
        """Elimina todas las filas duplicadas conservando la primera ocurrencia."""
        if self.dataframe is None:
            return

        dup_mask = self.dataframe.duplicated(keep="first")
        n_dups   = int(dup_mask.sum())

        if n_dups == 0:
            messagebox.showinfo("Sin cambios", "No hay filas duplicadas.")
            return

        confirm = messagebox.askyesno(
            "Confirmar eliminación",
            f"Se eliminarán {n_dups} fila(s) duplicadas, "
            "conservando la primera ocurrencia de cada registro.\n\n"
            "¿Deseas continuar?"
        )
        if not confirm:
            return

        self.dataframe = self.dataframe[~dup_mask].reset_index(drop=True)
        self.cell_formats    = {}
        self.edited_cells    = {}
        self.selected_cell   = None
        self.selected_column = None

        self._display_data()
        self._refresh_duplicates()
        self._refresh_auto_panel()
        messagebox.showinfo(
            "Eliminadas",
            f"{n_dups} fila(s) duplicada(s) eliminadas correctamente.")

    def _delete_duplicate_columns(self) -> None:
        """
        Elimina las columnas que son copias exactas de otra columna anterior,
        conservando siempre la primera ocurrencia de cada par.
        Los pares se calcularon en _refresh_duplicates y se almacenaron en
        self._exact_dup_col_pairs.
        """
        if self.dataframe is None:
            return

        pairs = getattr(self, "_exact_dup_col_pairs", [])
        if not pairs:
            messagebox.showinfo("Sin cambios",
                                "No hay columnas duplicadas exactas detectadas.\n"
                                "Ejecuta 'Analizar duplicados' primero.")
            return

        # Las columnas a eliminar son las copias (segundo elemento de cada par)
        cols_to_drop = [col_b for _, col_b in pairs]
        # Eliminar duplicados de la lista (por si una columna apareció en varios pares)
        cols_to_drop = list(dict.fromkeys(cols_to_drop))

        names = "\n".join(f'  • "{c}"' for c in cols_to_drop)
        confirm = messagebox.askyesno(
            "Confirmar eliminación",
            f"Se eliminarán {len(cols_to_drop)} columna(s) duplicada(s) "
            f"(se conserva la columna original de cada par):\n\n{names}\n\n"
            "¿Deseas continuar?"
        )
        if not confirm:
            return

        self.dataframe.drop(columns=cols_to_drop, inplace=True)
        self._exact_dup_col_pairs = []
        self.cell_formats    = {}
        self.edited_cells    = {}
        self.selected_cell   = None
        self.selected_column = None

        self._display_data()
        self._refresh_duplicates()
        self._refresh_auto_panel()
        messagebox.showinfo(
            "Eliminadas",
            f"{len(cols_to_drop)} columna(s) duplicada(s) eliminadas correctamente.")

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
            self._refresh_auto_panel()   # actualizar pestaña Automatización
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
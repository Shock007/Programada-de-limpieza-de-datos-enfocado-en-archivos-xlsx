import tkinter as tk
from tkinter import messagebox, filedialog, colorchooser
from tkinter import ttk
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


class DataCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XLSX Data Cleaner")
        self.root.geometry("1100x700")
        
        self.current_file = None
        self.dataframe = None
        self.edited_cells = {}
        self.cell_formats = {}  # Store formatting for each cell
        self.selected_cell = None  # Track selected cell (row, col)
        self.row_id_map = {}  # Map Treeview item IDs to dataframe indices
        
        # Main frame
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Button frame
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Upload button
        self.upload_btn = ttk.Button(
            self.button_frame,
            text="Enter your file",
            command=self.upload_file
        )
        self.upload_btn.pack(side=tk.LEFT, padx=5)
        
        # Save button
        self.save_btn = ttk.Button(
            self.button_frame,
            text="Save Changes",
            command=self.save_file,
            state=tk.DISABLED
        )
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # File label
        self.file_label = ttk.Label(self.button_frame, text="No file loaded")
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # Content frame (table + edit panel)
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)
        
        # Table frame with scrollbars
        self.table_frame = ttk.Frame(self.content_frame)
        self.table_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 10))
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)
        
        # Scrollbars
        vsb = ttk.Scrollbar(self.table_frame, orient=tk.VERTICAL)
        hsb = ttk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL)
        
        # Treeview
        self.tree = ttk.Treeview(
            self.table_frame,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Grid layout for table
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)
        
        # Edit Panel (Right side)
        self.edit_panel = ttk.LabelFrame(self.content_frame, text="Edit", padding=10)
        self.edit_panel.grid(row=0, column=1, sticky=tk.NSEW, padx=5)
        
        # Cell info label
        self.cell_info_label = ttk.Label(self.edit_panel, text="No cell selected", relief=tk.SUNKEN)
        self.cell_info_label.pack(fill=tk.X, pady=(0, 10))
        
        # Value editor
        ttk.Label(self.edit_panel, text="Value:").pack(anchor=tk.W, pady=(5, 0))
        self.value_entry = ttk.Entry(self.edit_panel, width=20)
        self.value_entry.pack(fill=tk.X, pady=(0, 10))
        
        # Font selection
        ttk.Label(self.edit_panel, text="Font:").pack(anchor=tk.W, pady=(5, 0))
        self.font_var = tk.StringVar(value="Calibri")
        self.font_combo = ttk.Combobox(
            self.edit_panel, 
            textvariable=self.font_var, 
            width=17,
            values=["Calibri", "Arial", "Times New Roman", "Courier New"]
        )
        self.font_combo.pack(fill=tk.X, pady=(0, 10))
        
        # Font size
        ttk.Label(self.edit_panel, text="Font Size:").pack(anchor=tk.W, pady=(5, 0))
        self.font_size_var = tk.StringVar(value="11")
        self.font_size_spin = ttk.Spinbox(
            self.edit_panel,
            from_=8,
            to=72,
            textvariable=self.font_size_var,
            width=17
        )
        self.font_size_spin.pack(fill=tk.X, pady=(0, 10))
        
        # Font color
        ttk.Label(self.edit_panel, text="Font Color:").pack(anchor=tk.W, pady=(5, 0))
        self.font_color_frame = ttk.Frame(self.edit_panel)
        self.font_color_frame.pack(fill=tk.X, pady=(0, 10))
        self.font_color_btn = tk.Button(
            self.font_color_frame,
            text="Choose Color",
            command=self.choose_font_color,
            bg="black",
            fg="white",
            width=17
        )
        self.font_color_btn.pack(fill=tk.X)
        self.font_color_var = tk.StringVar(value="000000")
        
        # Cell background color
        ttk.Label(self.edit_panel, text="Background Color:").pack(anchor=tk.W, pady=(5, 0))
        self.bg_color_frame = ttk.Frame(self.edit_panel)
        self.bg_color_frame.pack(fill=tk.X, pady=(0, 10))
        self.bg_color_btn = tk.Button(
            self.bg_color_frame,
            text="Choose Color",
            command=self.choose_bg_color,
            bg="white",
            fg="black",
            width=17
        )
        self.bg_color_btn.pack(fill=tk.X)
        self.bg_color_var = tk.StringVar(value="FFFFFF")
        
        # Alignment
        ttk.Label(self.edit_panel, text="Alignment:").pack(anchor=tk.W, pady=(5, 0))
        self.alignment_var = tk.StringVar(value="left")
        self.alignment_combo = ttk.Combobox(
            self.edit_panel,
            textvariable=self.alignment_var,
            width=17,
            values=["left", "center", "right"]
        )
        self.alignment_combo.pack(fill=tk.X, pady=(0, 10))
        
        # Number format
        ttk.Label(self.edit_panel, text="Number Format:").pack(anchor=tk.W, pady=(5, 0))
        self.number_format_var = tk.StringVar(value="General")
        self.number_format_combo = ttk.Combobox(
            self.edit_panel,
            textvariable=self.number_format_var,
            width=17,
            state="readonly",
            values=["General", "Number", "US Currency", "Colombian Currency", "Date", "Time", "Percentage", "Text"]
        )
        self.number_format_combo.pack(fill=tk.X, pady=(0, 10))
        
        # Apply button
        self.apply_btn = ttk.Button(
            self.edit_panel,
            text="Apply Formatting",
            command=self.apply_formatting,
            state=tk.DISABLED
        )
        self.apply_btn.pack(fill=tk.X, pady=(10, 0))
        
        # Bind double-click for editing
        self.tree.bind("<Double-1>", self.on_cell_double_click)
        self.tree.bind("<Button-1>", self.on_cell_click)
        self.tree.bind("<Return>", self.on_key_release)
        
    def upload_file(self):
        """Open file dialog to select xlsx file"""
        file_path = filedialog.askopenfilename(
            title="Select XLSX file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.current_file = file_path
                self.dataframe = pd.read_excel(file_path)
                self.edited_cells = {}
                self.cell_formats = {}
                self.display_data()
                self.file_label.config(text=f"Loaded: {os.path.basename(file_path)}")
                self.save_btn.config(state=tk.NORMAL)
                messagebox.showinfo("Success", "File loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")
    
    def display_data(self):
        """Display dataframe in treeview"""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.row_id_map = {}  # Reset mapping
        
        if self.dataframe is None:
            return
        
        # Set up columns
        columns = ["#0"] + list(self.dataframe.columns)
        self.tree["columns"] = list(self.dataframe.columns)
        
        # Define column headings
        self.tree.column("#0", width=50, anchor=tk.CENTER)
        self.tree.heading("#0", text="Row")
        
        for col in self.dataframe.columns:
            self.tree.column(col, width=100, anchor=tk.W)
            self.tree.heading(col, text=str(col))
        
        # Insert data and track row indices
        for idx, row in self.dataframe.iterrows():
            values = [str(val) for val in row]
            item_id = self.tree.insert("", tk.END, text=str(idx), values=values)
            self.row_id_map[item_id] = idx  # Map Treeview item ID to dataframe row index
    
    def on_cell_double_click(self, event):
        """Handle double-click to edit cell"""
        region = self.tree.identify_region(event.x, event.y)
        column = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        
        if region != "cell" or not row:
            return
        
        # Get cell position
        col_index = int(column) - 1
        
        # Get actual row index from mapping
        if row not in self.row_id_map:
            return
        row_index = self.row_id_map[row]
        
        if col_index < 0:
            return
        
        # Get current value
        item = self.tree.item(row)
        current_value = item["values"][col_index]
        
        # Get actual row index from mapping
        if row not in self.row_id_map:
            return
        row_index = self.row_id_map[row]
        
        # Create edit window
        self.edit_window(row_index, col_index, current_value)
    
    def on_cell_click(self, event):
        """Handle single-click to select cell for formatting"""
        region = self.tree.identify_region(event.x, event.y)
        column = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        
        if region != "cell" or not row:
            return
        
        # Get cell position
        col_index = int(column) - 1
        
        # Get actual row index from mapping
        if row not in self.row_id_map:
            return
        row_index = self.row_id_map[row]
        
        if col_index < 0:
            return
        
        # Store selected cell
        self.selected_cell = (row_index, col_index)
        
        # Get current value
        item = self.tree.item(row)
        current_value = item["values"][col_index]
        col_name = self.dataframe.columns[col_index]
        
        # Update cell info label
        self.cell_info_label.config(text=f"Cell: {col_name} (Row {row_index})")
        
        # Update value entry
        self.value_entry.delete(0, tk.END)
        self.value_entry.insert(0, str(current_value))
        
        # Load stored formatting if exists
        if (row_index, col_index) in self.cell_formats:
            fmt = self.cell_formats[(row_index, col_index)]
            self.font_var.set(fmt.get("font", "Calibri"))
            self.font_size_var.set(str(fmt.get("size", 11)))
            self.font_color_var.set(fmt.get("font_color", "000000"))
            self.bg_color_var.set(fmt.get("bg_color", "FFFFFF"))
            self.alignment_var.set(fmt.get("alignment", "left"))
            self.number_format_var.set(fmt.get("number_format", "General"))
            self.update_color_buttons()
        else:
            # Reset to defaults
            self.font_var.set("Calibri")
            self.font_size_var.set("11")
            self.font_color_var.set("000000")
            self.bg_color_var.set("FFFFFF")
            self.alignment_var.set("left")
            self.number_format_var.set("General")
            self.update_color_buttons()
        
        # Enable apply button
        self.apply_btn.config(state=tk.NORMAL)
    
    def choose_font_color(self):
        """Open color picker for font color"""
        color = colorchooser.askcolor(title="Choose Font Color")
        if color[1]:
            hex_color = color[1].lstrip("#")
            self.font_color_var.set(hex_color)
            self.update_color_buttons()
    
    def choose_bg_color(self):
        """Open color picker for background color"""
        color = colorchooser.askcolor(title="Choose Background Color")
        if color[1]:
            hex_color = color[1].lstrip("#")
            self.bg_color_var.set(hex_color)
            self.update_color_buttons()
    
    def update_color_buttons(self):
        """Update button colors to reflect selected colors"""
        try:
            self.font_color_btn.config(bg="#" + self.font_color_var.get())
            self.bg_color_btn.config(bg="#" + self.bg_color_var.get())
        except tk.TclError:
            pass
    
    def apply_formatting(self):
        """Apply formatting to selected cell"""
        if self.selected_cell is None:
            messagebox.showwarning("Warning", "Please select a cell first")
            return
        
        row_index, col_index = self.selected_cell
        
        # Store the formatting
        self.cell_formats[(row_index, col_index)] = {
            "font": self.font_var.get(),
            "size": int(self.font_size_var.get()),
            "font_color": self.font_color_var.get(),
            "bg_color": self.bg_color_var.get(),
            "alignment": self.alignment_var.get(),
            "number_format": self.number_format_var.get()
        }
        
        # Update the cell value in dataframe
        col_name = self.dataframe.columns[col_index]
        new_value = self.value_entry.get()
        self.dataframe.at[row_index, col_name] = new_value
        self.edited_cells[(row_index, col_index)] = new_value
        
        # Refresh display
        self.display_data()
        messagebox.showinfo("Success", "Formatting applied!")
    
    def edit_window(self, row_index, col_index, current_value):
        """Create window for editing cell"""
        edit_win = tk.Toplevel(self.root)
        edit_win.title("Edit Cell")
        edit_win.geometry("300x100")
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        ttk.Label(edit_win, text="Enter new value:").pack(pady=5)
        
        entry = ttk.Entry(edit_win)
        entry.insert(0, str(current_value))
        entry.pack(pady=5, padx=10, fill=tk.X)
        entry.focus()
        
        def save_edit():
            new_value = entry.get()
            
            # Store edit
            col_name = self.dataframe.columns[col_index]
            self.dataframe.at[row_index, col_name] = new_value
            self.edited_cells[(row_index, col_index)] = new_value
            
            # Update tree
            self.display_data()
            edit_win.destroy()
        
        btn_frame = ttk.Frame(edit_win)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Save", command=save_edit).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=edit_win.destroy).pack(side=tk.LEFT, padx=5)
        
        entry.bind("<Return>", lambda e: save_edit())
    
    def on_key_release(self, event):
        """Handle return key"""
        pass
    
    def save_file(self):
        """Save changes back to xlsx file with formatting"""
        if self.current_file is None or self.dataframe is None:
            messagebox.showwarning("Warning", "No file loaded")
            return
        
        try:
            # First, save the data with pandas
            self.dataframe.to_excel(self.current_file, index=False)
            
            # Now apply formatting using openpyxl
            wb = load_workbook(self.current_file)
            ws = wb.active
            
            # Apply formatting to each cell that has stored formatting
            for (row_index, col_index), fmt in self.cell_formats.items():
                # openpyxl uses 1-based indexing
                cell = ws.cell(row=row_index + 2, column=col_index + 1)  # +2 because row 1 is header
                
                # Apply font
                font_name = fmt.get("font", "Calibri")
                font_size = fmt.get("size", 11)
                font_color = fmt.get("font_color", "000000")
                cell.font = Font(name=font_name, size=font_size, color=font_color)
                
                # Apply background color
                bg_color = fmt.get("bg_color", "FFFFFF")
                if bg_color and bg_color.upper() != "FFFFFF":
                    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                
                # Apply alignment
                alignment = fmt.get("alignment", "left")
                horizontal_align = alignment if alignment in ["left", "center", "right"] else "left"
                cell.alignment = Alignment(horizontal=horizontal_align, vertical="center", wrap_text=True)
                
                # Apply number format
                number_format = fmt.get("number_format", "General")
                format_map = {
                    "General": "General",
                    "Number": "0.00",
                    "US Currency": "$#,##0.00",
                    "Colombian Currency": "[$$-2058]#,##0.00;[RED]-[$$-2058]#,##0.00",
                    "Date": "mm/dd/yyyy",
                    "Time": "hh:mm:ss",
                    "Percentage": "0.00%",
                    "Text": "@"
                }
                cell.number_format = format_map.get(number_format, "General")
            
            wb.save(self.current_file)
            messagebox.showinfo("Success", "File saved successfully with formatting!")
            self.edited_cells = {}
            self.cell_formats = {}
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DataCleanerApp(root)
    root.mainloop()

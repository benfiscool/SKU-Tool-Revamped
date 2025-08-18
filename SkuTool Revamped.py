import re   #Add ctrl+s, ctrl+z, save before quit, autosave
import os
import pickle
import subprocess
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import datetime
import tkinter.font as tkFont

VALUE_RE = re.compile(r'Name=([^|]+)\|Value=([^|]+?)(?=(?:Type=|$))')
web_service_process = None
def get_data_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

class SimpleAutocompleteEntry(tk.Entry):
    def __init__(self, master, suggestion_list, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.suggestion_list = suggestion_list
        self.popup_window = None
        self.listbox = None
        
        # Create StringVar for tracking changes
        self.var = tk.StringVar()
        self.config(textvariable=self.var)
        
        # Bind events
        self.var.trace_add("write", self.update_autocomplete)
        self.bind("<Down>", self.navigate_down)
        self.bind("<Escape>", self.hide_popup)
        self.bind("<FocusOut>", lambda e: self.after(100, self.check_focus))
        
    def update_autocomplete(self, *args):
        self.hide_popup()  # Hide any existing popup
        
        # Get current text and find matches
        text = self.var.get().strip()
        if not text:
            return
        
        # Find matching suggestions
        matches = [item for item in self.suggestion_list if item.lower().startswith(text.lower())]
        if not matches:
            return
            
        # Create popup window
        self.popup_window = tk.Toplevel(self)
        self.popup_window.overrideredirect(True)
        self.popup_window.attributes("-topmost", True)
        
        # Position popup below the entry
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        self.popup_window.geometry(f"+{x}+{y}")
        
        # Create listbox in popup
        self.listbox = tk.Listbox(self.popup_window, height=min(6, len(matches)))
        self.listbox.pack(fill="both", expand=True)
        
        # Fill listbox with matches
        for match in matches:
            self.listbox.insert(tk.END, match)
            
        # Bind listbox events
        self.listbox.bind("<ButtonRelease-1>", self.select_item)
        self.listbox.bind("<Return>", self.select_item)
        self.listbox.bind("<Escape>", lambda e: self.hide_popup())
    
    def navigate_down(self, event):
        if self.popup_window and self.listbox:
            self.listbox.focus_set()
            self.listbox.selection_set(0)
        return "break"
    
    def select_item(self, event=None):
        if self.listbox and self.popup_window:
            if self.listbox.curselection():
                index = self.listbox.curselection()[0]
                value = self.listbox.get(index)
                self.var.set(value)
                self.icursor(tk.END)
            self.hide_popup()
        return "break"
    
    def hide_popup(self, event=None):
        if self.popup_window:
            self.popup_window.destroy()
            self.popup_window = None
            self.listbox = None
    
    def check_focus(self):
        # Don't hide if focus is in the listbox
        focused = self.focus_get()
        if focused and self.popup_window and (focused == self.listbox or focused == self):
            return
        self.hide_popup()

class OptionsParserApp(tk.Tk):
    def __init__(self):
        ensure_web_service_running()
        super().__init__()
        # Default database folder and config to script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.options = {
            "database_folder": script_dir,
        }
        self._load_options()
        if "database_folder" not in self.options:
            self.options["database_folder"] = script_dir
        self.protocol("WM_DELETE_WINDOW", self._on_quit)
        self.title("Product Variation Configurator - Accessory Partners")
        self.geometry("1000x800")
        self.last_export_label = None
        self.input_df = None
        self.master_df = None
        self.entry_widgets = {}
        self.base_sku = ""
        self.current_prog_path = None
        self.unsaved_changes = False

        self._build_menu()
        self._build_ui()
        self.bind_all("<Button-1>", self._clear_highlight_if_needed, add='+')
        self.bind_all('<Control-s>', self._ctrl_s)
        self.bind_all('<Control-S>', self._ctrl_shift_s)
        self.bind_all('<Control-Shift-S>', self._ctrl_shift_s)
        # Show search on startup
        self.after(100, self.load_from_database)

        self.option_names = []  # Always define this attribute

    def _mark_unsaved(self, event=None):
        # Prevent marking as unsaved if Ctrl+S or Ctrl+Shift+S is pressed
        if event is not None and hasattr(event, "keysym"):
            if event.keysym.lower() == "s" and (event.state & 0x4):
                return
        self.unsaved_changes = True
        self._update_title()

    def _update_base_sku_from_entry(self, event=None):
        self.base_sku = self.base_sku_entry.get().strip()
        self._mark_unsaved()

    def _update_title(self):
        # Use the current database file if available, otherwise show Untitled
        if self.current_prog_path:
            fname = os.path.basename(self.current_prog_path)
        else:
            # Try to use the latest database file
            db_path = self._get_latest_database_path()
            fname = os.path.basename(db_path) if os.path.exists(db_path) else "Untitled"
        if self.unsaved_changes:
            fname += " *"
        self.title(f"Product Variation Configurator - Accessory Partners - {fname}")

    def _ctrl_s(self, event=None):
        """Save to temp file only (no upload)."""
        self.save_to_database(temp=True)
        self.unsaved_changes = False
        self._update_title()
        messagebox.showinfo("Saved", "Changes saved to temp file. Make sure to upload to cloud using ctrl + shift + s.")

    def _ctrl_shift_s(self, event=None):
        """Save to temp file and upload to cloud with timestamped name."""
        self.save_to_database(temp=True)
        push_database_to_cloud()
        self.unsaved_changes = False
        self._update_title()

    def _build_menu(self):
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Import CSV from BigCommerce...", command=self.load_csv)
        filemenu.add_command(label="Export Current SKU Breakout to Excel...", command=self.export_excel)
        filemenu.add_command(label="Rename Associated SKU...", command=self.rename_associated_sku)
        filemenu.add_separator()

        filemenu.add_command(label="Save Database Locally", command=self.save_to_database)
        filemenu.add_command(label="Save and Upload to Cloud", command=lambda:[self.save_to_database,self._ctrl_shift_s])
        filemenu.add_command(label="Revert to Backup Database", command=self.revert_to_backup)
        filemenu.add_separator()

        filemenu.add_command(label="Delete Base SKU from Database", command=self.delete_base_sku)
        filemenu.add_command(label="Export Pricing Database...", command=self.export_cost_db)
        filemenu.add_command(label="Import Pricing Database...", command=self.import_cost_db)
        filemenu.add_separator()

        filemenu.add_command(label="Exit", command=self._on_quit)
        menubar.add_cascade(label="File", menu=filemenu)

        # --- Options menu ---
        optionsmenu = tk.Menu(menubar, tearoff=0)

        # Manual Controls submenu
        manualmenu = tk.Menu(optionsmenu, tearoff=0)
        optionsmenu.add_command(label="Config...", command=self.open_options_window)
        optionsmenu.add_separator()
        manualmenu.add_command(label="Parse Table", command=lambda: [self.parse_options(), self._mark_unsaved()])
        manualmenu.add_command(label="Generate New SKUs", command=lambda: [self.generate_new_skus(), self._mark_unsaved()])
        manualmenu.add_command(label="Extract Base SKU", command=lambda: [self.extract_base_sku(), self._mark_unsaved()])
        optionsmenu.add_cascade(label="Manual Controls", menu=manualmenu)

        menubar.add_cascade(label="Options", menu=optionsmenu)

        self.config(menu=menubar)

    def _build_ui(self):
        ctrl = ttk.Frame(self)
        ctrl.pack(fill='x', padx=10, pady=5)
        spot_btn = ttk.Button(ctrl, text="Spot Check", command=self.open_spotcheck_window)
        spot_btn.pack(side='right', padx=5)

        # --- New row for Extract/Base SKU ---
        keyword_row = ttk.Frame(self)
        keyword_row.pack(fill='x', padx=10, pady=(0, 5))

        # Base SKU input
        ttk.Label(keyword_row, text="Base SKU:").pack(side='left')
        self.base_sku_entry = ttk.Entry(keyword_row, width=18)
        self.base_sku_entry.pack(side='left', padx=(0, 10))
        self.base_sku_entry.bind(
            "<Return>", #CHANGED
            lambda e: None if (e.keysym.lower() == "s" and (e.state & 0x4)) or self._is_modifier(e)
            else [self._update_base_sku_from_entry, self.generate_new_skus(), self._mark_unsaved()]
        )

        # Prefix input
        ttk.Label(keyword_row, text="Prefix:").pack(side='left')
        self.prefix_entry = ttk.Entry(keyword_row, width=10)
        self.prefix_entry.pack(side='left', padx=(0, 10))
        self.prefix_entry.bind(
            "<Return>", #CHANGED
            lambda e: None if (e.keysym.lower() == "s" and (e.state & 0x4)) or self._is_modifier(e)
            else [self.generate_new_skus(), self._mark_unsaved()]
        )

        self.last_export_label = ttk.Label(keyword_row, text="", foreground="gray")
        self.last_export_label.pack(side='left', padx=(10, 0))

        cost_import_btn = ttk.Button(ctrl, text="Import Pricing Data", command=self.open_cost_import_window)
        cost_import_btn.pack(side='right', padx=5)

        cost_db_btn = ttk.Button(ctrl, text="Open Pricing Database", command=self.open_cost_db_explorer)
        cost_db_btn.pack(side='right', padx=5)

        search_btn = ttk.Button(ctrl, text="🔍SKU Breakouts", command=self.load_from_database)
        search_btn.pack(side='right', padx=5)

        ttk.Label(ctrl, text="Base Price: $").pack(side='left')
        self.base_price_entry = ttk.Entry(ctrl, width=8)
        self.base_price_entry.pack(side='left', padx=(0,10))
        
        self.last_export_label.pack(side='left', padx=(10, 0))
        self.base_price_entry.bind(
            "<KeyRelease>",
            lambda e: None if (e.keysym.lower() == "s" and (e.state & 0x4)) else [self.generate_new_skus(), self._regenerate_left_preview(), self._mark_unsaved()]
        )
        self.base_price_entry.bind("<Return>", self._mark_unsaved)

        ttk.Label(ctrl, text="Base Weight: lb").pack(side='left')
        self.base_weight_entry = ttk.Entry(ctrl, width=8)
        self.base_weight_entry.pack(side='left', padx=(0,10))
        self.base_weight_entry.bind(
            "<KeyRelease>",
            lambda e: None if (e.keysym.lower() == "s" and (e.state & 0x4)) or self._is_modifier(e)
            else [self.generate_new_skus(), self._regenerate_left_preview(), self._mark_unsaved()]
        )
        self.base_weight_entry.bind("<Return>", self._mark_unsaved)

        # Top preview (breakout columns)
        self.in_tree = self._make_treeview(self, [], stretch=True, label="SKU Combinations Preview")
        self.in_tree.master.pack(fill='both', expand=True, padx=10, pady=(5,0))

        ttk.Separator(self, orient='horizontal').pack(fill='x', padx=10, pady=5)

        # --- Bottom preview (out_tree, breakdown_tree side by side) ---
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill='both', expand=True, padx=10, pady=(0,5))

        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)
        bottom_frame.rowconfigure(0, weight=1)

        # Bottom left: master name/value table
        self.out_tree = self._make_treeview(bottom_frame, ['Name','Value',"Add'l Cost","Add'l Weight"], stretch=True, label="Option Master Table")
        self.out_tree.master.grid(row=0, column=0, sticky='nsew')
        self.out_tree.bind("<Double-1>", self._on_out_tree_double_click)
        self.out_tree.bind("<<TreeviewSelect>>", self._on_out_tree_select)

        # Bottom right: breakdown panel
        self.breakdown_tree = self._make_treeview(bottom_frame, ['Name', 'Value', 'Cost ($)', 'Weight (lb)'], stretch=True, label="Cost/Weight Breakdown")
        self.breakdown_tree.master.grid(row=0, column=1, sticky='nsew')

        # Bind resize event
        bottom_frame.bind("<Configure>", self._resize_bottom_panels)

        # Bind selection in in_tree to populate breakdown_tree
        self.in_tree.bind("<<TreeviewSelect>>", self._on_in_tree_select)

        ttk.Separator(self, orient='horizontal').pack(fill='x', padx=10, pady=5)

        cfg = ttk.LabelFrame(self, text="Configure")
        cfg.pack(fill='x', padx=10, pady=5)
        row = ttk.Frame(cfg); row.pack(fill='x', pady=5)
        ttk.Label(row, text="Name:").grid(row=0, column=0)
        ttk.Button(row, text="◀", width=3, command=lambda: [self._prev_name(), self._mark_unsaved()]).grid(row=0, column=1)
        self.name_combo = ttk.Combobox(row, state='readonly', width=30)
        self.name_combo.grid(row=0, column=2, padx=5)
        self.name_combo.bind("<<ComboboxSelected>>", lambda e: [self.populate_value_grid(), self._mark_unsaved()])
        ttk.Button(row, text="▶", width=3, command=lambda: [self._next_name(), self._mark_unsaved()]).grid(row=0, column=3)

        # Header row for value/cost/weight, aligned with grid_container columns
        self.header_frame = ttk.Frame(cfg)
        self.header_frame.pack(fill='x', padx=5, pady=(0,0))
        ttk.Label(self.header_frame, text="Value", anchor='w', width=40).grid(row=0, column=0, padx=0, sticky='w')
        ttk.Label(self.header_frame, text="Add'l Cost ($)", anchor='w', width=15).grid(row=0, column=1, padx=5, sticky='w')
        ttk.Label(self.header_frame, text="Add'l Weight (lb)", anchor='w', width=15).grid(row=0, column=2, padx=5, sticky='w')
        ttk.Label(self.header_frame, text="Associated SKUs", anchor='w', width=25).grid(row=0, column=3, padx=5, sticky='w')

        self.grid_canvas = tk.Canvas(cfg, height=180)  # Set a reasonable height
        self.grid_canvas.pack(fill='both', expand=True, padx=5, pady=(0,5), side='left')
        self.grid_scrollbar = ttk.Scrollbar(cfg, orient='vertical', command=self.grid_canvas.yview)
        self.grid_scrollbar.pack(side='right', fill='y')
        self.grid_canvas.configure(yscrollcommand=self.grid_scrollbar.set)
        self.grid_container = ttk.Frame(self.grid_canvas)
        self.grid_container_id = self.grid_canvas.create_window((0, 0), window=self.grid_container, anchor='nw')

        def _on_grid_configure(event):
            self.grid_canvas.configure(scrollregion=self.grid_canvas.bbox("all"))
        self.grid_container.bind("<Configure>", _on_grid_configure)

        def _on_canvas_configure(event):
            self.grid_canvas.itemconfig(self.grid_container_id, width=event.width)
        self.grid_canvas.bind("<Configure>", _on_canvas_configure)

    def _make_treeview(self, parent, columns, stretch=False, label="Preview"):
        frame = ttk.Labelframe(parent, text=label)
        tree = ttk.Treeview(frame, columns=columns, show='headings')
        vsb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor='w', stretch=stretch)
        # Use grid for layout
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv")])
        if not path:
            return
        
        try:
            # First, read and parse the CSV data normally
            self.input_df = pd.read_csv(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV:\n{e}")
            return
    
        # Store the first row (base product) data before filtering
        base_row = None
        if not self.input_df.empty:
            base_row = self.input_df.iloc[0]
    
        # Filter out the first row and keep only variant rows
        self.input_df.dropna(subset=['Options'], how='all', inplace=True)
        self.input_df.reset_index(drop=True, inplace=True)
        if self.master_df is not None:
            self.master_df.reset_index(drop=True, inplace=True)
    
        self.option_names = sorted({
            name for blob in self.input_df['Options']
            for name, _ in VALUE_RE.findall(str(blob))
        })
    
        # Now determine website suffix after parsing
        try:
            # Read the entire CSV to check for website indicators
            with open(path, 'r', encoding='utf-8') as f:
                csv_content = f.read().lower()
            
            # Determine website and suffix
            suffix = ""
            is_suspension_superstore = 'suspensionsuperstore' in csv_content
            
            if is_suspension_superstore:
                suffix = ""
            else:
                suffix = " (MXT)"
        except Exception as e:
            messagebox.showerror("Error", f"Failed to determine website:\n{e}")
            return
    
        # --- Set base SKU directly from the base row (Column E) with suffix ---
        if base_row is not None and 'SKU' in base_row:
            try:
                base_sku_from_csv = str(base_row['SKU']).strip()
                if base_sku_from_csv and base_sku_from_csv.lower() != 'nan':
                    proposed_base_sku = base_sku_from_csv + suffix
                else:
                    proposed_base_sku = suffix  # Just the suffix if no SKU found
            except Exception:
                proposed_base_sku = suffix  # Just the suffix if error accessing SKU
        else:
            proposed_base_sku = suffix  # Just the suffix if no base row found
    
        # Check if base SKU already exists in database
        db_path = self._get_latest_database_path()
        if os.path.exists(db_path):
            try:
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
                if proposed_base_sku in db:
                    result = messagebox.askyesnocancel(
                        "Base SKU Already Exists",
                        f"The base SKU '{proposed_base_sku}' already exists in the database.\n\n"
                        "Choose an option:\n"
                        "• Yes: Overwrite the existing data\n"
                        "• No: Load the existing data instead\n"
                        "• Cancel: Stop import"
                    )
                    if result is None:  # Cancel
                        return
                    elif result is False:  # No - load existing data
                        # Load the existing base SKU data
                        state = db[proposed_base_sku]
                        self.input_df = pd.DataFrame(state['input_df'])
                        self.master_df = pd.DataFrame(state['master_df'])
                        in_tree_df = pd.DataFrame(state['in_tree_df'])
                        
                        # Restore option names
                        self.option_names = sorted({
                            name for blob in self.input_df['Options']
                            for name, _ in VALUE_RE.findall(str(blob))
                        })
                        
                        # Restore UI fields
                        self.base_price_entry.delete(0, 'end')
                        self.base_price_entry.insert(0, state.get('base_price', ''))
                        self.base_weight_entry.delete(0, 'end')
                        self.base_weight_entry.insert(0, state.get('base_weight', ''))
                        self.base_sku_entry.delete(0, 'end')
                        self.base_sku_entry.insert(0, state.get('base_sku', ''))
                        self.prefix_entry.delete(0, 'end')
                        self.prefix_entry.insert(0, state.get('prefix', ''))
                        self.base_sku = proposed_base_sku
                        self._show_last_export_time(self.base_sku)
                        
                        # Fill trees and UI
                        self._fill_tree(self.in_tree, in_tree_df)
                        self._fill_tree(self.out_tree, self.master_df)
                        
                        # Reconfigure the "Configure" panel
                        names = sorted(self.master_df['Name'].unique())
                        self.name_combo['values'] = names
                        if names:
                            self.name_combo.set(names[0])
                            self.populate_value_grid()
                        
                        self._regenerate_left_preview()
                        self.current_prog_path = db_path
                        self.unsaved_changes = False
                        self._update_title()
                        self._resize_bottom_panels()
                        return
                    # If result is True, continue with overwrite (normal flow)
            except Exception as e:
                print(f"Error checking database: {e}")
        
        # Continue with normal import flow
        self.base_sku = proposed_base_sku
        self.base_sku_entry.delete(0, 'end')
        self.base_sku_entry.insert(0, self.base_sku)
        self._show_last_export_time(self.base_sku)
    
        # --- Set base price and weight from first row (base product) ---
        if base_row is not None:
            # Try to get price from 'Price' column
            try:
                base_price = base_row.get('Price', 0)
                if pd.notna(base_price):
                    self.base_price_entry.delete(0, 'end')
                    self.base_price_entry.insert(0, str(base_price))
                else:
                    self.base_price_entry.delete(0, 'end')
                    self.base_price_entry.insert(0, "0")
            except Exception:
                self.base_price_entry.delete(0, 'end')
                self.base_price_entry.insert(0, "0")
            
            # Try to get weight from 'Weight' column
            try:
                base_weight = base_row.get('Weight', 0)
                if pd.notna(base_weight):
                    self.base_weight_entry.delete(0, 'end')
                    self.base_weight_entry.insert(0, str(base_weight))
                else:
                    self.base_weight_entry.delete(0, 'end')
                    self.base_weight_entry.insert(0, "0")
            except Exception:
                self.base_weight_entry.delete(0, 'end')
                self.base_weight_entry.insert(0, "0")
        else:
            # Fallback to 0 if no base row found
            self.base_price_entry.delete(0, 'end')
            self.base_price_entry.insert(0, "0")
            self.base_weight_entry.delete(0, 'end')
            self.base_weight_entry.insert(0, "0")
    
        # Auto-populate prefix based on website
        self.prefix_entry.delete(0, 'end')
        if not is_suspension_superstore:
            # If it's NOT SuspensionSuperstore (i.e., it's MaxTrac), auto-populate with "M"
            self.prefix_entry.insert(0, "M")
    
        # Clear current file path and mark as unsaved
        self.current_prog_path = None
        self.unsaved_changes = True
        self._update_title()
        
        # Parse the options and generate new SKUs
        self.parse_options()
        self.generate_new_skus()

    def parse_options(self):
        if self.input_df is None:
            return messagebox.showwarning("Warning", "Load a CSV first.")
        rows, seen = [], set()
        for blob in self.input_df['Options']:
            for name, val in VALUE_RE.findall(str(blob)):
                key = (name.strip(), val.strip())
                if key not in seen:
                    seen.add(key)
                    rows.append({
                        'Name': key[0],
                        'Value': key[1],
                        "Add'l Cost": '0',
                        "Add'l Weight": '0',
                        "Associated SKUs": ''
                    })
        # Sort by Name, then Value
        self.master_df = pd.DataFrame(rows)[['Name','Value',"Add'l Cost","Add'l Weight","Associated SKUs"]]
        self._fill_tree(self.out_tree, self.master_df)
    
        names = sorted(self.master_df['Name'].unique())
        self.name_combo['values'] = names
        if names:
            self.name_combo.set(names[0])
            self.populate_value_grid()
    
        # Removed: self.extract_base_sku() - no longer needed
        self.generate_new_skus()
        self._regenerate_left_preview()
        self._mark_unsaved()
        self._resize_bottom_panels()
    
    def extract_base_sku(self):
        if self.input_df is not None and 'SKU' in self.input_df.columns and not self.input_df.empty:
            sku = str(self.input_df['SKU'].iloc[0])
            base = sku.split('-')[0]
            self.base_sku = base
            self.base_sku_entry.delete(0, 'end')
            self.base_sku_entry.insert(0, base)
            self._mark_unsaved()

    def populate_value_grid(self):
        # Clear previous widgets
        for w in self.grid_container.winfo_children():
            w.destroy()
        self.entry_widgets.clear()
    
        sel = self.name_combo.get()
        sub = self.master_df[self.master_df['Name'] == sel].sort_values('Value')
        for _, row in sub.iterrows():
            orig_idx = row.name  # This is the index in master_df
            frame = ttk.Frame(self.grid_container)
            frame.pack(fill='x', pady=2)
            ttk.Label(frame, text=row['Value'], width=40, anchor='w').grid(row=0, column=0)
            ec = ttk.Entry(frame, width=15)
            ec.grid(row=0, column=1, padx=5)
            ec.insert(0, row["Add'l Cost"])
            ew = ttk.Entry(frame, width=15)
            ew.grid(row=0, column=2, padx=5)
            ew.insert(0, row["Add'l Weight"])
            esku = ttk.Entry(frame, width=25)  # Associated SKUs entry
            esku.grid(row=0, column=3, padx=5)
            esku.insert(0, row.get("Associated SKUs", ""))
    
            edit_btn = ttk.Button(frame, text="Edit SKUs", width=10, command=lambda idx=orig_idx: self.open_sku_cost_editor(idx))
            edit_btn.grid(row=0, column=4, padx=5)
    
            ec.bind("<FocusIn>", lambda e: e.widget.select_range(0, 'end'))
            ew.bind("<FocusIn>", lambda e: e.widget.select_range(0, 'end'))
            esku.bind("<Button-1>", lambda e: (edit_btn.focus_set(), "break")[1])
            esku.bind("<FocusIn>", lambda e: edit_btn.focus_set())
    
            self.entry_widgets[orig_idx] = {'Cost': ec, 'Weight': ew, 'SKUs': esku}
            ec.bind(
                "<KeyRelease>",
                lambda e, i=orig_idx: None if (e.keysym.lower() == "s" and (e.state & 0x4)) or self._is_modifier(e)
                else self._on_change(i, field_changed='Cost')
            )
            ew.bind(
                "<KeyRelease>",
                lambda e, i=orig_idx: None if (e.keysym.lower() == "s" and (e.state & 0x4)) or self._is_modifier(e)
                else self._on_change(i, field_changed='Weight')
            )
            esku.bind(
                "<KeyRelease>",
                lambda e, i=orig_idx: None if (e.keysym.lower() == "s" and (e.state & 0x4)) or self._is_modifier(e)
                else self._on_change(i, field_changed='SKUs')
            )
            ec.bind("<Return>", lambda e, i=orig_idx: self._on_change(i))
            ew.bind("<Return>", lambda e, i=orig_idx: self._on_change(i))
            esku.bind("<Return>", lambda e, i=orig_idx: self._on_change(i))
    
        # Update header to include Associated SKUs
        for widget in self.header_frame.winfo_children():
            widget.destroy()
        ttk.Label(self.header_frame, text="Value", anchor='w', width=40).grid(row=0, column=0, padx=0, sticky='w')
        ttk.Label(self.header_frame, text="Add'l Cost ($)", anchor='w', width=15).grid(row=0, column=1, padx=5, sticky='w')
        ttk.Label(self.header_frame, text="Add'l Weight (lb)", anchor='w', width=15).grid(row=0, column=2, padx=5, sticky='w')
        ttk.Label(self.header_frame, text="Associated SKUs", anchor='w', width=25).grid(row=0, column=3, padx=5, sticky='w')
    
        # Ensure the scrollregion is updated
        self.grid_canvas.update_idletasks()
        self.grid_canvas.configure(scrollregion=self.grid_canvas.bbox("all"))

    def _on_change(self, idx, field_changed=None):
        w = self.entry_widgets[idx]
        if field_changed == 'SKUs':
            # Only recalculate Add'l Cost if SKUs changed
            sku_costs = parse_associated_skus(w['SKUs'].get().strip())
            total_cost = sum(float(c) if c else 0 for _, c in sku_costs)
            w['Cost'].delete(0, 'end')
            w['Cost'].insert(0, str(total_cost))
        # Always update master_df with current values
        self.master_df.at[idx,"Add'l Cost"]   = w['Cost'].get().strip() or '0'
        self.master_df.at[idx,"Add'l Weight"] = w['Weight'].get().strip() or '0'
        self.master_df.at[idx,"Associated SKUs"] = w['SKUs'].get().strip() or ''
        self._regenerate_left_preview()
        self._fill_tree(self.out_tree, self.master_df)
        self.generate_new_skus()
        self._mark_unsaved()
        self._resize_bottom_panels()

    def _regenerate_left_preview(self):
        if self.input_df is None or self.master_df is None or not self.option_names:
            return
        try:
            bp = float(self.base_price_entry.get())
            bw = float(self.base_weight_entry.get())
        except ValueError:
            return
    
        # Prepare exploded options DataFrame
        
        sku_rows = []
        option_rows = []
        for idx, row in enumerate(self.input_df.itertuples(), start=1):
            sku = row.SKU
            options = list(VALUE_RE.findall(str(row.Options)))
            sku_rows.append({
                '#': idx,
                'SKU': sku,
                'Options': options,
                'New SKU': '',  # Will fill below
            })
            for n, v in options:
                option_rows.append({'#': idx, 'SKU': sku, 'Name': n, 'Value': v})
    
        sku_df = pd.DataFrame(sku_rows)
        option_df = pd.DataFrame(option_rows)
    
        # Merge with master_df to get costs/weights for each option
        if not option_df.empty:
            merged = option_df.merge(self.master_df, on=['Name', 'Value'], how='left')
            merged["Add'l Cost"] = pd.to_numeric(merged["Add'l Cost"], errors='coerce').fillna(0)
            merged["Add'l Weight"] = pd.to_numeric(merged["Add'l Weight"], errors='coerce').fillna(0)
            cost_weight = merged.groupby('#').agg({'Add\'l Cost': 'sum', 'Add\'l Weight': 'sum'}).reset_index()
        else:
            cost_weight = pd.DataFrame({'#': [], "Add'l Cost": [], "Add'l Weight": []})
    
        # Merge cost/weight sums back to sku_df
        sku_df = sku_df.merge(cost_weight, on='#', how='left').fillna({'Add\'l Cost': 0, 'Add\'l Weight': 0})
    
        # Calculate final price/weight
        sku_df['Price'] = (bp + sku_df["Add'l Cost"]).round(2)
        sku_df['Weight'] = (bw + sku_df["Add'l Weight"]).round(2)
    
        # Prepare option columns
        all_option_names = self.option_names
        for name in all_option_names:
            sku_df[name] = ''
    
        # Fill option columns
        for idx, row in sku_df.iterrows():
            for n, v in row['Options']:
                sku_df.at[idx, n] = v
    
        # Preserve existing New SKU if present
        current_df = self._get_tree_df(self.in_tree)
        existing_new_skus = {}
        if 'SKU' in current_df.columns and 'New SKU' in current_df.columns:
            for idx, row in current_df.iterrows():
                existing_new_skus[(row['SKU'], idx)] = row['New SKU']
    
        for idx, row in sku_df.iterrows():
            prev_new_sku = existing_new_skus.get((row['SKU'], idx), '')
            # If no existing New SKU, generate one using the same logic as generate_new_skus
            if not prev_new_sku:
                # Get the base SKU and clean it
                base_sku = self.base_sku_entry.get().strip() or self.base_sku
                if base_sku:
                    # Apply the same cleaning logic as generate_new_skus
                    base_sku_clean = re.sub(r'\([^)]*\)', '', base_sku)
                    base_sku_clean = base_sku_clean.replace(' ', '')
                    base_sku_clean = base_sku_clean.strip()
                    
                    # Get the prefix and handle zero removal
                    prefix = self.prefix_entry.get().strip()
                    zero_removal = 0
                    if prefix.startswith('-') and len(prefix) > 1 and prefix[1:].isdigit():
                        zero_removal = int(prefix[1:])
                        prefix = ""
                    
                    # Generate the new SKU with proper formatting
                    padding = max(1, 4 - zero_removal)
                    number_str = str(idx + 1).zfill(padding)
                    
                    if prefix:
                        prev_new_sku = f"{base_sku_clean}-{prefix}{number_str}"
                    else:
                        prev_new_sku = f"{base_sku_clean}-{number_str}"
            
            sku_df.at[idx, 'New SKU'] = prev_new_sku
    
        # Reorder columns for display
        display_cols = ['#', 'SKU', 'New SKU', 'Price', 'Weight'] + all_option_names
        self._fill_tree(self.in_tree, sku_df[display_cols])
        self._resize_bottom_panels()
    
        # --- Auto-resize columns based on content ---
        for col in ("SKU", "Price"):
            self.in_tree.column(col, width=tkFont.Font().measure(col) + 40, stretch=True)
            # Optionally, auto-size to content:
            max_width = max([self.in_tree.column(col, width=None), max((tkFont.Font().measure(str(self.in_tree.set(iid, col))) for iid in self.in_tree.get_children()), default=0) + 40])
            self.in_tree.column(col, width=max_width)

    def save_progress(self, path=None, show_dialog=True):
        if self.input_df is None:
            return
        base_sku = self.base_sku_entry.get().strip() or self.base_sku
        if not base_sku:
            messagebox.showwarning("Missing Base SKU", "Please enter a Base SKU before saving progress.")
            return
        if not path or show_dialog:
            path = filedialog.asksaveasfilename(
                defaultextension='.prog',
                initialfile=f"{base_sku}.prog",
                filetypes=[('Prog','*.prog')]
            )
            if not path:
                return
        in_tree_df = self._get_tree_df(self.in_tree)
        state = {
            'input_df': self.input_df,
            'master_df': self.master_df,
            'in_tree_df': in_tree_df,
            'base_price': self.base_price_entry.get(),
            'base_weight': self.base_weight_entry.get(),
            'base_sku': self.base_sku_entry.get().strip()
        }
        with open(path, 'wb') as f:
            pickle.dump(state, f)
        self.current_prog_path = path
        self.unsaved_changes = False
        self._update_title()
        if show_dialog:
            messagebox.showinfo("Saved Progress", f"Saved to {path}")

    def load_progress(self):
        path = filedialog.askopenfilename(filetypes=[('Prog','*.prog')])
        if not path:
            return
        try:
            with open(path, 'rb') as f:
                state = pickle.load(f)
        except Exception:
            messagebox.showerror("Error", "Invalid prog file.")
            return

        self.input_df  = pd.DataFrame(state['input_df'])
        self.master_df = pd.DataFrame(state['master_df'])
        in_tree_df = pd.DataFrame(state['in_tree_df'])

        self.base_price_entry.delete(0, 'end')
        self.base_price_entry.insert(0, state.get('base_price', ''))
        self.base_weight_entry.delete(0, 'end')
        self.base_weight_entry.insert(0, state.get('base_weight', ''))

        self.base_sku_entry.delete(0, 'end')
        self.base_sku_entry.insert(0, state.get('base_sku', ''))
        self.base_sku = self.base_sku_entry.get().strip()
        self._show_last_export_time(self.base_sku)

        self.option_names = sorted({
            name for blob in self.input_df['Options']
            for name, _ in VALUE_RE.findall(str(blob))
        })

        # --- Ensure in_tree_df is a DataFrame ---
        in_tree_df = state.get('in_tree_df')
        if in_tree_df is not None:
            if not isinstance(in_tree_df, pd.DataFrame):
                in_tree_df = pd.DataFrame(in_tree_df)
            self._fill_tree(self.in_tree, in_tree_df)
            self._resize_bottom_panels()
        else:
            # fallback: rebuild as before
            cols = ['#','SKU','New SKU','Price','Weight'] + self.option_names
            rows = []
            for idx, row in enumerate(self.input_df.itertuples(), start=1):
                entry = {'#': idx, 'SKU': row.SKU, 'New SKU': '', 'Price': '', 'Weight': ''}
                for n in self.option_names:
                    entry[n] = ''
                for n, v in VALUE_RE.findall(str(row.Options)):
                    entry[n] = v
                rows.append(entry)
            top_df = pd.DataFrame(rows, columns=cols)
            self._fill_tree(self.in_tree, top_df)
            self._resize_bottom_panels()
            self._regenerate_left_preview()  # Only call here if rebuilding

        # refill bottom preview (out_tree)
        self._fill_tree(self.out_tree, self.master_df)
        self._resize_bottom_panels()
        # reconfigure the “Configure” panel
        names = sorted(self.master_df['Name'].unique())
        self.name_combo['values'] = names
        if names:
            self.name_combo.set(names[0])
            self.populate_value_grid()

        # finally regenerate prices & weights on the left preview
        self._regenerate_left_preview()
        self.current_prog_path = path
        self.unsaved_changes = False
        self._update_title()

    def export_excel(self):
        if not self.base_price_entry.get().strip() or not self.base_weight_entry.get().strip():
            return messagebox.showwarning(
                "Missing Base",
                "Enter base price & weight before export."
            )
        df = self._get_tree_df(self.in_tree)
        base_sku = self.base_sku_entry.get().strip() or self.base_sku or ''
        if not base_sku:
            base_sku = 'output'
        path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            initialfile=f"{base_sku}.xlsx",
            filetypes=[('Excel','*.xlsx')]
        )
        if not path:
            return
        df.to_excel(path, index=False)
        if messagebox.askyesno("Exported", f"Saved to {path}\nOpen in Excel?"):
            try:
                subprocess.Popen(['start','excel',path], shell=True)
            except Exception:
                messagebox.showerror("Error","Can't open Excel.")
        self._update_last_export_time(base_sku)
        self._show_last_export_time(base_sku)

    def save_to_database(self, db_path=None, temp=False):
        folder = self._get_database_folder()
        if temp:
            db_path = os.path.join(folder, "sku_database_temp.json")
        else:
            dt = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            db_path = os.path.join(folder, f"sku_database_{dt}.json")
        base_sku = self.base_sku_entry.get().strip() or self.base_sku
        
        if not base_sku:
            # Don't save if no base SKU is set
            return
        if os.path.exists(self._get_latest_database_path()):
            with open(self._get_latest_database_path(), "r", encoding="utf-8") as f:
                db = json.load(f)
        else:
            db = {}
        in_tree_df = self._get_tree_df(self.in_tree)
        state = {
            'input_df': self.input_df.to_dict(orient='records') if self.input_df is not None else [],
            'master_df': self.master_df.to_dict(orient='records') if self.master_df is not None else [],
            'in_tree_df': in_tree_df.to_dict(orient='records') if in_tree_df is not None else [],
            'base_price': self.base_price_entry.get(),
            'base_weight': self.base_weight_entry.get(),
            'base_sku': base_sku,
            'prefix': self.prefix_entry.get().strip()  # Save the prefix
        }
        # Preserve last_export if present
        if base_sku in db and 'last_export' in db[base_sku]:
            state['last_export'] = db[base_sku]['last_export']
        db[base_sku] = state
        with open(db_path, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=2)
        if not temp:
            temp_path = os.path.join(folder, "sku_database_temp.json")
            if os.path.exists(temp_path):
                os.remove(temp_path)
            self.current_prog_path = db_path
            self.unsaved_changes = False
            self._update_title()
            messagebox.showinfo("Saved", f"Saved {base_sku} to {db_path}")

    def load_from_database(self):
        self.save_to_database(temp=True)
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            with open(db_path, "w", encoding="utf-8") as f:
                json.dump({}, f, indent=2)
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except (UnicodeDecodeError, json.JSONDecodeError):
            # Try to repair the file
            db = try_repair_json_file(db_path)
            if db is None:
                messagebox.showerror(
                    "Database Corrupted",
                    f"The database file could not be repaired. Please restore from backup."
                )
                db = {}
        base_skus = list(db.keys())
        if not base_skus:
            messagebox.showinfo("No Data", "No SKUs in database.")
            return

        # Create a search dialog with a listbox
        import tkinter.simpledialog

        class SkuSearchDialog(tk.Toplevel):
            def __init__(self, parent, skus):
                super().__init__(parent)
                self.title("Select Base SKU")
                self.geometry("400x400")
                self.selected = None
                self.all_skus = sorted(skus)  # Store all SKUs
                
                # Make dialog modal
                self.transient(parent)
                self.grab_set()
                
                tk.Label(self, text="Search:").pack(anchor='w', padx=10, pady=(10,0))
                self.var = tk.StringVar()
                self.var.trace_add('write', self.update_list)
                entry = tk.Entry(self, textvariable=self.var)
                entry.pack(fill='x', padx=10, pady=(0,10))
                
                # Create listbox with scrollbar
                list_frame = tk.Frame(self)
                list_frame.pack(fill='both', expand=True, padx=10, pady=(0,10))
                
                scrollbar = tk.Scrollbar(list_frame)
                scrollbar.pack(side='right', fill='y')
                
                self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
                self.listbox.pack(side='left', fill='both', expand=True)
                scrollbar.config(command=self.listbox.yview)
                
                # Initial population
                self.populate_listbox(self.all_skus)
                    
                self.listbox.bind('<Double-1>', self.select)
                self.listbox.bind('<Return>', self.select)
                
                btn_frame = tk.Frame(self)
                btn_frame.pack(fill='x', padx=10, pady=(0,10))
                tk.Button(btn_frame, text="Load", command=self.select).pack(side='right', padx=(5,0))
                tk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side='right')
                
                entry.focus_set()
                
                # Center the dialog
                self.center_on_parent(parent)
                
            def center_on_parent(self, parent):
                self.update_idletasks()
                x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
                y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
                self.geometry(f"+{x}+{y}")
                
            def populate_listbox(self, skus):
                self.listbox.delete(0, 'end')
                for sku in skus:
                    self.listbox.insert('end', sku)
                
            def update_list(self, *args):
                search_text = self.var.get().lower()
                if not search_text:
                    # Show all SKUs if search is empty
                    self.populate_listbox(self.all_skus)
                else:
                    # Filter SKUs based on search text
                    filtered_skus = [sku for sku in self.all_skus if search_text in sku.lower()]
                    self.populate_listbox(filtered_skus)
                        
            def select(self, event=None):
                selection = self.listbox.curselection()
                if selection:
                    self.selected = self.listbox.get(selection[0])
                    self.destroy()

        dlg = SkuSearchDialog(self, base_skus)
        self.wait_window(dlg)
        base_sku = dlg.selected
        if not base_sku:
            return
        state = db[base_sku]

        # --- Correct DataFrame loading for orient='records' ---
        self.input_df = pd.DataFrame(state['input_df'])
        self.master_df = pd.DataFrame(state['master_df'])
        in_tree_df = pd.DataFrame(state['in_tree_df'])

        # Restore option names for UI logic
        self.option_names = sorted({
            name for blob in self.input_df['Options']
            for name, _ in VALUE_RE.findall(str(blob))
        })

        self.base_price_entry.delete(0, 'end')
        self.base_price_entry.insert(0, state.get('base_price', ''))
        self.base_weight_entry.delete(0, 'end')
        self.base_weight_entry.insert(0, state.get('base_weight', ''))
        self.base_sku_entry.delete(0, 'end')
        self.base_sku_entry.insert(0, state.get('base_sku', ''))
        self.prefix_entry.delete(0, 'end')
        self.prefix_entry.insert(0, state.get('prefix', ''))  # Load the prefix
        self.base_sku = self.base_sku_entry.get().strip()
        self._show_last_export_time(self.base_sku)

        # Fill trees and UI
        self._fill_tree(self.in_tree, in_tree_df)
        self._fill_tree(self.out_tree, self.master_df)

        # Reconfigure the "Configure" panel
        names = sorted(self.master_df['Name'].unique())
        self.name_combo['values'] = names
        if names:
            self.name_combo.set(names[0])
            self.populate_value_grid()

        # finally regenerate prices & weights on the left preview
        self._regenerate_left_preview()
        self.current_prog_path = db_path
        self.unsaved_changes = False
        self._update_title()
        self._resize_bottom_panels()

    def open_cost_import_window(self):
        import_file = {"path": None}
        data = {"df": None}

        win = tk.Toplevel(self)
        win.title("Import Pricing Data")
        win.geometry("800x600")

        frm = ttk.Frame(win)
        frm.pack(fill='both', expand=True, padx=10, pady=10)

        prefix_frame = ttk.Frame(frm)
        prefix_frame.pack(fill='x', pady=5)
        ttk.Label(prefix_frame, text="Prefix to add to all part numbers:").pack(side='left')
        prefix_var = tk.StringVar()
        prefix_entry = ttk.Entry(prefix_frame, textvariable=prefix_var, width=15)
        prefix_entry.pack(side='left', padx=5)

        # Top: Import button
        import_btn = ttk.Button(frm, text="Import XLSX", command=lambda: import_xlsx())
        import_btn.pack(anchor='w')

        tree_frame = ttk.Frame(frm)
        tree_frame.pack(fill='both', expand=True, pady=10)

        tree = ttk.Treeview(tree_frame, show='headings')
        tree.grid(row=0, column=0, sticky='nsew')

        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        vsb.grid(row=0, column=1, sticky='ns')

        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
        hsb.grid(row=1, column=0, sticky='ew')

        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # Inside open_cost_import_window, after col_frame.pack(...)
        omit_frame = ttk.Frame(frm)
        omit_frame.pack(fill='x', pady=5)
        ttk.Label(omit_frame, text="Omit first N rows:").pack(side='left')
        omit_var = tk.IntVar(value=0)
        omit_spin = ttk.Spinbox(omit_frame, from_=0, to=100, width=5, textvariable=omit_var)
        omit_spin.pack(side='left', padx=5)

        # Column selection
        col_frame = ttk.Frame(frm)
        col_frame.pack(fill='x', pady=5)
        ttk.Label(col_frame, text="Part Number Column:").pack(side='left')
        part_col = ttk.Combobox(col_frame, state='readonly')
        part_col.pack(side='left', padx=5)
        ttk.Label(col_frame, text="Price Column:").pack(side='left')
        price_col = ttk.Combobox(col_frame, state='readonly')
        price_col.pack(side='left', padx=5)

        # Save button
        save_btn = ttk.Button(frm, text="Save to Cost DB", state='disabled')
        save_btn.pack(anchor='e', pady=5)

        # Explore button
        explore_btn = ttk.Button(frm, text="Explore Cost DB", command=self.open_cost_db_explorer)
        explore_btn.pack(anchor='e', pady=5)

        def import_xlsx():
            from tkinter import filedialog
            import pandas as pd
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if not path:
                return
            try:
                df = pd.read_excel(path, header=None)
            except Exception as e:
                messagebox.showerror("Import Error", f"Could not read file:\n{e}")
                return
            omit = omit_var.get()
            if omit > 0:
                df = df.iloc[omit:].reset_index(drop=True)
            # Rename columns to A, B, C, ...
            df.columns = [excel_colname(i) for i in range(len(df.columns))]
            import_file["path"] = path
            data["df"] = df
            tree.delete(*tree.get_children())
            tree["columns"] = list(df.columns)
            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, anchor='center')
            for row in df.itertuples(index=False):
                tree.insert('', 'end', values=row)
            part_col["values"] = list(df.columns)
            price_col["values"] = list(df.columns)

            def save_to_cost_db():
                df = data.get("df")
                if df is None:
                    return
                part = part_col.get()
                price = price_col.get()
                if not part or not price:
                    messagebox.showwarning("Select Columns", "Please select both columns.")
                    return
                cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
                if os.path.exists(cost_db_path):
                    with open(cost_db_path, "r", encoding="utf-8") as f:
                        cost_db = json.load(f)
                else:
                    cost_db = {}
                for _, row in df.iterrows():
                    raw_pn = row[part]
                    # If it's a float and ends with .0, convert to int first
                    if isinstance(raw_pn, float) and raw_pn.is_integer():
                        pn = prefix_var.get() + str(int(raw_pn))
                    else:
                        pn = prefix_var.get() + str(raw_pn).strip()
                    try:
                        pr = round(float(row[price]), 2)
                    except Exception:
                        pr = str(row[price]).strip()
                    if pn:
                        cost_db[pn] = pr
                with open(cost_db_path, "w", encoding="utf-8") as f:
                    json.dump(cost_db, f, indent=2)
                messagebox.showinfo("Saved", f"Imported {len(df)} parts to pricing database.")
                self.update_option_costs_from_cost_db()  # <-- Add this line
                self.recalculate_all_pricing()

            save_btn.config(command=save_to_cost_db)

        def highlight_columns(*args):
            # Remove all tags
            for iid in tree.get_children():
                for col in tree["columns"]:
                    tree.set(iid, col, tree.item(iid, 'values')[tree["columns"].index(col)])
            # Highlight selected columns
            for col in tree["columns"]:
                col_idx = tree["columns"].index(col)
                tag = ""
                if col == part_col.get():
                    tag = "sku_col"
                elif col == price_col.get():
                    tag = "price_col"
                tree.tag_configure("sku_col", background="#ffe599")   # yellow
                tree.tag_configure("price_col", background="#b6d7a8") # green
                for iid in tree.get_children():
                    tree.item(iid, tags=(tag,))
            if part_col.get() and price_col.get():
                save_btn["state"] = "normal"
            else:
                save_btn["state"] = "disabled"

        part_col.bind("<<ComboboxSelected>>", highlight_columns)
        price_col.bind("<<ComboboxSelected>>", highlight_columns)

        def excel_colname(idx):
            name = ""
            while idx >= 0:
                name = chr(idx % 26 + ord('A')) + name
                idx = idx // 26 - 1
            return name


    def open_cost_db_explorer(self):
        db_path = os.path.join(self._get_database_folder(), "cost_db.json")
        if not os.path.exists(db_path):
            return messagebox.showinfo("No Cost DB", "Pricing database not found. Import a cost file first.")

        with open(db_path, "r", encoding="utf-8") as f:
            cost_db = json.load(f)

        win = tk.Toplevel(self)
        win.title("Option Pricing Data")
        win.geometry("600x500")

        # --- Search bar ---
        search_frame = ttk.Frame(win)
        search_frame.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Label(search_frame, text="Search:").pack(side="left")
        search_var = tk.StringVar(win)
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side="left", padx=(5, 0), fill="x", expand=True)

        # Define the treeview
        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        tree = ttk.Treeview(tree_frame, show="headings", selectmode="browse")
        tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=vsb.set)
        tree["columns"] = ["Part Number", "Price"]
        tree.heading("Part Number", text="Part Number")
        tree.heading("Price", text="Price")
        tree.column("Part Number", width=200, anchor='w')
        tree.column("Price", width=100, anchor='center')

        # Helper to populate tree with filter
        def populate_tree(filter_text=""):
            try:
                with open(db_path, "r", encoding="utf-8") as f:
                    current_cost_db = json.load(f)
            except Exception:
                current_cost_db = {}
            # clear existing rows
            for iid in tree.get_children():
                tree.delete(iid)
            lowft = filter_text.lower()
            for part, price in current_cost_db.items():
                if lowft in str(part).lower():
                    tree.insert("", "end", values=(part, price))

        def on_search(*args):
            filter_text = search_var.get()
            populate_tree(filter_text)

        # bind the trace _after_ defining on_search
        search_var.trace_add("write", on_search)

        # --- Editing price cell on double-click ---
        def on_double_click(event):
            item = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            if not item or col != "#2":
                return
            x, y, width, height = tree.bbox(item, "Price")
            value = tree.set(item, "Price")
            entry = ttk.Entry(tree)
            entry.place(x=x, y=y, width=width, height=height)
            entry.insert(0, value)
            entry.focus_set()

            def save_edit(event=None):
                new_val = entry.get()
                tree.set(item, "Price", new_val)
                entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", lambda e: entry.destroy())

        tree.bind("<Double-1>", on_double_click)

        # --- Save button ---
        def save_all():
            for iid in tree.get_children():
                part, price = tree.item(iid, 'values')
                try:
                    price_val = float(price)
                except Exception:
                    price_val = price
                cost_db[part] = price_val
            with open(db_path, "w", encoding="utf-8") as f:
                json.dump(cost_db, f, indent=2)
            with open(db_path, "r", encoding="utf-8") as f:
                cost_db.clear()
                cost_db.update(json.load(f))
            self.update_option_costs_from_cost_db()
            self.recalculate_all_pricing()
            messagebox.showinfo("Saved", "All changes saved to pricing database. Make sure to upload to cloud using ctrl + shift + s.")

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x', padx=10, pady=(0, 10))
        ttk.Button(btn_frame, text="Save Changes", command=save_all).pack(side='right')

        self.update_option_costs_from_cost_db()
        populate_tree()  # Fill on open


    def _update_last_export_time(self, base_sku):
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            return
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
            if base_sku in db:
                db[base_sku]['last_export'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            else:
                # If base_sku not in db, add it as a dict with just last_export
                db[base_sku] = {'last_export': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            with open(db_path, "w", encoding="utf-8") as f:
                json.dump(db, f, indent=2)
        except Exception as e:
            print("Error updating last export time:", e)

    
    def _show_last_export_time(self, base_sku=None):
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            self.last_export_label.config(text="")
            return
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
            if not base_sku:
                base_sku = self.base_sku_entry.get().strip() or self.base_sku
            last_export = db.get(base_sku, {}).get('last_export', None)
            if last_export:
                self.last_export_label.config(text=f"Last Excel Export: {last_export}")
            else:
                self.last_export_label.config(text="No Excel export yet")
        except Exception as e:
            print("Error showing last export time:", e)
            self.last_export_label.config(text="")

    def _get_latest_database_path(self):
        folder = self._get_database_folder()
        temp_path = os.path.join(folder, "sku_database_temp.json")
        if os.path.exists(temp_path):
            return temp_path
        files = [f for f in os.listdir(folder) if f.startswith("sku_database_") and f.endswith(".json")]
        if not files:
            return os.path.join(folder, "sku_database.json")
        files.sort(reverse=True)
        return os.path.join(folder, files[0])
    
    def delete_base_sku(self):
        """Delete a base SKU from the database"""
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            messagebox.showinfo("No Database", "No database file found.")
            return
            
        # Load database once
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except json.JSONDecodeError:
            if try_repair_json_file(db_path):
                messagebox.showinfo("Database Repaired", "Database file was corrupted but has been repaired.")
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
            else:
                messagebox.showerror("Database Error", "Database file is corrupted and could not be repaired.")
                return
            
        base_skus = list(db.keys())
        if not base_skus:
            messagebox.showinfo("No Data", "No SKUs in database.")
            return
    
        # Create dialog to select a base SKU
        class SkuDeleteDialog(tk.Toplevel):
            def __init__(self, parent, skus):
                super().__init__(parent)
                self.title("Delete Base SKU")
                self.transient(parent)
                self.grab_set()
                self.geometry("400x400")
                self.selected = None
                
                frame = ttk.Frame(self)
                frame.pack(fill="both", expand=True, padx=10, pady=10)
                
                ttk.Label(frame, text="Select a base SKU to delete:").pack(fill="x", pady=5)
                
                # Search box
                search_frame = ttk.Frame(frame)
                search_frame.pack(fill="x", pady=5)
                ttk.Label(search_frame, text="Search:").pack(side="left")
                self.search_var = tk.StringVar()
                self.search_var.trace_add("write", self.filter_list)
                search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
                search_entry.pack(side="left", fill="x", expand=True, padx=5)
                
                # SKU listbox with scrollbar
                list_frame = ttk.Frame(frame)
                list_frame.pack(fill="both", expand=True, pady=5)
                scrollbar = ttk.Scrollbar(list_frame)
                scrollbar.pack(side="right", fill="y")
                
                self.sku_list = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
                self.sku_list.pack(side="left", fill="both", expand=True)
                scrollbar.config(command=self.sku_list.yview)
                
                # Store all SKUs and populate the list
                self.all_skus = sorted(skus)
                self.populate_list()
                
                # Buttons
                btn_frame = ttk.Frame(frame)
                btn_frame.pack(fill="x", pady=10)
                ttk.Button(btn_frame, text="Delete", command=self.on_delete).pack(side="right", padx=5)
                ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="right")
                
                # Focus search box
                search_entry.focus_set()
                
            def populate_list(self, skus=None):
                if skus is None:
                    skus = self.all_skus
                self.sku_list.delete(0, tk.END)
                for sku in skus:
                    self.sku_list.insert(tk.END, sku)
                    
            def filter_list(self, *args):
                search_text = self.search_var.get().lower()
                if not search_text:
                    self.populate_list()
                else:
                    filtered = [sku for sku in self.all_skus if search_text in sku.lower()]
                    self.populate_list(filtered)
                    
            def on_delete(self):
                if not self.sku_list.curselection():
                    messagebox.showwarning("Selection Required", "Please select a SKU to delete.")
                    return
                self.selected = self.sku_list.get(self.sku_list.curselection())
                self.destroy()
        
        # Show the dialog and wait for result
        dlg = SkuDeleteDialog(self, base_skus)
        self.wait_window(dlg)
        selected_sku = dlg.selected
        
        # If no SKU was selected, return
        if not selected_sku:
            return
            
        # Confirm deletion
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete base SKU '{selected_sku}'? This action cannot be undone."):
            return
            
        # Delete the SKU from the database
        del db[selected_sku]
        
        # Save the updated database
        with open(db_path, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=2)
            
        # Mark database as modified
        self._mark_unsaved()
        
        # Update UI elements if needed
        if hasattr(self, 'sku_combo') and self.sku_combo.winfo_exists():
            current = self.sku_combo.get()
            values = list(self.sku_combo['values'])
            if selected_sku in values:
                values.remove(selected_sku)
                self.sku_combo['values'] = values
                if current == selected_sku:
                    self.sku_combo.set('')
        
        # Show success message
        messagebox.showinfo("Success", f"Base SKU '{selected_sku}' has been deleted from the database.")
        
        # If option master tree is populated, update it
        self._update_sku_history()

    def open_options_window(self):
        win = tk.Toplevel(self)
        win.title("Options")
        win.resizable(True, True)  # Make window resizable

        frm = ttk.Frame(win)
        frm.pack(fill='both', expand=True, padx=20, pady=20)

        ttk.Label(frm, text="Database Folder:").grid(row=0, column=0, sticky='w')
        db_var = tk.StringVar(value=self.options.get("database_folder", ""))
        db_entry = ttk.Entry(frm, textvariable=db_var, width=35)
        db_entry.grid(row=0, column=1, sticky='w')
        def browse_folder():
            folder = filedialog.askdirectory(initialdir=get_data_dir())
            if folder:
                db_var.set(folder)
        ttk.Button(frm, text="Browse...", command=browse_folder).grid(row=0, column=2, padx=5)

        def save_opts():
            self.options["database_folder"] = db_var.get()
            self._save_options()
            win.destroy()
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=1, column=0, columnspan=3, sticky='e', pady=(15,0))
        ttk.Button(btn_frame, text="Save", command=save_opts).pack(side='right', padx=(0,5))
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side='right')

        win.update_idletasks()
        win.minsize(win.winfo_reqwidth(), win.winfo_reqheight())  # Scale to fit contents

    

    def open_spotcheck_window(self):
        if not hasattr(self, "option_names") or not self.option_names:
            messagebox.showwarning("No Options", "Parse a table first.")
            return

        win = tk.Toplevel(self)
        win.title("Spot Check Configurator")
        win.geometry("400x500")
        frm = ttk.Frame(win)
        frm.pack(fill='both', expand=True, padx=20, pady=20)

        # Store dropdowns and comboboxes
        dropdowns = {}
        comboboxes = {}
        for name in self.option_names:
            ttk.Label(frm, text=f"{name}:", font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(10,0))
            values = list(self.master_df[self.master_df['Name'] == name]['Value'])
            var = tk.StringVar()
            cb = ttk.Combobox(frm, values=values, textvariable=var, state='readonly')
            cb.pack(fill='x', pady=2)
            dropdowns[name] = var
            comboboxes[name] = cb
            # Do NOT set var.set(values[0]) -- leave blank by default

        result_lbl = ttk.Label(frm, text="", font=('Segoe UI', 10, 'bold'), foreground='blue')
        result_lbl.pack(pady=(20,0))

        def update_result(*args):
            selections = {name: var.get() for name, var in dropdowns.items()}
            if not all(selections.values()):
                result_lbl.config(text="")
                return
            df = self._get_tree_df(self.in_tree)
            match = df.copy()
            for name, val in selections.items():
                if name in match.columns:
                    match = match[match[name] == val]
            if not match.empty:
                row = match.iloc[0]
                result_lbl.config(
                    text=f"Old SKU: {row['SKU']}\nNew SKU: {row['New SKU']}\nPrice: ${row['Price']}\nWeight: {row['Weight']} lb"
                )
            else:
                result_lbl.config(text="No matching SKU found.")

        for var in dropdowns.values():
            var.trace_add('write', lambda *a: update_result())

        # Bottom frame for buttons, right-aligned
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x', side='bottom', anchor='e', pady=(0, 10), padx=20)

        def highlight():
            selections = {name: var.get() for name, var in dropdowns.items()}
            if not all(selections.values()):
                messagebox.showwarning("Incomplete", "Please select a value for every option.")
                return
            df = self._get_tree_df(self.in_tree)
            match = df.copy()
            for name, val in selections.items():
                if name in match.columns:
                    match = match[match[name] == val]
            if not match.empty:
                row_idx = match.index[0]
                item_id = self.in_tree.get_children()[row_idx]
                self.in_tree.selection_set(item_id)
                self.in_tree.see(item_id)
                self._on_in_tree_select(None)
            else:
                messagebox.showinfo("Not found", "No matching SKU found.")

        def close():
            win.destroy()

        ttk.Button(btn_frame, text="Highlight", command=highlight).pack(side='right', padx=(0, 10))
        ttk.Button(btn_frame, text="Close", command=close).pack(side='right')

    def _resize_bottom_panels(self, event=None):
        # Get the width of the bottom frame
        bottom_frame = event.widget if event else self.out_tree.master.master
        total_width = bottom_frame.winfo_width()
        if total_width < 100:  # Avoid dividing by zero or too small
            return
        panel_width = total_width // 2

        # Resize out_tree columns
        out_cols = self.out_tree['columns']
        if out_cols:
            col_width = max(60, panel_width // len(out_cols))
            for col in out_cols:
                self.out_tree.column(col, width=col_width, stretch=True)

        # Resize breakdown_tree columns
        bd_cols = self.breakdown_tree['columns']
        if bd_cols:
            col_width = max(60, panel_width // len(bd_cols))
            for col in bd_cols:
                               self.breakdown_tree.column(col, width=col_width, stretch=True)

    def _on_quit(self):
        if self.unsaved_changes:
            res = messagebox.askyesnocancel("Unsaved Changes", "You have unsaved changes. Save before quitting?")
            if res is None:
                return  # Cancel quit
            elif res:
                self.save_to_database(temp=True)
        push_database_to_cloud()
        global web_service_process
        if web_service_process is not None:
            try:
                import requests
                requests.post('http://localhost:5000/shutdown', timeout=2)
            except Exception:
                pass
            try:
                web_service_process.wait(timeout=5)
            except Exception:
                try:
                    web_service_process.kill()
                except Exception:
                    pass
            web_service_process = None
        self.destroy()

    def _on_out_tree_double_click(self, event):
        # Identify row and column
        tree = self.out_tree
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_idx = int(col_id.replace('#', '')) - 1
        columns = tree['columns']
        if col_idx >= len(columns):
            return
        col_name = columns[col_idx]
        values = tree.item(row_id, 'values')
        if not values or col_name not in ("Add'l Cost", "Add'l Weight", "Associated SKUs"):
            return
        # Find the option in the configure panel
        name = values[0]  # "Name" column
        value = values[1] # "Value" column
        # Set the name in the combobox
        self.name_combo.set(name)
        self.populate_value_grid()
        # Find the entry widget for this value
        for idx, row in self.master_df[self.master_df['Name'] == name].sort_values('Value').iterrows():
            if row['Value'] == value:
                widget = self.entry_widgets.get(idx)
                if widget:
                    if col_name == "Add'l Cost":
                        widget['Cost'].focus_set()
                        widget['Cost'].select_range(0, 'end')
                    elif col_name == "Add'l Weight":
                        widget['Weight'].focus_set()
                        widget['Weight'].select_range(0, 'end')
                    elif col_name == "Associated SKUs":
                        widget['SKUs'].focus_set()
                        widget['SKUs'].select_range(0, 'end')
                break

    def _on_out_tree_select(self, event):
        selected = self.out_tree.selection()
        if not selected:
            return
        item = selected[0]
        values = self.out_tree.item(item, 'values')
        if not values or len(values) < 2:
            return
        name, value = values[0], values[1]
        # Highlight all SKUs in in_tree that have this option value
        tree = self.in_tree
        tree.selection_remove(tree.selection())
        for iid in tree.get_children():
            row = tree.item(iid, 'values')
            columns = tree['columns']
            if name in columns:
                idx = columns.index(name)
                if row[idx] == value:
                    tree.selection_add(iid)
                    tree.see(iid)

    def _on_in_tree_select(self, event):
        # Show breakdown for selected SKU in in_tree
        selected = self.in_tree.selection()
        if not selected:
            return
        item = selected[0]
        values = self.in_tree.item(item, 'values')
        if not values:
            return
        # Map columns to values
        columns = self.in_tree['columns']
        row = dict(zip(columns, values))
        # Find matching options in master_df
        breakdown = []
        total_cost = 0
        total_weight = 0
    
        # Get base cost and weight
        try:
            base_cost = float(self.base_price_entry.get())
        except Exception:
            base_cost = 0
        try:
            base_weight = float(self.base_weight_entry.get())
        except Exception:
            base_weight = 0
    
        # Add base at the top
        breakdown.append(("BASE", "", base_cost, base_weight))
        total_cost += base_cost
        total_weight += base_weight
    
        for name in self.option_names:
            val = row.get(name, '')
            if val:
                match = self.master_df[(self.master_df['Name'] == name) & (self.master_df['Value'] == val)]
                if not match.empty:
                    cost = match.iloc[0].get("Add'l Cost", 0)
                    weight = match.iloc[0].get("Add'l Weight", 0)
                    try:
                        cost = float(cost)
                    except Exception:
                        cost = 0
                    try:
                        weight = float(weight)
                    except Exception:
                        weight = 0
                    breakdown.append((name, val, cost, weight))
                    total_cost += cost
                    total_weight += weight
    
        # Add total at the bottom
        breakdown.append(("TOTAL", "", round(total_cost, 2), round(total_weight, 2)))
    
        # Fill breakdown_tree
        self.breakdown_tree.delete(*self.breakdown_tree.get_children())
        for tup in breakdown:
            self.breakdown_tree.insert('', 'end', values=tup)

    def _fill_tree(self, tree, df):
        # Clear and fill a Treeview from a DataFrame
        tree.delete(*tree.get_children())
        if df is None or df.empty:
            return
        tree['columns'] = list(df.columns)
        for col in df.columns:
            tree.heading(col, text=col)
            # Optionally, auto-size columns or set a default width
            tree.column(col, width=120, stretch=True)
        for row in df.itertuples(index=False):
            tree.insert('', 'end', values=row)

    def _get_tree_df(self, tree):
        # Convert a Treeview to a DataFrame
        cols = tree['columns']
        if not cols or (isinstance(cols, str) and not cols.strip()):
            return pd.DataFrame()  # Return empty DataFrame if no columns
        rows = [tree.item(iid)['values'] for iid in tree.get_children()]
        return pd.DataFrame(rows, columns=cols)

    def _get_database_folder(self):
        # Return the current database folder
        return get_data_dir()

    def _is_modifier(self, event):
        # Utility: return True if event is a modifier key (Shift, Ctrl, etc.)
        return event.state & 0x4 or event.state & 0x1 or event.state & 0x20000


    def generate_new_skus(self):
        # Get the base SKU
        base_sku = self.base_sku_entry.get().strip() or self.base_sku
        if not base_sku:
            return
    
        # Remove text in parentheses and spaces for new SKU generation
        import re
        # Remove anything in parentheses (including the parentheses)
        base_sku_clean = re.sub(r'\([^)]*\)', '', base_sku)
        # Remove all spaces
        base_sku_clean = base_sku_clean.replace(' ', '')
        # Strip any remaining whitespace
        base_sku_clean = base_sku_clean.strip()
        
        if not base_sku_clean:
            return
    
        # Get the prefix
        prefix = self.prefix_entry.get().strip()
        
        # Check for special zero-removal prefixes
        zero_removal = 0
        if prefix.startswith('-') and len(prefix) > 1 and prefix[1:].isdigit():
            zero_removal = int(prefix[1:])
            prefix = ""  # Clear the prefix since it was just for zero removal
    
        # Get the DataFrame from the in_tree
        df = self._get_tree_df(self.in_tree)
        if df.empty or 'New SKU' not in df.columns:
            return
    
        # Generate new SKUs
        for idx in range(len(df)):
            # Calculate the number padding based on zero removal
            padding = max(1, 4 - zero_removal)  # Minimum padding of 1, default 4
            number_str = str(idx + 1).zfill(padding)
            
            if prefix:
                new_sku = f"{base_sku_clean}-{prefix}{number_str}"
            else:
                new_sku = f"{base_sku_clean}-{number_str}"
            
            df.at[idx, 'New SKU'] = new_sku
    
        # Check if the first SKU is over 20 characters and warn
        if not df.empty:
            first_sku = df.iloc[0]['New SKU']
            if len(first_sku) > 20:
                import tkinter.messagebox as messagebox
                messagebox.showwarning(
                    "SKU Length Warning", 
                    f"Generated SKUs are {len(first_sku)} characters long, which exceeds the 20 character limit.\n\n"
                    f"Example: {first_sku}\n\n"
                    "Consider using a shorter base SKU or prefix."
                )
    
        # Update the treeview
        self._fill_tree(self.in_tree, df)


    def _clear_highlight_if_needed(self, event=None):
        # Only clear if click was outside out_tree and in_tree
        widgets = [self.out_tree, self.in_tree]
        widget = event.widget if event else None
        if widget not in widgets:
            self.out_tree.selection_remove(self.out_tree.selection())
            self.in_tree.selection_remove(self.in_tree.selection())

    def _prev_name(self):
        # Move to previous option name in the configure panel
        names = self.name_combo['values']
        if not names:
            return
        idx = names.index(self.name_combo.get())
        if idx > 0:
            self.name_combo.set(names[idx - 1])
            self.populate_value_grid()

    def _next_name(self):
        # Move to next option name in the configure panel
        names = self.name_combo['values']
        if not names:
            return
        idx = names.index(self.name_combo.get())
        if idx < len(names) - 1:
            self.name_combo.set(names[idx + 1])
            self.populate_value_grid()

    def update_option_costs_from_cost_db(self):
        # Load cost DB
        cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
        if not os.path.exists(cost_db_path):
            return
        with open(cost_db_path, "r", encoding="utf-8") as f:
            cost_db = json.load(f)
        # Update all Associated SKUs with matching part numbers
        for idx, row in self.master_df.iterrows():
            assoc_str = row.get("Associated SKUs", "")
            sku_costs = parse_associated_skus(assoc_str)
            updated = False
            new_sku_costs = []
            for sku, cost, partnumber in sku_costs:
                # If partnumber is present and in cost_db, update cost
                if partnumber and partnumber in cost_db:
                    new_cost = str(cost_db[partnumber])
                    if new_cost != cost:
                        updated = True
                    new_sku_costs.append((sku, new_cost, partnumber))
                # If no partnumber, but sku is in cost_db, update cost
                elif not partnumber and sku in cost_db:
                    new_cost = str(cost_db[sku])
                    if new_cost != cost:
                        updated = True
                    new_sku_costs.append((sku, new_cost, partnumber))
                else:
                    new_sku_costs.append((sku, cost, partnumber))
            if updated:
                self.master_df.at[idx, "Associated SKUs"] = format_associated_skus(new_sku_costs)
                # Optionally, update Add'l Cost as sum of all
                try:
                    total_cost = sum(float(c) if c else 0 for _, c, _ in new_sku_costs)
                except Exception:
                    total_cost = 0
                self.master_df.at[idx, "Add'l Cost"] = str(total_cost)
        self._fill_tree(self.out_tree, self.master_df)
        self._regenerate_left_preview()
        self.populate_value_grid()

    # Add "Edit SKUs" button if multiple SKUs
    def open_sku_cost_editor(self, idx):
        if hasattr(self, "sku_editor_window") and self.sku_editor_window is not None and tk.Toplevel.winfo_exists(self.sku_editor_window):
            messagebox.showwarning("Window Already Open", "Please close the current 'Edit Associated SKUs' window before opening another.")
            return
        
        editor = tk.Toplevel(self)
        editor.title("Edit Associated SKUs")
        editor.geometry("600x350")  # Made wider by default
        editor.resizable(True, True)  # Make window resizable
        editor.transient(self)
        editor.grab_set()
        self.sku_editor_window = editor
        
        # Configure grid weights for resizing
        editor.columnconfigure(0, weight=1)
        editor.rowconfigure(999, weight=1)  # Space before buttons
        
        # Get cost database for pricing lookups
        cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
        if os.path.exists(cost_db_path):
            with open(cost_db_path, "r", encoding="utf-8") as f:
                cost_db = json.load(f)
            part_numbers = list(cost_db.keys())
        else:
            part_numbers = []
            cost_db = {}
        
        # Parse existing associated SKUs
        current_value = self.master_df.iloc[idx].get("Associated SKUs", "")
        sku_costs = parse_associated_skus(current_value)
        sku_list = [sku for sku, _, _ in sku_costs]
        
        rows = []  # Store references to entry widgets
        
        # Custom dropdown lists - one for each entry
        dropdowns = []
        
        def build_rows(editor, sku_list):
            # Clear existing rows
            for widget in editor.winfo_children():
                if widget not in [btn_frame, add_btn]:
                    widget.destroy()
            rows.clear()
            
            # Clear all dropdowns
            for dropdown in dropdowns:
                if dropdown[0].winfo_exists():
                    dropdown[0].destroy()
            dropdowns.clear()
        
            for i, sku in enumerate(sku_list):
                row_frame = ttk.Frame(editor)
                row_frame.grid(row=i, column=0, sticky='ew', padx=5, pady=2)
                row_frame.columnconfigure(1, weight=1)  # Make entry column expandable
                
                ttk.Label(row_frame, text="SKU / Part #:").grid(row=0, column=0, sticky='w')
                
                # Use Entry instead of Combobox - now fills available space
                sku_entry = ttk.Entry(row_frame)
                sku_entry.grid(row=0, column=1, sticky='ew', padx=(5, 5))
                sku_entry.insert(0, sku)
                rows.append(sku_entry)
                
                # Minus button anchored to the right
                minus_btn = ttk.Button(row_frame, text="−", width=2, command=lambda idx=i: remove_row(idx))
                minus_btn.grid(row=0, column=2, sticky='e')
                
                # Create dropdown listbox for this entry
                dropdown = tk.Listbox(editor, height=6)
                dropdown.place_forget()  # Hide initially
                dropdowns.append((dropdown, sku_entry))
                
                # Bind keyboard events for filtering
                def show_dropdown(event, entry=sku_entry, dropdown=dropdown):
                    value = entry.get().lower()
                    
                    # Filter the dropdown list based on entry text
                    dropdown.delete(0, tk.END)
                    matches = [item for item in part_numbers if value in item.lower()]
                    
                    for match in matches:
                        dropdown.insert(tk.END, match)
                    
                    if matches:
                        # Position dropdown below entry
                        x = entry.winfo_rootx() - editor.winfo_rootx()
                        y = entry.winfo_rooty() - editor.winfo_rooty() + entry.winfo_height()
                        dropdown.place(x=x, y=y, width=entry.winfo_width())
                    else:
                        dropdown.place_forget()
                    
                    # Color the entry based on whether it's a known SKU
                    if entry.get() in part_numbers:
                        entry.config(foreground='green')
                    else:
                        entry.config(foreground='black')
                
                # Select from dropdown and hide it
                def select_from_dropdown(event, entry=sku_entry, dropdown=dropdown):
                    if dropdown.curselection():
                        value = dropdown.get(dropdown.curselection())
                        entry.delete(0, tk.END)
                        entry.insert(0, value)
                        dropdown.place_forget()
                        entry.focus_set()
                    
                # Bind events
                sku_entry.bind('<KeyRelease>', show_dropdown)
                dropdown.bind('<ButtonRelease-1>', select_from_dropdown)
                dropdown.bind('<Return>', select_from_dropdown)
                
                # Hide dropdown when entry loses focus
                sku_entry.bind('<FocusOut>', lambda e, d=dropdown: 
                                editor.after(100, lambda: d.place_forget() if d.winfo_viewable() else None))
                
                # Navigate to dropdown with Down key
                sku_entry.bind('<Down>', lambda e, d=dropdown: 
                                (d.focus_set(), d.selection_set(0)) if d.winfo_viewable() else None)
        
            def remove_row(idx):
                new_sku_list = [se.get().strip() for j, se in enumerate(rows) if j != idx]
                build_rows(editor, new_sku_list)
        
        def add_row():
            new_sku_list = [se.get().strip() for se in rows]
            new_sku_list.append("")
            build_rows(editor, new_sku_list)
            if rows:  # Make sure rows exist before focusing
                rows[-1].focus_set()  # Focus the new row
        
        def save_and_close():
            new_skus = [se.get().strip() for se in rows if se.get().strip()]
            if not new_skus:
                self.entry_widgets[idx]['SKUs'].delete(0, 'end')
                self.entry_widgets[idx]['SKUs'].insert(0, '')
                # Update master_df and cost
                self.master_df.at[idx, "Associated SKUs"] = ''
                self.master_df.at[idx, "Add'l Cost"] = '0'
                self._fill_tree(self.out_tree, self.master_df)
                self._regenerate_left_preview()
                self._mark_unsaved()
                editor.destroy()
                return
    
            mapped_parts = []
            missing_parts = []
            for sku in new_skus:
                if sku in cost_db:
                    mapped_parts.append(sku)
                else:
                    missing_parts.append(sku)
    
            sku_costs = []
            for sku in new_skus:
                cost = str(cost_db[sku]) if sku in cost_db else '0'
                sku_costs.append((sku, cost, ''))
            new_str = format_associated_skus(sku_costs)
            self.entry_widgets[idx]['SKUs'].delete(0, 'end')
            self.entry_widgets[idx]['SKUs'].insert(0, new_str)
            total_cost = sum(float(cost_db[sku]) if sku in cost_db and str(cost_db[sku]).replace('.','',1).isdigit() else 0 for sku in new_skus)
            self.entry_widgets[idx]['Cost'].delete(0, 'end')
            self.entry_widgets[idx]['Cost'].insert(0, str(total_cost))
    
            # Update master_df and Option Master Table
            self.master_df.at[idx, "Associated SKUs"] = new_str
            self.master_df.at[idx, "Add'l Cost"] = str(total_cost)
            self._fill_tree(self.out_tree, self.master_df)
            self._regenerate_left_preview()
            self._mark_unsaved()
    
            editor.destroy()
    
            if mapped_parts:
                msg = "The following SKUs are mapped to imported part numbers and their prices will auto-update:\n"
                msg += "\n".join(mapped_parts)
                messagebox.showinfo("Associated SKU Mapping", msg)
            if missing_parts:
                msg = "Warning: The following SKUs are NOT in your imported pricing database. Their prices will NOT auto-update:\n"
                msg += "\n".join(missing_parts)
                msg += "\n\nWould you like to manually add pricing for these SKUs now?"
                if messagebox.askyesno("Missing Part Numbers", msg):
                    self.open_manual_cost_import(missing_parts)
    
            self.sku_editor_window = None
    
        # Bottom button frame - anchored to bottom
        btn_frame = ttk.Frame(editor)
        btn_frame.grid(row=1000, column=0, sticky='ew', padx=5, pady=5)
        btn_frame.columnconfigure(0, weight=1)  # Allow space to expand
        
        add_btn = ttk.Button(btn_frame, text="+ Add SKU", command=add_row)
        add_btn.grid(row=0, column=0, sticky='w')
        
        ttk.Button(btn_frame, text="Cancel", command=lambda: editor.destroy()).grid(row=0, column=1, sticky='e', padx=(0, 5))
        ttk.Button(btn_frame, text="Save & Close", command=save_and_close).grid(row=0, column=2, sticky='e')
        
        # Initial population
        build_rows(editor, sku_list)
        
        def on_close():
            self.sku_editor_window = None
            editor.destroy()
    
        editor.protocol("WM_DELETE_WINDOW", on_close)
        
        # Center the window on the screen
        editor.update_idletasks()
        width = editor.winfo_width()
        height = editor.winfo_height()
        x = (editor.winfo_screenwidth() // 2) - (width // 2)
        y = (editor.winfo_screenheight() // 2) - (height // 2)
        editor.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def open_manual_cost_import(self, missing_parts):
        win = tk.Toplevel(self)
        win.title("Manual Cost Entry")
        win.geometry("500x400")
    
        frm = ttk.Frame(win)
        frm.pack(fill='both', expand=True, padx=10, pady=10)
    
        ttk.Label(frm, text="Enter prices for missing SKUs/Part Numbers:", font=('Segoe UI', 11, 'bold')).pack(anchor='w', pady=(0,10))
    
        # Table headers
        table_frame = ttk.Frame(frm)
        table_frame.pack(fill='both', expand=True)
        ttk.Label(table_frame, text="Part Number", font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(table_frame, text="Price", font=('Segoe UI', 10, 'bold')).grid(row=0, column=1, padx=5, pady=5)
    
        entries = []
        for i, sku in enumerate(missing_parts):
            pn_entry = ttk.Entry(table_frame, width=25)
            pn_entry.grid(row=i+1, column=0, padx=5, pady=2)
            pn_entry.insert(0, sku)
            pn_entry.config(state='readonly')  # Prevent editing part number
            price_entry = ttk.Entry(table_frame, width=12)
            price_entry.grid(row=i+1, column=1, padx=5, pady=2)
            entries.append((pn_entry, price_entry))
    
        def save_prices():
            cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
            if os.path.exists(cost_db_path):
                with open(cost_db_path, "r", encoding="utf-8") as f:
                    cost_db = json.load(f)
            else:
                cost_db = {}
            updated = False
            for pn_entry, price_entry in entries:
                pn = pn_entry.get().strip()
                price = price_entry.get().strip()
                if pn and price:
                    try:
                        price_val = float(price)
                    except Exception:
                        messagebox.showwarning("Invalid Price", f"Invalid price for {pn}.")
                        return
                    cost_db[pn] = price_val
                    updated = True
            if updated:
                with open(cost_db_path, "w", encoding="utf-8") as f:
                    json.dump(cost_db, f, indent=2)
                self.update_option_costs_from_cost_db()
            win.destroy()
    
        btn_frame = ttk.Frame(frm)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Save", command=save_prices).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side='left', padx=5)
        self.recalculate_all_pricing()
    
    def export_cost_db(self):
        cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
        if not os.path.exists(cost_db_path):
            messagebox.showinfo("No Cost DB", "Pricing database not found. Import a cost file first.")
            return
        export_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            initialfile="cost_db_export.xlsx",
            filetypes=[('Excel', '*.xlsx')]
        )
        if not export_path:
            return
        try:
            with open(cost_db_path, "r", encoding="utf-8") as f:
                cost_db = json.load(f)
            import pandas as pd
            df = pd.DataFrame([
                {"Part Number": k, "Price": v}
                for k, v in cost_db.items()
            ])
            df.to_excel(export_path, index=False)
            if messagebox.askyesno("Exported", f"Exported pricing database to {export_path}\nOpen in Excel?"):
                try:
                    subprocess.Popen(['start', 'excel', export_path], shell=True)
                except Exception:
                    messagebox.showerror("Error", "Can't open Excel.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{e}")
        

    def _save_options(self):
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "options.json")
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.options, f, indent=2)
        except Exception as e:
            print("Error saving options:", e)
    
    def _load_options(self):
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "options.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    self.options.update(json.load(f))
            except Exception as e:
                print("Error loading options:", e)

    import tkinter as tk
    from tkinter import ttk
    def import_cost_db(self):
        if not messagebox.askyesno(
            "Overwrite Pricing Database",
            "Importing will completely overwrite the current pricing database. Are you sure you want to continue?"
        ):
            return
        import_path = filedialog.askopenfilename(
            filetypes=[('Excel', '*.xlsx'), ('CSV', '*.csv')]
        )
        if not import_path:
            return
        try:
            import pandas as pd
            if import_path.lower().endswith('.csv'):
                df = pd.read_csv(import_path)
            else:
                df = pd.read_excel(import_path)
            # Expect columns: Part Number, Price
            if "Part Number" not in df.columns or "Price" not in df.columns:
                messagebox.showerror("Import Error", "File must have 'Part Number' and 'Price' columns.")
                return
            cost_db = {}
            for _, row in df.iterrows():
                pn = str(row["Part Number"]).strip()
                price = row["Price"]
                try:
                    price = float(price)
                except Exception:
                    pass
                if pn:
                    cost_db[pn] = price

            cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
            if os.path.exists(cost_db_path):
                with open(cost_db_path, "r", encoding="utf-8") as f:
                    old_cost_db = json.load(f)
            else:
                old_cost_db = {}
            
            overridden = []
            for pn, price in cost_db.items():
                if pn in old_cost_db and old_cost_db[pn] != price:
                    overridden.append((pn, old_cost_db[pn], price))
            
            if overridden:
                win = tk.Toplevel(self)
                win.title("Overridden Part Numbers")
                win.geometry("500x400")
                frm = ttk.Frame(win)
                frm.pack(fill='both', expand=True, padx=10, pady=10)
                ttk.Label(frm, text="The following part numbers had their pricing overridden:").pack(anchor='w', pady=(0,10))
                tree = ttk.Treeview(frm, columns=("Part Number", "Old Price", "New Price"), show="headings")
                tree.heading("Part Number", text="Part Number")
                tree.heading("Old Price", text="Old Price")
                tree.heading("New Price", text="New Price")
                tree.column("Part Number", width=180)
                tree.column("Old Price", width=100)
                tree.column("New Price", width=100)
                tree.pack(fill='both', expand=True)
                for pn, old, new in overridden:
                    tree.insert('', 'end', values=(pn, old, new))
            
                def copy_to_clipboard():
                    rows = [("Part Number\tOld Price\tNew Price")]
                    for pn, old, new in overridden:
                        rows.append(f"{pn}\t{old}\t{new}")
                    self.clipboard_clear()
                    self.clipboard_append('\n'.join(rows))
                    self.update()  # now it stays on the clipboard after the window is closed
            
                btn = ttk.Button(frm, text="Copy", command=copy_to_clipboard)
                btn.pack(pady=10)

            cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
            with open(cost_db_path, "w", encoding="utf-8") as f:
                json.dump(cost_db, f, indent=2)

            # --- NEW: Check for missing SKUs in all Associated SKUs ---
            db_path = self._get_latest_database_path()
            if os.path.exists(db_path):
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
                missing_links = []
                for base_sku, state in db.items():
                    master_df = pd.DataFrame(state.get('master_df', []))
                    for idx, row in master_df.iterrows():
                        assoc_str = row.get("Associated SKUs", "")
                        sku_costs = parse_associated_skus(assoc_str)
                        for sku, cost, partnumber in sku_costs:
                            key = partnumber if partnumber else sku
                            if key and key not in cost_db:
                                missing_links.append({
                                    "Base SKU": base_sku,
                                    "Option Name": row.get("Name", ""),
                                    "Option Value": row.get("Value", ""),
                                    "Missing Part/SKU": key,
                                    "Old Price": cost
                                })
                if missing_links:
                    win = tk.Toplevel(self)
                    win.title("Missing Pricing References")
                    win.geometry("700x400")
                    ttk.Label(win, text="The following options reference SKUs/part numbers that are missing from the imported pricing database:").pack(anchor='w', padx=10, pady=(10,0))
                    tree = ttk.Treeview(win, columns=("Base SKU", "Option Name", "Option Value", "Missing Part/SKU", "Old Price"), show="headings")
                    for col in tree["columns"]:
                        tree.heading(col, text=col)
                        tree.column(col, width=120)
                    tree.pack(fill='both', expand=True, padx=10, pady=10)
                    for row in missing_links:
                        tree.insert('', 'end', values=(row["Base SKU"], row["Option Name"], row["Option Value"], row["Missing Part/SKU"], row["Old Price"]))

                    btn_frame = ttk.Frame(win)
                    btn_frame.pack(pady=10)

                    def do_reimport():
                        # Add missing SKUs back to cost_db with old price
                        cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
                        if os.path.exists(cost_db_path):
                            with open(cost_db_path, "r", encoding="utf-8") as f:
                                cost_db = json.load(f)
                        else:
                            cost_db = {}
                        for row in missing_links:
                            sku = row["Missing Part/SKU"]
                            price = row["Old Price"]
                            try:
                                price_val = float(price)
                            except Exception:
                                price_val = price
                            if sku and sku not in cost_db:
                                cost_db[sku] = price_val
                        with open(cost_db_path, "w", encoding="utf-8") as f:
                            json.dump(cost_db, f, indent=2)
                        win.destroy()
                        # Now recalculate after resolving
                        self.update_option_costs_from_cost_db()
                        self.recalculate_all_pricing()

                    def do_remove():
                        if not messagebox.askyesno(
                            "Remove SKUs",
                            "This will remove all the missing SKUs/part numbers from every option in every base SKU in your database. "
                            "This cannot be undone. Are you sure you want to proceed?"
                        ):
                            return
                        db_path = self._get_latest_database_path()
                        affected_base_skus = []
                        if os.path.exists(db_path):
                            with open(db_path, "r", encoding="utf-8") as f:
                                db = json.load(f)
                            to_remove = set(row["Missing Part/SKU"] for row in missing_links)
                            for base_sku, state in db.items():
                                master_df = pd.DataFrame(state.get('master_df', []))
                                changed = False
                                for idx, row in master_df.iterrows():
                                    assoc_str = row.get("Associated SKUs", "")
                                    sku_costs = parse_associated_skus(assoc_str)
                                    # Remove any sku/partnumber in to_remove
                                    new_sku_costs = [
                                        (sku, cost, partnumber)
                                        for sku, cost, partnumber in sku_costs
                                        if (partnumber if partnumber else sku) not in to_remove
                                    ]
                                    if len(new_sku_costs) != len(sku_costs):
                                        changed = True
                                        if new_sku_costs:
                                            master_df.at[idx, "Associated SKUs"] = format_associated_skus(new_sku_costs)
                                            try:
                                                total_cost = sum(float(c) if c else 0 for _, c, _ in new_sku_costs)
                                            except Exception:
                                                total_cost = 0
                                            master_df.at[idx, "Add'l Cost"] = str(total_cost)
                                        else:
                                            master_df.at[idx, "Associated SKUs"] = ""
                                            master_df.at[idx, "Add'l Cost"] = "0"
                                if changed:
                                    state['master_df'] = master_df.to_dict(orient='records')
                                    affected_base_skus.append(base_sku)
                            with open(db_path, "w", encoding="utf-8") as f:
                                json.dump(db, f, indent=2)
                    
                            # --- Update in-memory DataFrame if current base SKU was changed ---
                            current_base_sku = self.base_sku_entry.get().strip() or self.base_sku
                            if current_base_sku in db:
                                state = db[current_base_sku]
                                self.master_df = pd.DataFrame(state.get('master_df', []))
                                self._fill_tree(self.out_tree, self.master_df)
                                self.populate_value_grid()
                                self._regenerate_left_preview()
                        win.destroy()
                        self.reload_current_base_sku()
                        self.update_option_costs_from_cost_db()
                        self.recalculate_all_pricing(affected_base_skus=affected_base_skus)

                    ttk.Button(btn_frame, text="Reimport", command=do_reimport).pack(side='left', padx=10)
                    ttk.Button(btn_frame, text="Remove SKUs from all Options", command=do_remove).pack(side='left', padx=10)
            # ...continue with update_option_costs_from_cost_db and recalc...            

            self.update_option_costs_from_cost_db()
            messagebox.showinfo("Imported", f"Imported {len(cost_db)} parts to pricing database.")
            self.recalculate_all_pricing()
            self.reload_current_base_sku()
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import:\n{e}")

    def recalculate_all_pricing(self, affected_base_skus=None):
        import threading
    
        def do_recalc():
            db_path = self._get_latest_database_path()
            if not os.path.exists(db_path):
                self.after(0, loading_win.destroy)
                return
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
            cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
            if os.path.exists(cost_db_path):
                with open(cost_db_path, "r", encoding="utf-8") as f:
                    cost_db = json.load(f)
            else:
                cost_db = {}
    
            # Use a set for affected SKUs, initialized from argument if provided
            local_affected = set(affected_base_skus) if affected_base_skus else set()
    
            # For each base SKU, recalculate all pricing
            for base_sku, state in db.items():
                try:
                    master_df = pd.DataFrame(state['master_df'])
                    orig_master_df = master_df.copy(deep=True)
                    changed = False
                    # Update option costs from cost_db
                    for idx, row in master_df.iterrows():
                        assoc_str = row.get("Associated SKUs", "")
                        sku_costs = parse_associated_skus(assoc_str)
                        new_sku_costs = []
                        for sku, cost, partnumber in sku_costs:
                            if partnumber and partnumber in cost_db:
                                new_cost = str(cost_db[partnumber])
                                new_sku_costs.append((sku, new_cost, partnumber))
                            elif not partnumber and sku in cost_db:
                                new_cost = str(cost_db[sku])
                                new_sku_costs.append((sku, new_cost, partnumber))
                            else:
                                new_sku_costs.append((sku, cost, partnumber))
                        new_assoc_str = format_associated_skus(new_sku_costs)
                        if new_assoc_str != master_df.at[idx, "Associated SKUs"]:
                            changed = True
                            master_df.at[idx, "Associated SKUs"] = new_assoc_str
                        # Update Add'l Cost as sum of all
                        try:
                            total_cost = sum(float(c) if c else 0 for _, c, _ in new_sku_costs)
                        except Exception:
                            total_cost = 0
                        if str(total_cost) != str(master_df.at[idx, "Add'l Cost"]):
                            changed = True
                            master_df.at[idx, "Add'l Cost"] = str(total_cost)
                    if changed:
                        local_affected.add(base_sku)
                    state['master_df'] = master_df.to_dict(orient='records')
                    
                    # --- Recalculate in_tree_df for this base_sku ---
                    try:
                        # Rebuild input_df and option_names for this base_sku
                        input_df = pd.DataFrame(state.get('input_df', []))
                        option_names = sorted({
                            name for blob in input_df['Options']
                            for name, _ in VALUE_RE.findall(str(blob))
                        }) if not input_df.empty and 'Options' in input_df.columns else []
                        # Prepare exploded options DataFrame
                        sku_rows = []
                        option_rows = []
                        for idx, row in enumerate(input_df.itertuples(), start=1):
                            sku = row.SKU
                            options = list(VALUE_RE.findall(str(row.Options)))
                            sku_rows.append({
                                '#': idx,
                                'SKU': sku,
                                'Options': options,
                                'New SKU': '',
                            })
                            for n, v in options:
                                option_rows.append({'#': idx, 'SKU': sku, 'Name': n, 'Value': v})
    
                        sku_df = pd.DataFrame(sku_rows)
                        option_df = pd.DataFrame(option_rows)
    
                        # Merge with master_df to get costs/weights for each option
                        if not option_df.empty:
                            merged = option_df.merge(master_df, on=['Name', 'Value'], how='left')
                            merged["Add'l Cost"] = pd.to_numeric(merged["Add'l Cost"], errors='coerce').fillna(0)
                            merged["Add'l Weight"] = pd.to_numeric(merged["Add'l Weight"], errors='coerce').fillna(0)
                            cost_weight = merged.groupby('#').agg({'Add\'l Cost': 'sum', 'Add\'l Weight': 'sum'}).reset_index()
                        else:
                            cost_weight = pd.DataFrame({'#': [], "Add'l Cost": [], "Add'l Weight": []})
                        sku_df = sku_df.merge(cost_weight, on='#', how='left').fillna({'Add\'l Cost': 0, 'Add\'l Weight': 0})
                        # Calculate final price/weight
                        try:
                            bp = float(state.get('base_price', 0))
                        except Exception:
                            bp = 0
                        try:
                            bw = float(state.get('base_weight', 0))
                        except Exception:
                            bw = 0
                        sku_df['Price'] = (bp + sku_df["Add'l Cost"]).round(2)
                        sku_df['Weight'] = (bw + sku_df["Add'l Weight"]).round(2)
                        # Prepare option columns
                        for name in option_names:
                            sku_df[name] = ''
                        for idx2, row2 in sku_df.iterrows():
                            for n, v in row2['Options']:
                                sku_df.at[idx2, n] = v
                        # Generate new SKUs
                        base = state.get('base_sku', base_sku)
                        for idx2 in range(len(sku_df)):
                            sku_df.at[idx2, 'New SKU'] = f"{base}-{str(idx2+1).zfill(4)}"
                        display_cols = ['#', 'SKU', 'New SKU', 'Price', 'Weight'] + option_names
                        state['in_tree_df'] = sku_df[display_cols].to_dict(orient='records')
                    except Exception as e:
                        print(f"Error recalculating in_tree_df for {base_sku}: {e}")
                except Exception as e:
                    print(f"Error recalculating for {base_sku}: {e}")
    
            # Save updated db
            with open(db_path, "w", encoding="utf-8") as f:
                json.dump(db, f, indent=2)
    
            # Hide loading window and show affected SKUs
            def show_result():
                loading_win.destroy()
                db_path = self._get_latest_database_path()
                if os.path.exists(db_path):
                    with open(db_path, "r", encoding="utf-8") as f:
                        latest_db = json.load(f)
                else:
                    latest_db = db  # fallback to in-memory if file missing
    
                if local_affected:
                    msg = f"{len(local_affected)} base SKUs had price changes:\n\n" + "\n".join(local_affected)
                    if messagebox.askyesno("Recalculation Complete", msg + "\n\nWould you like to dump these to Excel?"):
                        self.dump_affected_to_excel(list(local_affected), latest_db)
                else:
                    msg = "No base SKUs were affected by the price change."
                    messagebox.showinfo("Recalculation Complete", msg)
            self.after(0, show_result)
    
        # Show loading window
        loading_win = tk.Toplevel(self)
        loading_win.title("Recalculating Pricing")
        loading_win.geometry("350x120")
        ttk.Label(loading_win, text="Recalculating all pricing...\nThis may take a while.", anchor="center").pack(expand=True, pady=20)
        pb = ttk.Progressbar(loading_win, mode="indeterminate")
        pb.pack(fill="x", padx=20, pady=10)
        pb.start(10)
        loading_win.grab_set()
        loading_win.transient(self)
        loading_win.update()
    
        # Run recalculation in a thread
        threading.Thread(target=do_recalc, daemon=True).start()

    def dump_affected_to_excel(self, affected_base_skus, db):
        import os
        import datetime
    
        # Create export directory
        db_dir = self._get_database_folder()
        now = datetime.datetime.now()
        export_dir = os.path.join(
            db_dir,
            "Excel Exports",
            now.strftime("%m-%d-%y_%H-%M")
        )
        os.makedirs(export_dir, exist_ok=True)
    
        for base_sku in affected_base_skus:
            state = db.get(base_sku)
            if not state:
                continue
            # Prepare DataFrame for export
            in_tree_df = pd.DataFrame(state.get('in_tree_df', []))
            if in_tree_df.empty:
                continue
            filename = f"{base_sku}.xlsx"
            filepath = os.path.join(export_dir, filename)
            try:
                in_tree_df.to_excel(filepath, index=False)
            except Exception as e:
                print(f"Failed to export {base_sku}: {e}")
    
        messagebox.showinfo("Excel Export", f"Exported {len(affected_base_skus)} base SKUs to:\n{export_dir}")
            # Automatically open the export folder
        try:
            os.startfile(export_dir)
        except Exception as e:
            print(f"Could not open folder: {e}")

    def revert_to_backup(self):
        from tkinter import filedialog, messagebox
        # Let user pick a backup file
        db_path = filedialog.askopenfilename(
            title="Select Database Backup",
            filetypes=[("JSON Database", "*.json")],
            initialdir = get_data_dir()
        )
        if not db_path:
            return
        # Confirm with user
        if not messagebox.askyesno("Revert to Backup", f"Load and set this backup as the current database?\n\n{db_path}"):
            return
        # Copy selected file to be the latest database
        folder = self._get_database_folder()
        latest_path = os.path.join(folder, "sku_database_temp.json")
        import shutil
        try:
            shutil.copy2(db_path, latest_path)
            # Set as current in config (if you track a current path, update it here)
            self.current_prog_path = latest_path
            self.unsaved_changes = False
            self._update_title()
            messagebox.showinfo("Reverted", f"Reverted to backup:\n{db_path}")
            # Reload UI from the reverted database
            self.load_from_database()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to revert to backup:\n{e}")
    
    def reload_current_base_sku(self):
        """Reloads the currently selected base SKU from disk and updates all relevant DataFrames and UI."""
        db_path = self._get_latest_database_path()
        current_base_sku = self.base_sku_entry.get().strip() or self.base_sku
        if os.path.exists(db_path) and current_base_sku:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
            if current_base_sku in db:
                state = db[current_base_sku]
                self.master_df = pd.DataFrame(state.get('master_df', []))
                self.input_df = pd.DataFrame(state.get('input_df', []))
                self._fill_tree(self.out_tree, self.master_df)
                self.populate_value_grid()
                self._regenerate_left_preview()

    def rename_associated_sku(self):
        self._ctrl_s()
        import tkinter.simpledialog
    
        # 1. Prompt for old and new associated SKU (part number)
        cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
        if not os.path.exists(cost_db_path):
            messagebox.showinfo("No Cost DB", "No cost database file found.")
            return
        with open(cost_db_path, "r", encoding="utf-8") as f:
            cost_db = json.load(f)
        part_numbers = list(cost_db.keys())
        if not part_numbers:
            messagebox.showinfo("No Data", "No part numbers in cost database.")
            return
    
        old_part = tkinter.simpledialog.askstring(
            "Rename Associated SKU",
            "Enter the associated SKU (part number) to rename:",
            initialvalue=part_numbers[0],
            parent=self
        )
        if not old_part or old_part not in cost_db:
            messagebox.showerror("Not Found", "Associated SKU not found in cost database.")
            return
        new_part = tkinter.simpledialog.askstring(
            "Rename Associated SKU",
            f"Enter the new name for '{old_part}':",
            parent=self
        )
        if not new_part or new_part == old_part:
            return
    
        # 2. Rename key in cost_db
        cost_db[new_part] = cost_db.pop(old_part)
        with open(cost_db_path, "w", encoding="utf-8") as f:
            json.dump(cost_db, f, indent=2)
    
        # 3. Update all Associated SKUs in all base SKUs in sku_database (both sku and partnumber fields)
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            messagebox.showinfo("No SKU DB", "No SKU database file found.")
            return
        with open(db_path, "r", encoding="utf-8") as f:
            db = json.load(f)
        changed_any = False
        for bsku, bstate in db.items():
            master_df = pd.DataFrame(bstate.get('master_df', []))
            changed = False
            for idx, row in master_df.iterrows():
                assoc_str = row.get("Associated SKUs", "")
                sku_costs = parse_associated_skus(assoc_str)
                new_sku_costs = []
                for sku, cost, partnumber in sku_costs:
                    # Replace old_part with new_part in BOTH sku and partnumber fields
                    new_sku_val = new_part if sku == old_part else sku
                    new_part_val = new_part if partnumber == old_part else partnumber
                    if new_sku_val != sku or new_part_val != partnumber:
                        changed = True
                    new_sku_costs.append((new_sku_val, cost, new_part_val))
                if changed:
                    master_df.at[idx, "Associated SKUs"] = format_associated_skus(new_sku_costs)
            if changed:
                bstate['master_df'] = master_df.to_dict(orient='records')
                changed_any = True
    
        # 4. Save updated sku_database if changed
        if changed_any:
            with open(db_path, "w", encoding="utf-8") as f:
                json.dump(db, f, indent=2)
    
        # 5. Recalculate pricing for all base SKUs and update UI
        if changed_any:
            self.recalculate_all_pricing()
            self.reload_current_base_sku()
    
        messagebox.showinfo("Renamed", f"Renamed associated SKU '{old_part}' to '{new_part}' everywhere.")
    

def parse_associated_skus(assoc_str, default_cost='0'):
    # Returns list of (sku, cost, partnumber) tuples
    result = []
    for part in assoc_str.split(','):
        part = part.strip()
        if not part:
            continue
        fields = part.split(':')
        if len(fields) == 3:
            sku, cost, partnumber = fields
        elif len(fields) == 2:
            sku, cost = fields
            partnumber = ''
        else:
            sku = fields[0]
            cost = default_cost
            partnumber = ''
        result.append((sku.strip(), cost.strip(), partnumber.strip()))
    return result


def format_associated_skus(sku_cost_list):
    # sku_cost_list: list of (sku, cost, partnumber)
    return ', '.join(f"{sku}:{cost}:{partnumber}" if partnumber else f"{sku}:{cost}" for sku, cost, partnumber in sku_cost_list)

import subprocess
import requests
import time
import os
def ensure_web_service_running():
    global web_service_process
    try:
        requests.get('http://localhost:5000')
    except Exception:
        # Not running, so launch it
        service_path = os.path.join(os.path.dirname(__file__), "ap_sku_tool_web_service.py")
        web_service_process = subprocess.Popen(['python', service_path])
        # Wait for service to start
        for _ in range(10):
            try:
                requests.get('http://localhost:5000')
                break
            except Exception:
                time.sleep(1)

def pull_database_from_cloud():
    # Ensure web service is running
    ensure_web_service_running()
    # Download main database
    resp = requests.post('http://localhost:5000/pull_latest_db', json={
        'prefix': 'sku_database_',
        'suffix': '.json'
    })
    data = resp.json()
    if data.get('status') != 'success':
        messagebox.showerror("Database Download Error", data.get('message', 'Unknown error'))
        return False
    # Copy the downloaded file to the data subfolder as sku_database_temp.json
    data_dir = get_data_dir()
    os.makedirs(data_dir, exist_ok=True)
    src = data['local_path']
    dst = os.path.join(data_dir, "sku_database_temp.json")
    import shutil
    shutil.copy2(src, dst)

    # Download cost database
    resp = requests.post('http://localhost:5000/pull_latest_db', json={
        'prefix': 'cost_db',
        'suffix': '.json'
    })
    data = resp.json()
    if data.get('status') == 'success':
        src = data['local_path']
        dst = os.path.join(data_dir, "cost_db.json")
        shutil.copy2(src, dst)
    else:
        messagebox.showwarning(
            "Pricing Database Not Updated",
            "No pricing database was found in the cloud. The program will use the existing local cost_db.json (if any)."
        )

    return True

def push_database_to_cloud():
    import threading
    import itertools

    root = tk._default_root
    splash, status_labels = show_upload_status_dialog(root)

    def set_status(idx, value):
        status_labels[idx]['text'] = value
        status_labels[idx]['foreground'] = 'black'
        splash.update_idletasks()

    def spin_status(idx, stop_event):
        spinner_cycle = itertools.cycle(['-', '\\', '|', '/'])
        while not stop_event.is_set():
            set_status(idx, next(spinner_cycle)) 
            status_labels[idx].update_idletasks()
            stop_event.wait(0.15)

    def do_upload():
        try:
            ensure_web_service_running()
            data_dir = get_data_dir()
            dt = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

            # Step 1: Upload main database
            spin1 = threading.Event()
            t1 = threading.Thread(target=spin_status, args=(0, spin1), daemon=True)
            t1.start()
            local_path = os.path.join(data_dir, "sku_database_temp.json")
            drive_filename = f"sku_database_{dt}.json"
            resp = requests.post('http://localhost:5000/upload', json={
                'file_path': local_path,
                'drive_filename': drive_filename
            })
            data = resp.json()
            spin1.set()
            set_status(0, '✓')
            if data.get('status') != 'success':
                set_status(0, 'X')
                messagebox.showerror("Upload Error", data.get('message', 'Unknown error'))

            # Step 2: Upload cost database
            spin2 = threading.Event()
            t2 = threading.Thread(target=spin_status, args=(1, spin2), daemon=True)
            t2.start()
            cost_db_path = os.path.join(data_dir, "cost_db.json")
            if os.path.exists(cost_db_path):
                cost_drive_filename = f"cost_db_{dt}.json"
                resp = requests.post('http://localhost:5000/upload', json={
                    'file_path': cost_db_path,
                    'drive_filename': cost_drive_filename
                })
                data = resp.json()
                spin2.set()
                set_status(1, '✓' if data.get('status') == 'success' else 'X')
                if data.get('status') != 'success':
                    messagebox.showerror("Upload Error", f"Cost DB: {data.get('message', 'Unknown error')}")
            else:
                spin2.set()
                set_status(1, '✓')
            # Short delay for polish
            time.sleep(0.3)
        finally:
            splash.destroy()
            messagebox.showinfo("Upload Complete", "Upload process finished.")

    threading.Thread(target=do_upload, daemon=True).start()

def show_upload_status_dialog(root):
    splash = tk.Toplevel(root)
    splash.title("Uploading...")
    splash.geometry("400x160")
    splash.resizable(False, False)
    splash.grab_set()
    splash.overrideredirect(True)
    splash.update_idletasks()
    w = splash.winfo_screenwidth()
    h = splash.winfo_screenheight()
    size = tuple(int(_) for _ in splash.geometry().split('+')[0].split('x'))
    x = w//2 - size[0]//2
    y = h//2 - size[1]//2
    splash.geometry(f"{size[0]}x{size[1]}+{x}+{y}")

    # Bring to front and set focus
    splash.lift()
    splash.attributes('-topmost', True)
    splash.focus_force()
    splash.after(100, lambda: splash.attributes('-topmost', False))

    steps = [
        "Uploading Main Database",
        "Uploading Pricing Database"
    ]
    status_labels = []
    for i, step in enumerate(steps):
        frame = ttk.Frame(splash)
        frame.pack(fill='x', pady=5, padx=20)
        lbl = ttk.Label(frame, text=step, anchor='w', width=32)
        lbl.pack(side='left')
        status = ttk.Label(frame, text='X', width=2, anchor='center', font=('Segoe UI', 14), foreground='black')
        status.pack(side='left', padx=10)
        status_labels.append(status)
    return splash, status_labels

def try_repair_json_file(filepath):
    # Try to repair a corrupted JSON file by truncating at the last valid bracket/brace
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        # Find the last closing brace or bracket
        last_brace = content.rfind('}')
        last_bracket = content.rfind(']')
        cut = max(last_brace, last_bracket)
        if cut == -1:
            return None  # Can't repair
        repaired = content[:cut+1]
        # Try to parse the repaired content
        try:
            data = json.loads(repaired)
            # If successful, overwrite the file with the repaired content
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(repaired)
            messagebox.showinfo("Database Repaired", f"Corrupted database file was repaired and loaded.")
            return data
        except Exception:
            return None
    except Exception:
        return None
    
import itertools

def show_loading_dialog(root):
    splash = tk.Toplevel(root)
    splash.title("Loading...")
    splash.geometry("400x220")
    splash.resizable(False, False)
    splash.grab_set()
    splash.overrideredirect(True)
    splash.update_idletasks()
    w = splash.winfo_screenwidth()
    h = splash.winfo_screenheight()
    size = tuple(int(_) for _ in splash.geometry().split('+')[0].split('x'))
    x = w//2 - size[0]//2
    y = h//2 - size[1]//2
    splash.geometry(f"{size[0]}x{size[1]}+{x}+{y}")

    steps = [
        "Launching Web Service",
        "Syncing SKU Breakouts with Cloud",
        "Syncing Option Pricing with Cloud",
        "Purging Old Database Files from Cloud",  # <-- Add this line
        "Launching Main Window"
    ]
    status_labels = []
    for i, step in enumerate(steps):
        frame = ttk.Frame(splash)
        frame.pack(fill='x', pady=5, padx=20)
        lbl = ttk.Label(frame, text=step, anchor='w', width=32)
        lbl.pack(side='left')
        status = ttk.Label(frame, text='X', width=2, anchor='center', font=('Segoe UI', 14), foreground='black')
        status.pack(side='left', padx=10)
        status_labels.append(status)
    return splash, status_labels

def set_status(idx, value, status_labels):
    status_labels[idx]['text'] = value
    status_labels[idx]['foreground'] = 'black'  # Always black

import dateutil.parser

def purge_old_cloud_databases():
    # List all database files in Drive
    resp = requests.post('http://localhost:5000/list', json={
        'prefix': 'sku_database_',
        'suffix': '.json'
    })
    data = resp.json()
    if data.get('status') != 'success':
        return
    files = data.get('files', [])
    # Parse times and find the most recent
    for f in files:
        tstr = f.get('createdTime') or f.get('modifiedTime')
        f['dt'] = dateutil.parser.isoparse(tstr) if tstr else None
    files_with_dt = [f for f in files if f['dt'] is not None]
    if not files_with_dt:
        return
    # Find the most recent file
    newest = max(files_with_dt, key=lambda f: f['dt'])
    newest_time = newest['dt']
    cutoff = newest_time - datetime.timedelta(hours=36)
    # Delete files more than 36 hours older than the newest
    for f in files_with_dt:
        if f['dt'] < cutoff:
            requests.post('http://localhost:5000/delete', json={'file_id': f['id']})

if __name__ == '__main__':
    import threading

    root = tk.Tk()
    root.withdraw()
    splash, status_labels = show_loading_dialog(root)

    spinner_cycle = itertools.cycle(['-', '\\', '|', '/'])

    def set_status(idx, value):
        status_labels[idx]['text'] = value
        status_labels[idx]['foreground'] = 'black'  # Always black
        splash.update_idletasks()

    def spin_status(idx, stop_event, status_labels):
        spinner_cycle = itertools.cycle(['-', '\\', '|', '/'])
        while not stop_event.is_set():
            set_status(idx, next(spinner_cycle))
            status_labels[idx].update_idletasks()
            stop_event.wait(0.15)

    def start_app():
        # Step 1: Launching Web Service
        spin1 = threading.Event()
        t1 = threading.Thread(target=spin_status, args=(0, spin1, status_labels))  # <-- pass status_labels
        t1.start()
        ensure_web_service_running()
        spin1.set()
        set_status(0, '✓')

        # Step 2: Syncing SKU Breakouts with Cloud
        spin2 = threading.Event()
        t2 = threading.Thread(target=spin_status, args=(1, spin2, status_labels))  # <-- pass status_labels
        t2.start()
        # Download main database only
        resp = requests.post('http://localhost:5000/pull_latest_db', json={
            'prefix': 'sku_database_',
            'suffix': '.json'
        })
        data = resp.json()
        if data.get('status') == 'success':
            script_dir = os.path.dirname(os.path.abspath(__file__))
            data_dir = get_data_dir()
            os.makedirs(data_dir, exist_ok=True)
            src = data['local_path']
            dst = os.path.join(data_dir, "sku_database_temp.json")
            import shutil
            shutil.copy2(src, dst)
        spin2.set()
        set_status(1, '✓')

        # Step 3: Syncing Option Pricing with Cloud
        spin3 = threading.Event()
        t3 = threading.Thread(target=spin_status, args=(2, spin3, status_labels))  # <-- pass status_labels
        t3.start()
        # Download cost database
        resp = requests.post('http://localhost:5000/pull_latest_db', json={
            'prefix': 'cost_db',
            'suffix': '.json'
        })
        data = resp.json()
        if data.get('status') == 'success':
            src = data['local_path']
            dst = os.path.join(data_dir, "cost_db.json")
            shutil.copy2(src, dst)
        spin3.set()
        set_status(2, '✓')

        # Step 4: Purging Old Database Files from Cloud
        spin4 = threading.Event()
        t4 = threading.Thread(target=spin_status, args=(3, spin4, status_labels))
        t4.start()
        # purge_old_cloud_databases()
        spin4.set()
        set_status(3, '✓')

        # Step 5: Launching Main Window
        spin4 = threading.Event()
        t4 = threading.Thread(target=spin_status, args=(3, spin4, status_labels))  # <-- pass status_labels
        t4.start()
        # Simulate a short delay for UI polish
        import time
        time.sleep(0.3)
        spin4.set()
        set_status(3, '✓')

        # Now launch the main app
        app = OptionsParserApp()
        splash.destroy()
        app.mainloop()

    threading.Thread(target=start_app).start()
    root.mainloop()


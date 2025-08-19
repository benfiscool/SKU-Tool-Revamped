import re   #Add ctrl+s, ctrl+z, save before quit, autosave
import os
import sys
import pickle
import subprocess
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import datetime
import tkinter.font as tkFont
import threading
import time
import calendar as cal
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

# Test mode flag - if True, uses local files only and doesn't sync with cloud
isTest = False

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
        if not isTest:
            ensure_web_service_running()
        super().__init__()
        
        # Initialize save lock to prevent concurrent writes
        self._save_lock = threading.Lock()
        
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
        self.bind_all('<Control-r>', lambda e: self.restore_window_visibility())  # Ctrl+R to restore window
        # Show search on startup - but in non-test mode, wait a bit longer for cloud sync to complete
        if isTest:
            self.after(100, self.load_from_database)
        else:
            # In normal mode, wait longer to ensure cloud sync is complete
            self.after(2000, self.load_from_database)

        self.option_names = []  # Always define this attribute

    def restore_window_visibility(self):
        """Force the window to be visible and properly restored - but respect user's size choice"""
        try:
            # Only restore if the window is actually invisible or minimized
            current_state = self.state()
            is_visible = self.winfo_viewable()
            
            # Only intervene if window is minimized or invisible
            if current_state == 'iconic' or not is_visible:
                self.deiconify()  # Ensure not minimized
                self.lift()  # Bring to front
                self.focus_force()  # Give focus
                self.update()  # Force update
                
                # If still not visible after deiconify, try repositioning
                if self.winfo_viewable() == 0:
                    self.geometry("+100+100")  # Move to visible area
                    self.update()
            elif current_state in ['normal', 'zoomed']:
                # Window is in a good state, just bring to front if needed
                self.lift()
                self.focus_force()
        except Exception as e:
            print(f"Warning: Could not restore window visibility: {e}")

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

    def _ensure_base_sku_suffix(self, base_sku, is_suspension_superstore=None):
        """Ensure base SKU has appropriate suffix (SSS) or (MXT)"""
        # If already has a suffix, return as-is
        if base_sku.endswith(" (MXT)") or base_sku.endswith(" (SSS)"):
            return base_sku
            
        # If we don't know the store type, assume SSS for now
        # This method can be called when manually creating base SKUs
        if is_suspension_superstore is None:
            is_suspension_superstore = True  # Default to SSS for manual entries
            
        if is_suspension_superstore:
            return base_sku + " (SSS)"
        else:
            return base_sku + " (MXT)"

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
        if not isTest:
            push_database_to_cloud()
        self.unsaved_changes = False
        self._update_title()

    def _build_menu(self):
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Import CSV from BigCommerce...", command=self.load_csv)
        filemenu.add_command(label="Export Current SKU Breakout to Excel...", command=self.export_excel)
        filemenu.add_command(label="Batch Export SKUs to Excel...", command=self.batch_export_excel)
        filemenu.add_command(label="Bulk Update Base Prices...", command=self.bulk_update_base_prices)
        filemenu.add_separator()

        filemenu.add_command(label="Save Database Locally", command=self.save_to_database)
        filemenu.add_command(label="Save and Upload to Cloud", command=lambda:[self.save_to_database,self._ctrl_shift_s] if not isTest else self.save_to_database)
        filemenu.add_command(label="Revert to Backup Database", command=self.revert_to_backup)
        filemenu.add_command(label="Restore from Local Backup to Cloud", command=lambda: restore_from_local_backup())
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

        # --- Reports menu ---
        reportsmenu = tk.Menu(menubar, tearoff=0)
        reportsmenu.add_command(label="Base SKU Summary Report...", command=self.generate_base_sku_summary_report)
        reportsmenu.add_separator()
        reportsmenu.add_command(label="Associated SKU Usage Report...", command=self.generate_associated_sku_report)
        reportsmenu.add_separator()
        reportsmenu.add_command(label="Base SKU Price Change Report...", command=self.generate_base_price_change_report)
        reportsmenu.add_command(label="Associated SKU Price Change Report...", command=self.generate_associated_price_change_report)
        menubar.add_cascade(label="Reports", menu=reportsmenu)

        # --- Utilities menu ---
        utilitiesmenu = tk.Menu(menubar, tearoff=0)
        utilitiesmenu.add_command(label="Launch HotKeys...", command=self.launch_hotkeys)
        utilitiesmenu.add_command(label="Launch Excel SKU Reorder Tool...", command=self.launch_excel_sku_reorder)
        utilitiesmenu.add_command(label="Duplicate Remover...", command=self.launch_duplicate_remover)
        menubar.add_cascade(label="Utilities", menu=utilitiesmenu)

        self.config(menu=menubar)

    def _build_ui(self):
        ctrl = ttk.Frame(self)
        ctrl.pack(fill='x', padx=10, pady=5)

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
            else [self._update_base_sku_from_entry(), self.generate_new_skus(), self._mark_unsaved()]
        )
        # Also update base_sku when typing (without pressing Enter)
        self.base_sku_entry.bind(
            "<KeyRelease>",
            lambda e: None if (e.keysym.lower() == "s" and (e.state & 0x4)) or self._is_modifier(e)
            else self._update_base_sku_from_entry()
        )
        # And when focus leaves the field
        self.base_sku_entry.bind("<FocusOut>", lambda e: self._update_base_sku_from_entry())

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
        self.out_tree = self._make_treeview(bottom_frame, ['Name','Value',"Add'l Price","Add'l Weight"], stretch=True, label="Option Master Table")
        self.out_tree.master.grid(row=0, column=0, sticky='nsew')
        self.out_tree.bind("<Double-1>", self._on_out_tree_double_click)
        self.out_tree.bind("<<TreeviewSelect>>", self._on_out_tree_select)

        # Bottom right: breakdown panel
        self.breakdown_tree = self._make_treeview(bottom_frame, ['Name', 'Value', 'Price ($)', 'Weight (lb)'], stretch=True, label="Cost/Weight Breakdown")
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
        ttk.Button(row, text="Copy Options From...", command=self.copy_options_from_base_sku).grid(row=0, column=4, padx=(10,0))

        # Header row for value/cost/weight, aligned with grid_container columns
        self.header_frame = ttk.Frame(cfg)
        self.header_frame.pack(fill='x', padx=5, pady=(0,0))
        ttk.Label(self.header_frame, text="Value", anchor='w', width=40).grid(row=0, column=0, padx=0, sticky='w')
        ttk.Label(self.header_frame, text="Add'l Price ($)", anchor='w', width=15).grid(row=0, column=1, padx=5, sticky='w')
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

        # Add mouse wheel scrolling support
        def _on_mousewheel(event):
            self.grid_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def _bind_mousewheel(event):
            self.grid_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_mousewheel(event):
            self.grid_canvas.unbind_all("<MouseWheel>")
        
        # Bind mouse wheel events when mouse enters the configure panel area
        cfg.bind("<Enter>", _bind_mousewheel)
        cfg.bind("<Leave>", _unbind_mousewheel)
        self.grid_canvas.bind("<Enter>", _bind_mousewheel)
        self.grid_canvas.bind("<Leave>", _unbind_mousewheel)

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
                suffix = " (SSS)"
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
        
        # Ensure current data is saved before updating export time
        self.save_to_database(temp=True)
        self._update_last_export_time(base_sku)
        self._show_last_export_time(base_sku)

    def batch_export_excel(self):
        """Export multiple SKUs to Excel from a list"""
        # Create dialog for SKU input
        dialog = tk.Toplevel(self)
        dialog.title("Batch Export SKUs to Excel")
        dialog.geometry("600x500")
        dialog.transient(self)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Instructions
        tk.Label(dialog, text="Batch Export SKUs to Excel", font=('Segoe UI', 12, 'bold')).pack(pady=10)
        tk.Label(dialog, text="Enter base SKUs to export (one per line):", font=('Segoe UI', 10)).pack(pady=(0,5))
        tk.Label(dialog, text="SKUs can include or exclude suffixes like (SSS) or (MXT)", font=('Segoe UI', 8), foreground='gray').pack(pady=(0,10))
        
        # Text area for SKU input
        text_frame = tk.Frame(dialog)
        text_frame.pack(fill='both', expand=True, padx=10, pady=(0,10))
        
        sku_text = tk.Text(text_frame, wrap='word', height=15)
        sku_text.pack(side='left', fill='both', expand=True)
        
        scrollbar = tk.Scrollbar(text_frame, orient='vertical', command=sku_text.yview)
        scrollbar.pack(side='right', fill='y')
        sku_text.configure(yscrollcommand=scrollbar.set)
        
        # Options frame
        options_frame = tk.Frame(dialog)
        options_frame.pack(fill='x', padx=10, pady=(0,10))
        
        # Suffix handling options
        tk.Label(options_frame, text="For SKUs with variants (SSS/MXT), export:", font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(0,5))
        
        suffix_var = tk.StringVar(value="both")
        tk.Radiobutton(options_frame, text="Both variants", variable=suffix_var, value="both").pack(anchor='w', padx=(20,0))
        tk.Radiobutton(options_frame, text="Only (SSS) variants", variable=suffix_var, value="sss").pack(anchor='w', padx=(20,0))
        tk.Radiobutton(options_frame, text="Only (MXT) variants", variable=suffix_var, value="mxt").pack(anchor='w', padx=(20,0))
        
        # Progress frame (initially hidden)
        progress_frame = tk.Frame(dialog)
        progress_label = tk.Label(progress_frame, text="")
        progress_label.pack()
        
        # Buttons
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill='x', padx=10, pady=(0,10))
        
        def export_batch():
            # Get SKU list
            sku_input = sku_text.get("1.0", tk.END).strip()
            if not sku_input:
                messagebox.showwarning("No SKUs", "Please enter at least one SKU to export.")
                return
            
            sku_list = [sku.strip() for sku in sku_input.split('\n') if sku.strip()]
            if not sku_list:
                messagebox.showwarning("No SKUs", "Please enter at least one valid SKU to export.")
                return
            
            # Show progress
            progress_frame.pack(fill='x', padx=10, pady=(0,5))
            progress_label.config(text="Processing SKUs...")
            dialog.update()
            
            try:
                # Process batch export with separate files (no folder selection needed)
                exported_files = self._process_batch_export_separate_files(
                    sku_list, 
                    suffix_var.get(), 
                    progress_label,
                    dialog
                )
                
                dialog.destroy()
                
                if exported_files:
                    message = f"Successfully exported {len(exported_files)} SKU files:\n\n"
                    for file_path in exported_files[:5]:  # Show first 5 files
                        message += f"• {os.path.basename(file_path)}\n"
                    if len(exported_files) > 5:
                        message += f"• ... and {len(exported_files) - 5} more files\n"
                    
                    message += f"\nFiles saved to:\n{exported_files[0].split(os.path.basename(exported_files[0]))[0]}"
                    
                    if messagebox.askyesno("Export Complete", f"{message}\n\nOpen export folder?"):
                        try:
                            subprocess.Popen(['explorer', os.path.dirname(exported_files[0])], shell=True)
                        except Exception:
                            messagebox.showerror("Error","Can't open folder.")
                else:
                    messagebox.showwarning("No Data", "No valid SKU data found to export.")
                    
            except Exception as e:
                progress_frame.pack_forget()
                messagebox.showerror("Export Error", f"Error during batch export:\n{e}")
        
        def cancel_export():
            dialog.destroy()
        
        tk.Button(btn_frame, text="Export to Excel", command=export_batch).pack(side='right', padx=(5,0))
        tk.Button(btn_frame, text="Cancel", command=cancel_export).pack(side='right')
        
        # Focus on text area
        sku_text.focus_set()

    def _process_batch_export(self, sku_list, ignore_suffix, update_pricing, progress_label, dialog):
        """Process the batch export and return combined DataFrame"""
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            messagebox.showerror("Database Error", "Database file not found.")
            return None
        
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except Exception as e:
            messagebox.showerror("Database Error", f"Error reading database: {e}")
            return None
        
        # Update pricing if requested
        if update_pricing:
            try:
                self.update_option_costs_from_cost_db()
                # Reload database to get updated pricing
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
            except Exception as e:
                print(f"Error updating pricing: {e}")
        
        combined_rows = []
        processed_count = 0
        
        for sku in sku_list:
            processed_count += 1
            progress_label.config(text=f"Processing {processed_count}/{len(sku_list)}: {sku}")
            dialog.update()
            
            # Find matching base SKU in database
            matching_key = self._find_matching_base_sku(sku, db.keys(), ignore_suffix)
            
            if matching_key:
                try:
                    state = db[matching_key]
                    
                    # Recreate the DataFrame for this SKU
                    input_df = pd.DataFrame(state['input_df'])
                    master_df = pd.DataFrame(state['master_df'])
                    
                    if input_df.empty or master_df.empty:
                        print(f"Skipping {sku}: empty data")
                        continue
                    
                    # Get base price and weight
                    try:
                        base_price = float(state.get('base_price', 0))
                        base_weight = float(state.get('base_weight', 0))
                    except (ValueError, TypeError):
                        print(f"Skipping {sku}: invalid base price/weight")
                        continue
                    
                    # Generate SKU combinations (similar to _regenerate_left_preview)
                    sku_rows = self._generate_sku_combinations(
                        input_df, master_df, base_price, base_weight, matching_key, state
                    )
                    
                    # Add base SKU identifier to each row
                    for row in sku_rows:
                        row['Base_SKU'] = matching_key
                    
                    combined_rows.extend(sku_rows)
                    
                except Exception as e:
                    print(f"Error processing {sku}: {e}")
                    continue
            else:
                print(f"SKU not found: {sku}")
        
        if combined_rows:
            result_df = pd.DataFrame(combined_rows)
            # Reorder columns to put Base_SKU first
            if 'Base_SKU' in result_df.columns:
                cols = ['Base_SKU'] + [col for col in result_df.columns if col != 'Base_SKU']
                result_df = result_df[cols]
            return result_df
        else:
            return pd.DataFrame()

    def _find_matching_base_sku(self, input_sku, available_skus, ignore_suffix):
        """Find matching base SKU, optionally ignoring suffixes"""
        input_sku = input_sku.strip()
        
        # First try exact match
        if input_sku in available_skus:
            return input_sku
        
        if ignore_suffix:
            # Remove suffix from input
            input_clean = re.sub(r'\s*\([^)]*\)\s*$', '', input_sku).strip()
            
            # Try to find match by comparing cleaned versions
            for available_sku in available_skus:
                available_clean = re.sub(r'\s*\([^)]*\)\s*$', '', available_sku).strip()
                if input_clean.lower() == available_clean.lower():
                    return available_sku
        
        return None

    def _generate_sku_combinations(self, input_df, master_df, base_price, base_weight, base_sku, state):
        """Generate SKU combinations for export (similar to _regenerate_left_preview logic)"""
        rows = []
        
        # Get option names
        option_names = sorted({
            name for blob in input_df['Options']
            for name, _ in VALUE_RE.findall(str(blob))
        })
        
        # Process each row in input_df
        for idx, row in enumerate(input_df.itertuples(), start=1):
            sku = row.SKU
            options = list(VALUE_RE.findall(str(row.Options)))
            
            # Calculate additional cost and weight
            total_add_cost = 0
            total_add_weight = 0
            
            for option_name, option_value in options:
                # Find matching row in master_df
                match = master_df[(master_df['Name'] == option_name) & (master_df['Value'] == option_value)]
                if not match.empty:
                    try:
                        add_cost = float(match.iloc[0]["Add'l Cost"] or 0)
                        add_weight = float(match.iloc[0]["Add'l Weight"] or 0)
                        total_add_cost += add_cost
                        total_add_weight += add_weight
                    except (ValueError, TypeError):
                        pass
            
            # Create row data
            row_data = {
                '#': idx,
                'SKU': str(sku),
                'New SKU': '',  # Will be filled if available
                'Price': round(float(base_price) + float(total_add_cost), 2),
                'Weight': round(float(base_weight) + float(total_add_weight), 2)
            }
            
            # Add option columns
            for option_name in option_names:
                row_data[str(option_name)] = ''
            
            # Fill option values
            for option_name, option_value in options:
                row_data[str(option_name)] = str(option_value)
            
            # Try to get New SKU from existing data if available
            in_tree_df = pd.DataFrame(state.get('in_tree_df', []))
            if not in_tree_df.empty and idx <= len(in_tree_df):
                try:
                    new_sku = in_tree_df.iloc[idx-1].get('New SKU', '')
                    if new_sku:
                        row_data['New SKU'] = str(new_sku)
                except:
                    pass
            
            rows.append(row_data)
        
        return rows

    def _process_batch_export_separate_files(self, sku_list, suffix_filter, progress_label, dialog):
        """Process batch export for separate files per base SKU"""
        try:
            # Load the database
            db_path = self._get_latest_database_path()
            if not os.path.exists(db_path):
                messagebox.showerror("Database Error", "Database file not found.", parent=dialog)
                return []
            
            with open(db_path, "r", encoding="utf-8") as f:
                sku_database = json.load(f)
            
            # Create timestamped export folder in Excel Exports
            project_root = os.path.dirname(os.path.abspath(__file__))
            excel_exports_dir = os.path.join(project_root, "Excel Exports")
            os.makedirs(excel_exports_dir, exist_ok=True)
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_folder = os.path.join(excel_exports_dir, f"BatchExport_{timestamp}")
            os.makedirs(output_folder, exist_ok=True)
            
            # Group SKUs by base SKU and handle variants
            base_sku_groups = {}
            
            for sku in sku_list:
                sku = sku.strip()
                
                # Find all matching SKUs considering variants
                matching_skus = self._find_matching_skus_with_variants(sku, suffix_filter, sku_database)
                
                for matched_sku in matching_skus:
                    # For separate files, each SKU gets its own group
                    if matched_sku not in base_sku_groups:
                        base_sku_groups[matched_sku] = []
                    base_sku_groups[matched_sku].append(matched_sku)
            
            total_base_skus = len(base_sku_groups)
            exported_files = []
            
            # Process each base SKU group
            for index, (sku_name, skus_in_group) in enumerate(base_sku_groups.items(), 1):
                progress_label.config(text=f"Processing {str(sku_name)} ({index}/{total_base_skus})...")
                dialog.update()
                
                try:
                    # Each group should only contain one SKU since we want separate files
                    result_df = self._process_single_base_sku_export(sku_name, skus_in_group, sku_database)
                    
                    if not result_df.empty:
                        # Sanitize filename using the exact SKU name
                        safe_sku_name = str(sku_name).replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                        filename = f"{safe_sku_name}.xlsx"
                        
                        filepath = os.path.join(output_folder, filename)
                        result_df.to_excel(filepath, index=False)
                        exported_files.append(filepath)
                        print(f"Exported {str(sku_name)} to {filepath}")
                    
                except Exception as e:
                    print(f"Error processing {str(sku_name)}: {str(e)}")
                    continue
            
            # Update progress
            progress_label.config(text="Export completed!")
            dialog.update()
            
            # Show completion message
            if exported_files:
                result = messagebox.askyesno(
                    "Export Complete", 
                    f"Successfully exported {len(exported_files)} files to:\n{output_folder}\n\nOpen the output folder?",
                    parent=dialog
                )
                
                if result:
                    # Open the output folder
                    try:
                        os.startfile(output_folder)
                    except:
                        # Fallback for other OS
                        import subprocess
                        subprocess.run(['explorer', output_folder], shell=True)
            else:
                messagebox.showwarning("Export Complete", "No files were exported.", parent=dialog)
            
            return exported_files
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Error during batch export: {str(e)}", parent=dialog)
            return []

    def _find_matching_skus_with_variants(self, input_sku, suffix_filter, sku_database):
        """Find matching SKUs considering variant preferences"""
        input_sku = input_sku.strip()
        available_skus = list(sku_database.keys())
        matching_skus = []
        
        # Remove suffix from input for base matching
        input_base = re.sub(r'\s*\([^)]*\)\s*$', '', input_sku).strip()
        
        # Find all SKUs that match the base name
        potential_matches = []
        for available_sku in available_skus:
            available_base = re.sub(r'\s*\([^)]*\)\s*$', '', available_sku).strip()
            if input_base.lower() == available_base.lower():
                potential_matches.append(available_sku)
        
        # If no base matches found, try exact match
        if not potential_matches and input_sku in available_skus:
            potential_matches = [input_sku]
        
        # Filter based on suffix preference
        if suffix_filter == "both":
            matching_skus = potential_matches
        elif suffix_filter == "sss":
            matching_skus = [sku for sku in potential_matches if "(SSS)" in sku.upper()]
        elif suffix_filter == "mxt":
            matching_skus = [sku for sku in potential_matches if "(MXT)" in sku.upper()]
        else:  # This handles the old ignore_suffix boolean for backwards compatibility
            matching_skus = potential_matches
        
        return matching_skus

    def _process_single_base_sku_export(self, base_sku, related_skus, sku_database):
        """Process export for a single base SKU and its related SKUs"""
        # Get base SKU data
        base_data = sku_database.get(base_sku, {})
        if not base_data:
            return pd.DataFrame()
        
        # Get the required data for processing
        input_df = pd.DataFrame(base_data.get('input_df', []))
        master_df = pd.DataFrame(base_data.get('master_df', []))
        
        # Ensure base_price and base_weight are properly converted to float
        try:
            base_price = float(base_data.get('base_price', 0))
        except (ValueError, TypeError):
            base_price = 0.0
            
        try:
            base_weight = float(base_data.get('base_weight', 0))
        except (ValueError, TypeError):
            base_weight = 0.0
        
        if input_df.empty:
            return pd.DataFrame()
        
        # Generate the SKU combinations
        state = {'in_tree_df': base_data.get('in_tree_df', [])}
        rows = self._generate_sku_combinations(input_df, master_df, base_price, base_weight, base_sku, state)
        
        # Convert to DataFrame and apply column ordering
        if rows:
            result_df = pd.DataFrame(rows)
            
            # Apply standard column ordering
            base_columns = ['#', 'SKU', 'New SKU', 'Price', 'Weight']
            option_columns = [col for col in result_df.columns if col not in base_columns]
            cols = base_columns + sorted(option_columns)
            cols = [col for col in cols if col in result_df.columns]
            
            if cols:
                result_df = result_df[cols]
            return result_df
        else:
            return pd.DataFrame()

    def _safe_write_json(self, data, file_path):
        """Safely write JSON data to file using atomic operations to prevent corruption."""
        import tempfile
        import shutil
        
        # Create a temporary file in the same directory
        temp_dir = os.path.dirname(file_path)
        try:
            # Write to temporary file first
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', 
                                           dir=temp_dir, delete=False, 
                                           suffix='.tmp') as temp_file:
                json.dump(data, temp_file, indent=2)
                temp_file.flush()
                os.fsync(temp_file.fileno())  # Force write to disk
                temp_path = temp_file.name
            
            # Atomically replace the original file
            if os.path.exists(file_path):
                # Create backup before replacing
                backup_path = file_path + '.backup'
                shutil.copy2(file_path, backup_path)
            
            # Atomic move (this is the key to preventing corruption)
            shutil.move(temp_path, file_path)
            
            # Clean up backup if successful
            backup_path = file_path + '.backup'
            if os.path.exists(backup_path):
                os.remove(backup_path)
                
            return True
            
        except Exception as e:
            print(f"Error in safe JSON write: {e}")
            # Clean up temp file if it exists
            if 'temp_path' in locals() and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass
            return False

    def save_to_database(self, db_path=None, temp=False):
        # Use lock to prevent concurrent saves that cause corruption
        with self._save_lock:
            return self._do_save_to_database(db_path, temp)
    
    def _do_save_to_database(self, db_path=None, temp=False):
        folder = self._get_database_folder()
        if temp:
            # Use unique temp filename to prevent conflicts
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            db_path = os.path.join(folder, f"sku_database_temp_{timestamp}.json")
        else:
            dt = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            db_path = os.path.join(folder, f"sku_database_{dt}.json")
        
        base_sku = self.base_sku_entry.get().strip()
        if not base_sku:
            # If the entry is empty, try the stored value as fallback
            base_sku = self.base_sku.strip() if hasattr(self, 'base_sku') and self.base_sku else ''
            
        # Update the stored base_sku to match what we're saving
        self.base_sku = base_sku
            
        if os.path.exists(self._get_latest_database_path()):
            try:
                with open(self._get_latest_database_path(), "r", encoding="utf-8") as f:
                    db = json.load(f)
            except (json.JSONDecodeError, UnicodeDecodeError) as e:
                print(f"Database corruption detected: {e}")
                messagebox.showwarning("Database Corruption", 
                    "Database file is corrupted. Attempting to repair...\n"
                    "This may take a moment.")
                
                # Try to repair the database
                repaired_db = self.try_repair_json_file(self._get_latest_database_path())
                if repaired_db is not None:
                    db = repaired_db
                    messagebox.showinfo("Database Repaired", 
                        "Database has been successfully repaired and loaded.\n"
                        "Your data should be intact.")
                else:
                    messagebox.showerror("Repair Failed", 
                        "Could not repair the database file.\n"
                        "Using empty database to prevent data loss.\n"
                        "Please check your data folder for backup files.")
                    db = {}
        else:
            db = {}
        in_tree_df = self._get_tree_df(self.in_tree)
        
        # DEBUG: Check what data we're putting into the state
        input_df_records = self.input_df.to_dict(orient='records') if self.input_df is not None else []
        master_df_records = self.master_df.to_dict(orient='records') if self.master_df is not None else []
        in_tree_df_records = in_tree_df.to_dict(orient='records') if in_tree_df is not None else []
        
        print(f"DEBUG: input_df_records count: {len(input_df_records)}")
        print(f"DEBUG: master_df_records count: {len(master_df_records)}")
        print(f"DEBUG: in_tree_df_records count: {len(in_tree_df_records)}")
        
        state = {
            'input_df': input_df_records,
            'master_df': master_df_records,
            'in_tree_df': in_tree_df_records,
            'base_price': self.base_price_entry.get(),
            'base_weight': self.base_weight_entry.get(),
            'base_sku': base_sku,
            'prefix': self.prefix_entry.get().strip()  # Save the prefix
        }
        # Preserve last_export if present
        if base_sku in db and 'last_export' in db[base_sku]:
            state['last_export'] = db[base_sku]['last_export']
        db[base_sku] = state
        
        # Use safe atomic write to prevent corruption
        if not self._safe_write_json(db, db_path):
            messagebox.showerror("Save Error", 
                "Failed to save database safely. Please try again.\n"
                "Your changes are preserved in memory.")
            return False
        if not temp:
            # Clean up old temp files (keep last 5 for safety)
            try:
                temp_files = [f for f in os.listdir(folder) if f.startswith("sku_database_temp_")]
                temp_files.sort()
                if len(temp_files) > 5:
                    for old_temp in temp_files[:-5]:
                        old_temp_path = os.path.join(folder, old_temp)
                        if os.path.exists(old_temp_path):
                            os.remove(old_temp_path)
            except Exception as e:
                print(f"Warning: Could not clean up temp files: {e}")
                
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
                # Database could not be repaired, offer backup restoration
                result = messagebox.askyesnocancel(
                    "Database Corrupted", 
                    "The database file could not be repaired.\n\n"
                    "Would you like to restore from your latest uncorrupted local backup?\n\n"
                    "• YES: Restore from local backup to cloud\n"
                    "• NO: Continue with empty database (not recommended)\n"
                    "• CANCEL: Abort operation"
                )
                
                if result is None:  # Cancel
                    return
                elif result:  # Yes - restore from backup
                    if restore_from_local_backup():
                        # After successful restore, try loading from the latest restored file
                        try:
                            # Get the fresh database path after restore
                            fresh_db_path = self._get_latest_database_path()
                            with open(fresh_db_path, "r", encoding="utf-8") as f:
                                db = json.load(f)
                        except:
                            db = {}
                    else:
                        db = {}
                else:  # No - continue with empty database
                    messagebox.showwarning("Using Empty Database", 
                        "Proceeding with empty database. "
                        "Your data may be lost. "
                        "Consider restoring from backup soon.")
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
                
                # Add KeyRelease binding as backup
                def on_key_release(event):
                    self.update_list(entry)
                
                entry.bind('<KeyRelease>', on_key_release)
                
                # Store reference to entry for update_list
                self.search_entry = entry
                
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
                
            def update_list(self, entry_widget=None, *args):
                if entry_widget and hasattr(entry_widget, 'get'):
                    search_text = entry_widget.get().lower()
                else:
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

        # --- Correct DataFrame loading for orient='records' with fallback for missing keys ---
        self.input_df = pd.DataFrame(state.get('input_df', []))
        self.master_df = pd.DataFrame(state.get('master_df', []))
        in_tree_df = pd.DataFrame(state.get('in_tree_df', []))

        # Restore option names for UI logic
        if not self.input_df.empty and 'Options' in self.input_df.columns:
            self.option_names = sorted({
                name for blob in self.input_df['Options']
                for name, _ in VALUE_RE.findall(str(blob))
            })
        else:
            self.option_names = []

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
        if not self.master_df.empty and 'Name' in self.master_df.columns:
            names = sorted(self.master_df['Name'].unique())
        else:
            names = []
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
            if not item:
                return
            
            if col == "#1":  # Part Number column
                # Handle part number renaming
                handle_part_number_rename(item)
            elif col == "#2":  # Price column
                # Handle price editing (existing functionality)
                handle_price_edit(item, event)

        def handle_part_number_rename(item):
            old_part_number = tree.set(item, "Part Number")
            
            # Create a simple dialog to get the new part number
            dialog = tk.Toplevel(win)
            dialog.title("Rename Associated SKU")
            dialog.geometry("400x200")
            dialog.transient(win)
            dialog.grab_set()
            
            # Center the dialog
            dialog.update_idletasks()
            x = win.winfo_x() + (win.winfo_width() // 2) - (dialog.winfo_width() // 2)
            y = win.winfo_y() + (win.winfo_height() // 2) - (dialog.winfo_height() // 2)
            dialog.geometry(f"+{x}+{y}")
            
            tk.Label(dialog, text=f"Rename part number:", font=('Segoe UI', 10, 'bold')).pack(pady=10)
            tk.Label(dialog, text=f"Current: {old_part_number}").pack(pady=5)
            tk.Label(dialog, text="New name:").pack(pady=(10,0))
            
            new_name_entry = tk.Entry(dialog, width=30)
            new_name_entry.pack(pady=5)
            new_name_entry.insert(0, old_part_number)
            new_name_entry.select_range(0, tk.END)
            new_name_entry.focus_set()
            
            result = {"confirmed": False, "new_name": ""}
            
            def confirm_rename():
                new_name = new_name_entry.get().strip()
                if not new_name:
                    messagebox.showwarning("Invalid Name", "Part number cannot be empty.")
                    return
                if new_name == old_part_number:
                    dialog.destroy()
                    return
                
                # Check if new name already exists
                if new_name in cost_db:
                    if not messagebox.askyesno("Name Exists", 
                        f"Part number '{new_name}' already exists in the database.\n"
                        "Do you want to merge with the existing entry?"):
                        return
                
                result["confirmed"] = True
                result["new_name"] = new_name
                dialog.destroy()
            
            def cancel_rename():
                dialog.destroy()
            
            btn_frame = tk.Frame(dialog)
            btn_frame.pack(pady=20)
            tk.Button(btn_frame, text="Rename", command=confirm_rename).pack(side='left', padx=5)
            tk.Button(btn_frame, text="Cancel", command=cancel_rename).pack(side='left', padx=5)
            
            # Bind Enter key to confirm
            new_name_entry.bind('<Return>', lambda e: confirm_rename())
            
            # Wait for dialog to close
            dialog.wait_window()
            
            if result["confirmed"]:
                perform_sku_rename(old_part_number, result["new_name"])

        def save_cost_database():
            """Save the current cost database to file"""
            try:
                with open(db_path, "w", encoding="utf-8") as f:
                    json.dump(cost_db, f, indent=2)
            except Exception as e:
                print(f"Error saving cost database: {e}")

        def perform_sku_rename(old_name, new_name):
            # Show confirmation dialog with details of what will be updated
            affected_instances = find_sku_usage(old_name)
            
            if not affected_instances:
                messagebox.showinfo("No Usage Found", 
                    f"Part number '{old_name}' is not currently used in any associated SKU fields.")
                # Still update the cost database
                update_cost_database_name(old_name, new_name)
                
                # Save the updated cost database
                save_cost_database()
                
                # Recalculate current pricing in case the user is working on related data
                try:
                    self.update_option_costs_from_cost_db()
                    self.recalculate_all_pricing()
                    self._regenerate_left_preview()
                except:
                    pass  # Ignore errors if no data is loaded
                
                populate_tree(search_var.get())
                return
            
            message = f"This will rename '{old_name}' to '{new_name}' and update:\n\n"
            for base_sku, options in affected_instances.items():
                message += f"Base SKU: {base_sku}\n"
                for option in options:
                    message += f"  • {option['Name']}: {option['Value']}\n"
                message += "\n"
            
            message += f"Total: {sum(len(options) for options in affected_instances.values())} option(s) across {len(affected_instances)} base SKU(s).\n\n"
            message += "Do you want to proceed with the rename?"
            
            if messagebox.askyesno("Confirm Rename", message):
                # Perform the actual rename
                affected_base_skus = list(affected_instances.keys())
                update_all_sku_references(old_name, new_name)
                update_cost_database_name(old_name, new_name)
                
                # Save the updated cost database immediately
                save_cost_database()
                
                # Use the main application's recalculate method for better reliability
                try:
                    self.update_option_costs_from_cost_db()
                    self.recalculate_all_pricing()
                except Exception as e:
                    print(f"Error during recalculation: {e}")
                
                # Also recalculate for specific base SKUs
                recalculate_pricing_for_base_skus(affected_base_skus)
                
                populate_tree(search_var.get())
                messagebox.showinfo("Rename Complete", 
                    f"Successfully renamed '{old_name}' to '{new_name}' across all instances.\n"
                    f"Pricing has been recalculated for {len(affected_base_skus)} affected base SKU(s).\n"
                    f"Original SKU '{old_name}' has been removed from pricing database.")

        def find_sku_usage(sku_name):
            """Find all instances where a SKU is used in associated SKU fields"""
            affected_instances = {}
            
            # Get the latest database
            db_path = self._get_latest_database_path()
            if not os.path.exists(db_path):
                return affected_instances
            
            try:
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
                
                for base_sku, base_data in db.items():
                    master_df_data = base_data.get('master_df', [])
                    affected_options = []
                    
                    for row in master_df_data:
                        assoc_str = row.get("Associated SKUs", "")
                        if assoc_str and sku_name in assoc_str:
                            # Parse and check if it's actually this SKU
                            sku_costs = parse_associated_skus(assoc_str)
                            for sku, cost, partnumber in sku_costs:
                                if sku == sku_name:
                                    affected_options.append({
                                        'Name': row.get('Name', ''),
                                        'Value': row.get('Value', ''),
                                        'Associated SKUs': assoc_str
                                    })
                                    break
                    
                    if affected_options:
                        affected_instances[base_sku] = affected_options
                        
            except Exception as e:
                print(f"Error finding SKU usage: {e}")
                
            return affected_instances

        def update_all_sku_references(old_name, new_name):
            """Update all references to the old SKU name with the new name"""
            # Get the latest database
            db_path = self._get_latest_database_path()
            if not os.path.exists(db_path):
                return
            
            try:
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
                
                modified = False
                for base_sku, base_data in db.items():
                    master_df_data = base_data.get('master_df', [])
                    
                    for row in master_df_data:
                        assoc_str = row.get("Associated SKUs", "")
                        if assoc_str and old_name in assoc_str:
                            # Parse, update, and rebuild the string
                            sku_costs = parse_associated_skus(assoc_str)
                            updated_parts = []
                            
                            for sku, cost, partnumber in sku_costs:
                                if sku == old_name:
                                    sku = new_name
                                    modified = True
                                
                                # Rebuild the part string
                                if partnumber:
                                    updated_parts.append(f"{sku}:{cost}:{partnumber}")
                                elif cost != '0':
                                    updated_parts.append(f"{sku}:{cost}")
                                else:
                                    updated_parts.append(sku)
                            
                            row["Associated SKUs"] = ", ".join(updated_parts)
                
                if modified:
                    # Save the updated database
                    with open(db_path, "w", encoding="utf-8") as f:
                        json.dump(db, f, indent=2)
                    
                    # Update the current loaded data if it matches
                    if hasattr(self, 'master_df') and self.master_df is not None:
                        current_base_sku = self.base_sku_entry.get().strip() or self.base_sku
                        if current_base_sku in db:
                            # Reload the current data
                            self.master_df = pd.DataFrame(db[current_base_sku]['master_df'])
                            self._fill_tree(self.out_tree, self.master_df)
                            
                            # Update the configure panel if visible
                            if hasattr(self, 'name_combo') and self.name_combo.get():
                                self.populate_value_grid()
                                
            except Exception as e:
                print(f"Error updating SKU references: {e}")
                messagebox.showerror("Update Error", f"Error updating SKU references: {e}")

        def update_cost_database_name(old_name, new_name):
            """Update the cost database to rename the SKU"""
            if old_name in cost_db:
                if new_name not in cost_db:
                    # Simple rename - move the price to the new name
                    cost_db[new_name] = cost_db[old_name]
                else:
                    # Merging - keep the existing price for new_name
                    print(f"Merging '{old_name}' into existing '{new_name}', keeping existing price")
                
                # Always delete the old entry
                del cost_db[old_name]

        def handle_price_edit(item, event):
            # Existing price editing functionality
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
                # Auto-save when Return key is pressed
                save_all()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", lambda e: entry.destroy())

        def recalculate_pricing_for_base_skus(base_sku_list):
            """Recalculate pricing for specific base SKUs after SKU rename"""
            try:
                # Get the latest database
                db_path = self._get_latest_database_path()
                if not os.path.exists(db_path):
                    return
                
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
                
                modified = False
                for base_sku in base_sku_list:
                    if base_sku not in db:
                        continue
                        
                    base_data = db[base_sku]
                    master_df_data = base_data.get('master_df', [])
                    
                    # Update costs for each option based on current cost database
                    for row in master_df_data:
                        assoc_str = row.get("Associated SKUs", "")
                        if assoc_str:
                            # Parse associated SKUs and recalculate total cost
                            sku_costs = parse_associated_skus(assoc_str)
                            total_cost = 0
                            
                            for sku, old_cost, partnumber in sku_costs:
                                # Get current price from cost database
                                current_price = cost_db.get(sku, 0)
                                try:
                                    total_cost += float(current_price)
                                except (ValueError, TypeError):
                                    total_cost += 0
                            
                            # Update the Add'l Cost field
                            old_cost = row.get("Add'l Cost", "0")
                            if str(total_cost) != str(old_cost):
                                row["Add'l Cost"] = str(total_cost)
                                modified = True
                
                if modified:
                    # Save the updated database
                    with open(db_path, "w", encoding="utf-8") as f:
                        json.dump(db, f, indent=2)
                    
                    # Update the current loaded data if it's one of the affected base SKUs
                    current_base_sku = self.base_sku_entry.get().strip() or self.base_sku
                    if current_base_sku in base_sku_list and current_base_sku in db:
                        # Reload the current data
                        self.master_df = pd.DataFrame(db[current_base_sku]['master_df'])
                        self._fill_tree(self.out_tree, self.master_df)
                        
                        # Regenerate the preview with updated pricing
                        self._regenerate_left_preview()
                        
                        # Update the configure panel if visible
                        if hasattr(self, 'name_combo') and self.name_combo.get():
                            self.populate_value_grid()
                            
            except Exception as e:
                print(f"Error recalculating pricing: {e}")

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

        def show_usage():
            """Show where the selected associated SKU is used"""
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a part number to see its usage.")
                return
            
            # Get the selected part number
            item = selection[0]
            part_number = tree.set(item, "Part Number")
            
            # Find usage across all base SKUs
            usage_instances = find_sku_usage(part_number)
            
            if not usage_instances:
                messagebox.showinfo("No Usage Found", 
                    f"Part number '{part_number}' is not currently used in any base SKU configurations.")
                return
            
            # Create usage display window
            usage_win = tk.Toplevel(win)
            usage_win.title(f"Usage of Part Number: {part_number}")
            usage_win.geometry("600x400")
            usage_win.transient(win)
            
            # Center the window
            usage_win.update_idletasks()
            x = win.winfo_x() + (win.winfo_width() // 2) - (usage_win.winfo_width() // 2)
            y = win.winfo_y() + (win.winfo_height() // 2) - (usage_win.winfo_height() // 2)
            usage_win.geometry(f"+{x}+{y}")
            
            tk.Label(usage_win, text=f"Part Number '{part_number}' is used in:", 
                    font=('Segoe UI', 10, 'bold')).pack(pady=10)
            
            # Create treeview for usage display
            usage_frame = ttk.Frame(usage_win)
            usage_frame.pack(fill='both', expand=True, padx=10, pady=(0,10))
            
            usage_tree = ttk.Treeview(usage_frame, show="headings", selectmode="browse")
            usage_tree["columns"] = ["Base SKU", "Option Name", "Option Value", "Associated SKUs"]
            
            for col in usage_tree["columns"]:
                usage_tree.heading(col, text=col)
                usage_tree.column(col, anchor='w')
            
            # Scrollbars
            v_scroll = ttk.Scrollbar(usage_frame, orient="vertical", command=usage_tree.yview)
            h_scroll = ttk.Scrollbar(usage_frame, orient="horizontal", command=usage_tree.xview)
            usage_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
            
            # Grid layout
            usage_tree.grid(row=0, column=0, sticky='nsew')
            v_scroll.grid(row=0, column=1, sticky='ns')
            h_scroll.grid(row=1, column=0, sticky='ew')
            usage_frame.grid_rowconfigure(0, weight=1)
            usage_frame.grid_columnconfigure(0, weight=1)
            
            # Populate usage tree
            for base_sku, options in usage_instances.items():
                for option in options:
                    usage_tree.insert("", "end", values=(
                        base_sku,
                        option['Name'],
                        option['Value'],
                        option['Associated SKUs']
                    ), tags=(base_sku,))
            
            # Handle double-click to load base SKU
            def on_usage_double_click(event):
                item = usage_tree.selection()
                if item:
                    base_sku = usage_tree.item(item[0])['values'][0]
                    load_base_sku_in_main(base_sku)
                    usage_win.destroy()
            
            usage_tree.bind("<Double-1>", on_usage_double_click)
            
            # Buttons
            btn_frame_usage = ttk.Frame(usage_win)
            btn_frame_usage.pack(fill='x', padx=10, pady=(0,10))
            
            def load_selected():
                selection = usage_tree.selection()
                if selection:
                    base_sku = usage_tree.item(selection[0])['values'][0]
                    load_base_sku_in_main(base_sku)
                    usage_win.destroy()
                else:
                    messagebox.showwarning("No Selection", "Please select a base SKU to load.")
            
            ttk.Button(btn_frame_usage, text="Load Selected Base SKU", command=load_selected).pack(side='right', padx=(5,0))
            ttk.Button(btn_frame_usage, text="Close", command=usage_win.destroy).pack(side='right')
            
            # Instructions
            tk.Label(usage_win, text="Double-click or select a row and click 'Load Selected Base SKU' to open it in the main window.", 
                    font=('Segoe UI', 8), foreground='gray').pack(pady=(0,5))

        def load_base_sku_in_main(base_sku):
            """Load the specified base SKU in the main application window"""
            try:
                # Get the latest database
                db_path = self._get_latest_database_path()
                if not os.path.exists(db_path):
                    messagebox.showerror("Database Error", "Database file not found.")
                    return
                
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
                
                if base_sku not in db:
                    messagebox.showerror("Base SKU Not Found", f"Base SKU '{base_sku}' not found in database.")
                    return
                
                # Load the base SKU data into the main window
                state = db[base_sku]
                
                # Update DataFrames
                self.input_df = pd.DataFrame(state['input_df'])
                self.master_df = pd.DataFrame(state['master_df'])
                in_tree_df = pd.DataFrame(state['in_tree_df'])
                
                # Restore option names
                self.option_names = sorted({
                    name for blob in self.input_df['Options']
                    for name, _ in VALUE_RE.findall(str(blob))
                })
                
                # Update UI fields
                self.base_price_entry.delete(0, 'end')
                self.base_price_entry.insert(0, state.get('base_price', ''))
                self.base_weight_entry.delete(0, 'end')
                self.base_weight_entry.insert(0, state.get('base_weight', ''))
                self.base_sku_entry.delete(0, 'end')
                self.base_sku_entry.insert(0, state.get('base_sku', ''))
                self.prefix_entry.delete(0, 'end')
                self.prefix_entry.insert(0, state.get('prefix', ''))
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
                
                # Regenerate pricing
                self._regenerate_left_preview()
                self.current_prog_path = db_path
                self.unsaved_changes = False
                self._update_title()
                self._resize_bottom_panels()
                
                # Close the cost explorer window
                win.destroy()
                
                
            except Exception as e:
                messagebox.showerror("Load Error", f"Error loading base SKU '{base_sku}':\n{e}")

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x', padx=10, pady=(0, 10))
        ttk.Button(btn_frame, text="Show Usage", command=show_usage).pack(side='left')
        ttk.Button(btn_frame, text="Exit", command=on_window_close).pack(side='right')

        # Auto-save when window is closed
        def on_window_close():
            save_all()  # Save changes
            win.destroy()
        
        win.protocol("WM_DELETE_WINDOW", on_window_close)

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
                # Only update existing entries - don't create new ones
                db[base_sku]['last_export'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(db_path, "w", encoding="utf-8") as f:
                    json.dump(db, f, indent=2)
            else:
                # DO NOT create new entries - this was the bug!
                # If the base SKU isn't in the database, something is wrong
                print(f"Warning: Attempted to update export time for non-existent base SKU: {base_sku}")
                print("This indicates the base SKU was not properly saved before export.")
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
            
        # Get all database files (including timestamped temp files)
        files = [f for f in os.listdir(folder) if f.startswith("sku_database") and f.endswith(".json")]
        if not files:
            return os.path.join(folder, "sku_database.json")
            
        # Sort by modification time (newest first) to get the most recent file
        file_paths = [(f, os.path.getmtime(os.path.join(folder, f))) for f in files]
        file_paths.sort(key=lambda x: x[1], reverse=True)
        
        latest_file = os.path.join(folder, file_paths[0][0])
        return latest_file
    
    def delete_base_sku(self):
        if not messagebox.askyesno("Delete Confirmation", "Are you sure you want to delete the currently loaded base SKU from the database?"):
            return
        base_sku = self.base_sku_entry.get().strip() or self.base_sku
        if not base_sku:
            messagebox.showwarning("No Base SKU", "No base SKU selected to delete.")
            return
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            messagebox.showwarning("No Database", "No database file found.")
            return
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
            if base_sku in db:
                del db[base_sku]
                with open(db_path, "w", encoding="utf-8") as f:
                    json.dump(db, f, indent=2)
                # Clear current UI
                self.input_df = None
                self.master_df = None
                self.base_sku = ""
                self.base_sku_entry.delete(0, 'end')
                self.base_price_entry.delete(0, 'end')
                self.base_weight_entry.delete(0, 'end')
                self.prefix_entry.delete(0, 'end')
                self.in_tree.delete(*self.in_tree.get_children())
                self.out_tree.delete(*self.out_tree.get_children())
                self.breakdown_tree.delete(*self.breakdown_tree.get_children())
                for widget in self.grid_container.winfo_children():
                    widget.destroy()
                self.name_combo['values'] = []
                self.name_combo.set('')
                self.unsaved_changes = False
                self._update_title()
                messagebox.showinfo("Deleted", f"Base SKU '{base_sku}' has been deleted from the database.")
            else:
                messagebox.showwarning("Not Found", f"Base SKU '{base_sku}' not found in database.")
        except Exception as e:
            messagebox.showerror("Delete Error", f"Failed to delete base SKU:\n{e}")

    def generate_base_price_change_report(self):
        """Generate a report showing price changes between two database versions for base SKUs"""
        self._generate_price_change_report(report_type="base")

    def generate_associated_price_change_report(self):
        """Generate a report showing price changes between two database versions for associated SKUs"""
        self._generate_price_change_report(report_type="associated")

    def _generate_price_change_report(self, report_type="base"):
        """
        Generate price change reports comparing two database versions
        report_type: "base" for base SKU prices, "associated" for associated SKU prices
        """
        # Get current database
        current_db_path = self._get_latest_database_path()
        if not os.path.exists(current_db_path):
            messagebox.showerror("Error", "No current database found.")
            return

        # Show calendar dialog to select old database date
        old_db_path = self._select_database_by_date()
        if not old_db_path:
            return

        try:
            # Load both databases
            with open(current_db_path, "r", encoding="utf-8") as f:
                current_db = json.load(f)
            with open(old_db_path, "r", encoding="utf-8") as f:
                old_db = json.load(f)

            # Extract dates from filenames for display
            current_date = self._extract_date_from_filename(current_db_path)
            old_date = self._extract_date_from_filename(old_db_path)

            if report_type == "base":
                report_data = self._compare_base_sku_prices(old_db, current_db, old_date, current_date)
                report_title = "Base SKU Price Change Report"
            else:
                report_data = self._compare_associated_sku_prices(old_db, current_db, old_date, current_date)
                report_title = "Associated SKU Price Change Report"

            if not report_data:
                messagebox.showinfo("No Changes", "No price changes found between the selected databases.")
                return

            # Show report window
            self._show_price_change_report(report_data, report_title, old_date, current_date, report_type)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report:\n{e}")

    def _select_database_by_date(self):
        """Show a calendar dialog to select a date and find the corresponding database file"""
        # Create calendar dialog
        dialog = tk.Toplevel(self)
        dialog.title("Select Database Date")
        dialog.geometry("400x350")
        dialog.transient(self)
        dialog.grab_set()

        selected_date = None
        
        # Create a simple calendar using standard tkinter widgets
        cal_frame = ttk.Frame(dialog)
        cal_frame.pack(fill='both', expand=True, padx=20, pady=20)

        ttk.Label(cal_frame, text="Select a date to find the corresponding database:", 
                 font=('Segoe UI', 11)).pack(pady=(0, 20))

        # Year and month selection
        date_frame = ttk.Frame(cal_frame)
        date_frame.pack(pady=10)

        current_date = datetime.datetime.now()
        
        ttk.Label(date_frame, text="Year:").grid(row=0, column=0, padx=5)
        year_var = tk.IntVar(value=current_date.year)
        year_spin = ttk.Spinbox(date_frame, from_=2020, to=2030, width=8, textvariable=year_var)
        year_spin.grid(row=0, column=1, padx=5)

        ttk.Label(date_frame, text="Month:").grid(row=0, column=2, padx=5)
        month_var = tk.IntVar(value=current_date.month)
        month_spin = ttk.Spinbox(date_frame, from_=1, to=12, width=8, textvariable=month_var)
        month_spin.grid(row=0, column=3, padx=5)

        ttk.Label(date_frame, text="Day:").grid(row=0, column=4, padx=5)
        day_var = tk.IntVar(value=current_date.day)
        day_spin = ttk.Spinbox(date_frame, from_=1, to=31, width=8, textvariable=day_var)
        day_spin.grid(row=0, column=5, padx=5)

        # Available databases list
        ttk.Label(cal_frame, text="Available databases for selected date:", 
                 font=('Segoe UI', 10)).pack(pady=(20, 5), anchor='w')
        
        # Listbox for available databases
        list_frame = ttk.Frame(cal_frame)
        list_frame.pack(fill='both', expand=True, pady=5)
        
        db_listbox = tk.Listbox(list_frame)
        db_listbox.pack(side='left', fill='both', expand=True)
        
        db_scroll = ttk.Scrollbar(list_frame, orient='vertical', command=db_listbox.yview)
        db_scroll.pack(side='right', fill='y')
        db_listbox.configure(yscrollcommand=db_scroll.set)

        def update_database_list(*args):
            """Update the list of available databases for the selected date"""
            try:
                selected_year = year_var.get()
                selected_month = month_var.get()
                selected_day = day_var.get()
                
                # Create target date
                target_date = datetime.date(selected_year, selected_month, selected_day)
                
                # Find databases for this date
                data_dir = self._get_database_folder()
                db_files = []
                
                for filename in os.listdir(data_dir):
                    if filename.startswith('sku_database_') and filename.endswith('.json') and filename != 'sku_database_temp.json':
                        try:
                            # Extract date from filename (format: sku_database_YYYY-MM-DD_HH-MM-SS.json)
                            date_part = filename.replace('sku_database_', '').replace('.json', '').split('_')[0]
                            file_date = datetime.datetime.strptime(date_part, '%Y-%m-%d').date()
                            
                            # Check if file date matches target date (within 3 days)
                            date_diff = abs((file_date - target_date).days)
                            if date_diff <= 3:
                                full_path = os.path.join(data_dir, filename)
                                file_time = datetime.datetime.fromtimestamp(os.path.getmtime(full_path))
                                db_files.append((filename, file_date, file_time, full_path, date_diff))
                        except (ValueError, IndexError):
                            continue
                
                # Sort by date difference (closest first), then by time (latest first)
                db_files.sort(key=lambda x: (x[4], -x[2].timestamp()))
                
                # Update listbox
                db_listbox.delete(0, tk.END)
                for filename, file_date, file_time, full_path, date_diff in db_files:
                    if date_diff == 0:
                        display_text = f"{filename} (Exact match - {file_time.strftime('%H:%M:%S')})"
                    else:
                        display_text = f"{filename} ({date_diff} days off - {file_time.strftime('%H:%M:%S')})"
                    db_listbox.insert(tk.END, display_text)
                
                # Store file paths for selection
                db_listbox.file_paths = [item[3] for item in db_files]
                
            except Exception as e:
                print(f"Error updating database list: {e}")

        # Bind date changes to update list
        year_var.trace_add('write', update_database_list)
        month_var.trace_add('write', update_database_list)
        day_var.trace_add('write', update_database_list)

        # Initial population
        update_database_list()

        # Button frame
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', padx=20, pady=20)

        selected_path = None

        def on_select():
            nonlocal selected_path
            selection = db_listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a database file.")
                return
            
            selected_path = db_listbox.file_paths[selection[0]]
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side='right', padx=(5, 0))
        ttk.Button(button_frame, text="Select", command=on_select).pack(side='right')

        # Center dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        self.wait_window(dialog)
        return selected_path

    def _extract_date_from_filename(self, filepath):
        """Extract and format date from database filename"""
        try:
            filename = os.path.basename(filepath)
            if filename == "sku_database_temp.json":
                return datetime.datetime.now().strftime('%Y-%m-%d')
            
            # Extract date part (format: sku_database_YYYY-MM-DD_HH-MM-SS.json)
            date_part = filename.replace('sku_database_', '').replace('.json', '')
            if '_' in date_part:
                date_str = date_part.split('_')[0]
                return date_str
            return "Unknown"
        except Exception:
            return "Unknown"

    def _compare_base_sku_prices(self, old_db, current_db, old_date, current_date):
        """Compare base SKU prices between two databases"""
        changes = []
        
        # Find common base SKUs
        common_skus = set(old_db.keys()) & set(current_db.keys())
        
        for base_sku in sorted(common_skus):
            try:
                old_price = float(old_db[base_sku].get('base_price', 0))
                current_price = float(current_db[base_sku].get('base_price', 0))
                
                if old_price != current_price:
                    old_weight = old_db[base_sku].get('base_weight', '0')
                    current_weight = current_db[base_sku].get('base_weight', '0')
                    
                    changes.append({
                        'Base SKU': base_sku,
                        'Item Type': 'Base Product',
                        'Option Name': '',
                        'Option Value': '',
                        f'Price ({old_date})': f"${old_price:.2f}",
                        f'Price ({current_date})': f"${current_price:.2f}",
                        'Price Change': f"${current_price - old_price:+.2f}",
                        'Old Weight': f"{old_weight} lb",
                        'Current Weight': f"{current_weight} lb"
                    })
            except (ValueError, TypeError):
                continue
        
        return changes

    def _compare_associated_sku_prices(self, old_db, current_db, old_date, current_date):
        """Compare associated SKU prices between two databases"""
        changes = []
        
        # Find common base SKUs
        common_skus = set(old_db.keys()) & set(current_db.keys())
        
        for base_sku in sorted(common_skus):
            try:
                old_master_df = pd.DataFrame(old_db[base_sku].get('master_df', []))
                current_master_df = pd.DataFrame(current_db[base_sku].get('master_df', []))
                
                # Compare each option
                if not old_master_df.empty and not current_master_df.empty:
                    for _, old_row in old_master_df.iterrows():
                        option_name = old_row.get('Name', '')
                        option_value = old_row.get('Value', '')
                        
                        # Find corresponding row in current database
                        current_matches = current_master_df[
                            (current_master_df['Name'] == option_name) & 
                            (current_master_df['Value'] == option_value)
                        ]
                        
                        if not current_matches.empty:
                            current_row = current_matches.iloc[0]
                            
                            try:
                                old_cost = float(old_row.get("Add'l Cost", 0))
                                current_cost = float(current_row.get("Add'l Cost", 0))
                                
                                if old_cost != current_cost:
                                    old_weight = old_row.get("Add'l Weight", '0')
                                    current_weight = current_row.get("Add'l Weight", '0')
                                    
                                    # Get associated SKUs info
                                    old_assoc = old_row.get("Associated SKUs", '')
                                    current_assoc = current_row.get("Associated SKUs", '')
                                    
                                    changes.append({
                                        'Base SKU': base_sku,
                                        'Item Type': 'Option',
                                        'Option Name': option_name,
                                        'Option Value': option_value,
                                        f'Price ({old_date})': f"${old_cost:.2f}",
                                        f'Price ({current_date})': f"${current_cost:.2f}",
                                        'Price Change': f"${current_cost - old_cost:+.2f}",
                                        'Old Weight': f"{old_weight} lb",
                                        'Current Weight': f"{current_weight} lb",
                                        'Old Associated SKUs': old_assoc,
                                        'Current Associated SKUs': current_assoc
                                    })
                            except (ValueError, TypeError):
                                continue
            except Exception as e:
                print(f"Error comparing {base_sku}: {e}")
                continue
        
        return changes

    def _show_price_change_report(self, report_data, report_title, old_date, current_date, report_type):
        """Display the price change report in a new window"""
        report_win = tk.Toplevel(self)
        report_win.title(report_title)
        report_win.geometry("1200x700")
        report_win.transient(self)

        # Header
        header_frame = ttk.Frame(report_win)
        header_frame.pack(fill='x', padx=10, pady=10)

        title_label = ttk.Label(header_frame, text=report_title, font=('Segoe UI', 16, 'bold'))
        title_label.pack()

        subtitle_label = ttk.Label(header_frame, 
                                  text=f"Comparing {old_date} vs {current_date} • {len(report_data)} changes found",
                                  font=('Segoe UI', 10))
        subtitle_label.pack(pady=(5, 0))

        # Export buttons
        button_frame = ttk.Frame(report_win)
        button_frame.pack(fill='x', padx=10, pady=(0, 10))

        def export_to_excel():
            export_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                initialfile=f"{report_title.replace(' ', '_')}_{old_date}_vs_{current_date}.xlsx",
                filetypes=[('Excel', '*.xlsx')]
            )
            if export_path:
                try:
                    df = pd.DataFrame(report_data)
                    df.to_excel(export_path, index=False)
                    if messagebox.askyesno("Exported", f"Report saved to {export_path}\nOpen in Excel?"):
                        try:
                            subprocess.Popen(['start', 'excel', export_path], shell=True)
                        except Exception:
                            messagebox.showerror("Error", "Can't open Excel.")
                except Exception as e:
                    messagebox.showerror("Export Error", f"Failed to export:\n{e}")

        def export_to_pdf():
            export_path = filedialog.asksaveasfilename(
                defaultextension='.pdf',
                initialfile=f"{report_title.replace(' ', '_')}_{old_date}_vs_{current_date}.pdf",
                filetypes=[('PDF', '*.pdf')]
            )
            if export_path:
                try:
                    self._generate_price_change_pdf(report_data, report_title, old_date, current_date, export_path, report_type)
                    if messagebox.askyesno("Exported", f"PDF report saved to {export_path}\nOpen PDF?"):
                        try:
                            subprocess.Popen(['start', export_path], shell=True)
                        except Exception:
                            messagebox.showerror("Error", "Can't open PDF viewer.")
                except Exception as e:
                    messagebox.showerror("Export Error", f"Failed to generate PDF:\n{e}")

        ttk.Button(button_frame, text="Export to Excel", command=export_to_excel).pack(side='right', padx=(0, 5))
        ttk.Button(button_frame, text="Generate PDF Report", command=export_to_pdf).pack(side='right')

        # Treeview for report data
        tree_frame = ttk.Frame(report_win)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Determine columns based on report type
        if report_type == "base":
            columns = ['Base SKU', 'Item Type', f'Price ({old_date})', f'Price ({current_date})', 
                      'Price Change', 'Old Weight', 'Current Weight']
        else:
            columns = ['Base SKU', 'Item Type', 'Option Name', 'Option Value', 
                      f'Price ({old_date})', f'Price ({current_date})', 'Price Change', 
                      'Old Weight', 'Current Weight', 'Old Associated SKUs', 'Current Associated SKUs']

        tree = ttk.Treeview(tree_frame, show='headings')
        tree['columns'] = columns

        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'Base SKU':
                tree.column(col, width=150, anchor='w')
            elif col in ['Option Name', 'Option Value']:
                tree.column(col, width=120, anchor='w')
            elif col in [f'Price ({old_date})', f'Price ({current_date})', 'Price Change']:
                tree.column(col, width=100, anchor='center')
            elif 'Associated SKUs' in col:
                tree.column(col, width=200, anchor='w')
            else:
                tree.column(col, width=100, anchor='center')

        # Add scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # Populate tree
        for item in report_data:
            values = [item.get(col, '') for col in columns]
            tree.insert('', 'end', values=values)

        # Close button
        close_frame = ttk.Frame(report_win)
        close_frame.pack(fill='x', padx=10, pady=(0, 10))
        ttk.Button(close_frame, text="Close", command=report_win.destroy).pack(side='right')

    def _generate_price_change_pdf(self, report_data, report_title, old_date, current_date, export_path, report_type):
        """Generate a professional PDF report for price changes"""
        try:
            # Create the PDF document
            doc = SimpleDocTemplate(export_path, pagesize=letter,
                                   rightMargin=72, leftMargin=72,
                                   topMargin=72, bottomMargin=18)

            # Get styles
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=30,
                alignment=1,  # Center
                textColor=colors.black
            )

            subtitle_style = ParagraphStyle(
                'CustomSubtitle',
                parent=styles['Normal'],
                fontSize=12,
                spaceAfter=20,
                alignment=1,  # Center
                textColor=colors.grey
            )

            # Build the story (content)
            story = []

            # Title
            title = Paragraph(report_title, title_style)
            story.append(title)

            # Subtitle
            subtitle = Paragraph(f"Comparing databases from {old_date} to {current_date}", subtitle_style)
            story.append(subtitle)

            # Summary
            total_changes = len(report_data)
            price_increases = len([item for item in report_data if float(item['Price Change'].replace('$', '').replace('+', '')) > 0])
            price_decreases = total_changes - price_increases

            summary_text = f"<b>Summary:</b><br/>"
            summary_text += f"Total price changes: {total_changes}<br/>"
            summary_text += f"Price increases: {price_increases}<br/>"
            summary_text += f"Price decreases: {price_decreases}"

            summary_para = Paragraph(summary_text, styles['Normal'])
            story.append(summary_para)
            story.append(Spacer(1, 20))

            # Prepare table data
            if report_type == "base":
                headers = ['Base SKU', 'Type', f'Price\n({old_date})', f'Price\n({current_date})', 'Change']
                col_widths = [2.5*inch, 1*inch, 1*inch, 1*inch, 1*inch]
            else:
                headers = ['Base SKU', 'Option Name', 'Option Value', f'Old\nPrice', f'New\nPrice', 'Change']
                col_widths = [1.8*inch, 1.5*inch, 1.5*inch, 0.8*inch, 0.8*inch, 0.8*inch]

            table_data = [headers]

            # Add data rows
            for item in report_data:
                if report_type == "base":
                    row = [
                        item['Base SKU'],
                        item['Item Type'],
                        item[f'Price ({old_date})'],
                        item[f'Price ({current_date})'],
                        item['Price Change']
                    ]
                else:
                    row = [
                        item['Base SKU'],
                        item['Option Name'],
                        item['Option Value'],
                        item[f'Price ({old_date})'],
                        item[f'Price ({current_date})'],
                        item['Price Change']
                    ]
                table_data.append(row)

            # Create table
            table = Table(table_data, colWidths=col_widths, repeatRows=1)

            # Apply table style
            table.setStyle(TableStyle([
                # Header row styling
                ('BACKGROUND', (0, 0), (-1, 0), colors.black),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),

                # Data rows styling
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),    # First column left-aligned
                ('ALIGN', (-3, 1), (-1, -1), 'CENTER'), # Price columns centered
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),

                # Grid lines
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

                # Padding
                ('TOPPADDING', (0, 1), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ]))

            story.append(table)

            # Add footer information
            story.append(Spacer(1, 30))
            footer_text = f"Report generated on: <b>{datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p')}</b>"
            footer_para = Paragraph(footer_text, styles['Normal'])
            story.append(footer_para)

            # Build PDF
            doc.build(story)

        except ImportError:
            messagebox.showerror("Missing Library", 
                "ReportLab library is required for PDF generation.\n\n"
                "Please install it using:\npip install reportlab")
            raise
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
                # Database could not be repaired, offer backup restoration
                result = messagebox.askyesnocancel(
                    "Database Corrupted", 
                    "The database file could not be repaired.\n\n"
                    "Would you like to restore from your latest uncorrupted local backup?\n\n"
                    "• YES: Restore from local backup to cloud\n"
                    "• NO: Cancel operation\n"
                    "• CANCEL: Abort operation"
                )
                
                if result is None or not result:  # Cancel or No
                    return
                else:  # Yes - restore from backup
                    if restore_from_local_backup():
                        # After successful restore, try loading again
                        try:
                            with open(db_path, "r", encoding="utf-8") as f:
                                db = json.load(f)
                        except:
                            messagebox.showerror("Restore Failed", "Could not load database after restore.")
                            return
                    else:
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

    def generate_associated_sku_report(self):
        """Generate a report showing all base SKUs that contain a specific associated SKU"""
        # First, get all available associated SKUs from the cost database
        cost_db_path = os.path.join(self._get_database_folder(), "cost_db.json")
        if not os.path.exists(cost_db_path):
            messagebox.showinfo("No Cost DB", "No cost database found. Import a cost file first.")
            return
            
        with open(cost_db_path, "r", encoding="utf-8") as f:
            cost_db = json.load(f)
        
        if not cost_db:
            messagebox.showinfo("No Data", "No part numbers in cost database.")
            return
            
        # Create dialog to select an associated SKU
        class AssociatedSkuSearchDialog(tk.Toplevel):
            def __init__(self, parent, available_skus):
                super().__init__(parent)
                self.title("Select Associated SKU")
                self.geometry("400x500")
                self.selected = None
                self.all_skus = sorted(available_skus)
                
                # Make dialog modal
                self.transient(parent)
                self.grab_set()
                
                frame = ttk.Frame(self)
                frame.pack(fill='both', expand=True, padx=10, pady=10)
                
                ttk.Label(frame, text="Search for Associated SKU:").pack(anchor='w', pady=(0,5))
                
                # Search entry
                self.var = tk.StringVar()
                self.var.trace_add('write', self.update_list)
                entry = tk.Entry(frame, textvariable=self.var)
                entry.pack(fill='x', pady=(0,10))
                
                # Create listbox with scrollbar
                list_frame = tk.Frame(frame)
                list_frame.pack(fill='both', expand=True, pady=(0,10))
                
                scrollbar = tk.Scrollbar(list_frame)
                scrollbar.pack(side='right', fill='y')
                
                self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
                self.listbox.pack(side='left', fill='both', expand=True)
                scrollbar.config(command=self.listbox.yview)
                
                # Initial population
                self.populate_listbox(self.all_skus)
                
                self.listbox.bind('<Double-1>', self.select)
                self.listbox.bind('<Return>', self.select)
                
                btn_frame = tk.Frame(frame)
                btn_frame.pack(fill='x', pady=(0,0))
                tk.Button(btn_frame, text="Generate Report", command=self.select).pack(side='right', padx=(5,0))
                tk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side='right')
                
                entry.focus_set()
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
                    self.populate_listbox(self.all_skus)
                else:
                    filtered = [sku for sku in self.all_skus if search_text in sku.lower()]
                    self.populate_listbox(filtered)
                        
            def select(self, event=None):
                selection = self.listbox.curselection()
                if selection:
                    self.selected = self.listbox.get(selection[0])
                    self.destroy()
        
        # Show the dialog
        dlg = AssociatedSkuSearchDialog(self, list(cost_db.keys()))
        self.wait_window(dlg)
        selected_sku = dlg.selected
        
        if not selected_sku:
            return
            
        # Now search through all base SKUs to find ones that contain this associated SKU
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            messagebox.showinfo("No Database", "No SKU database found.")
            return
            
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except json.JSONDecodeError:
            messagebox.showerror("Database Error", "Failed to read database file.")
            return
            
        # Search for the selected associated SKU in all base SKUs
        report_data = []
        for base_sku, state in db.items():
            master_df = pd.DataFrame(state.get('master_df', []))
            for idx, row in master_df.iterrows():
                assoc_str = row.get("Associated SKUs", "")
                if assoc_str:
                    sku_costs = parse_associated_skus(assoc_str)
                    for sku, cost, partnumber in sku_costs:
                        if sku == selected_sku or partnumber == selected_sku:
                            report_data.append({
                                'Base SKU': base_sku,
                                'Option Name': row.get('Name', ''),
                                'Option Value': row.get('Value', ''),
                                'Associated SKU': sku,
                                'Cost': cost,
                                'Part Number': partnumber
                            })
        
        if not report_data:
            messagebox.showinfo("No Results", f"Associated SKU '{selected_sku}' was not found in any base SKUs.")
            return
            
        # Create report window
        report_win = tk.Toplevel(self)
        report_win.title(f"Associated SKU Usage Report: {selected_sku}")
        report_win.geometry("900x600")
        
        # Add export button frame at top
        button_frame = ttk.Frame(report_win)
        button_frame.pack(fill='x', padx=10, pady=(10,0))
        
        ttk.Label(button_frame, text=f"Found {len(report_data)} occurrences of '{selected_sku}'", 
                 font=('Segoe UI', 11, 'bold')).pack(side='left')
        
        def export_to_excel():
            export_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                initialfile=f"Associated_SKU_Report_{selected_sku}.xlsx",
                filetypes=[('Excel', '*.xlsx')]
            )
            if export_path:
                try:
                    df = pd.DataFrame(report_data)
                    df.to_excel(export_path, index=False)
                    if messagebox.askyesno("Exported", f"Report saved to {export_path}\nOpen in Excel?"):
                        try:
                            subprocess.Popen(['start', 'excel', export_path], shell=True)
                        except Exception:
                            messagebox.showerror("Error", "Can't open Excel.")
                except Exception as e:
                    messagebox.showerror("Export Error", f"Failed to export:\n{e}")
        
        def export_to_pdf():
            export_path = filedialog.asksaveasfilename(
                defaultextension='.pdf',
                initialfile=f"Option_Containing_Report_{selected_sku}.pdf",
                filetypes=[('PDF', '*.pdf')]
            )
            if export_path:
                try:
                    self._generate_professional_pdf_report(report_data, selected_sku, export_path)
                    if messagebox.askyesno("Exported", f"Professional report saved to {export_path}\nOpen PDF?"):
                        try:
                            subprocess.Popen(['start', export_path], shell=True)
                        except Exception:
                            messagebox.showerror("Error", "Can't open PDF viewer.")
                except Exception as e:
                    messagebox.showerror("Export Error", f"Failed to generate PDF:\n{e}")
        
        ttk.Button(button_frame, text="Export to Excel", command=export_to_excel).pack(side='right', padx=(0,5))
        ttk.Button(button_frame, text="Generate PDF Report", command=export_to_pdf).pack(side='right')
        
        # Create treeview to display results
        tree_frame = ttk.Frame(report_win)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        tree = ttk.Treeview(tree_frame, show='headings')
        columns = ['Base SKU', 'Option Name', 'Option Value', 'Associated SKU', 'Cost', 'Part Number']
        tree['columns'] = columns
        
        for col in columns:
            tree.heading(col, text=col)
            if col == 'Base SKU':
                tree.column(col, width=150, anchor='w')
            elif col in ['Option Name', 'Option Value']:
                tree.column(col, width=120, anchor='w')
            elif col == 'Associated SKU':
                tree.column(col, width=100, anchor='center')
            elif col == 'Cost':
                tree.column(col, width=80, anchor='center')
            elif col == 'Part Number':
                tree.column(col, width=100, anchor='center')
        
        # Add scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        # Populate the tree
        for item in report_data:
            tree.insert('', 'end', values=[
                item['Base SKU'],
                item['Option Name'],
                item['Option Value'],
                item['Associated SKU'],
                item['Cost'],
                item['Part Number']
            ])
        
        # Add close button
        close_frame = ttk.Frame(report_win)
        close_frame.pack(fill='x', padx=10, pady=(0,10))
        ttk.Button(close_frame, text="Close", command=report_win.destroy).pack(side='right')

    def _generate_professional_pdf_report(self, report_data, selected_sku, export_path):
        """Generate a professional PDF report based on the associated SKU usage data"""
        try:
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.enums import TA_CENTER, TA_LEFT
        except ImportError:
            messagebox.showerror("Missing Library", 
                "ReportLab library is required for PDF generation.\n\n"
                "Please install it using:\npip install reportlab")
            return
        
        # Create the PDF document
        doc = SimpleDocTemplate(export_path, pagesize=letter,
                               rightMargin=72, leftMargin=72,
                               topMargin=72, bottomMargin=18)
        
        # Get styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            alignment=TA_CENTER,
            textColor=colors.black
        )
        
        subtitle_style = ParagraphStyle(
            'CustomSubtitle',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=20,
            alignment=TA_CENTER,
            textColor=colors.grey
        )
        
        # Build the story (content)
        story = []
        
        # Title
        title = Paragraph("Associated SKU Usage Report", title_style)
        story.append(title)
        
        # Subtitle with selected SKU
        subtitle = Paragraph(f"Base SKUs containing Associated SKU: <b>{selected_sku}</b>", subtitle_style)
        story.append(subtitle)
        
        story.append(Spacer(1, 20))
        
        # Prepare data for the table - group by base SKU and get price/weight info
        base_sku_summary = {}
        
        # Get base price and weight information from the database
        db_path = self._get_latest_database_path()
        db_info = {}
        if os.path.exists(db_path):
            try:
                with open(db_path, "r", encoding="utf-8") as f:
                    db = json.load(f)
                for base_sku, state in db.items():
                    base_price = state.get('base_price', '0')
                    base_weight = state.get('base_weight', '0')
                    db_info[base_sku] = {
                        'price': base_price,
                        'weight': base_weight
                    }
            except Exception:
                pass
        
        # Group report data by base SKU
        for item in report_data:
            base_sku = item['Base SKU']
            if base_sku not in base_sku_summary:
                base_sku_summary[base_sku] = {
                    'price': db_info.get(base_sku, {}).get('price', 'N/A'),
                    'weight': db_info.get(base_sku, {}).get('weight', 'N/A'),
                    'options': []
                }
            base_sku_summary[base_sku]['options'].append({
                'name': item['Option Name'],
                'value': item['Option Value'],
                'cost': item['Cost']
            })
        
        # Create table data
        table_data = [
            ['', 'BASE SKUS CONTAINING\n[' + selected_sku + ']', 'PRICE', 'WEIGHT']
        ]
        
        # Add data rows
        for i, (base_sku, info) in enumerate(sorted(base_sku_summary.items()), 1):
            price_text = f"${info['price']}" if info['price'] != 'N/A' else 'N/A'
            weight_text = f"{info['weight']} lb" if info['weight'] != 'N/A' else 'N/A'
            
            table_data.append([
                str(i),
                base_sku,
                price_text,
                weight_text
            ])
        
        # Create the table
        table = Table(table_data, colWidths=[0.5*inch, 4*inch, 1.5*inch, 1.5*inch])
        
        # Apply table style
        table.setStyle(TableStyle([
            # Header row styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.black),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            
            # Data rows styling
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # Row numbers centered
            ('ALIGN', (1, 1), (1, -1), 'LEFT'),    # Base SKUs left-aligned
            ('ALIGN', (2, 1), (-1, -1), 'CENTER'), # Price and weight centered
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            
            # Grid lines
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            
            # Padding
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ]))
        
        story.append(table)
        
        # Add summary information
        story.append(Spacer(1, 30))
        
        summary_text = f"Total Base SKUs found: <b>{len(base_sku_summary)}</b><br/>"
        summary_text += f"Total occurrences: <b>{len(report_data)}</b><br/>"
        summary_text += f"Generated on: <b>{datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p')}</b>"
        
        summary_para = Paragraph(summary_text, styles['Normal'])
        story.append(summary_para)
        
        # Build PDF
        doc.build(story)

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

    def generate_base_sku_summary_report(self):
        """Generate a report showing all base SKUs with their prices and weights"""
        # Load database
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            messagebox.showinfo("No Database", "No database file found.")
            return
        
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except json.JSONDecodeError:
            # Try to repair first
            repaired_db = self.try_repair_json_file(db_path)
            if repaired_db is not None:
                db = repaired_db
                messagebox.showinfo("Database Repaired", "Database was corrupted but has been repaired.")
            else:
                # Database could not be repaired, offer backup restoration
                result = messagebox.askyesnocancel(
                    "Database Corrupted", 
                    "The database file could not be repaired.\n\n"
                    "Would you like to restore from your latest uncorrupted local backup?\n\n"
                    "• YES: Restore from local backup to cloud\n"
                    "• NO: Cancel report generation\n"
                    "• CANCEL: Abort operation"
                )
                
                if result is None or not result:  # Cancel or No
                    return
                else:  # Yes - restore from backup
                    if restore_from_local_backup():
                        # After successful restore, try loading again
                        try:
                            with open(db_path, "r", encoding="utf-8") as f:
                                db = json.load(f)
                        except:
                            messagebox.showerror("Restore Failed", "Could not load database after restore.")
                            return
                    else:
                        return
        
        if not db:
            messagebox.showinfo("No Data", "No base SKUs found in database.")
            return
        
        # Prepare report data
        report_data = []
        for base_sku, state in db.items():
            base_price = state.get('base_price', 'N/A')
            base_weight = state.get('base_weight', 'N/A')
            
            # Extract store information from parentheses
            store = ''
            clean_base_sku = base_sku
            if '(' in base_sku and ')' in base_sku:
                # Extract content within parentheses
                start = base_sku.find('(')
                end = base_sku.find(')')
                if start < end:
                    store = base_sku[start+1:end].strip()
                    clean_base_sku = base_sku[:start].strip()
            
            # Collect all associated SKUs from all options
            associated_skus = set()
            if 'master_df' in state:
                for option in state['master_df']:
                    assoc_str = option.get("Associated SKUs", "")
                    if assoc_str:
                        # Split by comma and clean each SKU
                        skus = [sku.strip() for sku in assoc_str.split(',') if sku.strip()]
                        associated_skus.update(skus)
            
            # Convert to sorted comma-separated string
            associated_skus_str = ', '.join(sorted(associated_skus)) if associated_skus else 'None'
            
            # Try to convert to numbers for proper sorting/formatting
            try:
                price_num = float(base_price) if base_price != 'N/A' else 0
                price_display = f"${price_num:.2f}" if base_price != 'N/A' else 'N/A'
            except (ValueError, TypeError):
                price_display = str(base_price) if base_price else 'N/A'
            
            try:
                weight_num = float(base_weight) if base_weight != 'N/A' else 0
                weight_display = f"{weight_num:.2f} lb" if base_weight != 'N/A' else 'N/A'
            except (ValueError, TypeError):
                weight_display = str(base_weight) if base_weight else 'N/A'
            
            report_data.append({
                'Base SKU': clean_base_sku,
                'Store': store,
                'Base Price': price_display,
                'Base Weight': weight_display,
                'Associated SKUs': associated_skus_str
            })
        
        # Sort by Base SKU name
        report_data.sort(key=lambda x: x['Base SKU'])
        
        # Create report window
        report_win = tk.Toplevel(self)
        report_win.title("Base SKU Summary Report")
        report_win.geometry("1200x600")
        
        # Add header and export buttons
        header_frame = ttk.Frame(report_win)
        header_frame.pack(fill='x', padx=10, pady=(10,0))
        
        ttk.Label(header_frame, text=f"Base SKU Summary - {len(report_data)} Total SKUs", 
                 font=('Segoe UI', 12, 'bold')).pack(side='left')
        
        def export_to_excel():
            export_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                initialfile="Base_SKU_Summary_Report.xlsx",
                filetypes=[('Excel', '*.xlsx')]
            )
            if export_path:
                try:
                    # Create DataFrame for export
                    df = pd.DataFrame(report_data)
                    df.to_excel(export_path, index=False)
                    if messagebox.askyesno("Exported", f"Report saved to {export_path}\nOpen in Excel?"):
                        try:
                            subprocess.Popen(['start', 'excel', export_path], shell=True)
                        except Exception:
                            messagebox.showerror("Error", "Can't open Excel.")
                except Exception as e:
                    messagebox.showerror("Export Error", f"Failed to export:\n{e}")
        
        def export_to_pdf():
            export_path = filedialog.asksaveasfilename(
                defaultextension='.pdf',
                initialfile="Base_SKU_Summary_Report.pdf",
                filetypes=[('PDF', '*.pdf')]
            )
            if export_path:
                try:
                    self._generate_base_sku_summary_pdf(report_data, export_path)
                    if messagebox.askyesno("Exported", f"Professional report saved to {export_path}\nOpen PDF?"):
                        try:
                            subprocess.Popen(['start', export_path], shell=True)
                        except Exception:
                            messagebox.showerror("Error", "Can't open PDF viewer.")
                except Exception as e:
                    messagebox.showerror("Export Error", f"Failed to generate PDF:\n{e}")
        
        ttk.Button(header_frame, text="Export to Excel", command=export_to_excel).pack(side='right', padx=(0,5))
        ttk.Button(header_frame, text="Generate PDF Report", command=export_to_pdf).pack(side='right')
        
        # Create treeview to display results
        tree_frame = ttk.Frame(report_win)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        tree = ttk.Treeview(tree_frame, show='headings')
        columns = ['Base SKU', 'Store', 'Base Price', 'Base Weight', 'Associated SKUs']
        tree['columns'] = columns
        
        # Configure columns
        tree.heading('Base SKU', text='Base SKU')
        tree.column('Base SKU', width=250, anchor='w')
        
        tree.heading('Store', text='Store')
        tree.column('Store', width=80, anchor='center')
        
        tree.heading('Base Price', text='Base Price')
        tree.column('Base Price', width=100, anchor='center')
        
        tree.heading('Base Weight', text='Base Weight')
        tree.column('Base Weight', width=100, anchor='center')
        
        tree.heading('Associated SKUs', text='Associated SKUs')
        tree.column('Associated SKUs', width=300, anchor='w')
        
        # Add scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        # Populate the tree
        for item in report_data:
            tree.insert('', 'end', values=[
                item['Base SKU'],
                item['Store'],
                item['Base Price'],
                item['Base Weight'],
                item['Associated SKUs']
            ])
        
        # Add close button
        close_frame = ttk.Frame(report_win)
        close_frame.pack(fill='x', padx=10, pady=(0,10))
        ttk.Button(close_frame, text="Close", command=report_win.destroy).pack(side='right')

    def _generate_base_sku_summary_pdf(self, report_data, export_path):
        """Generate a professional PDF report for base SKU summary"""
        try:
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.enums import TA_CENTER, TA_LEFT
            import datetime
        except ImportError:
            messagebox.showerror("Missing Library", 
                "ReportLab is required for PDF generation. Please install it using:\npip install reportlab")
            return
        
        # Create the PDF document
        doc = SimpleDocTemplate(export_path, pagesize=letter,
                               rightMargin=72, leftMargin=72,
                               topMargin=72, bottomMargin=18)
        
        # Get styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            alignment=TA_CENTER,
            textColor=colors.black
        )
        
        subtitle_style = ParagraphStyle(
            'CustomSubtitle',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=20,
            alignment=TA_CENTER,
            textColor=colors.grey
        )
        
        # Build the story (content)
        story = []
        
        # Title
        title = Paragraph("Base SKU Summary Report", title_style)
        story.append(title)
        
        # Subtitle
        subtitle = Paragraph(f"Complete listing of all base SKUs with pricing and weight information", subtitle_style)
        story.append(subtitle)
        
        story.append(Spacer(1, 20))
        
        # Create table data
        table_data = [
            ['#', 'BASE SKU', 'STORE', 'PRICE', 'WEIGHT', 'ASSOCIATED SKUS']
        ]
        
        # Add data rows
        for i, item in enumerate(report_data, 1):
            table_data.append([
                str(i),
                item['Base SKU'],
                item['Store'],
                item['Base Price'],
                item['Base Weight'],
                item['Associated SKUs']
            ])
        
        # Create the table with adjusted column widths
        table = Table(table_data, colWidths=[0.3*inch, 2*inch, 0.6*inch, 0.8*inch, 0.8*inch, 2.3*inch])
        
        # Apply table style
        table.setStyle(TableStyle([
            # Header row styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.black),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            
            # Data rows styling
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # Row numbers centered
            ('ALIGN', (1, 1), (1, -1), 'LEFT'),    # Base SKUs left-aligned
            ('ALIGN', (2, 1), (2, -1), 'CENTER'),  # Store centered
            ('ALIGN', (3, 1), (4, -1), 'CENTER'),  # Price and weight centered
            ('ALIGN', (5, 1), (5, -1), 'LEFT'),    # Associated SKUs left-aligned
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            
            # Grid lines
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            
            # Padding
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ]))
        
        story.append(table)
        
        # Add generation timestamp
        story.append(Spacer(1, 30))
        
        summary_text = f"Total Base SKUs: <b>{len(report_data)}</b><br/>"
        summary_text += f"Generated on: <b>{datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p')}</b>"
        
        summary_para = Paragraph(summary_text, styles['Normal'])
        story.append(summary_para)
        
        # Build PDF
        doc.build(story)

    def launch_hotkeys(self):
        """Launch HotKeys.py as a child process"""
        import subprocess
        import os
        
        # Get the path to HotKeys.py in the same directory as this script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        hotkeys_path = os.path.join(script_dir, "HotKeys.py")
        
        if not os.path.exists(hotkeys_path):
            messagebox.showerror("File Not Found", f"HotKeys.py not found at:\n{hotkeys_path}")
            return
        
        try:
            # Launch HotKeys.py as a separate process
            subprocess.Popen([sys.executable, hotkeys_path], cwd=script_dir)
        except Exception as e:
            messagebox.showerror("Launch Error", f"Failed to launch HotKeys:\n{e}")

    def launch_excel_sku_reorder(self):
        """Launch excel_sku_reorder_gui.py as a child process"""
        import subprocess
        import os
        
        # Get the path to excel_sku_reorder_gui.py in the same directory as this script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_tool_path = os.path.join(script_dir, "excel_sku_reorder_gui.py")
        
        if not os.path.exists(excel_tool_path):
            messagebox.showerror("File Not Found", f"excel_sku_reorder_gui.py not found at:\n{excel_tool_path}")
            return
        
        try:
            # Launch excel_sku_reorder_gui.py as a separate process
            subprocess.Popen([sys.executable, excel_tool_path], cwd=script_dir)
        except Exception as e:
            messagebox.showerror("Launch Error", f"Failed to launch Excel SKU Reorder Tool:\n{e}")

    def launch_duplicate_remover(self):
        """Launch duplicate_remover.py as a child process"""
        import subprocess
        import os
        
        # Get the path to duplicate_remover.py in the same directory as this script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        duplicate_remover_path = os.path.join(script_dir, "duplicate_remover.py")
        
        if not os.path.exists(duplicate_remover_path):
            messagebox.showerror("File Not Found", f"duplicate_remover.py not found at:\n{duplicate_remover_path}")
            return
        
        try:
            # Launch duplicate_remover.py as a separate process
            subprocess.Popen([sys.executable, duplicate_remover_path], cwd=script_dir)
        except Exception as e:
            messagebox.showerror("Launch Error", f"Failed to launch Duplicate Remover:\n{e}")

    def bulk_update_base_prices(self):
        """Open a window to bulk update base prices for all base SKUs"""
        # Load database
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            messagebox.showinfo("No Database", "No database file found.")
            return
        
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except json.JSONDecodeError:
            messagebox.showerror("Database Error", "Database file is corrupted.")
            return
        
        if not db:
            messagebox.showinfo("No Data", "No base SKUs found in database.")
            return
        
        # Create bulk update window
        bulk_win = tk.Toplevel(self)
        bulk_win.title("Bulk Update Base Prices")
        bulk_win.geometry("800x600")
        bulk_win.transient(self)
        bulk_win.grab_set()
        
        # Header
        header_frame = ttk.Frame(bulk_win)
        header_frame.pack(fill='x', padx=10, pady=(10,0))
        
        ttk.Label(header_frame, text=f"Bulk Update Base Prices - {len(db)} Base SKUs", 
                 font=('Segoe UI', 12, 'bold')).pack(side='left')
        
        # Instructions
        instructions = ttk.Label(bulk_win, 
            text="Double-click a price cell to edit. Changes will be applied when you click 'Save All Changes'.",
            font=('Segoe UI', 9), foreground='gray')
        instructions.pack(padx=10, pady=(5,10))
        
        # Create treeview for editing
        tree_frame = ttk.Frame(bulk_win)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=(0,10))
        
        # Add search functionality
        search_frame = ttk.Frame(tree_frame)
        search_frame.pack(fill='x', pady=(0,5))
        ttk.Label(search_frame, text="Search:").pack(side='left')
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side='left', padx=(5,0))
        
        # Add scrollbars container
        tree_container = ttk.Frame(tree_frame)
        tree_container.pack(fill='both', expand=True)
        
        tree = ttk.Treeview(tree_container, show='headings', selectmode='browse')
        columns = ['Base SKU', 'Store', 'Current Base Price', 'New Base Price']
        tree['columns'] = columns
        
        # Configure columns
        tree.heading('Base SKU', text='Base SKU')
        tree.column('Base SKU', width=250, anchor='w')
        
        tree.heading('Store', text='Store')
        tree.column('Store', width=100, anchor='center')
        
        tree.heading('Current Base Price', text='Current Base Price')
        tree.column('Current Base Price', width=150, anchor='center')
        
        tree.heading('New Base Price', text='New Base Price')
        tree.column('New Base Price', width=150, anchor='center')
        
        tree.pack(side='left', fill='both', expand=True)
        
        vsb = ttk.Scrollbar(tree_container, orient='vertical', command=tree.yview)
        vsb.pack(side='right', fill='y')
        tree.configure(yscrollcommand=vsb.set)
        
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
        hsb.pack(side='bottom', fill='x')
        tree.configure(xscrollcommand=hsb.set)
        
        # Data storage for tracking changes
        original_data = {}
        modified_data = {}
        
        # Populate tree with current data
        def populate_tree(filter_text=""):
            # Clear existing items
            for item in tree.get_children():
                tree.delete(item)
            
            original_data.clear()
            filter_lower = filter_text.lower()
            
            for base_sku, state in db.items():
                # Extract store information
                store = ''
                clean_base_sku = base_sku
                if '(' in base_sku and ')' in base_sku:
                    start = base_sku.find('(')
                    end = base_sku.find(')')
                    if start < end:
                        store = base_sku[start+1:end].strip()
                        clean_base_sku = base_sku[:start].strip()
                
                # Apply filter
                if filter_text and filter_lower not in clean_base_sku.lower() and filter_lower not in store.lower():
                    continue
                
                current_price = state.get('base_price', '')
                
                # Format price for display
                try:
                    price_num = float(current_price) if current_price else 0
                    price_display = f"${price_num:.2f}"
                except (ValueError, TypeError):
                    price_display = str(current_price) if current_price else 'N/A'
                
                item_id = tree.insert('', 'end', values=[
                    clean_base_sku,
                    store, 
                    price_display,
                    price_display  # Start with current price as new price
                ])
                
                # Store original data
                original_data[item_id] = {
                    'base_sku': base_sku,  # Keep full base SKU with store
                    'clean_base_sku': clean_base_sku,
                    'store': store,
                    'current_price': current_price,
                    'new_price': current_price
                }
        
        # Search functionality
        def on_search(*args):
            populate_tree(search_var.get())
        
        search_var.trace_add("write", on_search)
        
        # Edit functionality
        def on_double_click(event):
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            
            if not item or column != '#4':  # Only allow editing the "New Base Price" column
                return
            
            # Get current values
            current_values = tree.item(item, 'values')
            current_new_price = current_values[3]
            
            # Remove $ and convert to number for editing
            if current_new_price.startswith('$'):
                edit_value = current_new_price[1:]
            else:
                edit_value = current_new_price
            
            # Create edit entry
            bbox = tree.bbox(item, column)
            if not bbox:
                return
            
            edit_entry = tk.Entry(tree)
            edit_entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
            edit_entry.insert(0, edit_value)
            edit_entry.select_range(0, tk.END)
            edit_entry.focus_set()
            
            def save_edit():
                try:
                    new_value = edit_entry.get().strip()
                    if new_value:
                        # Validate it's a number
                        price_num = float(new_value)
                        formatted_price = f"${price_num:.2f}"
                        
                        # Update tree display
                        values = list(current_values)
                        values[3] = formatted_price
                        tree.item(item, values=values)
                        
                        # Update stored data
                        if item in original_data:
                            original_data[item]['new_price'] = str(price_num)
                            modified_data[item] = original_data[item]
                    
                    edit_entry.destroy()
                except ValueError:
                    messagebox.showerror("Invalid Price", "Please enter a valid number.")
                    edit_entry.focus_set()
            
            def cancel_edit():
                edit_entry.destroy()
            
            edit_entry.bind('<Return>', lambda e: save_edit())
            edit_entry.bind('<Escape>', lambda e: cancel_edit())
            edit_entry.bind('<FocusOut>', lambda e: save_edit())
        
        tree.bind('<Double-1>', on_double_click)
        
        # Buttons
        button_frame = ttk.Frame(bulk_win)
        button_frame.pack(fill='x', padx=10, pady=(0,10))
        
        def save_all_changes():
            if not modified_data:
                messagebox.showinfo("No Changes", "No changes to save.")
                return
            
            # Confirm with user
            change_count = len(modified_data)
            if not messagebox.askyesno("Confirm Changes", 
                f"This will update base prices for {change_count} base SKU(s). "
                "All pricing will be recalculated. Continue?"):
                return
            
            try:
                # Apply changes to database
                for item_id, data in modified_data.items():
                    base_sku = data['base_sku']
                    new_price = data['new_price']
                    
                    if base_sku in db:
                        db[base_sku]['base_price'] = new_price
                
                # Save database
                with open(db_path, "w", encoding="utf-8") as f:
                    json.dump(db, f, indent=2)
                
                # Update current session if we're editing the current base SKU
                current_base_sku = self.base_sku_entry.get().strip()
                if current_base_sku:
                    for data in modified_data.values():
                        clean_sku = data['clean_base_sku']
                        store = data['store']
                        full_sku = f"{clean_sku} ({store})" if store else clean_sku
                        
                        if full_sku == current_base_sku or clean_sku == current_base_sku:
                            self.base_price_entry.delete(0, 'end')
                            self.base_price_entry.insert(0, data['new_price'])
                            break
                
                # Recalculate all pricing
                self.recalculate_all_pricing()
                
                # Auto-save to cloud
                self._ctrl_shift_s()
                
                messagebox.showinfo("Success", 
                    f"Successfully updated {change_count} base price(s). "
                    "All pricing has been recalculated and saved to cloud.")
                
                bulk_win.destroy()
                
            except Exception as e:
                messagebox.showerror("Save Error", f"Error saving changes:\n{e}")
        
        def cancel_changes():
            if modified_data:
                if messagebox.askyesno("Unsaved Changes", 
                    "You have unsaved changes. Are you sure you want to cancel?"):
                    bulk_win.destroy()
            else:
                bulk_win.destroy()
        
        ttk.Button(button_frame, text="Save All Changes", command=save_all_changes).pack(side='right', padx=(5,0))
        ttk.Button(button_frame, text="Cancel", command=cancel_changes).pack(side='right')
        
        # Add bulk operations
        bulk_ops_frame = ttk.Frame(button_frame)
        bulk_ops_frame.pack(side='left')
        
        def apply_percentage_change():
            percentage_str = percentage_entry.get().strip()
            if not percentage_str:
                return
            
            try:
                percentage = float(percentage_str)
                
                # Apply to all visible items
                for item in tree.get_children():
                    data = original_data[item]
                    current_price = float(data['current_price']) if data['current_price'] else 0
                    new_price = current_price * (1 + percentage / 100)
                    
                    # Update display
                    values = list(tree.item(item, 'values'))
                    values[3] = f"${new_price:.2f}"
                    tree.item(item, values=values)
                    
                    # Update stored data
                    data['new_price'] = str(new_price)
                    modified_data[item] = data
                
                messagebox.showinfo("Applied", f"Applied {percentage:+.1f}% change to all visible items.")
            
            except ValueError:
                messagebox.showerror("Invalid Percentage", "Please enter a valid percentage number.")
        
        ttk.Label(bulk_ops_frame, text="Apply % change to all:").pack(side='left')
        percentage_entry = ttk.Entry(bulk_ops_frame, width=8)
        percentage_entry.pack(side='left', padx=2)
        ttk.Button(bulk_ops_frame, text="Apply", command=apply_percentage_change).pack(side='left', padx=2)
        
        # Initial population
        populate_tree()
        
        # Set focus to search
        search_entry.focus_set()

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
        # Handle unsaved changes first
        if self.unsaved_changes:
            res = messagebox.askyesnocancel("Unsaved Changes", "You have unsaved changes. Save before quitting?")
            if res is None:
                return  # Cancel quit
            elif res:
                self.save_to_database(temp=True)
        
        # Ask about cloud upload only if not in test mode
        if not isTest:
            upload_choice = messagebox.askyesnocancel(
                "Upload to Cloud", 
                "Do you want to upload your databases to cloud storage before closing?\n\n" +
                "• Yes: Upload then close\n" +
                "• No: Close without uploading\n" +
                "• Cancel: Don't close",
                icon='question'
            )
            
            if upload_choice is None:
                return  # Cancel quit - don't close
            elif upload_choice:
                # User wants to upload - start upload and wait for completion
                self._quit_after_upload()
                return  # Don't continue with immediate quit
            # If upload_choice is False, skip upload and proceed to quit
        
        # Immediate quit (no upload or test mode)
        self._shutdown_and_quit()

    def _quit_after_upload(self):
        """Start upload process and quit only after it completes"""
        # Show upload dialog and start upload
        upload_dialog, status_labels = show_upload_status_dialog(self)

        # Keep track of scheduled after IDs so we can cancel them before destroying the dialog
        upload_after_ids = []

        def set_status(idx, value):
            """Update a status label safely on the dialog and record the after id."""
            def update_ui():
                try:
                    if upload_dialog.winfo_exists():
                        status_labels[idx]['text'] = value
                        if value == '✓':
                            status_labels[idx]['foreground'] = 'green'
                        elif value == 'X':
                            status_labels[idx]['foreground'] = 'red'
                        else:
                            status_labels[idx]['foreground'] = 'orange'
                        upload_dialog.update_idletasks()
                except Exception:
                    # Widget likely destroyed or gone; ignore
                    pass

            try:
                aid = upload_dialog.after(0, update_ui)
                upload_after_ids.append(aid)
            except Exception:
                pass

        def schedule_on_dialog(callable_fn):
            try:
                aid = upload_dialog.after(0, callable_fn)
                upload_after_ids.append(aid)
                return aid
            except Exception:
                return None

        def spin_status(idx, stop_event):
            import itertools
            spinner_cycle = itertools.cycle(['⏳', '🔄', '⏳', '🔄'])
            while not stop_event.is_set():
                if stop_event.wait(0.3):
                    break

                def update_spinner():
                    try:
                        if upload_dialog.winfo_exists() and not stop_event.is_set():
                            status_labels[idx]['text'] = next(spinner_cycle)
                            status_labels[idx]['foreground'] = 'orange'
                            upload_dialog.update_idletasks()
                    except Exception:
                        # Widget likely destroyed, ignore
                        pass

                # Schedule the spinner update and record the after id
                schedule_on_dialog(update_spinner)

        def do_upload_then_quit():
            import threading
            success = True
            try:
                ensure_web_service_running()
                data_dir = get_data_dir()
                dt = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

                # Step 1: Upload main database
                spin1 = threading.Event()
                t1 = threading.Thread(target=spin_status, args=(0, spin1), daemon=True)
                t1.start()

                local_path = get_latest_database_path()
                drive_filename = f"sku_database_{dt}.json"
                print(f"DEBUG: Uploading database file: {local_path}")

                resp = requests.post('http://localhost:5000/upload', json={
                    'file_path': local_path,
                    'drive_filename': drive_filename
                })
                try:
                    data = resp.json()
                except Exception:
                    data = {}
                # Log server response for debugging
                try:
                    print(f"DEBUG: upload resp status={resp.status_code}, body={resp.text}")
                except Exception:
                    pass

                spin1.set()
                t1.join(timeout=1)

                if data.get('status') == 'success':
                    set_status(0, '✓')
                else:
                    set_status(0, 'X')
                    success = False

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
                    try:
                        data = resp.json()
                    except Exception:
                        data = {}
                    try:
                        print(f"DEBUG: upload resp status={resp.status_code}, body={resp.text}")
                    except Exception:
                        pass

                    spin2.set()
                    t2.join(timeout=1)

                    if data.get('status') == 'success':
                        set_status(1, '✓')
                    else:
                        set_status(1, 'X')
                        success = False
                else:
                    spin2.set()
                    t2.join(timeout=1)
                    set_status(1, '✓')

                # Brief pause to show final status
                time.sleep(1.0)

            except Exception as e:
                success = False
                print(f"Upload error: {e}")

            finally:
                # Schedule the final cleanup and quit in the main thread
                def finish_and_quit():
                    # allow reassigning dialog and labels for retry
                    nonlocal upload_dialog, status_labels, upload_after_ids
                    try:
                        # Cancel any pending after callbacks scheduled on the dialog
                        try:
                            for aid in upload_after_ids:
                                try:
                                    upload_dialog.after_cancel(aid)
                                except Exception:
                                    pass
                        except Exception:
                            pass

                        # Destroy the dialog now that spinners are cancelled
                        if upload_dialog.winfo_exists():
                            upload_dialog.destroy()
                    except Exception:
                        pass

                    # If upload succeeded, quit; otherwise prompt the user
                    if success:
                        try:
                            self._shutdown_and_quit()
                        except Exception:
                            try:
                                self.quit()
                            except Exception:
                                pass
                        return

                    # Upload failed: let the user decide (Retry / Close Anyway / Cancel)
                    try:
                        # modal choice: Retry, Close Anyway, Cancel
                        choice = messagebox.askretrycancel(
                            "Upload Failed",
                            "Uploading failed. Retry upload? (Cancel will keep the app open)",
                            parent=self
                        )
                    except Exception:
                        # If parent/dialog gone or messagebox fails, default to not closing
                        choice = True  # Treat as Retry

                    if choice is True:
                        # User selected Retry -> start another background upload using a fresh dialog
                        try:
                            # Recreate dialog and status labels and reset after-ids tracker
                            upload_dialog, status_labels = show_upload_status_dialog(self)
                            upload_after_ids = []
                            import threading
                            # Small delay to allow dialog to initialize then start upload thread
                            def restart():
                                try:
                                    time.sleep(0.2)
                                except Exception:
                                    pass
                                t = threading.Thread(target=do_upload_then_quit, daemon=True)
                                t.start()
                            try:
                                upload_dialog.after(200, restart)
                            except Exception:
                                restart()
                        except Exception:
                            # If we can't recreate dialog, keep app open
                            pass
                        return
                    else:
                        # User selected Cancel/CloseAnyway (askretrycancel returns False for Cancel)
                        # If they explicitly chose to not retry, ask to Close Anyway
                        try:
                            close_any = messagebox.askyesno(
                                "Close Anyway?",
                                "Do you want to close the application anyway?",
                                parent=self
                            )
                        except Exception:
                            close_any = False

                        if close_any:
                            try:
                                self._shutdown_and_quit()
                            except Exception:
                                try:
                                    self.quit()
                                except Exception:
                                    pass
                            return
                        else:
                            # Keep the application open; do not quit
                            return

                try:
                    upload_dialog.after(500, finish_and_quit)
                except Exception:
                    # If scheduling on dialog fails, run immediately on main thread
                    finish_and_quit()

        # Start upload in background thread
        upload_thread = threading.Thread(target=do_upload_then_quit, daemon=True)
        upload_thread.start()

    def _shutdown_and_quit(self):
        """Shutdown web service and quit application"""
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
        self.quit()  # Exit the Tkinter main loop
        self.destroy()  # Destroy the window

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

    def copy_options_from_base_sku(self):
        """Copy option costs, weights, and associated SKUs from another base SKU"""
        # First check if we have any options to copy to
        if not hasattr(self, 'master_df') or self.master_df is None or self.master_df.empty:
            messagebox.showwarning("No Options", "Load a product configuration first.")
            return
        
        # Get current base SKU to exclude from selection
        current_base_sku = self.base_sku_entry.get().strip() or self.base_sku
        
        # Load database to get available base SKUs
        db_path = self._get_latest_database_path()
        if not os.path.exists(db_path):
            messagebox.showinfo("No Database", "No database file found.")
            return
        
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except json.JSONDecodeError:
            messagebox.showerror("Database Error", "Database file is corrupted.")
            return
        
        # Get available base SKUs (excluding current one)
        available_skus = [sku for sku in db.keys() if sku != current_base_sku]
        if not available_skus:
            messagebox.showinfo("No Options", "No other base SKUs available to copy from.")
            return
        
        # Create dialog to select source base SKU (reuse the search dialog pattern)
        class SourceSkuDialog(tk.Toplevel):
            def __init__(self, parent, skus):
                super().__init__(parent)
                self.title("Select Base SKU to Copy Options From")
                self.geometry("450x400")
                self.selected = None
                self.all_skus = sorted(skus)
                
                # Make dialog modal
                self.transient(parent)
                self.grab_set()
                
                tk.Label(self, text="Search for source base SKU:").pack(anchor='w', padx=10, pady=(10,0))
                self.var = tk.StringVar()
                self.var.trace_add('write', self.update_list)
                entry = tk.Entry(self, textvariable=self.var)
                entry.pack(fill='x', padx=10, pady=(0,10))
                
                # Add KeyRelease binding as backup
                def on_key_release(event):
                    self.update_list(entry)
                entry.bind('<KeyRelease>', on_key_release)
                self.search_entry = entry
                
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
                tk.Button(btn_frame, text="Copy Options", command=self.select).pack(side='right', padx=(5,0))
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
                
            def update_list(self, entry_widget=None, *args):
                if entry_widget:
                    search_text = entry_widget.get().lower()
                else:
                    search_text = self.var.get().lower()
                if not search_text:
                    self.populate_listbox(self.all_skus)
                else:
                    filtered_skus = [sku for sku in self.all_skus if search_text in sku.lower()]
                    self.populate_listbox(filtered_skus)
                        
            def select(self, event=None):
                selection = self.listbox.curselection()
                if selection:
                    self.selected = self.listbox.get(selection[0])
                    self.destroy()
        
        # Show dialog
        dlg = SourceSkuDialog(self, available_skus)
        self.wait_window(dlg)
        source_sku = dlg.selected
        
        if not source_sku:
            return
        
        # Load source SKU data
        source_state = db[source_sku]
        source_master_df = pd.DataFrame(source_state.get('master_df', []))
        
        if source_master_df.empty:
            messagebox.showinfo("No Options", f"Source base SKU '{source_sku}' has no option data.")
            return
        
        # Find matching options and copy data
        copied_count = 0
        skipped_count = 0
        
        for idx, current_row in self.master_df.iterrows():
            current_name = current_row['Name']
            current_value = current_row['Value']
            
            # Find matching option in source
            matches = source_master_df[
                (source_master_df['Name'] == current_name) & 
                (source_master_df['Value'] == current_value)
            ]
            
            if not matches.empty:
                source_row = matches.iloc[0]
                
                # Copy the data
                self.master_df.at[idx, 'Add\'l Cost'] = source_row.get('Add\'l Cost', '0')
                self.master_df.at[idx, 'Add\'l Weight'] = source_row.get('Add\'l Weight', '0') 
                self.master_df.at[idx, 'Associated SKUs'] = source_row.get('Associated SKUs', '')
                
                copied_count += 1
            else:
                skipped_count += 1
        
        # Update UI
        self._fill_tree(self.out_tree, self.master_df)
        self.populate_value_grid()
        self._regenerate_left_preview()
        self._mark_unsaved()
        
        # Show results
        if copied_count > 0:
            message = f"Successfully copied option data for {copied_count} matching options from '{source_sku}'."
            if skipped_count > 0:
                message += f"\n\n{skipped_count} options were skipped (no matching Name+Value found in source)."
            messagebox.showinfo("Copy Complete", message)
        else:
            messagebox.showinfo("No Matches", f"No matching options found between current product and '{source_sku}'.")

    def update_option_costs_from_cost_db(self):
        # Check if master_df is loaded
        if self.master_df is None:
            return
            
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
        
        # Ensure at least one empty row if no SKUs exist
        if not sku_list:
            sku_list = [""]
        
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
                    
                    # Filter the dropdown list based on entry text (search both name and price)
                    dropdown.delete(0, tk.END)
                    matches = []
                    
                    for item in part_numbers:
                        item_price = str(cost_db.get(item, 0))
                        # Check if search value matches item name or price
                        if (value in item.lower() or 
                            value in item_price.lower()):
                            matches.append(item)
                    
                    # Display matches with prices in parentheses
                    for match in matches:
                        price = cost_db.get(match, 0)
                        display_text = f"{match} (${price})"
                        dropdown.insert(tk.END, display_text)
                    
                    if matches:
                        # Position dropdown below entry
                        x = entry.winfo_rootx() - editor.winfo_rootx()
                        y = entry.winfo_rooty() - editor.winfo_rooty() + entry.winfo_height()
                        dropdown.place(x=x, y=y, width=max(entry.winfo_width(), 300))  # Make dropdown wider to show prices
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
                        display_value = dropdown.get(dropdown.curselection())
                        # Extract just the SKU name (before the price in parentheses)
                        sku_name = display_value.split(' (')[0] if ' (' in display_value else display_value
                        entry.delete(0, tk.END)
                        entry.insert(0, sku_name)
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
                self.populate_value_grid()  # Refresh the configure panel
                self._on_in_tree_select(None)  # Refresh the breakdown panel
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
            self.populate_value_grid()  # Refresh the configure panel
            self._on_in_tree_select(None)  # Refresh the breakdown panel
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
                self.recalculate_all_pricing()
            win.destroy()
    
        btn_frame = ttk.Frame(frm)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Save", command=save_prices).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side='left', padx=5)
    
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
            if not self._safe_write_json(self.options, config_path):
                print("Error saving options: Safe write failed")
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
                        # Clean the base SKU (remove parentheses and spaces)
                        base_clean = re.sub(r'\([^)]*\)', '', base)
                        base_clean = base_clean.replace(' ', '')
                        base_clean = base_clean.strip()
                        
                        # Get prefix if available
                        prefix = state.get('prefix', '')
                        zero_removal = 0
                        if prefix and prefix.startswith('-') and len(prefix) > 1 and prefix[1:].isdigit():
                            zero_removal = int(prefix[1:])
                            prefix = ""
                        
                        for idx2 in range(len(sku_df)):
                            # Calculate padding based on zero removal
                            padding = max(1, 4 - zero_removal)
                            number_str = str(idx2 + 1).zfill(padding)
                            
                            if prefix:
                                new_sku = f"{base_clean}-{prefix}{number_str}"
                            else:
                                new_sku = f"{base_clean}-{number_str}"
                            
                            sku_df.at[idx2, 'New SKU'] = new_sku
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
                    try:
                        with open(db_path, "r", encoding="utf-8") as f:
                            latest_db = json.load(f)
                    except (json.JSONDecodeError, UnicodeDecodeError) as e:
                        # Try to repair the JSON file
                        latest_db = try_repair_json_file(db_path)
                        if latest_db is None:
                            messagebox.showerror("Database Error", 
                                f"Database file is corrupted and cannot be repaired.\n"
                                f"Error: {str(e)}\n\n"
                                "Using in-memory data instead. Please save and upload to cloud to fix the database.")
                            latest_db = db  # fallback to in-memory data
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
        ttk.Label(loading_win, text="Recalculating all pricing...\nThis should be quick (hopefully).", anchor="center").pack(expand=True, pady=20)
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
        messagebox.showinfo("Whoops...", "Contact Ben at 602 617 3531")
        # from tkinter import filedialog, messagebox
        # # Let user pick a backup file
        # db_path = filedialog.askopenfilename(
        #     title="Select Database Backup",
        #     filetypes=[("JSON Database", "*.json")],
        #     initialdir = get_data_dir()
        # )
        # if not db_path:
        #     return
        # # Confirm with user
        # if not messagebox.askyesno("Revert to Backup", f"Load and set this backup as the current database?\n\n{db_path}"):
        #     return
        # # Copy selected file to be the latest database
        # folder = self._get_database_folder()
        # latest_path = os.path.join(folder, "sku_database_temp.json")
        # import shutil
        # try:
        #     shutil.copy2(db_path, latest_path)
        #     # Set as current in config (if you track a current path, update it here)
        #     self.current_prog_path = latest_path
        #     self.unsaved_changes = False
        #     self._update_title()
        #     messagebox.showinfo("Reverted", f"Reverted to backup:\n{db_path}")
        #     # Reload UI from the reverted database
        #     self.load_from_database()
        # except Exception as e:
        #     messagebox.showerror("Error", f"Failed to revert to backup:\n{e}")
    
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
import sys
import threading

def ensure_web_service_running():
    """Ensure the helper Flask web service is running.

    When running from a PyInstaller bundle (frozen), there is no standalone
    Python script on disk to spawn. In that case import the web service
    module and run it in a background thread. Otherwise spawn the script
    using the local Python interpreter.
    """
    global web_service_process
    try:
        requests.get('http://localhost:5000')
        return
    except Exception:
        pass

    # Not running, so start it.
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle — run the Flask app in a background thread
        try:
            # Import the module packaged inside the bundle
            import ap_sku_tool_web_service as websvc

            def _run_service():
                try:
                    # Use 127.0.0.1 for predictable binding
                    websvc.app.run(host='127.0.0.1', port=5000, threaded=True)
                except Exception as e:
                    print("Web service failed to start inside bundle:", e)

            t = threading.Thread(target=_run_service, daemon=True)
            t.start()
            web_service_process = None
        except Exception as e:
            print("Failed to start internal web service:", e)
    else:
        # Normal (non-frozen) execution: spawn the helper script as a separate process
        service_path = os.path.join(os.path.dirname(__file__), "ap_sku_tool_web_service.py")
        web_service_process = subprocess.Popen([sys.executable, service_path])

    # Wait for service to start (give it a few seconds)
    for _ in range(10):
        try:
            requests.get('http://localhost:5000')
            break
        except Exception:
            time.sleep(1)

def check_database_integrity():
    """Check if current database files are corrupted and offer backup restore."""
    data_dir = get_data_dir()
    
    issues = []
    
    # Check SKU database
    sku_db_path = os.path.join(data_dir, "sku_database_temp.json")
    if os.path.exists(sku_db_path):
        try:
            with open(sku_db_path, 'r', encoding='utf-8') as f:
                json.load(f)
        except (json.JSONDecodeError, UnicodeDecodeError):
            issues.append("SKU database")
    
    # Check cost database
    cost_db_path = os.path.join(data_dir, "cost_db.json")
    if os.path.exists(cost_db_path):
        try:
            with open(cost_db_path, 'r', encoding='utf-8') as f:
                json.load(f)
        except (json.JSONDecodeError, UnicodeDecodeError):
            issues.append("Cost database")
    
    if issues:
        issues_text = " and ".join(issues)
        result = messagebox.askyesno(
            "Database Corruption Detected",
            f"Current {issues_text} file(s) are corrupted.\n\n"
            f"Would you like to restore from your latest uncorrupted local backup?"
        )
        
        if result:
            restore_from_local_backup()
    
    return len(issues) == 0

def format_timestamp_for_display(timestamp_str):
    """Convert timestamp from filename format to human-readable format."""
    try:
        # Handle format: 2025-07-16_15-40-22
        if len(timestamp_str) == 19 and timestamp_str.count('-') == 4 and timestamp_str.count('_') == 1:
            # Parse: YYYY-MM-DD_HH-MM-SS
            date_part, time_part = timestamp_str.split('_')
            year, month, day = date_part.split('-')
            hour, minute, second = time_part.split('-')
            
            dt = datetime.datetime(int(year), int(month), int(day), int(hour), int(minute), int(second))
            return dt.strftime("%B %d, %Y at %I:%M:%S %p")
        
        # Handle other possible formats as fallback
        else:
            return f"File: {timestamp_str}"
            
    except (ValueError, IndexError):
        return f"File: {timestamp_str}"

def find_latest_uncorrupted_databases():
    """Find the newest pair of non-corrupted SKU and cost databases from local files."""
    data_dir = get_data_dir()
    if not os.path.exists(data_dir):
        return None, None
    
    # Get all timestamps from both SKU and cost databases
    sku_timestamps = {}
    cost_timestamps = {}
    
    # Collect all SKU database files and their timestamps
    for filename in os.listdir(data_dir):
        if filename.startswith('sku_database_') and filename.endswith('.json'):
            # Skip temp files
            if '_temp' in filename or filename.endswith('_temp.json'):
                continue
                
            try:
                timestamp_part = filename.replace('sku_database_', '').replace('.json', '')
                if timestamp_part and timestamp_part != 'temp':
                    sku_timestamps[timestamp_part] = filename
            except:
                continue
    
    # Collect all cost database files and their timestamps
    for filename in os.listdir(data_dir):
        if filename.startswith('cost_db_') and filename.endswith('.json'):
            # Skip temp files
            if '_temp' in filename or filename.endswith('_temp.json'):
                continue
                
            try:
                timestamp_part = filename.replace('cost_db_', '').replace('.json', '')
                if timestamp_part and timestamp_part != 'temp':
                    cost_timestamps[timestamp_part] = filename
            except:
                continue
    
    # Find matching timestamps, sorted by newest first
    matching_timestamps = sorted(
        set(sku_timestamps.keys()) & set(cost_timestamps.keys()), 
        reverse=True
    )
    
    def validate_database_file(file_path):
        """Validate that a database file exists, is not empty, and contains valid data structure."""
        try:
            if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
                return False
                
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            # Check if data is empty
            if not data or (isinstance(data, (dict, list)) and len(data) == 0):
                return False
            
            # Additional validation based on file type
            filename = os.path.basename(file_path)
            
            if filename.startswith('sku_database_'):
                # SKU database should be a dict with SKU entries
                if not isinstance(data, dict):
                    return False
                
                # Check if at least one entry looks like valid SKU data
                for sku_name, sku_data in data.items():
                    if not isinstance(sku_data, dict):
                        continue
                    
                    # Valid SKU data should have these keys
                    required_keys = ['input_df', 'master_df', 'base_price', 'base_weight']
                    if any(key in sku_data for key in required_keys):
                        # Found at least one valid-looking SKU entry
                        return True
                
                # No valid SKU entries found
                return False
                
            elif filename.startswith('cost_db_'):
                # Cost database should be a dict with part numbers and prices
                if not isinstance(data, dict):
                    return False
                
                # Check if at least one entry looks like valid cost data
                for part_num, cost_data in data.items():
                    if isinstance(cost_data, (int, float)) and cost_data > 0:
                        # Found at least one valid cost entry
                        return True
                    elif isinstance(cost_data, dict) and 'cost' in cost_data:
                        # Alternative cost data structure
                        return True
                
                # No valid cost entries found
                return False
            
            # Unknown file type, just check it's not empty
            return True
                
        except (json.JSONDecodeError, UnicodeDecodeError, OSError, FileNotFoundError, ValueError):
            return False
    
    # Go through matching timestamps from newest to oldest until we find a valid pair
    for timestamp in matching_timestamps:
        sku_filename = sku_timestamps[timestamp]
        cost_filename = cost_timestamps[timestamp]
        
        sku_path = os.path.join(data_dir, sku_filename)
        cost_path = os.path.join(data_dir, cost_filename)
        
        # Validate both files
        if validate_database_file(sku_path) and validate_database_file(cost_path):
            print(f"Found valid database pair: {sku_filename} and {cost_filename}")
            return sku_path, cost_path
        else:
            print(f"Skipping corrupted database pair: {sku_filename} and {cost_filename}")
            continue
    
    # If no matching pairs found, try to find just the newest valid SKU database
    # and any valid cost database (even if timestamps don't match)
    print("No matching timestamp pairs found, looking for any valid databases...")
    
    # Find newest valid SKU database
    valid_sku_path = None
    for timestamp in sorted(sku_timestamps.keys(), reverse=True):
        sku_path = os.path.join(data_dir, sku_timestamps[timestamp])
        if validate_database_file(sku_path):
            valid_sku_path = sku_path
            break
    
    # Find newest valid cost database
    valid_cost_path = None
    for timestamp in sorted(cost_timestamps.keys(), reverse=True):
        cost_path = os.path.join(data_dir, cost_timestamps[timestamp])
        if validate_database_file(cost_path):
            valid_cost_path = cost_path
            break
    
    if valid_sku_path:
        print(f"Found valid SKU database: {os.path.basename(valid_sku_path)}")
        if valid_cost_path:
            print(f"Found valid cost database: {os.path.basename(valid_cost_path)}")
        else:
            print("No valid cost database found")
    else:
        print("No valid databases found at all")
    
    return valid_sku_path, valid_cost_path

def restore_from_local_backup():
    """Restore databases from local backup to cloud."""
    sku_db_path, cost_db_path = find_latest_uncorrupted_databases()
    
    if not sku_db_path:
        messagebox.showerror("No Backup Found", 
            "No uncorrupted local database backups were found.\n"
            "Cannot restore from local files.")
        return False
    
    # Show what we found with human-readable timestamps
    sku_name = os.path.basename(sku_db_path)
    cost_name = os.path.basename(cost_db_path) if cost_db_path else "None found"
    
    # Extract and format timestamps for display
    sku_timestamp = sku_name.replace('sku_database_', '').replace('.json', '')
    sku_date = format_timestamp_for_display(sku_timestamp)
    
    cost_date = "N/A"
    if cost_db_path:
        cost_timestamp = cost_name.replace('cost_db_', '').replace('.json', '')
        cost_date = format_timestamp_for_display(cost_timestamp)
    
    result = messagebox.askyesno("Restore from Local Backup", 
        f"Found latest uncorrupted local databases:\n\n"
        f"SKU Database:\n"
        f"  File: {sku_name}\n"
        f"  Created: {sku_date}\n\n"
        f"Cost Database:\n"
        f"  File: {cost_name}\n"
        f"  Created: {cost_date}\n\n"
        f"This will upload these files to the cloud, overwriting any corrupted data.\n"
        f"Do you want to proceed with the restore?")
    
    if not result:
        return False
    
    try:
        # Double-check the backup files before proceeding
        # Validate SKU database file
        try:
            if not os.path.exists(sku_db_path) or os.path.getsize(sku_db_path) == 0:
                raise ValueError("SKU backup file is missing or empty")
            
            with open(sku_db_path, 'r', encoding='utf-8') as f:
                sku_data = json.load(f)
                if not sku_data:
                    raise ValueError("SKU backup file contains no data")
        except (json.JSONDecodeError, UnicodeDecodeError, OSError) as e:
            messagebox.showerror("Backup Validation Failed", 
                f"The selected SKU backup file is corrupted:\n{str(e)}\n\n"
                f"Cannot proceed with restore.")
            return False
        
        # Validate cost database file if present
        if cost_db_path:
            try:
                if not os.path.exists(cost_db_path) or os.path.getsize(cost_db_path) == 0:
                    cost_db_path = None  # Skip cost DB if invalid
                else:
                    with open(cost_db_path, 'r', encoding='utf-8') as f:
                        cost_data = json.load(f)
                        if not cost_data:
                            cost_db_path = None  # Skip cost DB if empty
            except (json.JSONDecodeError, UnicodeDecodeError, OSError):
                cost_db_path = None  # Skip cost DB if corrupted
        
        # Create timestamped copies for upload
        data_dir = get_data_dir()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        
        # Copy SKU database
        restore_sku_path = os.path.join(data_dir, f"sku_database_{timestamp}.json")
        import shutil
        shutil.copy2(sku_db_path, restore_sku_path)
        
        # Copy cost database if available
        restore_cost_path = None
        if cost_db_path:
            restore_cost_path = os.path.join(data_dir, f"cost_db_{timestamp}.json")
            shutil.copy2(cost_db_path, restore_cost_path)
        
        # Upload to cloud
        ensure_web_service_running()
        
        # Upload SKU database
        resp = requests.post('http://localhost:5000/push_db', json={
            'local_path': restore_sku_path
        })
        
        if resp.json().get('status') != 'success':
            messagebox.showerror("Upload Failed", 
                f"Failed to upload SKU database: {resp.json().get('message', 'Unknown error')}")
            return False
        
        # Upload cost database
        if restore_cost_path:
            resp = requests.post('http://localhost:5000/push_db', json={
                'local_path': restore_cost_path
            })
            
            if resp.json().get('status') != 'success':
                messagebox.showwarning("Partial Upload", 
                    f"SKU database uploaded successfully, but cost database failed: {resp.json().get('message', 'Unknown error')}")
        
        messagebox.showinfo("Restore Complete", 
            "Local backup databases have been successfully restored to the cloud.\n"
            "The corrupted cloud data has been replaced with your latest uncorrupted local files.")
        
        return True
        
    except Exception as e:
        messagebox.showerror("Restore Failed", 
            f"Error during restore process: {str(e)}")
        return False

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
    
    # Test if downloaded SKU database is corrupted
    sku_corrupted = False
    try:
        # Check file size first
        if not os.path.exists(dst) or os.path.getsize(dst) == 0:
            sku_corrupted = True
            print(f"Downloaded SKU database is empty or missing")
        else:
            with open(dst, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if not data:  # Check if data is empty dict/list
                    sku_corrupted = True
                    print(f"Downloaded SKU database contains no data")
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        sku_corrupted = True
        print(f"Downloaded SKU database is corrupted: {e}")

    # Download cost database
    cost_corrupted = False
    resp = requests.post('http://localhost:5000/pull_latest_db', json={
        'prefix': 'cost_db',
        'suffix': '.json'
    })
    data = resp.json()
    if data.get('status') == 'success':
        src = data['local_path']
        dst_cost = os.path.join(data_dir, "cost_db.json")
        shutil.copy2(src, dst_cost)
        
        # Test if downloaded cost database is corrupted
        try:
            # Check file size first
            if not os.path.exists(dst_cost) or os.path.getsize(dst_cost) == 0:
                cost_corrupted = True
                print(f"Downloaded cost database is empty or missing")
            else:
                with open(dst_cost, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if not data:  # Check if data is empty dict/list
                        cost_corrupted = True
                        print(f"Downloaded cost database contains no data")
        except (json.JSONDecodeError, UnicodeDecodeError) as e:
            cost_corrupted = True
            print(f"Downloaded cost database is corrupted: {e}")
    else:
        messagebox.showwarning(
            "Pricing Database Not Updated",
            "No pricing database was found in the cloud. The program will use the existing local cost_db.json (if any)."
        )

    # If either database is corrupted, offer to restore from local backup
    if sku_corrupted or cost_corrupted:
        corruption_msg = []
        if sku_corrupted:
            corruption_msg.append("SKU database")
        if cost_corrupted:
            corruption_msg.append("Cost database")
        
        corruption_text = " and ".join(corruption_msg)
        
        result = messagebox.askyesnocancel(
            "Database Corruption Detected",
            f"The downloaded {corruption_text} from the cloud is corrupted.\n\n"
            f"Would you like to restore from your latest uncorrupted local backup?\n\n"
            f"• YES: Restore from local backup to cloud\n"
            f"• NO: Continue with corrupted data (not recommended)\n"
            f"• CANCEL: Abort download process"
        )
        
        if result is None:  # Cancel
            return False
        elif result:  # Yes - restore from backup
            if restore_from_local_backup():
                # After successful restore, re-download the restored data
                return pull_database_from_cloud()
            else:
                return False
        else:  # No - continue with corrupted data
            messagebox.showwarning("Using Corrupted Data", 
                "Proceeding with corrupted database files. "
                "Data integrity issues may occur. "
                "Consider restoring from backup soon.")

    return True

def get_latest_database_path():
    """Standalone function to get the latest database file path"""
    folder = get_data_dir()
        
    # Get all database files (including timestamped temp files)
    files = [f for f in os.listdir(folder) if f.startswith("sku_database") and f.endswith(".json")]
    if not files:
        return os.path.join(folder, "sku_database.json")
        
    # Sort by modification time (newest first) to get the most recent file
    file_paths = [(f, os.path.getmtime(os.path.join(folder, f))) for f in files]
    file_paths.sort(key=lambda x: x[1], reverse=True)
    
    latest_file = os.path.join(folder, file_paths[0][0])
    return latest_file

def push_database_to_cloud():
    import threading
    import itertools

    # Get the main window (should be OptionsParserApp instance)
    main_window = None
    for widget in tk._default_root.winfo_children():
        if isinstance(widget, tk.Toplevel) or hasattr(widget, 'title'):
            main_window = widget
            break
    
    if main_window is None:
        main_window = tk._default_root

    splash, status_labels = show_upload_status_dialog(main_window)
    
    def set_status(idx, value):
        # Schedule UI updates in the main thread
        def update_ui():
            if splash.winfo_exists():
                status_labels[idx]['text'] = value
                if value == '✓':
                    status_labels[idx]['foreground'] = 'green'
                elif value == 'X':
                    status_labels[idx]['foreground'] = 'red'
                else:
                    status_labels[idx]['foreground'] = 'orange'
                splash.update_idletasks()
        
        splash.after(0, update_ui)

    def spin_status(idx, stop_event):
        spinner_cycle = itertools.cycle(['⏳', '🔄', '⏳', '🔄'])
        while not stop_event.is_set():
            if stop_event.wait(0.3):  # Check every 300ms
                break
            def update_spinner():
                if splash.winfo_exists() and not stop_event.is_set():
                    status_labels[idx]['text'] = next(spinner_cycle)
                    status_labels[idx]['foreground'] = 'orange'
                    splash.update_idletasks()
            splash.after(0, update_spinner)

    def do_upload():
        success = True
        try:
            ensure_web_service_running()
            data_dir = get_data_dir()
            dt = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

            # Step 1: Upload main database
            spin1 = threading.Event()
            t1 = threading.Thread(target=spin_status, args=(0, spin1), daemon=True)
            t1.start()
            
            local_path = get_latest_database_path()  # Use the latest database file!
            drive_filename = f"sku_database_{dt}.json"
            print(f"DEBUG: Uploading database file: {local_path}")
            
            resp = requests.post('http://localhost:5000/upload', json={
                'file_path': local_path,
                'drive_filename': drive_filename
            })
            data = resp.json()
            
            spin1.set()
            t1.join(timeout=1)  # Wait for spinner to stop
            
            if data.get('status') == 'success':
                set_status(0, '✓')
            else:
                set_status(0, 'X')
                success = False
                splash.after(0, lambda: messagebox.showerror("Upload Error", 
                    data.get('message', 'Unknown error'), parent=splash))

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
                t2.join(timeout=1)  # Wait for spinner to stop
                
                if data.get('status') == 'success':
                    set_status(1, '✓')
                else:
                    set_status(1, 'X')
                    success = False
                    splash.after(0, lambda: messagebox.showerror("Upload Error", 
                        f"Cost DB: {data.get('message', 'Unknown error')}", parent=splash))
            else:
                spin2.set()
                t2.join(timeout=1)
                set_status(1, '✓')
            
            # Brief pause to show final status
            time.sleep(1.0)
            
        except Exception as e:
            success = False
            splash.after(0, lambda: messagebox.showerror("Upload Error", 
                f"Unexpected error: {str(e)}", parent=splash))
        
        finally:
            # Schedule the cleanup and message in the main thread
            def finish_upload():
                try:
                    if splash.winfo_exists():
                        splash.destroy()
                    
                    if success:
                        messagebox.showinfo("Upload Complete", 
                            "Upload process finished successfully.", parent=main_window)
                    else:
                        messagebox.showwarning("Upload Incomplete", 
                            "Upload process finished with errors.", parent=main_window)
                except:
                    pass
            
            splash.after(500, finish_upload)

    # Start upload in background thread
    upload_thread = threading.Thread(target=do_upload, daemon=True)
    upload_thread.start()

def show_upload_status_dialog(parent):
    splash = tk.Toplevel(parent)
    splash.title("Uploading to Cloud...")
    splash.geometry("480x200")
    splash.resizable(False, False)
    
    # Make it truly modal and prevent closing
    splash.transient(parent)
    splash.grab_set()
    splash.protocol("WM_DELETE_WINDOW", lambda: None)  # Disable close button
    splash.focus_set()
    
    # Remove minimize/maximize buttons
    splash.attributes('-toolwindow', True)
    
    # Keep on top during upload
    splash.attributes('-topmost', True)
    
    # Center the dialog
    splash.update_idletasks()
    w = splash.winfo_screenwidth()
    h = splash.winfo_screenheight()
    size = tuple(int(_) for _ in splash.geometry().split('+')[0].split('x'))
    x = w//2 - size[0]//2
    y = h//2 - size[1]//2
    splash.geometry(f"{size[0]}x{size[1]}+{x}+{y}")

    # Main content frame
    main_frame = tk.Frame(splash, bg='white', padx=20, pady=15)
    main_frame.pack(fill='both', expand=True)

    # Title
    title_label = tk.Label(main_frame, text="📤 Uploading to Cloud Storage", 
                          font=('Segoe UI', 14, 'bold'), bg='white', fg='#2c3e50')
    title_label.pack(pady=(0, 5))
    
    # Instruction
    instruction_label = tk.Label(main_frame, 
                               text="Please wait while your databases are uploaded...\nDo not close this window.", 
                               font=('Segoe UI', 9), bg='white', fg='#7f8c8d', justify='center')
    instruction_label.pack(pady=(0, 15))

    # Upload steps
    steps = [
        "📄 Uploading Main Database",
        "💰 Uploading Pricing Database"
    ]
    
    status_labels = []
    for i, step in enumerate(steps):
        # Step frame
        step_frame = tk.Frame(main_frame, bg='white')
        step_frame.pack(fill='x', pady=5)
        
        # Step label
        step_label = tk.Label(step_frame, text=step, anchor='w', width=30,
                             font=('Segoe UI', 10), bg='white', fg='#34495e')
        step_label.pack(side='left', fill='x', expand=True)
        
        # Status indicator
        status = tk.Label(step_frame, text='⏳', width=3, anchor='center', 
                         font=('Segoe UI', 16), bg='white', fg='orange')
        status.pack(side='right')
        
        status_labels.append(status)
    
    # Warning at bottom
    warning_frame = tk.Frame(main_frame, bg='#fff3cd', relief='solid', bd=1)
    warning_frame.pack(fill='x', pady=(15, 0))
    
    warning_label = tk.Label(warning_frame, 
                           text="⚠️ Upload in progress - Please do not close this window",
                           font=('Segoe UI', 9, 'bold'), bg='#fff3cd', fg='#856404', 
                           pady=8)
    warning_label.pack()
    
    # Force display and bring to front
    splash.update()
    splash.lift()
    splash.focus_force()
    
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
    
    # Make it modal but not completely blocking
    splash.transient(root)
    splash.grab_set()
    
    # Center the dialog
    splash.update_idletasks()
    w = splash.winfo_screenwidth()
    h = splash.winfo_screenheight()
    size = tuple(int(_) for _ in splash.geometry().split('+')[0].split('x'))
    x = w//2 - size[0]//2
    y = h//2 - size[1]//2
    splash.geometry(f"{size[0]}x{size[1]}+{x}+{y}")
    
    # Bring to front but don't stay on top
    splash.lift()
    splash.focus_force()

    steps = [
        "Launching Web Service",
        "Syncing SKU Breakouts with Cloud",
        "Syncing Option Pricing with Cloud",
        "Purging Old Database Files from Cloud",
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

def purge_old_cloud_databases():
    import dateutil.parser
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
        # Schedule UI updates in the main thread
        root.after(0, lambda: _update_status(idx, value))
    
    def _update_status(idx, value):
        if splash.winfo_exists():
            status_labels[idx]['text'] = value
            status_labels[idx]['foreground'] = 'black'  # Always black
            splash.update_idletasks()

    def spin_status(idx, stop_event, status_labels):
        spinner_cycle = itertools.cycle(['-', '\\', '|', '/'])
        while not stop_event.is_set():
            root.after(0, lambda v=next(spinner_cycle): _update_status(idx, v))
            stop_event.wait(0.15)

    def start_app():
        # Step 1: Launching Web Service
        spin1 = threading.Event()
        t1 = threading.Thread(target=spin_status, args=(0, spin1, status_labels), daemon=True)
        t1.start()
        ensure_web_service_running()
        spin1.set()
        set_status(0, '✓')

        # Helper: POST JSON with retries and safe JSON decode
        def post_json_with_retries(url, payload, retries=6, backoff=0.6, timeout=5):
            import requests
            import time
            last_exc = None
            for attempt in range(1, retries + 1):
                try:
                    resp = requests.post(url, json=payload, timeout=timeout)
                except requests.exceptions.RequestException as e:
                    last_exc = e
                    if attempt < retries:
                        time.sleep(backoff)
                        backoff *= 1.5
                        continue
                    return {"status": "error", "message": f"Request failed: {e}"}

                # Got a response, try to parse JSON safely
                try:
                    return resp.json()
                except Exception as je:
                    # Return an error dict with some context
                    text = (resp.text[:1000] if resp.text is not None else '')
                    return {"status": "error", "message": f"Invalid JSON response: {je}", "text": text, "code": resp.status_code}

        # Step 2: Syncing SKU Breakouts with Cloud
        spin2 = threading.Event()
        t2 = threading.Thread(target=spin_status, args=(1, spin2, status_labels), daemon=True)
        t2.start()
        # Download main database only
        data = post_json_with_retries('http://localhost:5000/pull_latest_db', {
            'prefix': 'sku_database_',
            'suffix': '.json'
        })
        if isinstance(data, dict) and data.get('status') == 'success':
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
        t3 = threading.Thread(target=spin_status, args=(2, spin3, status_labels), daemon=True)
        t3.start()
        # Download cost database
        data = post_json_with_retries('http://localhost:5000/pull_latest_db', {
            'prefix': 'cost_db',
            'suffix': '.json'
        })
        if isinstance(data, dict) and data.get('status') == 'success':
            src = data['local_path']
            dst = os.path.join(data_dir, "cost_db.json")
            shutil.copy2(src, dst)
        spin3.set()
        set_status(2, '✓')

        # Step 4: Purging Old Database Files from Cloud
        spin4 = threading.Event()
        t4 = threading.Thread(target=spin_status, args=(3, spin4, status_labels), daemon=True)
        t4.start()
        # purge_old_cloud_databases()
        spin4.set()
        set_status(3, '✓')

        # Step 5: Launching Main Window
        spin5 = threading.Event()
        t5 = threading.Thread(target=spin_status, args=(4, spin5, status_labels), daemon=True)
        t5.start()
        # Simulate a short delay for UI polish
        import time
        time.sleep(0.3)
        spin5.set()
        set_status(4, '✓')

        # Signal that initialization is complete
        # Schedule the main app launch in the main thread
        root.after(500, launch_main_app)

    def launch_main_app():
        # This runs in the main thread
        try:
            splash.destroy()
        except:
            pass  # In case splash was already destroyed
        
        # Ensure the root window is properly hidden
        root.withdraw()
        
        # Create and show main app
        app = OptionsParserApp()
        
        # Force window to be visible and properly positioned
        app.restore_window_visibility()
        
        # Add restoration bindings for future window management issues
        def on_map_event(event):
            if event.widget == app:
                # Only restore if window was actually invisible
                current_state = app.state()
                if current_state == 'iconic':
                    app.after(50, app.restore_window_visibility)
        
        def on_focus_event(event):
            if event.widget == app:
                # Just bring to front, don't force state changes
                app.after(10, lambda: app.lift())
        
        app.bind('<Map>', on_map_event)
        app.bind('<FocusIn>', on_focus_event)
        
        # Also bind to state changes - but only fix real problems
        def on_state_change():
            try:
                current_state = app.state()
                is_visible = app.winfo_viewable()
                
                # Only intervene if there's actually a problem
                # (maximized but invisible, or minimized when shouldn't be)
                if current_state == 'zoomed' and not is_visible:
                    # Window is maximized but not visible - this is the real problem
                    app.after(100, app.restore_window_visibility)
                elif current_state == 'iconic' and is_visible:
                    # Window claims to be minimized but is visible - unusual state
                    app.after(100, app.restore_window_visibility)
                # Don't interfere with normal maximized or normal states
            except:
                pass
        
        # Check state only a few times at startup, then stop
        check_count = 0
        def check_visibility():
            nonlocal check_count
            if check_count < 3:  # Only check 3 times total
                on_state_change()
                check_count += 1
                app.after(1000, check_visibility)  # Check again in 1 second
        
        app.after(500, check_visibility)
        
        app.mainloop()

    if not isTest:
        threading.Thread(target=start_app, daemon=True).start()
        root.mainloop()
    else:
        # In test mode, skip the splash and web service, just start the app directly
        root.destroy()
        app = OptionsParserApp()
        app.mainloop()


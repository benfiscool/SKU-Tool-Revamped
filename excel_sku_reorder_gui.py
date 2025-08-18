#!/usr/bin/env python3
"""
Excel SKU Reorder Script with GUI

This script reorders rows in an Excel file based on a provided SKU order list.
The script preserves all row data while reordering based on the SKU column.

Usage:
    python excel_sku_reorder_gui.py [input_file] [sku_list_file] [output_file]
    python excel_sku_reorder_gui.py --gui  # Launch GUI mode
    
    If no arguments provided, GUI mode will be launched.
"""

import pandas as pd
import sys
import os
from typing import List, Optional, Tuple
import argparse
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading


class ExcelSKUReorderGUI:
    """GUI application for Excel SKU reordering."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel SKU Reorder Tool")
        self.root.geometry("850x700")  # Increased height to accommodate new UI elements
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.sku_list = []
        self.reorder_column = tk.StringVar(value="SKU")  # Default to SKU column
        self.available_columns = []
        
        self.setup_gui()
        
    def setup_gui(self):
        """Set up the GUI components."""
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel SKU Reorder Tool", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input file selection
        ttk.Label(main_frame, text="Input Excel File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(main_frame, text="Browse", command=self.browse_input_file).grid(row=1, column=2, padx=(5, 0))
        
        # Output file selection
        ttk.Label(main_frame, text="Output Excel File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_file, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(main_frame, text="Browse", command=self.browse_output_file).grid(row=2, column=2, padx=(5, 0))
        
        # Column selection for reordering
        ttk.Label(main_frame, text="Reorder Based On:").grid(row=3, column=0, sticky=tk.W, pady=(20, 5))
        
        # Column selection frame
        column_frame = ttk.LabelFrame(main_frame, text="Select Column to Reorder By", padding="10")
        column_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        column_frame.columnconfigure(0, weight=1)
        
        # Radio buttons for column selection
        self.column_radio_frame = ttk.Frame(column_frame)
        self.column_radio_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Default radio buttons (will be updated when file is loaded)
        self.update_column_radio_buttons(['SKU', 'New SKU'])
        
        # Button to analyze file columns
        ttk.Button(column_frame, text="Analyze File Columns", 
                  command=self.analyze_file_columns).grid(row=1, column=0, pady=(10, 0), sticky=tk.W)
        
        # SKU list input methods
        ttk.Label(main_frame, text="SKU List Input:").grid(row=5, column=0, sticky=tk.W, pady=(20, 5))
        
        # SKU input frame
        sku_frame = ttk.Frame(main_frame)
        sku_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        sku_frame.columnconfigure(1, weight=1)
        
        # SKU input methods
        ttk.Button(sku_frame, text="Load from File", command=self.load_sku_file).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(sku_frame, text="Enter Manually", command=self.enter_sku_manually).grid(row=0, column=1, sticky=tk.W)
        
        # SKU list display
        ttk.Label(main_frame, text="Current SKU List:").grid(row=7, column=0, sticky=tk.W, pady=(20, 5))
        
        # SKU list text area
        self.sku_text = scrolledtext.ScrolledText(main_frame, height=8, width=70)
        self.sku_text.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Configure text area to expand
        main_frame.rowconfigure(8, weight=1)
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=9, column=0, columnspan=3, pady=(20, 0))
        
        ttk.Button(button_frame, text="Clear SKU List", command=self.clear_sku_list).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Process File", command=self.process_file).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(20, 0))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready", foreground="green")
        self.status_label.grid(row=11, column=0, columnspan=3, pady=(10, 0))
        
    def browse_input_file(self):
        """Browse for input Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # Auto-generate output filename if not set
            if not self.output_file.get():
                base_name = os.path.splitext(filename)[0]
                output_name = f"{base_name}_reordered.xlsx"
                self.output_file.set(output_name)
            # Automatically analyze file columns when input file is selected
            self.analyze_file_columns()

    def update_column_radio_buttons(self, columns):
        """Update the radio buttons for column selection."""
        # Clear existing radio buttons
        for widget in self.column_radio_frame.winfo_children():
            widget.destroy()
        
        # Create new radio buttons
        for i, column in enumerate(columns):
            rb = ttk.Radiobutton(self.column_radio_frame, text=column, 
                               variable=self.reorder_column, value=column)
            rb.grid(row=0, column=i, sticky=tk.W, padx=(0, 15))
        
        # Set default selection if available
        if 'SKU' in columns:
            self.reorder_column.set('SKU')
        elif 'New SKU' in columns:
            self.reorder_column.set('New SKU')
        elif columns:
            self.reorder_column.set(columns[0])

    def analyze_file_columns(self):
        """Analyze the input file and update available columns."""
        if not self.input_file.get() or not os.path.exists(self.input_file.get()):
            messagebox.showerror("Error", "Please select a valid input Excel file first")
            return
        
        try:
            # Read just the header to get column names
            df = pd.read_excel(self.input_file.get(), nrows=0)
            self.available_columns = list(df.columns)
            
            # Update radio buttons with actual columns
            self.update_column_radio_buttons(self.available_columns)
            
            self.status_label.config(
                text=f"Found {len(self.available_columns)} columns: {', '.join(self.available_columns)}", 
                foreground="blue"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to analyze file columns: {e}")
            self.status_label.config(text="Failed to analyze file", foreground="red")
    
    def browse_output_file(self):
        """Browse for output Excel file."""
        filename = filedialog.asksaveasfilename(
            title="Save Output Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
    
    def load_sku_file(self):
        """Load SKU list from a text file."""
        filename = filedialog.askopenfilename(
            title="Select SKU List File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    skus = [line.strip() for line in f.readlines() if line.strip()]
                self.sku_list = skus
                self.update_sku_display()
                self.status_label.config(text=f"Loaded {len(skus)} SKUs from file", foreground="green")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load SKU file: {e}")
    
    def enter_sku_manually(self):
        """Open dialog to enter SKUs manually."""
        dialog = SKUInputDialog(self.root)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            self.sku_list = dialog.result
            self.update_sku_display()
            self.status_label.config(text=f"Entered {len(dialog.result)} SKUs manually", foreground="green")
    
    def update_sku_display(self):
        """Update the SKU list display."""
        self.sku_text.delete(1.0, tk.END)
        for i, sku in enumerate(self.sku_list, 1):
            self.sku_text.insert(tk.END, f"{i:3d}. {sku}\n")
    
    def clear_sku_list(self):
        """Clear the SKU list."""
        self.sku_list = []
        self.sku_text.delete(1.0, tk.END)
        self.status_label.config(text="SKU list cleared", foreground="blue")
    
    def process_file(self):
        """Process the Excel file with the current settings."""
        # Validate inputs
        if not self.input_file.get():
            messagebox.showerror("Error", "Please select an input Excel file")
            return
        
        if not self.output_file.get():
            messagebox.showerror("Error", "Please specify an output Excel file")
            return
        
        if not self.sku_list:
            messagebox.showerror("Error", "Please provide a SKU list")
            return
        
        if not self.reorder_column.get():
            messagebox.showerror("Error", "Please select a column to reorder by")
            return
        
        if not os.path.exists(self.input_file.get()):
            messagebox.showerror("Error", "Input file does not exist")
            return
        
        # Start processing in a separate thread
        self.progress.start()
        self.status_label.config(text=f"Processing using '{self.reorder_column.get()}' column...", foreground="orange")
        
        thread = threading.Thread(target=self.process_file_thread)
        thread.daemon = True
        thread.start()
    
    def process_file_thread(self):
        """Process the file in a separate thread."""
        try:
            success = reorder_excel_by_sku(
                self.input_file.get(),
                self.output_file.get(),
                self.sku_list,
                self.reorder_column.get()  # Pass the selected column
            )
            
            # Update UI in main thread
            self.root.after(0, self.process_complete, success)
            
        except Exception as e:
            self.root.after(0, self.process_error, str(e))
    
    def process_complete(self, success):
        """Handle completion of file processing."""
        self.progress.stop()
        
        if success:
            self.status_label.config(text="File processed successfully!", foreground="green")
            
            # Ask if user wants to open the file
            result = messagebox.askyesnocancel(
                "Success", 
                "Excel file has been reordered successfully!\n\nOpen the file now?",
                icon='question'
            )
            
            if result:  # User clicked Yes
                try:
                    import subprocess
                    import platform
                    
                    output_file = self.output_file.get()
                    
                    # Open file using the default application based on OS
                    if platform.system() == 'Windows':
                        subprocess.run(['start', output_file], shell=True, check=True)
                    elif platform.system() == 'Darwin':  # macOS
                        subprocess.run(['open', output_file], check=True)
                    else:  # Linux and others
                        subprocess.run(['xdg-open', output_file], check=True)
                        
                    self.status_label.config(text="File opened successfully!", foreground="green")
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open file: {e}")
                    self.status_label.config(text="File processed but could not open", foreground="orange")
            elif result is False:  # User clicked No
                messagebox.showinfo("Complete", "File processing complete!")
            # If result is None (Cancel), do nothing additional
                
        else:
            self.status_label.config(text="Processing failed", foreground="red")
            messagebox.showerror("Error", "Failed to process the Excel file")
    
    def process_error(self, error_msg):
        """Handle processing error."""
        self.progress.stop()
        self.status_label.config(text="Processing failed", foreground="red")
        messagebox.showerror("Error", f"An error occurred: {error_msg}")


class SKUInputDialog:
    """Dialog for manual SKU input."""
    
    def __init__(self, parent):
        self.result = None
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Enter SKUs")
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (400 // 2)
        self.dialog.geometry(f"500x400+{x}+{y}")
        
        self.setup_dialog()
    
    def setup_dialog(self):
        """Set up the dialog components."""
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="Enter SKUs one per line or comma-separated:")
        instructions.pack(pady=(0, 10))
        
        # Text area
        self.text_area = scrolledtext.ScrolledText(main_frame, height=15, width=60)
        self.text_area.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))
        
        ttk.Button(button_frame, text="OK", command=self.ok_clicked).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Cancel", command=self.cancel_clicked).pack(side=tk.LEFT)
        
        # Focus on text area
        self.text_area.focus_set()
    
    def ok_clicked(self):
        """Handle OK button click."""
        text = self.text_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showerror("Error", "Please enter at least one SKU")
            return
        
        # Parse SKUs (support both line-separated and comma-separated)
        skus = []
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if ',' in line:
                # Comma-separated
                skus.extend([sku.strip() for sku in line.split(',') if sku.strip()])
            elif line:
                # Line-separated
                skus.append(line)
        
        if not skus:
            messagebox.showerror("Error", "No valid SKUs found")
            return
        
        self.result = skus
        self.dialog.destroy()
    
    def cancel_clicked(self):
        """Handle Cancel button click."""
        self.dialog.destroy()


def read_sku_list(sku_file_path: str) -> List[str]:
    """
    Read SKU list from a text file (one SKU per line).
    
    Args:
        sku_file_path: Path to the text file containing SKUs
        
    Returns:
        List of SKUs in the desired order
    """
    try:
        with open(sku_file_path, 'r', encoding='utf-8') as f:
            skus = [line.strip() for line in f.readlines() if line.strip()]
        return skus
    except FileNotFoundError:
        print(f"Error: SKU file '{sku_file_path}' not found.")
        return []
    except Exception as e:
        print(f"Error reading SKU file: {e}")
        return []


def reorder_excel_by_sku(input_file: str, output_file: str, sku_order: List[str], sku_column: str = 'SKU') -> bool:
    """
    Reorder Excel rows based on SKU column matching the provided SKU order.
    
    Args:
        input_file: Path to the input Excel file
        output_file: Path to the output Excel file
        sku_order: List of SKUs in the desired order
        sku_column: Name of the column to use for reordering (default: 'SKU')
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Read the Excel file
        print(f"Reading Excel file: {input_file}")
        df = pd.read_excel(input_file)
        
        # Check if specified column exists
        if sku_column not in df.columns:
            print(f"Error: '{sku_column}' column not found in the Excel file.")
            print(f"Available columns: {list(df.columns)}")
            return False
          # Clean the SKU data to remove any whitespace issues
        df[sku_column] = df[sku_column].astype(str).str.strip()
        sku_order_cleaned = [str(sku).strip() for sku in sku_order]
        
        print(f"Found {len(df)} rows in the original file.")
        print(f"{sku_column} column contains {len(df[sku_column].unique())} unique values.")
        print(f"Processing {len(sku_order_cleaned)} SKUs from your list.")
        
        # Debug: show first few SKUs from each source
        print(f"\nFirst 5 values from {sku_column} column:")
        for i, sku in enumerate(df[sku_column].head().tolist()):
            print(f"  {i+1}. '{sku}' (length: {len(sku)})")
        
        print(f"\nFirst 5 SKUs from your list:")
        for i, sku in enumerate(sku_order_cleaned[:5]):
            print(f"  {i+1}. '{sku}' (length: {len(sku)})")
        
        # Create a list to hold the reordered rows
        reordered_rows = []
        found_skus = []
        
        print(f"\nProcessing SKUs in the following order:")
        for i, sku in enumerate(sku_order_cleaned, 1):
            print(f"  {i}. {sku}")
        
        # Process each SKU in the desired order
        for i, sku in enumerate(sku_order_cleaned):
            # Find all rows with this SKU
            matching_rows = df[df[sku_column] == sku]
            if not matching_rows.empty:
                print(f"Found {len(matching_rows)} row(s) for SKU: '{sku}' (position {i+1})")
                # Add all matching rows to our result (preserves duplicates if any)
                reordered_rows.append(matching_rows)
                found_skus.append(sku)
            else:
                print(f"SKU not found in Excel: '{sku}' (position {i+1})")
                # Debug: check for similar SKUs
                similar_skus = df[df[sku_column].str.contains(sku[:10] if len(sku) >= 10 else sku, na=False)][sku_column].unique()
                if len(similar_skus) > 0:
                    print(f"  Similar SKUs found: {similar_skus[:3]}")
        
        # Check if we found any matches
        if not reordered_rows:
            print("Warning: No matching SKUs found between the file and the provided SKU list.")
            return False
        
        # Concatenate all the reordered rows
        df_sorted = pd.concat(reordered_rows, ignore_index=True)
        
        print(f"\nFinal order in output:")
        for i, sku in enumerate(df_sorted[sku_column].tolist(), 1):
            print(f"  {i}. {sku}")
          # Save to output file
        print(f"Saving reordered data to: {output_file}")
        df_sorted.to_excel(output_file, index=False)
        
        print(f"Successfully reordered {len(df_sorted)} rows based on {sku_column} column.")
        
        # Verify the order is correct
        output_sku_order = df_sorted[sku_column].tolist()
        expected_order = [sku for sku in sku_order if sku in found_skus]
        
        print(f"\nOrder verification:")
        print(f"Expected order (SKUs found): {expected_order}")
        print(f"Actual output order: {output_sku_order}")
        
        if output_sku_order == expected_order:
            print("✅ Order is correct!")
        else:
            print("❌ Order mismatch detected!")
            print("Differences:")
            for i, (expected, actual) in enumerate(zip(expected_order, output_sku_order)):
                if expected != actual:
                    print(f"  Position {i+1}: Expected '{expected}', Got '{actual}'")
          # Show some statistics
        matched_skus = set(df_sorted[sku_column].unique())
        requested_skus = set(sku_order)
        missing_skus = requested_skus - matched_skus
        
        if missing_skus:
            print(f"Note: {len(missing_skus)} SKUs from your list were not found in the {sku_column} column:")
            for sku in sorted(missing_skus):
                print(f"  - {sku}")
        
        return True
        
    except FileNotFoundError:
        print(f"Error: Input file '{input_file}' not found.")
        return False
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        return False


def get_user_input() -> tuple:
    """
    Get input from user interactively.
    
    Returns:
        Tuple of (input_file, output_file, sku_list)
    """
    print("\n=== Excel SKU Reorder Tool ===")
    print("This tool will reorder rows in an Excel file based on a SKU list.\n")
    
    # Get input file
    while True:
        input_file = input("Enter the path to your input Excel file: ").strip().strip('"')
        if os.path.exists(input_file):
            break
        print(f"File not found: {input_file}")
    
    # Get output file
    output_file = input("Enter the path for the output Excel file: ").strip().strip('"')
    if not output_file:
        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_reordered.xlsx"
        print(f"Using default output file: {output_file}")
    
    # Get SKU list - either from file or manual input
    print("\nHow would you like to provide the SKU list?")
    print("1. From a text file (one SKU per line)")
    print("2. Enter SKUs manually (comma-separated)")
    
    choice = input("Enter your choice (1 or 2): ").strip()
    
    if choice == "1":
        while True:
            sku_file = input("Enter the path to your SKU list file: ").strip().strip('"')
            if os.path.exists(sku_file):
                sku_list = read_sku_list(sku_file)
                if sku_list:
                    break
                print("SKU file is empty or couldn't be read.")
            else:
                print(f"File not found: {sku_file}")
    else:
        sku_input = input("Enter SKUs separated by commas: ").strip()
        sku_list = [sku.strip() for sku in sku_input.split(',') if sku.strip()]
    
    print(f"\nProcessing {len(sku_list)} SKUs...")
    return input_file, output_file, sku_list


def main():
    """Main function to handle command line arguments or launch GUI."""
    
    parser = argparse.ArgumentParser(description='Excel SKU Reorder Tool')
    parser.add_argument('--gui', action='store_true', help='Launch GUI mode')
    parser.add_argument('input_file', nargs='?', help='Input Excel file')
    parser.add_argument('output_file', nargs='?', help='Output Excel file')
    parser.add_argument('sku_list_file', nargs='?', help='SKU list file')
    
    args = parser.parse_args()
    
    # Determine mode
    if args.gui or (not args.input_file and not args.output_file and not args.sku_list_file):
        # Launch GUI
        root = tk.Tk()
        app = ExcelSKUReorderGUI(root)
        root.mainloop()
        
    elif args.input_file and args.output_file and args.sku_list_file:
        # Command line mode
        sku_list = read_sku_list(args.sku_list_file)
        if not sku_list:
            sys.exit(1)
        
        success = reorder_excel_by_sku(args.input_file, args.output_file, sku_list, 'SKU')  # Default to SKU column
        
        if success:
            print("\n✅ File processing completed successfully!")
        else:
            print("\n❌ File processing failed.")
            sys.exit(1)
            
    else:
        # Interactive mode
        input_file, output_file, sku_list = get_user_input()
        
        success = reorder_excel_by_sku(input_file, output_file, sku_list, 'SKU')  # Default to SKU column
        
        if success:
            print("\n✅ File processing completed successfully!")
        else:
            print("\n❌ File processing failed.")
            sys.exit(1)


if __name__ == "__main__":
    main()

import pandas as pd
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk
import os

def remove_duplicates_and_export():
    """
    Program to remove duplicates from comma-separated phrases and export to Excel
    """
    
    def process_data():
        # Get the input text
        input_text = text_area.get("1.0", tk.END).strip()
        
        if not input_text:
            messagebox.showwarning("No Input", "Please enter some comma-separated phrases.")
            return
        
        # Split by commas and/or line breaks and clean up each phrase
        import re
        # Split by comma, newline, or carriage return, then clean each phrase
        phrases = re.split(r'[,\n\r]+', input_text)
        phrases = [phrase.strip() for phrase in phrases if phrase.strip()]
        
        # Remove colon and everything after it if checkbox is enabled
        if remove_colon_var.get():
            phrases = [phrase.split(':')[0].strip() for phrase in phrases]
            # Filter out any phrases that became empty after colon removal
            phrases = [phrase for phrase in phrases if phrase]
        
        if not phrases:
            messagebox.showwarning("No Data", "No valid phrases found.")
            return
        
        # Remove duplicates while preserving order
        unique_phrases = []
        seen = set()
        for phrase in phrases:
            phrase_lower = phrase.lower()  # Case-insensitive duplicate detection
            if phrase_lower not in seen:
                unique_phrases.append(phrase)
                seen.add(phrase_lower)
        
        # Display results
        result_text = f"Original count: {len(phrases)}\n"
        result_text += f"Unique count: {len(unique_phrases)}\n"
        result_text += f"Duplicates removed: {len(phrases) - len(unique_phrases)}\n\n"
        result_text += "Unique phrases:\n" + "\n".join(f"{i+1}. {phrase}" for i, phrase in enumerate(unique_phrases))
        
        result_area.delete("1.0", tk.END)
        result_area.insert("1.0", result_text)
        
        # Store unique phrases for export
        process_data.unique_phrases = unique_phrases
    
    def export_to_excel():
        if not hasattr(process_data, 'unique_phrases') or not process_data.unique_phrases:
            messagebox.showwarning("No Data", "Please process some data first.")
            return
        
        # Ask user where to save
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')],
            title="Save Excel file"
        )
        
        if not file_path:
            return
        
        try:
            # Create DataFrame
            df = pd.DataFrame({
                'Unique Phrases': process_data.unique_phrases
            })
            
            # Export to Excel
            df.to_excel(file_path, index=False, sheet_name='Unique Phrases')
            
            messagebox.showinfo("Success", f"Excel file saved successfully!\n\nFile: {file_path}\nUnique phrases: {len(process_data.unique_phrases)}")
            
            # Automatically open the file
            try:
                os.startfile(file_path)  # Windows
            except:
                try:
                    os.system(f'open "{file_path}"')  # macOS
                except:
                    try:
                        os.system(f'xdg-open "{file_path}"')  # Linux
                    except:
                        messagebox.showinfo("Info", "File saved but couldn't open automatically.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file:\n{str(e)}")
    
    def clear_all():
        text_area.delete("1.0", tk.END)
        result_area.delete("1.0", tk.END)
        if hasattr(process_data, 'unique_phrases'):
            delattr(process_data, 'unique_phrases')
    
    # Create main window
    root = tk.Tk()
    root.title("Duplicate Remover - Excel Export")
    root.geometry("800x600")
    
    # Create checkbox variable
    remove_colon_var = tk.BooleanVar(value=True)  # Default to enabled
    
    # Main frame
    main_frame = ttk.Frame(root, padding="10")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    # Configure grid weights
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    main_frame.columnconfigure(1, weight=1)
    main_frame.rowconfigure(2, weight=1)
    main_frame.rowconfigure(6, weight=1)
    
    # Title
    title_label = ttk.Label(main_frame, text="Duplicate Remover & Excel Exporter", 
                           font=('Arial', 14, 'bold'))
    title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
    
    # Input section
    input_label = ttk.Label(main_frame, text="Enter phrases (comma or line separated):")
    input_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
    
    # Input text area with scrollbar
    input_frame = ttk.Frame(main_frame)
    input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
    input_frame.columnconfigure(0, weight=1)
    input_frame.rowconfigure(0, weight=1)
    
    text_area = tk.Text(input_frame, height=8, wrap=tk.WORD)
    text_area.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    input_scrollbar = ttk.Scrollbar(input_frame, orient=tk.VERTICAL, command=text_area.yview)
    input_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    text_area.config(yscrollcommand=input_scrollbar.set)
    
    # Options section
    options_frame = ttk.Frame(main_frame)
    options_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
    
    remove_colon_checkbox = ttk.Checkbutton(
        options_frame, 
        text="Remove colon (:) and everything after it from each phrase",
        variable=remove_colon_var
    )
    remove_colon_checkbox.pack(anchor='w')
    
    # Buttons frame
    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(10, 10))
    
    ttk.Button(button_frame, text="Process Data", command=process_data).pack(side=tk.LEFT, padx=(0, 5))
    ttk.Button(button_frame, text="Export to Excel", command=export_to_excel).pack(side=tk.LEFT, padx=(0, 5))
    ttk.Button(button_frame, text="Clear All", command=clear_all).pack(side=tk.LEFT, padx=(0, 5))
    
    # Results section
    result_label = ttk.Label(main_frame, text="Results:")
    result_label.grid(row=5, column=0, sticky=tk.W, pady=(10, 5))
    
    # Results text area with scrollbar
    result_frame = ttk.Frame(main_frame)
    result_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
    result_frame.columnconfigure(0, weight=1)
    result_frame.rowconfigure(0, weight=1)
    
    result_area = tk.Text(result_frame, height=12, wrap=tk.WORD, state=tk.NORMAL)
    result_area.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    result_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=result_area.yview)
    result_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    result_area.config(yscrollcommand=result_scrollbar.set)
    
    # Instructions
    instructions = """
Instructions:
1. Enter your phrases in the text area above (separated by commas or line breaks)
2. Use the checkbox to remove colons and everything after them (enabled by default)
3. Click "Process Data" to remove duplicates
4. Click "Export to Excel" to save the unique phrases to an Excel file
5. The program removes duplicates case-insensitively (Apple = apple = APPLE)
6. Each unique phrase will be on its own row in the Excel output
7. Supports both comma-separated and line-separated input formats

Example with colon removal:
Input: "phrase1: 123, phrase2: 456"
Output: "phrase1, phrase2"
    """
    
    result_area.insert("1.0", instructions)
    
    root.mainloop()

if __name__ == "__main__":
    try:
        remove_duplicates_and_export()
    except ImportError as e:
        print("Error: Missing required library.")
        print("Please install pandas with: pip install pandas openpyxl")
        print(f"Full error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

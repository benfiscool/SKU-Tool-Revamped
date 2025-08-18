# Created by Ben
import time
import sys
import tkinter as tk
import keyboard
import pyperclip
import re
import threading

# --- feature state and hotkey handles ---
title_enabled = False
seq_enabled = False
copy_enabled = False
math_enabled = False
clipboard_fix_enabled = False
title_handle = None
seq_handle = None
copy_handle = None
math_handle = None
calc_handle = None
clipboard_monitor_thread = None
clipboard_monitor_running = False

# Math calculator state
math_expression = ""
capturing_math = False

words_list = []
seq_index = 0
auto_enter = False

# Clipboard monitoring state
last_clipboard_content = ""

# --- clipboard monitoring function ---
def monitor_clipboard():
    """Monitor clipboard for changes and fix multiple spaces"""
    global last_clipboard_content, clipboard_monitor_running
    
    while clipboard_monitor_running:
        try:
            current_content = pyperclip.paste()
            if current_content != last_clipboard_content:
                # Check if the content has multiple consecutive spaces
                if re.search(r'  +', current_content):  # Two or more consecutive spaces
                    # Fix multiple spaces by replacing them with single spaces
                    fixed_content = re.sub(r' +', ' ', current_content)
                    # Only update if something actually changed
                    if fixed_content != current_content:
                        pyperclip.copy(fixed_content)
                        print(f"Fixed clipboard content: multiple spaces removed")  # Debug
                
                last_clipboard_content = pyperclip.paste()  # Update with the current (possibly fixed) content
            
            time.sleep(0.1)  # Check every 100ms
        except Exception as e:
            print(f"Clipboard monitoring error: {e}")
            time.sleep(0.5)  # Wait longer on error

# --- core functions ---
def title_case_selected():
    previous = pyperclip.paste()
    keyboard.press_and_release('ctrl+c')
    time.sleep(0.05)
    txt = pyperclip.paste()
    def cap(w): return w[:1].upper() + w[1:] if w else ''
    transformed = ' '.join(cap(w) for w in txt.split(' '))
    pyperclip.copy(transformed)
    time.sleep(0.05)
    keyboard.press_and_release('ctrl+v')
    time.sleep(0.05)
    pyperclip.copy(previous)


def copy_selected_to_entry():
    """Copy selected text to the entry field, overwriting existing content"""
    previous = pyperclip.paste()
    keyboard.press_and_release('ctrl+c')
    time.sleep(0.05)
    selected_text = pyperclip.paste()
    
    # Restore original clipboard
    pyperclip.copy(previous)
    
    # Clear entry and add selected text
    if selected_text.strip():
        entry.delete(0, tk.END)
        entry.insert(0, selected_text.strip())
        load_list()  # Update the word list immediately


def calculate_math():
    """Calculate the captured math expression and replace it"""
    global math_expression, capturing_math
    
    print(f"Calculate called. Expression: '{math_expression}', Capturing: {capturing_math}")  # Debug
    
    if not math_expression.strip():
        print("No expression to calculate")  # Debug
        return
    
    try:
        # Clean the expression - only allow numbers, operators, parentheses, and decimal points
        if re.match(r'^[0-9+\-*/().\s]+$', math_expression):
            try:
                # Evaluate the expression safely
                result = eval(math_expression)
                print(f"Calculation result: {result}")  # Debug
                
                # Format the result (remove .0 for whole numbers)
                if isinstance(result, float) and result.is_integer():
                    result = int(result)
                
                # Calculate how many characters to delete (= + expression)
                chars_to_delete = len(math_expression) + 1  # +1 for the = sign
                print(f"Deleting {chars_to_delete} characters")  # Debug
                
                # Delete the expression by simulating backspace
                for _ in range(chars_to_delete):
                    keyboard.press_and_release('backspace')
                    time.sleep(0.01)
                
                # Type the result
                keyboard.write(str(result))
                
            except Exception as e:
                print(f"Calculation failed: {e}")  # Debug
                # If calculation fails, do nothing
                pass
        else:
            print("Expression contains invalid characters")  # Debug
        
        # Reset the math expression
        math_expression = ""
        capturing_math = False
        update_title()
        
    except Exception as e:
        print(f"Error in calculate_math: {e}")  # Debug
        # Reset on any error
        math_expression = ""
        capturing_math = False
        update_title()


def update_title():
    """Update window title to show math capture status"""
    if capturing_math and math_expression:
        root.title(f"Hotkey Manager - Math: ={math_expression}")
    elif capturing_math:
        root.title("Hotkey Manager - Math: =")
    else:
        root.title("Hotkey Manager")


def on_key_event(event):
    """Handle key events for math expression capture"""
    global math_expression, capturing_math
    
    if not math_enabled:
        return
    
    # Check if this is a key press (not release)
    if event.event_type != keyboard.KEY_DOWN:
        return
    
    key_name = event.name
    
    # Start capturing when = is pressed
    if key_name == '=' and not capturing_math:
        capturing_math = True
        math_expression = ""
        print("Started math capture")  # Debug
        update_title()
        return
    
    # If we're capturing math input
    if capturing_math:
        print(f"Captured key: {key_name}")  # Debug
        
        # Stop capturing on Enter, Escape, or Tab
        if key_name in ['enter', 'esc', 'tab']:
            capturing_math = False
            math_expression = ""
            print("Stopped math capture")  # Debug
            update_title()
            return
        
        # Handle backspace
        if key_name == 'backspace':
            if math_expression:
                math_expression = math_expression[:-1]
                print(f"Backspace - expression now: '{math_expression}'")  # Debug
                update_title()
            else:
                # If expression is empty, stop capturing (user backspaced to the = sign)
                capturing_math = False
                print("Stopped math capture (backspaced to empty)")  # Debug
                update_title()
                # Also delete the = sign that was originally typed
                keyboard.press_and_release('backspace')
            return
        
        # Capture numeric and operator keys
        if key_name.isdigit():
            math_expression += key_name
            update_title()
        elif key_name in ['+', '-', '*', '/', '(', ')', '.', 'decimal']:
            # Handle the decimal key from number pad
            if key_name == 'decimal':
                math_expression += '.'
            else:
                math_expression += key_name
            update_title()
        elif key_name == 'space':
            # Allow spaces in expressions for readability
            math_expression += ' '
            update_title()
        elif key_name in ['shift', 'ctrl', 'alt']:
            # Ignore modifier keys - don't stop capturing
            return
        else:
            # Any other key stops math capture
            print(f"Stopped math capture due to key: {key_name}")  # Debug
            capturing_math = False
            math_expression = ""
            update_title()


def sequential_paste():
    global seq_index, auto_enter
    if not words_list:
        return
    
    if auto_enter:
        # Paste all words with Enter between them (but not after the last one)
        for i, word in enumerate(words_list):
            keyboard.write(word.strip())
            # Press enter after each word except the last one
            if i < len(words_list) - 1:
                time.sleep(0.1)
                keyboard.press_and_release('enter')
                time.sleep(0.1)  # Extra delay between each paste
        
        # Reset index for next time
        seq_index = 0
    else:
        # Original behavior: paste one word at a time
        word = words_list[seq_index].strip()
        seq_index += 1
        
        # Type the word directly
        keyboard.write(word)
        
        # Loop back to beginning when done
        if seq_index >= len(words_list):
            seq_index = 0

# --- toggle logic ---
def toggle_title():
    global title_enabled, title_handle
    if title_enabled:
        if title_handle:
            keyboard.remove_hotkey(title_handle)
        title_handle = None
        title_enabled = False
    else:
        title_handle = keyboard.add_hotkey('ctrl+q', title_case_selected, suppress=False)
        title_enabled = True


def toggle_seq():
    global seq_enabled, seq_handle
    if seq_enabled:
        if seq_handle:
            keyboard.remove_hotkey(seq_handle)
        seq_handle = None
        seq_enabled = False
    else:
        seq_handle = keyboard.add_hotkey('f1', sequential_paste, suppress=False)
        seq_enabled = True


def toggle_copy():
    global copy_enabled, copy_handle
    if copy_enabled:
        if copy_handle:
            keyboard.remove_hotkey(copy_handle)
        copy_handle = None
        copy_enabled = False
    else:
        copy_handle = keyboard.add_hotkey('f2', copy_selected_to_entry, suppress=False)
        copy_enabled = True


def toggle_math():
    global math_enabled, math_handle, calc_handle
    if math_enabled:
        if math_handle:
            keyboard.unhook(math_handle)
        if calc_handle:
            keyboard.remove_hotkey(calc_handle)
        math_handle = None
        calc_handle = None
        math_enabled = False
    else:
        # Hook all key events to capture math expressions
        math_handle = keyboard.hook(on_key_event)
        # Also add hotkey for calculation - using F4 which is less likely to conflict
        calc_handle = keyboard.add_hotkey('f4', calculate_math, suppress=False)
        math_enabled = True


def toggle_clipboard_fix():
    global clipboard_fix_enabled, clipboard_monitor_thread, clipboard_monitor_running, last_clipboard_content
    if clipboard_fix_enabled:
        # Stop monitoring
        clipboard_monitor_running = False
        if clipboard_monitor_thread and clipboard_monitor_thread.is_alive():
            clipboard_monitor_thread.join(timeout=1)
        clipboard_monitor_thread = None
        clipboard_fix_enabled = False
        print("Clipboard monitoring stopped")  # Debug
    else:
        # Start monitoring
        clipboard_monitor_running = True
        last_clipboard_content = pyperclip.paste()  # Initialize with current clipboard
        clipboard_monitor_thread = threading.Thread(target=monitor_clipboard, daemon=True)
        clipboard_monitor_thread.start()
        clipboard_fix_enabled = True
        print("Clipboard monitoring started")  # Debug


def toggle_auto_enter():
    global auto_enter
    auto_enter = auto_enter_var.get()


def load_list():
    global words_list, seq_index
    raw = entry.get()
    words_list = [w for w in raw.split(',') if w.strip()]
    seq_index = 0

def clear_seq_list():
    """Clear the sequential paste list"""
    global words_list, seq_index
    words_list = []
    seq_index = 0
    entry.delete(0, tk.END)

def on_entry_change(*args):
    """Automatically update the word list when entry changes"""
    load_list()

def on_entry_click(event):
    """Select all text when entry is clicked"""
    entry.select_range(0, tk.END)
    entry.icursor(tk.END)

# --- build the GUI ---
root = tk.Tk()
root.title("Hotkey Manager")

root.resizable(False, False)
root.attributes("-topmost", True)

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

title_var = tk.BooleanVar(value=False)
seq_var = tk.BooleanVar(value=False)
copy_var = tk.BooleanVar(value=False)
math_var = tk.BooleanVar(value=False)
auto_enter_var = tk.BooleanVar(value=False)
clipboard_fix_var = tk.BooleanVar(value=False)

chk1 = tk.Checkbutton(frame, text="Enable Title‑Case (Ctrl+Q)",
                      variable=title_var, command=toggle_title)
chk1.grid(row=0, column=0, sticky='w')

chk2 = tk.Checkbutton(frame, text="Enable Copy Selected to List (F2)",
                      variable=copy_var, command=toggle_copy)
chk2.grid(row=1, column=0, sticky='w', pady=(5,0))

chk3 = tk.Checkbutton(frame, text="Enable Math Calculator (F4)",
                      variable=math_var, command=toggle_math)
chk3.grid(row=2, column=0, sticky='w', pady=(5,0))

chk4 = tk.Checkbutton(frame, text="Enable Sequential Paste (F1)",
                      variable=seq_var, command=toggle_seq)
chk4.grid(row=3, column=0, sticky='w', pady=(5,0))

chk5 = tk.Checkbutton(frame, text="Auto-press Enter after each paste",
                      variable=auto_enter_var, command=toggle_auto_enter)
chk5.grid(row=4, column=0, sticky='w', pady=(5,0))

chk6 = tk.Checkbutton(frame, text="Auto-fix multiple spaces in clipboard",
                      variable=clipboard_fix_var, command=toggle_clipboard_fix)
chk6.grid(row=5, column=0, sticky='w', pady=(5,0))

tk.Label(frame, text="Comma‑separated words for sequential paste:").grid(row=6, column=0, sticky='w', pady=(10,0))
entry = tk.Entry(frame, width=50)
entry.grid(row=7, column=0, sticky='ew', padx=(0,5))

clear_btn = tk.Button(frame, text="Clear", command=clear_seq_list)
clear_btn.grid(row=7, column=1)

# Set up auto-update when entry loses focus or Enter is pressed
def update_on_focus_out(event):
    load_list()

def update_on_enter(event):
    load_list()

entry.bind('<FocusOut>', update_on_focus_out)
entry.bind('<Return>', update_on_enter)
entry.bind('<Button-1>', on_entry_click)
entry.bind('<FocusIn>', on_entry_click)

# clean exit if window closed
def on_close():
    global clipboard_monitor_running
    try:
        # Stop clipboard monitoring
        clipboard_monitor_running = False
        
        # Remove hotkeys safely
        if title_handle:      
            try: keyboard.remove_hotkey(title_handle)
            except: pass
        if seq_handle:        
            try: keyboard.remove_hotkey(seq_handle)
            except: pass
        if copy_handle:       
            try: keyboard.remove_hotkey(copy_handle)
            except: pass
        if math_handle:       
            try: keyboard.unhook(math_handle)
            except: pass
        if calc_handle:       
            try: keyboard.remove_hotkey(calc_handle)
            except: pass
            
        # Clear all hotkeys and hooks just in case
        keyboard.unhook_all()
    except:
        pass
    
    root.destroy()
    sys.exit(0)

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from flask import Flask, request, jsonify
import webbrowser
import pickle
import io

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/drive.file']
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, 'credentials.json')
TOKEN_FILE = os.path.join(SCRIPT_DIR, 'token.pickle')

# Configuration for Google Drive folder
DRIVE_FOLDER_NAME = "SKU Tool Database"  # Change this to your desired folder name

app = Flask(__name__)

def get_or_create_folder(service, folder_name, parent_id='root'):
    """Get existing folder or create it if it doesn't exist"""
    # Search for the folder
    results = service.files().list(
        q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false and '{parent_id}' in parents",
        spaces='drive',
        fields="files(id, name)"
    ).execute()
    folders = results.get('files', [])
    
    if folders:
        return folders[0]['id']
    else:
        # Create the folder
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id]
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        return folder.get('id')

def sign_out():
    """Remove stored credentials to force re-authentication"""
    if os.path.exists(TOKEN_FILE):
        os.remove(TOKEN_FILE)
        messagebox.showinfo("Signed Out", "Successfully signed out. Next operation will require re-authentication with a different account.")

def get_drive_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                # If refresh fails, delete the token file and re-authenticate
                print(f"Token refresh failed: {e}")
                if os.path.exists(TOKEN_FILE):
                    os.remove(TOKEN_FILE)
                creds = None
        
        if not creds:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError("Missing credentials.json for Google Drive API.")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    
    return build('drive', 'v3', credentials=creds)

def upload_file(local_path, drive_filename):
    service = get_drive_service()
    
    # Get or create the target folder
    folder_id = get_or_create_folder(service, DRIVE_FOLDER_NAME)
    
    # Check if file already exists in the folder
    results = service.files().list(
        q=f"name='{drive_filename}' and trashed=false and '{folder_id}' in parents",
        spaces='drive', 
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    
    if files:
        # Update existing file
        file_id = files[0]['id']
        media = MediaFileUpload(local_path, resumable=True)
        updated = service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    else:
        # Create new file in the folder
        file_metadata = {
            'name': drive_filename,
            'parents': [folder_id]
        }
        media = MediaFileUpload(local_path, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return file.get('id')

def download_file(file_id, local_filename):
    service = get_drive_service()
    # Always save to script directory
    local_path = os.path.join(DATA_DIR, os.path.basename(local_filename))
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(local_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    return local_path

@app.route('/upload', methods=['POST'])
def api_upload():
    file_path = request.json.get('file_path')
    drive_filename = request.json.get('drive_filename') or os.path.basename(file_path)
    try:
        file_id = upload_file(file_path, drive_filename)
        return jsonify({"status": "success", "file_id": file_id, "message": f"Uploaded {file_path} as {drive_filename}"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/download', methods=['POST'])
def api_download():
    file_id = request.json.get('file_id')
    dest_filename = request.json.get('dest_path')  # This is now just the filename
    try:
        local_path = download_file(file_id, dest_filename)
        return jsonify({"status": "success", "message": f"Downloaded {file_id} to {local_path}"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/latest_db_file_id', methods=['POST'])
def api_latest_db_file_id():
    """
    Returns the file ID of the latest database file matching the given prefix and suffix.
    Example request: {"prefix": "sku_database_", "suffix": ".json"}
    """
    prefix = request.json.get('prefix', 'sku_database_')
    suffix = request.json.get('suffix', '.json')
    service = get_drive_service()
    
    # Get or create the target folder
    folder_id = get_or_create_folder(service, DRIVE_FOLDER_NAME)
    
    # List all files matching the pattern in the folder
    results = service.files().list(
        q=f"name contains '{prefix}' and name contains '{suffix}' and trashed=false and '{folder_id}' in parents",
        spaces='drive',
        fields="files(id, name)",
        pageSize=1000
    ).execute()
    files = results.get('files', [])
    # Filter and sort by name (descending)
    matching = [f for f in files if f['name'].startswith(prefix) and f['name'].endswith(suffix)]
    if not matching:
        return jsonify({"status": "error", "message": "No matching files found"}), 404
    matching.sort(key=lambda f: f['name'], reverse=True)
    latest = matching[0]
    return jsonify({"status": "success", "file_id": latest['id'], "name": latest['name']})

@app.route('/pull_latest_db', methods=['POST'])
def api_pull_latest_db():
    """
    Downloads the latest database file matching the given prefix and suffix to the script directory.
    Example request: {"prefix": "sku_database_", "suffix": ".json"}
    """
    prefix = request.json.get('prefix', 'sku_database_')
    suffix = request.json.get('suffix', '.json')
    service = get_drive_service()
    
    # Get or create the target folder
    folder_id = get_or_create_folder(service, DRIVE_FOLDER_NAME)
    
    results = service.files().list(
        q=f"name contains '{prefix}' and name contains '{suffix}' and trashed=false and '{folder_id}' in parents",
        spaces='drive',
        fields="files(id, name)",
        pageSize=1000
    ).execute()
    files = results.get('files', [])
    matching = [f for f in files if f['name'].startswith(prefix) and f['name'].endswith(suffix)]
    if not matching:
        def show_error():
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Database Not Found", f"No database file found matching '{prefix}*{suffix}' in Google Drive folder '{DRIVE_FOLDER_NAME}'.")
            root.destroy()
        threading.Thread(target=show_error).start()
        return jsonify({"status": "error", "message": "No matching database file found"}), 404
    matching.sort(key=lambda f: f['name'], reverse=True)
    latest = matching[0]
    local_path = os.path.join(DATA_DIR, latest['name'])
    try:
        download_file(latest['id'], latest['name'])
        def show_success():
            root = tk.Tk()
            root.withdraw()
            # messagebox.showinfo("Database Downloaded", f"Downloaded {latest['name']} successfully.")
            root.destroy()
        threading.Thread(target=show_success).start()
        return jsonify({"status": "success", "file_id": latest['id'], "name": latest['name'], "local_path": local_path})
    except Exception as e:
        def show_error():
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Download Error", f"Failed to download {latest['name']}:\n{e}")
            root.destroy()
        threading.Thread(target=show_error).start()
        return jsonify({"status": "error", "message": str(e)}), 500
    
@app.route('/shutdown', methods=['POST'])
def shutdown():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is not None:
        func()
        return 'Server shutting down...'
    else:
        # Fallback: force exit if not running with Werkzeug
        import os
        os._exit(0)

def run_flask():
    app.run(port=5000, debug=False, use_reloader=False)

def manual_upload():
    path = filedialog.askopenfilename(title="Select file to upload")
    if not path:
        return
    try:
        file_id = upload_file(path, os.path.basename(path))
        messagebox.showinfo("Upload", f"Uploaded {os.path.basename(path)}\nFile ID: {file_id}")
    except Exception as e:
        messagebox.showerror("Upload Error", str(e))

def manual_download():
    service = get_drive_service()
    
    # Get or create the target folder
    folder_id = get_or_create_folder(service, DRIVE_FOLDER_NAME)
    
    # List files in the folder for user to pick
    results = service.files().list(
        q=f"trashed=false and '{folder_id}' in parents", 
        spaces='drive', 
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if not files:
        messagebox.showinfo("Download", f"No files found in Drive folder '{DRIVE_FOLDER_NAME}'.")
        return
    # Simple selection dialog
    win = tk.Toplevel()
    win.title("Select file to download or delete")
    ttk.Label(win, text="Select file:").pack(padx=10, pady=5)
    file_var = tk.StringVar()
    file_list = ttk.Combobox(win, textvariable=file_var, values=[f"{f['name']} ({f['id']})" for f in files], width=60)
    file_list.pack(padx=10, pady=5)
    ttk.Label(win, text="Save as:").pack(padx=10, pady=(10,0))
    save_var = tk.StringVar()
    save_entry = ttk.Entry(win, textvariable=save_var, width=60)
    save_entry.pack(padx=10, pady=5)
    def do_download():
        idx = file_list.current()
        if idx < 0:
            messagebox.showwarning("Select", "Please select a file.")
            return
        file_id = files[idx]['id']
        dest_path = save_var.get()
        if not dest_path:
            messagebox.showwarning("Save as", "Please enter a destination path.")
            return
        try:
            download_file(file_id, dest_path)
            messagebox.showinfo("Download", f"Downloaded to {dest_path}")
            win.destroy()
        except Exception as e:
            messagebox.showerror("Download Error", str(e))
    def do_delete():
        idx = file_list.current()
        if idx < 0:
            messagebox.showwarning("Select", "Please select a file to delete.")
            return
        file_id = files[idx]['id']
        filename = files[idx]['name']
        if not messagebox.askyesno("Delete File", f"Are you sure you want to delete '{filename}'?"):
            return
        try:
            service.files().delete(fileId=file_id).execute()
            messagebox.showinfo("Deleted", f"Deleted '{filename}' from Google Drive.")
            win.destroy()
        except Exception as e:
            messagebox.showerror("Delete Error", str(e))
    btn_frame = ttk.Frame(win)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="Download", command=do_download).pack(side='left', padx=5)
    ttk.Button(btn_frame, text="Delete", command=do_delete).pack(side='left', padx=5)
    ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side='left', padx=5)
    win.grab_set()

def launch_gui():
    root = tk.Tk()
    root.wm_state('iconic')
    root.title("AP Sku Tool Web Service")
    ttk.Label(root, text="AP Sku Tool Web Service Running").pack(padx=20, pady=10)
    ttk.Button(root, text="Manual Upload", command=manual_upload).pack(pady=5)
    ttk.Button(root, text="Manual Download", command=manual_download).pack(pady=5)
    ttk.Button(root, text="Sign Out (Switch Account)", command=sign_out).pack(pady=5)
    ttk.Button(root, text="Open Google Drive", command=lambda: webbrowser.open("https://drive.google.com")).pack(pady=5)
    def quit_app():
        try:
            import requests
            requests.post('http://localhost:5000/shutdown')
        except Exception:
            pass
        root.destroy()
    ttk.Button(root, text="Quit", command=quit_app).pack(pady=5)
    root.mainloop()

@app.route('/delete', methods=['POST'])
def api_delete():
    """
    Delete a file from Google Drive by file_id or by name.
    Request: {"file_id": "..."} or {"filename": "..."}
    """
    service = get_drive_service()
    
    # Get the target folder
    folder_id = get_or_create_folder(service, DRIVE_FOLDER_NAME)
    
    file_id = request.json.get('file_id')
    filename = request.json.get('filename')
    try:
        if not file_id and filename:
            # Find file by name in the folder
            results = service.files().list(
                q=f"name='{filename}' and trashed=false and '{folder_id}' in parents",
                spaces='drive', 
                fields="files(id, name)"
            ).execute()
            files = results.get('files', [])
            if not files:
                return jsonify({"status": "error", "message": f"No file named {filename} found in folder '{DRIVE_FOLDER_NAME}'."}), 404
            file_id = files[0]['id']
        if not file_id:
            return jsonify({"status": "error", "message": "No file_id or filename provided."}), 400
        service.files().delete(fileId=file_id).execute()
        return jsonify({"status": "success", "message": f"Deleted file {file_id}"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/list', methods=['POST'])
def api_list():
    """
    List files in Google Drive matching a prefix and suffix.
    Request: {"prefix": "...", "suffix": "..."}
    """
    prefix = request.json.get('prefix', '')
    suffix = request.json.get('suffix', '')
    service = get_drive_service()
    
    # Get the target folder
    folder_id = get_or_create_folder(service, DRIVE_FOLDER_NAME)
    
    results = service.files().list(
        q=f"name contains '{prefix}' and name contains '{suffix}' and trashed=false and '{folder_id}' in parents",
        spaces='drive',
        fields="files(id, name, createdTime, modifiedTime)",
        pageSize=1000
    ).execute()
    files = results.get('files', [])
    # Only return files that start/end with prefix/suffix
    matching = [f for f in files if f['name'].startswith(prefix) and f['name'].endswith(suffix)]
    return jsonify({"status": "success", "files": matching})

if __name__ == '__main__':
    threading.Thread(target=run_flask, daemon=True).start()
    launch_gui()
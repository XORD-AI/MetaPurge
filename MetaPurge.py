import os
import sys
import time
import winsound
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
import pikepdf
import pywintypes
import win32file
import win32con
from datetime import datetime

# --- CONFIGURATION ---
TARGET_DATE = datetime(1990, 1, 1, 0, 0, 0)
TIMESTAMP = TARGET_DATE.timestamp()

# XORD Brand Colors
COLOR_BG = "#1a1a1a"        # Near-black background
COLOR_PANEL = "#2d2d2d"     # Dark grey for panels
COLOR_ACCENT = "#00ff88"    # Matrix green
COLOR_GOLD = "#FFD700"      # XORD Gold
COLOR_TEXT = "#ffffff"      # White text
COLOR_SUBTEXT = "#888888"   # Grey subtext

def change_file_creation_time(path, date_obj, retries=3):
    """Forces the Windows 'Date Created' timestamp to 1990. Retries if locked."""
    for attempt in range(retries):
        try:
            wintime = pywintypes.Time(date_obj)
            winfile = win32file.CreateFile(
                path, win32file.GENERIC_WRITE,
                win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
                None, win32file.OPEN_EXISTING,
                win32file.FILE_ATTRIBUTE_NORMAL, None
            )
            win32file.SetFileTime(winfile, wintime, wintime, wintime)
            winfile.close()
            return True
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(0.3)
            continue
    return False

def get_cleaned_path(file_path):
    """Generates output filename."""
    base, ext = os.path.splitext(file_path)
    return f"{base}_cleaned{ext}"

def scrub_image(file_path, new_path):
    """Removes EXIF data from Images by saving a fresh copy."""
    try:
        img = Image.open(file_path)
        data = list(img.getdata())
        clean_img = Image.new(img.mode, img.size)
        clean_img.putdata(data)
        clean_img.save(new_path)
        img.close()
        clean_img.close()
        return True, None
    except Exception as e:
        return False, str(e)

def scrub_pdf(file_path, new_path):
    """Aggressively removes all PDF metadata."""
    try:
        with pikepdf.open(file_path) as pdf:
            if '/Metadata' in pdf.Root:
                del pdf.Root['/Metadata']
            
            if pdf.docinfo is not None:
                keys_to_delete = list(pdf.docinfo.keys())
                for key in keys_to_delete:
                    del pdf.docinfo[key]
            
            pdf.save(new_path)
        
        time.sleep(0.1)
        with pikepdf.open(new_path, allow_overwriting_input=True) as pdf:
            if pdf.docinfo is not None:
                keys_to_delete = list(pdf.docinfo.keys())
                for key in keys_to_delete:
                    del pdf.docinfo[key]
            
            if '/Metadata' in pdf.Root:
                del pdf.Root['/Metadata']
            
            pdf.save(new_path)
        
        return True, None
    except Exception as e:
        return False, str(e)

def process_file(file_path):
    """Orchestrator: Creates cleaned copy and nukes timestamp."""
    if os.path.isdir(file_path):
        return "SKIPPED", None, "Folder"
    
    if not os.path.isfile(file_path):
        return "FAILED", None, "File not found"
    
    ext = os.path.splitext(file_path)[1].lower()
    
    supported = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.webp', '.gif', '.pdf']
    if ext not in supported:
        return "SKIPPED", None, f"Unsupported: {ext}"
    
    new_path = get_cleaned_path(file_path)
    
    # Check if cleaned version already exists
    if os.path.exists(new_path):
        answer = messagebox.askyesno(
            "Already Cleaned",
            f"'{os.path.basename(new_path)}' already exists.\n\nOverwrite?",
            icon="question"
        )
        if not answer:
            return "SKIPPED", None, "User cancelled"
    
    if ext == '.pdf':
        success, error = scrub_pdf(file_path, new_path)
    else:
        success, error = scrub_image(file_path, new_path)

    if success:
        time.sleep(0.2)
        os.utime(new_path, (TIMESTAMP, TIMESTAMP))
        change_file_creation_time(new_path, TARGET_DATE)
        return "CLEANED", new_path, None
    else:
        return "FAILED", None, error

def flash_status(color, text, revert=True):
    """Flash the status indicator."""
    dot_canvas.itemconfig(dot_id, fill=color, outline=color)
    lbl_status.config(fg=color, text=text)
    if revert:
        root.after(2000, lambda: reset_status())

def reset_status():
    """Reset status to default."""
    dot_canvas.itemconfig(dot_id, fill=COLOR_ACCENT, outline=COLOR_ACCENT)
    lbl_status.config(fg=COLOR_ACCENT, text="ACTIVE: Awaiting Files")

def drop(event):
    """Handles the file drop event."""
    files = event.data
    file_list = root.tk.splitlist(files)
    
    flash_status("#ffaa00", "PROCESSING...", revert=False)
    root.update()

    log_box.config(state=tk.NORMAL)
    log_box.delete(1.0, tk.END)

    cleaned_count = 0
    skipped_count = 0
    failed_count = 0
    
    for f in file_list:
        status, new_file, error = process_file(f)
        fname = os.path.basename(f)

        if status == "CLEANED":
            cleaned_count += 1
            cleaned_name = os.path.basename(new_file)
            msg = f"✔ {cleaned_name}\n"
            log_box.insert(tk.END, msg, "success")
        elif status == "SKIPPED":
            skipped_count += 1
            msg = f"⊘ {fname} — {error}\n"
            log_box.insert(tk.END, msg, "warning")
        else:
            failed_count += 1
            msg = f"✖ {fname} — {error}\n"
            log_box.insert(tk.END, msg, "fail")
        
        root.update_idletasks()

    summary = f"\nDone: {cleaned_count} cleaned"
    if skipped_count > 0:
        summary += f", {skipped_count} skipped"
    if failed_count > 0:
        summary += f", {failed_count} failed"
    log_box.insert(tk.END, summary, "info")
    log_box.config(state=tk.DISABLED)
    
    if cleaned_count > 0:
        flash_status(COLOR_ACCENT, f"DONE: {cleaned_count} file(s) cleaned")
        winsound.MessageBeep(winsound.MB_OK)
    elif failed_count > 0:
        flash_status("#ff4444", "FAILED")
        winsound.MessageBeep(winsound.MB_ICONHAND)
    else:
        flash_status("#ffaa00", "NOTHING TO DO")

# --- GUI SETUP ---
root = TkinterDnD.Tk()
root.title("XORD MetaPurge")
root.geometry("500x450")
root.configure(bg=COLOR_BG)
root.resizable(False, False)

# === HEADER ===
header_frame = tk.Frame(root, bg=COLOR_BG, pady=15)
header_frame.pack(fill=tk.X)

lbl_title = tk.Label(
    header_frame,
    text="METAPURGE",
    bg=COLOR_BG,
    fg=COLOR_TEXT,
    font=("Segoe UI", 20, "bold")
)
lbl_title.pack()

sub_frame = tk.Frame(header_frame, bg=COLOR_BG)
sub_frame.pack()

lbl_sub1 = tk.Label(
    sub_frame,
    text="The Digital Laundry by ",
    bg=COLOR_BG,
    fg=COLOR_SUBTEXT,
    font=("Segoe UI", 10)
)
lbl_sub1.pack(side=tk.LEFT)

lbl_sub2 = tk.Label(
    sub_frame,
    text="XORD",
    bg=COLOR_BG,
    fg=COLOR_GOLD,
    font=("Segoe UI", 10, "bold")
)
lbl_sub2.pack(side=tk.LEFT)

# === STATUS INDICATOR (CENTERED) ===
status_frame = tk.Frame(root, bg=COLOR_BG, pady=5)
status_frame.pack(fill=tk.X)

status_inner = tk.Frame(status_frame, bg=COLOR_BG)
status_inner.pack()

dot_canvas = tk.Canvas(status_inner, width=16, height=16, bg=COLOR_BG, highlightthickness=0)
dot_canvas.pack(side=tk.LEFT, padx=(0, 8))
dot_id = dot_canvas.create_oval(2, 2, 14, 14, fill=COLOR_ACCENT, outline=COLOR_ACCENT)

lbl_status = tk.Label(
    status_inner,
    text="ACTIVE: Awaiting Files",
    bg=COLOR_BG,
    fg=COLOR_ACCENT,
    font=("Segoe UI", 10)
)
lbl_status.pack(side=tk.LEFT)

# === DROP ZONE ===
drop_frame = tk.Frame(root, bg=COLOR_PANEL, bd=1, relief="solid")
drop_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

lbl_drop = tk.Label(
    drop_frame,
    text="DRAG FILES HERE",
    bg=COLOR_PANEL,
    fg=COLOR_ACCENT,
    font=("Segoe UI", 12)
)
lbl_drop.place(relx=0.5, rely=0.35, anchor="center")

lbl_formats = tk.Label(
    drop_frame,
    text="PDF / JPG / PNG / BMP / TIFF / WEBP / GIF",
    bg=COLOR_PANEL,
    fg=COLOR_SUBTEXT,
    font=("Segoe UI", 9)
)
lbl_formats.place(relx=0.5, rely=0.50, anchor="center")

lbl_location = tk.Label(
    drop_frame,
    text="Cleaned files saved in same location",
    bg=COLOR_PANEL,
    fg="#666666",
    font=("Segoe UI", 8)
)
lbl_location.place(relx=0.5, rely=0.75, anchor="center")

# === LOG AREA ===
log_frame = tk.Frame(root, bg=COLOR_BG)
log_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

log_box = tk.Text(
    log_frame,
    height=6,
    bg="#0d0d0d",
    fg=COLOR_TEXT,
    font=("Consolas", 9),
    relief="flat",
    state=tk.DISABLED,
    insertbackground=COLOR_TEXT
)
log_box.pack(fill=tk.X)

log_box.config(state=tk.NORMAL)
log_box.insert(tk.END, "Ready. Waiting for files...", "info")
log_box.config(state=tk.DISABLED)

log_box.tag_config("success", foreground=COLOR_ACCENT)
log_box.tag_config("warning", foreground="#ffaa00")
log_box.tag_config("fail", foreground="#ff4444")
log_box.tag_config("info", foreground=COLOR_SUBTEXT)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', drop)

root.mainloop()

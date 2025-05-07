import os
import shutil
import subprocess
from datetime import datetime
from tkinter import Tk, Toplevel, Label, Entry, Button, StringVar, messagebox
from tkinter import ttk
import threading
import tkinter
import configparser

# --- Load configuration ---
config = configparser.ConfigParser()
config.read('config.ini')

# Read UNC paths from config
ori_dir   = config.get('network', 'ori_dir')
dst_dir   = config.get('network', 'dst_dir')
merge_dir = config.get('network', 'merge_dir')

# credentials
username = config.get('credentials', 'username')
pwd_a     = config.get('credentials', 'pwd_a')
pwd_b     = config.get('credentials', 'pwd_b')

def map_network_drive_cmd(local_drive, remote_path, username, password):
    cmd = ["net", "use", local_drive, remote_path]
    if username and password:
        cmd.extend([password, "/user:" + username])
    try:
        subprocess.run(cmd, check=True)
        print(f"Drive {local_drive} successfully mapped to {remote_path}")
    except subprocess.CalledProcessError as e:
        print("Error mapping drive:", e)

map_network_drive_cmd("A:", ori_dir, username, pwd_a)
map_network_drive_cmd("B:", dst_dir, username, pwd_b)

files = [os.path.join(ori_dir, f) for f in os.listdir(ori_dir)
         if os.path.isfile(os.path.join(ori_dir, f))]
if not files:
    messagebox.showerror("No Files Found", "No files found in the source directory.")
    exit(1)

latest_file = max(files, key=os.path.getmtime)
print(f"Latest file found: {latest_file}")

# --- Create Main Tkinter Window & Start Mainloop ---
root = Tk()
root.withdraw()

# --- Patch PhotoImage for Thread-Safe Creation ---
from tkinter import PhotoImage as OriginalPhotoImage
def safe_PhotoImage(*args, **kwargs):
    if threading.current_thread() != threading.main_thread():
        event = threading.Event()
        result = {}
        def create_image():
            result["img"] = OriginalPhotoImage(*args, **kwargs)
            event.set()
        root.after(0, create_image)
        event.wait()
        return result["img"]
    else:
        return OriginalPhotoImage(*args, **kwargs)
tkinter.PhotoImage = safe_PhotoImage

# --- Helper to center a window ---
def center_window(window, width, height):
    window.update_idletasks()
    sw, sh = window.winfo_screenwidth(), window.winfo_screenheight()
    x = (sw - width) // 2
    y = (sh - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

def get_dwg_input():
    def validate_4digits(new_value):
        return new_value == "" or (new_value.isdigit() and len(new_value) <= 4)

    top = Toplevel(root)
    top.title("Enter MO Number, Preview & Confirm")
    center_window(top, 1100, 900)

    main_frame = tkinter.Frame(top)
    main_frame.pack(fill="both", expand=True, padx=20, pady=10)

    # --- Input section ---
    input_frame = tkinter.Frame(main_frame)
    input_frame.pack(pady=(0,10))
    for col in (0,2,4): input_frame.grid_columnconfigure(col, weight=0)
    for col in (1,3):   input_frame.grid_columnconfigure(col, weight=1)

    Label(input_frame, text="Year",    font=("TkDefaultFont",10,"bold")) \
        .grid(row=0, column=0, padx=5)
    Label(input_frame, text="Project", font=("TkDefaultFont",10,"bold")) \
        .grid(row=0, column=2, padx=5)
    Label(input_frame, text="Sequence",font=("TkDefaultFont",10,"bold")) \
        .grid(row=0, column=4, padx=5)

    year_var    = StringVar(top)
    project_var = StringVar(top)
    seq_var     = StringVar(top)

    cy   = datetime.now().year
    opts = [f"{str(y)[0]}{str(y)[2:]}" for y in range(cy-2, cy+3)]
    cb   = ttk.Combobox(input_frame, textvariable=year_var,
                        values=opts, state="readonly",
                        width=8, justify="center")
    cb.grid(row=1, column=0, padx=5)
    cb.current(2)

    dash_font = ("TkDefaultFont",15,"bold")
    Label(input_frame, text="-", font=dash_font).grid(row=1, column=1)
    vc = (top.register(validate_4digits), '%P')
    Entry(input_frame, textvariable=project_var, validate="key",
          validatecommand=vc, width=8).grid(row=1, column=2)
    Label(input_frame, text="-", font=dash_font).grid(row=1, column=3)
    Entry(input_frame, textvariable=seq_var, validate="key",
          validatecommand=vc, width=8).grid(row=1, column=4)

    # --- Preview & Confirm section ---
    preview_frame = tkinter.Frame(main_frame)
    preview_frame.pack(fill="both", pady=(10,0), expand=True)

    def build_preview():
        # Clear previous preview widgets
        for w in preview_frame.winfo_children():
            w.destroy()

        # MO Number label
        y = year_var.get() or opts[2]
        p = project_var.get().zfill(4)
        s = seq_var.get().zfill(4)
        Label(preview_frame,
              text=f"MO Number: {y} - {p} - {s}",
              font=("TkDefaultFont",12,"bold"))\
          .pack(pady=5)

        # New file name label
        ext = os.path.splitext(latest_file)[1]
        newname = f"D{y}{p}{s}{ext}"
        Label(preview_frame,
              text=f"New File Name: {newname}",
              font=("TkDefaultFont",10))\
          .pack(pady=5)

        # --- PDF Viewer block restored ---
        try:
            from tkPDFViewer import tkPDFViewer as pdf
        except ImportError:
            messagebox.showerror("Missing Module",
                                 "tkPDFViewer is required for PDF preview.")
            top.destroy()
            return

        pdf_holder = tkinter.Frame(preview_frame)
        pdf_holder.pack(fill="both", expand=True, pady=5)
        viewer = pdf.ShowPdf()
        pdf_widget = viewer.pdf_view(
            pdf_holder,
            pdf_location=latest_file,
            width=120,   # adjust these to taste
            height=40
        )
        pdf_widget.pack(fill="both", expand=True)

        # Confirm/Cancel buttons
        btn_frame = tkinter.Frame(preview_frame)
        btn_frame.pack(pady=10)
        Button(btn_frame, text="Confirm", command=on_confirm)\
            .grid(row=0, column=0, padx=5)
        Button(btn_frame, text="Cancel",  command=on_cancel)\
            .grid(row=0, column=1, padx=5)

    def on_confirm():
        # Read fresh values
        y = year_var.get() or opts[2]
        p = project_var.get().zfill(4)
        s = seq_var.get().zfill(4)

        ext = os.path.splitext(latest_file)[1]
        fn  = f"D{y}{p}{s}{ext}"
        dest = os.path.join(dst_dir, fn)
        i = 1
        while os.path.exists(dest):
            fn   = f"D{y}{p}{s}-R{i:03}{ext}"
            dest = os.path.join(dst_dir, fn)
            i  += 1

        shutil.copy2(latest_file, dest)
        messagebox.showinfo("File Renamed", f"Copied to:\n{dest}")
        top.confirmed = True
        top.destroy()

    def on_cancel():
        top.destroy()

    # initial render
    build_preview()

    # if you want live‑update as you type:
    project_var.trace_add('write', lambda *a: build_preview())
    seq_var.trace_add('write',     lambda *a: build_preview())

    top.grab_set()
    top.wait_window()
    if getattr(top, 'confirmed', False):
        return True
    return None



# --- Main flow ---
def main_flow():
    if not get_dwg_input():
        messagebox.showerror("Cancelled", "No DWG input received or operation cancelled.")
    root.destroy()

root.after(0, main_flow)
root.mainloop()

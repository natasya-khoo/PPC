import os
import shutil
import subprocess
from datetime import datetime
from tkinter import (Tk, Toplevel, Label, Entry, Button, StringVar,
                     messagebox, Listbox, ttk)
import threading
import tkinter
import configparser

# --- Load configuration ---
config = configparser.ConfigParser()
config.read('config.ini')

# Read UNC paths from config
i_ori_dir   = config.get('network', 'ori_dir')
i_dst_dir   = config.get('network', 'dst_dir')
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
        print(f"Drive {local_drive} mapped to {remote_path}")
    except subprocess.CalledProcessError as e:
        print("Error mapping drive:", e)

map_network_drive_cmd("A:", i_ori_dir, username, pwd_a)
map_network_drive_cmd("B:", i_dst_dir, username, pwd_b)

# gather files in source directory
files = [os.path.join(i_ori_dir, f) for f in os.listdir(i_ori_dir)
         if os.path.isfile(os.path.join(i_ori_dir, f))]
if not files:
    messagebox.showerror("No Files Found", "No files in source directory.")
    exit(1)

# --- Tkinter setup ---
root = Tk()
root.withdraw()

# thread-safe PhotoImage patch
from tkinter import PhotoImage as OriginalPhotoImage

def safe_PhotoImage(*args, **kwargs):
    if threading.current_thread() is not threading.main_thread():
        event = threading.Event(); result = {}
        def create():
            result['img'] = OriginalPhotoImage(*args, **kwargs)
            event.set()
        root.after(0, create); event.wait()
        return result['img']
    return OriginalPhotoImage(*args, **kwargs)

tkinter.PhotoImage = safe_PhotoImage

# center-window helper
def center_window(window, w, h):
    window.update_idletasks()
    sw, sh = window.winfo_screenwidth(), window.winfo_screenheight()
    x = (sw - w) // 2; y = (sh - h) // 2
    window.geometry(f"{w}x{h}+{x}+{y}")

# --- Main dialog with sidebar ---
def get_dwg_input():
    def validate_digits(val):
        return val == '' or (val.isdigit() and len(val) <= 4)

    top = Toplevel(root)
    top.title("Enter MO Number, Preview & Confirm")
    center_window(top, 1200, 900)

    # --- NEW: use a PanedWindow so the sidebar is resizable by dragging its sash ---
    paned = ttk.PanedWindow(top, orient='horizontal')
    paned.pack(fill='both', expand=True, padx=10, pady=10)

    sidebar = tkinter.Frame(paned, width=250, bg='#f0f0f0')
    content = tkinter.Frame(paned)
    paned.add(sidebar, weight=1)   # weight=1 makes sidebar take a proportional share
    paned.add(content, weight=4)   # weight=4 makes content larger by default

    # --- NEW: sidebar file list with its own vertical scrollbar ---
    sb_scroll = tkinter.Scrollbar(sidebar, orient='vertical')
    lb = Listbox(sidebar, yscrollcommand=sb_scroll.set)
    sb_scroll.config(command=lb.yview)
    sb_scroll.pack(side='right', fill='y', padx=(0,5), pady=5)
    lb.pack(side='left', fill='both', expand=True, padx=(5,0), pady=5)

    for fpath in files:
        lb.insert('end', os.path.basename(fpath))
    lb.selection_set(0)

    selected_file = {'path': files[0]}
    def on_select(evt):
        idx = evt.widget.curselection()
        if not idx: return
        selected_file['path'] = files[idx[0]]
        build_preview()
    lb.bind('<<ListboxSelect>>', on_select)

    # --- Content: inputs + preview ---
    main_frame = tkinter.Frame(content)
    main_frame.pack(fill='both', expand=True)

    # Input row
    input_frame = tkinter.Frame(main_frame)
    input_frame.pack(pady=(0,10))
    for c in (0,2,4): input_frame.grid_columnconfigure(c, weight=0)
    for c in (1,3):   input_frame.grid_columnconfigure(c, weight=1)

    Label(input_frame, text="Year",    font=("TkDefaultFont",10,"bold"))
    Label(input_frame, text="Project", font=("TkDefaultFont",10,"bold"))
    Label(input_frame, text="Sequence",font=("TkDefaultFont",10,"bold"))
    Label(input_frame, text="Year").grid(row=0,column=0,padx=5)
    Label(input_frame, text="Project").grid(row=0,column=2,padx=5)
    Label(input_frame, text="Sequence").grid(row=0,column=4,padx=5)

    year_var    = StringVar(top)
    project_var = StringVar(top)
    seq_var     = StringVar(top)

    cy   = datetime.now().year
    opts = [f"{str(y)[0]}{str(y)[2:]}" for y in range(cy-2, cy+3)]
    cb   = ttk.Combobox(input_frame, textvariable=year_var, values=opts,
                        state="readonly", width=8, justify="center")
    cb.grid(row=1,column=0,padx=5); cb.current(2)

    dash = ("TkDefaultFont",15,"bold")
    Label(input_frame, text="-", font=dash).grid(row=1,column=1)
    vc = (top.register(validate_digits), '%P')
    Entry(input_frame, textvariable=project_var, validate="key",
          validatecommand=vc, width=8).grid(row=1,column=2)
    Label(input_frame, text="-", font=dash).grid(row=1,column=3)
    Entry(input_frame, textvariable=seq_var, validate="key",
          validatecommand=vc, width=8).grid(row=1,column=4)

    # Preview area
    preview_frame = tkinter.Frame(main_frame)
    preview_frame.pack(fill='both', expand=True)

    def build_preview():
        # clear
        for w in preview_frame.winfo_children(): w.destroy()

        # MO Number label
        y = year_var.get() or opts[2]
        p = project_var.get().zfill(4)
        s = seq_var.get().zfill(4)
        Label(preview_frame,
              text=f"MO Number: {y} - {p} - {s}",
              font=("TkDefaultFont",12,"bold"))
        Label(preview_frame,
              text=f"MO Number: {y} - {p} - {s}").pack(pady=5)

        # Filename preview
        ext = os.path.splitext(selected_file['path'])[1]
        newname = f"D{y}{p}{s}{ext}"
        Label(preview_frame,
              text=f"New File Name: {newname}",
              font=("TkDefaultFont",10)).pack(pady=5)

        # PDF Viewer
        try:
            from tkPDFViewer import tkPDFViewer as pdf
        except ImportError:
            messagebox.showerror("Missing Module",
                                 "tkPDFViewer is required.")
            top.destroy(); return

        holder = tkinter.Frame(preview_frame)
        holder.pack(fill='both', expand=True, pady=5)
        viewer = pdf.ShowPdf()
        pdf_widget = viewer.pdf_view(
            holder,
            pdf_location=selected_file['path'],
            width=120, height=40
        )
        pdf_widget.pack(fill='both', expand=True)

        # buttons
        btns = tkinter.Frame(preview_frame)
        btns.pack(pady=10)
        Button(btns, text="Confirm", command=on_confirm).grid(row=0,column=0,padx=5)
        Button(btns, text="Cancel",  command=on_cancel).grid(row=0,column=1,padx=5)

    def on_confirm():
        y = year_var.get() or opts[2]
        p = project_var.get().zfill(4)
        s = seq_var.get().zfill(4)
        ext = os.path.splitext(selected_file['path'])[1]
        fn  = f"D{y}{p}{s}{ext}"
        dest = os.path.join(i_dst_dir, fn)
        i = 1
        while os.path.exists(dest):
            fn   = f"D{y}{p}{s}-R{i:03}{ext}"
            dest = os.path.join(i_dst_dir, fn)
            i  += 1
        shutil.copy2(selected_file['path'], dest)
        messagebox.showinfo("File Renamed", f"Copied:{dest}")
        top.confirmed = True; top.destroy()

    def on_cancel():
        top.destroy()

    # initial render
    build_preview()
    project_var.trace_add('write', lambda *a: build_preview())
    seq_var.trace_add('write',     lambda *a: build_preview())

    top.grab_set()
    top.wait_window()
    return getattr(top, 'confirmed', False)

# --- Main flow ---
def main_flow():
    if not get_dwg_input():
        messagebox.showerror("Cancelled", "Operation cancelled.")
    root.destroy()

root.after(0, main_flow)
root.mainloop()

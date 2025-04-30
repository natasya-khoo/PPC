import os
import shutil
import subprocess
from datetime import datetime
from tkinter import Tk, Toplevel, Label, Entry, Button, StringVar, messagebox
from tkinter import ttk
import threading
import tkinter

# --- Network Drive Mapping & File Handling ---

ori_dir = r"\\172.16.20.13\Share Folder\Level3"
dst_dir = r"\\172.16.30.120\SVR-Drive\CANSG\DWG"
merge_dir = r"\\192.168.4.163\SVR-Drive\CANSG\DWG"

def map_network_drive_cmd(local_drive, remote_path, username, password):
    cmd = ["net", "use", local_drive, remote_path]
    if username and password:
        cmd.extend([password, "/user:" + username])
    try:
        subprocess.run(cmd, check=True)
        print(f"Drive {local_drive} successfully mapped to {remote_path}")
    except subprocess.CalledProcessError as e:
        print("Error mapping drive:", e)

map_network_drive_cmd("A:", ori_dir, "cantal", "eYlvK72e")
map_network_drive_cmd("B:", dst_dir, "cantal", "123456")

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

# --- Helper Function to Center a Window on Screen ---
def center_window(window, width, height):
    window.update_idletasks()  # Update "requested size" from geometry manager
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

# --- Custom Dialog for DWG Input, Integrated PDF Preview & Confirmation ---
def get_dwg_input():
    # Validation callback: allow empty string or up to 4 digits only.
    def validate_4digits(new_value):
        if new_value == "":
            return True
        return new_value.isdigit() and len(new_value) <= 4

    # Create one Toplevel window for the entire dialog.
    top = Toplevel(root)
    top.title("Enter MO Number, Preview & Confirm")
    # Set the Toplevel to a reasonably large size and center it.
    center_window(top, 1100, 900)

    main_frame = tkinter.Frame(top)
    main_frame.pack(fill="both", expand=True, padx=20, pady=10)

    # ---------- Section 1: MO Number Input (Top Section) ----------
    input_frame = tkinter.Frame(main_frame)
    input_frame.pack(fill="x", pady=(0, 10))
    
    # Configure grid columns so that the inputs are centered.
    input_frame.grid_columnconfigure(0, weight=1)
    input_frame.grid_columnconfigure(1, weight=0)
    input_frame.grid_columnconfigure(2, weight=1)
    input_frame.grid_columnconfigure(3, weight=0)
    input_frame.grid_columnconfigure(4, weight=1)

    # Header row: Labels above each input.
    Label(input_frame, text="Year", font=("TkDefaultFont", 10, "bold"), anchor="center") \
        .grid(row=0, column=0, padx=5, pady=5)
    Label(input_frame, text="Project", font=("TkDefaultFont", 10, "bold"), anchor="center") \
        .grid(row=0, column=2, padx=5, pady=5)
    Label(input_frame, text="Sequence", font=("TkDefaultFont", 10, "bold"), anchor="center") \
        .grid(row=0, column=4, padx=5, pady=5)

    # Input row: MO number entries with bold dash separators.
    year_var = StringVar(top)
    current_year = datetime.now().year
    start_year = current_year - 2
    end_year = current_year + 2
    year_options = [str(year)[0] + str(year)[2:] for year in range(start_year, end_year + 1)]
    combobox_year = ttk.Combobox(input_frame, textvariable=year_var, values=year_options,
                                 state="readonly", width=20, justify="center")
    combobox_year.grid(row=1, column=0, padx=5, pady=5)
    current_modified = str(current_year)[0] + str(current_year)[2:]
    combobox_year.current(year_options.index(current_modified))

    dash_font = ("TkDefaultFont", 15, "bold")
    Label(input_frame, text="-", font=dash_font).grid(row=1, column=1, padx=10, pady=5)

    validate_cmd = (top.register(validate_4digits), '%P')
    project_var = StringVar(top)
    entry_project = Entry(input_frame, textvariable=project_var, validate="key",
                          validatecommand=validate_cmd, width=20)
    entry_project.grid(row=1, column=2, padx=5, pady=5)

    Label(input_frame, text="-", font=dash_font).grid(row=1, column=3, padx=10, pady=5)
    seq_var = StringVar(top)
    entry_seq = Entry(input_frame, textvariable=seq_var, validate="key",
                      validatecommand=validate_cmd, width=20)
    entry_seq.grid(row=1, column=4, padx=5, pady=5)

    # ---------- Section 2: Submit Button (to validate MO Number) ----------
    # Create the submit button with a widget reference.
    def on_submit():
        selected_year = year_var.get()
        project = project_var.get()
        seq = seq_var.get()
        if not (project.isdigit() and seq.isdigit()):
            messagebox.showerror("Invalid input", "Project and sequence numbers must be numeric.")
            return
        top.selected_year = selected_year
        top.dwg_project = project.zfill(4)
        top.dwg_seq = seq.zfill(4)
        submit_button.grid_forget()
        dwg_number = f"D{selected_year}{project.zfill(4)}{seq.zfill(4)}"
        _, file_extension = os.path.splitext(latest_file)
        new_file_name = f"{dwg_number}{file_extension}"
        # Build the Preview & Confirmation section.
        build_preview_section(new_file_name)

    submit_button = Button(input_frame, text="Submit", command=on_submit)
    submit_button.grid(row=2, column=0, columnspan=5, pady=10)

    # ---------- Section 3: Preview & Confirmation (Middle Section) ----------
    # This frame will be displayed once the MO input is submitted.
    preview_frame = tkinter.Frame(main_frame)

    def build_preview_section(new_file_name):
        for widget in preview_frame.winfo_children():
            widget.destroy()
        preview_frame.pack(fill="both", pady=(10, 0))
        
        # Display the MO Number in the middle section.
        mo_str = f"{top.selected_year} - {top.dwg_project} - {top.dwg_seq}"
        Label(preview_frame, text=f"MO Number: {mo_str}",
              font=("TkDefaultFont", 12, "bold")).pack(pady=5)
        
        # Optionally, display the new file name.
        Label(preview_frame, text=f"New File Name: {new_file_name}",
              font=("TkDefaultFont", 10)).pack(pady=5)
        
        # PDF Preview Section.
        try:
            from tkPDFViewer import tkPDFViewer as pdf
        except ImportError:
            messagebox.showerror("Missing Module", "tkPDFViewer is required for PDF preview.")
            top.destroy()
            return
        pdf_frame = tkinter.Frame(preview_frame)
        pdf_frame.pack(pady=5)
        viewer = pdf.ShowPdf()
        pdf_display = viewer.pdf_view(pdf_frame, pdf_location=latest_file, width=120, height=40)
        pdf_display.pack(anchor="center", pady=5)
        
        # Confirmation Buttons.
        button_frame = tkinter.Frame(preview_frame)
        button_frame.pack(pady=10)
        Button(button_frame, text="Confirm", command=on_confirm) \
            .pack(side="left", padx=10)
        Button(button_frame, text="Cancel", command=on_cancel) \
            .pack(side="left", padx=10)

    def on_confirm():
        top.confirmed = True
        top.destroy()

    def on_cancel():
        top.destroy()

    top.grab_set()
    top.wait_window(top)

    if hasattr(top, "confirmed") and top.confirmed:
        return getattr(top, 'selected_year', None), getattr(top, 'dwg_project', None), getattr(top, 'dwg_seq', None)
    else:
        return None

# --- Main Flow Function ---
def main_flow():
    mo_data = get_dwg_input()
    if not mo_data:
        messagebox.showerror("Cancelled", "No DWG input received or operation cancelled.")
        root.destroy()
        return

    selected_year, project, sequence = mo_data

    prefix = f"D{selected_year}"
    dwg_number = f"{prefix}{project}{sequence}"
    _, file_extension = os.path.splitext(latest_file)
    new_file_name = f"{dwg_number}{file_extension}"
    destination_file_path = os.path.join(dst_dir, new_file_name)

    counter = 1
    while os.path.exists(destination_file_path):
        new_file_name = f"{dwg_number}-R{counter:03}{file_extension}"
        destination_file_path = os.path.join(dst_dir, new_file_name)
        counter += 1

    shutil.copy2(latest_file, destination_file_path)
    print(f"File copied and renamed to: {destination_file_path}")
    messagebox.showinfo("File Renamed", f"File copied and renamed to:\n{destination_file_path}")

root.after(0, main_flow)
root.mainloop()

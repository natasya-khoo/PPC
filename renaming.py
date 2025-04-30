import os
import shutil
from tkinter import Tk, Toplevel, Label, Entry, Button, StringVar, messagebox
from tkinter import ttk  # Import ttk for better themed widgets
import subprocess
from datetime import datetime

# --- Existing setup ---
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

map_network_drive_cmd("A:", ori_dir , "cantal", "eYlvK72e")
map_network_drive_cmd("B:", dst_dir, "cantal", "123456")

# --- File Handling ---
files = [os.path.join(ori_dir, f) for f in os.listdir(ori_dir)
         if os.path.isfile(os.path.join(ori_dir, f))]

if not files:
    print("No files found in the source directory.")
    exit(1)

latest_file = max(files, key=os.path.getmtime)
print(f"Latest file found: {latest_file}")

# --- GUI Setup ---
root = Tk()
root.withdraw()

# --- Custom Dialog with ttk.Combobox for the Year ---
def get_dwg_input():
    def on_submit():
        # Retrieve the selected year and other input values.
        selected_year = year_var.get()
        project = project_var.get()
        seq = seq_var.get()
        if not (project.isdigit() and seq.isdigit()):
            messagebox.showerror("Invalid input", "Project and sequence numbers must be numeric.")
            return
        # Store values in the dialog window attributes.
        top.selected_year = selected_year
        top.dwg_project = project.zfill(3)  # Ensure at least three digits for project
        top.dwg_seq = seq.zfill(4)          # Ensure four digits for sequence
        top.destroy()

    top = Toplevel()
    top.title("Enter MO Number")

    # Labels for each field.
    Label(top, text="Year:").grid(row=0, column=0, padx=10, pady=5)
    Label(top, text="Project Number (e.g. 148):").grid(row=1, column=0, padx=10, pady=5)
    Label(top, text="Sequence Number (e.g. 1):").grid(row=2, column=0, padx=10, pady=5)

    # --- Dropdown using ttk.Combobox ---
    year_var = StringVar(top)
    current_year = datetime.now().year  # e.g., 2025
    start_year = current_year - 2         # 2023
    end_year = current_year + 2           # 2027
    # Build the list of modified years. For example: 2023 becomes "223".
    year_options = [str(year)[0] + str(year)[2:] for year in range(start_year, end_year + 1)]
    
    # Create a readonly combobox for the year dropdown.
    combobox = ttk.Combobox(top, textvariable=year_var, values=year_options, state="readonly", width=17)
    combobox.grid(row=0, column=1, padx=10, pady=5)
    # Set the current year as the default selection.
    current_modified = str(current_year)[0] + str(current_year)[2:]
    combobox.current(year_options.index(current_modified))

    # --- Entry for Project Number ---
    project_var = StringVar(top)
    Entry(top, textvariable=project_var).grid(row=1, column=1, padx=10, pady=5)

    # --- Entry for Sequence Number ---
    seq_var = StringVar(top)
    Entry(top, textvariable=seq_var).grid(row=2, column=1, padx=10, pady=5)

    Button(top, text="Submit", command=on_submit).grid(row=3, columnspan=2, pady=10)

    top.grab_set()
    root.wait_window(top)
    return getattr(top, 'selected_year', None), getattr(top, 'dwg_project', None), getattr(top, 'dwg_seq', None)



# Get user input from the dialog.
selected_year, project, sequence = get_dwg_input()
if not selected_year or not project or not sequence:
    messagebox.showerror("Cancelled", "No DWG input received.")
    root.destroy()
    exit()

# Build the DWG number using the selected dropdown year.
prefix = f"MS{selected_year}"
dwg_number = f"{prefix}{project}{sequence}"

# --- File Copy with Name Handling ---
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
root.destroy()

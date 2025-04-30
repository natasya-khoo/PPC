import subprocess

def map_network_drive_cmd(local_drive, remote_path, username="cantal", password="eYlvK72e"):
    cmd = ["net", "use", local_drive, remote_path]
    if username and password:
        cmd.extend([password, "/user:" + username])
    try:
        subprocess.run(cmd, check=True)
        print(f"Drive {local_drive} successfully mapped to {remote_path}")
    except subprocess.CalledProcessError as e:
        print("Error mapping drive:", e)

# Example usage:
map_network_drive_cmd("Z:", r"\\172.16.20.13\Share Folder\Printer")
#map_network_drive_cmd("Z:", r"\\172.16.30.120\Share Folder\Printer")

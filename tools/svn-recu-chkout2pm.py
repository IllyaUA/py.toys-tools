import configparser
import os
import stat
import subprocess
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import filedialog, scrolledtext

import psutil

MONITORING_DATA_PATH = os.path.expanduser(r"~\AppData\Roaming\TortoiseSVN\MonitoringData.ini")

def get_svn_subfolders(repo_url):
    """Retrieve the list of immediate subdirectories in an SVN repository using 'svn list --xml'."""
    try:
        result = subprocess.run(["svn", "list", "--xml", repo_url], capture_output=True, text=True, check=True)
        root = ET.fromstring(result.stdout)
        return [entry.find('name').text for entry in root.findall("./list/entry") if entry.get('kind') == 'dir']
    except subprocess.CalledProcessError as e:
        log_message(f"Error accessing {repo_url}: {e}")
        return []

def normalize_path(path):
    """Ensure the path uses the correct OS-specific separator."""
    return os.path.normpath(path)

def remove_readonly_flag(file_path):
    """Remove readonly flag from a file if it is set."""
    if os.path.exists(file_path):
        os.chmod(file_path, stat.S_IWRITE)


def close_project_monitor():
    """Close all running instances of TortoiseSVN Project Monitor."""
    found = False  # Track if any processes were found

    for process in psutil.process_iter(attrs=['pid', 'name']):
        if process.info['name'] == "TSVNProjectMonitor.exe":
            found = True
            log_message(f"Closing TortoiseSVN Project Monitor (PID: {process.info['pid']})...")
            process.terminate()  # Graceful termination

            try:
                process.wait(timeout=5)  # Wait for process to exit
                log_message(f"TortoiseSVN Project Monitor (PID: {process.info['pid']}) closed.")
            except psutil.TimeoutExpired:
                log_message(f"Process {process.info['pid']} did not close in time. Forcing shutdown...")
                process.kill()  # Force kill
                log_message(f"TortoiseSVN Project Monitor (PID: {process.info['pid']}) forcibly closed.")

    if not found:
        log_message("✅ No running instances of TortoiseSVN Project Monitor found.")


def add_to_tortoise_project_monitor(local_checkout_path, repo_root, repo_url):
    """Append the new SVN working copy to TortoiseSVN's MonitoringData.ini."""
    if not os.path.exists(MONITORING_DATA_PATH):
        print(f"Error: MonitoringData.ini not found at {MONITORING_DATA_PATH}")
        return
    remove_readonly_flag(MONITORING_DATA_PATH)
    close_project_monitor()
    config = configparser.RawConfigParser()
    config.read(MONITORING_DATA_PATH)
    print(MONITORING_DATA_PATH)

    # Find the next available item number
    existing_items = [section for section in config.sections() if section.startswith("item_")]
    next_item_number = (
        max((int(item.split("_")[1]) for item in existing_items), default=0)
        + 1
    )
    # next_item_number = max([int(item.split("_")[1]) for item in existing_items], default=0) + 1
    new_section = f"item_{next_item_number:03}"  # Ensures item naming is consistent

    # Check if the project is already in MonitoringData.ini
    for section in existing_items:
        if config.get(section, "wcPathOrUrl", fallback="") == local_checkout_path:
            print(f"Project {local_checkout_path} already exists in Project Monitor.")
            return

    # Add new section
    project_name = f"_{os.path.basename(local_checkout_path)}"
    config[new_section] = {
        "root": repo_root,
        "wcPathOrUrl": local_checkout_path,
        "pp_sUrl": repo_url,
        "Name": project_name,
        "interval": "30",
        "unreadItems": "0",
    }

    # Write the updated configuration
    with open(MONITORING_DATA_PATH, "w") as configfile:
        config.write(configfile, space_around_delimiters=False)

    print(f"✅ Added to TortoiseSVN Monitor: {local_checkout_path}")


def checkout_trunks():
    """Recursively checks out only 'trunk' folders from each project and adds them to Project Monitor."""
    repo_url = repo_entry.get()
    local_path = local_path_entry.get()
    os.makedirs(local_path, exist_ok=True)

    subfolders = get_svn_subfolders(repo_url)
    if "trunk" in subfolders:
        trunk_url = f"{repo_url}/trunk"
        local_checkout_path = normalize_path(os.path.join(local_path, os.path.basename(repo_url)))
        log_message(f"Checking out trunk: {trunk_url} to {local_checkout_path}")
        subprocess.run(["svn", "checkout", trunk_url, local_checkout_path])
        add_to_tortoise_project_monitor(local_checkout_path, repo_url, trunk_url)
    else:
        for subfolder in subfolders:
            trunk_url = f"{repo_url}/{subfolder}/trunk"
            local_checkout_path = normalize_path(os.path.join(local_path, subfolder))

            if 'trunk' in get_svn_subfolders(f"{repo_url}/{subfolder}"):
                log_message(f"Checking out trunk: {trunk_url} to {local_checkout_path}")
                subprocess.run(["svn", "checkout", trunk_url, local_checkout_path])
                add_to_tortoise_project_monitor(local_checkout_path, repo_url, trunk_url)
            else:
                log_message(f"Skipping {subfolder}, no 'trunk' found.")


def browse_folder():
    folder_selected = filedialog.askdirectory()
    local_path_entry.delete(0, tk.END)
    local_path_entry.insert(0, folder_selected)


def log_message(message):
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)


# GUI Setup
root = tk.Tk()
root.title("SVN Trunk Checkout")
root.geometry("600x400")

tk.Label(root, text="SVN Repository URL:").pack()
repo_entry = tk.Entry(root, width=80)
repo_entry.insert(0, "http://subversion.name.local/svn/Repo")
repo_entry.pack()

tk.Label(root, text="Local Checkout Folder:").pack()
local_path_frame = tk.Frame(root)
local_path_frame.pack()
local_path_entry = tk.Entry(local_path_frame, width=60)
local_path_entry.pack(side=tk.LEFT)
browse_button = tk.Button(local_path_frame, text="Browse", command=browse_folder)
browse_button.pack(side=tk.RIGHT)

tk.Button(root, text="Proceed", command=checkout_trunks).pack()

log_text = scrolledtext.ScrolledText(root, height=15, width=70)
log_text.pack()

if __name__ == "__main__":
    root.mainloop()
    print("🔄 Restart TortoiseSVN Project Monitor to apply changes.")
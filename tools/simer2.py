import os
import time
import random
import ctypes
import win32gui
import win32con
import pythoncom
import pyautogui
import subprocess
import threading
import tkinter as tk
import win32com.client
from pynput import keyboard
from tkinter import scrolledtext

#print(pythoncom.__file__)

# Constants to prevent sleep/lock
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002


running = False
user_detection_enabled = False
script_path = r"C:\temp\script.py"

def prevent_sleep():
    """Prevent Windows from sleeping."""
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)

def allow_sleep():
    """Allow Windows to sleep."""
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

def log_message(message):
    log_field.config(state=tk.NORMAL)
    log_field.insert(tk.END, message + "\n")
    log_field.config(state=tk.DISABLED)
    log_field.yview(tk.END)

def load_snippets():
    try:
        with open("dblib.py", "r") as f:
            lines = [line.strip() for line in f if line.strip()]
        return lines
    except Exception as e:
        log_message(f"Error loading snippets: {e}")
        return ["Default snippet: dblib.py not found or empty."]

snippets = load_snippets()

def open_script_in_pycharm(script_path):
    """Open the script in PyCharm."""
    with open(script_path, 'w') as f:
        f.write("# New Python script\n")
    try:
        subprocess.Popen(f'start pycharm "{script_path}"', shell=True)
        log_message(f"Opened script in PyCharm: {script_path}")
    except Exception as e:
        log_message(f"Error opening script: {e}")

def get_or_create_word_doc():
    """Find an open Word document named 'py.simer', or create a new one if it doesn't exist."""
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True

    # Check if "py.simer" is already open
    for doc in word.Documents:
        if doc.Name.lower() == "py.simer.docx":
            log_message("Using existing Word document: py.simer")
            doc.Application.Activate()
            bring_window_to_front("WINWORD.EXE")
            return doc

    # Create new document if not found
    doc = word.Documents.Add()
    doc.SaveAs2(os.path.join(os.environ["TEMP"], "py.simer.docx"))  # Save in temp directory
    doc.Application.Activate()
    bring_window_to_front("WINWORD.EXE")
    log_message("Created new Word document: py.simer")
    return doc

def bring_window_to_front(process_name="WINWORD.EXE"):
    """Finds the MS Word window and brings it to the foreground."""
    def enum_window_callback(hwnd, result):
        """Callback to find the right window by process name."""
        if win32gui.IsWindowVisible(hwnd) and process_name.lower() in win32gui.GetWindowText(hwnd).lower():
            result.append(hwnd)

    hwnd_list = []
    win32gui.EnumWindows(enum_window_callback, hwnd_list)

    if hwnd_list:
        hwnd = hwnd_list[0]  # Get the first matching window

        # Restore if minimized
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        # Bring to foreground
        win32gui.SetForegroundWindow(hwnd)
        log_message(f"Brought {process_name} to front.")
    else:
        log_message(f"Could not find window: {process_name}")

def random_mouse_movement():
    """Move the mouse in a curved path with randomness."""
    current_x, current_y = pyautogui.position()
    x_end = random.randint(0, pyautogui.size()[0])
    y_end = random.randint(0, pyautogui.size()[1])
    curve_x = random.randint(-100, 100)
    curve_y = random.randint(-100, 100)
    for i in range(10):
        x = current_x + (x_end - current_x) * i / 10 + random.uniform(-curve_x, curve_x)
        y = current_y + (y_end - current_y) * i / 10 + random.uniform(-curve_y, curve_y)
        pyautogui.moveTo(x, y, duration=random.uniform(0.1, 0.5))


def random_typing():
    """Type text character-by-character with random delays.
       Move the mouse between snippets, not during typing."""
    global running, user_detection_enabled
    pythoncom.CoInitialize()
    try:
        prevent_sleep()
        user_detection_enabled = True
        doc = get_or_create_word_doc()
        time.sleep(2)
        #doc.Content.InsertAfter("Started\n")
        log_message("Automation started.")

        # Disable PyAutoGUI fail-safe temporarily
        pyautogui.FAILSAFE = False

        while running:
            snippet = random.choice(snippets)
            # log_message(f"Typing: {snippet}")
            for char in snippet:
                pyautogui.write(char, interval=random.uniform(0.1, 0.4))
            pyautogui.write("\n")
            time.sleep(random.randint(1, 3))
            # Move the mouse only after finishing typing the snippet.
            random_mouse_movement()

    finally:
        # Re-enable PyAutoGUI fail-safe
        pyautogui.FAILSAFE = True
        allow_sleep()
        pythoncom.CoUninitialize()
        log_message("Automation stopped.")



def on_key_press(key):
    """Stop only if ESC or SPACE is pressed."""
    global running, user_detection_enabled
    if user_detection_enabled:
        if key == keyboard.Key.esc: # or key == keyboard.Key.space:
            running = False
            user_detection_enabled = False
            toggle_btn.config(text="Start", bg="white")
            log_message("ESC or SPACE detected! Stopping automation.")

def toggle_automation():
    """Start or stop simulation."""
    global running
    if not running:
        running = True
        toggle_btn.config(text="Stop", bg="gray")
        threading.Thread(target=random_typing, daemon=True).start()
        open_script_in_pycharm(script_path)
        log_message("Starting automation...")
    else:
        running = False
        toggle_btn.config(text="Start", bg="white")
        log_message("Stopping automation...")

# GUI Setup
root = tk.Tk()
root.title("MS Word Automation")
root.geometry("+0+0")  # Floating window at (0,0)
root.attributes('-topmost', True)  # Always on top

toggle_btn = tk.Button(root, text="Start", command=toggle_automation,
                       width=20, height=2, bg="white")
toggle_btn.pack(pady=10)

# Scrollable Log Window
log_frame = tk.Frame(root)
log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
log_field = scrolledtext.ScrolledText(root, height=5, width=40, state=tk.DISABLED)
log_field.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Keyboard listener for ESC & SPACE only
keyboard_listener = keyboard.Listener(on_press=on_key_press)
keyboard_listener.start()

root.mainloop()

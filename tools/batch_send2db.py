import json
import logging
import threading
import tkinter as tk
from datetime import datetime
# from typing import Any, Tuple, cast
from tkinter import filedialog, messagebox, ttk, scrolledtext

import jsonlines
import requests
import sqlparse
from requests.auth import HTTPBasicAuth

import dblib

# Dummy credentials
USERNAME = "usrname"
PASSWORD = "usrpwd"

# Example JSON record
json_record = {
    "sn": "51202410091259960000013617609",
    "origin": "zollner",
    "ItemNr": "95666",
    "WorkPlan": "A00864",
    "WorkPlanIndex": 1,
    "bom_name": "95666",
    "bom_index": 0,
    "serie": "none",
    "ffmaterial": "2053611-02",
    "ffauftrag": "13617609",
    "timestamp": "2024-10-15T16:09:17",
    "verification": True,
    "comment": "",
    "status": "Active",
    "testinfo": [
        {
            "id": "InspStatusReady",
            "description": "Endprüfung bestanden",
            "limits": [1.0, 1.0],
            "value": True,
            "unit": "---",
            "ok": True
        }
    ]
}



def log_query(query, params):
    """
    Logs the query with parameters substituted for debugging.
    """
    formatted_query = query
    for key, value in params.items():
        # Format values correctly for SQL (strings need quotes)
        if isinstance(value, str):
            value = f"'{value}'"
        elif isinstance(value, bool):
            value = "TRUE" if value else "FALSE"
        elif value is None:
            value = "NULL"
        formatted_query = formatted_query.replace(f"%({key})s", str(value))

    logging.debug("Executing query:\n" + sqlparse.format(formatted_query, reindent=True))


class JSONToDBApp:
    def __init__(self):
        self.root = tk.Tk()  # Create the root Tkinter window
        self.root.title("JSON to DB App")
        self.root.geometry("600x300")

        # Load the dbconfig once
        self.dbconfig = dblib.load_config()  # Load DB configuration
        self.db_pool = None
        if not self.dbconfig:
            messagebox.showerror("Error", "Failed to load database configuration.")
            self.root.quit()  # Exit the app if the config loading fails

        # Label for file selection
        self.file_label = tk.Label(self.root, text="No file selected")
        self.file_label.pack(pady=10)

        # Button to open file dialog
        self.select_button = tk.Button(self.root, text="Select JSON File", command=self.select_file)
        self.select_button.pack(pady=5)

        # Button to send data to DB
        self.send_button = tk.Button(self.root, text="Send to DB", command=self.send_data, state=tk.DISABLED)
        self.send_button.pack(pady=10)

        # Button: Send to JSONL Server
        self.send_jsonl_button = tk.Button(self.root, text="Send to JSONL Server", command=self.send_to_jsonl_server,
                                           state=tk.DISABLED)
        self.send_jsonl_button.pack(pady=5)

        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=5)

        # Abort button (initially disabled)
        self.abort_button = tk.Button(self.root, text="Abort sending", command=self.abort_sending, state=tk.DISABLED)
        self.abort_button.pack(pady=5)

        # Frame to group time labels
        self.time_frame = tk.Frame(self.root)
        self.time_frame.pack(pady=5)
        self.start_time = None
        self.process_running = False

        self.time_start_label = tk.Label(self.time_frame, text="Start Time: 00:00:00")
        self.time_start_label.pack(side="left", padx=5)

        self.time_elapsed_label = tk.Label(self.time_frame, text="Elapsed: 00:00:00.0")
        self.time_elapsed_label.pack(side="left", padx=5)

        # self.time_remaining_label = tk.Label(self.time_frame, text="Remaining: 00:00:00.0")
        # self.time_remaining_label.pack(side="left", padx=5)

        # JSON data placeholder
        self.json_data = None
        self.sending_thread = None  # will hold the reference to the sending thread
        self.abort_event = threading.Event()  # flag to signal abort

        # Log output with scroll
        self.log_output = scrolledtext.ScrolledText(self.root, height=10, width=70, wrap=tk.WORD)
        self.log_output.pack(pady=5)

        logging.basicConfig(
            filename=f"batch_send2db_{datetime.now().strftime('%Y_%m_%d_%H_%M_%S')}.log",
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
            encoding="utf-8",
            force=True,
        )

    def select_file(self):
        """Open file dialog to select a JSON or JSONL file and load it accordingly."""
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json;*.jsonl")])

        if file_path:
            self.file_label.config(text=f"Selected file: {file_path}")
            self.load_json(file_path)  # Used for both JSON and JSONL files

    def load_json(self, file_path):
        """Load and parse JSON or JSONL data, handling errors gracefully."""
        try:
            with (jsonlines.open(file_path, "r") if file_path.endswith(".jsonl")
            else open(file_path, "r", encoding="utf-8")) as file:
                data = list(file) if file_path.endswith(".jsonl") else json.load(file)

            if isinstance(data, list):
                self.json_data = data
                self.send_button.config(state=tk.NORMAL)
                self.send_jsonl_button.config(state=tk.NORMAL)
                self.log(f"File contains #records: {len(data)}")
            else:
                raise ValueError("Invalid JSON format: Expected a list of records.")

        except (UnicodeDecodeError, json.JSONDecodeError, jsonlines.InvalidLineError, ValueError) as e:
            messagebox.showerror("Error", f"Invalid JSON: {e}")
            self.json_data = []  # Reset buffer
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load JSON: {e}")

    def log(self, message, log_to_error_file=False):
        """Log messages to the UI text box and optionally to the error log."""
        # Log to the UI text box
        self.log_output.insert(tk.END, message + "\n")
        self.log_output.yview(tk.END)

        # Optionally log to the error log
        if log_to_error_file:
            with open("error_log.txt", "a", encoding="utf-8") as log_file:
                log_file.write(f"{datetime.now()}: {message}\n")

    def authenticate(self):
        """Authenticate with the server using the provided credentials."""
        auth_url = "https://127.0.0.1:5444/token"
        try:
            # Attempt to get the token using HTTP Basic Authentication
            response = requests.get(auth_url, auth=HTTPBasicAuth(USERNAME, PASSWORD))
            if response.status_code == 200:
                # Assume the server returns a JSON response with the token
                token = response.json().get('token')
                if token:
                    self.log(f"Authentication successful. Token: {token}")
                    return token
                else:
                    self.log("Authentication failed: No token returned.")
            else:
                self.log(f"Authentication failed. Status code: {response.status_code}")
        except Exception as e:
            self.log(f"Error during authentication: {e}")
        return None


    def send_in_background(self, target):
        self.time_start_label.config(text=f"Process Start Time: {self.start_time.strftime('%H:%M:%S')}")
        self.time_elapsed_label.config(text="Elapsed: 00:00:00")
        self.update_elapsed_time()
        self.log(f"Processing started at {self.start_time}")

        # Enable Abort button once sending starts
        self.abort_button.config(state=tk.NORMAL)
        # Clear any previous abort signal
        self.abort_event.clear()

        try:
            if target == "db":
                # Initialize pool if not already done
                if not self.db_pool:
                    self.db_pool = dblib.get_db_pool(self.dbconfig)
                # Pass abort_event to dblib.send_data
                if dblib.send_data(self.json_data, self.root, self.progress, self.db_pool,
                                   abort_event=self.abort_event):
                    self.log("Data successfully sent to the database.")
                else:
                    self.log("Data transfer to database failed.")
            elif target == "jsonl":
                self._send_to_jsonl_server()
        except Exception as e:
            self.log(f"Error sending data: {e}")
        finally:
            self.log("Sending process finished.")
            # Disable Abort button when done
            self.abort_button.config(state=tk.DISABLED)

    def send_data(self):
        """Sends data to the database in the background."""
        if self.json_data:
            self.log("Sending data to database...")
            self.start_time = datetime.now()
            self.sending_thread = threading.Thread(target=self.send_in_background, args=("db",), daemon=True)
            self.sending_thread.start()

    def send_to_jsonl_server(self):
        """Sends JSON data to the JSONL server in the background."""
        if self.json_data:
            self.log("Sending data to JSONL server...")
            self.start_time = datetime.now()
            self.sending_thread = threading.Thread(target=self.send_in_background, args=("jsonl",), daemon=True)
            self.sending_thread.start()

    def abort_sending(self):
        self.log("Abort requested. Waiting for current batch to finish...")
        self.abort_event.set()  # Signal abort

    def _send_to_jsonl_server(self):
        jsonl_srv_url = "https://127.0.0.1:5444/data"  # Use HTTPS
        if not self.json_data:
            self.log("No JSON data loaded.")
            return
        try:
            self.log("Initializing connection with JSONL server...")

            # Basic Authentication headers
            auth = HTTPBasicAuth(USERNAME, PASSWORD)

            headers = {"Content-Type": "application/json"}
            timeout = 2  # seconds
            batch_size = 10000
            with requests.Session() as session:
                # Send the start message with authentication
                start_msg = {"control": "start", "total_records": len(self.json_data)}
                response = session.post(jsonl_srv_url, json=start_msg, headers=headers, auth=auth, timeout=timeout,
                                        verify=False)  # 'verify=True' to use default CA bundle
                self.log(f"Start message response: {response.status_code} - {response.text}")
                if response.status_code != 200:
                    self.log("Start message failed. Aborting data transfer.")
                    return
                # Send the data in batches with authentication
                for i in range(0, len(self.json_data), batch_size):
                    if self.abort_event.is_set():
                        self.log("Abort flag set. Halting further JSONL batches.")
                        break
                    batch = self.json_data[i:i + batch_size]
                    response = session.post(jsonl_srv_url, json=batch, headers=headers, auth=auth, timeout=timeout,
                                            verify=False) # /path/to/certificate.crt' for testing self-signed.  # 'verify=True' for SSL verification
                    self.log(f"Sent batch {i // batch_size + 1}, status: {response.status_code}")
                    if response.status_code != 200:
                        self.log(f"Failed to send batch {i // batch_size + 1}, aborting further transfers.")
                        return
                # Send the end message with authentication
                end_msg = {"control": "end"}
                response = session.post(jsonl_srv_url, json=end_msg, headers=headers, auth=auth, timeout=timeout,
                                        verify=False)  # 'verify=True'
                self.log(f"End message response: {response.status_code} - {response.text}")
                self.log("Data transfer complete.")
        except Exception as e:
            self.log(f"Error sending to JSONL server: {e}", log_to_error_file=True)

    def update_elapsed_time(self):
        """Update the elapsed time label while the sending thread is still alive."""
        if self.sending_thread and self.sending_thread.is_alive():
            elapsed = datetime.now() - self.start_time
            self.time_elapsed_label.config(text=f"Elapsed: {str(elapsed).split('.')[0]}")
            self.root.after(100, self.update_elapsed_time)  # Update every 100ms


# Create and run the application
if __name__ == "__main__":
    app = JSONToDBApp()  # Only create the app instance
    app.root.mainloop()  # Start the Tkinter mainloop using the root window in the app instance

import logging
import os
import threading
import tkinter as tk
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, ttk

import dblib
import jsonlines


class App:
    def __init__(self, root_window):
        self.root_window = root_window
        self.root_window.title("XML Jots Data Processing App.IG")
        self.folder_path = tk.StringVar(value="C:/Prod/test jots")
        self.abort_event = threading.Event()
        self.start_time = None
        self.total_files = 0
        self.processed_files = 0
        self.processing_done = False  # flag to stop periodic updates when done
        self.progress_update_interval = 1000  # in milliseconds
        self.origin_output_files = {}  # Keep track of output files for each origin
        self.setup_ui()

    def setup_ui(self):
        """Set up the UI elements for the application."""
        tk.Label(self.root_window, text="Folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(self.root_window, textvariable=self.folder_path, width=50).grid(row=0, column=1)
        tk.Button(self.root_window, text="Browse", command=self.browse_folder).grid(row=0, column=2)
        self.process_button = tk.Button(self.root_window, text="Process",
                                        command=self.start_processing, state=tk.DISABLED)
        self.abort_button = tk.Button(self.root_window, text="Abort",
                                      command=self.abort_processing, state=tk.DISABLED)
        self.process_button.grid(row=3, column=0, columnspan=3)
        self.abort_button.grid(row=4, column=0, columnspan=3)

        # Log output frame with scrollbar
        log_frame = tk.Frame(self.root_window)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        log_scrollbar = tk.Scrollbar(log_frame, orient="vertical")
        log_scrollbar.pack(side="right", fill="y")
        self.log_output = tk.Text(log_frame, height=10, width=70, state=tk.DISABLED, wrap="word",
                                  yscrollcommand=log_scrollbar.set)
        self.log_output.pack(side="left", fill="both", expand=True)
        log_scrollbar.config(command=self.log_output.yview)

        # Progress bar and time tracking labels
        self.progress = ttk.Progressbar(self.root_window, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=6, column=0, columnspan=2)
        self.time_start_label = tk.Label(self.root_window, text="Process Start Time: 00:00:00")
        self.time_elapsed_label = tk.Label(self.root_window, text="Elapsed Time: 00:00:00.0")
        self.time_remaining_label = tk.Label(self.root_window, text="Remaining Time: 00:00:00.0")
        self.files_processed_label = tk.Label(self.root_window, text="Files Processed: 0")
        self.time_start_label.grid(row=7, column=0)
        self.time_elapsed_label.grid(row=7, column=1)
        self.time_remaining_label.grid(row=7, column=2)
        self.files_processed_label.grid(row=6, column=2)

    def update_progress(self):
        """Periodic update for the progress bar and labels."""
        if self.total_files > 0:
            percent = (self.processed_files / self.total_files) * 100
            self.progress["value"] = percent
            self.files_processed_label.config(text=f"Files Processed: {self.processed_files}")
            elapsed = (datetime.now() - self.start_time).total_seconds()
            if self.processed_files > 0:
                remaining = max(0, ((elapsed / self.processed_files) * self.total_files) - elapsed)
            else:
                remaining = 0
            format_time = lambda s: f"{int(s//3600):02}:{int((s % 3600)//60):02}:{int(s % 60):02}.{int((s % 1)*10):1}"
            self.time_elapsed_label.config(text=f"Elapsed Time: {format_time(elapsed)}")
            self.time_remaining_label.config(text=f"Remaining Time: {format_time(remaining)}")
        # Schedule next update if processing is not done.
        if not self.processing_done:
            self.root_window.after(self.progress_update_interval, self.update_progress)

    def log(self, message, error=False, to_file=True):
        self.root_window.after(0, self._log_ui_update, message)
        if to_file:
            (logging.error if error else logging.info)(message)

    def _log_ui_update(self, message):
        self.log_output.config(state=tk.NORMAL)
        self.log_output.insert(tk.END, message + "\n")
        self.log_output.yview(tk.END)
        self.log_output.config(state=tk.DISABLED)

    def browse_folder(self):
        self.process_button.config(state=tk.DISABLED)
        if folder := filedialog.askdirectory():
            self.folder_path.set(folder)
            self.process_button.config(state=tk.NORMAL)

    def abort_processing(self):
        self.abort_event.set()
        self.abort_button.config(state=tk.DISABLED)
        self.process_button.config(state=tk.NORMAL)
        self.log("Processing aborted!", error=True)

    def start_processing(self):
        # Reset progress bar and time labels at the start
        self.progress["value"] = 0
        self.time_elapsed_label.config(text="Elapsed Time: 00:00:00.0")
        self.time_remaining_label.config(text="Remaining Time: 00:00:00.0")
        self.files_processed_label.config(text="Files Processed: 0")

        # Initialize tracking variables
        self.processed_files = 0  # Reset the processed files count
        self.total_files = 0  # Reset the total files count
        self.start_time = datetime.now()  # Set the start time for the new process
        self.processing_done = False  # Ensure the process is not marked as done
        self.abort_event.clear()  # Reset the abort event

        # Update the start time label
        self.time_start_label.config(text=f"Process Start Time: {self.start_time.strftime('%H:%M:%S')}")

        if folder := self.folder_path.get():
            self._extracted_from_start_processing_20(folder)

    # TODO Rename this here and in `start_processing`
    def _extracted_from_start_processing_20(self, folder):
        self.abort_button.config(state=tk.NORMAL)
        self.process_button.config(state=tk.DISABLED)
        self.log(f"Processing started at {self.start_time}")

        logging.basicConfig(
            filename=f"xml2json_{self.start_time.strftime('%Y_%m_%d_%H_%M_%S')}.log",
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
            encoding="utf-8",
            force=True,
        )

        # Start periodic progress updates
        self.root_window.after(self.progress_update_interval, self.update_progress)

        # Start processing in a separate thread
        threading.Thread(target=self.process_folder, args=(folder,), daemon=True).start()

    def process_folder(self, folder):
        try:
            self.log("Reading the folder... Please wait.")
            files = list(Path(folder).glob("*.xml"))
            self.total_files = len(files)
            if self.total_files == 0:
                self.log("No XML files found.", error=True)
                self.processing_done = True
                return

            self.log(f"Processing {self.total_files} files in {folder}.")
            # Group records by origin; each group is a list of JSON records.
            grouped = defaultdict(list)

            # Update output files with a new timestamp at the start of each process
            timestamp = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
            self.origin_output_files = {origin: Path(f"{origin}_processed_data_{timestamp}.jsonl")
                                        for origin in set()}

            with ThreadPoolExecutor(max_workers=os.cpu_count() * 2) as executor:
                future_to_file = {executor.submit(self.safe_process_file, file): file for file in files}
                for future in as_completed(future_to_file):
                    if self.abort_event.is_set():
                        break
                    result = future.result()
                    self.processed_files += 1

                    if result is not None:
                        origin = result.get("origin", "unknown")
                        grouped[origin].append(result)
                        # Dump batch for this origin if batch reaches 10,000 records.
                        if len(grouped[origin]) >= 10000:
                            self.append_to_output_file(origin, grouped[origin])
                            grouped[origin].clear()

            # Dump remaining records for all origins
            for origin, records in grouped.items():
                if records:
                    self.append_to_output_file(origin, records)

            self.log(f"Processing completed. Total time elapsed: {datetime.now() - self.start_time}")

        except Exception as e:
            self.log(f"Error processing folder {folder}: {e}", error=True)
        finally:
            self.processing_done = True
            self.process_button.config(state=tk.NORMAL)
            self.abort_button.config(state=tk.DISABLED)
            self.abort_event.clear()

    def safe_process_file(self, file):
        if self.abort_event.is_set():
            return None
        try:
            return dblib.parse_xml_to_json(file)
        except Exception as e:
            self.log(f"Error parsing {file.name}: {e}", error=True)
            return None

    def append_to_output_file(self, origin, records):
        """Append the processed records to the output file for the origin."""
        if origin not in self.origin_output_files:
            # If no file for this origin yet, create it with the timestamped name.
            timestamp = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
            output_file = Path(f"{origin}_processed_data_{timestamp}.jsonl")
            self.origin_output_files[origin] = output_file
        else:
            # Use the existing output file for this origin
            output_file = self.origin_output_files[origin]

        try:
            with jsonlines.open(output_file, mode="a") as writer:
                writer.write_all(records)
            self.log(f"Wrote {len(records)} records to {output_file}")
        except Exception as e:
            self.log(f"Error writing to {output_file}: {e}", error=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    tk.Label(root, text="Server Log:").grid(row=8, column=0, columnspan=3, sticky="w")
    root.mainloop()

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfMerger

class PDFMergerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Merger")
        self.geometry("600x400")
        self.pdf_files = []  # List to store full PDF file paths

        self.create_widgets()
        self.setup_dnd()

    def create_widgets(self):
        # Main frame for listbox and buttons
        main_frame = tk.Frame(self)
        main_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Listbox to display PDF files
        self.listbox = tk.Listbox(main_frame, selectmode=tk.SINGLE)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Frame for Up, Down, and other buttons
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)

        tk.Button(btn_frame, text="Add Files", command=self.add_files).pack(pady=5)
        tk.Button(btn_frame, text="Up", command=self.move_up).pack(pady=5)
        tk.Button(btn_frame, text="Down", command=self.move_down).pack(pady=5)
        tk.Button(btn_frame, text="Remove", command=self.remove_file).pack(pady=5)
        tk.Button(btn_frame, text="Merge", command=self.merge_files).pack(pady=20)

        # Log field for debugging messages
        self.log_field = tk.Text(self, height=5, state=tk.DISABLED)
        self.log_field.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    def setup_dnd(self):
        # Register the listbox as a drop target and bind drop event
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.on_drop)

    def log_message(self, message):
        self.log_field.config(state=tk.NORMAL)
        self.log_field.insert(tk.END, message + "\n")
        self.log_field.config(state=tk.DISABLED)
        self.log_field.yview(tk.END)

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.listbox.insert(tk.END, os.path.basename(file))
        self.log_message("Added files.")

    def on_drop(self, event):
        # event.data contains file paths separated by spaces (possibly quoted)
        files = self.tk.splitlist(event.data)
        for file in files:
            if file.lower().endswith('.pdf') and file not in self.pdf_files:
                self.pdf_files.append(file)
                self.listbox.insert(tk.END, os.path.basename(file))
        self.log_message("Dropped files added.")

    def move_up(self):
        selected = self.listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        if idx == 0:
            return
        # Swap the selected file with the one above it
        self.pdf_files[idx], self.pdf_files[idx - 1] = self.pdf_files[idx - 1], self.pdf_files[idx]
        self.refresh_listbox()
        self.listbox.selection_set(idx - 1)
        self.log_message("Moved file up.")

    def move_down(self):
        selected = self.listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        if idx == len(self.pdf_files) - 1:
            return
        # Swap the selected file with the one below it
        self.pdf_files[idx], self.pdf_files[idx + 1] = self.pdf_files[idx + 1], self.pdf_files[idx]
        self.refresh_listbox()
        self.listbox.selection_set(idx + 1)
        self.log_message("Moved file down.")

    def remove_file(self):
        selected = self.listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        self.pdf_files.pop(idx)
        self.refresh_listbox()
        self.log_message("Removed file.")

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for file in self.pdf_files:
            self.listbox.insert(tk.END, os.path.basename(file))

    def merge_files(self):
        if len(self.pdf_files) < 2:
            messagebox.showwarning("Insufficient Files", "Please add at least two PDF files to merge.")
            return

        output_file = os.path.normpath(filedialog.asksaveasfilename(defaultextension=".pdf",
                                                   filetypes=[("PDF Files", "*.pdf")]))
        if not output_file:
            return

        merger = PdfMerger()
        try:
            for pdf in self.pdf_files:
                merger.append(pdf)
            merger.write(output_file)
            merger.close()
            messagebox.showinfo("Success", f"PDFs merged successfully into {output_file}")
            self.log_message(f"Merged files into: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log_message(f"Error during merge: {e}")

if __name__ == "__main__":
    app = PDFMergerApp()
    app.mainloop()

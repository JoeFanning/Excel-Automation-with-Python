import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path


class ExcelAutomationGUI:
    def __init__(self, pipeline_callback):
        """
        pipeline_callback: The function in main.py that runs your backend data engine.
        """
        self.pipeline_callback = pipeline_callback

        # 1. Initialize permanent application window
        self.root = tk.Tk()
        self.root.title("Weekly Sales Automation Engine")
        self.root.geometry("500x320")
        self.root.config(bg="#f8f9fa")

        # Center window on screen
        self.root.eval('tk::PlaceWindow . center')

        self._build_ui()

    def _build_ui(self):
        # Header Title
        title = tk.Label(
            self.root,
            text="Weekly Sales Data Pipeline",
            font=("Segoe UI", 16, "bold"),
            bg="#f8f9fa", fg="#212529"
        )
        title.pack(pady=(30, 10))

        # Instructions
        intro_text = (
            "Welcome! This desktop tool automates your weekly reporting.\n\n"
            "1. Click the button below to select raw sales spreadsheets.\n"
            "2. The system will merge, clean, sort, and calculate metrics.\n"
            "3. A visual dashboard will generate inside the 'output' folder.\n"
            "4. A unified summary email will automatically go to your client."
        )
        desc = tk.Label(
            self.root, text=intro_text, font=("Segoe UI", 9),
            bg="#f8f9fa", fg="#495057", justify="left"
        )
        desc.pack(pady=10, padx=25)

        # Main Action Button
        self.run_btn = tk.Button(
            self.root,
            text="Select Files & Launch Pipeline",
            font=("Segoe UI", 11, "bold"),
            bg="#198754", fg="white",  # Green theme for success/execution
            activebackground="#157347", activeforeground="white",
            padx=20, pady=8, cursor="hand2",
            command=self.trigger_file_picker
        )
        self.run_btn.pack(pady=20)

    def trigger_file_picker(self):
        # Default explorer look path directly to project input directory
        input_dir = Path("input").resolve()
        input_dir.mkdir(exist_ok=True)
        Path("output").mkdir(exist_ok=True)

        file_paths = filedialog.askopenfilenames(
            title="Select Weekly Sales Excel Files",
            initialdir=input_dir,
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        files = list(file_paths)

        if not files:
            messagebox.showwarning("Staged", "No files selected. Pipeline aborted.")
            return

        # Disable button during calculation to prevent double clicks
        self.run_btn.config(text="Processing Data... Please Wait", state="disabled", bg="#6c757d")
        self.root.update_idletasks()

        # Hand off control back to your main.py engine
        success, message = self.pipeline_callback(files)

        # Restore button state
        self.run_btn.config(text="Select Files & Launch Pipeline", state="normal", bg="#198754")

        # Show native alert window based on the pipeline outcome
        if success:
            messagebox.showinfo("Success", message)
            self.root.destroy()  # Close application when entirely completed
        else:
            messagebox.showerror("Pipeline Failure", message)

    def run(self):
        self.root.mainloop()

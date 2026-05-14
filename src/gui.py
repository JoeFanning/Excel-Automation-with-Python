import os
import tkinter as tk
from tkinter import filedialog, messagebox
import logging


class ExcelAutomationGUI:
    def __init__(self, pipeline_callback):
        """
        Initializes the clean Desktop User Interface window framework.
        """
        self.pipeline_callback = pipeline_callback

        # Connect safely back to our single initialized logging stream
        self.logger = logging.getLogger("ExcelAutomation")

        self.root = tk.Tk()
        self.root.title("Excel Automation Pipeline Engine")
        self.root.geometry("500x350")
        self.root.resizable(False, False)

        # UI Elements Layout Layout Configuration
        self.label = tk.Label(
            self.root,
            text="Excel Sales Dashboard Automation Engine",
            font=("Arial", 14, "bold")
        )
        self.label.pack(pady=20)

        self.status_label = tk.Label(
            self.root,
            text="Status: Ready to process data",
            font=("Arial", 10),
            fg="blue"
        )
        self.status_label.pack(pady=10)

        self.process_button = tk.Button(
            self.root,
            text="Select Excel Files & Run Pipeline",
            command=self.trigger_file_picker,
            font=("Arial", 11, "bold"),
            bg="#2da44e",
            fg="white",
            padx=10,
            pady=5
        )
        self.process_button.pack(pady=30)

    def trigger_file_picker(self):
        """Opens a file dialog window allowing users to multi-select target spreadsheets."""
        self.status_label.config(text="Status: Selecting source spreadsheets...", fg="orange")
        self.root.update()

        files = filedialog.askopenfilenames(
            title="Select Monthly Sales Sheets",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not files:
            self.status_label.config(text="Status: Ready to process data", fg="blue")
            return

        self.status_label.config(text="Status: Executing data calculations...", fg="orange")
        self.root.update()

        # Trigger the core processing pipeline engine hook
        success, message = self.pipeline_callback(files)

        if success:
            self.status_label.config(text="Status: Success! Email report sent.", fg="green")
            messagebox.showinfo("Pipeline Completed", message)
        else:
            self.status_label.config(text="Status: Critical Error encountered.", fg="red")
            messagebox.showerror("Pipeline Failed", message)

    def synchronize_cloud_vault(self):
        """
        Communicates with GitHub API tracking logs securely.
        """
        # FIX: Wiped out the raw hardcoded plain-text 'ghp_...' token!
        # Fetches securely out of your system memory workspace environments instead
        github_token = os.environ.get("GITHUB_TOKEN")

        if not github_token:
            self.logger.warning("Local GUI: GITHUB_TOKEN environment variable is not defined.")
            return False

        self.logger.info("Contacting GitHub API to check for existing remote file SHA...")

        url = "github.com"
        headers = {
            "Authorization": f"Bearer {github_token}",
            "User-Agent": "Excel-Automation-Pipeline-App",
            "Accept": "application/vnd.github+json"
        }

        # Communicates safely over HTTPS using explicit network timeout gates
        try:
            import requests
            response = requests.get(url, headers=headers, timeout=15)
            if response.status_code == 200:
                self.logger.info("Successfully fetched remote tracking metrics metadata.")
                return response.json().get("sha")
            else:
                self.logger.error(f"GitHub API handshake failed with code: {response.status_code}")
                return None
        except Exception as e:
            self.logger.error(f"Cloud synchronization failure: {e}")
            return None

    def run(self):
        """Starts up the permanent graphical application frame environment loop."""
        self.root.mainloop()

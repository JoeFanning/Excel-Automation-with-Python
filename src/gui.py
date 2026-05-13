import tkinter as tk
from tkinter import filedialog

# Open pop window to select files
def get_files_window():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring to front

    file_paths = filedialog.askopenfilenames(
        title="Select Excel Files"
    )

    root.destroy()
    #return a tuple of file names
    return list(file_paths)

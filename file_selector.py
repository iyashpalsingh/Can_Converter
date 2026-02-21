#file_selector.py
from tkinter import Tk, filedialog


def select_file():
    root = Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="Select BUSMASTER / candump log file",
        filetypes=[
            ("Log Files", "*.asc *.log *.txt"),
            ("All Files", "*.*")
        ]
    )

    return list(files)

def detect_file_type(input_file):
    with open(input_file, "r", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            if line.startswith("***BUSMASTER"):
                return "BUSMASTER"

            if line.startswith("(") and "can" in line and "#" in line:
                return "CANDUMP"

            if "Frame Id" in line and "Data(Hex)" in line:
                return "TABLETXT"
            
            break

    return "UNKNOWN"
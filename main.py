# main.py
import os
import sys

# 🔴 REQUIRED for PyInstaller
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, base_path)

from file_selector import select_file, detect_file_type
from busmaster_parser import parse_busmaster
from candump_parser import parse_can_dump
from tabletxt_parser import parse_tabletxt


def main():
    files = select_file()

    if not files:
        print("No file selected. Exiting.")
        return

    for input_file in files:

        print(f"\nProcessing: {input_file}")

        ft = detect_file_type(input_file)

        if ft == "BUSMASTER":
            parse_busmaster(input_file)

        elif ft == "CANDUMP":
            parse_can_dump(input_file)

        elif ft == "TABLETXT":
            parse_tabletxt(input_file)

        else:
            print(f"Unknown file format: {input_file}")


if __name__ == "__main__":
    main()
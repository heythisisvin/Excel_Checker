import sys
import tkinter as tk
from gui import ExcelScannerGUI

# Import CLI cleanup modules
from cleanup import remove_excel_objects
from cleanup_styles import cleanup_excel_file

import argparse


def run_cli():
    parser = argparse.ArgumentParser(description="Excel Scanner CLI")

    parser.add_argument("file", help="Path to Excel file")
    parser.add_argument("--check", action="store_true", help="Check for corruption")
    parser.add_argument("--analyze", action="store_true", help="Analyze performance")
    parser.add_argument("--cleanup", action="store_true", help="Full cleanup")
    parser.add_argument("--cleanup_styles", action="store_true", help="Cleanup excessive styles only")
    parser.add_argument("-o", "--output", help="Output file path")

    args = parser.parse_args()

    # FULL CLEANUP
    if args.cleanup:
        output = cleanup_excel_file(args.file, args.output)
        print(f"Full cleanup complete → {output}")
        return

    # STYLES-ONLY CLEANUP
    if args.cleanup_styles:
        output = cleanup_styles_file(args.file, args.output)
        print(f"Style cleanup complete → {output}")
        return

    print("No valid action selected. Use --help for options.")


def run_gui():
    root = tk.Tk()
    app = ExcelScannerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    # If no command-line arguments → open GUI
    if len(sys.argv) == 1:
        run_gui()
    else:
        run_cli()

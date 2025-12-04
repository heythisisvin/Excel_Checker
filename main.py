import sys
import tkinter as tk
from gui import ExcelScannerGUI

# Cleanup modules
from cleanup import remove_excel_objects, remove_excessive_styles
from cleanup_styles import cleanup_excel_file
from cleanup_styles import cleanup_styles_file

import argparse


def run_cli():
    parser = argparse.ArgumentParser(description="Excel Scanner CLI Tool")

    parser.add_argument("file", help="Path to Excel file")
    parser.add_argument("--cleanup", action="store_true", help="Full cleanup (objects, links, styles)")
    parser.add_argument("--cleanup_styles", action="store_true", help="Cleanup styles only")
    parser.add_argument("-o", "--output", help="Output file path")

    args = parser.parse_args()

    if args.cleanup:
        output = cleanup_excel_file(args.file, args.output)
        print(f"Full cleanup complete → {output}")
        return

    if args.cleanup_styles:
        output = cleanup_styles_file(args.file, args.output)
        print(f"Style-only cleanup complete → {output}")
        return

    print("No valid action selected. Use --help for more options.")


def run_gui():
    root = tk.Tk()
    app = ExcelScannerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    # If run without arguments → launch GUI
    if len(sys.argv) == 1:
        run_gui()
    else:
        run_cli()

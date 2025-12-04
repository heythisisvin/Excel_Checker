import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from analyzer import analyze_xlsx
from basic_corruption_checker import check_excel_corruption
from report_generator import generate_report

# CLEANUP MODULES
from cleanup import remove_excel_objects, remove_excessive_styles  # full cleanup
from cleanup_styles import cleanup_styles_file                 # styles-only cleanup
from cleanup_styles import cleanup_excel_file

class ExcelScannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Scanner Tool")

        self.file_path = None
        self.analysis_result = None

        # -----------------------------------
        # Select Excel file
        # -----------------------------------
        self.btn_select = tk.Button(root, text="Select Excel File", command=self.choose_file)
        self.btn_select.pack(pady=10)

        # -----------------------------------
        # Output box
        # -----------------------------------
        self.output_box = scrolledtext.ScrolledText(root, width=80, height=20)
        self.output_box.pack(pady=10)

        # -----------------------------------
        # Buttons
        # -----------------------------------
        self.btn_save = tk.Button(root, text="Save Report", command=self.save_report)
        self.btn_save.pack(pady=5)

        self.cleanup_button = tk.Button(root, text="Full Cleanup", command=self.run_cleanup)
        self.cleanup_button.pack(pady=5)

        self.cleanup_styles_button = tk.Button(root, text="Cleanup Styles Only", command=self.run_cleanup_styles)
        self.cleanup_styles_button.pack(pady=5)

    # ------------------------------------------------------------------
    # CLEANUP FUNCTIONS
    # ------------------------------------------------------------------
    def run_cleanup(self):
        if not self.file_path:
            messagebox.showerror("Error", "No file selected!")
            return

        try:
            output = cleanup_excel_file(self.file_path)
            messagebox.showinfo("Success", f"Full cleanup completed.\nSaved to:\n{output}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run_cleanup_styles(self):
        if not self.file_path:
            messagebox.showerror("Error", "No file selected!")
            return

        try:
            output = cleanup_styles_file(self.file_path)
            messagebox.showinfo("Success", f"Style cleanup completed.\nSaved to:\n{output}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ------------------------------------------------------------------
    # FILE SELECTION
    # ------------------------------------------------------------------
    def choose_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")]
        )

        if not self.file_path:
            return

        self.output_box.delete(1.0, tk.END)
        self.output_box.insert(tk.END, f"Selected file:\n{self.file_path}\n\n")

        try:
            corruption_result = check_excel_corruption(self.file_path)
            analysis_result = analyze_xlsx(self.file_path)

            # Combine both results
            self.analysis_result = f"{corruption_result}\n\n{analysis_result}"

            self.output_box.insert(tk.END, self.analysis_result)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to analyze file:\n{e}")

    # ------------------------------------------------------------------
    # SAVE REPORT
    # ------------------------------------------------------------------
    def save_report(self):
        if not self.analysis_result:
            messagebox.showwarning("No Report", "Please analyze a file first.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Report", "*.txt")]
        )

        if not save_path:
            return

        try:
            generate_report(self.analysis_result, save_path)
            messagebox.showinfo("Saved", f"Report saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save report:\n{e}")

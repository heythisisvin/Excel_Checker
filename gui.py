import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from analyzer import analyze_xlsx
from basic_corruption_checker import check_excel_corruption
from report_generator import generate_report


class ExcelScannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Scanner Tool")

        self.file_path = None
        self.analysis_result = None

        # ------------------------------
        # File selection button
        # ------------------------------
        self.btn_select = tk.Button(root, text="Select Excel File", command=self.choose_file)
        self.btn_select.pack(pady=10)

        # ------------------------------
        # Analysis result box
        # ------------------------------
        self.output_box = scrolledtext.ScrolledText(root, width=80, height=20)
        self.output_box.pack(pady=10)

        # ------------------------------
        # Save report button (NEW)
        # ------------------------------
        self.btn_save = tk.Button(root, text="Save Report", command=self.save_report)
        self.btn_save.pack(pady=10)

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

            self.analysis_result = f"{corruption_result}\n\n{analysis_result}"

            self.output_box.insert(tk.END, self.analysis_result)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to analyze file:\n{e}")

    # -------------------------------------------------------
    # New Button Function: Save report to a chosen directory
    # -------------------------------------------------------
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

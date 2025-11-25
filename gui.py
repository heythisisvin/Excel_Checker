import tkinter as tk
from tkinter import filedialog, messagebox
from analyzer import analyze_xlsx
from basic_corruption_checker import check_xlsx_structure

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Scanner")
        self.geometry("400x200")

        btn = tk.Button(self, text="Choose Excel File", command=self.choose_file)
        btn.pack(pady=40)

    def choose_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")]
        )

        if not file_path:
            return

        try:
            result_corruption = check_xlsx_structure(file_path)
            result_analysis = analyze_xlsx(file_path)

            messagebox.showinfo(
                "Scan Complete",
                f"Corruption Check:\n{result_corruption}\n\n"
                f"Analysis:\n{result_analysis}"
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = App()
    app.mainloop()

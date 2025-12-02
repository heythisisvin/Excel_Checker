import win32com.client as win32
import os

def remove_excel_objects(input_file, output_file=None):
    """
    Removes shapes, charts, OLE objects, and form controls
    from all worksheets in an Excel workbook.
    Requires Windows + Excel installed.
    """

    excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False

    try:
        workbook = excel.Workbooks.Open(os.path.abspath(input_file))

        for sheet in workbook.Worksheets:
            print(f"Cleaning sheet: {sheet.Name}")

            # Remove Shapes (pictures, buttons, charts, smartart, icons)
            try:
                count_shapes = sheet.Shapes.Count
                for i in range(count_shapes, 0, -1):
                    sheet.Shapes(i).Delete()
                print(f"  Removed {count_shapes} shapes")
            except:
                print("  No shapes found or deletion error")

            # Remove OLEObjects (embedded files, activex)
            try:
                count_ole = sheet.OLEObjects().Count
                for i in range(count_ole, 0, -1):
                    sheet.OLEObjects().Item(i).Delete()
                print(f"  Removed {count_ole} OLE objects")
            except:
                print("  No OLE objects found")

            # Remove ChartObjects
            try:
                count_charts = sheet.ChartObjects().Count
                for i in range(count_charts, 0, -1):
                    sheet.ChartObjects().Item(i).Delete()
                print(f"  Removed {count_charts} charts")
            except:
                print("  No ChartObjects found")

        # Save cleaned file
        if not output_file:
            output_file = input_file.replace(".xlsx", "_cleaned.xlsx")

        workbook.SaveAs(os.path.abspath(output_file))
        workbook.Close()

        print(f"\nCleanup completed. Saved as: {output_file}")

        return output_file

    except Exception as e:
        print("Error:", e)

    finally:
        excel.Quit()




if __name__ == "__main__":
    # Example usage
    input_path = r"C:\Users\local-u\PycharmProjects\Excel_Checker_Ver2\99.xlsx"
    remove_excel_objects(input_path)

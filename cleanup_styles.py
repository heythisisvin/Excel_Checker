# cleanup.py
# Module to remove external links and excessive styles from an Excel workbook

import openpyxl
from openpyxl import load_workbook

def remove_external_links(wb):
    """
    Remove all external links from the workbook.
    """
    if hasattr(wb, "external_links"):
        wb.external_links = []
    if hasattr(wb, "_external_links"):
        wb._external_links = []

    # Some external references are stored inside defined names
    if hasattr(wb, "defined_names"):
        new_names = []
        for defined_name in wb.defined_names.definedName:
            if "!" in defined_name.attr_text and "[" in defined_name.attr_text:
                # This looks like an external reference
                continue
            new_names.append(defined_name)

        wb.defined_names.definedName = new_names


def remove_excessive_styles(wb):
    """
    Remove unused cell styles.
    WARNING: openpyxl does NOT support removing styles safely at runtime.
             Instead, the best approach is to reset style to 'Normal'.
    """
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                cell.style = "Normal"  # Reset formatting to default


def remove_drawing_objects(wb):
    """
    Remove drawings, images, charts, shapes, smart art.
    """

    for sheet in wb.worksheets:
        # Remove sheet drawings (charts/images)
        if hasattr(sheet, "_images"):
            sheet._images = []

        if hasattr(sheet, "_charts"):
            sheet._charts = []

        if hasattr(sheet, "drawing"):
            sheet.drawing = None

        if hasattr(sheet, "_rels"):
            sheet._rels = {}

        # Remove comments
        if hasattr(sheet, "comments"):
            sheet.comments = []


def remove_pivot_caches(wb):
    """
    Remove pivot cache definitions (helps reduce file size).
    """
    if hasattr(wb, "_pivots"):
        wb._pivots = []

    if hasattr(wb, "_pivot_caches"):
        wb._pivot_caches = []


def cleanup_excel_file(input_file, output_file=None):
    """
    Full cleanup pipeline:
        - Remove external links
        - Remove excessive styles
        - Remove drawing objects
        - Remove pivot caches
    """
    if output_file is None:
        output_file = input_file.replace(".xlsx", "_CLEANED.xlsx")

    wb = load_workbook(input_file)

    remove_external_links(wb)
    remove_excessive_styles(wb)
    remove_drawing_objects(wb)
    remove_pivot_caches(wb)

    wb.save(output_file)
    return output_file

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Excel Cleanup Tool")
    parser.add_argument("input_file", help="Path to the Excel file to clean")
    parser.add_argument("-o", "--output", help="Output cleaned file path (optional)")

    args = parser.parse_args()

    output = cleanup_excel_file(args.input_file, args.output)
    print(f"✔ Cleanup complete! Saved to: {output}")

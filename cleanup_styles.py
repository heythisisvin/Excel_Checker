from openpyxl import load_workbook


# ---------------------------
# REMOVE EXTERNAL LINKS
# ---------------------------
def remove_external_links(wb):
    """
    Removes all external links, external references, and formulas referring to other files.
    """

    # Remove external links in defined names
    names_to_remove = []
    for name in wb.defined_names:
        if name.attr_text and "!" in name.attr_text and ".xlsx" in name.attr_text.lower():
            names_to_remove.append(name.name)

    for name in names_to_remove:
        del wb.defined_names[name]

    # Remove external links in sheets
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f" and cell.value:
                    if "[" in cell.value and "]" in cell.value:  # external reference pattern
                        cell.value = None  # clear formula
                        cell.data_type = "n"  # set normal type


# ---------------------------
# REMOVE EXCESSIVE STYLES
# ---------------------------
def remove_excessive_styles(wb):
    """
    Removes all styles and resets cells to default to reduce file size.
    """
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.style = "Normal"  # reset style


# ---------------------------
# REMOVE OBJECTS
# ---------------------------
def remove_drawings(wb):
    """
    Removes shapes, charts, images, and drawings.
    """
    for ws in wb.worksheets:
        ws._images = []
        ws._charts = []
        ws._drawings = None


# ---------------------------
# FULL CLEANUP WORKFLOW
# ---------------------------
def cleanup_excel_file(input_file, output_file=None):
    wb = load_workbook(input_file, data_only=False)

    print("Removing external links...")
    remove_external_links(wb)

    print("Removing styles...")
    remove_excessive_styles(wb)

    print("Removing drawings / objects...")
    remove_drawings(wb)

    # Save
    if not output_file:
        output_file = input_file.replace(".xlsx", "_CLEANED.xlsx")

    wb.save(output_file)
    return output_file


# ---------------------------
# STANDALONE EXECUTION
# ---------------------------
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Excel Cleanup Tool")
    parser.add_argument("input_file", help="Path to the Excel file to clean")
    parser.add_argument("-o", "--output", help="Output cleaned file path (optional)")

    args = parser.parse_args()

    output = cleanup_excel_file(args.input_file, args.output)
    print(f"✔ Cleanup complete! Saved to: {output}")

"""basic_corruption_checker.py
Quick check whether an XLSX file is a valid ZIP container and verifies core parts exist.
"""

import zipfile
import sys

CORE_FILES = [
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/sharedStrings.xml',
    'xl/styles.xml',
]


def check_excel_corruption(path: str) -> str:
    """
    Returns a human-readable message only (string),
    compatible with GUI output.
    """
    try:
        with zipfile.ZipFile(path, 'r') as z:
            bad = z.testzip()
            if bad is not None:
                return f"[CORRUPT] Bad file entry detected: {bad}"

            names = set(z.namelist())
            missing = [f for f in CORE_FILES if f not in names]
            if missing:
                return f"[WARNING] Missing core components: {missing}"

            return "[OK] Excel structure looks valid."

    except zipfile.BadZipFile:
        return "[ERROR] Not a valid ZIP file (Excel XLSX required)."
    except Exception as e:
        return f"[ERROR] Exception occurred: {e}"


# CLI execution
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python basic_corruption_checker.py <file.xlsx>")
        sys.exit(1)

    path = sys.argv[1]

    result = check_excel_corruption(path)
    print(result)

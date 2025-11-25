"""basic_corruption_checker.py
Quickly tests whether an XLSX file is a valid ZIP container and verifies common parts exist.
"""
import zipfile
import sys
from typing import Tuple

CORE_FILES = [
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/sharedStrings.xml',
    'xl/styles.xml',
]


def check_excel_corruption(path: str) -> Tuple[bool, str]:
    try:
        with zipfile.ZipFile(path, 'r') as z:
            bad = z.testzip()
            if bad is not None:
                return False, f"Corrupt/Bad file entry: {bad}"
            names = set(z.namelist())
            missing = [f for f in CORE_FILES if f not in names]
            if missing:
                return False, f"Missing core parts: {missing}"
            return True, "Structure OK"
    except zipfile.BadZipFile:
        return False, "Not a valid ZIP / not an XLSX file"
    except Exception as e:
        return False, f"Exception: {e}"


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python basic_corruption_checker.py <file.xlsx>')
        sys.exit(1)
    path = sys.argv[1]
    ok, msg = check_xlsx_structure(path)
    print('OK' if ok else 'CORRUPT/PROBLEM', '-', msg)

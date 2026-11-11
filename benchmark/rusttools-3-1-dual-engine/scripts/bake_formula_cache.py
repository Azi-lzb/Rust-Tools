# -*- coding: utf-8 -*-
"""Bake formula cached values into fixtures via Excel COM.

openpyxl writes formulas without cached results; both print engines read the
cached value only. Opening each fixture in Excel, recalculating and saving
plants the cache — including cached date values for the 1900/1904 fixtures.

Usage: python -I scripts/bake_formula_cache.py <xlsx> [<xlsx> ...]
"""
import sys
from pathlib import Path

import win32com.client


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        for raw in sys.argv[1:]:
            path = Path(raw).resolve()
            book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
            try:
                book.Application.CalculateFullRebuild()
                book.Save()
                print("baked", path.name)
            finally:
                book.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()

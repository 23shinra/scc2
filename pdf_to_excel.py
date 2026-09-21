#!/usr/bin/env python3
"""Переделать PDF-сканы реестра сделок в Excel.

    python3 pdf_to_excel.py
    python3 pdf_to_excel.py pdfs/файл.pdf
    python3 pdf_to_excel.py ~/Downloads/новый.pdf

Без аргументов берёт все PDF из папки pdfs/.
Excel сохраняется в папку excel/, один файл на каждый PDF.
"""

from convert_rfc import main

if __name__ == "__main__":
    main()

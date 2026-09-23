#!/usr/bin/env python3
"""PDF-сканы реестра сделок РФЦ → Excel.

Ежедневная синхронизация (новые с сайта, уже скачанные — skip):

    python3 pdf_to_excel.py
    python3 pdf_to_excel.py --sync

Конвертация локальных PDF:

    python3 pdf_to_excel.py --local
    python3 pdf_to_excel.py --local pdfs/файл.pdf

Принудительно перекачать всё:

    python3 pdf_to_excel.py --sync --force
"""

from __future__ import annotations

import sys


def main() -> None:
    args = sys.argv[1:]
    if "--local" in args:
        from convert_rfc import main as convert_main

        sys.argv = [sys.argv[0], *[a for a in args if a != "--local"]]
        convert_main()
        return

    # default: sync from rfc.kz
    from sync_rfc import sync

    force = "--force" in args
    sync(force=force)


if __name__ == "__main__":
    main()

# scc2 — PDF реестров сделок → Excel

Скрипт переводит сканы итогов централизованных торгов (PDF-картинки) в Excel.

## Запуск (macOS)

```bash
python3 -m pip install --user pymupdf pillow openpyxl numpy
python3 pdf_to_excel.py
```

Один файл:

```bash
python3 pdf_to_excel.py pdfs/файл.pdf
```

Результат: папка `excel/` — по одному `.xlsx` на каждый PDF.

## Структура

| Путь | Назначение |
|------|------------|
| `pdf_to_excel.py` | Точка входа |
| `convert_rfc.py` | Логика OCR и сборки Excel |
| `ocr.swift` | Распознавание текста через Apple Vision |
| `pdfs/` | Исходные PDF |
| `excel/` | Готовые Excel |

Нужен macOS: текст со сканов читает Vision (`ocr.swift` → бинарник `ocr`).

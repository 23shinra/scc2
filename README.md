# scc2 — PDF реестров сделок → Excel

Скрипт забирает итоги централизованных торгов с [rfc.kz](https://rfc.kz/ru/power-market/prices-and-rates/results-of-centralized-trading/),
пропускает уже скачанные и новые переводит в Excel.

## Запуск

```bash
python3 -m pip install --user pymupdf pillow openpyxl numpy
python3 pdf_to_excel.py          # sync: только новые
python3 pdf_to_excel.py --force  # перекачать всё заново
python3 pdf_to_excel.py --local  # только локальные pdfs/
```

PDF → `pdfs/`, Excel → `excel/` (обе папки в `.gitignore`).
Учёт скачанного: `state/manifest.json`.

## Автозапуск раз в сутки (macOS, 08:00)

```bash
chmod +x install_daily.sh run_daily.sh
./install_daily.sh install
```

Снять: `./install_daily.sh uninstall`  
Лог: `state/daily.log`

## Файлы

| Путь | Назначение |
|------|------------|
| `pdf_to_excel.py` | Точка входа |
| `sync_rfc.py` | Проверка сайта + download + skip |
| `convert_rfc.py` | OCR скана → Excel |
| `ocr.swift` | Apple Vision |
| `run_daily.sh` / `install_daily.sh` | Суточный запуск |

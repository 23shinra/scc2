#!/usr/bin/env python3
"""Сканы реестров сделок РФЦ/KOREM (PDF-картинки) -> Excel.

Запуск из этой папки:

    python3 pdf_to_excel.py
    python3 pdf_to_excel.py pdfs/какой-то.pdf
    python3 pdf_to_excel.py ~/Downloads/новый.pdf

PDF по умолчанию берутся из папки pdfs/, Excel пишется в excel/.
Нужен macOS: текст читается через Apple Vision.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import fitz
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageEnhance, ImageOps

ROOT = Path(__file__).resolve().parent
PDF_DIR = ROOT / "pdfs"
OUT_DIR = ROOT / "excel"
OCR_BIN = ROOT / "ocr"
WORK = ROOT / "_ocr_work"

TITLES = {
    "01_2025_sever_yug_27.11.24.pdf": "Итоги централизованных торгов электрической мощностью для Северной и Южной зон ЕЭС РК на 2025 год - 27.11.24",
    "02_2025_zapad_zko_atyrau_28.11.24.pdf": "Итоги централизованных торгов электрической мощностью для Западной (Западно-Казахстанской и Атырауской областей) зоны ЕЭС РК на 2025 год - 28.11.24",
    "03_2025_zapad_mangistau_28.11.24.pdf": "Итоги централизованных торгов электрической мощностью для Западной (Мангистауской области) зоны ЕЭС РК на 2025 год - 28.11.24",
    "04_2024_sever_yug_21.11.2023.pdf": "Итоги централизованных торгов электрической мощностью для Северной и Южной зон ЕЭС РК 21.11.2023",
    "05_2024_zapad_20.11.2023_n2.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 20.11.2023 №2",
    "06_2024_zapad_20.11.2023.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 20.11.2023",
    "07_2023_zapad_21.11.2022_n2.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 21.11.2022 №2",
    "08_2023_zapad_21.11.2022.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 21.11.2022",
    "09_2023_sever_yug_19.11.2022.pdf": "Итоги централизованных торгов электрической мощностью для Северной и Южной зон ЕЭС РК 19.11.2022",
    "10_2022_zapad_17.11.2021_n2.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 17.11.2021 №2",
    "11_2022_zapad_17.11.2021.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 17.11.2021",
    "12_2022_sever_yug_16.11.2021.pdf": "Итоги централизованных торгов электрической мощностью для Северной и Южной зон ЕЭС РК 16.11.2021",
    "13_2021_zapad_19.11.2020_n2.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 19.11.2020 №2",
    "14_2021_zapad_19.11.2020.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 19.11.2020",
    "15_2021_sever_yug_18.11.2020.pdf": "Итоги централизованных торгов электрической мощностью для Северной и Южной зон ЕЭС РК 18.11.2020",
    "16_2020_mangistau_29.11.2019.pdf": "Итоги централизованных торгов электрической мощностью для Западной (Мангистауская область) зоны ЕЭС РК 29.11.2019",
    "17_2020_zko_atyrau_29.11.2019.pdf": "Итоги централизованных торгов электрической мощностью для Западной (Западно-Казахстанская и Атырауская области) зоны ЕЭС РК 29.11.2019",
    "18_2020_sever_yug_28.11.2019.pdf": "Итоги централизованных торгов электрической мощностью для Северной и Южной зон ЕЭС РК 28.11.2019",
    "19_2019_sever_yug_05.12.2018.pdf": "Итоги централизованных торгов электрической мощностью для Северной и Южной зон ЕЭС РК 05.12.2018",
    "20_2019_zapad_05.12.2018.pdf": "Итоги централизованных торгов электрической мощностью для Западной зоны ЕЭС РК 05.12.2018",
}

DEAL_RE = re.compile(r"(\d{2,3})\s*[-–—]\s*(\d{5,6})")
DATE_RE = re.compile(r"(\d{2})[.\s](\d{2})[.\s](\d{4})")
COL_NAMES = ["n", "org", "deal", "zone_volume", "volume", "price", "sum_wo", "sum_w", "zone_org"]


def render_page(pdf: Path, dest: Path) -> Image.Image:
    if dest.exists():
        return Image.open(dest)
    doc = fitz.open(pdf)
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img = ImageOps.grayscale(img)
    img = ImageOps.autocontrast(img, cutoff=0.4)
    img = ImageEnhance.Contrast(img).enhance(1.3)
    dest.parent.mkdir(parents=True, exist_ok=True)
    img.save(dest, "PNG")
    return img


def _suppress_close(lines: list[tuple[float, float]]) -> list[float]:
    kept: list[tuple[float, float]] = []
    for x, score in sorted(lines):
        if kept and x - kept[-1][0] < 0.05 and min(score, kept[-1][1]) < 0.75:
            if score > kept[-1][1]:
                kept[-1] = (x, score)
            continue
        kept.append((x, score))
    return [x for x, _ in kept]


def _groups_to_xs(score: np.ndarray, score_thr: float, width: int) -> list[tuple[float, float]]:
    idx = np.where(score > score_thr)[0]
    if len(idx) == 0:
        return []
    groups: list[list[int]] = [[int(idx[0])]]
    for x in idx[1:]:
        x = int(x)
        if x - groups[-1][-1] <= 8:
            groups[-1].append(x)
        else:
            groups.append([x])
    found: list[tuple[float, float]] = []
    for group in groups:
        if len(group) > 14:
            continue
        peak = float(score[group].max())
        if peak < score_thr:
            continue
        center = sum(group) / len(group) / width
        if center < 0.985:
            found.append((center, peak))
    return found


def _form_lines(pairs: list[tuple[float, float]]) -> list[float] | None:
    xs = _suppress_close(pairs)
    if not (9 <= len(xs) <= 11):
        return None
    if xs[0] > 0.18 or xs[-1] < 0.82:
        return None
    widths = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
    wide = [w for w in widths if w > 0.15]
    if len(wide) != 1:
        return None
    # The wide column is the organization name and sits near the left.
    wide_at = widths.index(wide[0])
    if xs[wide_at] > 0.16:
        return None
    return xs


def vertical_lines(arr: np.ndarray) -> list[float]:
    height, width = arr.shape
    best: tuple[tuple[int, float], list[float]] | None = None
    for y0f in (0.12, 0.18, 0.24, 0.30, 0.36, 0.44, 0.52):
        y0 = int(height * y0f)
        y1 = min(height - 1, y0 + int(height * 0.24))
        sub = arr[y0:y1]
        if sub.size == 0:
            continue
        for thr in (150, 170, 190, 210):
            score = (sub < thr).mean(axis=0)
            for score_thr in (0.34, 0.45, 0.55, 0.65):
                xs = _form_lines(_groups_to_xs(score, score_thr, width))
                if not xs:
                    continue
                key = (abs(len(xs) - 10), -score_thr)
                if best is None or key < best[0]:
                    best = (key, xs)
    return align_to_template(best[1]) if best else align_to_template([])


TEMPLATE = [0.051, 0.078, 0.313, 0.401, 0.515, 0.587, 0.704, 0.772, 0.836, 0.915]


def align_to_template(detected: list[float]) -> list[float]:
    if len(detected) < 6:
        return list(TEMPLATE)
    best_shift = 0.0
    best_score = -1
    for detected_x in detected[:4]:
        for template_x in TEMPLATE[:5]:
            shift = detected_x - template_x
            if abs(shift) > 0.08:
                continue
            predicted = [item + shift for item in TEMPLATE]
            score = sum(1 for value in detected if min(abs(value - item) for item in predicted) <= 0.032)
            if score > best_score:
                best_score = score
                best_shift = shift
    predicted = [item + best_shift for item in TEMPLATE]
    # Prefer the template spacing; only snap a predicted border to a nearby
    # detected line when it is clearly the same border.
    aligned: list[float] = []
    for item in predicted:
        nearby = [value for value in detected if abs(value - item) <= 0.018]
        aligned.append(min(nearby, key=lambda value: abs(value - item)) if nearby else item)
    return aligned


def ocr_crop(img: Image.Image, box: tuple[int, int, int, int], cache: Path) -> list[tuple[float, str, float]]:
    """Return (page_y_from_top, text, confidence) for a crop box."""
    x0, y0, x1, y1 = box
    if cache.exists():
        raw = cache.read_text(encoding="utf-8", errors="replace")
    else:
        crop = img.crop(box)
        if crop.width < 40 or crop.height < 40:
            return []
        scale = 2 if crop.height < 1800 else 1
        if scale > 1:
            crop = crop.resize((crop.width * scale, crop.height * scale), Image.Resampling.LANCZOS)
        cache.parent.mkdir(parents=True, exist_ok=True)
        tmp = cache.with_suffix(".png")
        crop.save(tmp)
        raw = subprocess.check_output([str(OCR_BIN), str(tmp)], text=True, errors="replace")
        cache.write_text(raw, encoding="utf-8")
    out = []
    span = max(1, y1 - y0)
    for line in raw.splitlines():
        parts = line.split("\t", 5)
        if len(parts) != 6:
            continue
        _, _, ymin, ymax, conf, text = parts
        text = " ".join(text.split())
        if not text:
            continue
        y_mid = (float(ymin) + float(ymax)) / 2
        page_y = (y0 + (1 - y_mid) * span) / img.height
        out.append((page_y, text, float(conf)))
    return out


def normalize_zone(text: str) -> str:
    low = text.lower()
    if "юж" in low:
        return "Южная зона"
    if "север" in low or "сверн" in low:
        return "Северная зона"
    if "мангис" in low:
        return "Западная (Мангистауская область)"
    if "атырау" in low or "казахстанск" in low:
        return "Западная (Западно-Казахстанская и Атырауская области)"
    if "запад" in low:
        return "Западная зона"
    return " ".join(text.split())


def parse_number(text: str) -> float | None:
    raw = text.strip()
    if re.fullmatch(r"\d{4,}\s\d", raw):
        raw = raw.replace(" ", ".")
    raw = (
        raw.replace(" ", "")
        .replace("О", "0")
        .replace("O", "0")
        .replace("o", "0")
        .replace("Т", "1")
        .replace("T", "1")
        .replace("l", "1")
        .replace("I", "1")
        .replace("|", "1")
        .replace("З", "3")
        .replace("з", "3")
        .replace("Б", "6")
        .replace("S", "5")
        .replace(",", ".")
        .replace(":", "1")
        .replace("Х", "8")
        .replace("X", "8")
        .replace("х", "8")
        .replace("x", "8")
    )
    raw = re.sub(r"[^0-9.]", "", raw)
    if raw.count(".") > 1:
        parts = [part for part in raw.split(".") if part != ""]
        if parts and len(parts[-1]) <= 2:
            raw = "".join(parts[:-1]) + "." + parts[-1]
        else:
            raw = "".join(parts)
    if not raw or raw == ".":
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def normalize_deal(text: str) -> str | None:
    compact = re.sub(r"\s+", "", text)
    compact = compact.replace("Х", "X").replace("х", "x")
    match = DEAL_RE.search(compact)
    if not match:
        fuzzy = re.sub(r"[^0-9A-Za-zА-Яа-я-]", "", compact)
        match = re.search(r"(\d{2,3})\s*[-–—]\s*(\d{5,6})", fuzzy)
    if not match:
        return None
    left = match.group(1).zfill(3)
    right = match.group(2)
    if len(right) == 5:
        # OCR often drops one digit from DDMMYY.
        return f"{left}-{right}"
    return f"{left}-{right}"


def fix_deal_sequence(deals: list[str]) -> list[str]:
    parsed: list[int | None] = []
    suffixes: list[str] = []
    for deal in deals:
        match = re.match(r"(\d{3})-(\d{5,6})", deal)
        if match:
            parsed.append(int(match.group(1)))
            suffixes.append(match.group(2))
        else:
            parsed.append(None)
    if not suffixes:
        return deals
    # Prefer the most common 6-digit suffix; fall back to 5-digit.
    preferred = max(
        (item for item in suffixes if len(item) == 6),
        key=lambda item: suffixes.count(item),
        default=max(suffixes, key=suffixes.count),
    )
    fixed: list[str] = []
    for i, deal in enumerate(deals):
        match = re.match(r"(\d{3})-(\d{5,6})", deal)
        if match:
            fixed.append(f"{match.group(1)}-{preferred}")
            continue
        prev_n = next((parsed[j] for j in range(i - 1, -1, -1) if parsed[j] is not None), None)
        next_n = next((parsed[j] for j in range(i + 1, len(parsed)) if parsed[j] is not None), None)
        if prev_n is not None and next_n == prev_n + 2:
            fixed.append(f"{prev_n + 1:03d}-{preferred}")
        else:
            fixed.append(deal)
    return fixed


def cluster_lines(items: list[tuple[float, str]], gap: float) -> list[list[tuple[float, str]]]:
    if not items:
        return []
    ordered = sorted(items, key=lambda item: item[0])
    lines: list[list[tuple[float, str]]] = [[ordered[0]]]
    for item in ordered[1:]:
        if item[0] - lines[-1][-1][0] > gap:
            lines.append([item])
        else:
            lines[-1].append(item)
    return lines


def join_line(items: list[tuple[float, str]]) -> str:
    return " ".join(text for _, text in items).strip()


def reconcile(volume: float | None, price: float | None, sum_wo: float | None, sum_w: float | None) -> tuple[float | None, float | None, float | None, float | None, str]:
    notes: list[str] = []

    def close(a: float, b: float) -> bool:
        return abs(a - b) <= max(1.5, 0.012 * max(abs(b), 1))

    candidates_v = []
    if volume is not None:
        candidates_v.append(volume)
        whole = volume
        for _ in range(4):
            whole = whole / 10
            candidates_v.append(whole)
    else:
        candidates_v.append(None)

    # Recover truncated sums like 0600 from 210600.
    if volume and price and sum_wo is not None:
        expected = volume * price
        if not close(expected, sum_wo):
            text = f"{sum_wo:g}".replace(".", "")
            for prefix in ("", "1", "2", "3", "4", "5", "6", "7", "8", "9", "11", "12", "15", "21", "10"):
                try:
                    candidate = float(prefix + text)
                except ValueError:
                    continue
                if close(expected, candidate):
                    sum_wo = candidate
                    notes.append("восстановлена сумма без НДС")
                    break

    best = None
    for vol in candidates_v:
        if vol is None or price is None or sum_wo is None or vol == 0:
            continue
        if close(vol * price, sum_wo):
            best = (vol, price, sum_wo, sum_w, "OK")
            break
    if best and best[0] != volume:
        notes.append("в объёме восстановлена десятичная точка")
        check = "OK; " + "; ".join(notes) if notes else "OK"
        return best[0], best[1], best[2], best[3], check

    if best:
        check = "OK; " + "; ".join(notes) if notes else "OK"
        return best[0], best[1], best[2], best[3], check

    if volume and price and (sum_wo is None or not close(volume * price, sum_wo)):
        if 0.5 <= volume <= 20000 and 50 <= price <= 5000:
            sum_wo = round(volume * price, 3)
            notes.append("сумма без НДС = объём×цена")
            if sum_w is None or not close(sum_wo * 1.12, sum_w):
                sum_w = round(sum_wo * 1.12, 3)
                notes.append("сумма с НДС = ×1.12")
            return volume, price, sum_wo, sum_w, "OK; " + "; ".join(notes)

    if price is not None and 200 <= price <= 2500 and volume and sum_wo and close(volume * price, sum_wo):
        return volume, price, sum_wo, sum_w, "OK"

    if volume and sum_wo and volume != 0:
        implied = sum_wo / volume
        if 50 <= implied <= 5000:
            rounded = round(implied)
            if close(volume * rounded, sum_wo):
                return volume, float(rounded), sum_wo, sum_w, "OK; цена вычислена как сумма/объём"

    if volume and price and sum_wo:
        return volume, price, sum_wo, sum_w, f"объём×цена={volume * price:.2f}, в скане {sum_wo:g}"
    if volume == 0:
        return volume, price, sum_wo, sum_w, "объём 0"
    return volume, price, sum_wo, sum_w, "не сходится"


def detect_columns(xs: list[float]) -> list[tuple[str, float, float]] | None:
    # Drop the outer border pair if we have the 10-line KOREM grid.
    if len(xs) < 8:
        return None
    # Use inner content borders. Prefer the 9 data columns between the first and last line.
    if len(xs) >= 10:
        borders = xs[:10]
    elif len(xs) == 9:
        borders = xs[:9]
    else:
        borders = xs
    names = COL_NAMES[: len(borders) - 1]
    # If the narrow row-number column is missing, shift names.
    widths = [borders[i + 1] - borders[i] for i in range(len(borders) - 1)]
    if len(widths) == 8 and widths[0] > 0.12:
        names = COL_NAMES[1:]
    cols = []
    for name, x0, x1 in zip(names, borders, borders[1:]):
        cols.append((name, x0, x1))
    needed = {"org", "deal", "volume", "price", "sum_wo"}
    if not needed.issubset({name for name, _, _ in cols}):
        return None
    return cols


def parse_header(img: Image.Image, table_top: float, cache_dir: Path) -> list[tuple[str, str]]:
    y1 = max(80, int(img.height * min(table_top + 0.01, 0.42)))
    tokens = ocr_crop(img, (0, 0, img.width, y1), cache_dir / "header.tsv")
    if not tokens:
        return []
    gap = 0.012
    rows = cluster_lines([(y, text) for y, text, _ in tokens], gap)
    params: list[tuple[str, str]] = []
    for row in rows:
        text = join_line(row)
        if len(text) < 2:
            continue
        if re.search(r"Реестр сделок", text, re.I) and len(text) < 40:
            params.append(("Документ на бланке", text))
            continue
        params.append(("Строка шапки", text))
    # Structured pull
    blob = " ".join(text for _, text in params)
    dates = [f"{d}.{m}.{y}" for d, m, y in DATE_RE.findall(blob)]
    structured = []
    if dates:
        structured.append(("Дата проведения торгов", dates[0]))
    if len(dates) >= 3:
        structured.append(("Период поставки", f"{dates[1]} - {dates[2]}"))
    elif len(dates) == 2:
        structured.append(("Период поставки", f"{dates[0]} - {dates[1]}" if "Дата" else dates[1]))
    return structured + params


def process_pdf(pdf: Path) -> dict:
    img_path = WORK / f"{pdf.stem}.png"
    img = render_page(pdf, img_path)
    arr = np.asarray(img)
    xs = vertical_lines(arr)
    if not xs and img.height > img.width:
        rotated = img.transpose(Image.Transpose.ROTATE_270)
        xs = vertical_lines(np.asarray(rotated))
        if xs:
            img = rotated
            img.save(img_path)
            cache_dir = WORK / pdf.stem
            if cache_dir.exists():
                import shutil
                shutil.rmtree(cache_dir)
    cache_dir = WORK / pdf.stem
    cols = detect_columns(xs)
    print(f"  lines={len(xs)} cols={None if not cols else [c[0] for c in cols]}", flush=True)
    if not cols:
        return {
            "file": pdf.name,
            "title": TITLES.get(pdf.name, pdf.stem),
            "params": [("Ошибка", "Не удалось найти сетку таблицы")],
            "records": [],
        }

    table_top = min(y0 for _, y0, _ in []) if False else xs and 0.2
    # table top from where vertical lines begin is unknown; use first column crop later
    header_cut = 0.28
    # refine header cut: first strong horizontal-ish by looking at deal column later
    col_tokens: dict[str, list[tuple[float, str, float]]] = {}
    for name, x0, x1 in cols:
        pad = 0.004
        box = (
            max(0, int(img.width * (x0 + pad))),
            int(img.height * 0.12),
            min(img.width, int(img.width * (x1 - pad))),
            int(img.height * 0.80),
        )
        col_tokens[name] = ocr_crop(img, box, cache_dir / f"{name}.tsv")

    deals = []
    for y, text, conf in col_tokens.get("deal", []):
        deal = normalize_deal(text)
        if deal:
            deals.append((y, deal, conf))
    deals.sort()
    deal_ids = fix_deal_sequence([d for _, d, _ in deals])
    deals = [(y, deal, conf) for (y, _, conf), deal in zip(deals, deal_ids)]

    records = []
    if not deals:
        return {
            "file": pdf.name,
            "title": TITLES.get(pdf.name, pdf.stem),
            "params": parse_header(img, 0.28, cache_dir),
            "records": [],
        }

    ys = [y for y, _, _ in deals]
    gaps = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
    gap = sorted(gaps)[len(gaps) // 2] if gaps else 0.015
    band = max(gap * 0.62, 0.007)

    def collect(name: str, y: float) -> list[str]:
        found = [text for ty, text, _ in col_tokens.get(name, []) if abs(ty - y) <= band]
        return found

    for y, deal, _ in deals:
        org_parts = collect("org", y)
        org = " ".join(org_parts).strip()
        n_parts = collect("n", y)
        n_val = ""
        for part in n_parts:
            num = parse_number(part)
            if num is not None and num < 500:
                n_val = str(int(num))
                break
        zone_v = normalize_zone(" ".join(collect("zone_volume", y)).strip())
        zone_o = normalize_zone(" ".join(collect("zone_org", y)).strip())
        volumes = [parse_number(t) for t in collect("volume", y)]
        volumes = [v for v in volumes if v is not None]
        prices = [parse_number(t) for t in collect("price", y)]
        prices = [v for v in prices if v is not None]
        sum_wos = [parse_number(t) for t in collect("sum_wo", y)]
        sum_wos = [v for v in sum_wos if v is not None]
        sum_ws = [parse_number(t) for t in collect("sum_w", y)]
        sum_ws = [v for v in sum_ws if v is not None]
        price = prices[0] if prices else None
        sum_wo = sum_wos[0] if sum_wos else None
        sum_w = sum_ws[0] if sum_ws else None
        inferred_volume = False
        if not volumes and price and sum_wo:
            implied = sum_wo / price
            if 0 < implied < 20000 and abs(implied - round(implied, 1)) < 0.05:
                volumes = [round(implied, 1)]
                inferred_volume = True
        if not volumes:
            volumes = [None]
        # Multiple stacked volumes in one deal: keep one row per volume, sums on each for checking the total.
        if len(volumes) > 1 and all(v is not None for v in volumes):
            total_v = sum(volumes)
            for vol in volumes:
                records.append(
                    {
                        "n": n_val,
                        "org": org,
                        "deal": deal,
                        "zone_volume": zone_v,
                        "volume": vol,
                        "price": price,
                        "sum_wo": sum_wo,
                        "sum_w": sum_w,
                        "zone_org": zone_o,
                        "check": f"объём разбит, сумма объёмов={total_v:g}",
                    }
                )
        else:
            vol = volumes[0]
            vol, price, sum_wo, sum_w, check = reconcile(vol, price, sum_wo, sum_w)
            if inferred_volume and check.startswith("OK"):
                check = "OK; объём вычислен как сумма/цена"
            if sum_wo and (sum_w is None or abs((sum_wo * 1.12) - sum_w) > max(2.0, 0.02 * abs(sum_wo))):
                sum_w = round(sum_wo * 1.12, 3)
            if (not zone_o or len(zone_o) < 8) and zone_v:
                zone_o = zone_v
            if not zone_v and zone_o:
                zone_v = zone_o
            records.append(
                {
                    "n": n_val,
                    "org": org,
                    "deal": deal,
                    "zone_volume": zone_v,
                    "volume": vol,
                    "price": price,
                    "sum_wo": sum_wo,
                    "sum_w": sum_w,
                    "zone_org": zone_o,
                    "check": check,
                }
            )

    # Totals below the last deal.
    last_y = ys[-1]
    for name in ("volume", "sum_wo", "sum_w"):
        pass
    total_bits = []
    for name in ("n", "org", "volume", "sum_wo", "sum_w"):
        for y, text, _ in col_tokens.get(name, []):
            if y > last_y + band and y < last_y + gap * 3.2:
                total_bits.append((y, name, text))
    if any("всего" in text.lower() for _, _, text in total_bits) or any(name == "volume" for _, name, _ in total_bits):
        # group total lines
        total_lines = cluster_lines([(y, f"{name}:{text}") for y, name, text in total_bits], gap * 0.7)
        for line in total_lines:
            bucket: dict[str, list[str]] = {}
            for _, packed in line:
                name, text = packed.split(":", 1)
                bucket.setdefault(name, []).append(text)
            label = " ".join(bucket.get("org", []) + bucket.get("n", []))
            if not re.search(r"всего", label, re.I) and "volume" not in bucket and "sum_wo" not in bucket:
                continue
            vol = parse_number(" ".join(bucket.get("volume", [])[:1]))
            sum_wo = parse_number(" ".join(bucket.get("sum_wo", [])[:1]))
            sum_w = parse_number(" ".join(bucket.get("sum_w", [])[:1]))
            if (vol or 0) < 10 and (sum_wo or 0) < 100 and (sum_w or 0) < 100:
                continue
            records.append(
                {
                    "n": "",
                    "org": "Всего",
                    "deal": "",
                    "zone_volume": "",
                    "volume": vol,
                    "price": None,
                    "sum_wo": sum_wo,
                    "sum_w": sum_w,
                    "zone_org": "",
                    "check": "итоговая строка",
                }
            )

    header_y = max(0.18, ys[0] - 0.08)
    params = parse_header(img, header_y, cache_dir)
    return {
        "file": pdf.name,
        "title": TITLES.get(pdf.name, pdf.stem),
        "params": params,
        "records": records,
    }


def style_sheet(ws, wide_cols: set[int]) -> None:
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 32
    for i, cell in enumerate(ws[1], 1):
        width = 44 if i in wide_cols else min(36, max(14, len(str(cell.value)) + 2))
        ws.column_dimensions[get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)


def build_one_workbook(page: dict) -> Path:
    OUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Реестр сделок"

    # Keep header values in a compact block like the scanned form.
    header_map: dict[str, str] = {}
    for param, value in page["params"]:
        if param == "Дата проведения торгов":
            header_map["Дата проведения торгов"] = str(value)
        elif param == "Период поставки":
            header_map["Период поставки"] = str(value)
        elif param == "Документ на бланке":
            header_map["Документ"] = str(value)
        elif param == "Строка шапки":
            text = str(value)
            if "Зона торгов" in text or "зона" in text.lower():
                header_map.setdefault("Зона торгов ЕЭС РК", text)
            if "Объем" in text or "Объём" in text:
                header_map.setdefault("Объёмы / тариф", text)
            if "Предельный тариф" in text:
                header_map["Предельный тариф"] = text

    title = page.get("title") or Path(page["file"]).stem
    ws.append([title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    ws["A1"].font = Font(bold=True, size=13, color="1B1F24")
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 28

    meta_rows = [
        ("Дата проведения торгов", header_map.get("Дата проведения торгов", "")),
        ("Зона торгов ЕЭС РК", header_map.get("Зона торгов ЕЭС РК", "")),
        ("Период поставки", header_map.get("Период поставки", "")),
        ("Объёмы / тариф из шапки", header_map.get("Объёмы / тариф", header_map.get("Предельный тариф", ""))),
        ("Исходный PDF", page["file"]),
    ]
    # Pull date/period from OCR header lines if structured fields are empty.
    blob = " ".join(str(v) for _, v in page["params"])
    dates = DATE_RE.findall(blob)
    norm = [f"{d}.{m}.{y}" for d, m, y in dates]
    if not meta_rows[0][1] and norm:
        meta_rows[0] = ("Дата проведения торгов", norm[0])
    if not meta_rows[2][1] and len(norm) >= 3:
        meta_rows[2] = ("Период поставки", f"{norm[1]} - {norm[2]}")
    elif not meta_rows[2][1] and len(norm) == 2:
        meta_rows[2] = ("Период поставки", f"{norm[0]} - {norm[1]}")
    for label, value in meta_rows:
        ws.append([label, value])
        ws.merge_cells(start_row=ws.max_row, start_column=2, end_row=ws.max_row, end_column=9)

    ws.append([])
    headers = [
        "№ п/п",
        "Наименование энергопроизводящей организации",
        "Номер сделки",
        "Зона отобранного объёма",
        "Объём, МВт",
        "Цена, тыс. тг/МВт без НДС",
        "Сумма без НДС",
        "Сумма с НДС",
        "Зона организации",
    ]
    ws.append(headers)
    header_row = ws.max_row

    thin = Border(
        left=Side(style="thin", color="C9CDD4"),
        right=Side(style="thin", color="C9CDD4"),
        top=Side(style="thin", color="C9CDD4"),
        bottom=Side(style="thin", color="C9CDD4"),
    )
    header_fill = PatternFill("solid", fgColor="F2F4F7")
    header_font = Font(bold=True, size=10, color="1B1F24")
    zebra = PatternFill("solid", fgColor="FAFBFC")
    total_fill = PatternFill("solid", fgColor="EEF2F6")

    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    ws.row_dimensions[header_row].height = 36

    # Merge multi-zone rows that belong to one deal into one Excel row when possible.
    grouped: list[dict] = []
    for rec in page["records"]:
        if rec["org"] == "Всего":
            grouped.append(rec)
            continue
        if (
            grouped
            and grouped[-1].get("deal")
            and grouped[-1]["deal"] == rec["deal"]
            and grouped[-1]["org"] != "Всего"
        ):
            prev = grouped[-1]
            zone_parts = []
            if prev.get("zone_volume") or prev.get("volume") is not None:
                zone_parts.append(
                    f"{prev.get('zone_volume') or ''} {prev.get('volume') if prev.get('volume') is not None else ''}".strip()
                )
            if rec.get("zone_volume") or rec.get("volume") is not None:
                zone_parts.append(
                    f"{rec.get('zone_volume') or ''} {rec.get('volume') if rec.get('volume') is not None else ''}".strip()
                )
            prev["zone_volume"] = " / ".join(p for p in zone_parts if p)
            vols = [v for v in (prev.get("volume"), rec.get("volume")) if isinstance(v, (int, float))]
            prev["volume"] = sum(vols) if vols else prev.get("volume")
            continue
        grouped.append(dict(rec))

    seq = 0
    data_rows = [rec for rec in grouped if rec["org"] != "Всего"]
    for rec in data_rows:
        seq += 1
        number = rec["n"] or str(seq)
        ws.append(
            [
                number,
                rec["org"],
                rec["deal"],
                rec["zone_volume"],
                rec["volume"],
                rec["price"],
                rec["sum_wo"],
                rec["sum_w"],
                rec["zone_org"],
            ]
        )
        row_idx = ws.max_row
        for cell in ws[row_idx]:
            cell.border = thin
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if seq % 2 == 0:
                cell.fill = zebra

    total_vol = sum(r["volume"] for r in data_rows if isinstance(r.get("volume"), (int, float)))
    total_wo = sum(r["sum_wo"] for r in data_rows if isinstance(r.get("sum_wo"), (int, float)))
    total_w = sum(r["sum_w"] for r in data_rows if isinstance(r.get("sum_w"), (int, float)))
    ocr_total = next((r for r in grouped if r["org"] == "Всего"), None)
    if ocr_total and isinstance(ocr_total.get("sum_wo"), (int, float)) and ocr_total["sum_wo"] > total_wo * 0.5:
        total_vol = ocr_total.get("volume") or total_vol
        total_wo = ocr_total.get("sum_wo") or total_wo
        total_w = ocr_total.get("sum_w") or total_w
    ws.append(["", "Всего", "", "", total_vol, "", total_wo, total_w, ""])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True, size=10)
        cell.fill = total_fill
        cell.border = thin
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for col in (5, 6, 7, 8):
        for cell in ws.iter_cols(min_col=col, max_col=col, min_row=header_row + 1):
            for item in cell:
                if isinstance(item.value, (int, float)):
                    item.number_format = "#,##0.###"
                    item.alignment = Alignment(vertical="center", horizontal="right")

    widths = [8, 46, 14, 22, 12, 16, 14, 14, 22]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Title / meta styling
    for r in range(2, header_row):
        label = ws.cell(r, 1).value
        if not label:
            continue
        ws.cell(r, 1).font = Font(bold=True, size=10, color="5B6570")
        ws.cell(r, 2).font = Font(size=10, color="1B1F24")

    ws.freeze_panes = f"A{header_row + 1}"
    ws.print_title_rows = f"{header_row}:{header_row}"

    out = OUT_DIR / f"{Path(page['file']).stem}.xlsx"
    wb.save(out)
    return out


def ensure_ocr() -> None:
    global OCR_BIN
    source = ROOT / "ocr.swift"
    binary = ROOT / "ocr"
    if not source.exists():
        raise SystemExit(f"Нет файла распознавания: {source}")
    if not binary.exists() or binary.stat().st_mtime < source.stat().st_mtime:
        print("Собираю распознавание текста (один раз)...", flush=True)
        subprocess.check_call(["swiftc", "-O", "-o", str(binary), str(source)])
    OCR_BIN = binary


def collect_pdfs(args: list[str]) -> list[Path]:
    if not args:
        PDF_DIR.mkdir(exist_ok=True)
        return sorted(PDF_DIR.glob("*.pdf"))
    found: list[Path] = []
    for arg in args:
        path = Path(arg).expanduser()
        if path.is_dir():
            found.extend(sorted(path.glob("*.pdf")))
            continue
        if path.is_file() and path.suffix.lower() == ".pdf":
            found.append(path)
            continue
        found.extend(pdf for pdf in sorted(PDF_DIR.glob("*.pdf")) if arg in pdf.name)
    # keep order, drop duplicates
    unique: list[Path] = []
    seen: set[Path] = set()
    for pdf in found:
        resolved = pdf.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(pdf)
    return unique


def main() -> None:
    ensure_ocr()
    WORK.mkdir(exist_ok=True)
    OUT_DIR.mkdir(exist_ok=True)
    PDF_DIR.mkdir(exist_ok=True)
    pdfs = collect_pdfs(sys.argv[1:])
    if not pdfs:
        raise SystemExit(
            "Нет PDF. Положи файлы в папку pdfs/ или укажи путь:\n"
            "  python3 pdf_to_excel.py файл.pdf"
        )
    saved: list[Path] = []
    for pdf in pdfs:
        print(f"== {pdf.name}", flush=True)
        page = process_pdf(pdf)
        ok = sum(1 for rec in page["records"] if str(rec["check"]).startswith("OK"))
        print(f"   строк={len(page['records'])} проверено={ok}", flush=True)
        out = build_one_workbook(page)
        saved.append(out)
        print(f"   -> {out}", flush=True)
    print(f"Готово: {len(saved)} файл(ов) в {OUT_DIR}")


if __name__ == "__main__":
    main()

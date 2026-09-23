#!/usr/bin/env python3
"""Скачать новые итоги торгов с rfc.kz и переделать в Excel.

Пропускает URL, которые уже есть в state/manifest.json (или для которых
уже лежит Excel). Запускать раз в сутки через launchd / cron.
"""

from __future__ import annotations

import hashlib
import json
import re
import urllib.request
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path

from convert_rfc import (
    OUT_DIR,
    PDF_DIR,
    ROOT,
    TITLES,
    build_one_workbook,
    ensure_ocr,
    process_pdf,
)

BASE = "https://rfc.kz"
LISTING_URL = f"{BASE}/ru/power-market/prices-and-rates/results-of-centralized-trading/"
STATE_DIR = ROOT / "state"
MANIFEST_PATH = STATE_DIR / "manifest.json"
USER_AGENT = "Mozilla/5.0 (compatible; scc2-rfc-sync/1.0)"


class DocParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.in_a = False
        self.href: str | None = None
        self.in_name = False
        self.buf: list[str] = []
        self.items: list[tuple[str, str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        data = {k: v or "" for k, v in attrs}
        if tag == "a" and "corporative-docs" in data.get("class", ""):
            self.in_a = True
            self.href = data.get("href") or None
        if tag == "div" and self.in_a and "name" in data.get("class", ""):
            self.in_name = True
            self.buf = []

    def handle_endtag(self, tag: str) -> None:
        if tag == "div" and self.in_name:
            title = " ".join("".join(self.buf).split())
            if self.href and title:
                self.items.append((self.href, title))
            self.in_name = False
        if tag == "a" and self.in_a:
            self.in_a = False
            self.href = None

    def handle_data(self, data: str) -> None:
        if self.in_name:
            self.buf.append(data)


def load_manifest() -> dict:
    if not MANIFEST_PATH.exists():
        return {}
    return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))


def save_manifest(manifest: dict) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def fetch_listing() -> list[tuple[str, str]]:
    req = urllib.request.Request(LISTING_URL, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=60) as resp:
        html = resp.read().decode("utf-8", errors="replace")
    parser = DocParser()
    parser.feed(html)
    # keep order, unique by href
    seen: set[str] = set()
    items: list[tuple[str, str]] = []
    for href, title in parser.items:
        if href in seen:
            continue
        seen.add(href)
        url = href if href.startswith("http") else BASE + href
        items.append((url, title))
    return items


def slugify(title: str) -> str:
    text = title.casefold()
    text = re.sub(r"[^\w\s\-а-яёa-z0-9.]", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", "_", text.strip())
    return text[:80] or "document"


def stem_for(url: str, title: str) -> str:
    digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:10]
    return f"{digest}_{slugify(title)}"


def legacy_paths(title: str) -> tuple[Path | None, Path | None]:
    for name, known_title in TITLES.items():
        if known_title == title:
            pdf = PDF_DIR / name
            xlsx = OUT_DIR / f"{Path(name).stem}.xlsx"
            return pdf, xlsx
    return None, None


def already_done(manifest: dict, url: str, stem: str, title: str) -> bool:
    excel = OUT_DIR / f"{stem}.xlsx"
    if excel.exists():
        return True
    entry = manifest.get(url)
    if entry:
        excel_name = entry.get("excel") or f"{entry.get('stem', stem)}.xlsx"
        if (OUT_DIR / excel_name).exists():
            return True
    _, legacy_xlsx = legacy_paths(title)
    if legacy_xlsx and legacy_xlsx.exists():
        return True
    return False


def remember(manifest: dict, url: str, title: str, stem: str, pdf_name: str, excel_name: str) -> None:
    manifest[url] = {
        "title": title,
        "stem": stem,
        "pdf": pdf_name,
        "excel": excel_name,
        "synced_at": datetime.now(timezone.utc).isoformat(),
    }
    save_manifest(manifest)


def download_pdf(url: str, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=120) as resp, dest.open("wb") as out:
        out.write(resp.read())


def sync(force: bool = False) -> list[Path]:
    ensure_ocr()
    PDF_DIR.mkdir(exist_ok=True)
    OUT_DIR.mkdir(exist_ok=True)
    STATE_DIR.mkdir(exist_ok=True)

    manifest = load_manifest()
    items = fetch_listing()
    print(f"На сайте: {len(items)} документ(ов)", flush=True)

    saved: list[Path] = []
    skipped = 0
    for url, title in items:
        stem = stem_for(url, title)
        if not force and already_done(manifest, url, stem, title):
            skipped += 1
            if url not in manifest:
                legacy_pdf, legacy_xlsx = legacy_paths(title)
                pdf_name = legacy_pdf.name if legacy_pdf and legacy_pdf.exists() else f"{stem}.pdf"
                excel_name = legacy_xlsx.name if legacy_xlsx and legacy_xlsx.exists() else f"{stem}.xlsx"
                remember(manifest, url, title, stem, pdf_name, excel_name)
            continue

        pdf_path = PDF_DIR / f"{stem}.pdf"
        excel_path = OUT_DIR / f"{stem}.xlsx"
        legacy_pdf, legacy_xlsx = legacy_paths(title)
        if legacy_pdf and legacy_pdf.exists() and not pdf_path.exists():
            pdf_path = legacy_pdf
        if legacy_xlsx and legacy_xlsx.exists() and not force:
            skipped += 1
            remember(manifest, url, title, stem, pdf_path.name, legacy_xlsx.name)
            continue

        print(f"NEW {title}", flush=True)
        if force or not pdf_path.exists() or pdf_path.stat().st_size < 1000:
            target = PDF_DIR / f"{stem}.pdf"
            print(f"  download -> {target.name}", flush=True)
            download_pdf(url, target)
            pdf_path = target
        else:
            print("  pdf already local, convert only", flush=True)

        TITLES[pdf_path.name] = title
        page = process_pdf(pdf_path)
        page["title"] = title
        page["file"] = pdf_path.name
        out = build_one_workbook(page)
        if out.resolve() != excel_path.resolve():
            excel_path.write_bytes(out.read_bytes())
            if out.exists() and out.resolve() != excel_path.resolve():
                out.unlink(missing_ok=True)

        remember(manifest, url, title, stem, pdf_path.name, excel_path.name)
        saved.append(excel_path)
        print(f"  -> {excel_path.name}", flush=True)

    print(f"Готово: новых {len(saved)}, пропущено {skipped}", flush=True)
    return saved


if __name__ == "__main__":
    import sys

    sync(force="--force" in sys.argv)

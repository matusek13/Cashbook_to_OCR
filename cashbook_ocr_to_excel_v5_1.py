#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cashbook OCR -> Excel (v5.1)

Bazuję na działającym v5 i dodaję postprocessing po OCR:
- globalna naprawa numeracji,
- czyszczenie śmieciowych tokenów w kategoriach / opisie,
- normalizacja kwot,
- próba odzyskania brakującej wpłaty/wypłaty na podstawie bilansu,
- usuwanie pustych / duchowych rekordów.

Uruchomienie:
    python cashbook_ocr_to_excel_v5_1.py "snipping tool-20260530T093141Z-3-001.zip" wynik.xlsx
"""

from __future__ import annotations

import argparse
import concurrent.futures as cf
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

import cv2
import numpy as np
import pandas as pd
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image


LANG = "deu+pol+eng"
MIN_CONF = 18

OUTPUT_COLUMNS = [
    "data",
    "nr",
    "opis",
    "kategoria",
    "proc",
    "vat",
    "wplata",
    "wyplata",
    "bilans",
    "source_file",
    "source_row",
]

NOISE_TOKENS = {
    "", "n", "u", "i", "l", "|", "'", '"', "rt", "ae", "st", "ss", "ss!", "on",
    "0", "o", "n!", "n'", "a", "x", "v"
}

AMOUNT_COLS = {"vat", "wplata", "wyplata", "bilans"}
TEXT_COLS = {"opis", "kategoria", "proc"}
COL_ORDER = ["data", "nr", "opis", "kategoria", "proc", "vat", "wplata", "wyplata", "bilans"]

# Fallback granice kolumn dla ~1078 px szerokości
FALLBACK_X_BOUNDS = [0, 88, 150, 432, 585, 625, 724, 845, 966, 1078]


def natural_key(path: Path):
    s = path.name.lower()
    nums = re.findall(r"\d+", s)
    return (int(nums[0]) if nums else 10**9, "niepelne" in s, s)


def ensure_tesseract_available() -> None:
    if shutil.which("tesseract") is None:
        raise RuntimeError(
            "Nie znaleziono programu tesseract w PATH. "
            "Zainstaluj Tesseract OCR i dodaj go do PATH."
        )


def list_pngs(folder: Path) -> List[Path]:
    files = [p for p in folder.rglob("*") if p.suffix.lower() == ".png"]
    return sorted(files, key=natural_key)


def extract_zip(zip_path: Path, temp_dir: Path) -> List[Path]:
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(temp_dir)
    return list_pngs(temp_dir)


def cluster_peaks(indices: np.ndarray) -> List[int]:
    if len(indices) == 0:
        return []
    clusters = []
    start = prev = int(indices[0])
    for v in indices[1:]:
        v = int(v)
        if v == prev + 1:
            prev = v
        else:
            clusters.append((start + prev) // 2)
            start = prev = v
    clusters.append((start + prev) // 2)
    return clusters


def unique_bounds(peaks: List[int], max_value: int) -> List[int]:
    bounds = sorted(set([0, max_value] + [int(p) for p in peaks if 0 <= int(p) <= max_value]))
    return bounds


def detect_grid_lines(gray: np.ndarray) -> Tuple[List[int], List[int]]:
    th = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 31, 15
    )

    h, w = gray.shape[:2]
    row_kernel_w = max(20, w // 30)
    col_kernel_h = max(20, h // 30)

    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (row_kernel_w, 1))
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, col_kernel_h))

    horiz = cv2.dilate(cv2.erode(th, horiz_kernel), horiz_kernel)
    vert = cv2.dilate(cv2.erode(th, vert_kernel), vert_kernel)

    proj_y = horiz.sum(axis=1)
    y_idx = np.where(proj_y > proj_y.max() * 0.35)[0]
    y_lines = cluster_peaks(y_idx)

    proj_x = vert.sum(axis=0)
    x_idx = np.where(proj_x > proj_x.max() * 0.35)[0]
    x_lines = cluster_peaks(x_idx)

    return y_lines, x_lines


def remove_table_lines(gray: np.ndarray) -> np.ndarray:
    th = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 31, 15
    )
    h, w = gray.shape[:2]
    row_kernel_w = max(20, w // 30)
    col_kernel_h = max(20, h // 30)

    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (row_kernel_w, 1))
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, col_kernel_h))

    horiz = cv2.morphologyEx(th, cv2.MORPH_OPEN, horiz_kernel)
    vert = cv2.morphologyEx(th, cv2.MORPH_OPEN, vert_kernel)

    mask = cv2.bitwise_or(horiz, vert)
    cleaned = cv2.bitwise_and(th, cv2.bitwise_not(mask))
    out = 255 - cleaned
    return out


def clean_text(text: str) -> str:
    if not text:
        return ""
    t = (
        text.replace("‘", "")
        .replace("’", "")
        .replace("`", "")
        .replace("´", "")
        .replace("|", " ")
        .replace("¦", " ")
        .replace("•", " ")
        .replace("—", " ")
        .replace("–", " ")
    )
    t = re.sub(r"\s+", " ", t).strip()
    return t


def drop_noise_tokens(text: str) -> str:
    t = clean_text(text)
    if not t:
        return ""
    if t.lower() in NOISE_TOKENS:
        return ""
    if len(t) <= 2 and not re.search(r"\d", t):
        return ""
    return t


def parse_date(text: str) -> str:
    if not text:
        return ""
    t = clean_text(str(text))
    m = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", t)
    if m:
        return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}.{m.group(3)}"
    digits = re.sub(r"\D", "", t)
    if len(digits) == 8:
        return f"{digits[:2]}.{digits[2:4]}.{digits[4:]}"
    if len(digits) == 7:
        digits = "0" + digits
        return f"{digits[:2]}.{digits[2:4]}.{digits[4:]}"
    return t


def parse_amount(text: str) -> str:
    if not text:
        return ""
    t = clean_text(str(text)).replace("€", "").replace("EUR", "").replace("eur", "")
    t = re.sub(r"\s+", "", t)
    m = re.search(r"[-+]?\d[\d.,]*\d|[-+]?\d", t)
    if not m:
        return ""
    num = m.group(0)

    if "," in num and "." in num:
        if num.rfind(",") > num.rfind("."):
            num = num.replace(".", "")
            return num
        num = num.replace(",", "")
        num = num.replace(".", ",")
        return num

    if "," in num:
        left, right = num.split(",", 1)
        left_digits = re.sub(r"\D", "", left)
        right_digits = re.sub(r"\D", "", right)
        if 1 <= len(right_digits) <= 2:
            return f"{left_digits},{right_digits.zfill(2)}"
        digits = left_digits + right_digits
        if len(digits) > 2:
            return f"{digits[:-2]},{digits[-2:]}"
        return digits

    if "." in num:
        left, right = num.split(".", 1)
        left_digits = re.sub(r"\D", "", left)
        right_digits = re.sub(r"\D", "", right)
        if 1 <= len(right_digits) <= 2:
            return f"{left_digits},{right_digits.zfill(2)}"
        digits = left_digits + right_digits
        if len(digits) > 2:
            return f"{digits[:-2]},{digits[-2:]}"
        return digits

    digits = re.sub(r"\D", "", num)
    if not digits:
        return ""
    if len(digits) <= 2:
        return f"{digits},00"
    return f"{digits[:-2]},{digits[-2:]}"


def parse_nr(text: str) -> str:
    if not text:
        return ""
    digits = re.sub(r"\D", "", clean_text(str(text)))
    return digits


def amount_to_float(text: str) -> Optional[float]:
    txt = parse_amount(text)
    if not txt:
        return None
    txt = txt.replace(".", "").replace(",", ".") if txt.count(",") == 1 else txt
    try:
        return float(txt)
    except Exception:
        return None


def float_to_amount(value: float) -> str:
    return f"{value:.2f}".replace(".", ",")


def choose_row_from_y(y: float, y_bounds: List[int]) -> int:
    return max(0, min(len(y_bounds) - 2, int(np.searchsorted(y_bounds, y, side="right") - 1)))


def choose_col_from_x(x: float, x_bounds: List[int]) -> int:
    return max(0, min(len(x_bounds) - 2, int(np.searchsorted(x_bounds, x, side="right") - 1)))


def row_is_meaningful(row: Dict[str, str]) -> bool:
    fields = [
        row.get("data", ""),
        row.get("nr", ""),
        row.get("opis", ""),
        row.get("vat", ""),
        row.get("wplata", ""),
        row.get("wyplata", ""),
        row.get("bilans", ""),
    ]
    if any(f for f in fields[:2]):
        return True
    if any(re.search(r"\d", row.get(k, "")) for k in ("vat", "wplata", "wyplata", "bilans")):
        return True
    if len(row.get("opis", "")) >= 3:
        return True
    return False


def merge_very_similar_rows(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    out = []
    prev = None
    for r in rows:
        key = (
            r.get("data", ""),
            r.get("nr", ""),
            r.get("opis", ""),
            r.get("kategoria", ""),
            r.get("proc", ""),
            r.get("vat", ""),
            r.get("wplata", ""),
            r.get("wyplata", ""),
            r.get("bilans", ""),
            r.get("source_file", ""),
        )
        if key == prev:
            continue
        out.append(r)
        prev = key
    return out


def normalize_ocr_cell(raw: str, col: str) -> str:
    raw = clean_text(raw)
    if col == "data":
        return parse_date(raw)
    if col == "nr":
        return parse_nr(raw)
    if col in AMOUNT_COLS:
        return parse_amount(raw)
    if col in TEXT_COLS:
        return drop_noise_tokens(raw)
    return raw


def ocr_page(image_path: Path) -> List[Dict[str, str]]:
    if "niepelne" in image_path.name.lower():
        return []

    gray0 = cv2.imread(str(image_path), cv2.IMREAD_GRAYSCALE)
    if gray0 is None:
        return []

    gray = cv2.resize(gray0, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)

    y_lines, x_lines = detect_grid_lines(gray)
    y_bounds = unique_bounds(y_lines, gray.shape[0])
    x_bounds = unique_bounds(x_lines, gray.shape[1])

    if len(y_bounds) < 3:
        y_bounds = [0, gray.shape[0]]
    if len(x_bounds) < 3:
        x_bounds = FALLBACK_X_BOUNDS.copy()
        x_bounds[-1] = gray.shape[1]

    cleaned = remove_table_lines(gray)
    pil = Image.fromarray(cleaned)

    df = pytesseract.image_to_data(
        pil,
        lang=LANG,
        config="--oem 1 --psm 6",
        output_type=pytesseract.Output.DATAFRAME,
    )

    if df is None or df.empty:
        return []

    df = df.dropna(subset=["text"]).copy()
    df["conf"] = pd.to_numeric(df["conf"], errors="coerce")
    df = df[df["conf"] >= MIN_CONF].copy()
    if df.empty:
        return []

    df["xcenter"] = df["left"] + df["width"] / 2.0
    df["ycenter"] = df["top"] + df["height"] / 2.0
    df["row"] = df["ycenter"].apply(lambda y: choose_row_from_y(y, y_bounds))
    df["col"] = df["xcenter"].apply(lambda x: choose_col_from_x(x, x_bounds))

    rows: List[Dict[str, str]] = []

    for row_idx, sub in df.groupby("row", sort=True):
        cells: Dict[int, List[Tuple[float, str]]] = {}
        for _, r in sub.iterrows():
            txt = clean_text(str(r["text"]))
            if not txt or txt in {"|", "-", "_"}:
                continue
            if len(txt) <= 2 and not re.search(r"\d", txt) and txt.lower() not in {"nr"}:
                continue
            cells.setdefault(int(r["col"]), []).append((float(r["left"]), txt))

        if not cells:
            continue

        out = {k: "" for k in OUTPUT_COLUMNS}

        for c_idx, c_name in enumerate(COL_ORDER):
            parts = cells.get(c_idx, [])
            parts.sort(key=lambda z: z[0])
            raw = clean_text(" ".join(p for _, p in parts))

            if c_name in {"kategoria", "proc"}:
                raw = drop_noise_tokens(raw)
            elif c_name == "opis":
                raw = drop_noise_tokens(raw)
            elif c_name == "nr":
                raw = parse_nr(raw)
            elif c_name in AMOUNT_COLS:
                raw = parse_amount(raw)
            elif c_name == "data":
                raw = parse_date(raw)

            out[c_name] = raw

        joined = " ".join(out[c] for c in COL_ORDER).lower()
        if any(k in joined for k in ("datum", "beschreibung", "restbetrag", "eingang", "ausgang")):
            continue

        if not row_is_meaningful(out):
            continue

        out["source_file"] = image_path.name
        out["source_row"] = str(row_idx)
        rows.append(out)

    return rows


def fix_row_fields(rows: List[Dict[str, str]]) -> None:
    """
    Postprocessing po złożeniu wszystkich stron:
    - globalna numeracja,
    - czyszczenie krótkich śmieci w text cols,
    - próba odzyskania wpłata/wypłata z bilansu,
    - domykanie braków.
    """
    # 1) globalne czyszczenie tekstu
    for r in rows:
        r["data"] = parse_date(r.get("data", ""))
        r["nr"] = parse_nr(r.get("nr", ""))
        r["opis"] = drop_noise_tokens(r.get("opis", ""))
        r["kategoria"] = drop_noise_tokens(r.get("kategoria", ""))
        r["proc"] = drop_noise_tokens(r.get("proc", ""))
        for k in ("vat", "wplata", "wyplata", "bilans"):
            r[k] = parse_amount(r.get(k, ""))

    # 2) globalna naprawa sekwencji numerów
    prev_nr: Optional[int] = None
    for r in rows:
        raw = parse_nr(r.get("nr", ""))
        nr_val = int(raw) if raw.isdigit() else None

        if prev_nr is None:
            if nr_val is not None:
                r["nr"] = str(nr_val)
                prev_nr = nr_val
            continue

        expected = prev_nr + 1
        if nr_val is None:
            r["nr"] = str(expected)
            prev_nr = expected
            continue

        if abs(nr_val - expected) > 2 or len(str(nr_val)) < len(str(expected)):
            r["nr"] = str(expected)
            prev_nr = expected
        else:
            r["nr"] = str(nr_val)
            prev_nr = nr_val

    # 3) próba naprawy pieniędzy na bazie bilansu
    prev_bal: Optional[float] = None
    for r in rows:
        bal = amount_to_float(r.get("bilans", ""))
        wp = amount_to_float(r.get("wplata", ""))
        wy = amount_to_float(r.get("wyplata", ""))

        if prev_bal is not None and bal is not None:
            diff = round(bal - prev_bal, 2)

            # jeśli brakuje kwoty i bilans wskazuje kierunek transakcji, uzupełnij
            if wp is None and wy is None:
                if diff > 0:
                    r["wplata"] = float_to_amount(diff)
                    wp = diff
                elif diff < 0:
                    r["wyplata"] = float_to_amount(abs(diff))
                    wy = abs(diff)

            # jeśli jest tylko jedna kwota, a bilans wskazuje przeciwny kierunek,
            # przerzuć ją do właściwej kolumny
            if wp is None and wy is not None and diff > 0 and abs(wy - diff) <= max(0.05, diff * 0.02):
                r["wplata"] = float_to_amount(wy)
                r["wyplata"] = ""
                wp, wy = wy, None
            elif wy is None and wp is not None and diff < 0 and abs(wp - abs(diff)) <= max(0.05, abs(diff) * 0.02):
                r["wyplata"] = float_to_amount(wp)
                r["wplata"] = ""
                wp, wy = None, wp

        if bal is not None:
            prev_bal = bal

    # 4) usuwanie ewidentnych śmieci / duplikatów
    cleaned: List[Dict[str, str]] = []
    seen = set()
    for r in rows:
        if not row_is_meaningful(r):
            continue
        key = (
            r.get("data", ""),
            r.get("nr", ""),
            r.get("opis", ""),
            r.get("kategoria", ""),
            r.get("proc", ""),
            r.get("vat", ""),
            r.get("wplata", ""),
            r.get("wyplata", ""),
            r.get("bilans", ""),
        )
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(r)

    rows[:] = cleaned


def write_excel(rows: List[Dict[str, str]], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekordy"

    headers = [
        "Data",
        "Nr",
        "Opis",
        "Kategoria",
        "%",
        "VAT",
        "Wpłata",
        "Wypłata",
        "Bilans",
        "Źródło pliku",
        "Wiersz OCR",
    ]
    ws.append(headers)

    for r in rows:
        ws.append([
            r.get("data", ""),
            r.get("nr", ""),
            r.get("opis", ""),
            r.get("kategoria", ""),
            r.get("proc", ""),
            r.get("vat", ""),
            r.get("wplata", ""),
            r.get("wyplata", ""),
            r.get("bilans", ""),
            r.get("source_file", ""),
            r.get("source_row", ""),
        ])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 12, "B": 10, "C": 35, "D": 18, "E": 8, "F": 12,
        "G": 14, "H": 14, "I": 14, "J": 24, "K": 12
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(out_path)


def main() -> None:
    ensure_tesseract_available()

    parser = argparse.ArgumentParser(description="Konwersja skanów cashbook PNG -> Excel.")
    parser.add_argument("input_zip", type=Path, help="ZIP z PNG.")
    parser.add_argument("output_xlsx", type=Path, help="Plik wynikowy XLSX.")
    parser.add_argument("--workers", type=int, default=max(1, (os.cpu_count() or 2) - 1),
                        help="Liczba procesów OCR.")
    args = parser.parse_args()

    if not args.input_zip.exists():
        raise FileNotFoundError(args.input_zip)

    with tempfile.TemporaryDirectory() as td:
        temp_dir = Path(td)
        images = extract_zip(args.input_zip, temp_dir)
        images = [p for p in images if "niepelne" not in p.name.lower()]

        if not images:
            raise RuntimeError("Nie znaleziono plików PNG w archiwum.")

        all_rows: List[Dict[str, str]] = []

        # zostawiam OCR sekwencyjny jako domyślnie najstabilniejszy wariant
        if args.workers <= 1:
            for idx, img in enumerate(images, start=1):
                try:
                    rows = ocr_page(img)
                    all_rows.extend(rows)
                    print(f"[{idx}/{len(images)}] OK  {img.name}: {len(rows)} wierszy")
                except Exception as e:
                    print(f"[{idx}/{len(images)}] BŁĄD {img.name}: {e}")
        else:
            with cf.ThreadPoolExecutor(max_workers=args.workers) as ex:
                futures = {ex.submit(ocr_page, img): img for img in images}
                for idx, fut in enumerate(cf.as_completed(futures), start=1):
                    img = futures[fut]
                    try:
                        rows = fut.result()
                        all_rows.extend(rows)
                        print(f"[{idx}/{len(images)}] OK  {img.name}: {len(rows)} wierszy")
                    except Exception as e:
                        print(f"[{idx}/{len(images)}] BŁĄD {img.name}: {e}")

        def sort_key(r):
            return (natural_key(Path(r.get("source_file", ""))), int(r.get("source_row", "0") or 0))

        all_rows.sort(key=sort_key)

        fix_row_fields(all_rows)

        # po czyszczeniu jeszcze raz usuń duplikaty
        all_rows = merge_very_similar_rows(all_rows)

        write_excel(all_rows, args.output_xlsx)
        print(f"Zapisano: {args.output_xlsx}  |  wierszy: {len(all_rows)}")


if __name__ == "__main__":
    main()

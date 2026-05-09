from pathlib import Path

import pandas as pd


COLUMN_HINTS = {
    "item_number": ["no", "nomor", "item", "kode"],
    "work_title": ["pekerjaan", "judul", "uraian pekerjaan"],
    "description": ["uraian", "deskripsi", "keterangan", "spesifikasi"],
    "material_service_name": ["material", "barang", "jasa", "nama"],
    "volume": ["volume", "vol", "qty", "kuantitas"],
    "unit": ["satuan", "unit", "uom"],
    "unit_price": ["harga satuan", "harga", "price", "harsat"],
    "total_price": ["jumlah", "total", "subtotal", "nilai"],
    "notes": ["catatan", "notes", "remark"],
}


def load_excel_or_csv(file_path):
    path = Path(file_path)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        try:
            df = pd.read_csv(path)
        except UnicodeDecodeError:
            df = pd.read_csv(path, encoding="latin-1")
        return {"dataframe": df, "sheet": "CSV", "warning": ""}
    sheets = pd.read_excel(path, sheet_name=None)
    first_name = next(iter(sheets))
    return {"dataframe": sheets[first_name], "sheet": first_name, "warning": f"Sheet aktif: {first_name}"}


def load_rab_excel_items(file_path):
    """Best-effort extractor for Indonesian RAB workbooks with title rows and item sheets."""
    path = Path(file_path)
    if path.suffix.lower() not in [".xlsx", ".xls"]:
        return None
    sheets = pd.read_excel(path, sheet_name=None, header=None)
    title = _extract_title(sheets) or path.stem
    rows = []
    for sheet_name, raw in sheets.items():
        sheet_upper = str(sheet_name).upper()
        if "REALISASI" in sheet_upper:
            rows.extend(_extract_realisasi_rows(raw, sheet_name, title))
        else:
            rows.extend(_extract_rab_rows(raw, sheet_name, title))
    if not rows:
        return None
    return pd.DataFrame(rows)


def _extract_title(sheets):
    for raw in sheets.values():
        for _, row in raw.iterrows():
            values = ["" if pd.isna(v) else str(v).strip() for v in row.tolist()]
            for idx, value in enumerate(values):
                if value.lower() in ("nama pekerjaan", "judul pekerjaan", "pekerjaan"):
                    for candidate in values[idx + 1 :]:
                        candidate = candidate.strip()
                        if candidate and candidate != ":":
                            return candidate[1:].strip() if candidate.startswith(":") else candidate
    return ""


def _extract_rab_rows(raw, sheet_name, title):
    rows = []
    current_section = ""
    for idx, row in raw.iterrows():
        no = _cell(row, 1)
        desc = _cell(row, 2)
        if desc and not _is_number(no) and not _looks_like_header(desc):
            current_section = desc
        if not (_is_number(no) and desc and not _is_number(desc) and not _looks_like_header(desc)):
            continue
        total = _first_non_empty([_cell(row, 11), _cell(row, 10), _cell(row, 9)])
        rows.append(
            {
                "row_id": str(len(rows) + 1),
                "judul_rab": title,
                "item_per_rab": desc,
                "section": current_section,
                "sheet": sheet_name,
                "volume": _cell(row, 7),
                "unit": _cell(row, 6),
                "unit_price": _cell(row, 8),
                "total_price": total,
                "notes": "",
                "review_text": " | ".join(x for x in [title, current_section, desc] if x),
            }
        )
    return rows


def _extract_realisasi_rows(raw, sheet_name, title):
    rows = []
    for _, row in raw.iterrows():
        item = _cell(row, 0)
        if not item or item.upper() == "REALISASI":
            continue
        total = _first_non_empty([_cell(row, 1), _cell(row, 7)])
        if not item and not total:
            continue
        rows.append(
            {
                "row_id": str(len(rows) + 1),
                "judul_rab": title,
                "item_per_rab": item,
                "section": "Realisasi",
                "sheet": sheet_name,
                "volume": "",
                "unit": "",
                "unit_price": "",
                "total_price": total,
                "notes": _cell(row, 2),
                "review_text": " | ".join(x for x in [title, "Realisasi", item, _cell(row, 2)] if x),
            }
        )
    return rows


def _cell(row, pos):
    if pos >= len(row):
        return ""
    value = row.iloc[pos]
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _first_non_empty(values):
    for value in values:
        if value not in ("", None):
            return value
    return ""


def _is_number(value):
    try:
        float(str(value).strip())
        return True
    except Exception:
        return False


def _looks_like_header(text):
    lower = str(text).lower()
    return any(token in lower for token in ["nama barang", "material", "jasa", "harga", "jumlah"])


def detect_columns(df):
    detected = {}
    normalized = {str(c).lower().strip(): c for c in df.columns}
    for target, hints in COLUMN_HINTS.items():
        for lc, original in normalized.items():
            if any(h in lc for h in hints):
                detected[target] = original
                break
    return detected


def combine_selected_text_columns(df, text_columns):
    if not text_columns:
        raise ValueError("Pilih minimal satu kolom teks untuk review.")
    existing = [c for c in text_columns if c in df.columns]
    combined = df[existing].fillna("").astype(str).agg(" | ".join, axis=1)
    return combined


def normalize_dataframe(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

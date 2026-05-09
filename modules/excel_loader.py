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


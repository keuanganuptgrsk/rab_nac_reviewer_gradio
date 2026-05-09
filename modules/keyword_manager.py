import pandas as pd

from . import db


def import_keywords_from_excel(file_path):
    frame = pd.read_excel(file_path).fillna("")
    required = {"category", "keyword"}
    missing = required - set(frame.columns)
    if missing:
        raise ValueError(f"Kolom wajib hilang: {', '.join(sorted(missing))}")
    added = 0
    for _, row in frame.iterrows():
        if not str(row.get("keyword", "")).strip():
            continue
        keyword_id = db.add_keyword(
            row.get("category", "Umum"),
            row.get("keyword", ""),
            row.get("description", ""),
            row.get("reference", ""),
            row.get("severity", "medium") or "medium",
            row.get("status", "active") or "active",
            row.get("notes", ""),
            "excel_import",
        )
        for syn in str(row.get("synonyms", "")).split(";"):
            syn = syn.strip()
            if syn:
                db.add_synonym(keyword_id, syn)
        added += 1
    return added


def export_keyword_database(path):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        pd.DataFrame(db.get_keywords(False)).to_excel(writer, sheet_name="NAC Keywords", index=False)
        pd.DataFrame(db.get_synonyms(False)).to_excel(writer, sheet_name="Synonyms", index=False)
        pd.DataFrame(db.get_allowable(False)).to_excel(writer, sheet_name="Allowable", index=False)
        pd.DataFrame(db.get_exceptions(False)).to_excel(writer, sheet_name="Exceptions", index=False)
    return path


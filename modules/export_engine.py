from datetime import datetime
from pathlib import Path

import pandas as pd

from . import db


EXPORT_DIR = Path(__file__).resolve().parents[1] / "exports"


def export_review_excel(results):
    EXPORT_DIR.mkdir(exist_ok=True)
    path = EXPORT_DIR / f"rab_nac_review_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.xlsx"
    findings = pd.DataFrame(results)
    if findings.empty:
        findings = pd.DataFrame(columns=[
            "row_id", "source_file", "page_or_sheet", "original_text", "normalized_text", "item_description",
            "volume", "unit", "unit_price", "total_price", "matched_keyword", "matched_category", "match_type",
            "fuzzy_score", "semantic_score", "allowable_score", "final_confidence", "confidence_label",
            "explanation", "recommended_action", "redaction_suggestion", "suggested_synonym_candidate",
            "suggested_synonym_for_keyword", "synonym_suggestion_confidence", "synonym_suggestion_reason",
            "user_feedback", "reviewer_notes",
        ])
    summary = _summary(findings)
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        findings.to_excel(writer, sheet_name="Findings", index=False)
        findings[["row_id", "redaction_suggestion", "recommended_action"]].to_excel(writer, sheet_name="Suggestions", index=False)
        pd.DataFrame(db.get_feedback()).to_excel(writer, sheet_name="Feedback Log", index=False)
        findings[["row_id", "matched_keyword", "matched_category", "match_type", "fuzzy_score", "semantic_score"]].to_excel(writer, sheet_name="Keyword Matches", index=False)
        pd.DataFrame(db.get_keywords(False)).to_excel(writer, sheet_name="NAC Keyword Database Snapshot", index=False)
    return str(path)


def _summary(findings):
    rows = [
        {"metric": "Disclaimer", "value": "Hasil deteksi adalah bantuan awal untuk review internal. Keputusan final tetap harus divalidasi oleh reviewer yang memahami PMK, kebijakan internal, dan konteks pekerjaan."},
        {"metric": "total_items_reviewed", "value": len(findings)},
    ]
    if "confidence_label" in findings:
        for label, count in findings["confidence_label"].value_counts().items():
            rows.append({"metric": f"count_{label}", "value": int(count)})
    high = findings[findings.get("confidence_label", pd.Series(dtype=str)).isin(["Tinggi", "Sangat tinggi"])]
    rows.append({"metric": "count_potential_nac_high_very_high", "value": len(high)})
    if "total_price" in findings:
        values = pd.to_numeric(findings["total_price"], errors="coerce")
        rows.append({"metric": "total_value_detected_rows", "value": float(values.sum(skipna=True) or 0)})
        for label, group in findings.assign(_value=values).groupby("confidence_label", dropna=False):
            rows.append({"metric": f"total_value_{label}", "value": float(group["_value"].sum(skipna=True) or 0)})
    return pd.DataFrame(rows)


def export_feedback_logs():
    EXPORT_DIR.mkdir(exist_ok=True)
    path = EXPORT_DIR / f"feedback_log_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.xlsx"
    pd.DataFrame(db.get_feedback()).to_excel(path, index=False)
    return str(path)

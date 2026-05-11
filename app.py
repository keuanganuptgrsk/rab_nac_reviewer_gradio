import os
import html
from pathlib import Path

import gradio as gr
import pandas as pd

from modules import db
from modules.excel_loader import combine_selected_text_columns, detect_columns, load_excel_or_csv, load_rab_excel_items, normalize_dataframe
from modules.export_engine import (
    export_all_materials_excel,
    export_all_materials_pdf,
    export_feedback_logs,
    export_potential_nac_pdf,
    export_review_excel,
)
from modules.feedback_engine import learning_summary
from modules.keyword_manager import export_keyword_database, import_keywords_from_excel
from modules.nac_detector import detect_items
from modules.ocr_engine import extract_text_from_image, extract_text_from_pdf_scan
from modules.pdf_loader import extract_text_from_pdf
from modules.version import APP_RELEASE_NOTES, APP_RELEASE_TITLE, APP_VERSION, version_banner


DISCLAIMER = (
    "Hasil deteksi adalah bantuan awal untuk review internal. Keputusan final tetap harus divalidasi oleh reviewer "
    "yang memahami PMK, kebijakan internal, dan konteks pekerjaan."
)
BASE_DIR = Path(__file__).resolve().parent
db.init_db()

FLUENTLY_THEME_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600;700;800&display=swap');

:root {
    --rab-bg: #E8EDF2;
    --rab-panel: #FFFFFF;
    --rab-ink: #2C3947;
    --rab-muted: #547A95;
    --rab-accent: #C2A56D;
    --rab-line: #d5dee7;
    --rab-soft-blue: #edf4f8;
    --rab-soft-gold: #fbf5e8;
    --rab-soft-green: #eef7ef;
    --rab-soft-rose: #fff1f2;
}

html,
body,
.gradio-container {
    background: var(--rab-bg) !important;
    color: var(--rab-ink) !important;
    font-family: "Geist", ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
}

.gradio-container {
    width: min(100%, 1440px) !important;
    max-width: 1440px !important;
    margin: 0 auto !important;
    padding: 28px 28px 48px !important;
}

#app-hero {
    padding: 8px 0 22px !important;
    margin: 0 !important;
    background: transparent !important;
    border: 0 !important;
    box-shadow: none !important;
    overflow: visible !important;
}

#app-hero .version-pill {
    display: block !important;
    margin: 0 0 12px !important;
    padding: 0 !important;
    color: var(--rab-muted) !important;
    background: transparent !important;
    border: 0 !important;
    font-size: 14px !important;
    line-height: 1.4 !important;
    font-weight: 600 !important;
}

#app-hero h1 {
    margin: 0 !important;
    color: var(--rab-ink) !important;
    font-size: clamp(30px, 4vw, 52px) !important;
    line-height: 1.12 !important;
    font-weight: 800 !important;
    letter-spacing: -0.02em !important;
    overflow: visible !important;
}

.tabs {
    background: var(--rab-panel) !important;
    border: 1px solid var(--rab-line) !important;
    border-radius: 18px !important;
    box-shadow: 0 18px 48px rgba(44, 57, 71, 0.08) !important;
    overflow: visible !important;
}

div[role="tablist"] {
    display: flex !important;
    flex-wrap: wrap !important;
    gap: 8px !important;
    align-items: center !important;
    padding: 12px !important;
    min-height: 0 !important;
    height: auto !important;
    background: #f6f9fb !important;
    border-bottom: 1px solid var(--rab-line) !important;
    border-radius: 18px 18px 0 0 !important;
    overflow: visible !important;
}

button[role="tab"] {
    min-height: 42px !important;
    height: auto !important;
    padding: 10px 16px !important;
    border-radius: 12px !important;
    color: var(--rab-muted) !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    font-weight: 700 !important;
    white-space: nowrap !important;
    line-height: 1.2 !important;
}

button[role="tab"][aria-selected="true"] {
    color: #ffffff !important;
    background: var(--rab-ink) !important;
    border-color: var(--rab-ink) !important;
}

div[role="tabpanel"] {
    padding: 28px !important;
    background: var(--rab-panel) !important;
    overflow: visible !important;
}

h1, h2, h3,
.prose h1, .prose h2, .prose h3 {
    color: var(--rab-ink) !important;
    letter-spacing: -0.015em !important;
}

.gradio-container label,
.gradio-container .prose,
.gradio-container p {
    color: var(--rab-ink) !important;
}

.gradio-container button:not([role="tab"]),
.gradio-container .download-button {
    border-radius: 12px !important;
    font-weight: 700 !important;
    min-height: 44px !important;
    transition: transform 120ms ease, background 120ms ease, border-color 120ms ease !important;
}

.gradio-container button:not([role="tab"]):active {
    transform: translateY(1px) !important;
}

.primary > button,
button[variant="primary"] {
    background: var(--rab-muted) !important;
    color: #ffffff !important;
    border: 1px solid var(--rab-muted) !important;
}

input,
textarea,
select {
    border-radius: 12px !important;
    border-color: var(--rab-line) !important;
    color: var(--rab-ink) !important;
}

#rab-upload-file,
#keyword-upload-file {
    background: #ffffff !important;
    border: 1px dashed #a9bac9 !important;
    border-radius: 16px !important;
    overflow: visible !important;
}

#rab-upload-file {
    min-height: 170px !important;
}

.review-output-panel {
    margin-top: 24px !important;
    padding: 0 !important;
    background: transparent !important;
    border: 0 !important;
    overflow: visible !important;
}

.review-output-panel > div,
.review-output-panel .form,
.review-output-panel [data-testid="html"],
.review-output-panel [data-testid="dataframe"] {
    overflow: visible !important;
    max-width: 100% !important;
}

.findings-panel {
    display: grid;
    gap: 14px;
    margin: 18px 0 22px;
}

.findings-toolbar {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 12px;
}

.findings-kpi {
    padding: 16px;
    border: 1px solid var(--rab-line);
    border-radius: 16px;
    background: #ffffff;
}

.findings-kpi strong {
    display: block;
    color: var(--rab-ink);
    font-size: 28px;
    line-height: 1;
    margin-bottom: 8px;
}

.findings-kpi span {
    color: var(--rab-muted);
    font-size: 13px;
    font-weight: 650;
}

.finding-card {
    display: grid;
    grid-template-columns: 72px minmax(220px, 1fr) 140px 140px;
    gap: 16px;
    align-items: center;
    padding: 16px;
    border: 1px solid var(--rab-line);
    border-radius: 16px;
    background: #ffffff;
    box-shadow: 0 8px 24px rgba(44, 57, 71, 0.06);
}

.finding-row {
    width: 48px;
    height: 48px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 14px;
    background: var(--rab-soft-blue);
    color: var(--rab-ink);
    font-weight: 800;
}

.finding-item {
    color: var(--rab-ink);
    font-size: 17px;
    line-height: 1.35;
    font-weight: 800;
}

.finding-meta,
.redaction-copy,
.keyword-desc,
.learning-copy,
.simple-note {
    color: var(--rab-muted);
}

.confidence-score,
.redaction-score {
    color: var(--rab-ink);
    font-weight: 800;
    font-size: 24px;
}

.confidence-bar {
    height: 8px;
    margin-top: 8px;
    border-radius: 999px;
    background: #dbe4ec;
    overflow: hidden;
}

.confidence-fill {
    height: 100%;
    border-radius: 999px;
}

.confidence-pill,
.materials-level,
.severity-badge,
.alias-chip,
.keyword-chip {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 999px;
    padding: 6px 10px;
    font-size: 12px;
    font-weight: 750;
    border: 1px solid rgba(44, 57, 71, 0.12);
    background: var(--rab-soft-gold);
    color: var(--rab-ink);
}

.level-sedang,
.fill-sedang {
    background: var(--rab-accent) !important;
    color: #ffffff !important;
}

.level-tinggi,
.fill-tinggi {
    background: #9c7d3d !important;
    color: #ffffff !important;
}

.level-sangat-tinggi,
.fill-sangat-tinggi {
    background: #8f3d46 !important;
    color: #ffffff !important;
}

.keyword-workspace,
.learning-grid {
    display: grid;
    gap: 16px;
}

.keyword-search-panel,
.learning-card,
.redaction-search,
.redaction-result,
.versioning-panel,
.simple-note,
.empty-findings {
    border: 1px solid var(--rab-line);
    border-radius: 16px;
    background: #ffffff;
    padding: 16px;
    box-shadow: none;
}

.keyword-search-head {
    display: flex;
    justify-content: space-between;
    gap: 16px;
    align-items: center;
    margin-bottom: 12px;
}

.keyword-search-title,
.learning-title,
.redaction-title {
    color: var(--rab-ink);
    font-weight: 800;
}

.keyword-chip-row,
.alias-row {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
}

.keyword-card-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
    gap: 14px;
}

.keyword-card {
    position: relative;
    min-height: 230px;
    padding: 18px 18px 58px;
    border: 1px solid var(--rab-line);
    border-radius: 18px;
    background: #ffffff;
}

.keyword-card:nth-child(3n + 1) { background: var(--rab-soft-gold); }
.keyword-card:nth-child(3n + 2) { background: var(--rab-soft-blue); }
.keyword-card:nth-child(3n + 3) { background: var(--rab-soft-green); }

.keyword-card-top {
    display: flex;
    justify-content: space-between;
    gap: 12px;
    align-items: flex-start;
}

.keyword-name {
    color: var(--rab-ink);
    font-size: 22px;
    line-height: 1.2;
    font-weight: 800;
}

.keyword-category,
.alias-label {
    color: var(--rab-ink);
    font-weight: 700;
    margin-top: 8px;
}

.keyword-delete-btn {
    position: absolute !important;
    right: 16px;
    bottom: 16px;
    min-height: 32px !important;
    height: 32px !important;
    padding: 6px 12px !important;
    border-radius: 10px !important;
    background: var(--rab-ink) !important;
    color: #ffffff !important;
    border: 0 !important;
    font-size: 12px !important;
}

.hidden-delete-control {
    position: absolute !important;
    width: 1px !important;
    height: 1px !important;
    opacity: 0 !important;
    pointer-events: none !important;
}

.learning-grid {
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
}

.learning-number {
    color: var(--rab-ink);
    font-size: 32px;
    line-height: 1;
    font-weight: 850;
}

.learning-label {
    color: var(--rab-muted);
    margin-top: 8px;
    font-weight: 650;
}

.learning-row {
    display: flex;
    justify-content: space-between;
    gap: 12px;
    padding: 10px 0;
    border-top: 1px solid var(--rab-line);
}

.learning-count {
    color: #ffffff;
    background: var(--rab-muted);
    border-radius: 999px;
    padding: 3px 9px;
    font-weight: 800;
}

.redaction-result {
    display: grid;
    grid-template-columns: 150px 1fr;
    gap: 16px;
    align-items: center;
    margin-top: 16px;
}

.dataframe,
[data-testid="dataframe"] {
    color: var(--rab-ink) !important;
}

@media (max-width: 760px) {
    .gradio-container {
        padding: 18px 12px 36px !important;
    }

    div[role="tabpanel"] {
        padding: 18px !important;
    }

    .findings-toolbar,
    .finding-card,
    .redaction-result {
        grid-template-columns: 1fr;
    }
}
"""


def _file_path(file_obj):
    return file_obj.name if hasattr(file_obj, "name") else str(file_obj)


def handle_upload(file_obj):
    if file_obj is None:
        empty = gr.update(choices=[], value=None)
        return pd.DataFrame(), empty, empty, empty, empty, empty, "Upload file terlebih dahulu.", {}
    path = Path(_file_path(file_obj))
    suffix = path.suffix.lower()
    if suffix in [".xlsx", ".xls", ".csv"]:
        rab_items = load_rab_excel_items(path) if suffix in [".xlsx", ".xls"] else None
        if rab_items is not None and not rab_items.empty:
            frame = normalize_dataframe(rab_items)
            loaded = {"sheet": "RAB/REALISASI", "warning": "Format RAB terdeteksi. Review otomatis menggunakan Judul dan Item per RAB."}
        else:
            loaded = load_excel_or_csv(path)
            frame = normalize_dataframe(loaded["dataframe"])
        detected = detect_columns(frame)
        cols = list(frame.columns)
        text_defaults = ["review_text"] if "review_text" in cols else [c for k, c in detected.items() if k in ("work_title", "description", "material_service_name", "notes")]
        text_defaults = text_defaults or cols[:1]
        preview = frame.head(30)
        state = {"kind": "table", "path": str(path), "sheet": loaded["sheet"], "data": frame.to_dict("records"), "columns": cols, "detected": detected}
        msg = f"{loaded.get('warning','')} Kolom terdeteksi: {detected}"
        return (
            preview,
            gr.update(choices=cols, value=text_defaults),
            gr.update(choices=cols, value=detected.get("volume")),
            gr.update(choices=cols, value=detected.get("unit")),
            gr.update(choices=cols, value=detected.get("unit_price")),
            gr.update(choices=cols, value=detected.get("total_price")),
            msg,
            state,
        )
    if suffix == ".pdf":
        chunks, warning, scanned = extract_text_from_pdf(path)
        settings = db.get_settings()
        if scanned:
            ocr_text, ocr_note = extract_text_from_pdf_scan(path, settings.get("ocr_mode", "auto"))
            if ocr_text.strip():
                chunks = [{"page_or_sheet": "OCR PDF", "text": ocr_text}]
                warning = ocr_note
            else:
                warning = f"{warning} {ocr_note}"
        frame = pd.DataFrame(chunks)
        state = {"kind": "chunks", "path": str(path), "data": chunks, "source_quality": "ocr" if scanned else "digital_pdf"}
        return frame, gr.update(choices=["text"], value=["text"]), gr.update(choices=[], value=None), gr.update(choices=[], value=None), gr.update(choices=[], value=None), gr.update(choices=[], value=None), warning, state
    if suffix in [".png", ".jpg", ".jpeg"]:
        text, note = extract_text_from_image(path, db.get_settings().get("ocr_mode", "auto"))
        chunks = [{"page_or_sheet": "Image OCR", "text": text}] if text else []
        return pd.DataFrame(chunks), gr.update(choices=["text"], value=["text"]), gr.update(choices=[], value=None), gr.update(choices=[], value=None), gr.update(choices=[], value=None), gr.update(choices=[], value=None), note, {"kind": "chunks", "path": str(path), "data": chunks, "source_quality": "ocr"}
    empty = gr.update(choices=[], value=None)
    return pd.DataFrame(), empty, empty, empty, empty, empty, "Format file tidak didukung.", {}


def build_items(upload_state, text_columns, volume_col, unit_col, unit_price_col, total_price_col):
    if not upload_state:
        return [], "Upload file dahulu."
    path = Path(upload_state.get("path", ""))
    if upload_state.get("kind") == "table":
        frame = pd.DataFrame(upload_state["data"])
        combined = combine_selected_text_columns(frame, text_columns)
        items = []
        for idx, text in combined.items():
            row = frame.loc[idx]
            items.append({
                "row_id": str(row.get("row_id", idx + 1)),
                "source_file": path.name,
                "page_or_sheet": row.get("sheet", upload_state.get("sheet", "")),
                "original_text": text,
                "item_description": text,
                "judul_rab": row.get("judul_rab", ""),
                "item_per_rab": row.get("item_per_rab", text),
                "section": row.get("section", ""),
                "volume": row.get(volume_col, "") if volume_col else "",
                "unit": row.get(unit_col, "") if unit_col else "",
                "unit_price": row.get(unit_price_col, "") if unit_price_col else "",
                "total_price": row.get(total_price_col, "") if total_price_col else "",
                "source_quality": "table",
            })
        return items, f"{len(items)} baris siap direview."
    items = []
    for i, chunk in enumerate(upload_state.get("data", []), start=1):
        text = chunk.get("text", "")
        for part_no, part in enumerate(_chunk_text(text), start=1):
            items.append({
                "row_id": f"{i}.{part_no}",
                "source_file": path.name,
                "page_or_sheet": chunk.get("page_or_sheet", ""),
                "original_text": part,
                "item_description": part,
                "source_quality": upload_state.get("source_quality", "text"),
            })
    return items, f"{len(items)} chunk teks siap direview."


def _chunk_text(text, max_len=900):
    lines = [ln.strip() for ln in str(text).splitlines() if ln.strip()]
    chunks, current = [], ""
    for line in lines:
        if len(current) + len(line) > max_len and current:
            chunks.append(current)
            current = line
        else:
            current = f"{current} {line}".strip()
    if current:
        chunks.append(current)
    return chunks or ([text] if text else [])


def run_review(upload_state, text_columns, volume_col, unit_col, unit_price_col, total_price_col, sort_by, sort_order):
    items, msg = build_items(upload_state, text_columns, volume_col, unit_col, unit_price_col, total_price_col)
    if not items:
        return (
            gr.update(visible=True),
            gr.update(value=f"Review belum dapat dijalankan. {msg}", visible=True),
            gr.update(value="", visible=False),
            gr.update(value=all_materials_dataframe([]), visible=False),
            [],
            gr.update(visible=False),
            gr.update(visible=False),
        )
    try:
        results = detect_items(items, db.get_settings())
    except Exception as exc:
        return (
            gr.update(visible=True),
            gr.update(value=f"Review gagal diproses: {exc}", visible=True),
            gr.update(value="", visible=False),
            gr.update(value=all_materials_dataframe([]), visible=False),
            [],
            gr.update(visible=False),
            gr.update(visible=False),
        )
    return (
        gr.update(visible=True),
        gr.update(value=f"Review selesai. {msg} Output pertama menampilkan confidence Sedang sampai Sangat tinggi. Output kedua menampilkan seluruh item RAB. {DISCLAIMER}", visible=True),
        gr.update(value=render_findings_cards(results, sort_by, sort_order), visible=True),
        gr.update(value=all_materials_dataframe(results), visible=True),
        results,
        gr.update(visible=True),
        gr.update(visible=True),
    )


def auto_run_review(upload_state):
    if not upload_state:
        return "", "", [], "Upload file dahulu."
    detected = upload_state.get("detected", {})
    columns = upload_state.get("columns", [])
    text_columns = ["review_text"] if "review_text" in columns else [
        col for key, col in detected.items() if key in ("work_title", "description", "material_service_name", "notes")
    ]
    text_columns = text_columns or columns[:1]
    review_outputs = run_review(
        upload_state,
        text_columns,
        detected.get("volume") or ("volume" if "volume" in columns else None),
        detected.get("unit") or ("unit" if "unit" in columns else None),
        detected.get("unit_price") or ("unit_price" if "unit_price" in columns else None),
        detected.get("total_price") or ("total_price" if "total_price" in columns else None),
        "Confidence",
        "Tinggi ke rendah",
    )
    return review_outputs[0], review_outputs[1], review_outputs[2], review_outputs[3]


def review_summary_dataframe(results):
    frame = pd.DataFrame(results or [])
    if frame.empty:
        return frame
    frame = frame[frame["confidence_label"].isin(["Sedang", "Tinggi", "Sangat tinggi"])].copy()
    columns = [
        "row_id",
        "source_file",
        "page_or_sheet",
        "judul_rab",
        "section",
        "item_per_rab",
        "matched_category",
        "matched_keyword",
        "match_type",
        "final_confidence",
        "confidence_label",
        "explanation",
        "redaction_suggestion",
    ]
    for col in columns:
        if col not in frame.columns:
            frame[col] = ""
    return frame[columns].rename(
        columns={
            "row_id": "Row",
            "source_file": "File",
            "page_or_sheet": "Sheet",
            "judul_rab": "Judul",
            "section": "Bagian",
            "item_per_rab": "Item per RAB",
            "matched_category": "Kategori",
            "matched_keyword": "Keyword",
            "match_type": "Tipe Deteksi",
            "final_confidence": "Confidence",
            "confidence_label": "Confidence Level",
            "explanation": "Alasan Deteksi",
            "redaction_suggestion": "Sugesti Perubahan Redaksi",
        }
    )


def all_materials_dataframe(results):
    frame = pd.DataFrame(results or [])
    columns = ["row_id", "item_per_rab", "matched_category", "final_confidence", "confidence_label"]
    labels = {
        "row_id": "Row",
        "item_per_rab": "Item RAB",
        "matched_category": "Kategori NAC",
        "final_confidence": "Confidence %",
        "confidence_label": "Confidence Level",
    }
    if frame.empty:
        return pd.DataFrame(columns=list(labels.values()))
    for col in columns:
        if col not in frame.columns:
            frame[col] = ""
    frame = frame[columns].copy()
    frame["matched_category"] = frame["matched_category"].replace("", "-").fillna("-")
    frame["final_confidence"] = pd.to_numeric(frame["final_confidence"], errors="coerce").fillna(0).round(2)
    frame["_row_sort"] = pd.to_numeric(frame["row_id"], errors="coerce")
    frame = frame.sort_values("_row_sort", na_position="last").drop(columns=["_row_sort"])
    return frame.rename(columns=labels)


def render_all_materials_table(results):
    frame = pd.DataFrame(results or [])
    if frame.empty:
        return _empty_findings("Belum ada tabel material. Upload RAB lalu tekan Run NAC Review.")
    columns = ["row_id", "item_per_rab", "matched_category", "final_confidence", "confidence_label"]
    for col in columns:
        if col not in frame.columns:
            frame[col] = ""
    frame = frame[columns].copy()
    frame["matched_category"] = frame["matched_category"].replace("", "-").fillna("-")
    frame["item_per_rab"] = frame["item_per_rab"].replace("", "-").fillna("-")
    frame["confidence_label"] = frame["confidence_label"].replace("", "Sangat rendah").fillna("Sangat rendah")
    frame["final_confidence"] = pd.to_numeric(frame["final_confidence"], errors="coerce").fillna(0).round(2)
    frame["_row_sort"] = frame["row_id"].apply(_row_sort_key)
    frame = frame.sort_values("_row_sort", na_position="last")

    rows = []
    for _, row in frame.iterrows():
        score = _safe_float(row.get("final_confidence", 0))
        level = str(row.get("confidence_label") or "Sangat rendah")
        level_class = _level_class(level)
        rows.append(
            "<tr "
            f"data-row='{html.escape(str(row.get('_row_sort', '999999')))}' "
            f"data-name='{html.escape(str(row.get('item_per_rab') or '').lower())}' "
            f"data-category='{html.escape(str(row.get('matched_category') or '').lower())}' "
            f"data-confidence='{score:.4f}' "
            f"data-level='{html.escape(level.lower())}'>"
            f"<td>{html.escape(str(row.get('row_id') or '-'))}</td>"
            f"<td>{html.escape(str(row.get('item_per_rab') or '-'))}</td>"
            f"<td>{html.escape(str(row.get('matched_category') or '-'))}</td>"
            f"<td><span class='materials-confidence'>{score:.2f}%</span></td>"
            f"<td><span class='materials-level level-{level_class}'>{html.escape(level)}</span></td>"
            "</tr>"
        )
    return (
        "<section class='materials-table-panel'>"
        "<div class='materials-table-head'>"
        "<div>"
        "<div class='materials-table-title'>Tabel Seluruh Item RAB</div>"
        f"<div class='materials-table-copy'>{len(frame)} item ditampilkan. Klik header kolom untuk sort langsung di UI.</div>"
        "</div>"
        "</div>"
        "<div class='materials-table-wrap'>"
        "<table class='materials-table' data-sort-dir='asc'>"
        "<thead><tr>"
        "<th><button class='materials-sort-header' data-sort='row' type='button'>Row</button></th>"
        "<th><button class='materials-sort-header' data-sort='name' type='button'>Item RAB</button></th>"
        "<th><button class='materials-sort-header' data-sort='category' type='button'>Kategori NAC</button></th>"
        "<th><button class='materials-sort-header' data-sort='confidence' type='button'>Confidence</button></th>"
        "<th><button class='materials-sort-header' data-sort='level' type='button'>Confidence Level</button></th>"
        "</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
        "</div>"
        "</section>"
    )


def render_sorted_findings(results, sort_by, sort_order):
    return gr.update(value=render_findings_cards(results, sort_by, sort_order), visible=bool(results))


def reset_analysis_visibility():
    return (
        gr.update(visible=False),
        gr.update(value="", visible=False),
        gr.update(value="", visible=False),
        gr.update(value=all_materials_dataframe([]), visible=False),
        gr.update(value="", visible=False),
        gr.update(visible=False),
        gr.update(visible=False),
    )


def render_findings_cards(results, sort_by="Confidence", sort_order="Tinggi ke rendah"):
    frame = pd.DataFrame(results or [])
    if frame.empty:
        return _empty_findings("Belum ada hasil review. Upload RAB lalu tekan Run NAC Review.")
    total_count = len(frame)
    if "final_confidence" not in frame.columns:
        frame["final_confidence"] = 0
    frame["final_confidence"] = pd.to_numeric(frame["final_confidence"], errors="coerce").fillna(0)
    max_confidence = float(frame["final_confidence"].max()) if not frame.empty else 0.0
    frame = frame[frame["confidence_label"].isin(["Sedang", "Tinggi", "Sangat tinggi"])].copy()
    if frame.empty:
        return _empty_findings("Tidak ada item dengan confidence Sedang hingga Sangat tinggi.")
    frame["_row_sort"] = frame["row_id"].apply(_row_sort_key)
    if sort_by == "Row":
        ascending = sort_order == "Rendah ke tinggi"
        frame = frame.sort_values(["_row_sort", "final_confidence"], ascending=[ascending, False])
    else:
        ascending = sort_order == "Rendah ke tinggi"
        frame = frame.sort_values(["final_confidence", "_row_sort"], ascending=[ascending, True])
    cards = [
        "<div class='findings-panel'>",
        "<div class='findings-toolbar'>"
        f"<div class='findings-kpi'><strong>{total_count}</strong><span>Total item RAB terbaca</span></div>"
        f"<div class='findings-kpi'><strong>{len(frame)}</strong><span>Potensi NAC confidence sedang-tinggi</span></div>"
        f"<div class='findings-kpi'><strong>{max_confidence:.0f}%</strong><span>Confidence tertinggi</span></div>"
        "</div>",
    ]
    for _, row in frame.iterrows():
        score = _safe_float(row.get("final_confidence", 0))
        label = str(row.get("confidence_label", "Sedang"))
        level_class = _level_class(label)
        item = html.escape(str(row.get("item_per_rab") or row.get("item_description") or row.get("original_text") or "-"))
        row_id = html.escape(str(row.get("row_id") or "-"))
        keyword = html.escape(str(row.get("matched_keyword") or "-"))
        category = html.escape(str(row.get("matched_category") or "-"))
        cards.append(
            "<div class='finding-card'>"
            f"<div><div class='finding-row'>{row_id}</div></div>"
            f"<div><div class='finding-item'>{item}</div><div class='finding-meta'>{category} | {keyword}</div></div>"
            f"<div><div class='confidence-score'>{score:.0f}%</div>"
            f"<div class='confidence-bar'><div class='confidence-fill fill-{level_class}' style='width:{min(max(score, 0), 100):.0f}%'></div></div></div>"
            f"<div><span class='confidence-pill level-{level_class}'>{html.escape(label)}</span></div>"
            "</div>"
        )
    cards.append("</div>")
    return "".join(cards)


def _empty_findings(message):
    return f"<div class='empty-findings'>{html.escape(message)}</div>"


def _safe_float(value):
    try:
        return float(value)
    except Exception:
        return 0.0


def _level_class(label):
    normalized = str(label).lower().replace(" ", "-")
    if normalized == "sangat-tinggi":
        return "sangat-tinggi"
    if normalized == "tinggi":
        return "tinggi"
    return "sedang"


def _row_sort_key(value):
    try:
        return float(str(value).split(".")[0])
    except Exception:
        return 999999


def filter_results(results, label, category, match_type, keyword, medium_high, manual_only):
    frame = pd.DataFrame(results or [])
    if frame.empty:
        return frame
    if label and label != "Semua":
        frame = frame[frame["confidence_label"] == label]
    if category:
        frame = frame[frame["matched_category"].astype(str).str.contains(category, case=False, na=False)]
    if match_type and match_type != "Semua":
        frame = frame[frame["match_type"] == match_type]
    if keyword:
        frame = frame[frame["matched_keyword"].astype(str).str.contains(keyword, case=False, na=False)]
    if medium_high:
        frame = frame[frame["confidence_label"].isin(["Sedang", "Tinggi", "Sangat tinggi"])]
    if manual_only:
        frame = frame[frame["recommended_action"].astype(str).str.contains("Review Manual", case=False, na=False)]
    return review_summary_dataframe(frame.to_dict("records"))


def save_row_feedback(results, row_id, feedback_type, redaction, notes):
    frame = pd.DataFrame(results or [])
    if frame.empty or not row_id:
        return "Pilih row_id hasil review."
    row = frame[frame["row_id"].astype(str) == str(row_id)]
    if row.empty:
        return "row_id tidak ditemukan."
    rec = row.iloc[0].to_dict()
    db.save_feedback(row_id, rec.get("original_text", ""), rec.get("matched_keyword", ""), feedback_type, redaction, notes)
    return "Feedback tersimpan ke SQLite."


def approve_suggested_synonym(results, row_id, weight):
    frame = pd.DataFrame(results or [])
    if frame.empty or not row_id:
        return "Pilih row_id yang memiliki suggested synonym.", render_keyword_cards("")
    row = frame[frame["row_id"].astype(str) == str(row_id)]
    if row.empty:
        return "row_id tidak ditemukan.", render_keyword_cards("")
    rec = row.iloc[0].to_dict()
    candidate = str(rec.get("suggested_synonym_candidate", "") or "").strip()
    keyword = str(rec.get("suggested_synonym_for_keyword", "") or rec.get("matched_keyword", "") or "").strip()
    if not candidate or not keyword:
        return "Baris ini belum memiliki kandidat sinonim dari model.", render_keyword_cards("")
    keyword_row = db.get_keyword_by_text(keyword)
    if not keyword_row:
        return f"Keyword induk tidak ditemukan: {keyword}", render_keyword_cards("")
    if db.synonym_exists(keyword_row["id"], candidate):
        return "Sinonim sudah ada di database.", render_keyword_cards(keyword)
    db.add_synonym(keyword_row["id"], candidate, float(weight or 0.85), "active")
    db.save_feedback(
        row_id,
        rec.get("original_text", ""),
        keyword,
        "Add as Synonym",
        candidate,
        "Approved model-suggested synonym",
    )
    return f"Sinonim '{candidate}' ditambahkan untuk keyword '{keyword}'.", render_keyword_cards(keyword)


def add_keyword_ui(category, keyword, description, reference, severity, status, notes):
    if not keyword:
        return "Keyword wajib diisi.", refresh_keywords()
    db.add_keyword(category or "Umum", keyword, description, reference, severity, status, notes)
    return "Keyword NAC ditambahkan.", refresh_keywords()


def add_synonym_ui(keyword_id, synonym, weight):
    if not keyword_id or not synonym:
        return "Pilih keyword dan isi sinonim.", refresh_synonyms()
    db.add_synonym(int(keyword_id), synonym, float(weight or 0.9))
    return "Sinonim ditambahkan.", refresh_synonyms()


def update_keyword_status_ui(keyword_id, status):
    if not keyword_id:
        return "Isi keyword id.", refresh_keywords()
    db.update_keyword_status(int(keyword_id), status)
    return "Status keyword diperbarui.", refresh_keywords()


def update_synonym_status_ui(synonym_id, status):
    if not synonym_id:
        return "Isi synonym id.", refresh_synonyms()
    db.update_synonym_status(int(synonym_id), status)
    return "Status synonym diperbarui.", refresh_synonyms()


def update_allowable_status_ui(allowable_id, status):
    if not allowable_id:
        return "Isi allowable id.", refresh_allowable()
    db.update_allowable_status(int(allowable_id), status)
    return "Status allowable diperbarui.", refresh_allowable()


def update_exception_status_ui(exception_id, status):
    if not exception_id:
        return "Isi exception id.", refresh_exceptions()
    db.update_exception_status(int(exception_id), status)
    return "Status exception diperbarui.", refresh_exceptions()


def add_allowable_ui(category, keyword, description):
    if not keyword:
        return "Keyword allowable wajib diisi.", refresh_allowable()
    db.add_allowable(category or "Teknis", keyword, description)
    return "Allowable keyword ditambahkan.", refresh_allowable()


def add_exception_ui(keyword_id, pattern, reason, action, adjustment):
    if not pattern:
        return "Pattern exception wajib diisi.", refresh_exceptions()
    db.add_exception(int(keyword_id) if keyword_id else None, pattern, reason, action, float(adjustment or 25))
    return "Exception ditambahkan.", refresh_exceptions()


def refresh_keywords():
    return pd.DataFrame(db.get_keywords(False))


def refresh_synonyms():
    return pd.DataFrame(db.get_synonyms(False))


def refresh_allowable():
    return pd.DataFrame(db.get_allowable(False))


def refresh_exceptions():
    return pd.DataFrame(db.get_exceptions(False))


def _auto_aliases(keyword):
    base = str(keyword or "").strip().lower()
    if not base:
        return []
    variants = []
    if not base.startswith("biaya "):
        variants.append(f"biaya {base}")
    if " " in base:
        variants.append(base.replace("biaya ", ""))
    if "honorarium" in base:
        variants.extend(["honor", "fee"])
    if "konsumsi" in base:
        variants.extend(["makan minum", "jamuan", "snack"])
    if "transport" in base:
        variants.extend(["transportasi", "bantuan transport"])
    cleaned = []
    for item in variants:
        item = item.strip()
        if item and item != base and item not in cleaned:
            cleaned.append(item)
    return cleaned[:5]


def _infer_keyword_metadata(keyword):
    text = str(keyword or "").lower()
    rules = [
        (["konsumsi", "makan", "minum", "snack", "jamuan", "catering"], "Rapat/Jamuan", "high", "Kandidat biaya konsumsi/jamuan; validasi konteks kegiatan dan aturan internal."),
        (["hadiah", "souvenir", "doorprize", "oleh-oleh", "cinderamata"], "Pribadi/Hadiah", "high", "Kandidat biaya hadiah/cinderamata; perlu validasi allowability."),
        (["pegawai", "tunjangan", "cuti", "fasilitas", "seragam"], "Pegawai", "high", "Kandidat biaya terkait pegawai; cek pemisahan komponen allowable/non-allowable."),
        (["denda", "sanksi", "penalti"], "Denda/Sanksi", "high", "Kandidat denda/sanksi; biasanya perlu review khusus."),
        (["honor", "narasumber", "fee", "uang saku", "pulsa"], "Personel/Operasional", "medium", "Kandidat biaya personel/operasional; validasi dasar pembayaran dan output kegiatan."),
        (["transport", "perjalanan", "akomodasi"], "Transportasi/Personel", "medium", "Kandidat biaya perjalanan/transport; pastikan terkait langsung dengan pekerjaan teknis."),
    ]
    for tokens, category, severity, notes in rules:
        if any(token in text for token in tokens):
            return category, severity, notes
    return "Umum", "medium", "Keyword ditambahkan dari UI sederhana; wajib validasi PMK/kebijakan internal."


def add_keyword_simple_ui(keyword):
    keyword = str(keyword or "").strip()
    if not keyword:
        return "Isi nama keyword NAC terlebih dahulu.", render_keyword_cards(""), ""
    existing = db.get_keyword_by_text(keyword)
    if existing:
        return f"Keyword '{keyword}' sudah ada.", render_keyword_cards(keyword), ""
    category, severity, notes = _infer_keyword_metadata(keyword)
    keyword_id = db.add_keyword(
        category,
        keyword,
        notes,
        "USER",
        severity,
        "active",
        "Metadata kategori/confidence dasar dipilih otomatis oleh sistem.",
    )
    aliases = _auto_aliases(keyword)
    for alias in aliases:
        if not db.synonym_exists(keyword_id, alias):
            db.add_synonym(keyword_id, alias, 0.85, "active")
    msg = f"Keyword '{keyword}' ditambahkan."
    if aliases:
        msg += " Sistem menambahkan kandidat sinonim/parafrasa otomatis: " + ", ".join(aliases) + "."
    return msg, render_keyword_cards(keyword), ""


def keyword_delete_choices():
    return [f"{row['id']} | {row['keyword']}" for row in db.get_keywords(False) if row.get("status") == "active"]


def delete_keyword_simple_ui(selection):
    if not selection:
        return "Pilih keyword yang ingin dihapus dari daftar aktif.", render_keyword_cards(""), gr.update(choices=keyword_delete_choices(), value=None)
    keyword_id = str(selection).split("|", 1)[0].strip()
    try:
        db.update_keyword_status(int(keyword_id), "inactive")
    except Exception as exc:
        return f"Gagal menghapus keyword: {exc}", render_keyword_cards(""), gr.update(choices=keyword_delete_choices(), value=None)
    return "Keyword dihapus dari daftar aktif. Data tidak dihapus permanen agar tetap audit-friendly.", render_keyword_cards(""), gr.update(choices=keyword_delete_choices(), value=None)


def delete_keyword_by_id_ui(keyword_id):
    keyword_id = str(keyword_id or "").strip()
    if not keyword_id:
        return "", render_keyword_cards("")
    try:
        db.update_keyword_status(int(keyword_id), "inactive")
    except Exception as exc:
        return f"Gagal menghapus keyword: {exc}", render_keyword_cards("")
    return "Keyword dihapus dari daftar aktif. Data tidak dihapus permanen agar tetap audit-friendly.", render_keyword_cards("")


def import_keywords_simple_ui(file_obj):
    if file_obj is None:
        return "Upload file Excel keyword dahulu.", render_keyword_cards(""), ""
    try:
        count = import_keywords_from_excel(_file_path(file_obj))
    except Exception as exc:
        return f"Import gagal: {exc}", render_keyword_cards(""), ""
    return f"{count} keyword berhasil diimpor dari Excel.", render_keyword_cards(""), ""


def render_keyword_cards(query=""):
    query_l = str(query or "").strip().lower()
    keywords = [row for row in db.get_keywords(False) if row.get("status") == "active"]
    synonyms = db.get_synonyms(False)
    aliases_by_keyword = {}
    for syn in synonyms:
        aliases_by_keyword.setdefault(syn.get("nac_keyword_id"), []).append(syn.get("synonym", ""))

    if query_l:
        filtered = []
        for row in keywords:
            haystack = " ".join(
                [
                    str(row.get("keyword", "")),
                    str(row.get("category", "")),
                    str(row.get("description", "")),
                    " ".join(aliases_by_keyword.get(row.get("id"), [])),
                ]
            ).lower()
            if query_l in haystack:
                filtered.append(row)
        keywords = filtered

    active_count = sum(1 for row in keywords if row.get("status") == "active")
    top_keywords = keywords[:7]
    chip_colors = ["chip-blue", "chip-yellow", "chip-green", "chip-purple", "chip-red"]
    chips = "".join(
        f"<span class='keyword-chip {chip_colors[i % len(chip_colors)]}'>{html.escape(str(row.get('keyword', '')))}</span>"
        for i, row in enumerate(top_keywords)
    )
    cards = []
    for row in keywords:
        severity = str(row.get("severity") or "medium").replace("_", "-")
        aliases = aliases_by_keyword.get(row.get("id"), [])
        alias_html = "".join(f"<span class='alias-chip'>{html.escape(str(alias))}</span>" for alias in aliases[:8])
        if not alias_html:
            alias_html = "<span class='alias-chip'>semantic/fuzzy otomatis</span>"
        desc = html.escape(str(row.get("description") or "Keyword demo/user; wajib divalidasi reviewer."))
        status = html.escape(str(row.get("status") or "active"))
        keyword_id = html.escape(str(row.get("id") or ""))
        cards.append(
            "<div class='keyword-card'>"
            "<div class='keyword-card-top'>"
            "<div>"
            f"<div class='keyword-name'>{html.escape(str(row.get('keyword', '')))}</div>"
            f"<div class='keyword-category'>{html.escape(str(row.get('category', 'Umum')))} | {status}</div>"
            "</div>"
            f"<span class='severity-badge severity-{severity}'>{html.escape(str(row.get('severity', 'medium')))}</span>"
            "</div>"
            f"<div class='keyword-desc'>{desc}</div>"
            "<div class='alias-label'>Sinonim/parafrasa yang dipakai sistem</div>"
            f"<div class='alias-row'>{alias_html}</div>"
            f"<button class='keyword-delete-btn' type='button' data-delete-keyword-id='{keyword_id}'>Hapus</button>"
            "</div>"
        )

    if not cards:
        cards.append("<div class='empty-findings'>Tidak ada keyword yang cocok dengan pencarian.</div>")

    return (
        "<div class='keyword-workspace'>"
        "<div class='keyword-search-panel'>"
        "<div class='keyword-search-head'>"
        f"<div class='keyword-search-title'>{html.escape(query or 'Keyword NAC')}</div>"
        "<div class='keyword-search-icon'>Search</div>"
        "</div>"
        "<div class='keyword-suggestion-strip'>"
        "<div class='keyword-hint'>Suggested Keywords: keyword NAC aktif yang akan dipakai sistem saat review</div>"
        f"<div class='keyword-chip-row'>{chips}</div>"
        "</div>"
        "</div>"
        f"<div class='simple-note'>{len(keywords)} keyword ditampilkan, {active_count} aktif. Sistem juga memakai fuzzy matching dan semantic similarity jika fitur semantic di Settings aktif, sehingga user tidak perlu memasukkan semua parafrasa secara manual.</div>"
        f"<div class='keyword-card-grid'>{''.join(cards)}</div>"
        "</div>"
    )


def import_keywords_ui(file_obj):
    if file_obj is None:
        return "Upload file import dahulu.", refresh_keywords()
    count = import_keywords_from_excel(_file_path(file_obj))
    return f"{count} keyword diimpor.", refresh_keywords()


def export_keywords_ui():
    path = BASE_DIR / "exports" / "nac_keyword_database.xlsx"
    path.parent.mkdir(exist_ok=True)
    return export_keyword_database(path)


def learning_ui():
    fp, fn, new_kw, syn, model_syn, exc, fb_hist = learning_summary()
    return render_learning_dashboard(fp, fn, new_kw, syn, model_syn, exc, fb_hist)


def _top_rows(frame, title, empty_text):
    if frame is None or frame.empty:
        return (
            "<div class='learning-card'>"
            f"<div class='learning-title'>{html.escape(title)}</div>"
            f"<div class='learning-copy'>{html.escape(empty_text)}</div>"
            "</div>"
        )
    rows_html = ""
    for _, row in frame.head(5).iterrows():
        item = row.get("item") or row.get("suggested_synonym") or row.get("matched_keyword") or "-"
        count = row.get("count", 0)
        rows_html += (
            "<div class='learning-row'>"
            f"<div class='learning-item'>{html.escape(str(item))}</div>"
            f"<div class='learning-count'>{html.escape(str(count))}</div>"
            "</div>"
        )
    return (
        "<div class='learning-card'>"
        f"<div class='learning-title'>{html.escape(title)}</div>"
        f"<div class='learning-list'>{rows_html}</div>"
        "</div>"
    )


def render_learning_dashboard(fp, fn, new_kw, syn, model_syn, exc, fb_hist):
    feedback_count = 0 if fb_hist is None or fb_hist.empty else len(fb_hist)
    fp_count = 0 if fp is None or fp.empty else int(fp["count"].sum())
    fn_count = 0 if fn is None or fn.empty else int(fn["count"].sum())
    syn_count = 0 if syn is None or syn.empty else int(syn["count"].sum())
    return (
        "<div class='keyword-workspace'>"
        "<div class='learning-grid'>"
        f"<div class='learning-card'><div class='learning-number'>{feedback_count}</div><div class='learning-label'>Total feedback reviewer</div></div>"
        f"<div class='learning-card'><div class='learning-number'>{fp_count}</div><div class='learning-label'>False positive perlu exception</div></div>"
        f"<div class='learning-card'><div class='learning-number'>{fn_count}</div><div class='learning-label'>False negative perlu keyword</div></div>"
        f"<div class='learning-card'><div class='learning-number'>{syn_count}</div><div class='learning-label'>Sinonim disarankan</div></div>"
        "</div>"
        "<div class='learning-grid'>"
        + _top_rows(fp, "Paling sering ditandai Bukan NAC", "Belum ada feedback false positive.")
        + _top_rows(fn, "Paling sering dikoreksi sebagai NAC", "Belum ada feedback false negative.")
        + _top_rows(new_kw, "Kandidat keyword baru", "Belum ada usulan keyword baru.")
        + _top_rows(syn, "Kandidat sinonim", "Belum ada usulan sinonim.")
        + _top_rows(exc, "Kandidat exception", "Belum ada pola exception.")
        + "</div>"
        "<div class='simple-note'>Dashboard ini memakai feedback reviewer untuk membantu menyarankan keyword, sinonim, dan exception. Tidak ada retraining otomatis atau keputusan final.</div>"
        "</div>"
    )


def analyze_redaction_ui(text):
    text = str(text or "").strip()
    if not text:
        return "<div class='empty-findings'>Ketik satu kalimat redaksi RAB untuk dianalisa.</div>"
    result = detect_items(
        [
            {
                "row_id": "redaksi",
                "source_file": "input_manual",
                "page_or_sheet": "Analisa Redaksi",
                "original_text": text,
                "item_description": text,
                "item_per_rab": text,
            }
        ],
        db.get_settings(),
    )[0]
    score = float(result.get("final_confidence", 0) or 0)
    label = result.get("confidence_label", "-")
    category = result.get("matched_category") or "Tidak ada kategori kuat"
    keyword = result.get("matched_keyword") or "-"
    suggestion = result.get("redaction_suggestion") or "Tidak ada saran khusus."
    explanation = result.get("explanation") or ""
    return (
        "<div class='redaction-result'>"
        f"<div><div class='redaction-score'>{score:.0f}%</div><span class='confidence-pill level-{_level_class(label)}'>{html.escape(label)}</span></div>"
        "<div>"
        f"<div class='redaction-title'>Potensi NAC: {html.escape(category)} | Keyword: {html.escape(keyword)}</div>"
        f"<div class='redaction-copy'>{html.escape(explanation)}</div>"
        f"<div class='redaction-copy'><strong>Saran klarifikasi:</strong> {html.escape(suggestion)}</div>"
        "</div>"
        "</div>"
    )


def export_review_ui(results):
    return export_review_excel(results or [])


def export_potential_pdf_ui(results):
    return export_potential_nac_pdf(results or [])


def export_all_pdf_ui(results):
    return export_all_materials_pdf(results or [])


def export_all_excel_ui(results):
    return export_all_materials_excel(results or [])


def backup_ui():
    return db.backup_db()


def restore_ui(file_obj):
    if file_obj is None:
        return "Upload backup SQLite dahulu."
    db.restore_db(_file_path(file_obj))
    return "Backup berhasil direstore."


def save_settings_ui(model, enable_semantic, enable_stemming, fuzzy, semantic, exact_w, syn_w, fuzzy_w, sem_w, sev_w, feedback_w, allowable_w, ocr_mode):
    values = {
        "embedding_model": model,
        "enable_semantic": str(enable_semantic).lower(),
        "enable_stemming": str(enable_stemming).lower(),
        "fuzzy_threshold": fuzzy,
        "semantic_threshold": semantic,
        "exact_weight": exact_w,
        "synonym_weight": syn_w,
        "fuzzy_weight": fuzzy_w,
        "semantic_weight": sem_w,
        "severity_weight": sev_w,
        "feedback_weight": feedback_w,
        "allowable_penalty_weight": allowable_w,
        "ocr_mode": ocr_mode,
        "semantic_user_configured": "true",
    }
    for k, v in values.items():
        db.save_setting(k, v)
    return "Settings tersimpan."


def save_simple_settings_ui(review_mode, semantic_mode, ocr_mode):
    sensitivity = {
        "Ketat": {"fuzzy_threshold": "86", "semantic_threshold": "72"},
        "Seimbang": {"fuzzy_threshold": "78", "semantic_threshold": "60"},
        "Lebih sensitif": {"fuzzy_threshold": "68", "semantic_threshold": "52"},
    }.get(review_mode, {"fuzzy_threshold": "78", "semantic_threshold": "60"})
    semantic_enabled = "true" if semantic_mode == "Aktif" else "false"
    values = {
        "enable_semantic": semantic_enabled,
        "enable_stemming": "false",
        "fuzzy_threshold": sensitivity["fuzzy_threshold"],
        "semantic_threshold": sensitivity["semantic_threshold"],
        "ocr_mode": ocr_mode,
        "semantic_user_configured": "true",
    }
    for key, value in values.items():
        db.save_setting(key, value)
    return f"Settings tersimpan: mode review {review_mode}, semantic {semantic_mode}, OCR {ocr_mode}."


APP_INTERACTIONS_JS = """
() => {
  if (window.__rabNacDeleteKeywordBound) return;
  window.__rabNacDeleteKeywordBound = true;
  document.addEventListener("click", (event) => {
    const button = event.target.closest("[data-delete-keyword-id]");
    if (button) {
      event.preventDefault();
      const keywordId = button.getAttribute("data-delete-keyword-id");
      const input = document.querySelector("#delete-kw-id textarea, #delete-kw-id input");
      const trigger = document.querySelector("#delete-kw-trigger button");
      if (!input || !trigger) return;
      input.value = keywordId;
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
      trigger.click();
      return;
    }

    const sortButton = event.target.closest(".materials-sort-header");
    if (!sortButton) return;
    event.preventDefault();
    const table = sortButton.closest("table");
    const tbody = table?.querySelector("tbody");
    if (!table || !tbody) return;
    const sortKey = sortButton.getAttribute("data-sort");
    const previousKey = table.getAttribute("data-sort-key");
    const previousDir = table.getAttribute("data-sort-dir") || "asc";
    const nextDir = previousKey === sortKey && previousDir === "asc" ? "desc" : "asc";
    table.setAttribute("data-sort-key", sortKey);
    table.setAttribute("data-sort-dir", nextDir);
    const numericKeys = new Set(["row", "confidence"]);
    const rows = Array.from(tbody.querySelectorAll("tr"));
    rows.sort((a, b) => {
      const av = a.dataset[sortKey] || "";
      const bv = b.dataset[sortKey] || "";
      let result;
      if (numericKeys.has(sortKey)) {
        result = (parseFloat(av) || 0) - (parseFloat(bv) || 0);
      } else {
        result = av.localeCompare(bv, "id", { sensitivity: "base", numeric: true });
      }
      return nextDir === "asc" ? result : -result;
    });
    rows.forEach((row) => tbody.appendChild(row));
  });
}
"""


def reset_db_ui():
    db.reset_demo_database()
    return "Demo database direset."


def review_started_ui():
    return (
        gr.update(value="Memproses NAC review... mohon tunggu sebentar.", visible=True),
        gr.update(interactive=False, value="Sedang memproses..."),
    )


def review_finished_ui():
    return gr.update(interactive=True, value="Run NAC Review")


def app():
    settings = db.get_settings()
    theme = gr.themes.Soft(
        primary_hue="blue",
        secondary_hue="blue",
        neutral_hue="slate",
        font=["DM Sans", "Inter", "ui-sans-serif", "system-ui", "sans-serif"],
        font_mono=["ui-monospace", "SFMono-Regular", "Consolas", "monospace"],
    )
    with gr.Blocks(title="RAB NAC Reviewer Copilot", theme=theme, css=FLUENTLY_THEME_CSS) as demo:
        upload_state = gr.State({})
        results_state = gr.State([])
        gr.Markdown(
            f"""
<section id="app-hero">
  <div class="version-pill"><span>Versi {APP_VERSION}</span></div>
  <h1>RAB NAC Reviewer Copilot</h1>
</section>
"""
        )
        with gr.Tabs():
            with gr.Tab("Upload RAB"):
                file_in = gr.File(
                    label="Upload RAB",
                    file_types=[".xlsx", ".xls", ".csv", ".pdf", ".png", ".jpg", ".jpeg"],
                    height=160,
                    elem_id="rab-upload-file",
                )
                upload_msg = gr.Markdown()
                preview = gr.Dataframe(label="Preview / Extracted Rows", visible=False)
                text_cols = gr.Dropdown(label="Kolom teks untuk digabung dan direview", multiselect=True, visible=False)
                volume_col = gr.Dropdown(label="Volume", visible=False)
                unit_col = gr.Dropdown(label="Unit", visible=False)
                unit_price_col = gr.Dropdown(label="Unit Price", visible=False)
                total_price_col = gr.Dropdown(label="Total Price", visible=False)
                run_btn = gr.Button("Run NAC Review", variant="primary")
                run_status = gr.Markdown(visible=False)
                with gr.Group(visible=False, elem_classes=["review-output-panel"]) as review_output_panel:
                    result_msg = gr.Markdown(visible=False)
                    with gr.Group(visible=False) as sort_panel:
                        with gr.Row():
                            sort_by = gr.Radio(["Confidence", "Row"], value="Confidence", label="Urutkan berdasarkan")
                            sort_order = gr.Radio(["Tinggi ke rendah", "Rendah ke tinggi"], value="Tinggi ke rendah", label="Arah urutan")
                    auto_results_df = gr.HTML(visible=False)
                    all_materials_df = gr.Dataframe(
                        label="Tabel Seluruh Item RAB",
                        value=all_materials_dataframe([]),
                        headers=["Row", "Item RAB", "Kategori NAC", "Confidence %", "Confidence Level"],
                        interactive=False,
                        wrap=True,
                        visible=False,
                    )
                    with gr.Group(visible=False) as export_panel:
                        with gr.Row():
                            export_potential_pdf_btn = gr.DownloadButton("Export PDF Rangkuman Potensi NAC")
                            export_all_pdf_btn = gr.DownloadButton("Export PDF Seluruh Material RAB")
                            export_all_excel_btn = gr.DownloadButton("Export Excel Seluruh Material RAB")
            with gr.Tab("Analisa Redaksi NAC"):
                gr.Markdown("Ketik satu kalimat redaksi RAB. Sistem akan menghitung potensi NAC sebagai bantuan awal review internal.")
                with gr.Group(elem_classes=["redaction-search"]):
                    redaction_text = gr.Textbox(
                        label="Analisa Redaksi NAC",
                        placeholder="Contoh: biaya konsumsi rapat koordinasi",
                        lines=2,
                    )
                    redaction_btn = gr.Button("Analisa Redaksi", variant="primary")
                redaction_result = gr.HTML(value=analyze_redaction_ui(""))
            with gr.Tab("Database NAC"):
                gr.Markdown("Kelola keyword NAC dengan cara sederhana. Seed database bersifat demo dan wajib divalidasi dengan PMK/kebijakan internal.")
                with gr.Accordion("Tambah keyword NAC", open=True):
                    simple_kw = gr.Textbox(label="Keyword NAC", placeholder="Contoh: uang saku, honorarium, biaya representasi")
                    simple_add_btn = gr.Button("Tambah Keyword NAC", variant="primary")
                    gr.Markdown("Kategori, confidence dasar, catatan, dan kandidat sinonim/parafrasa akan dipilih otomatis oleh sistem.")
                with gr.Accordion("Upload Excel keyword NAC", open=False, elem_id="keyword-upload-accordion"):
                    gr.Markdown("Kolom minimal: `category` dan `keyword`. Jika hanya punya daftar keyword, buat satu kolom bernama `keyword`.")
                    import_file = gr.File(label="Upload Excel Keyword NAC", file_types=[".xlsx"], elem_id="keyword-upload-file", height=96)
                    import_btn = gr.Button("Import Keyword dari Excel")
                    export_kw_btn = gr.Button("Export Database Keyword")
                    export_kw_file = gr.File(label="Download Keyword DB")
                delete_kw_id = gr.Textbox(elem_id="delete-kw-id", elem_classes=["hidden-delete-control"], label="delete_keyword_id")
                delete_kw_btn = gr.Button("Hapus Keyword NAC", elem_id="delete-kw-trigger", elem_classes=["hidden-delete-control"])
                kw_msg = gr.Markdown()
                keyword_search = gr.Textbox(label="Cari keyword NAC", placeholder="Contoh: konsumsi, honorarium, transport, hadiah")
                keyword_cards = gr.HTML(value=render_keyword_cards(""))
            with gr.Tab("Feedback & Learning"):
                learn_btn = gr.Button("Refresh Learning Dashboard", variant="primary")
                learning_html = gr.HTML(value=learning_ui())
            with gr.Tab("Settings"):
                gr.Markdown("Pengaturan dibuat sederhana untuk reviewer finance. Mode default `Seimbang` direkomendasikan.")
                review_mode = gr.Radio(["Ketat", "Seimbang", "Lebih sensitif"], value="Seimbang", label="Mode Review")
                semantic_mode = gr.Radio(["Nonaktif", "Aktif"], value="Aktif" if settings.get("enable_semantic", "false") == "true" else "Nonaktif", label="Deteksi Sinonim/Parafrasa Otomatis")
                ocr_mode = gr.Radio(["auto", "disabled"], value="auto" if settings.get("ocr_mode", "auto") != "disabled" else "disabled", label="OCR PDF Scan/Gambar")
                save_set = gr.Button("Simpan Settings", variant="primary")
                reset_db = gr.Button("Reset demo database")
                settings_msg = gr.Markdown()
                gr.HTML(
                    "<section class='versioning-panel'>"
                    "<h3>Versioning</h3>"
                    f"<p>{version_banner()}</p>"
                    "<p>Setiap penambahan fitur wajib memperbarui <code>modules/version.py</code> dan <code>CHANGELOG.md</code> dengan versi, judul, tanggal, dan keterangan.</p>"
                    "</section>"
                )

        file_in.change(
            handle_upload,
            file_in,
            [preview, text_cols, volume_col, unit_col, unit_price_col, total_price_col, upload_msg, upload_state],
        ).then(
            reset_analysis_visibility,
            outputs=[review_output_panel, result_msg, auto_results_df, all_materials_df, run_status, sort_panel, export_panel],
        )
        run_btn.click(
            review_started_ui,
            outputs=[run_status, run_btn],
            queue=False,
        ).then(
            run_review,
            [upload_state, text_cols, volume_col, unit_col, unit_price_col, total_price_col, sort_by, sort_order],
            [
                review_output_panel,
                result_msg,
                auto_results_df,
                all_materials_df,
                results_state,
                sort_panel,
                export_panel,
            ],
        ).then(
            review_finished_ui,
            outputs=run_btn,
        )
        sort_by.change(render_sorted_findings, [results_state, sort_by, sort_order], auto_results_df)
        sort_order.change(render_sorted_findings, [results_state, sort_by, sort_order], auto_results_df)
        export_potential_pdf_btn.click(
            export_potential_pdf_ui,
            inputs=results_state,
            outputs=export_potential_pdf_btn,
        )
        export_all_pdf_btn.click(
            export_all_pdf_ui,
            inputs=results_state,
            outputs=export_all_pdf_btn,
        )
        export_all_excel_btn.click(
            export_all_excel_ui,
            inputs=results_state,
            outputs=export_all_excel_btn,
        )
        redaction_btn.click(analyze_redaction_ui, redaction_text, redaction_result)
        redaction_text.submit(analyze_redaction_ui, redaction_text, redaction_result)
        redaction_text.change(analyze_redaction_ui, redaction_text, redaction_result)
        keyword_search.change(render_keyword_cards, keyword_search, keyword_cards)
        simple_add_btn.click(add_keyword_simple_ui, simple_kw, [kw_msg, keyword_cards, delete_kw_id])
        import_btn.click(import_keywords_simple_ui, import_file, [kw_msg, keyword_cards, delete_kw_id])
        delete_kw_btn.click(delete_keyword_by_id_ui, delete_kw_id, [kw_msg, keyword_cards])
        export_kw_btn.click(export_keywords_ui, outputs=export_kw_file)
        learn_btn.click(learning_ui, outputs=learning_html)
        save_set.click(save_simple_settings_ui, [review_mode, semantic_mode, ocr_mode], settings_msg)
        reset_db.click(reset_db_ui, outputs=settings_msg)
        demo.load(None, None, None, js=APP_INTERACTIONS_JS)
    return demo


if __name__ == "__main__":
    app().queue().launch()

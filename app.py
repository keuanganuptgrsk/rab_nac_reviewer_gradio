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
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');

:root {
    --rab-bg: #ffffff;
    --rab-primary: #0b1220;
    --rab-secondary: #475569;
    --rab-accent: #1665d6;
    --rab-text: #0b1220;
    --rab-heading: #1665d6;
    --rab-muted: #64748b;
    --rab-border: #dbe7f7;
    --rab-soft: #f8fbff;
    --rab-accent-soft: #eef6ff;
}

body,
.gradio-container {
    background: var(--rab-bg) !important;
    color: var(--rab-text) !important;
    font-family: "DM Sans", Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
}

.gradio-container {
    max-width: 1440px !important;
    margin: 0 auto !important;
    padding: 28px 34px 44px !important;
}

h1, h2, h3, .prose h1, .prose h2, .prose h3 {
    font-family: "DM Sans", Inter, sans-serif !important;
    color: var(--rab-heading) !important;
    letter-spacing: 0 !important;
}

#app-hero {
    background:
        radial-gradient(circle at 82% 18%, rgba(22, 101, 214, 0.16), transparent 32%),
        linear-gradient(135deg, #ffffff 0%, #f8fbff 48%, #eef6ff 100%);
    border: 1px solid var(--rab-border);
    border-radius: 8px;
    padding: 30px 32px;
    margin-bottom: 18px;
    box-shadow: 0 18px 48px rgba(15, 23, 42, 0.06);
}

#app-hero h1 {
    font-size: clamp(34px, 5vw, 72px);
    line-height: 1.02;
    margin: 0 0 12px;
    font-weight: 800;
}

#app-hero p {
    color: var(--rab-secondary);
    font-size: 18px;
    line-height: 1.55;
    max-width: 960px;
}

#app-hero .version-pill {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    color: #ffffff;
    background: var(--rab-primary);
    border: 1px solid #1e293b;
    border-radius: 999px;
    padding: 8px 13px;
    font-size: 13px;
    font-weight: 800;
    margin-bottom: 14px;
}

#app-hero .disclaimer {
    color: var(--rab-primary);
    font-weight: 600;
}

.tabs {
    border-bottom: 1px solid var(--rab-border) !important;
}

div[role="tabpanel"] {
    min-height: 720px;
    padding-top: 18px;
}

.tab-nav,
div[role="tablist"] {
    display: flex !important;
    flex-wrap: nowrap !important;
    overflow-x: auto !important;
    overflow-y: hidden !important;
    gap: 4px !important;
    scrollbar-width: thin;
}

.tab-nav button,
div[role="tablist"] button,
button[role="tab"] {
    flex: 0 0 auto !important;
    white-space: nowrap !important;
}

button[aria-label="More"],
button[title="More"],
.tab-nav button[aria-label="More"],
.tab-nav button[title="More"] {
    display: none !important;
}

.tab-nav button,
.tabs button {
    font-family: "DM Sans", Inter, sans-serif !important;
    font-weight: 650 !important;
    color: var(--rab-muted) !important;
    border-radius: 8px 8px 0 0 !important;
}

.tab-nav button.selected,
.tabs button.selected {
    color: var(--rab-accent) !important;
    border-bottom-color: var(--rab-accent) !important;
}

button.primary,
.primary > button,
button[variant="primary"] {
    background: var(--rab-accent) !important;
    border-color: var(--rab-accent) !important;
    color: #ffffff !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
}

button {
    border-radius: 8px !important;
    font-family: "DM Sans", Inter, sans-serif !important;
    font-weight: 700 !important;
}

.form, .panel, .block, .wrap, .container {
    border-radius: 8px !important;
}

label, .label-wrap span {
    color: var(--rab-text) !important;
    font-weight: 650 !important;
}

input, textarea, select {
    border-color: var(--rab-border) !important;
    border-radius: 8px !important;
}

.dataframe, .table-wrap {
    border-radius: 8px !important;
    border-color: var(--rab-border) !important;
}

table {
    font-family: "Geist", Inter, sans-serif !important;
}

th {
    background: var(--rab-soft) !important;
    color: var(--rab-primary) !important;
    font-weight: 700 !important;
}

td {
    color: var(--rab-text) !important;
}

a {
    color: var(--rab-accent) !important;
}

.gradio-container .prose p {
    color: var(--rab-secondary);
}

.markdown-code, code {
    border-radius: 6px !important;
}

.findings-panel {
    display: flex;
    flex-direction: column;
    gap: 12px;
    margin-top: 14px;
}

.findings-toolbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 14px;
    padding: 14px 16px;
    border: 1px solid var(--rab-border);
    background: linear-gradient(135deg, #f8fbff 0%, #eef6ff 100%);
    border-radius: 8px;
}

.findings-summary-copy strong {
    color: var(--rab-primary);
    font-size: 18px;
}

.findings-summary-copy span {
    color: var(--rab-secondary);
    font-size: 14px;
}

.sort-controls {
    display: flex;
    align-items: center;
    gap: 8px;
}

.sort-controls label {
    color: var(--rab-secondary) !important;
    font-size: 13px;
}

.finding-card {
    display: grid;
    grid-template-columns: 80px minmax(240px, 1fr) 130px 160px;
    align-items: center;
    gap: 16px;
    padding: 16px;
    border: 1px solid var(--rab-border);
    background: #ffffff;
    border-radius: 8px;
    box-shadow: 0 12px 32px rgba(15, 23, 42, 0.06);
}

.finding-card:hover {
    border-color: rgba(22, 101, 214, 0.45);
    box-shadow: 0 18px 42px rgba(22, 101, 214, 0.10);
}

.finding-row {
    width: 48px;
    height: 48px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 8px;
    background: var(--rab-accent-soft);
    color: var(--rab-accent);
    font-weight: 800;
    font-size: 18px;
}

.finding-item {
    color: var(--rab-primary);
    font-weight: 750;
    font-size: 17px;
    line-height: 1.35;
}

.finding-meta {
    color: var(--rab-muted);
    font-size: 13px;
    margin-top: 4px;
}

.confidence-score {
    font-weight: 800;
    font-size: 22px;
    color: var(--rab-primary);
}

.confidence-bar {
    height: 8px;
    width: 100%;
    background: #e2e8f0;
    border-radius: 999px;
    overflow: hidden;
    margin-top: 7px;
}

.confidence-fill {
    height: 100%;
    border-radius: 999px;
}

.confidence-pill {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 112px;
    border-radius: 999px;
    padding: 9px 13px;
    font-size: 14px;
    font-weight: 800;
}

.level-sedang {
    color: #92400e;
    background: #fef3c7;
}

.level-tinggi {
    color: #9a3412;
    background: #ffedd5;
}

.level-sangat-tinggi {
    color: #991b1b;
    background: #fee2e2;
}

.fill-sedang {
    background: #f59e0b;
}

.fill-tinggi {
    background: #f97316;
}

.fill-sangat-tinggi {
    background: #dc2626;
}

.empty-findings {
    padding: 22px;
    border: 1px solid var(--rab-border);
    background: #ffffff;
    border-radius: 8px;
    color: var(--rab-secondary);
    font-weight: 650;
}

.keyword-workspace {
    display: flex;
    flex-direction: column;
    gap: 16px;
}

.keyword-search-panel {
    border: 2px solid #b7dcf5;
    box-shadow: 0 0 0 5px rgba(186, 230, 253, 0.55);
    background: #ffffff;
    border-radius: 8px;
    overflow: hidden;
}

.keyword-search-head {
    display: grid;
    grid-template-columns: 1fr auto;
    align-items: center;
    gap: 14px;
    padding: 16px 18px;
    border-bottom: 1px solid var(--rab-border);
}

.keyword-search-title {
    color: var(--rab-primary);
    font-weight: 800;
    font-size: 20px;
}

.keyword-search-icon {
    color: #94a3b8;
    font-size: 22px;
}

.keyword-suggestion-strip {
    padding: 14px 18px 18px;
}

.keyword-hint {
    color: #9ca3af;
    font-size: 15px;
    margin-bottom: 12px;
}

.keyword-chip-row {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
}

.keyword-chip {
    display: inline-flex;
    align-items: center;
    border-radius: 8px;
    padding: 9px 14px;
    font-size: 15px;
    font-weight: 800;
    color: var(--rab-primary);
}

.chip-blue { background: #e0f2fe; }
.chip-yellow { background: #fef3c7; }
.chip-green { background: #dcfce7; }
.chip-purple { background: #ede9fe; }
.chip-red { background: #fee2e2; }

.keyword-card-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
    gap: 14px;
}

.keyword-card {
    border: 1px solid var(--rab-border);
    background: #ffffff;
    border-radius: 8px;
    padding: 16px;
    box-shadow: 0 12px 32px rgba(15, 23, 42, 0.06);
}

.keyword-card:hover {
    border-color: rgba(22, 101, 214, 0.45);
    box-shadow: 0 18px 42px rgba(22, 101, 214, 0.10);
}

.keyword-card-top {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 10px;
}

.keyword-name {
    color: var(--rab-primary);
    font-size: 19px;
    font-weight: 850;
    line-height: 1.2;
}

.keyword-category {
    color: var(--rab-secondary);
    font-size: 13px;
    font-weight: 700;
    margin-top: 5px;
}

.severity-badge {
    border-radius: 999px;
    padding: 7px 10px;
    font-size: 12px;
    font-weight: 850;
    white-space: nowrap;
}

.severity-medium { background: #fef3c7; color: #92400e; }
.severity-high { background: #ffedd5; color: #9a3412; }
.severity-very-high { background: #fee2e2; color: #991b1b; }
.severity-low, .severity-very-low { background: #e0f2fe; color: #075985; }

.keyword-desc {
    color: var(--rab-secondary);
    font-size: 13px;
    line-height: 1.45;
    margin-top: 10px;
}

.alias-label {
    color: var(--rab-primary);
    font-size: 12px;
    font-weight: 800;
    margin: 14px 0 8px;
}

.alias-row {
    display: flex;
    flex-wrap: wrap;
    gap: 7px;
}

.alias-chip {
    border-radius: 999px;
    padding: 6px 9px;
    background: var(--rab-accent-soft);
    color: var(--rab-accent);
    font-size: 12px;
    font-weight: 750;
}

.simple-note {
    border: 1px solid var(--rab-border);
    background: linear-gradient(135deg, #f8fbff 0%, #eef6ff 100%);
    border-radius: 8px;
    padding: 14px 16px;
    color: var(--rab-secondary);
    font-weight: 650;
}

.learning-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
    gap: 14px;
}

.learning-card {
    border: 1px solid var(--rab-border);
    border-radius: 8px;
    padding: 16px;
    background: #ffffff;
    box-shadow: 0 12px 30px rgba(15, 23, 42, 0.06);
}

.learning-number {
    font-size: 34px;
    font-weight: 850;
    color: var(--rab-accent);
}

.learning-label {
    color: var(--rab-secondary);
    font-size: 13px;
    font-weight: 750;
}

.learning-title {
    color: var(--rab-primary);
    font-size: 17px;
    font-weight: 850;
    margin-bottom: 10px;
}

.learning-copy {
    color: var(--rab-secondary);
    font-size: 14px;
    line-height: 1.5;
}

.learning-list {
    display: grid;
    gap: 10px;
    margin-top: 14px;
}

.learning-row {
    display: grid;
    grid-template-columns: 1fr auto;
    gap: 10px;
    align-items: center;
    padding: 12px 14px;
    background: var(--rab-soft);
    border: 1px solid var(--rab-border);
    border-radius: 8px;
}

.learning-item {
    color: var(--rab-primary);
    font-weight: 750;
}

.learning-count {
    background: var(--rab-accent);
    color: #ffffff;
    border-radius: 999px;
    padding: 5px 10px;
    font-weight: 850;
    font-size: 12px;
}

.redaction-search {
    border: 2px solid #b7dcf5;
    box-shadow: 0 0 0 5px rgba(186, 230, 253, 0.55);
    border-radius: 8px;
    padding: 16px;
    background: #ffffff;
}

.redaction-result {
    display: grid;
    grid-template-columns: 170px 1fr;
    gap: 16px;
    align-items: center;
    margin-top: 16px;
    padding: 18px;
    border: 1px solid var(--rab-border);
    border-radius: 8px;
    background: linear-gradient(135deg, #f8fbff 0%, #eef6ff 100%);
}

.redaction-score {
    color: var(--rab-accent);
    font-size: 42px;
    font-weight: 900;
    line-height: 1;
}

.redaction-title {
    color: var(--rab-primary);
    font-size: 18px;
    font-weight: 850;
}

.redaction-copy {
    color: var(--rab-secondary);
    line-height: 1.5;
    margin-top: 6px;
}

@media (max-width: 820px) {
    .finding-card {
        grid-template-columns: 58px 1fr;
    }

    .finding-card > div:nth-child(3),
    .finding-card > div:nth-child(4) {
        grid-column: 2;
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
            gr.update(value="", visible=False),
            gr.update(value=pd.DataFrame(), visible=False),
            [],
            msg,
            gr.update(value=pd.DataFrame(), visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
        )
    results = detect_items(items, db.get_settings())
    summary = review_summary_dataframe(results)
    all_materials = all_materials_dataframe(results)
    return (
        gr.update(value=render_findings_cards(results, sort_by, sort_order), visible=True),
        gr.update(value=all_materials, visible=True),
        results,
        f"Review selesai. {msg} Ditampilkan hanya confidence Sedang sampai Sangat tinggi. {DISCLAIMER}",
        gr.update(value=summary, visible=True),
        gr.update(visible=True),
        gr.update(visible=True),
        gr.update(visible=True),
        gr.update(visible=True),
        gr.update(visible=True),
    )


def auto_run_review(upload_state):
    if not upload_state:
        return pd.DataFrame(), [], "Upload file dahulu."
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
    return review_outputs[0], review_outputs[2], review_outputs[3]


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
        "item_per_rab": "Nama Material",
        "matched_category": "Kategori NAC",
        "final_confidence": "Confidence %",
        "confidence_label": "Kategori Confidence Level",
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


def render_sorted_findings(results, sort_by, sort_order):
    return gr.update(value=render_findings_cards(results, sort_by, sort_order), visible=bool(results))


def reset_analysis_visibility():
    return (
        gr.update(visible=False),
        gr.update(visible=False),
        gr.update(value="", visible=False),
        gr.update(value=pd.DataFrame(), visible=False),
    )


def render_findings_cards(results, sort_by="Confidence", sort_order="Tinggi ke rendah"):
    frame = pd.DataFrame(results or [])
    if frame.empty:
        return _empty_findings("Belum ada hasil review. Upload RAB lalu tekan Run NAC Review.")
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
        "<div class='findings-summary-copy'>"
        f"<strong>{len(frame)} potensi NAC perlu review</strong><br>"
        "<span>Fokus: Row, Item RAB, Confidence, dan Confidence Level.</span>"
        "</div>"
        "<div class='sort-controls'><label>Gunakan kontrol sort di atas untuk mengurutkan temuan.</label></div>"
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
        return "Isi nama keyword NAC terlebih dahulu.", render_keyword_cards(""), gr.update(choices=keyword_delete_choices(), value=None)
    existing = db.get_keyword_by_text(keyword)
    if existing:
        return f"Keyword '{keyword}' sudah ada.", render_keyword_cards(keyword), gr.update(choices=keyword_delete_choices(), value=None)
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
    return msg, render_keyword_cards(keyword), gr.update(choices=keyword_delete_choices(), value=None)


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


def import_keywords_simple_ui(file_obj):
    if file_obj is None:
        return "Upload file Excel keyword dahulu.", render_keyword_cards(""), gr.update(choices=keyword_delete_choices(), value=None)
    try:
        count = import_keywords_from_excel(_file_path(file_obj))
    except Exception as exc:
        return f"Import gagal: {exc}", render_keyword_cards(""), gr.update(choices=keyword_delete_choices(), value=None)
    return f"{count} keyword berhasil diimpor dari Excel.", render_keyword_cards(""), gr.update(choices=keyword_delete_choices(), value=None)


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
  <div class="version-pill">Versi {APP_VERSION}</div>
  <h1>RAB NAC Reviewer Copilot</h1>
</section>
"""
        )
        with gr.Tabs():
            with gr.Tab("Upload RAB"):
                file_in = gr.File(label="Upload RAB", file_types=[".xlsx", ".xls", ".csv", ".pdf", ".png", ".jpg", ".jpeg"])
                upload_msg = gr.Markdown()
                preview = gr.Dataframe(label="Preview / Extracted Rows", visible=False)
                text_cols = gr.Dropdown(label="Kolom teks untuk digabung dan direview", multiselect=True, visible=False)
                volume_col = gr.Dropdown(label="Volume", visible=False)
                unit_col = gr.Dropdown(label="Unit", visible=False)
                unit_price_col = gr.Dropdown(label="Unit Price", visible=False)
                total_price_col = gr.Dropdown(label="Total Price", visible=False)
                run_btn = gr.Button("Run NAC Review", variant="primary")
                run_status = gr.Markdown(visible=False)
                with gr.Group(visible=False) as sort_panel:
                    with gr.Row():
                        sort_by = gr.Radio(["Confidence", "Row"], value="Confidence", label="Urutkan berdasarkan")
                        sort_order = gr.Radio(["Tinggi ke rendah", "Rendah ke tinggi"], value="Tinggi ke rendah", label="Arah urutan")
                auto_results_df = gr.HTML(visible=False)
                with gr.Group(visible=False) as export_panel:
                    with gr.Row():
                        export_potential_pdf_btn = gr.DownloadButton(
                            "Export PDF Rangkuman Potensi NAC",
                            value=export_potential_pdf_ui,
                            inputs=results_state,
                        )
                        export_all_pdf_btn = gr.DownloadButton(
                            "Export PDF Seluruh Material RAB",
                            value=export_all_pdf_ui,
                            inputs=results_state,
                        )
                        export_all_excel_btn = gr.DownloadButton(
                            "Export Excel Seluruh Material RAB",
                            value=export_all_excel_ui,
                            inputs=results_state,
                        )
                all_materials_df = gr.Dataframe(label="Tabel Seluruh Material RAB", visible=False)
                result_msg = gr.Markdown(visible=False)
                results_df = gr.Dataframe(label="Review Hasil", visible=False)
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
                with gr.Accordion("Upload Excel keyword NAC", open=False):
                    gr.Markdown("Kolom minimal: `category` dan `keyword`. Jika hanya punya daftar keyword, buat satu kolom bernama `keyword`.")
                    import_file = gr.File(label="Upload Excel Keyword NAC", file_types=[".xlsx"])
                    import_btn = gr.Button("Import Keyword dari Excel")
                    export_kw_btn = gr.Button("Export Database Keyword")
                    export_kw_file = gr.File(label="Download Keyword DB")
                with gr.Accordion("Hapus keyword NAC", open=False):
                    delete_kw_choice = gr.Dropdown(choices=keyword_delete_choices(), label="Keyword NAC yang akan dihapus")
                    delete_kw_btn = gr.Button("Hapus Keyword NAC", variant="stop")
                    gr.Markdown("Keyword dibuat inactive, bukan dihapus permanen, agar riwayat audit tetap aman.")
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
                with gr.Accordion("Backup data", open=False):
                    backup_btn = gr.Button("Export SQLite database backup")
                    backup_file = gr.File(label="SQLite Backup")
                    restore_file = gr.File(label="Import SQLite database backup", file_types=[".db"])
                    restore_btn = gr.Button("Restore Backup")
                    restore_msg = gr.Markdown()
                reset_db = gr.Button("Reset demo database")
                settings_msg = gr.Markdown()
                gr.Markdown(
                    "### Versioning\n"
                    f"{version_banner()}\n\n"
                    "Setiap penambahan fitur wajib memperbarui `modules/version.py` dan `CHANGELOG.md` "
                    "dengan versi, judul, tanggal, dan keterangan."
                )

        file_in.change(
            handle_upload,
            file_in,
            [preview, text_cols, volume_col, unit_col, unit_price_col, total_price_col, upload_msg, upload_state],
        ).then(
            reset_analysis_visibility,
            outputs=[sort_panel, export_panel, auto_results_df, all_materials_df],
        )
        run_btn.click(
            review_started_ui,
            outputs=[run_status, run_btn],
            queue=False,
        ).then(
            run_review,
            [upload_state, text_cols, volume_col, unit_col, unit_price_col, total_price_col, sort_by, sort_order],
            [
                auto_results_df,
                all_materials_df,
                results_state,
                result_msg,
                results_df,
                sort_panel,
                export_panel,
                export_potential_pdf_btn,
                export_all_pdf_btn,
                export_all_excel_btn,
            ],
        ).then(
            review_finished_ui,
            outputs=run_btn,
        )
        sort_by.change(render_sorted_findings, [results_state, sort_by, sort_order], auto_results_df)
        sort_order.change(render_sorted_findings, [results_state, sort_by, sort_order], auto_results_df)
        redaction_btn.click(analyze_redaction_ui, redaction_text, redaction_result)
        redaction_text.submit(analyze_redaction_ui, redaction_text, redaction_result)
        redaction_text.change(analyze_redaction_ui, redaction_text, redaction_result)
        keyword_search.change(render_keyword_cards, keyword_search, keyword_cards)
        simple_add_btn.click(add_keyword_simple_ui, simple_kw, [kw_msg, keyword_cards, delete_kw_choice])
        import_btn.click(import_keywords_simple_ui, import_file, [kw_msg, keyword_cards, delete_kw_choice])
        delete_kw_btn.click(delete_keyword_simple_ui, delete_kw_choice, [kw_msg, keyword_cards, delete_kw_choice])
        export_kw_btn.click(export_keywords_ui, outputs=export_kw_file)
        learn_btn.click(learning_ui, outputs=learning_html)
        backup_btn.click(backup_ui, outputs=backup_file)
        restore_btn.click(restore_ui, restore_file, restore_msg)
        save_set.click(save_simple_settings_ui, [review_mode, semantic_mode, ocr_mode], settings_msg)
        reset_db.click(reset_db_ui, outputs=settings_msg)
    return demo


if __name__ == "__main__":
    app().queue().launch()

import os
import html
from pathlib import Path

import gradio as gr
import pandas as pd

from modules import db
from modules.excel_loader import combine_selected_text_columns, detect_columns, load_excel_or_csv, load_rab_excel_items, normalize_dataframe
from modules.export_engine import export_feedback_logs, export_review_excel
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

.findings-summary {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 14px;
    padding: 14px 16px;
    border: 1px solid var(--rab-border);
    background: linear-gradient(135deg, #f8fbff 0%, #eef6ff 100%);
    border-radius: 8px;
}

.findings-summary strong {
    color: var(--rab-primary);
    font-size: 18px;
}

.findings-summary span {
    color: var(--rab-secondary);
    font-size: 14px;
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


def run_review(upload_state, text_columns, volume_col, unit_col, unit_price_col, total_price_col):
    items, msg = build_items(upload_state, text_columns, volume_col, unit_col, unit_price_col, total_price_col)
    if not items:
        return gr.update(value="", visible=False), [], msg, gr.update(value=pd.DataFrame(), visible=False)
    results = detect_items(items, db.get_settings())
    summary = review_summary_dataframe(results)
    return (
        gr.update(value=render_findings_cards(results), visible=True),
        results,
        f"Review selesai. {msg} Ditampilkan hanya confidence Sedang sampai Sangat tinggi. {DISCLAIMER}",
        gr.update(value=summary, visible=True),
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
    results_df, results, msg = run_review(
        upload_state,
        text_columns,
        detected.get("volume") or ("volume" if "volume" in columns else None),
        detected.get("unit") or ("unit" if "unit" in columns else None),
        detected.get("unit_price") or ("unit_price" if "unit_price" in columns else None),
        detected.get("total_price") or ("total_price" if "total_price" in columns else None),
    )
    return results_df, results, msg


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


def render_findings_cards(results):
    frame = pd.DataFrame(results or [])
    if frame.empty:
        return _empty_findings("Belum ada hasil review. Upload RAB lalu tekan Run NAC Review.")
    frame = frame[frame["confidence_label"].isin(["Sedang", "Tinggi", "Sangat tinggi"])].copy()
    if frame.empty:
        return _empty_findings("Tidak ada item dengan confidence Sedang hingga Sangat tinggi.")
    frame = frame.sort_values(["final_confidence", "row_id"], ascending=[False, True])
    cards = [
        "<div class='findings-panel'>",
        "<div class='findings-summary'>"
        f"<strong>{len(frame)} potensi NAC perlu review</strong>"
        "<span>Ditampilkan sebagai kartu agar mudah discan. Fokus: Row, Item RAB, Confidence, dan Confidence Level.</span>"
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
        return "Pilih row_id yang memiliki suggested synonym.", refresh_synonyms()
    row = frame[frame["row_id"].astype(str) == str(row_id)]
    if row.empty:
        return "row_id tidak ditemukan.", refresh_synonyms()
    rec = row.iloc[0].to_dict()
    candidate = str(rec.get("suggested_synonym_candidate", "") or "").strip()
    keyword = str(rec.get("suggested_synonym_for_keyword", "") or rec.get("matched_keyword", "") or "").strip()
    if not candidate or not keyword:
        return "Baris ini belum memiliki kandidat sinonim dari model.", refresh_synonyms()
    keyword_row = db.get_keyword_by_text(keyword)
    if not keyword_row:
        return f"Keyword induk tidak ditemukan: {keyword}", refresh_synonyms()
    if db.synonym_exists(keyword_row["id"], candidate):
        return "Sinonim sudah ada di database.", refresh_synonyms()
    db.add_synonym(keyword_row["id"], candidate, float(weight or 0.85), "active")
    db.save_feedback(
        row_id,
        rec.get("original_text", ""),
        keyword,
        "Add as Synonym",
        candidate,
        "Approved model-suggested synonym",
    )
    return f"Sinonim '{candidate}' ditambahkan untuk keyword '{keyword}'.", refresh_synonyms()


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
    return learning_summary()


def export_review_ui(results):
    return export_review_excel(results or [])


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
                auto_results_df = gr.HTML(visible=False)
            with gr.Tab("Review Hasil"):
                result_msg = gr.Markdown()
                with gr.Row():
                    label_filter = gr.Dropdown(["Semua", "Sangat rendah", "Rendah", "Sedang", "Tinggi", "Sangat tinggi"], value="Semua", label="Confidence")
                    category_filter = gr.Textbox(label="Kategori")
                    match_filter = gr.Dropdown(["Semua", "exact", "synonym", "fuzzy", "semantic", "none"], value="Semua", label="Match Type")
                    keyword_filter = gr.Textbox(label="Keyword")
                with gr.Row():
                    medium_high = gr.Checkbox(value=True, label="Show only medium to very high confidence")
                    manual_only = gr.Checkbox(label="Show only manual review items")
                results_df = gr.Dataframe(label="Review Hasil", interactive=True)
                with gr.Row():
                    fb_row = gr.Textbox(label="row_id")
                    fb_type = gr.Dropdown(["Correct NAC", "Not NAC", "Manual Review", "Confidence Too High", "Confidence Too Low", "Add as New NAC Keyword", "Add as Synonym", "Add Exception"], label="Feedback")
                fb_redaction = gr.Textbox(label="User Suggested Redaction / Saran Redaksi/Klarifikasi")
                fb_notes = gr.Textbox(label="Reviewer Notes")
                fb_btn = gr.Button("Simpan Feedback")
                fb_msg = gr.Markdown()
                gr.Markdown("### Suggested Synonyms dari Model HF\nKandidat ini berasal dari semantic/fuzzy match dan harus divalidasi reviewer sebelum masuk database.")
                with gr.Row():
                    syn_row = gr.Textbox(label="row_id kandidat sinonim")
                    syn_approve_weight = gr.Number(value=0.85, label="Synonym weight")
                    approve_syn_btn = gr.Button("Approve Suggested Synonym")
                approve_syn_msg = gr.Markdown()
            with gr.Tab("Database NAC"):
                gr.Markdown("Seed database bersifat demo dan wajib divalidasi dengan PMK/kebijakan internal.")
                kw_table = gr.Dataframe(value=refresh_keywords(), label="NAC Keywords")
                with gr.Row():
                    kw_cat = gr.Textbox(label="Category")
                    kw_text = gr.Textbox(label="Tambah Keyword NAC")
                    kw_sev = gr.Dropdown(["very_low", "low", "medium", "high", "very_high"], value="medium", label="Severity")
                    kw_status = gr.Dropdown(["active", "inactive", "deprecated", "needs_review"], value="active", label="Status")
                kw_desc = gr.Textbox(label="Description")
                kw_ref = gr.Textbox(label="Reference")
                kw_notes = gr.Textbox(label="Notes")
                add_kw = gr.Button("Tambah Keyword NAC")
                kw_msg = gr.Markdown()
                with gr.Row():
                    kw_status_id = gr.Number(label="Keyword ID untuk update status", precision=0)
                    kw_status_new = gr.Dropdown(["active", "inactive", "deprecated", "needs_review"], value="inactive", label="Status baru")
                    kw_status_btn = gr.Button("Update Status Keyword")
                syn_table = gr.Dataframe(value=refresh_synonyms(), label="Sinonim")
                with gr.Row():
                    syn_parent = gr.Dropdown(choices=[str(r["id"]) for r in db.get_keywords(False)], label="Keyword ID")
                    syn_text = gr.Textbox(label="Tambah Sinonim")
                    syn_weight = gr.Number(value=0.9, label="Weight")
                add_syn = gr.Button("Tambah Sinonim")
                syn_msg = gr.Markdown()
                with gr.Row():
                    syn_status_id = gr.Number(label="Synonym ID untuk update status", precision=0)
                    syn_status_new = gr.Dropdown(["active", "inactive", "deprecated", "needs_review"], value="inactive", label="Status baru")
                    syn_status_btn = gr.Button("Update Status Synonym")
                allow_table = gr.Dataframe(value=refresh_allowable(), label="Allowable Keywords")
                with gr.Row():
                    allow_cat = gr.Textbox(label="Category", value="Teknis")
                    allow_kw = gr.Textbox(label="Add Allowable Keyword")
                allow_desc = gr.Textbox(label="Description")
                add_allow = gr.Button("Tambah Allowable")
                allow_msg = gr.Markdown()
                with gr.Row():
                    allow_status_id = gr.Number(label="Allowable ID untuk update status", precision=0)
                    allow_status_new = gr.Dropdown(["active", "inactive", "deprecated", "needs_review"], value="inactive", label="Status baru")
                    allow_status_btn = gr.Button("Update Status Allowable")
                exc_table = gr.Dataframe(value=refresh_exceptions(), label="Exceptions")
                with gr.Row():
                    exc_parent = gr.Dropdown(choices=[""] + [str(r["id"]) for r in db.get_keywords(False)], label="NAC Keyword ID nullable")
                    exc_pattern = gr.Textbox(label="Tambah Exception")
                    exc_action = gr.Dropdown(["lower_confidence", "ignore", "manual_review"], value="lower_confidence", label="Action")
                    exc_adjust = gr.Number(value=25, label="Weight Adjustment")
                exc_reason = gr.Textbox(label="Reason")
                add_exc = gr.Button("Tambah Exception")
                exc_msg = gr.Markdown()
                with gr.Row():
                    exc_status_id = gr.Number(label="Exception ID untuk update status", precision=0)
                    exc_status_new = gr.Dropdown(["active", "inactive", "deprecated", "needs_review"], value="inactive", label="Status baru")
                    exc_status_btn = gr.Button("Update Status Exception")
                import_file = gr.File(label="Import NAC keyword database from Excel", file_types=[".xlsx"])
                import_btn = gr.Button("Import")
                export_kw_btn = gr.Button("Export NAC keyword database")
                export_kw_file = gr.File(label="Download Keyword DB")
            with gr.Tab("Feedback & Learning"):
                learn_btn = gr.Button("Refresh Learning Dashboard")
                fp = gr.Dataframe(label="Most Common False Positives")
                fn = gr.Dataframe(label="Most Common False Negatives")
                sug_kw = gr.Dataframe(label="Frequently Suggested New NAC Keywords")
                sug_syn = gr.Dataframe(label="Frequently Suggested Synonyms")
                model_syn = gr.Dataframe(label="Model-Suggested Synonyms Approved by Reviewer")
                sug_exc = gr.Dataframe(label="Suggested Exceptions")
                fb_hist = gr.Dataframe(label="Feedback History")
            with gr.Tab("Export Excel"):
                export_btn = gr.Button("Export Excel")
                export_file = gr.File(label="Export Excel")
                export_fb_btn = gr.Button("Export feedback logs")
                export_fb_file = gr.File(label="Feedback Excel")
                backup_btn = gr.Button("Export SQLite database backup")
                backup_file = gr.File(label="SQLite Backup")
                restore_file = gr.File(label="Import SQLite database backup", file_types=[".db"])
                restore_btn = gr.Button("Restore Backup")
                restore_msg = gr.Markdown()
            with gr.Tab("Settings"):
                model = gr.Dropdown(
                    [
                        "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2",
                        "intfloat/multilingual-e5-small",
                        "intfloat/multilingual-e5-base",
                        "firqaaa/indo-sentence-bert-base",
                        "firqaaa/indo-sentence-bert-large",
                    ],
                    value=settings.get("embedding_model"),
                    label="Embedding model name",
                    allow_custom_value=True,
                )
                enable_sem = gr.Checkbox(value=settings.get("enable_semantic", "false") == "true", label="Enable semantic matching")
                enable_stem = gr.Checkbox(value=settings.get("enable_stemming", "false") == "true", label="Enable stemming")
                with gr.Row():
                    fuzzy_thr = gr.Number(value=float(settings.get("fuzzy_threshold", 78)), label="Fuzzy threshold")
                    semantic_thr = gr.Number(value=float(settings.get("semantic_threshold", 60)), label="Semantic threshold")
                    ocr_mode = gr.Radio(["auto", "disabled", "easyocr", "paddleocr", "tesseract"], value=settings.get("ocr_mode", "auto"), label="OCR mode")
                with gr.Row():
                    exact_w = gr.Number(value=float(settings.get("exact_weight", 0.25)), label="Exact match weight")
                    syn_w = gr.Number(value=float(settings.get("synonym_weight", 0.25)), label="Synonym match weight")
                    fuzzy_w = gr.Number(value=float(settings.get("fuzzy_weight", 0.20)), label="Fuzzy match weight")
                    sem_w = gr.Number(value=float(settings.get("semantic_weight", 0.30)), label="Semantic match weight")
                with gr.Row():
                    sev_w = gr.Number(value=float(settings.get("severity_weight", 0.10)), label="Severity weight")
                    feedback_w = gr.Number(value=float(settings.get("feedback_weight", 0.10)), label="Feedback adjustment weight")
                    allowable_w = gr.Number(value=float(settings.get("allowable_penalty_weight", 0.20)), label="Allowable competitor penalty weight")
                save_set = gr.Button("Save Settings")
                reset_db = gr.Button("Reset demo database")
                settings_msg = gr.Markdown()
                rebuild_btn = gr.Button("Rebuild embeddings")
                rebuild_msg = gr.Markdown("Embeddings dibangun lazy saat review; tombol ini hanya penanda refresh cache pada versi demo.")
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
        )
        run_btn.click(
            review_started_ui,
            outputs=[run_status, run_btn],
            queue=False,
        ).then(
            run_review,
            [upload_state, text_cols, volume_col, unit_col, unit_price_col, total_price_col],
            [auto_results_df, results_state, result_msg, results_df],
        ).then(
            review_finished_ui,
            outputs=run_btn,
        )
        for control in [label_filter, category_filter, match_filter, keyword_filter, medium_high, manual_only]:
            control.change(filter_results, [results_state, label_filter, category_filter, match_filter, keyword_filter, medium_high, manual_only], results_df)
        fb_btn.click(save_row_feedback, [results_state, fb_row, fb_type, fb_redaction, fb_notes], fb_msg)
        approve_syn_btn.click(approve_suggested_synonym, [results_state, syn_row, syn_approve_weight], [approve_syn_msg, syn_table])
        add_kw.click(add_keyword_ui, [kw_cat, kw_text, kw_desc, kw_ref, kw_sev, kw_status, kw_notes], [kw_msg, kw_table])
        add_syn.click(add_synonym_ui, [syn_parent, syn_text, syn_weight], [syn_msg, syn_table])
        kw_status_btn.click(update_keyword_status_ui, [kw_status_id, kw_status_new], [kw_msg, kw_table])
        syn_status_btn.click(update_synonym_status_ui, [syn_status_id, syn_status_new], [syn_msg, syn_table])
        add_allow.click(add_allowable_ui, [allow_cat, allow_kw, allow_desc], [allow_msg, allow_table])
        allow_status_btn.click(update_allowable_status_ui, [allow_status_id, allow_status_new], [allow_msg, allow_table])
        add_exc.click(add_exception_ui, [exc_parent, exc_pattern, exc_reason, exc_action, exc_adjust], [exc_msg, exc_table])
        exc_status_btn.click(update_exception_status_ui, [exc_status_id, exc_status_new], [exc_msg, exc_table])
        import_btn.click(import_keywords_ui, import_file, [kw_msg, kw_table])
        export_kw_btn.click(export_keywords_ui, outputs=export_kw_file)
        learn_btn.click(learning_ui, outputs=[fp, fn, sug_kw, sug_syn, model_syn, sug_exc, fb_hist])
        export_btn.click(export_review_ui, results_state, export_file)
        export_fb_btn.click(export_feedback_logs, outputs=export_fb_file)
        backup_btn.click(backup_ui, outputs=backup_file)
        restore_btn.click(restore_ui, restore_file, restore_msg)
        save_set.click(save_settings_ui, [model, enable_sem, enable_stem, fuzzy_thr, semantic_thr, exact_w, syn_w, fuzzy_w, sem_w, sev_w, feedback_w, allowable_w, ocr_mode], settings_msg)
        reset_db.click(reset_db_ui, outputs=settings_msg)
    return demo


if __name__ == "__main__":
    app().queue().launch()

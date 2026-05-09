import os
from pathlib import Path

import gradio as gr
import pandas as pd

from modules import db
from modules.excel_loader import combine_selected_text_columns, detect_columns, load_excel_or_csv, normalize_dataframe
from modules.export_engine import export_feedback_logs, export_review_excel
from modules.feedback_engine import learning_summary
from modules.keyword_manager import export_keyword_database, import_keywords_from_excel
from modules.nac_detector import detect_items
from modules.ocr_engine import extract_text_from_image, extract_text_from_pdf_scan
from modules.pdf_loader import extract_text_from_pdf


DISCLAIMER = (
    "Hasil deteksi adalah bantuan awal untuk review internal. Keputusan final tetap harus divalidasi oleh reviewer "
    "yang memahami PMK, kebijakan internal, dan konteks pekerjaan."
)
BASE_DIR = Path(__file__).resolve().parent
db.init_db()


def _file_path(file_obj):
    return file_obj.name if hasattr(file_obj, "name") else str(file_obj)


def handle_upload(file_obj):
    if file_obj is None:
        empty = gr.update(choices=[], value=None)
        return pd.DataFrame(), empty, empty, empty, empty, empty, "Upload file terlebih dahulu.", {}
    path = Path(_file_path(file_obj))
    suffix = path.suffix.lower()
    if suffix in [".xlsx", ".xls", ".csv"]:
        loaded = load_excel_or_csv(path)
        frame = normalize_dataframe(loaded["dataframe"])
        detected = detect_columns(frame)
        cols = list(frame.columns)
        text_defaults = [c for k, c in detected.items() if k in ("work_title", "description", "material_service_name", "notes")]
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
                "row_id": str(idx + 1),
                "source_file": path.name,
                "page_or_sheet": upload_state.get("sheet", ""),
                "original_text": text,
                "item_description": text,
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
        return pd.DataFrame(), [], msg
    results = detect_items(items, db.get_settings())
    return pd.DataFrame(results), results, f"Review selesai. {msg} {DISCLAIMER}"


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
    return frame


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
    }
    for k, v in values.items():
        db.save_setting(k, v)
    return "Settings tersimpan."


def reset_db_ui():
    db.reset_demo_database()
    return "Demo database direset."


def app():
    settings = db.get_settings()
    with gr.Blocks(title="RAB NAC Reviewer Copilot") as demo:
        upload_state = gr.State({})
        results_state = gr.State([])
        gr.Markdown("# RAB NAC Reviewer Copilot\n**Bukan Keputusan Final.** " + DISCLAIMER)
        with gr.Tabs():
            with gr.Tab("Upload RAB"):
                file_in = gr.File(label="Upload RAB", file_types=[".xlsx", ".xls", ".csv", ".pdf", ".png", ".jpg", ".jpeg"])
                upload_msg = gr.Markdown()
                preview = gr.Dataframe(label="Preview / Extracted Rows")
                text_cols = gr.Dropdown(label="Kolom teks untuk digabung dan direview", multiselect=True)
                with gr.Row():
                    volume_col = gr.Dropdown(label="Volume")
                    unit_col = gr.Dropdown(label="Unit")
                    unit_price_col = gr.Dropdown(label="Unit Price")
                    total_price_col = gr.Dropdown(label="Total Price")
                run_btn = gr.Button("Run NAC Review", variant="primary")
                file_in.change(
                    handle_upload,
                    file_in,
                    [preview, text_cols, volume_col, unit_col, unit_price_col, total_price_col, upload_msg, upload_state],
                )
            with gr.Tab("Review Hasil"):
                result_msg = gr.Markdown()
                with gr.Row():
                    label_filter = gr.Dropdown(["Semua", "Sangat rendah", "Rendah", "Sedang", "Tinggi", "Sangat tinggi"], value="Semua", label="Confidence")
                    category_filter = gr.Textbox(label="Kategori")
                    match_filter = gr.Dropdown(["Semua", "exact", "synonym", "fuzzy", "semantic", "none"], value="Semua", label="Match Type")
                    keyword_filter = gr.Textbox(label="Keyword")
                with gr.Row():
                    medium_high = gr.Checkbox(label="Show only medium to very high confidence")
                    manual_only = gr.Checkbox(label="Show only manual review items")
                results_df = gr.Dataframe(label="Review Hasil", interactive=True)
                with gr.Row():
                    fb_row = gr.Textbox(label="row_id")
                    fb_type = gr.Dropdown(["Correct NAC", "Not NAC", "Manual Review", "Confidence Too High", "Confidence Too Low", "Add as New NAC Keyword", "Add as Synonym", "Add Exception"], label="Feedback")
                fb_redaction = gr.Textbox(label="User Suggested Redaction / Saran Redaksi/Klarifikasi")
                fb_notes = gr.Textbox(label="Reviewer Notes")
                fb_btn = gr.Button("Simpan Feedback")
                fb_msg = gr.Markdown()
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
                enable_sem = gr.Checkbox(value=settings.get("enable_semantic", "true") == "true", label="Enable semantic matching")
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

        run_btn.click(run_review, [upload_state, text_cols, volume_col, unit_col, unit_price_col, total_price_col], [results_df, results_state, result_msg])
        for control in [label_filter, category_filter, match_filter, keyword_filter, medium_high, manual_only]:
            control.change(filter_results, [results_state, label_filter, category_filter, match_filter, keyword_filter, medium_high, manual_only], results_df)
        fb_btn.click(save_row_feedback, [results_state, fb_row, fb_type, fb_redaction, fb_notes], fb_msg)
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
        learn_btn.click(learning_ui, outputs=[fp, fn, sug_kw, sug_syn, sug_exc, fb_hist])
        export_btn.click(export_review_ui, results_state, export_file)
        export_fb_btn.click(export_feedback_logs, outputs=export_fb_file)
        backup_btn.click(backup_ui, outputs=backup_file)
        restore_btn.click(restore_ui, restore_file, restore_msg)
        save_set.click(save_settings_ui, [model, enable_sem, enable_stem, fuzzy_thr, semantic_thr, exact_w, syn_w, fuzzy_w, sem_w, sev_w, feedback_w, allowable_w, ocr_mode], settings_msg)
        reset_db.click(reset_db_ui, outputs=settings_msg)
    return demo


if __name__ == "__main__":
    app().queue().launch()

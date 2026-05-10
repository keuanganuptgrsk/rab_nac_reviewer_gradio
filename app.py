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
@import url('https://fonts.googleapis.com/css2?family=Fira+Sans:wght@300;400;500;600;700;800&family=Fira+Code:wght@400;500;600;700&family=Montserrat:wght@400;500;600;700;800;900&display=swap');

:root {
    --rab-bg: #ffffff;
    --rab-primary: #263d5b;
    --rab-secondary: #4b637f;
    --rab-accent: #49b6e5;
    --rab-success: #16a34a;
    --rab-warning: #d97706;
    --rab-danger: #dc2626;
    --rab-text: #111827;
    --rab-heading: #263d5b;
    --rab-muted: #64748b;
    --rab-border: #263d5b;
    --rab-soft: #f6fbff;
    --rab-accent-soft: #e7f7fd;
    --rab-paper: #fffef8;
}

body,
.gradio-container {
    background:
        radial-gradient(circle at 10% 12%, rgba(73, 182, 229, 0.12), transparent 25%),
        radial-gradient(circle at 88% 8%, rgba(217, 119, 6, 0.10), transparent 22%),
        linear-gradient(180deg, #ffffff 0%, #f8fcff 100%) !important;
    color: var(--rab-text) !important;
    font-family: "Montserrat", Inter, ui-sans-serif, system-ui, sans-serif !important;
}

.gradio-container {
    max-width: 1440px !important;
    margin: 0 auto !important;
    padding: 28px 34px 44px !important;
}

h1, h2, h3, .prose h1, .prose h2, .prose h3 {
    font-family: "Space Grotesk", "Montserrat", Inter, sans-serif !important;
    color: var(--rab-heading) !important;
    letter-spacing: 0 !important;
}

#app-hero {
    background:
        linear-gradient(135deg, rgba(73, 182, 229, 0.20) 0%, rgba(255, 255, 255, 0.92) 52%, rgba(255, 246, 218, 0.92) 100%);
    border: 3px solid var(--rab-border);
    border-radius: 18px 22px 16px 24px;
    padding: 30px 32px;
    margin-bottom: 18px;
    box-shadow: 8px 8px 0 rgba(38, 61, 91, 0.16);
    transform: rotate(-0.15deg);
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
    border-radius: 999px 900px 999px 840px;
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
    border-bottom: 3px solid var(--rab-border) !important;
    position: sticky;
    top: 0;
    z-index: 20;
    background: rgba(255, 255, 255, 0.96);
    backdrop-filter: blur(8px);
}

div[role="tabpanel"] {
    min-height: 780px;
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
    font-family: "Montserrat", Inter, sans-serif !important;
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
    border: 3px solid var(--rab-primary) !important;
    color: var(--rab-primary) !important;
    border-radius: 13px 17px 12px 18px !important;
    font-weight: 700 !important;
    box-shadow: 5px 5px 0 rgba(38, 61, 91, 0.22) !important;
}

button {
    border-radius: 13px 17px 12px 18px !important;
    font-family: "Montserrat", Inter, sans-serif !important;
    font-weight: 700 !important;
}

.form, .panel, .block, .wrap, .container {
    border-radius: 18px 16px 22px 14px !important;
}

label, .label-wrap span {
    color: var(--rab-text) !important;
    font-weight: 650 !important;
}

input, textarea, select {
    border: 3px solid var(--rab-border) !important;
    border-radius: 16px 13px 18px 12px !important;
    background: var(--rab-paper) !important;
}

.dataframe, .table-wrap {
    border-radius: 18px 14px 20px 16px !important;
    border-color: var(--rab-border) !important;
}

table {
    font-family: "JetBrains Mono", ui-monospace, monospace !important;
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
    border: 3px solid var(--rab-border);
    background: var(--rab-paper);
    border-radius: 18px 14px 22px 16px;
    box-shadow: 6px 6px 0 rgba(38, 61, 91, 0.15);
    transform: rotate(-0.08deg);
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
    border-radius: 13px 17px 12px 18px;
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
    border: 3px solid var(--rab-border);
    box-shadow: 7px 7px 0 rgba(73, 182, 229, 0.20);
    background: var(--rab-paper);
    border-radius: 18px 14px 22px 16px;
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
    border: 2px solid rgba(38, 61, 91, 0.22);
    border-radius: 13px 17px 12px 18px;
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
    border: 3px solid var(--rab-border);
    background: var(--rab-paper);
    border-radius: 18px 14px 22px 16px;
    padding: 16px;
    box-shadow: 7px 7px 0 rgba(38, 61, 91, 0.12);
    transform: rotate(-0.12deg);
}

.keyword-card:hover {
    border-color: var(--rab-accent);
    box-shadow: 9px 9px 0 rgba(73, 182, 229, 0.24);
    transform: translate(-2px, -2px) rotate(0.08deg);
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
    border: 2px solid rgba(38, 61, 91, 0.22);
    border-radius: 999px 880px 940px 840px;
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
    border: 2px solid rgba(73, 182, 229, 0.30);
    border-radius: 999px 880px 940px 840px;
    padding: 6px 9px;
    background: var(--rab-accent-soft);
    color: var(--rab-accent);
    font-size: 12px;
    font-weight: 750;
}

.keyword-delete-btn {
    width: 100%;
    margin-top: 14px;
    border: 3px solid var(--rab-danger);
    background: #fff1f2;
    color: #991b1b;
    border-radius: 13px 17px 12px 18px;
    padding: 9px 12px;
    font-family: "Montserrat", Inter, sans-serif;
    font-weight: 850;
    cursor: pointer;
    box-shadow: 4px 4px 0 rgba(220, 38, 38, 0.18);
}

.keyword-delete-btn:hover {
    transform: translate(-1px, -1px);
    box-shadow: 6px 6px 0 rgba(220, 38, 38, 0.22);
}

.hidden-delete-control {
    position: absolute !important;
    width: 1px !important;
    height: 1px !important;
    overflow: hidden !important;
    opacity: 0 !important;
    pointer-events: none !important;
}

/* Perspective-inspired UI override: bright SaaS funnel style with stable full-width panels */
body,
.gradio-container {
    background:
        radial-gradient(circle at 14% 18%, rgba(18, 184, 134, 0.18), transparent 30%),
        radial-gradient(circle at 86% 8%, rgba(184, 255, 67, 0.22), transparent 28%),
        linear-gradient(180deg, #f7fff8 0%, #ffffff 48%, #f8fafc 100%) !important;
    font-family: "Montserrat", Inter, ui-sans-serif, system-ui, sans-serif !important;
}

.gradio-container {
    width: min(100%, 1280px) !important;
    max-width: 1280px !important;
    min-width: 0 !important;
    box-sizing: border-box !important;
    padding: 34px 28px 56px !important;
}

.gradio-container > *,
.gradio-container .tabs,
.gradio-container [role="tablist"],
.gradio-container [role="tabpanel"],
.gradio-container [data-testid="tabs"],
.gradio-container .tabitem,
.gradio-container .form,
.gradio-container .block,
.gradio-container .wrap,
.gradio-container .container,
#app-hero {
    width: 100% !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
}

h1, h2, h3, .prose h1, .prose h2, .prose h3,
#app-hero h1,
.keyword-name,
.finding-item,
.redaction-title,
.learning-title {
    font-family: "Space Grotesk", "Montserrat", Inter, sans-serif !important;
}

#app-hero {
    background:
        radial-gradient(circle at 74% 18%, rgba(184, 255, 67, 0.28), transparent 34%),
        linear-gradient(135deg, #092013 0%, #0a3b24 48%, #10b981 100%) !important;
    border: 0 !important;
    border-radius: 32px !important;
    box-shadow: 0 30px 80px rgba(15, 118, 82, 0.24) !important;
    transform: none !important;
    padding: 42px 44px !important;
}

#app-hero h1,
#app-hero .version-pill {
    color: #ffffff !important;
}

#app-hero .version-pill {
    background: rgba(184, 255, 67, 0.18) !important;
    border: 1px solid rgba(184, 255, 67, 0.38) !important;
    border-radius: 999px !important;
}

.tabs {
    border: 1px solid rgba(15, 118, 82, 0.12) !important;
    border-radius: 28px !important;
    background: rgba(255, 255, 255, 0.92) !important;
    box-shadow: 0 24px 64px rgba(15, 23, 42, 0.08) !important;
    overflow: hidden !important;
    margin: 0 auto !important;
}

div[role="tabpanel"] {
    min-height: 780px !important;
    padding: 32px !important;
    background: #ffffff !important;
}

.tab-nav button,
.tabs button,
button[role="tab"] {
    font-family: "Montserrat", Inter, sans-serif !important;
    border-radius: 999px !important;
    color: #64748b !important;
    font-weight: 800 !important;
}

.tab-nav button.selected,
.tabs button.selected,
button[role="tab"][aria-selected="true"] {
    color: #062414 !important;
    background: linear-gradient(135deg, #b8ff43 0%, #36e181 100%) !important;
}

button.primary,
.primary > button,
button[variant="primary"],
.gradio-container button:not([role="tab"]) {
    border: 0 !important;
    border-radius: 999px !important;
    background: linear-gradient(135deg, #b8ff43 0%, #22c55e 100%) !important;
    color: #062414 !important;
    box-shadow: 0 18px 34px rgba(34, 197, 94, 0.24) !important;
    transform: none !important;
    font-family: "Montserrat", Inter, sans-serif !important;
}

input, textarea, select {
    border: 1px solid rgba(15, 118, 82, 0.18) !important;
    border-radius: 18px !important;
    background: #f8fffb !important;
    font-family: "Montserrat", Inter, sans-serif !important;
}

.finding-card,
.keyword-card,
.learning-card,
.redaction-result,
.simple-note,
.keyword-search-panel,
.redaction-search {
    border: 1px solid rgba(15, 118, 82, 0.12) !important;
    border-radius: 28px !important;
    background: rgba(255, 255, 255, 0.96) !important;
    box-shadow: 0 22px 56px rgba(15, 23, 42, 0.08) !important;
    transform: none !important;
}

.keyword-card:hover,
.finding-card:hover {
    box-shadow: 0 26px 64px rgba(15, 118, 82, 0.16) !important;
    transform: translateY(-2px) !important;
}

.keyword-chip,
.alias-chip,
.severity-badge,
.confidence-pill,
.learning-count {
    border: 0 !important;
    border-radius: 999px !important;
    background: linear-gradient(135deg, rgba(184, 255, 67, 0.40), rgba(34, 197, 94, 0.18)) !important;
    color: #064e3b !important;
}

.keyword-delete-btn {
    border: 0 !important;
    border-radius: 14px !important;
    background: linear-gradient(135deg, #ef4444 0%, #f97316 100%) !important;
    color: #ffffff !important;
    box-shadow: 0 14px 28px rgba(220, 38, 38, 0.20) !important;
    font-family: "Montserrat", Inter, sans-serif !important;
}

.gradio-container .gap,
.gradio-container .row,
.gradio-container .column,
.gradio-container .form,
.gradio-container .block,
.gradio-container .block > div,
.gradio-container .wrap,
.gradio-container .contain,
.gradio-container .file-preview,
.gradio-container label,
.gradio-container .upload-container,
.gradio-container [data-testid="block-info"],
.gradio-container [data-testid="file"],
.gradio-container [data-testid="file"] > div,
.gradio-container [data-testid="file"] .wrap,
.gradio-container [data-testid="file"] .upload-container,
.gradio-container [data-testid="textbox"],
.gradio-container [data-testid="button"],
.gradio-container [data-testid="radio"],
.gradio-container [data-testid="checkbox"],
.gradio-container [data-testid="dropdown"],
.gradio-container [data-testid="dataframe"],
.gradio-container [data-testid="html"] {
    max-width: 100% !important;
    box-sizing: border-box !important;
}

.gradio-container [data-testid="file"] {
    min-height: 160px !important;
    overflow: hidden !important;
}

.gradio-container [data-testid="file"] .upload-container {
    min-height: 140px !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    overflow: hidden !important;
}

#rab-upload-file,
#rab-upload-file > div,
#rab-upload-file .wrap,
#rab-upload-file .upload-container {
    min-height: 160px !important;
    height: 160px !important;
    overflow: hidden !important;
    border: 1px solid var(--rab-border) !important;
    border-radius: 12px !important;
    box-shadow: none !important;
}

#rab-upload-file [data-testid="file"],
#rab-upload-file .upload-container {
    background: #ffffff !important;
}

#rab-upload-file,
#rab-upload-file > div,
#rab-upload-file .wrap {
    background: #ffffff !important;
}

#keyword-upload-file,
#keyword-upload-file > div,
#keyword-upload-file .wrap,
#keyword-upload-file .upload-container,
#keyword-upload-file [data-testid="file"] {
    min-height: 96px !important;
    height: 96px !important;
    background: #ffffff !important;
    border: 1px solid var(--rab-border) !important;
    border-radius: 12px !important;
    box-shadow: none !important;
    overflow: hidden !important;
}

#keyword-upload-file button,
#keyword-upload-file [data-testid="button"],
#keyword-upload-file label {
    background: transparent !important;
    color: #2C3947 !important;
    border: 1px solid var(--rab-border) !important;
    box-shadow: none !important;
}

#keyword-upload-file svg,
#keyword-upload-file [data-testid="upload-icon"],
#keyword-upload-file .upload-icon,
#keyword-upload-file p,
#keyword-upload-file span:not(.file-name) {
    display: none !important;
}

#keyword-upload-file {
    position: relative;
    font-size: 0 !important;
    color: transparent !important;
}

#keyword-upload-file * {
    font-size: 0 !important;
    color: transparent !important;
}

#keyword-upload-file::after {
    content: "Click to Upload";
    position: absolute;
    left: 50%;
    top: 50%;
    transform: translate(-50%, -50%);
    color: #547A95;
    font-size: 16px;
    font-weight: 800;
    pointer-events: none;
}

.gradio-container button[aria-expanded]:focus,
.gradio-container button[aria-expanded]:focus-visible {
    outline: 2px solid rgba(84, 122, 149, 0.24) !important;
    outline-offset: 2px !important;
}

.gradio-container label:has(input[type="radio"]:checked),
.gradio-container label:has(input[type="checkbox"]:checked),
.gradio-container label:has(input[type="radio"]:checked) *,
.gradio-container label:has(input[type="checkbox"]:checked) * {
    color: #ffffff !important;
}

#rab-upload-file button,
#rab-upload-file [data-testid="button"],
#rab-upload-file label {
    display: none !important;
}

#rab-upload-file svg,
#rab-upload-file [data-testid="upload-icon"],
#rab-upload-file .upload-icon {
    display: none !important;
}

#rab-upload-file {
    font-size: 0 !important;
    color: transparent !important;
}

#rab-upload-file button,
#rab-upload-file label {
    font-size: 14px !important;
}

#rab-upload-file span,
#rab-upload-file p,
#rab-upload-file div:not(:has(button)) {
    color: transparent !important;
}

#rab-upload-file::after {
    content: "Click to Upload";
    position: absolute;
    left: 50%;
    top: 48%;
    transform: translate(-50%, -50%);
    color: #547A95;
    font-size: 18px;
    font-weight: 800;
    pointer-events: none;
}

#rab-upload-file:focus,
#rab-upload-file:focus-within,
#rab-upload-file *:focus,
#rab-upload-file *:focus-visible {
    outline: 3px solid rgba(194, 165, 109, 0.32) !important;
    outline-offset: 2px !important;
    border-color: rgba(194, 165, 109, 0.45) !important;
}

.gradio-container [role="tabpanel"] > div,
.gradio-container [role="tabpanel"] > div > div {
    width: 100% !important;
    max-width: 100% !important;
}

.gradio-container .tabs > div {
    width: 100% !important;
}

.gradio-container button,
.gradio-container textarea,
.gradio-container input {
    max-width: 100% !important;
}

/* UI/UX Pro Max final system: Trust & Authority for finance review tools */
:root {
    --rab-primary: #2C3947;
    --rab-secondary: #547A95;
    --rab-accent: #C2A56D;
    --rab-accent-strong: #A9894E;
    --rab-accent-soft: #F3EBD8;
    --rab-text: #2C3947;
    --rab-heading: #2C3947;
    --rab-muted: #547A95;
    --rab-border: #C9D4DF;
    --rab-soft: #E8EDF2;
    --rab-paper: #ffffff;
    --rab-danger: #ef4444;
}

html,
body,
.gradio-container {
    min-width: 0 !important;
    overflow-x: hidden !important;
}

body,
.gradio-container {
    background:
        radial-gradient(circle at 12% 0%, rgba(194, 165, 109, 0.16), transparent 30%),
        linear-gradient(180deg, #E8EDF2 0%, #ffffff 46%, #E8EDF2 100%) !important;
    color: var(--rab-text) !important;
    font-family: "Fira Sans", Inter, system-ui, sans-serif !important;
}

.gradio-container {
    width: min(calc(100vw - 48px), 1320px) !important;
    max-width: 1320px !important;
    margin: 0 auto !important;
    padding: 42px 32px 64px !important;
}

.gradio-container > *,
#app-hero,
.tabs,
.tabs > div,
div[role="tablist"],
div[role="tabpanel"] {
    width: 100% !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
}

h1, h2, h3,
.prose h1, .prose h2, .prose h3,
#app-hero h1,
.keyword-name,
.finding-item,
.redaction-title,
.learning-title {
    color: var(--rab-heading) !important;
    font-family: "Fira Code", "Fira Sans", ui-monospace, monospace !important;
    letter-spacing: -0.01em !important;
}

#app-hero {
    position: relative;
    overflow: visible !important;
    display: grid;
    gap: 14px;
    background: transparent !important;
    border-radius: 0 !important;
    border: 0 !important;
    box-shadow: none !important;
    padding: 12px 0 24px !important;
    margin: 0 !important;
}

#app-hero::after {
    content: none;
}

#app-hero h1 {
    color: #2C3947 !important;
    font-size: clamp(34px, 4vw, 56px) !important;
    line-height: 1.18 !important;
    max-width: 920px;
    margin: 0 !important;
    overflow: visible !important;
}

#app-hero p,
#app-hero .hero-copy {
    display: none !important;
}

#app-hero .version-pill,
#app-hero .hero-badge {
    width: fit-content;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    color: #2C3947 !important;
    background: transparent !important;
    border: 0 !important;
    border-radius: 999px !important;
    padding: 0 !important;
    font-family: "Fira Sans", Inter, sans-serif !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    line-height: 1.5 !important;
}

.tabs {
    overflow: visible !important;
    background: rgba(255, 255, 255, 0.92) !important;
    border: 1px solid var(--rab-border) !important;
    border-radius: 16px !important;
    box-shadow: 0 10px 28px rgba(15, 23, 42, 0.05) !important;
    backdrop-filter: none;
}

div[role="tablist"] {
    display: grid !important;
    grid-template-columns: repeat(5, minmax(0, 1fr)) !important;
    align-items: center !important;
    gap: 0 !important;
    min-height: 58px !important;
    height: 58px !important;
    padding: 0 !important;
    overflow: hidden !important;
    background: #E8EDF2 !important;
    border-bottom: 1px solid var(--rab-border) !important;
    border-radius: 16px 16px 0 0 !important;
}

button[role="tab"],
.tab-nav button,
.tabs button {
    display: inline-flex !important;
    align-items: center !important;
    min-height: 58px !important;
    height: 58px !important;
    padding: 0 12px !important;
    margin: 0 !important;
    line-height: 1 !important;
    justify-content: center !important;
    white-space: nowrap !important;
    border: 0 !important;
    border-radius: 0 !important;
    background: transparent !important;
    color: #475569 !important;
    font-family: "Fira Sans", Inter, sans-serif !important;
    font-weight: 700 !important;
    overflow: hidden !important;
    box-shadow: none !important;
    transform: none !important;
    transition: background 160ms ease, color 160ms ease !important;
}

button[role="tab"]::before,
button[role="tab"]::after {
    display: none !important;
}

button[role="tab"]:focus,
button[role="tab"]:focus-visible {
    outline: 3px solid rgba(194, 165, 109, 0.35) !important;
    outline-offset: 2px !important;
}

button[role="tab"][aria-selected="true"],
.tab-nav button.selected,
.tabs button.selected {
    background: #2C3947 !important;
    color: #ffffff !important;
    box-shadow: none !important;
}

div[role="tabpanel"] {
    min-height: 760px !important;
    padding: 28px !important;
    background: #ffffff !important;
}

.gradio-container button:not([role="tab"]),
button.primary,
.primary > button,
button[variant="primary"] {
    min-height: 46px !important;
    border: 0 !important;
    border-radius: 10px !important;
    background: #547A95 !important;
    color: #ffffff !important;
    box-shadow: none !important;
    font-family: "Fira Sans", Inter, sans-serif !important;
    font-weight: 800 !important;
    transition: transform 160ms ease, box-shadow 160ms ease, background 160ms ease !important;
}

.gradio-container button:not([role="tab"]):hover {
    background: #2C3947 !important;
    color: #ffffff !important;
    transform: none !important;
}

.gradio-container button:not([role="tab"]) *,
button.primary *,
.primary > button *,
button[variant="primary"] * {
    color: inherit !important;
}

.gradio-container button[aria-expanded],
.gradio-container button[aria-expanded]:hover {
    min-height: 50px !important;
    background: transparent !important;
    color: #2C3947 !important;
    border: 0 !important;
    box-shadow: none !important;
}

footer,
.footer,
#footer,
[data-testid="footer"],
.built-with,
.api-info,
a[href*="gradio.app"],
button[aria-label="Use via API"] {
    display: none !important;
}

input, textarea, select,
.gradio-container .wrap,
.gradio-container .block,
.gradio-container .form {
    border-color: var(--rab-border) !important;
    border-radius: 16px !important;
    font-family: "Fira Sans", Inter, sans-serif !important;
}

input, textarea, select {
    background: #ffffff !important;
    color: var(--rab-text) !important;
    min-height: 46px !important;
}

label,
.label-wrap span {
    color: #334155 !important;
    font-family: "Fira Sans", Inter, sans-serif !important;
    font-weight: 700 !important;
}

.finding-card,
.keyword-card,
.learning-card,
.redaction-result,
.simple-note,
.keyword-search-panel,
.redaction-search,
.findings-toolbar {
    border: 1px solid var(--rab-border) !important;
    border-radius: 12px !important;
    background: #ffffff !important;
    box-shadow: none !important;
    transform: none !important;
}

.keyword-card {
    position: relative !important;
    padding-bottom: 62px !important;
}

.redaction-search,
.redaction-search > div,
.redaction-search .block,
.redaction-search .form,
.redaction-search .wrap {
    border: 0 !important;
    box-shadow: none !important;
    background: transparent !important;
    padding: 0 !important;
}

.redaction-search textarea,
.redaction-search input {
    border: 1px solid var(--rab-border) !important;
    box-shadow: none !important;
    background: #ffffff !important;
}

.keyword-card:hover,
.finding-card:hover,
.learning-card:hover {
    border-color: rgba(194, 165, 109, 0.65) !important;
    box-shadow: none !important;
    transform: none !important;
}

.keyword-card:nth-child(6n + 1) { background: #fff7e6 !important; border-color: #e5c98f !important; }
.keyword-card:nth-child(6n + 2) { background: #eef6fb !important; border-color: #b8cfdf !important; }
.keyword-card:nth-child(6n + 3) { background: #f3f7ef !important; border-color: #c9dabf !important; }
.keyword-card:nth-child(6n + 4) { background: #f7f0f6 !important; border-color: #dcc4d8 !important; }
.keyword-card:nth-child(6n + 5) { background: #f4f0e8 !important; border-color: #d6c6a8 !important; }
.keyword-card:nth-child(6n + 6) { background: #eff3f8 !important; border-color: #c7d4e4 !important; }

.keyword-chip,
.alias-chip,
.severity-badge,
.confidence-pill,
.learning-count {
    border: 1px solid rgba(34, 197, 94, 0.18) !important;
    border-radius: 999px !important;
    background: #F3EBD8 !important;
    color: #2C3947 !important;
    font-family: "Fira Sans", Inter, sans-serif !important;
    font-weight: 800 !important;
}

.keyword-delete-btn {
    position: absolute;
    right: 18px;
    bottom: 18px;
    width: auto;
    min-height: 34px !important;
    margin: 0;
    padding: 7px 10px !important;
    background: transparent !important;
    color: #991b1b !important;
    box-shadow: none !important;
    border: 1px solid #fecaca !important;
    border-radius: 999px !important;
    font-size: 12px !important;
    line-height: 1.2 !important;
}

.keyword-delete-btn:hover {
    background: #ef4444 !important;
    color: #ffffff !important;
}

.versioning-panel {
    padding: 16px 18px 18px !important;
    border: 1px solid var(--rab-border);
    border-radius: 12px;
    background: #ffffff;
    line-height: 1.65;
    overflow: visible;
}

.versioning-panel h3,
.versioning-panel p {
    margin: 0 0 10px !important;
    line-height: 1.65 !important;
}

.redaction-score,
.learning-number,
.confidence-score {
    color: #0f172a !important;
    font-family: "Fira Code", ui-monospace, monospace !important;
}

.dataframe,
.table-wrap,
table {
    border-radius: 16px !important;
    border-color: var(--rab-border) !important;
}

#app-hero {
    padding-left: 14px !important;
}

#app-hero .version-pill {
    padding-left: 6px !important;
}

#rab-upload-file {
    position: relative !important;
    cursor: pointer !important;
}

#rab-upload-file button,
#rab-upload-file [data-testid="button"],
#rab-upload-file label {
    display: block !important;
    position: absolute !important;
    inset: 0 !important;
    width: 100% !important;
    height: 100% !important;
    min-height: 160px !important;
    padding: 0 !important;
    margin: 0 !important;
    opacity: 0 !important;
    background: transparent !important;
    border: 0 !important;
    cursor: pointer !important;
    z-index: 2 !important;
}

#rab-upload-file::after {
    z-index: 3 !important;
}

#keyword-upload-accordion,
#keyword-upload-accordion > div,
#keyword-upload-accordion .block,
#keyword-upload-accordion .wrap,
#keyword-upload-accordion .form {
    border: 0 !important;
    box-shadow: none !important;
    background: transparent !important;
}

#keyword-upload-accordion button,
#keyword-upload-accordion button:hover,
#keyword-upload-accordion button:focus,
#keyword-upload-accordion button:focus-visible,
#keyword-upload-accordion *:focus,
#keyword-upload-accordion *:focus-visible {
    min-height: 48px !important;
    background: transparent !important;
    color: #2C3947 !important;
    border: 0 !important;
    box-shadow: none !important;
    outline: 0 !important;
}

#keyword-upload-file {
    position: relative !important;
    min-height: 86px !important;
    height: 86px !important;
    border: 1px solid var(--rab-border) !important;
    border-radius: 12px !important;
    background: #ffffff !important;
    box-shadow: none !important;
    overflow: hidden !important;
    cursor: pointer !important;
}

#keyword-upload-file > div,
#keyword-upload-file .wrap,
#keyword-upload-file .upload-container,
#keyword-upload-file [data-testid="file"] {
    min-height: 84px !important;
    height: 84px !important;
    border: 0 !important;
    border-radius: 0 !important;
    background: transparent !important;
    box-shadow: none !important;
    outline: 0 !important;
}

#keyword-upload-file button,
#keyword-upload-file [data-testid="button"],
#keyword-upload-file label {
    display: block !important;
    position: absolute !important;
    inset: 0 !important;
    width: 100% !important;
    height: 100% !important;
    min-height: 84px !important;
    padding: 0 !important;
    margin: 0 !important;
    opacity: 0 !important;
    background: transparent !important;
    border: 0 !important;
    cursor: pointer !important;
    z-index: 2 !important;
}

#keyword-upload-file::after {
    z-index: 3 !important;
}

#keyword-upload-file::before {
    content: "";
    position: absolute;
    inset: 0;
    border: 1px solid var(--rab-border);
    border-radius: 12px;
    pointer-events: none;
    z-index: 1;
}

th {
    background: #f8fafc !important;
    color: #0f172a !important;
}

@media (max-width: 920px) {
    .gradio-container {
        width: min(100%, calc(100vw - 24px)) !important;
        padding: 18px 12px 36px !important;
    }

    div[role="tablist"] {
        display: flex !important;
        overflow-x: auto !important;
    }

    button[role="tab"] {
        flex: 0 0 auto !important;
        min-width: 160px !important;
    }

    div[role="tabpanel"] {
        padding: 18px !important;
    }
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
            f"<button class='keyword-delete-btn' type='button' data-delete-keyword-id='{keyword_id}'>Hapus keyword</button>"
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


DELETE_KEYWORD_JS = """
() => {
  if (window.__rabNacDeleteKeywordBound) return;
  window.__rabNacDeleteKeywordBound = true;
  document.addEventListener("click", (event) => {
    const button = event.target.closest("[data-delete-keyword-id]");
    if (!button) return;
    event.preventDefault();
    const keywordId = button.getAttribute("data-delete-keyword-id");
    const input = document.querySelector("#delete-kw-id textarea, #delete-kw-id input");
    const trigger = document.querySelector("#delete-kw-trigger button");
    if (!input || !trigger) return;
    input.value = keywordId;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    trigger.click();
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
  <div class="version-pill">Versi {APP_VERSION}</div>
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
        simple_add_btn.click(add_keyword_simple_ui, simple_kw, [kw_msg, keyword_cards, delete_kw_id])
        import_btn.click(import_keywords_simple_ui, import_file, [kw_msg, keyword_cards, delete_kw_id])
        delete_kw_btn.click(delete_keyword_by_id_ui, delete_kw_id, [kw_msg, keyword_cards])
        export_kw_btn.click(export_keywords_ui, outputs=export_kw_file)
        learn_btn.click(learning_ui, outputs=learning_html)
        save_set.click(save_simple_settings_ui, [review_mode, semantic_mode, ocr_mode], settings_msg)
        reset_db.click(reset_db_ui, outputs=settings_msg)
        demo.load(None, None, None, js=DELETE_KEYWORD_JS)
    return demo


if __name__ == "__main__":
    app().queue().launch()

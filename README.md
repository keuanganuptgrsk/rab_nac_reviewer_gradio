---
title: RAB NAC Reviewer Copilot
colorFrom: blue
colorTo: gray
sdk: gradio
sdk_version: 6.14.0
app_file: app.py
python_version: 3.11
pinned: false
license: mit
---

# RAB NAC Reviewer Copilot

RAB NAC Reviewer Copilot adalah aplikasi web internal berbasis Gradio untuk membantu review awal dokumen RAB Indonesia dan mendeteksi **Potensi NAC** dalam konteks komponen subsidi PLN/Kemenkeu/PMK.

> Hasil deteksi adalah bantuan awal untuk review internal. Keputusan final tetap harus divalidasi oleh reviewer yang memahami PMK, kebijakan internal, dan konteks pekerjaan.

## Tujuan

Aplikasi ini membantu reviewer meningkatkan kejelasan, kepatuhan, auditability, dan redaksi yang selaras dengan PMK/kebijakan internal. Aplikasi ini tidak dirancang untuk menyembunyikan, menyamarkan, memanipulasi, atau menghindari NAC aktual.

## Kemampuan Utama

- Versi saat ini: `v0.6.5 - Finance Friendly Review Flow`.
- Upload RAB Excel, CSV, PDF, dan image.
- Excel/CSV adalah alur paling andal.
- PDF digital diekstrak dengan PyMuPDF.
- PDF scan/image OCR bersifat optional best-effort.
- Deteksi hybrid: exact keyword, sinonim, fuzzy matching, semantic similarity opsional, allowable competitor, exception, dan feedback.
- SQLite lokal untuk database keyword, exception, settings, feedback, dan backup.
- Export Excel multi-sheet untuk audit/review.

## Versioning

Setiap penambahan fitur wajib memperbarui:

- `modules/version.py` untuk versi aktif, judul rilis, dan keterangan singkat yang tampil di UI.
- `CHANGELOG.md` untuk catatan versi, judul, tanggal, dan detail perubahan.

Format versi yang digunakan: `MAJOR.MINOR.PATCH`.

## Instalasi Lokal

```bash
cd rab_nac_reviewer_gradio
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Buka URL lokal yang ditampilkan oleh Gradio.

## Deploy ke Hugging Face Spaces

Recommended Hugging Face Space setting:

- SDK: Gradio
- Hardware: CPU Basic
- Python app file: `app.py`
- Persistent data: gunakan Export/Import Backup secara rutin
- Best input format: Excel/CSV RAB
- OCR: optional best-effort, tergantung dependency yang tersedia

Upload isi folder ini ke Space Gradio. `data/app.db` akan dibuat otomatis saat startup.

## Database NAC

Tab **Database NAC** menyediakan:

- Tambah keyword NAC
- Tambah sinonim
- Tambah allowable keyword
- Tambah exception
- Import/export database keyword Excel

Seed keyword bersifat demo dan wajib divalidasi dengan PMK/kebijakan internal.

Template import:

- `category`
- `keyword`
- `synonyms` dipisah dengan titik koma
- `description`
- `reference`
- `severity`
- `status`
- `notes`

Status yang disarankan: `active`, `inactive`, `deprecated`, `needs_review`. Data tidak hard delete secara default.

## Cara Confidence Bekerja

Formula ringkas:

```text
final_score =
0.25 * exact_or_synonym_score
+ 0.20 * fuzzy_score
+ 0.30 * semantic_score
+ 0.10 * severity_score
+ 0.10 * feedback_adjustment
- 0.20 * allowable_competitor_penalty
- exception_penalty
```

Kategori confidence:

- 0-24%: Sangat rendah
- 25-44%: Rendah
- 45-64%: Sedang
- 65-84%: Tinggi
- 85-100%: Sangat tinggi

Jika sinyal NAC dan allowable sama-sama kuat, hasil diarahkan ke **Perlu Review Manual**.

## Feedback Learning

Tidak ada retraining kompleks. Feedback digunakan untuk:

- Menurunkan skor pola yang sering ditandai `Not NAC`
- Menaikkan skor pola yang sering ditandai `Correct NAC`
- Menyarankan exception untuk false positive berulang
- Menyarankan keyword/sinonim baru dari feedback reviewer
- Menampilkan kandidat sinonim dari semantic/fuzzy match model Hugging Face untuk disetujui reviewer

Kandidat sinonim dari model tidak otomatis masuk database. Reviewer harus menekan **Approve Suggested Synonym** pada baris hasil review yang sesuai. Jika sinonim atau exception ditambahkan, database langsung dipakai pada review berikutnya. Semantic embeddings dibangun secara lazy saat review.

## OCR

OCR tidak diwajibkan agar aplikasi tetap ringan di Hugging Face Spaces CPU Basic. Default requirements tidak memasang OCR berat.

Opsional:

```bash
pip install easyocr
pip install paddleocr paddlepaddle
pip install pytesseract
```

Untuk Tesseract, binary Tesseract OCR juga harus tersedia di sistem.

Jika OCR gagal, aplikasi tetap berjalan untuk Excel/CSV dan PDF berbasis teks.

## Export dan Backup

Tab **Export Excel** menyediakan:

- Export hasil review Excel dengan sheet Summary, Findings, Suggestions, Feedback Log, Keyword Matches, dan NAC Keyword Database Snapshot.
- Export feedback logs.
- Export SQLite database backup.
- Import SQLite database backup.

Pada hosting gratis, filesystem tidak selalu ideal untuk catatan bisnis permanen. Export backup secara rutin.

## Workflow Rekomendasi

1. Upload RAB Excel/CSV.
2. Pastikan kolom teks, volume, satuan, harga satuan, dan total benar.
3. Jalankan review.
4. Filter `Sedang`, `Tinggi`, dan `Sangat tinggi`.
5. Beri feedback untuk false positive, false negative, dan item yang perlu manual review.
6. Tambahkan sinonim/exception jika pola berulang.
7. Export Excel untuk dokumentasi review.

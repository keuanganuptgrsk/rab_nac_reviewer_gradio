# Changelog

Semua penambahan fitur harus dicatat dengan format: versi, judul, tanggal, dan keterangan.

## v0.5.0 - Tabtion-Inspired UI Refresh - 2026-05-09

- Mengadopsi arah visual modern seperti design system Tabtion dengan warna `#DCFCE7`, `#0B61A2`, `#15803D`, `#FFFFFF`, dan `#1D1D1F`.
- Menambahkan font role Manrope untuk heading dan Geist untuk body melalui CSS import.
- Mengubah header aplikasi menjadi hero ringkas berisi versi, judul, keterangan rilis, dan disclaimer.
- Merapikan tab, tombol, input, dan tabel agar lebih konsisten sebagai workspace review internal.

## v0.4.1 - Audit-Safe Redaction Suggestions - 2026-05-09

- Sugesti redaksi diubah menjadi format `Usulan redaksi` yang lebih singkat dan langsung bisa direview.
- Menambahkan template saran spesifik untuk narasumber, konsumsi/catering, hadiah/souvenir, pakaian/perlengkapan non-teknis, dan transportasi.
- Saran tidak mengganti substansi biaya menjadi istilah generik yang dapat menyamarkan NAC; jika tidak allowable, sistem menyarankan pemisahan dari komponen allowable.

## v0.4.0 - Auto Review RAB Upload - 2026-05-09

- Upload Excel RAB langsung menjalankan NAC review otomatis.
- Parser membaca `Nama Pekerjaan` sebagai `Judul` dan item dari sheet `RAB` / `REALISASI`.
- Tabel hasil otomatis menampilkan hanya confidence `Sedang`, `Tinggi`, dan `Sangat tinggi`.
- Kolom hasil ringkas mencakup `Judul`, `Item per RAB`, `Kategori`, `Confidence Level`, dan `Sugesti Perubahan Redaksi`.
- Seed demo ditambah untuk pola biaya seperti konsumsi, catering, prasmanan, snack, minuman, doorprize, oleh-oleh, cinderamata, dan fee narasumber.

## v0.3.0 - Model Suggested Synonym Approval - 2026-05-09

- Model semantic/fuzzy dapat mengusulkan kandidat sinonim dari hasil review.
- Reviewer harus menekan `Approve Suggested Synonym` sebelum kandidat masuk database.
- Kandidat sinonim yang disetujui tercatat di feedback learning.

## v0.2.0 - Hugging Face Space Deployment - 2026-05-09

- Menambahkan metadata Hugging Face Spaces dan konfigurasi Python 3.11.
- Menyelaraskan dependency Gradio untuk deployment Space.
- Menambahkan file template Excel keyword demo ke repository.

## v0.1.0 - Initial Internal Reviewer App - 2026-05-09

- Aplikasi Gradio dengan tab Upload RAB, Review Hasil, Database NAC, Feedback & Learning, Export Excel, dan Settings.
- SQLite lokal untuk keyword NAC, sinonim, allowable keyword, exception, feedback, settings, dan backup.
- Deteksi hybrid exact, synonym, fuzzy, semantic opsional, allowable competitor, exception, dan feedback adjustment.
- Export Excel multi-sheet untuk hasil review dan database snapshot.

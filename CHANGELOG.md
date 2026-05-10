# Changelog

Semua penambahan fitur harus dicatat dengan format: versi, judul, tanggal, dan keterangan.

## v0.7.4 - Contrast and Cleanup Pass - 2026-05-10

- Memberi ruang aman pada header dan versioning agar teks tidak terpotong di sisi kiri/atas/bawah.
- Menghapus subtitle hero `Internal review workspace...` dari tampilan awal.
- Memperbaiki kontras teks pada tombol/dropdown gelap agar tetap terbaca.
- Membuat upload Excel keyword NAC cukup outline tanpa background gelap dan tinggi lebih proporsional.
- Memindahkan tombol hapus keyword ke pojok kanan bawah card dengan ukuran lebih kecil.
- Menghapus fitur backup data dari tab Settings.

## v0.7.3 - Minimal Header and Soft Keyword Cards - 2026-05-10

- Menghapus card/gradient pada header dan menyembunyikan subtitle hero agar judul tidak terpotong.
- Menghilangkan bar coklat pada area upload dengan menyembunyikan label/button internal file upload.
- Menyederhanakan tab `Analisa Redaksi NAC` agar hanya memakai satu border input tanpa lapisan abu-abu.
- Mengubah keyword cards menjadi soft colorful palette dan mengecilkan tombol hapus agar tidak mendominasi konten.

## v0.7.2 - Clean Tabs and Surfaces - 2026-05-10

- Membuat tab rata dalam bar tanpa pill yang menjorok sehingga tidak ada elemen yang menghalangi tab/konten.
- Mengurangi card shadow, radius, dan hover transform agar tidak menghalangi tulisan.
- Memadatkan upload area menjadi 160px dan menyederhanakan surface-nya.
- Menyembunyikan footer Gradio agar tidak mengganggu tampilan aplikasi.

## v0.7.1 - ColorHunt Palette and Tab Polish - 2026-05-10

- Mengganti color palette ke ColorHunt: `#E8EDF2`, `#2C3947`, `#547A95`, `#C2A56D`.
- Memperbesar tinggi tab list dan tab button agar label tidak terpotong dan tetap center secara vertikal.
- Menyederhanakan area upload dengan menyembunyikan teks `Drop File Here` dan hanya menonjolkan `Click to Upload`.
- Mengganti focus ring upload dari hijau ke aksen gold sesuai palette baru.

## v0.7.0 - Trust Authority UI Rebuild - 2026-05-10

- Menggunakan skill `ui-ux-pro-max` untuk memilih arah desain `Trust & Authority` yang lebih cocok untuk aplikasi finance/compliance.
- Membangun ulang visual system menjadi slate/navy + green status, Fira Sans/Fira Code, glass cards halus, dan hierarchy yang lebih profesional.
- Menambahkan hero copy yang menjelaskan fungsi workspace internal.
- Memperkuat stabilitas width tab di desktop dan mobile dengan shell 1320px, grid tab desktop, dan horizontal tab scroll mobile.

## v0.6.9 - Perspective Gradient UI - 2026-05-10

- Mengubah UI mengikuti referensi Gradient Pro/Perspective: hero hijau gelap, aksen neon green, pill navigation, dan glass cards.
- Memperkuat aturan width untuk tab panel, row, column, group, dan komponen Gradio agar tidak mengecil saat berpindah tab.
- Menjaga card keyword dan hasil review tetap full-width di dalam shell aplikasi.

## v0.6.8 - Gradient UI and Stable Width - 2026-05-10

- Mengubah UI dari doodle ke TypeUI Gradient-inspired style: purple-to-pink gradients, Montserrat, Space Grotesk, glass cards, dan gradient actions.
- Mengunci shell, tab list, dan tab panel ke width yang konsisten agar layout tidak berubah saat pindah tab.
- Memperbaiki hapus keyword langsung dari card memakai event delegation global agar tetap berjalan setelah card dirender ulang.
- Mencatat bahwa TypeUI CLI tidak dapat dijalankan di environment lokal karena `npx` tidak tersedia, sehingga skill diterapkan dari dokumentasi TypeUI Gradient.

## v0.6.7 - Doodle UI and Card Delete - 2026-05-10

- Mengubah UI ke arah doodle/sketch-inspired dengan border tebal, shadow offset, dan font Delius Swash Caps.
- Menstabilkan layout antar tab dengan panel height dan sticky tab bar.
- Menghapus form/dropdown `Hapus keyword NAC` dari Database NAC.
- Menambahkan tombol hapus langsung di setiap card keyword NAC.
- Memprioritaskan PaddleOCR sebagai OCR optional saat mode OCR `auto`, dengan fallback ke EasyOCR dan Tesseract.

## v0.6.6 - Cleaner Upload and Keyword Delete - 2026-05-09

- Menambahkan fitur hapus keyword NAC dari daftar aktif dengan soft-delete `inactive`.
- Menyembunyikan kontrol sort dan tombol export sebelum RAB dianalisa.
- Mengganti output file export menjadi tiga tombol download saja tanpa panel file besar.
- Menambahkan tinggi minimum panel tab agar posisi halaman lebih stabil saat berpindah tab.
- Menyembunyikan keyword inactive dari kartu Database NAC.

## v0.6.5 - Finance Friendly Review Flow - 2026-05-09

- Menghapus tab `Review Hasil` dan `Export Excel` dari navigasi utama.
- Menambahkan tab `Analisa Redaksi NAC` untuk mengecek satu kalimat redaksi seperti search box.
- Memindahkan tambah keyword dan upload Excel keyword ke bagian atas tab `Database NAC`.
- Menyederhanakan tambah keyword agar user cukup mengisi keyword, sementara kategori, confidence dasar, catatan, dan kandidat sinonim dipilih otomatis.
- Menyederhanakan tab `Settings` menjadi mode review, semantic/parafrasa otomatis, OCR, backup, dan reset demo database.
- Mengubah `Feedback & Learning` menjadi dashboard kartu interaktif.

## v0.6.4 - Simple NAC Keyword Cards - 2026-05-09

- Menyederhanakan tab `Database NAC` untuk user pemula.
- Mengganti tabel database mentah menjadi search panel dan kartu keyword NAC.
- Menampilkan sinonim/parafrasa sebagai chip yang dipakai sistem untuk review.
- Menambahkan form sederhana untuk tambah keyword NAC tanpa pengaturan teknis seperti ID, weight, dan status.
- Menambahkan import Excel keyword dari tab yang sama serta export database keyword.

## v0.6.3 - RAB Persekot Parser Support - 2026-05-09

- Memperbaiki parser Excel untuk layout `RAB Persekot` yang memakai `No` di kolom pertama dan `Uraian Kegiatan` di kolom kedua.
- Membaca sub-item tanpa nomor seperti `Makan & Minum` dan `Snack` sebagai item RAB tersendiri.
- Mencegah section/group sebelumnya menempel ke item bernomor berikutnya agar mengurangi false positive.
- Menambahkan keyword demo yang perlu validasi internal: `uang saku`, `honorarium`, `pulsa petugas`, dan `bantuan transport eksternal`.

## v0.6.2 - Finding Sort and Export Pack - 2026-05-09

- Menyederhanakan arah sort menjadi `Tinggi ke rendah` dan `Rendah ke tinggi`.
- Menambahkan tabel seluruh material RAB tanpa filter confidence.
- Tabel seluruh material berisi Row, Nama Material, Kategori NAC, Confidence %, dan Kategori Confidence Level.
- Menambahkan export PDF untuk rangkuman potensi NAC yang perlu direview.
- Menambahkan export PDF dan Excel untuk tabel seluruh material RAB.

## v0.6.1 - Sortable NAC Finding Cards - 2026-05-09

- Menambahkan kontrol sort untuk kartu temuan NAC.
- Reviewer dapat mengurutkan berdasarkan `Row` atau `Confidence`.
- Kartu hasil akan di-render ulang tanpa menjalankan ulang review.

## v0.6.0 - Interactive NAC Finding Cards - 2026-05-09

- Mengganti tampilan hasil review utama dari tabel menjadi kartu interaktif.
- Kartu temuan hanya menampilkan informasi inti: Row, Item RAB, Confidence, dan Confidence Level.
- Menambahkan warna berbeda untuk confidence `Sedang`, `Tinggi`, dan `Sangat tinggi`.
- Mengoptimalkan tampilan untuk asumsi satu Excel per analisa agar hasil lebih mudah discan.

## v0.5.6 - Responsive Review Button - 2026-05-09

- Menambahkan status proses langsung setelah tombol `Run NAC Review` ditekan.
- Tombol dinonaktifkan sementara dan label berubah menjadi `Sedang memproses...` selama review berjalan.
- Semantic matching dimatikan secara default agar review Excel lebih cepat di Hugging Face CPU Basic.
- Semantic matching tetap bisa dinyalakan manual dari Settings untuk review lanjutan.

## v0.5.5 - Stable Tab Bar - 2026-05-09

- Menjaga tab navigasi tetap horizontal dan tidak berubah menjadi menu titik tiga.
- Menambahkan horizontal scroll untuk tab saat ruang layar sempit.
- Mengurangi pergeseran layout pada area navigasi utama.

## v0.5.4 - Manual Review Start - 2026-05-09

- Pada sesi awal tab Upload hanya menampilkan upload file dan tombol `Run NAC Review`.
- Tabel `Preview / Extracted Rows` disembunyikan dari UI utama.
- Tabel hasil review disembunyikan sampai tombol `Run NAC Review` ditekan.
- Auto-run saat upload dimatikan agar pengguna punya kontrol eksplisit kapan review dijalankan.

## v0.5.3 - Minimal Hero Header - 2026-05-09

- Header utama hanya menampilkan nomor versi dan judul aplikasi.
- Menghilangkan keterangan rilis dan disclaimer panjang dari area hero agar tampilan lebih bersih.
- Detail versi tetap tersedia di Settings dan `CHANGELOG.md`.

## v0.5.2 - Simplified Upload Flow - 2026-05-09

- Menyembunyikan kontrol mapping kolom teknis pada tab Upload RAB.
- Alur utama menjadi lebih sederhana: upload file, lihat preview, lalu hasil review otomatis tampil.
- Logika auto-detect kolom tetap berjalan di backend untuk mendukung review dan export.

## v0.5.1 - Fluently-Inspired UI Refresh - 2026-05-09

- Mengganti arah visual dari Tabtion-inspired ke Fluently-inspired sesuai design tokens yang diberikan.
- Menggunakan warna primary `#0B1220`, secondary `#475569`, accent/link `#1665D6`, background `#FFFFFF`, dan heading biru.
- Mengganti font role heading/body menjadi DM Sans.
- Memperbarui hero, tab, tombol, input, dan tabel agar terasa modern, high-energy, dan tetap cocok untuk workflow review RAB.

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

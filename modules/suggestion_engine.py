def generate_suggestion(text, keyword="", category="", confidence=0, allowable_score=0, exception_hit=False):
    text_l = (text or "").lower()
    risk = "Potensi NAC perlu ditinjau dalam konteks PMK/kebijakan internal."
    if exception_hit or ("konsumsi" in text_l and ("bahan bakar" in text_l or "genset" in text_l)):
        return (
            "Kemungkinan false positive pada konteks konsumsi bahan bakar. "
            "Perlakukan sebagai konteks teknis dan pertimbangkan exception bila pola ini berulang."
        )
    if "perjalanan" in text_l or "transport" in text_l:
        return (
            "Klarifikasi bahwa biaya perjalanan/transportasi terkait langsung dengan pekerjaan teknis. "
            "Tambahkan lokasi, surat tugas/work order, tanggal inspeksi, output laporan, dan unit penanggung jawab."
        )
    if "konsumsi" in text_l or "jamuan" in text_l or "snack" in text_l or "coffee" in text_l:
        return (
            "Review apakah biaya konsumsi/jamuan dapat diperhitungkan. Jika terkait langsung dengan kegiatan teknis, "
            "jelaskan tujuan, output, peserta, tanggal/lokasi, dan dasar pendukung. Jika tidak, pisahkan dari komponen allowable."
        )
    if confidence >= 45:
        return (
            f"{risk} Perjelas ruang lingkup teknis, dasar aturan internal/PMK, referensi pekerjaan, lokasi, output/deliverable, "
            "dan pisahkan komponen allowable dan non-allowable bila bercampur. Minta dokumen pendukung dan review manual."
        )
    if allowable_score > 50:
        return "Konteks terlihat lebih dekat ke aktivitas teknis/allowable; tetap simpan dokumentasi pendukung untuk audit."
    return "Tidak ada saran redaksi khusus. Lakukan review manual bila konteks pekerjaan belum jelas."


def recommended_action(confidence, allowable_score=0):
    if confidence >= 65 and allowable_score >= 60:
        return "Perlu Review Manual"
    if confidence >= 85:
        return "Potensi NAC tinggi - validasi reviewer dan pisahkan komponen bila perlu"
    if confidence >= 45:
        return "Perlu Review Manual"
    return "Monitor / dokumentasikan konteks"


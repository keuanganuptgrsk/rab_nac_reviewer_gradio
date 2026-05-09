APP_VERSION = "0.4.0"
APP_RELEASE_TITLE = "Auto Review RAB Upload"
APP_RELEASE_NOTES = (
    "Upload Excel RAB otomatis membaca judul dan item per RAB, menjalankan NAC review, "
    "serta menampilkan temuan confidence Sedang hingga Sangat tinggi dalam tabel ringkas."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"


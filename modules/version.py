APP_VERSION = "0.4.1"
APP_RELEASE_TITLE = "Audit-Safe Redaction Suggestions"
APP_RELEASE_NOTES = (
    "Sugesti perubahan redaksi dibuat lebih singkat, spesifik per kategori, dan tetap audit-able "
    "tanpa menyamarkan substansi biaya yang berpotensi NAC."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

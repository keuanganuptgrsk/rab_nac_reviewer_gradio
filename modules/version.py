APP_VERSION = "0.6.7"
APP_RELEASE_TITLE = "Doodle UI and Card Delete"
APP_RELEASE_NOTES = (
    "UI diganti ke gaya doodle, layout tab dibuat lebih konsisten, keyword bisa dihapus langsung dari card, dan PaddleOCR diprioritaskan sebagai OCR optional."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

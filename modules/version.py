APP_VERSION = "0.7.9"
APP_RELEASE_TITLE = "Tesseract OCR Enablement"
APP_RELEASE_NOTES = (
    "OCR PDF scan diaktifkan melalui pytesseract dan packages.txt Tesseract untuk Hugging Face Spaces."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

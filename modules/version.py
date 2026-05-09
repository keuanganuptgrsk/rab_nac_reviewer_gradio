APP_VERSION = "0.6.1"
APP_RELEASE_TITLE = "Sortable NAC Finding Cards"
APP_RELEASE_NOTES = (
    "Kartu temuan potensi NAC kini bisa diurutkan berdasarkan Row atau Confidence agar reviewer lebih cepat memprioritaskan analisa."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

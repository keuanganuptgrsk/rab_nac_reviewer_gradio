APP_VERSION = "0.6.0"
APP_RELEASE_TITLE = "Interactive NAC Finding Cards"
APP_RELEASE_NOTES = (
    "Temuan potensi NAC kini ditampilkan sebagai kartu interaktif, bukan tabel, dengan fokus pada Row, Item RAB, Confidence, dan Confidence Level."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

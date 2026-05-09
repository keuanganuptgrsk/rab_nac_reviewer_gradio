APP_VERSION = "0.5.5"
APP_RELEASE_TITLE = "Stable Tab Bar"
APP_RELEASE_NOTES = (
    "Tab navigasi dibuat stabil agar tidak berubah menjadi menu titik tiga; jika layar sempit, tab tetap horizontal dan dapat discroll."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

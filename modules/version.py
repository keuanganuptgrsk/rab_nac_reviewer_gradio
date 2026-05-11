APP_VERSION = "0.8.3"
APP_RELEASE_TITLE = "Export Button Chain Fix"
APP_RELEASE_NOTES = (
    "Tombol export tidak lagi otomatis berjalan saat hasil review berubah; export hanya diproses saat tombol diklik."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

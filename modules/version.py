APP_VERSION = "0.5.2"
APP_RELEASE_TITLE = "Simplified Upload Flow"
APP_RELEASE_NOTES = (
    "Kontrol mapping kolom upload disembunyikan agar alur review lebih sederhana: upload file, "
    "preview data, lalu hasil NAC review otomatis langsung tampil."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

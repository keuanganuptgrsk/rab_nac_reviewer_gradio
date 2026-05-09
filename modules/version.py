APP_VERSION = "0.6.3"
APP_RELEASE_TITLE = "RAB Persekot Parser Support"
APP_RELEASE_NOTES = (
    "Parser Excel kini membaca layout RAB Persekot dengan sub-item tanpa nomor, menjaga section agar tidak memicu false positive lintas item, dan menambah keyword demo operasional/personel."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

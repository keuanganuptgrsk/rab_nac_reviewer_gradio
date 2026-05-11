APP_VERSION = "0.8.2"
APP_RELEASE_TITLE = "Run Review Output Chain Fix"
APP_RELEASE_NOTES = (
    "Output Run NAC Review distabilkan agar hanya mengupdate card summary, tabel HTML, state hasil, status, sort panel, dan export panel."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

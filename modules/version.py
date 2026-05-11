APP_VERSION = "0.8.4"
APP_RELEASE_TITLE = "Visible Review Output Fix"
APP_RELEASE_NOTES = (
    "Output review dibuat dalam satu panel eksplisit dengan card summary dan tabel DataFrame agar hasil langsung terlihat setelah Run NAC Review."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

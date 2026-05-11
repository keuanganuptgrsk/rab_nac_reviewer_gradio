APP_VERSION = "0.8.0"
APP_RELEASE_TITLE = "Review Output UX Audit"
APP_RELEASE_NOTES = (
    "Output Run NAC Review diubah menjadi ringkasan kartu dan tabel HTML seluruh material yang stabil serta bisa disort langsung."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

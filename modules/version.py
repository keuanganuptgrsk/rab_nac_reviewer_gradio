APP_VERSION = "0.5.6"
APP_RELEASE_TITLE = "Responsive Review Button"
APP_RELEASE_NOTES = (
    "Tombol Run NAC Review sekarang langsung menampilkan status proses, sementara semantic matching dimatikan secara default agar review lebih cepat di hosting gratis."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

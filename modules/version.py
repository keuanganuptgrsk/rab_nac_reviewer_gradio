APP_VERSION = "0.5.3"
APP_RELEASE_TITLE = "Minimal Hero Header"
APP_RELEASE_NOTES = (
    "Header utama dibuat lebih ringkas dengan hanya menampilkan nomor versi dan judul aplikasi, "
    "tanpa keterangan rilis panjang di halaman utama."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

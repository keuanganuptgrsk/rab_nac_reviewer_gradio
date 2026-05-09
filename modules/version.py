APP_VERSION = "0.5.4"
APP_RELEASE_TITLE = "Manual Review Start"
APP_RELEASE_NOTES = (
    "Tab Upload dibuat lebih bersih pada sesi awal: tabel preview dan hasil disembunyikan sampai pengguna menekan Run NAC Review."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

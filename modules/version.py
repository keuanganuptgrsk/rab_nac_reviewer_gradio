APP_VERSION = "0.7.7"
APP_RELEASE_TITLE = "Version Safe Padding"
APP_RELEASE_NOTES = (
    "Label versi diberi spacer dan padding ekstra agar huruf awal tidak terpotong di pojok kiri atas."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

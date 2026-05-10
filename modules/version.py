APP_VERSION = "0.7.2"
APP_RELEASE_TITLE = "Clean Tabs and Surfaces"
APP_RELEASE_NOTES = (
    "Tab dibuat rata tanpa overlay, card/shadow dikurangi agar tulisan tidak terhalang, upload area dipadatkan, dan footer Gradio disembunyikan."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

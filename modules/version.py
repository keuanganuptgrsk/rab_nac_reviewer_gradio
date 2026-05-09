APP_VERSION = "0.6.5"
APP_RELEASE_TITLE = "Finance Friendly Review Flow"
APP_RELEASE_NOTES = (
    "Navigasi dipangkas, Database NAC cukup input keyword/upload Excel, Settings disederhanakan, Learning Dashboard dibuat berbasis kartu, dan tab Analisa Redaksi NAC ditambahkan."
)


def version_banner():
    return f"Versi {APP_VERSION} - {APP_RELEASE_TITLE}: {APP_RELEASE_NOTES}"

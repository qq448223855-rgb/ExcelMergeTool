# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


PROJECT_ROOT = Path(SPECPATH).parent
APP_BASENAME = "Excel合并工具V1.0.1"

EXCLUDED_BINARY_PATHS = {
    "PySide6/QtNetwork.abi3.so",
    "PySide6/QtDBus.abi3.so",
}
EXCLUDED_BINARY_PREFIXES = (
    "PySide6/Qt/lib/QtNetwork.framework/",
    "PySide6/Qt/lib/QtSvg.framework/",
    "PySide6/Qt/plugins/generic/",
    "PySide6/Qt/plugins/iconengines/",
    "PySide6/Qt/plugins/imageformats/",
    "PySide6/Qt/plugins/networkinformation/",
    "PySide6/Qt/plugins/tls/",
)
EXCLUDED_PLATFORM_PLUGINS = {
    "PySide6/Qt/plugins/platforms/libqminimal.dylib",
    "PySide6/Qt/plugins/platforms/libqoffscreen.dylib",
}
EXCLUDED_DATA_PATHS = {
    "QtNetwork",
    "QtSvg",
}
EXCLUDED_DATA_PREFIXES = (
    "PySide6/Qt/lib/QtNetwork.framework/",
    "PySide6/Qt/lib/QtSvg.framework/",
)


def keep_binary(entry):
    destination = entry[0].replace("\\", "/")
    if destination in EXCLUDED_BINARY_PATHS:
        return False
    if destination in EXCLUDED_PLATFORM_PLUGINS:
        return False
    return not destination.startswith(EXCLUDED_BINARY_PREFIXES)


def keep_data(entry):
    destination = entry[0].replace("\\", "/")
    if destination in EXCLUDED_DATA_PATHS:
        return False
    if destination.startswith(EXCLUDED_DATA_PREFIXES):
        return False

    translations_prefix = "PySide6/Qt/translations/"
    if not destination.startswith(translations_prefix):
        return True

    filename = destination.rsplit("/", 1)[-1]
    return filename.startswith("qtbase_")


a = Analysis(
    [str(PROJECT_ROOT / "main.py")],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "PySide6.QtNetwork",
        "PySide6.QtDBus",
        "PySide6.QtSvg",
    ],
    noarchive=False,
    optimize=1,
)
a.binaries = [entry for entry in a.binaries if keep_binary(entry)]
a.datas = [entry for entry in a.datas if keep_data(entry)]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_BASENAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch="arm64",
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name=APP_BASENAME,
)

app = BUNDLE(
    coll,
    name=f"{APP_BASENAME}.app",
    icon=None,
    bundle_identifier="com.huang.excelmerger",
    version="1.0.1",
    info_plist={
        "CFBundleDisplayName": APP_BASENAME,
        "CFBundleName": APP_BASENAME,
        "CFBundleShortVersionString": "1.0.1",
        "CFBundleVersion": "2",
        "LSMinimumSystemVersion": "11.0",
    },
)

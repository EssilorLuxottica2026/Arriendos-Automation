import os
from pathlib import Path

import webview


project_root = Path(SPECPATH)
webview_root = Path(webview.__file__).resolve().parent
debug_build = os.environ.get("ARRIENDOS_BUILD_DEBUG") == "1"

datas = [
    (str(project_root / "templates"), "templates"),
    (str(project_root / "static"), "static"),
]

for filename in (
    "CONTROL_ARRI_ADMON.xlsx",
    "Contratos_con_condiciones.xlsx",
    "PRORATEO.xlsx",
    "Cuadro_de_distribucion.xls",
    "FACTURAS_CONTABILIZADAS.xlsx",
    "Arriendos_Macro.xlsm",
):
    source = project_root / "data" / "input" / filename
    if source.exists():
        datas.append((str(source), "data/input"))

a = Analysis(
    ["desktop_app.py"],
    pathex=[str(project_root / "src")],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "webview.platforms.winforms",
        "webview.platforms.edgechromium",
    ],
    hookspath=[str(webview_root / "__pyinstaller")],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "charset_normalizer",
        "tkinter",
        "matplotlib",
        "IPython",
        "jupyter",
        "numpy.f2py",
        "pdfminer",
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ArriendosAutomationDebug" if debug_build else "ArriendosAutomation",
    debug="all" if debug_build else False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=debug_build,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    icon=str(project_root / "assets" / "arriendos_app_icon.ico"),
    codesign_identity=None,
    entitlements_file=None,
)

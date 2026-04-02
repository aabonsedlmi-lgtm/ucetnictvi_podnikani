# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_dir = Path(__file__).resolve().parent
run_script = project_dir / "run_fakturace.py"
icon_file = project_dir / "fakturace_icon.ico"

hidden_imports = [
    "flask",
    "openpyxl",
    "openpyxl.utils.datetime",
    "qrcode",
    "reportlab",
    "reportlab.lib",
    "reportlab.pdfgen",
    "pypdf",
    "pypdfium2",
    "rapidocr_onnxruntime",
]

a = Analysis(
    [str(run_script)],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="FakturaceStudio",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_file) if icon_file.exists() else None,
)

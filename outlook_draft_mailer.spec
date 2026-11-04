# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_dir = Path(SPECPATH)

a = Analysis(
    [str(project_dir / "mailer_gui.py")],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[],
    hiddenimports=[
        "openpyxl",
        "pythoncom",
        "pywintypes",
        "win32timezone",
        "win32com.client",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Outlook Draft Mailer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Outlook Draft Mailer",
)

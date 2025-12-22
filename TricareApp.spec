# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
from pathlib import Path
import sys
import streamlit as st

# Bundle Streamlit metadata and hidden imports so importlib.metadata can find them.
_st_datas, _st_bins, _st_hidden = collect_all("streamlit")

# Include Streamlit static assets explicitly (avoid dev server on port 3000).
_st_root = Path(st.__file__).parent
_st_static = _st_root / "static"
_extra_datas = [(str(_st_static), "streamlit/static")] if _st_static.exists() else []

# Ensure the main app script is available after extraction (_MEIPASS).
_app_file = Path("app.py")
_processor_file = Path("processor.py")
for _f in (_app_file, _processor_file):
    if _f.exists():
        _extra_datas.append((_f.as_posix(), "."))

_hidden = list(_st_hidden) + ["processor"]

a = Analysis(
    ['launch.py'],
    # Use current working directory as base; __file__ may be undefined when spec is run.
    pathex=[str(Path(sys.argv[0]).resolve().parent)],
    binaries=_st_bins,
    datas=_st_datas + _extra_datas,
    hiddenimports=_hidden,
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
    name='TricareApp',
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
)

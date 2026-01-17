# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

import sys
from PyInstaller.utils.hooks import collect_submodules

# Project base path
pathex = ['.']

# Conservative excludes: modules not used by this project
excludes = [
    'numpy', 'scipy', 'matplotlib', 'pandas', 'Crypto', 'win32com',
    'tkinter.test', 'test', 'unittest', 'distutils', 'pkg_resources',
]

# Some DLLs/APIs better to avoid UPX (PyInstaller may still override)
upx_exclude = [
    'VCRUNTIME140.dll', 'VCRUNTIME140_1.dll'
]

a = Analysis(
    ['Alldesk.py'],
    pathex=pathex,
    binaries=[],
    datas=[('Alldesk.xlsx', '.'), ('exe', 'exe')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, block_cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Alldesk',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=upx_exclude,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='EVERDURA_black.ico',
)

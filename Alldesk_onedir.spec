# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for building a one-dir (non-onefile) bundle of Alldesk.

Usage:
  pyinstaller Alldesk_onedir.spec

This spec will:
 - include `Alldesk.xlsx` next to the executable
 - copy the entire `exe/` folder into the bundled `dist/Alldesk/` under `exe/`
 - add common hidden-imports used by this project (adjust if needed)
"""
import os
import glob
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT

block_cipher = None

HERE = os.path.abspath(os.getcwd())

# collect all files under exe/ as binaries to be placed into dist/Alldesk/exe/
exe_dir = os.path.join(HERE, 'exe')
binaries = []
if os.path.isdir(exe_dir):
    for f in glob.glob(os.path.join(exe_dir, '*')):
        if os.path.isfile(f):
            # (src, destdir)
            binaries.append((f, 'exe'))

# include Alldesk.xlsx next to the executable (if present)
datas = []
xlsx = os.path.join(HERE, 'Alldesk.xlsx')
if os.path.isfile(xlsx):
    datas.append((xlsx, '.'))

# Add hidden imports commonly needed for pywin32/com automation and pywinauto
hidden_imports = [
    'comtypes',
    'pywinauto',
    'win32timezone',
]

# Exclude common heavy or test modules to reduce bundle size (adjust as needed)
excludes = [
    'tests',
    'unittest',
    'email',
    'pkg_resources',
    # 'importlib_metadata' removed due to PyInstaller hook aliasing
]

a = Analysis([
    os.path.join(HERE, 'Alldesk.py')
],
    pathex=[HERE],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden_imports,
    excludes=excludes,
    hookspath=[],
    runtime_hooks=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='Alldesk',
          icon=os.path.join(HERE, 'EVERDURA_black.ico'),
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
          console=False)

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=False,
               # 某些系統 DLL 與 python 內核檔不宜被 UPX 壓縮，列為例外
               upx_exclude=['python3.dll', 'vcruntime140.dll', 'msvcp140.dll'],
               name='Alldesk')
# 如果 PyInstaller 在 _internal 放了一份 Alldesk.xlsx，簡化處理：
# - 若 dist 根目錄已有同名檔，刪除 _internal 的副本
# - 否則用 shutil.move 將檔案移回 dist 根目錄
internal_xlsx = os.path.join(HERE, 'dist', 'Alldesk', '_internal', 'Alldesk.xlsx')
target_xlsx = os.path.join(HERE, 'dist', 'Alldesk', 'Alldesk.xlsx')
if os.path.exists(internal_xlsx):
    try:
        import shutil
        if os.path.exists(target_xlsx):
            os.remove(internal_xlsx)
        else:
            shutil.move(internal_xlsx, target_xlsx)
    except Exception:
        # 非致命，忽略錯誤並繼續完成 build
        pass

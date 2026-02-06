# -*- mode: python ; coding: utf-8 -*-
import os
import sys

# 動態檢查圖示檔案
# Windows: lioil.ico，macOS: lioil.icns
datas_list = [('exe', 'exe')]

# 根據平台設定應用程式名稱
app_name = 'Alldesk-macos' if sys.platform == 'darwin' else 'Alldesk'

# 根據運行平台選擇圖示
if sys.platform == 'darwin':  # macOS
    if os.path.exists('lioil.icns'):
        datas_list.insert(0, ('lioil.icns', '.'))
else:  # Windows
    if os.path.exists('lioil.ico'):
        datas_list.insert(0, ('lioil.ico', '.'))

# 設定 EXE 圖示參數
icon_param = []
if sys.platform == 'darwin':
    if os.path.exists('lioil.icns'):
        icon_param = ['lioil.icns']
else:
    if os.path.exists('lioil.ico'):
        icon_param = ['lioil.ico']

a = Analysis(
    ['Alldesk.py'],
    pathex=[],
    binaries=[],
    datas=datas_list,
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_param,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=app_name,
)

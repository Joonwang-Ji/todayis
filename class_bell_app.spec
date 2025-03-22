# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['class_bell_app.py'],
    pathex=[],
    binaries=[('.venv\\Lib\\site-packages\\PyQt5\\Qt5\\bin\\Qt5Widgets.dll', 'PyQt5\\Qt5\\bin'), ('.venv\\Lib\\site-packages\\PyQt5\\Qt5\\bin\\Qt5Core.dll', 'PyQt5\\Qt5\\bin'), ('.venv\\Lib\\site-packages\\PyQt5\\Qt5\\bin\\Qt5Gui.dll', 'PyQt5\\Qt5\\bin')],
    datas=[('start_bell.wav', '.'), ('end_bell.wav', '.'), ('todayis_qr.jpeg', '.'), ('start_icon.png', '.'), ('end_icon.png', '.'), ('icon.ico', '.'), ('.venv\\Lib\\site-packages\\PyQt5\\Qt5\\plugins\\platforms', 'PyQt5\\Qt5\\plugins\\platforms')],
    hiddenimports=['PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets'],
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
    name='class_bell_app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='class_bell_app',
)

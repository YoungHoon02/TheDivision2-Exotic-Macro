# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['d2macro.py'],
    pathex=[],
    binaries=[],
    datas=[('images', 'images')],
    hiddenimports=['keyboard'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'cv2', 'PyQt5', 'pygame', 'scipy', 'matplotlib', 'setuptools', 'pkg_resources', 'pyautogui', 'pydirectinput'],
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
    name='D2Macro',
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
    icon='images/div2_icon.ico',
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=True,
)

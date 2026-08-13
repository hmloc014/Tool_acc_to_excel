# -*- mode: python -*-

block_cipher = None

a = Analysis(
    ["acc_to_excel_app.py"],
    pathex=[],
    binaries=[],
    datas=[("images/icon6.ico", ".")],
    hiddenimports=["pythoncom", "pywintypes", "win32com.client"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name="Acc to Excel",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon="images/icon6.ico",
)

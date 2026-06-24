# -*- mode: python ; coding: utf-8 -*-
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('config.json', '.'), ('splash.png', '.'), ('icon.png', '.')],
    hiddenimports=[
        'customtkinter', 'selenium', 'win32com', 'win32com.client',
        'pywintypes', 'sqlite3', 'PIL', 'PIL.Image', 'PIL.ImageTk',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [],
    exclude_binaries=True,
    name='PregateMail',
    debug=False,
    console=False,
    upx=True,
    icon='icon.ico',
)
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, name='PregateMail')

# -*- mode: python ; coding: utf-8 -*-
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('config.json', '.'), ('splash.png', '.'), ('icon.png', '.')],
    hiddenimports=[
        'customtkinter',
        'win32com', 'win32com.client', 'win32gui', 'win32con',
        'win32process', 'win32api', 'pywintypes',
        'sqlite3', 'glob',
        'PIL', 'PIL.Image', 'PIL.ImageTk',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=['selenium','pyautogui','pyperclip'],
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

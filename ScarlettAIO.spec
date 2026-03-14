# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller build spec for Scarlett AIO.
#
# Build with:
#   pip install pyinstaller
#   pyinstaller ScarlettAIO.spec
#
# Output: dist\Scarlett AIO\
#   Zip that folder and distribute it.  Users unzip, put their config.json
#   inside, and double-click "Scarlett AIO.exe".
#
import os
import customtkinter

block_cipher = None

# customtkinter ships theme JSON files and widget images that must be bundled.
ctk_path = os.path.dirname(customtkinter.__file__)

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # customtkinter data files (themes, images)
        (ctk_path, 'customtkinter'),
        # Icons bundled next to the exe
        ('AIO.ico',  '.'),
        ('scar.ico', '.'),
        # Config template so users have a reference on a fresh install
        ('config.example.json', '.'),
    ],
    hiddenimports=[
        # pdfminer / pdfplumber dynamic imports
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.converter',
        'pdfminer.pdfdocument',
        'pdfminer.pdfinterp',
        'pdfminer.pdfpage',
        'pdfminer.pdfparser',
        # PIL is pulled in by customtkinter
        'PIL._tkinter_finder',
        # pandas optional engines
        'openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Scarlett AIO',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # No console window (windowed app)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='AIO.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Scarlett AIO',
)

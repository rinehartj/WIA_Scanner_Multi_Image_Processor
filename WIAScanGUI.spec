# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['WIAScanGUI.py'],  # Replace with your actual script name
    pathex=[],
    binaries=[],
    datas=[
        ('tools/exiftool.exe', 'tools'),
        ('tools/exiftool_files', 'tools/exiftool_files'),
    ],
    hiddenimports=[
        'win32com.client',
        'pythoncom',
        'pywintypes',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='WIAScannerGUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window for GUI
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add your icon path here if you have one: icon='icon.ico'
)
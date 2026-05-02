from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

added_files = [
    ("app/config",   "app/config"),
    ("app/examples", "app/examples"),
    ("app/assembler/requirements.txt", "app/assembler"),
]

a = Analysis(
    ["app/launcher.py"],
    pathex=["."],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        "watchdog.observers.inotify",
        "watchdog.observers.fsevents",
        "watchdog.observers.kqueue",
        "watchdog.observers.polling",
        "bibtexparser",
        "mammoth",
        "PIL._tkinter_finder",
    ],
    hookspath=[],
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
    name="EditorialSystem",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="EditorialSystem",
)

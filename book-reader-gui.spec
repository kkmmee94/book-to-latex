# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files


a = Analysis(
    ['book_to_latex_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/tessdata/ara.traineddata', 'assets/tessdata'),
        ('assets/tessdata/LICENSE', 'assets/tessdata'),
        ('assets/tessdata/README.md', 'assets/tessdata'),
    ] + collect_data_files('certifi'),
    hiddenimports=['PIL.Image', 'pytesseract'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # These packages are installed only for the optional Streamlit interface.
    # Keeping them out of the desktop bundle saves tens of megabytes. The
    # pytesseract/PyMuPDF features used by the desktop app do not require them.
    excludes=['altair', 'numpy', 'pandas', 'pyarrow', 'streamlit'],
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
    name='book-reader',
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
)

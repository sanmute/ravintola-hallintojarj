# -*- mode: python ; coding: utf-8 -*-
# Rakennus:  pyinstaller ruokalistasuunnittelija.spec
# Tulos:     dist/Ruokalistasuunnittelija/Ruokalistasuunnittelija.exe

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = [
    ('templates', 'templates'),
    ('static', 'static'),
    ('templates_src', 'templates_src'),   # brändätty menupohja
]
datas += collect_data_files('pymupdf')
datas += collect_data_files('reportlab')

hiddenimports = (
    collect_submodules('pymupdf')
    + collect_submodules('openpyxl')
    + ['webview.platforms.edgechromium', 'webview.platforms.winforms']
)

a = Analysis(
    ['desktop.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    excludes=['tkinter', 'matplotlib', 'scipy', 'numpy.f2py'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Ruokalistasuunnittelija',
    console=False,           # ei mustaa komentoikkunaa
    icon=None,               # lisää 'kesti.ico' kun logo on saatavilla
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name='Ruokalistasuunnittelija',
)

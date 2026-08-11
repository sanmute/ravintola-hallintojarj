# -*- mode: python ; coding: utf-8 -*-
# Rakennus:  pyinstaller ruokalistasuunnittelija.spec
# Tulos:     dist/Ruokalistasuunnittelija.exe  (YKSI tiedosto — onefile-tila)
#
# Onefile-tilassa koko sovellus (Python-runtime + kaikki datat, ml. ~150-200 MB
# käännösmalli) puretaan käyttäjän Temp-kansioon JOKAISEN käynnistyksen
# yhteydessä, ei vain ensimmäisellä kerralla — tämä on tarkoituksellinen
# kompromissi verrattuna aiempaan onedir-tilaan:
#   + Yksi tiedosto voidaan kopioida minne tahansa ilman erillistä
#     _internal-kansiota (onedir-tilassa pelkän .exe:n kopiointi/raahaus
#     työpöydälle irrottaa sen tästä kansiosta eikä sovellus käynnisty).
#   - Käynnistys on hitaampi (puretaan joka kerta uudelleen).
#   - Jotkin virustorjuntaohjelmat merkitsevät onefile-PyInstaller-paketteja
#     herkemmin vääriksi positiivisiksi kuin onedir-paketteja.

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = [
    ('templates', 'templates'),
    ('templates_src', 'templates_src'),   # brändätty menupohja
    ('models', 'models'),                 # paikallinen en->fi käännösmalli (ctranslate2)
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
    a.binaries,
    a.datas,
    [],
    name='Ruokalistasuunnittelija',
    console=False,           # ei mustaa komentoikkunaa
    icon='kesti.ico',
)

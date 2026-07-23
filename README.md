# Ruokalistasuunnittelija

Flask-pohjainen työpöytäsovellus ravintolan ruokalistojen suunnitteluun. Ajetaan
paikallisesti selaimessa, paketoitu PyInstallerilla yhdeksi .exe-tiedostoksi
keittiön koneelle — ei vaadi Python-asennusta käyttökoneella.

## Mitä sovellus tekee

- **Luo ruokalistoja** vuoden mittaisiksi (52 viikkoa) tai yksittäisille
  vuodenajoille (talvi, kevät, kesä, syksy). Jokaiselle arkipäivälle (ma–pe)
  valitaan pääruoka, keitto ja salaatti.
- **Kiertää reseptit järkevästi**: keitot eivät toistu alle 2 viikon välein,
  salaatit kiertävät vuorotellen, ja sama pääruokakategoria voi toistua
  korkeintaan kahdesti viikossa.
- **"Vain uudet reseptit" -tila**: ruokalistan voi generoida käyttäen vain
  reseptiuudistuksen jälkeen lisättyjä reseptejä.
- **Reseptien muokkaus jälkikäteen**: yksittäisen päivän ruoan voi vaihtaa,
  ja sovellus ehdottaa sopivia korvaavia reseptejä.
- **Resepteille voi asettaa "vain käsin valittava" -lipun**, jolloin
  automaattigeneraattori ei koskaan valitse niitä itse (esim. kastikkeet).
- **Reseptien poisto** poistaa reseptin myös aktiivisista ruokalistoista
  (kaskadina) ja kertoo, mitkä listat vaikuttuvat — käyttäjää kehotetaan
  generoimaan ne tarvittaessa uudelleen.
- **Reseptien haku verkosta**: skrapperi lukee schema.org/JSON-LD-tiedot
  tunnetuilta resepti­sivustoilta, mukaan lukien annosmäärän ("5 annosta"
  tms.), jos sivu sen ilmoittaa.
- **Ainesosanäkymässä** annosmäärää voi skaalata kertoimella (esim. ×2) —
  kerroin on vain näyttöä varten eikä tallennu reseptiin, ja sen voi
  nollata.
- **PoweResta-vienti**: valitut reseptit voi viedä Excel-tiedostoksi
  PoweResta-tuontia varten (yksi välilehti per resepti, määrät kiloina).
- **Käyttäjäroolit**: `admin`, `muokkaus` ja `katselu` — pääkäyttäjällä on
  pääsy mm. reseptien joukkopoistoon ja auditointiin.
- **Automaattiset varmuuskopiot** tietokannasta (`backup.py`,
  `nightly_backup.py` + ajastettu Windows-tehtävä).

## Käyttöönotto

1. Asenna riippuvuudet: `pip install -r requirements.txt`
2. Käynnistä sovellus: `python app.py` tai `KAYNNISTA.bat`
3. Avaa selaimessa osoite, jonka sovellus tulostaa konsoliin (oletuksena
   `http://localhost:5001`)

Valmiiksi paketoitu .exe rakennetaan `RAKENNA_EXE.bat`-skriptillä
(`ruokalistasuunnittelija.spec`-tiedoston mukaan). Käyttäjädata (tietokanta,
viennit, arviointijono) tallentuu ajon aikana kansioon
`%LOCALAPPDATA%\Ruokalistasuunnittelija\`, jotta se säilyy .exe:n
uudelleenrakentamisen yli.

Tarkemmat käyttöohjeet: [OHJEET.md](OHJEET.md) ja
[OHJEET_HALLINTA.md](OHJEET_HALLINTA.md).

## Koodirakenne

```
app.py                    Flask-sovellus (reitit, API)
auth.py                   Kirjautuminen ja roolit
meal_plan_db.py           Tietokantamallit (SQLite)
meal_plan_generator.py    Ruokalistan generointialgoritmi
meal_plan_modifier.py     Yksittäisten aterioiden vaihto
meal_plan_exporter.py     Excel-vienti
menu_generator.py         Viikkomenun kokoaminen (PDF/DOCX-pohjaa varten)
menu_pdf_generator.py     PDF-viikkomenu
poweresta_exporter.py     PoweResta-Excel-vienti
recipe_scraper_v2.py      Reseptien haku verkosta (JSON-LD)
recipe_classifier.py      Reseptien luokittelu auditointia varten
order_catalog.py          Tilausluettelo
backup.py / nightly_backup.py   Tietokannan varmuuskopiointi
desktop.py                PyInstaller-käynnistyspiste
templates/index.html      Käyttöliittymä (yksisivuinen, suomeksi)
```

## Teknologia

- Python 3.10+, Flask
- SQLite
- openpyxl (Excel-vienti/-tuonti)
- BeautifulSoup4 (reseptien haku)
- PyInstaller (jakelupaketti)

# Ruokalistasuunnittelija — Ravintolan hallintajärjestelmä

Tämä paketti lisää sovellukseen varmuuskopioinnin, lähiverkkokäytön sekä
Windows-työpöytäsovelluksen. Sovellus on yksinkäyttäjäjärjestelmä — ei
käyttäjätunnuksia, ei kirjautumista, kaikki toiminnot ovat aina käytössä.

---

## 1. Asennus (kehityskoneella)

Kopioi nämä tiedostot projektikansioon `app.py`:n viereen:

- `backup.py`
- `desktop.py`
- `ruokalistasuunnittelija.spec`
- `RAKENNA_EXE.bat`

Lisää sitten `app.py`-tiedostoon:

Importtien perään:

```python
from backup import init_backup
```

Heti `OUTPUT_DIR`-rivien jälkeen:

```python
init_backup(app, DB_PATH)
```

---

## 2. Varmuuskopiot

- Luo kopio: `POST /api/backup`
- Listaa kopiot: `GET /api/backup`
- Lataa kopio: `GET /api/backup/<nimi>`
- Palauta: `POST /api/backup/restore` — `{"name": "meal_plans_20260707_120000.db"}`

Kopiot tallentuvat `varmuuskopiot/`-kansioon aikaleimalla.
30 uusinta säilytetään, vanhemmat siivotaan automaattisesti.
Palautuksen yhteydessä nykytilasta otetaan aina turvakopio ensin.

---

## 3. Käyttö lähiverkossa (suositeltu tapa!)

Sovelluksen ei tarvitse olla asennettuna jokaiselle koneelle.
Riittää että **yksi kone** (esim. keittiön PC) ajaa sovellusta —
muut avaavat sen selaimella:

```
http://<koneen-nimi>:5001
```

Työpöytäsovelluksen ikkunan otsikossa näkyy valmiiksi osoite,
jonka voi jakaa muille.

Jos Windowsin palomuuri kysyy lupaa ensimmäisellä käynnistyksellä,
salli yhteys **yksityisissä verkoissa**.

---

## 4. Työpöytäsovelluksen rakentaminen (EXE)

Kehityskoneella:

1. Kaksoisklikkaa `RAKENNA_EXE.bat`
2. Odota — valmis sovellus syntyy kansioon
   `dist\Ruokalistasuunnittelija\`
3. Kopioi **koko kansio** kohdekoneelle (esim. muistitikulla tai verkkolevyltä)
4. Tee työpöydälle pikakuvake tiedostoon `Ruokalistasuunnittelija.exe`

Kohdekoneella **ei tarvita Pythonia, pipiä eikä järjestelmänvalvojan
oikeuksia** — kaikki on paketissa mukana.

Tietokanta (`meal_plans.db`) ja varmuuskopiot syntyvät sovelluskansion
sisään. Jos haluat tietokannan verkkolevylle, muuta `app.py`:n
`DB_PATH`-rivi osoittamaan sinne ennen paketointia.

---

## 5. Automaattinen käynnistys (valinnainen)

Palvelinkoneella: paina `Win+R`, kirjoita `shell:startup`, ja kopioi
pikakuvake avautuvaan kansioon. Sovellus käynnistyy aina kirjautumisen
yhteydessä.

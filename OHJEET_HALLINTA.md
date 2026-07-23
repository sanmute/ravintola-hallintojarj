# Ruokalistasuunnittelija — Ravintolan hallintajärjestelmä

Tämä paketti lisää sovellukseen käyttäjätunnukset ja roolit,
varmuuskopioinnin, lähiverkkokäytön sekä Windows-työpöytäsovelluksen.

---

## 1. Asennus (kehityskoneella)

Kopioi nämä tiedostot projektikansioon `app.py`:n viereen:

- `auth.py`
- `backup.py`
- `desktop.py`
- `ruokalistasuunnittelija.spec`
- `RAKENNA_EXE.bat`

Lisää sitten `app.py`-tiedostoon **kolme riviä**.

Importtien perään (rivin 20 tienoille):

```python
from auth import init_auth
from backup import init_backup
```

Heti `OUTPUT_DIR`-rivien jälkeen (rivin 27 tienoille):

```python
init_auth(app, DB_PATH)
init_backup(app, DB_PATH)
```

Siinä kaikki — kaikki nykyiset reitit ja ominaisuudet toimivat kuten ennenkin,
mutta nyt kirjautumisen takana.

---

## 2. Käyttäjät ja roolit

Ensimmäisellä käynnistyksellä luodaan tunnus:

| Tunnus | Salasana | Rooli |
|--------|----------|-------|
| admin  | vaihda123 | admin |

**Vaihda salasana heti ensimmäisen kirjautumisen jälkeen!**

Roolit:

- **admin** — kaikki oikeudet, käyttäjien hallinta, varmuuskopiot (esihenkilö)
- **muokkaus** — reseptien ja listojen luonti, muokkaus ja poisto (keittiövastaavat)
- **katselu** — vain katselu ja vientitiedostojen lataus (muu henkilökunta)

Käyttäjien hallinta (admin):

- Listaa: `GET /api/users`
- Lisää: `POST /api/users` — `{"username": "...", "password": "...", "role": "muokkaus"}`
- Poista: `DELETE /api/users/<id>`
- Oma salasananvaihto (kaikki): `POST /api/users/password` — `{"old": "...", "new": "..."}`

(Halutessasi rakennamme näille myöhemmin oman välilehden käyttöliittymään —
API on jo valmis.)

---

## 3. Varmuuskopiot (admin)

- Luo kopio: `POST /api/backup`
- Listaa kopiot: `GET /api/backup`
- Lataa kopio: `GET /api/backup/<nimi>`
- Palauta: `POST /api/backup/restore` — `{"name": "meal_plans_20260707_120000.db"}`

Kopiot tallentuvat `varmuuskopiot/`-kansioon aikaleimalla.
30 uusinta säilytetään, vanhemmat siivotaan automaattisesti.
Palautuksen yhteydessä nykytilasta otetaan aina turvakopio ensin.

---

## 4. Käyttö lähiverkossa (suositeltu tapa!)

Sovelluksen ei tarvitse olla asennettuna jokaiselle koneelle.
Riittää että **yksi kone** (esim. keittiön PC) ajaa sovellusta —
muut avaavat sen selaimella:

```
http://<koneen-nimi>:5001
```

Työpöytäsovelluksen ikkunan otsikossa näkyy valmiiksi osoite,
jonka voi jakaa muille. Jokainen kirjautuu omalla tunnuksellaan.

Jos Windowsin palomuuri kysyy lupaa ensimmäisellä käynnistyksellä,
salli yhteys **yksityisissä verkoissa**.

---

## 5. Työpöytäsovelluksen rakentaminen (EXE)

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

## 6. Automaattinen käynnistys (valinnainen)

Palvelinkoneella: paina `Win+R`, kirjoita `shell:startup`, ja kopioi
pikakuvake avautuvaan kansioon. Sovellus käynnistyy aina kirjautumisen
yhteydessä.

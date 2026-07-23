# Ruokalistasuunnittelija — Käyttöönotto-ohje

Toimiva versio: reseptit → ruokalista → Excel & Kespro-tilauspohja.

## Asennus Windowsille (5 min)

1. Pura ZIP haluamaasi kansioon (esim. `C:\Ruokalistat`)
2. Kaksoisklikkaa **seed_and_start.bat** → lisää 97 aloitusreseptiä (nykyisen ruokalistan ruoat)
3. Kaksoisklikkaa **KAYNNISTA.bat** → asentaa riippuvuudet ja käynnistää sovelluksen
4. Avaa selaimessa: **http://localhost:5001**

Vaatimus: Python 3.10+ asennettuna (sama kuin reseptimuuntimessa).

## Välilehdet

### 📖 Reseptit
- Tilastonäkymä: reseptit kausittain ja kategorioittain
- Lisää yksittäisiä reseptejä käsin (nimi + kausi + kategoria riittää)
- Muokkaa ja poista reseptejä

### 📥 Tuo reseptejä
Kolme tapaa, suositusjärjestyksessä:
1. **PoweResta-Excel-tuonti** — tuo reseptimuuntimen tuottamat Excelit suoraan.
   Organisaation omat ~450 reseptiä ovat paras reseptilähde: täysin omat,
   ruokavaliotiedot säilyvät, ei lupakysymyksiä.
2. **Liitä tekstinä** — kopioi resepti itse mistä tahansa (sivu, kirja, muistiinpano)
   ja liitä; nimi, raaka-aineet ja ohjeet tunnistetaan automaattisesti.
3. **Skannatut ruokalistataulukot (OCR)** — lataa valokuvat/skannaukset
   viikkoruokalistoista (viikonpäivät sarakkeina). Tekstintunnistus poimii
   ruoat automaattisesti: kuvan suunta tunnistetaan itse, taulukko jäsennetään
   päivittäin ja aterioittain, ja rivinvaihtojen katkomat nimet korjataan.
   Oletuksena vain pääateriat (Lounas + Päivällinen); perustarvikkeet kuten
   "Juomapaketti" ohitetaan aina. Useita kuvia voi ladata kerralla.
   **Vaatimus:** Tesseract OCR + suomen kielipaketti:
   - Windows: UB Mannheim -asennuspaketti (github.com/UB-Mannheim/tesseract),
     valitse asennuksessa "Finnish"
   - macOS: `brew install tesseract tesseract-lang`
   - Linux: `apt install tesseract-ocr tesseract-ocr-fin`
4. **Hae osoitteesta** — vain sivustoilla, joiden käyttöehdot sallivat
   automaattisen haun. ⚠️ Esim. K-Ruoka kieltää sen ehdoissaan — sieltä
   reseptit tuodaan tavoilla 1–3. Tarkista aina lähdesivuston ehdot ensin.

Kaikki kolme arvaavat kategorian ja kauden, ja tulokset menevät Tarkistusjonoon.

### ✅ Tarkistusjono
- Verkosta haetut reseptit odottavat täällä luokittelua
- Arvatut kausi/kategoria on esivalittu — tarkista ja korjaa tarvittaessa
- Poista turhat, "Vie hyväksytyt tietokantaan" → vain täysin luokitellut siirtyvät

### 📅 Ruokalistat
- Valitse kausi ja pituus (4 / 13 / 26 / 52 viikkoa)
- 19 ateriaa/viikko, sama resepti max 2× / 4 vk kierto, kategoriat tasapainossa
- **📥 Excel** → seinätaulumuotoinen ruokalista (värikoodattu)
- **📄 Viikkomenu** → löytyy nyt "Selaa & muokkaa" -välilehdeltä, jokaisen
  viikon kohdalta erikseen (ei 52 kerralla). Kaksi muotoa:
  - **PDF (oletus)** → käyttää ravintolan omaa ilmettä: Kesti-logo, taustakuva
    ja Setlementti-merkki tuodaan suoraan alkuperäisestä menupohjasta
    (templates_src/menu_template.pdf). Ma–Pe, keitto + pääruoka +
    ruokavaliokoodit (L, GL) automaattisesti, päivämäärät oikein.
  - **Word** → muokattava versio ilman grafiikkaa, jos tekstiä pitää hioa.
  Oman pohjan vaihto: korvaa templates_src/menu_template.pdf uudella
  PDF:llä ja poista templates_src/branded_background.png (se luodaan
  uudelleen automaattisesti). Hinnat/yhteystiedot: menu_pdf_generator.py.
- **🛒 Kespro** → viikkokohtainen tilauspohja Excelinä

### 👁️ Selaa & muokkaa
- Klikkaa mitä tahansa ateriaa → vaihda toiseen saman kategorian ruokaan
- Syy + tekijä kirjataan → muutosloki näkyy sivun alalaidassa
- Tämä on "lyhyellä varoitusajalla" -ominaisuus toimitusongelmia varten


## Kausiteemat (52 vk = 4 kautta)

Ruokalistat-välilehden oletusvalinta **"Koko vuosi — 4 kausiteemaa"**
luo 52 viikon suunnitelman, jossa viikot jaetaan kalenterin mukaan:

| Kausi | Viikot |
|---|---|
| Talvi | 1–9 ja 49–52 |
| Kevät | 10–22 |
| Kesä | 23–35 |
| Syksy | 36–48 |

Reseptien kausilogiikka:
- **Ympärivuotinen (kaikki)** — käytössä joka kaudella. Suurin osa
  perusruoista (lihapullat, lasagne, keitot) kuuluu tähän.
  Aloitusreseptit on merkitty näin.
- **Kausiresepti (talvi/kevät/kesä/syksy)** — käytössä vain omalla
  kaudellaan. Esim. parsakeitto→kevät, kantarellikeitto→syksy,
  grillattu lohi ja uudet perunat→kesä.

Kerääjä arvaa kauden raaka-aineista (parsa, raparperi → kevät;
kurpitsa, sieni, riista → syksy; uudet perunat, mansikka → kesä;
joulu, laskiainen → talvi). Jos selvää kausimerkkiä ei ole, kausi
jää tyhjäksi ja tarkistusjonossa valitaan yleensä "Ympärivuotinen".

Yhden kauden listan voi yhä luoda erikseen (esim. vain kevät, 13 vk).

## Reseptien hakeminen verkosta

**Huom: tarkista aina lähdesivuston käyttöehdot ennen automaattista hakua.**
Omien PoweResta-Excelien tuonti ja tekstin liittäminen eivät vaadi lupia.

Tekniikka: haku lukee schema.org Recipe -merkinnät (JSON-LD), joita
Ruoka.fi, K-Ruoka, Valio ja Kotikokki käyttävät — kestää sivustojen
ulkoasumuutokset. Haku odottaa 1 s/pyyntö (kohteliaisuus lähdesivustolle).

Komentorivi toimii edelleen (esim. suuriin eriin tai ajastettuihin ajoihin):

```
python recipe_scraper_v2.py https://www.ruoka.fi/reseptit/lohikeitto
python recipe_scraper_v2.py --listing https://www.ruoka.fi/reseptit/keitot 30
python recipe_scraper_v2.py --file urls.txt
```

## Kespro-tilauspohja

Tilauspohja sisältää sarakkeet: Kespro-tuotenumero, Tuote, Yksikkö,
Määrä, Toimituspäivä, Huomiot.

- Jos resepteille on syötetty raaka-aineet määrineen → pohja laskee
  viikon yhteismäärät automaattisesti
- Jos ei → pohja listaa viikon ruoat tilauksen tueksi

Kun esimies on selvittänyt Kesprolta integraatiovaihtoehdot
(API / SFTP / EDI), pohjan saa muunnettua suoraan heidän muotoonsa —
sarakerakenne on jo valmiiksi tilauksen mukainen.

## Tiedostot

| Tiedosto | Tarkoitus |
|---|---|
| app.py | Flask-sovellus (käyttöliittymä + API) |
| meal_plan_db.py | Tietokanta (SQLite, meal_plans.db) |
| meal_plan_generator.py | Ruokalista-algoritmi |
| meal_plan_exporter.py | Excel-vienti |
| meal_plan_modifier.py | Aterioiden vaihto + muutosloki |
| recipe_scraper_v2.py | Reseptikerääjä (JSON-LD) |
| seed_recipes.py | 97 aloitusreseptiä nykyiseltä listalta |
| KAYNNISTA.bat | Käynnistys |
| seed_and_start.bat | Ensimmäinen käyttökerta |

## Varmuuskopiointi

Kaikki data on yhdessä tiedostossa: **meal_plans.db**.
Kopioi se talteen säännöllisesti (esim. viikoittain).

## Tunnetut rajoitukset (vielä)

- Raaka-ainemäärät eivät tule kerääjältä jäsenneltynä — ne pitää
  syöttää käsin tai lisätä jäsennys myöhemmin (Kespro-pohja toimii
  silti ruokalistapohjaisena)
- Kerääjän kausiarvaus perustuu avainsanoihin — epävarmoissa tapauksissa kausi jää tyhjäksi ja vahvistetaan käsin
- Ravintoarvot eivät vielä mukana (JSON-LD sisältää ne usein —
  helppo lisätä seuraavassa vaiheessa)

"""
backup.py — Varmuuskopiot Ruokalistasuunnittelijalle.

Käyttöönotto app.py:ssä:

    from backup import init_backup
    init_backup(app, DB_PATH)

Reitit:
  POST /api/backup                     — luo aikaleimatun varmuuskopion
  GET  /api/backup                     — listaa varmuuskopiot
  GET  /api/backup/<name>              — lataa varmuuskopio tiedostona
  POST /api/backup/restore             — palauta varmuuskopio {"name": "..."}
  GET  /api/backup/download-package    — lataa koko sovellus + data ZIP-pakettina

Käyttää sqlite3:n omaa backup-API:a, joten kopio on eheä myös
kun sovellus on käytössä. Vanhat kopiot siivotaan automaattisesti
(oletuksena säilytetään 30 uusinta).

Yöllinen automaattinen varmuuskopiointi hoidetaan Windowsin Task
Schedulerilla, joka ajaa erillistä nightly_backup.py-skriptiä (ks.
nightly_backup.py ja nightly_backup.bat) — riippumatta siitä onko
tämä Flask-sovellus käynnissä. Sen luomat tiedostot noudattavat
samaa meal_plans_YYYYMMDD_HHMMSS.db-nimeämiskäytäntöä, joten ne
näkyvät automaattisesti tämän moduulin /api/backup-listauksessa.
"""

import os
import re
import sqlite3
import shutil
import tempfile
from datetime import datetime

from flask import jsonify, request, send_file

_NAME_RE = re.compile(r'^meal_plans_\d{8}_\d{6}\.db$')

_README_TXT = """RUOKALISTASUUNNITTELIJA — PIKAOHJE
===================================

Lataa sovellus ja pura ZIP-kansio. Kaksoisklikkaa Ruokalistasuunnittelija.exe.

Tarkemmat ohjeet:
- ASENNA.txt    — asennus ja käyttöönotto
- PALAUTUS.txt  — palautus varmuuskopioista
- TUKI.txt      — vianmääritys
"""

_ASENNA_TXT = """RUOKALISTASUUNNITTELIJA — ASENNUS JA KÄYTTÖÖNOTTO
=================================================

1. PAKKAUKSEN PURKAMINEN
   - Lataa tiedosto koneellesi
   - Pura ZIP-kansio (oikea-klikkaus → Pura kaikki)
   - Siirry kansioon "Ruokalistasuunnittelija"

2. SOVELLUKSEN KÄYNNISTÄMINEN
   - Kaksoisklikkaa "Ruokalistasuunnittelija.exe"
   - Sovellus avautuu selainikkunaan, suoraan pääsivulle

3. TIETOKANNAN PALAUTUS
   - Sovellus käyttää automaattisesti mukana olevaa "meal_plans.db"-tietokantaa
   - Kaikki reseptit ja ruokalistat ovat siihen tallennettuina
   - Älä kopioi tai muuta meal_plans.db-tiedostoa muuten kuin varmuuskopioiden kautta

4. USEALTA KONEELTA KÄYTTÖ
   - Yksi kone ajaa sovellusta (palvelin)
   - Muut koneet avaavat sen selaimessa: http://<palvelin-koneen-nimi>:5001
   - (Esim. http://keittiokone:5001)

5. PÄIVITTÄINEN VARMUUSKOPIOINTI
   - Avaa "Hallinto"-välilehti → Varmuuskopiot
   - Klikkaa "Luo varmuuskopio" päivittäin
   - Lataa väliajoin myös koko järjestelmän ZIP-paketti ("Lataa koko sovellus
     varmuuskopiona") ja säilytä se erillään tästä koneesta (esim. pilvipalvelu, USB)

ONGELMAT? Katso TUKI.txt
"""

_PALAUTUS_TXT = """RUOKALISTASUUNNITTELIJA — PALAUTUS VARMUUSKOPIOISTA
===================================================

Jos sovellus kaatuu tai tietokanta vioittuu:

1. PALAUTUS VANHASTA PAKETISTA
   - Jos sinulla on vanhempi ruokalistasuunnittelija_backup_*.zip:
   - Pura se ja kopioi "meal_plans.db" uuteen asennukseen
   - Käynnistä sovellus uudelleen

2. PALAUTUS VARMUUSKOPIOISTA SOVELLUKSEN SISÄLLÄ
   - Avaa sovellus
   - Mene "Hallinto" → "Varmuuskopiot"
   - Valitse vanha varmuuskopio listasta
   - Klikkaa "Palauta"
   - Sovellus ottaa ensin turvakopion nykytilasta, sitten palauttaa valitun kopion
   - Käynnistä sovellus (exe) uudelleen palautuksen jälkeen

3. MANUAALINEN PALAUTUS (VIIMEINEN KEINO)
   - Pysäytä sovellus kokonaan
   - Kopioi varmuuskopion tiedosto (esim. meal_plans_20260101_120000.db) tilalle
     nimellä "meal_plans.db" sovelluskansioon
   - Käynnistä sovellus uudelleen

4. DATAN SIIRTÄMINEN UUDELLE KONEELLE
   - Lataa viimeisin "ruokalistasuunnittelija_backup_*.zip" (Hallinto-välilehdeltä
     tai tallennuspaikastasi)
   - Pura uudelle koneelle
   - Sovellus käynnistyy täydellä datalla

HUOMIO: meal_plans.db sisältää kaiken — reseptit ja ruokalistat.
Varmuuskopioi se säännöllisesti!
"""

_TUKI_TXT = """RUOKALISTASUUNNITTELIJA — TUKI JA VIANMÄÄRITYS
==============================================

YLEISIÄ ONGELMIA:

1. "Sovellus ei käynnisty"
   - Tarkista, ettei antivirus estä .exe-tiedostoa
   - Salli ohjelma, jos antivirus kysyy lupaa
   - Kokeile poistaa ja purkaa paketti uudelleen

2. "Selain avautuu, mutta 'Yhteyden epäonnistuminen'"
   - Tarkista, ettei toista Ruokalistasuunnittelijaa ole jo käynnissä
   - Tarkista palomuuri: sallitaan portti 5001
   - Käynnistä sovellus uudelleen

3. "Tietokanta vioittunut" (sovellus kaatuu)
   - Palauta varmuuskopioista (ks. PALAUTUS.txt)

4. "Muista koneista ei pääse sisään"
   - Tarkista, että palvelinkone on päällä
   - Tarkista koneen osoite: http://<koneen-nimi>:5001
   - Tarkista palomuuri: sallitaan paikallisen verkon liikenne portissa 5001

5. "Excel-tuonti ei toimi"
   - Tarkista tiedoston muoto (täytyy olla PoweResta-muotoinen .xlsx)
   - Tarkista, että Excelin välilehdet alkavat "Nimi"-rivillä

6. "OCR-tuonti on hidas"
   - Se on normaalia — kuvasta lukeminen kestää
   - Odota kunnes kaikki kuvat on käsitelty ennen selaimen sulkemista

YHTEYSTIEDOT:
- [Kehittäjän nimi ja yhteystiedot]
- Organisaatiosi IT-tuki — [puhelinnumero, email]

VIIMEISENÄ KEINONA: Käytä varmuuskopioita palauttaaksesi edellinen tunnettu hyvä tila.
"""


def _do_backup(db_path, backup_dir, keep=30):
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    dest = os.path.join(backup_dir, f'meal_plans_{stamp}.db')
    src = sqlite3.connect(db_path, timeout=10)
    dst = sqlite3.connect(dest)
    with dst:
        src.backup(dst)          # eheä kopio myös käytön aikana
    src.close(); dst.close()

    # siivoa vanhat
    backups = sorted(f for f in os.listdir(backup_dir) if _NAME_RE.match(f))
    for old in backups[:-keep]:
        os.remove(os.path.join(backup_dir, old))
    return os.path.basename(dest)


def init_backup(app, db_path, backup_dir=None, keep=30):
    backup_dir = backup_dir or os.path.join(os.path.dirname(db_path), 'varmuuskopiot')

    @app.route('/api/backup', methods=['POST'])
    def backup_create():
        name = _do_backup(db_path, backup_dir, keep)
        return jsonify({'ok': True, 'name': name})

    @app.route('/api/backup')
    def backup_list():
        os.makedirs(backup_dir, exist_ok=True)
        items = []
        for f in sorted(os.listdir(backup_dir), reverse=True):
            if _NAME_RE.match(f):
                p = os.path.join(backup_dir, f)
                items.append({'name': f,
                              'size_kb': round(os.path.getsize(p) / 1024),
                              'modified': datetime.fromtimestamp(
                                  os.path.getmtime(p)).strftime('%d.%m.%Y %H:%M')})
        return jsonify(items)

    @app.route('/api/backup/<name>')
    def backup_download(name):
        if not _NAME_RE.match(name):
            return jsonify({'error': 'Virheellinen tiedostonimi'}), 400
        return send_file(os.path.join(backup_dir, name), as_attachment=True)

    @app.route('/api/backup/restore', methods=['POST'])
    def backup_restore():
        name = (request.get_json(force=True) or {}).get('name', '')
        if not _NAME_RE.match(name):
            return jsonify({'error': 'Virheellinen tiedostonimi'}), 400
        src = os.path.join(backup_dir, name)
        if not os.path.exists(src):
            return jsonify({'error': 'Varmuuskopiota ei löydy'}), 404
        # varmuuden vuoksi: kopio nykytilasta ennen palautusta
        safety = _do_backup(db_path, backup_dir, keep)
        shutil.copy2(src, db_path)
        return jsonify({'ok': True, 'restored': name, 'safety_copy': safety})

    @app.route('/api/backup/restore-upload', methods=['POST'])
    def backup_restore_upload():
        """Palauta tietokanta käyttäjän koneelta ladatusta .db-tiedostosta
        (esim. USB-tikulta tai toiselta koneelta) — eroaa /api/backup/restore
        siinä, että lähde ei ole jo palvelimella oleva varmuuskopio.
        Validoi ensin että tiedosto on aidosti käyttökelpoinen
        Ruokalistasuunnittelija-tietokanta ennen ylikirjoitusta, ja ottaa
        turvakopion nykytilasta ennen palautusta, aivan kuten tavallinen
        palautuskin."""
        if 'file' not in request.files or not request.files['file'].filename:
            return jsonify({'error': 'Valitse .db-tiedosto'}), 400
        f = request.files['file']
        if not f.filename.lower().endswith('.db'):
            return jsonify({'error': 'Tiedoston pitää olla .db-muotoinen'}), 400

        tmp_fd, tmp_path = tempfile.mkstemp(suffix='.db')
        os.close(tmp_fd)
        f.save(tmp_path)

        # Yhteys PITÄÄ sulkea ennen os.remove():a — Windows ei salli avoimena
        # olevan tiedoston poistoa, mikä aiemmin aiheutti PermissionErrorin
        # joka peitti alkuperäisen, käyttäjälle näytettävän virheilmoituksen.
        tables = None
        test_conn = sqlite3.connect(tmp_path)
        try:
            tables = {r[0] for r in test_conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
        except sqlite3.DatabaseError:
            pass
        finally:
            test_conn.close()

        if tables is None:
            os.remove(tmp_path)
            return jsonify({'error': 'Tiedosto ei ole kelvollinen SQLite-tietokanta'}), 400

        required = {'recipes', 'meal_plans', 'meal_plan_days'}
        if not required.issubset(tables):
            os.remove(tmp_path)
            return jsonify({'error': 'Tiedosto ei näytä Ruokalistasuunnittelijan '
                                     'tietokannalta (vaaditut taulut puuttuvat)'}), 400

        safety = _do_backup(db_path, backup_dir, keep)
        shutil.copy2(tmp_path, db_path)
        os.remove(tmp_path)
        return jsonify({'ok': True, 'message': 'Tietokanta palautettu ladatusta tiedostosta',
                        'safety_copy': safety})

    @app.route('/api/backup/download-package')
    def backup_download_package():
        """Kokoaa koko sovelluksen (exe + tietokanta + resurssit + ohjeet)
        yhdeksi ladattavaksi ZIP-paketiksi offline-palautusta varten."""
        import zipfile
        import io

        project_root = os.path.dirname(os.path.abspath(db_path))
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        zip_name = f'ruokalistasuunnittelija_backup_{stamp}.zip'
        root_folder = 'Ruokalistasuunnittelija'

        # eheä tietokantakopio sqlite:n omalla backup-API:lla — ei raakaa
        # tiedostokopiota, joka voisi jäädä puolitiehen jos sovellus kirjoittaa samaan aikaan
        tmp_db_fd, tmp_db_path = tempfile.mkstemp(suffix='.db')
        os.close(tmp_db_fd)
        buf = io.BytesIO()
        try:
            src = sqlite3.connect(db_path, timeout=10)
            dst = sqlite3.connect(tmp_db_path)
            with dst:
                src.backup(dst)
            src.close(); dst.close()

            with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.write(tmp_db_path, f'{root_folder}/meal_plans.db')

                exe_path = os.path.join(project_root, 'dist', 'Ruokalistasuunnittelija',
                                        'Ruokalistasuunnittelija.exe')
                exe_warning = ''
                if os.path.exists(exe_path):
                    zf.write(exe_path, f'{root_folder}/Ruokalistasuunnittelija.exe')
                    py_files = [f for f in os.listdir(project_root) if f.endswith('.py')]
                    newest_py = max((os.path.getmtime(os.path.join(project_root, f))
                                     for f in py_files), default=0)
                    if os.path.getmtime(exe_path) < newest_py:
                        exe_warning = ('\n\nHUOM: Pakettiin liitetty .exe on vanhempi kuin '
                                       'sovelluksen lähdekoodi tällä koneella. Aja RAKENNA_EXE.bat '
                                       'ennen paketin luontia, jos haluat viimeisimmät '
                                       'ominaisuudet mukaan.')
                else:
                    exe_warning = ('\n\nHUOM: Valmista .exe-tiedostoa ei löytynyt '
                                   '(dist-kansio puuttuu). Aja ensin RAKENNA_EXE.bat.')

                templates_src = os.path.join(project_root, 'templates_src')
                if os.path.isdir(templates_src):
                    for dirpath, _, filenames in os.walk(templates_src):
                        for fname in filenames:
                            fp = os.path.join(dirpath, fname)
                            arc = os.path.join(root_folder, 'templates_src',
                                               os.path.relpath(fp, templates_src))
                            zf.write(fp, arc)

                docs = {
                    'README.txt': _README_TXT,
                    'ASENNA.txt': _ASENNA_TXT,
                    'PALAUTUS.txt': _PALAUTUS_TXT,
                    'TUKI.txt': _TUKI_TXT + exe_warning,
                }
                for fname, content in docs.items():
                    zf.writestr(f'{root_folder}/{fname}', content)
        finally:
            try:
                os.remove(tmp_db_path)
            except OSError:
                pass

        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name=zip_name,
                         mimetype='application/zip')

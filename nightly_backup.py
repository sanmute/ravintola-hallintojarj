#!/usr/bin/env python
"""
nightly_backup.py — Itsenäinen tietokannan varmuuskopiointiskripti Windowsin
Task Scheduleria varten.

Varmuuskopioi meal_plans.db -> varmuuskopiot/-kansioon. Ei riipu Flaskista
tai muusta sovelluksesta, joten se toimii vaikka Ruokalistasuunnittelija
ei olisi käynnissä. Tuottaa saman meal_plans_YYYYMMDD_HHMMSS.db-nimisen
tiedoston kuin sovelluksen sisäinen "Luo varmuuskopio" (backup.py), joten
Task Schedulerin tekemät kopiot näkyvät automaattisesti myös Hallinto-
välilehden varmuuskopiolistassa ja Tietoa-sivun "viimeisin varmuuskopio"
-tiedossa.

Ajastus: ks. nightly_backup.bat + Windowsin Task Scheduler
    (OHJEET_HALLINTA.md, kohta "Automaattiset varmuuskopiot").

Aja manuaalisesti testataksesi:
    python nightly_backup.py
"""
import os
import sqlite3
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'meal_plans.db')
BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'varmuuskopiot')
MAX_BACKUPS = 30


def backup_database():
    """Luo aikaleimatun varmuuskopion tietokannasta. Palauttaa True/False."""
    if not os.path.exists(DB_PATH):
        print(f'[BACKUP] VIRHE: Tietokantaa ei löydy: {DB_PATH}')
        return False

    os.makedirs(BACKUP_DIR, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_name = f'meal_plans_{timestamp}.db'
    backup_path = os.path.join(BACKUP_DIR, backup_name)

    try:
        # SQLite:n oma backup-API — turvallinen, ei lukitse tietokantaa
        # vaikka sovellus olisi samaan aikaan käynnissä ja kirjoittamassa.
        src = sqlite3.connect(DB_PATH, timeout=10)
        dst = sqlite3.connect(backup_path)
        with dst:
            src.backup(dst)
        src.close()
        dst.close()

        print(f'[BACKUP] OK: {backup_name}')
        cleanup_old_backups()
        return True

    except Exception as e:
        print(f'[BACKUP] VIRHE: {e}')
        return False


def cleanup_old_backups():
    """Pitää vain MAX_BACKUPS uusinta tiedostoa (kaikki meal_plans_*.db-kopiot,
    riippumatta siitä onko ne luonut tämä skripti vai sovelluksen oma
    "Luo varmuuskopio" -painike — ne jakavat saman kansion ja nimeämiskäytännön)."""
    if not os.path.isdir(BACKUP_DIR):
        return

    files = []
    for fname in os.listdir(BACKUP_DIR):
        if fname.startswith('meal_plans_') and fname.endswith('.db'):
            fpath = os.path.join(BACKUP_DIR, fname)
            files.append((fpath, os.path.getmtime(fpath)))

    if len(files) > MAX_BACKUPS:
        files.sort(key=lambda x: x[1], reverse=True)
        for fpath, _ in files[MAX_BACKUPS:]:
            try:
                os.remove(fpath)
                print(f'[BACKUP] Poistettu vanha: {os.path.basename(fpath)}')
            except OSError as e:
                print(f'[BACKUP] Poisto epäonnistui ({fpath}): {e}')


if __name__ == '__main__':
    ok = backup_database()
    raise SystemExit(0 if ok else 1)

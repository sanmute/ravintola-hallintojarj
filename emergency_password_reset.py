#!/usr/bin/env python
"""
emergency_password_reset.py — Viimeinen keino salasanan palautukseen.

KÄYTÄ VAIN JOS: käyttäjä on lukittu ulos EIKÄ recovery-koodia ole
(ks. "Hallinto" -> "Salasanan palautuskoodi" sovelluksessa, tai
kirjautumissivun "Unohditko salasanan?" -linkki).

Tämä skripti vaatii SUORAN pääsyn palvelinkoneelle (tiedostojärjestelmään),
jossa meal_plans.db sijaitsee — sitä ei voi ajaa etänä verkon yli. Tämä on
tarkoituksellista: salasanan nollaus ilman mitään todennusta olisi
tietoturvariski, jos sen voisi tehdä kuka tahansa verkossa. Fyysinen/
tiedostojärjestelmätason pääsy koneelle on siksi ainoa hyväksytty ehto.

Käyttö (komentoriviltä, samasta kansiosta missä app.py ja meal_plans.db ovat):

    python emergency_password_reset.py

Skripti listaa käyttäjät, kysyy kenen salasana nollataan, pyytää uuden
salasanan kahdesti ja pyytää vielä erillisen vahvistuksen ennen kuin
mitään tallennetaan.
"""
import getpass
import os
import sqlite3
import sys

from auth import _hash_password

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'meal_plans.db')


def main():
    if not os.path.exists(DB_PATH):
        print(f'VIRHE: Tietokantaa ei löydy: {DB_PATH}')
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    users = conn.execute('SELECT id, username, role FROM users ORDER BY username').fetchall()
    if not users:
        print('VIRHE: Käyttäjiä ei löytynyt tietokannasta.')
        sys.exit(1)

    print('Käyttäjät:')
    for u in users:
        print(f"  [{u['id']}] {u['username']}  ({u['role']})")

    raw = input('\nKenen käyttäjän salasana nollataan? (käyttäjätunnus): ').strip()
    user = conn.execute('SELECT * FROM users WHERE username = ?', (raw,)).fetchone()
    if not user:
        print(f"VIRHE: Käyttäjää '{raw}' ei löydy.")
        sys.exit(1)

    new_password = getpass.getpass('Uusi salasana (vähintään 6 merkkiä): ')
    if len(new_password) < 6:
        print('VIRHE: Salasanan pitää olla vähintään 6 merkkiä.')
        sys.exit(1)
    confirm = getpass.getpass('Vahvista uusi salasana: ')
    if new_password != confirm:
        print('VIRHE: Salasanat eivät täsmää.')
        sys.exit(1)

    ans = input(f"\nNollataanko käyttäjän '{user['username']}' ({user['role']}) salasana? "
                f"Kirjoita KYLLA vahvistaaksesi: ").strip()
    if ans != 'KYLLA':
        print('Peruttu — mitään ei muutettu.')
        sys.exit(0)

    conn.execute('UPDATE users SET password_hash = ? WHERE id = ?',
                (_hash_password(new_password), user['id']))
    conn.commit()
    conn.close()
    print(f"\nOK: käyttäjän '{user['username']}' salasana vaihdettu. "
          f"Kirjaudu sisään uudella salasanalla ja luo tarvittaessa uusi recovery-koodi.")


if __name__ == '__main__':
    main()

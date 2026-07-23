"""
desktop.py — Työpöytäversion käynnistin.

Käynnistää Flask-palvelimen taustasäikeessä ja avaa sovelluksen
natiivissa Windows-ikkunassa (Edge WebView2, valmiina Win10/11:ssä).

Kehityskäyttö:   python desktop.py
Paketointi:      pyinstaller ruokalistasuunnittelija.spec

Palvelin kuuntelee myös lähiverkossa (host 0.0.0.0), joten muut
koneet pääsevät samaan sovellukseen selaimella:
    http://<tämän-koneen-nimi>:5001
"""

import socket
import threading
import time
import urllib.request

import webview  # pywebview

import app as flask_app_module
from app import app, DB_PATH
from meal_plan_db import MealPlanDB
from meal_plan_modifier import MealModifier

PORT = 5001


def _lan_address():
    """Selvitä koneen lähiverkko-osoite näytettäväksi käyttäjälle."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'


def _run_server():
    MealPlanDB(DB_PATH)
    MealModifier(DB_PATH)
    app.run(debug=False, host='0.0.0.0', port=PORT, use_reloader=False)


def _wait_for_server(url, timeout=15):
    start = time.time()
    while time.time() - start < timeout:
        try:
            urllib.request.urlopen(url, timeout=1)
            return True
        except Exception:
            time.sleep(0.25)
    return False


def main():
    t = threading.Thread(target=_run_server, daemon=True)
    t.start()

    url = f'http://127.0.0.1:{PORT}'
    _wait_for_server(url)

    lan = _lan_address()
    window = webview.create_window(
        f'Ruokalistasuunnittelija  —  lähiverkossa: http://{lan}:{PORT}',
        url,
        width=1280,
        height=850,
        min_size=(1000, 700),
    )
    webview.start()  # blocks until window closed; daemon server dies with it


if __name__ == '__main__':
    main()

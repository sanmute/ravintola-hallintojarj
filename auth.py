"""
auth.py — Käyttäjähallinta ja roolit Ruokalistasuunnittelijalle.

Roolit:
  admin   — kaikki oikeudet + käyttäjien hallinta + varmuuskopiot
  muokkaus — saa luoda/muokata/poistaa reseptejä ja listoja
  katselu — saa vain katsella ja ladata vientitiedostoja

Käyttöönotto app.py:ssä:

    from auth import init_auth
    init_auth(app, DB_PATH)

Ensimmäisellä käynnistyksellä luodaan käyttäjä:
    tunnus: admin    salasana: vaihda123
Vaihda salasana heti Käyttäjät-sivulta!
"""

import os
import hashlib
import hmac
import secrets
import sqlite3
from functools import wraps

from flask import (request, session, redirect, url_for, jsonify,
                   render_template_string)

ROLES = ('admin', 'muokkaus', 'katselu')
_ITERATIONS = 200_000


# ---------------------------------------------------------------- passwords
def _hash_password(password: str, salt: bytes = None):
    salt = salt or secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, _ITERATIONS)
    return salt.hex() + '$' + dk.hex()


def _verify_password(password: str, stored: str) -> bool:
    try:
        salt_hex, dk_hex = stored.split('$')
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(dk_hex)
        dk = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, _ITERATIONS)
        return hmac.compare_digest(dk, expected)
    except Exception:
        return False


# ---------------------------------------------------------------- recovery codes
# Käytetään samaa PBKDF2-hajautusta kuin salasanoille (_hash_password/_verify_password) —
# recovery-koodi on rakenteellisesti vain toinen "salasana" per käyttäjä.
_CODE_ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'  # ei 0/O/1/I/L -sekaannuksia


def _generate_recovery_code():
    raw = ''.join(secrets.choice(_CODE_ALPHABET) for _ in range(12))
    return f'KESTI-{raw[0:4]}-{raw[4:8]}-{raw[8:12]}'


# ---------------------------------------------------------------- db
def _ensure_tables(db_path):
    conn = sqlite3.connect(db_path, timeout=10)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK (role IN ('admin','muokkaus','katselu')),
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS recovery_codes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL UNIQUE REFERENCES users(id) ON DELETE CASCADE,
            code_hash TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now')),
            used_at TEXT
        )
    """)
    # first run: create default admin
    n = conn.execute('SELECT COUNT(*) FROM users').fetchone()[0]
    if n == 0:
        conn.execute(
            'INSERT INTO users (username, password_hash, role) VALUES (?,?,?)',
            ('admin', _hash_password('vaihda123'), 'admin'))
        print('  >> Luotu oletuskäyttäjä  admin / vaihda123  — VAIHDA SALASANA!')
    conn.commit()
    conn.close()


def _get_user(db_path, username):
    conn = sqlite3.connect(db_path, timeout=10)
    conn.row_factory = sqlite3.Row
    row = conn.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
    conn.close()
    return row


# ---------------------------------------------------------------- login page
_LOGIN_HTML = """
<!doctype html><html lang="fi"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Kirjaudu — Ruokalistasuunnittelija</title>
<style>
  body { font-family: 'Segoe UI', system-ui, sans-serif; background:#f4f6f4;
         display:flex; align-items:center; justify-content:center; min-height:100vh; margin:0; }
  .card { background:#fff; padding:2.2rem 2.5rem; border-radius:12px;
          box-shadow:0 4px 20px rgba(0,0,0,.08); width:320px; }
  h1 { font-size:1.25rem; margin:0 0 .3rem; color:#2d4a2d; }
  p.sub { margin:0 0 1.4rem; color:#777; font-size:.85rem; }
  label { display:block; font-size:.85rem; color:#444; margin-bottom:.25rem; }
  input { width:100%; padding:.55rem .7rem; margin-bottom:1rem; border:1px solid #ccc;
          border-radius:6px; font-size:1rem; box-sizing:border-box; }
  input:focus { outline:2px solid #5CB85C; border-color:#5CB85C; }
  button { width:100%; padding:.6rem; background:#5CB85C; color:#fff; border:0;
           border-radius:6px; font-size:1rem; cursor:pointer; }
  button:hover { background:#4aa04a; }
  .err { background:#fdecea; color:#b23b3b; padding:.5rem .7rem; border-radius:6px;
         font-size:.85rem; margin-bottom:1rem; }
  .ok { background:#eafaf1; color:#1e7e45; padding:.5rem .7rem; border-radius:6px;
        font-size:.85rem; margin-bottom:1rem; }
  .link { display:block; text-align:center; margin-top:1rem; font-size:.82rem; }
  .link a { color:#5CB85C; text-decoration:none; cursor:pointer; }
  .link a:hover { text-decoration:underline; }
  .hint { font-size:.78rem; color:#888; margin:-.6rem 0 1rem; }
</style></head><body>
<div class="card">
  <h1>Ruokalistasuunnittelija</h1>

  <div id="loginBox">
    <p class="sub">Kirjaudu sisään jatkaaksesi</p>
    {% if error %}<div class="err">{{ error }}</div>{% endif %}
    <form method="post" action="{{ url_for('auth_login') }}">
      <label for="u">Käyttäjätunnus</label>
      <input id="u" name="username" autocomplete="username" autofocus required>
      <label for="p">Salasana</label>
      <input id="p" name="password" type="password" autocomplete="current-password" required>
      <button type="submit">Kirjaudu sisään</button>
    </form>
    <p class="link"><a onclick="showRecovery()">Unohditko salasanan? Käytä recovery-koodia</a></p>
  </div>

  <div id="recoveryBox" style="display:none">
    <p class="sub">Palauta salasana recovery-koodilla</p>
    <div id="recoveryMsg"></div>
    <label for="rc">Recovery-koodi</label>
    <input id="rc" placeholder="KESTI-XXXX-XXXX-XXXX" autocomplete="off">
    <label for="rp">Uusi salasana</label>
    <input id="rp" type="password" autocomplete="new-password">
    <p class="hint">Vähintään 6 merkkiä.</p>
    <button onclick="submitRecovery()">Palauta salasana</button>
    <p class="link"><a onclick="showLogin()">Takaisin kirjautumiseen</a></p>
  </div>
</div>
<script>
function showRecovery() {
  document.getElementById('loginBox').style.display = 'none';
  document.getElementById('recoveryBox').style.display = 'block';
}
function showLogin() {
  document.getElementById('recoveryBox').style.display = 'none';
  document.getElementById('loginBox').style.display = 'block';
}
async function submitRecovery() {
  const code = document.getElementById('rc').value.trim();
  const newPassword = document.getElementById('rp').value;
  const msgEl = document.getElementById('recoveryMsg');
  msgEl.innerHTML = '';
  try {
    const r = await fetch('/api/auth/reset-with-recovery-code', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({recovery_code: code, new_password: newPassword})
    });
    const d = await r.json();
    if (r.ok && d.success) {
      msgEl.innerHTML = '<div class="ok">' + d.message + '</div>';
      document.getElementById('rc').value = '';
      document.getElementById('rp').value = '';
      setTimeout(showLogin, 1800);
    } else {
      msgEl.innerHTML = '<div class="err">' + (d.error || 'Tuntematon virhe') + '</div>';
    }
  } catch (e) {
    msgEl.innerHTML = '<div class="err">Yhteysvirhe: ' + e.message + '</div>';
  }
}
</script>
</body></html>
"""


# ---------------------------------------------------------------- decorators
def role_required(*roles):
    """Use on individual routes if you need finer control than the global guard."""
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if session.get('role') not in roles:
                return jsonify({'error': 'Ei käyttöoikeutta'}), 403
            return fn(*a, **kw)
        return wrapper
    return deco


# ---------------------------------------------------------------- init
def init_auth(app, db_path, secret_file=None):
    """Wire login, logout, user management and a global role guard into app."""
    _ensure_tables(db_path)

    # Persistent session secret so logins survive restarts
    secret_file = secret_file or os.path.join(os.path.dirname(db_path), '.secret_key')
    if os.path.exists(secret_file):
        app.secret_key = open(secret_file, 'rb').read()
    else:
        app.secret_key = secrets.token_bytes(32)
        with open(secret_file, 'wb') as f:
            f.write(app.secret_key)

    app.config.setdefault('PERMANENT_SESSION_LIFETIME', 60 * 60 * 12)  # 12 h työvuoro

    # ---- routes -------------------------------------------------
    @app.route('/login', methods=['GET', 'POST'])
    def auth_login():
        if request.method == 'POST':
            user = _get_user(db_path, request.form.get('username', '').strip())
            if user and _verify_password(request.form.get('password', ''), user['password_hash']):
                session.permanent = True
                session['user_id'] = user['id']
                session['username'] = user['username']
                session['role'] = user['role']
                return redirect('/')
            return render_template_string(_LOGIN_HTML,
                                          error='Väärä tunnus tai salasana'), 401
        return render_template_string(_LOGIN_HTML, error=None)

    @app.route('/logout')
    def auth_logout():
        session.clear()
        return redirect(url_for('auth_login'))

    @app.route('/api/me')
    def auth_me():
        return jsonify({'username': session.get('username'),
                        'role': session.get('role')})

    # ---- user management (admin only) ----------------------------
    @app.route('/api/users')
    @role_required('admin')
    def auth_users():
        conn = sqlite3.connect(db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        rows = conn.execute('SELECT id, username, role, created_at FROM users ORDER BY username').fetchall()
        conn.close()
        return jsonify([dict(r) for r in rows])

    @app.route('/api/users', methods=['POST'])
    @role_required('admin')
    def auth_add_user():
        d = request.get_json(force=True)
        username = (d.get('username') or '').strip()
        password = d.get('password') or ''
        role = d.get('role')
        if not username or len(password) < 6 or role not in ROLES:
            return jsonify({'error': 'Anna tunnus, vähintään 6 merkin salasana ja rooli'}), 400
        conn = sqlite3.connect(db_path, timeout=10)
        try:
            conn.execute('INSERT INTO users (username, password_hash, role) VALUES (?,?,?)',
                         (username, _hash_password(password), role))
            conn.commit()
        except sqlite3.IntegrityError:
            return jsonify({'error': 'Tunnus on jo olemassa'}), 409
        finally:
            conn.close()
        return jsonify({'ok': True})

    @app.route('/api/users/<int:uid>', methods=['DELETE'])
    @role_required('admin')
    def auth_del_user(uid):
        if uid == session.get('user_id'):
            return jsonify({'error': 'Et voi poistaa omaa tunnustasi'}), 400
        conn = sqlite3.connect(db_path, timeout=10)
        conn.execute('DELETE FROM users WHERE id = ?', (uid,))
        conn.commit(); conn.close()
        return jsonify({'ok': True})

    @app.route('/api/users/password', methods=['POST'])
    def auth_change_password():
        """Oman salasanan vaihto (kaikki roolit)."""
        d = request.get_json(force=True)
        user = _get_user(db_path, session.get('username', ''))
        if not user or not _verify_password(d.get('old', ''), user['password_hash']):
            return jsonify({'error': 'Vanha salasana väärin'}), 400
        if len(d.get('new', '')) < 6:
            return jsonify({'error': 'Uuden salasanan pitää olla vähintään 6 merkkiä'}), 400
        conn = sqlite3.connect(db_path, timeout=10)
        conn.execute('UPDATE users SET password_hash = ? WHERE id = ?',
                     (_hash_password(d['new']), user['id']))
        conn.commit(); conn.close()
        return jsonify({'ok': True})

    # ---- recovery code: salasanan palautus jos se unohtuu ---------
    @app.route('/api/auth/recovery-code/generate', methods=['POST'])
    def auth_generate_recovery_code():
        """Luo kertakäyttöisen recovery-koodin nykyiselle kirjautuneelle
        admin-käyttäjälle. Vaatii nykyisen salasanan vahvistuksen."""
        if session.get('role') != 'admin':
            return jsonify({'error': 'Vain ylläpito voi luoda recovery-koodin'}), 403
        d = request.get_json(force=True)
        user = _get_user(db_path, session.get('username', ''))
        if not user or not _verify_password(d.get('password', ''), user['password_hash']):
            return jsonify({'error': 'Salasana on väärä'}), 401

        code = _generate_recovery_code()
        conn = sqlite3.connect(db_path, timeout=10)
        conn.execute(
            '''INSERT INTO recovery_codes (user_id, code_hash, created_at, used_at)
               VALUES (?,?,datetime('now'),NULL)
               ON CONFLICT(user_id) DO UPDATE SET
                   code_hash=excluded.code_hash, created_at=excluded.created_at, used_at=NULL''',
            (user['id'], _hash_password(code)))
        conn.commit()
        conn.close()
        return jsonify({'recovery_code': code,
                        'message': 'Tallenna tämä koodi turvalliseen paikkaan. Aiempi koodi (jos oli) '
                                   'lakkasi toimimasta.'})

    @app.route('/api/auth/reset-with-recovery-code', methods=['POST'])
    def auth_reset_with_recovery_code():
        """Palauta salasana recovery-koodilla — ei vaadi kirjautumista."""
        d = request.get_json(force=True)
        code = (d.get('recovery_code') or '').strip().upper()
        new_password = d.get('new_password') or ''
        if not code or not new_password:
            return jsonify({'error': 'Recovery-koodi ja uusi salasana vaaditaan'}), 400
        if len(new_password) < 6:
            return jsonify({'error': 'Uuden salasanan pitää olla vähintään 6 merkkiä'}), 400

        conn = sqlite3.connect(db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        candidates = conn.execute(
            'SELECT user_id, code_hash FROM recovery_codes WHERE used_at IS NULL').fetchall()
        match = next((c for c in candidates if _verify_password(code, c['code_hash'])), None)
        if not match:
            conn.close()
            return jsonify({'error': 'Recovery-koodi on väärä tai jo käytetty'}), 401

        conn.execute('UPDATE users SET password_hash = ? WHERE id = ?',
                     (_hash_password(new_password), match['user_id']))
        conn.execute("UPDATE recovery_codes SET used_at = datetime('now') WHERE user_id = ?",
                     (match['user_id'],))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Salasana vaihdettu. Voit nyt kirjautua sisään.'})

    # ---- global guard --------------------------------------------
    WRITE_METHODS = {'POST', 'PUT', 'DELETE', 'PATCH'}
    PUBLIC_ENDPOINTS = {'auth_login', 'auth_reset_with_recovery_code', 'static'}

    @app.before_request
    def _guard():
        if request.endpoint in PUBLIC_ENDPOINTS or request.endpoint is None:
            return
        if 'user_id' not in session:
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Kirjaudu sisään'}), 401
            return redirect(url_for('auth_login'))
        # katselu-rooli: vain lukeminen
        if request.method in WRITE_METHODS and session.get('role') == 'katselu':
            return jsonify({'error': 'Katselutunnuksella ei voi tehdä muutoksia'}), 403

"""
order_catalog.py — Viikoittainen tilausluettelo (tuotekatalogi).

Kokoaa viikon ruokalistan reseptien raaka-aineet yhteen tilauslistaksi,
jonka määriä käyttäjä voi säätää kysynnän ja varastotilanteen mukaan.
Säädöt säilyvät, vaikka lista laskettaisiin uudelleen (esim. ruokalistan
muututtua): laskennallinen määrä päivittyy, käsin asetettu määrä pysyy.

Käyttöönotto app.py:ssä (initien jälkeen):

    from order_catalog import init_order_catalog
    init_order_catalog(app, DB_PATH)

Reitit:
  POST   /api/catalog/generate           {plan_id, week}  luo/päivitä
  GET    /api/catalog?plan_id=&week=                      hae rivit
  PATCH  /api/catalog/item/<id>          {adjusted_qty|note|checked}
  POST   /api/catalog/item               {catalog_id, name, unit, qty}
  DELETE /api/catalog/item/<id>
  GET    /api/catalog/export?plan_id=&week=               Excel
"""

import io
import os
import sqlite3
from datetime import datetime, timedelta

from flask import jsonify, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Font
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def _conn(db_path):
    conn = sqlite3.connect(db_path, timeout=10)
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_tables(db_path):
    conn = _conn(db_path)
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS order_catalogs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meal_plan_id INTEGER NOT NULL,
            week_number INTEGER NOT NULL,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now')),
            UNIQUE (meal_plan_id, week_number)
        );
        CREATE TABLE IF NOT EXISTS order_catalog_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            catalog_id INTEGER NOT NULL REFERENCES order_catalogs(id)
                ON DELETE CASCADE,
            name TEXT NOT NULL,
            unit TEXT DEFAULT '',
            calculated_qty REAL,          -- resepteistä laskettu
            adjusted_qty REAL,            -- käyttäjän säätämä (NULL = käytä laskettua)
            note TEXT DEFAULT '',
            checked INTEGER DEFAULT 0,    -- 'kerätty/tilattu' -rasti
            manual INTEGER DEFAULT 0,     -- käsin lisätty rivi
            UNIQUE (catalog_id, name, unit)
        );
    """)
    conn.commit()
    conn.close()


def _aggregate_ingredients(conn, plan_id, week):
    """Sum ingredient quantities for one week of a meal plan."""
    return conn.execute(
        '''SELECT i.name_fi AS name, ri.unit AS unit,
                  SUM(ri.quantity) AS total
           FROM meal_plan_days d
           JOIN recipe_ingredients ri ON ri.recipe_id = d.recipe_id
           JOIN ingredients i ON i.id = ri.ingredient_id
           WHERE d.meal_plan_id = ? AND d.week_number = ?
           GROUP BY i.name_fi, ri.unit
           ORDER BY i.name_fi COLLATE NOCASE''', (plan_id, week)).fetchall()


def init_order_catalog(app, db_path):
    _ensure_tables(db_path)

    # ---------------------------------------------------- generate / refresh
    @app.route('/api/catalog/generate', methods=['POST'])
    def catalog_generate():
        d = request.get_json(force=True)
        plan_id, week = d.get('plan_id'), d.get('week')
        if not plan_id or not week:
            return jsonify({'error': 'plan_id ja week vaaditaan'}), 400

        conn = _conn(db_path)
        try:
            ings = _aggregate_ingredients(conn, plan_id, week)

            cat = conn.execute(
                'SELECT id FROM order_catalogs WHERE meal_plan_id=? AND week_number=?',
                (plan_id, week)).fetchone()
            if cat:
                cat_id = cat['id']
            else:
                cur = conn.execute(
                    'INSERT INTO order_catalogs (meal_plan_id, week_number) VALUES (?,?)',
                    (plan_id, week))
                cat_id = cur.lastrowid

            # Upsert calculated rows; user adjustments (adjusted_qty, note,
            # checked) are preserved because we only touch calculated_qty.
            new_names = set()
            for ing in ings:
                new_names.add((ing['name'], ing['unit'] or ''))
                conn.execute(
                    '''INSERT INTO order_catalog_items
                           (catalog_id, name, unit, calculated_qty)
                       VALUES (?,?,?,?)
                       ON CONFLICT(catalog_id, name, unit)
                       DO UPDATE SET calculated_qty = excluded.calculated_qty''',
                    (cat_id, ing['name'], ing['unit'] or '',
                     round(ing['total'], 3)))

            # Remove auto rows whose ingredient no longer appears (keep
            # manual rows and rows the user adjusted or annotated).
            for row in conn.execute(
                    '''SELECT id, name, unit FROM order_catalog_items
                       WHERE catalog_id=? AND manual=0 AND adjusted_qty IS NULL
                         AND (note IS NULL OR note='') AND checked=0''',
                    (cat_id,)).fetchall():
                if (row['name'], row['unit'] or '') not in new_names:
                    conn.execute('DELETE FROM order_catalog_items WHERE id=?',
                                 (row['id'],))

            conn.execute(
                "UPDATE order_catalogs SET updated_at=datetime('now') WHERE id=?",
                (cat_id,))
            conn.commit()
            n = conn.execute(
                'SELECT COUNT(*) c FROM order_catalog_items WHERE catalog_id=?',
                (cat_id,)).fetchone()['c']
        finally:
            conn.close()
        if not ings and n == 0:
            return jsonify({'catalog_id': cat_id, 'items': n,
                            'warning': 'Viikon resepteillä ei ole raaka-ainetietoja. '
                                       'Voit lisätä rivejä käsin.'})
        return jsonify({'catalog_id': cat_id, 'items': n})

    # ---------------------------------------------------- read
    @app.route('/api/catalog')
    def catalog_get():
        plan_id = request.args.get('plan_id', type=int)
        week = request.args.get('week', type=int)
        conn = _conn(db_path)
        try:
            cat = conn.execute(
                'SELECT * FROM order_catalogs WHERE meal_plan_id=? AND week_number=?',
                (plan_id, week)).fetchone()
            if not cat:
                return jsonify({'catalog': None, 'items': []})
            items = conn.execute(
                '''SELECT * FROM order_catalog_items WHERE catalog_id=?
                   ORDER BY checked, name COLLATE NOCASE''',
                (cat['id'],)).fetchall()
        finally:
            conn.close()
        return jsonify({'catalog': dict(cat),
                        'items': [dict(i) for i in items]})

    # ---------------------------------------------------- edit one row
    @app.route('/api/catalog/item/<int:item_id>', methods=['PATCH'])
    def catalog_item_patch(item_id):
        d = request.get_json(force=True)
        sets, vals = [], []
        if 'adjusted_qty' in d:
            v = d['adjusted_qty']
            if v is not None:
                try:
                    v = round(float(v), 3)
                    if v < 0:
                        return jsonify({'error': 'Määrä ei voi olla negatiivinen'}), 400
                except (TypeError, ValueError):
                    return jsonify({'error': 'Virheellinen määrä'}), 400
            sets.append('adjusted_qty=?'); vals.append(v)
        if 'note' in d:
            sets.append('note=?'); vals.append(str(d['note'])[:200])
        if 'checked' in d:
            sets.append('checked=?'); vals.append(1 if d['checked'] else 0)
        if not sets:
            return jsonify({'error': 'Ei muutettavia kenttiä'}), 400
        vals.append(item_id)
        conn = _conn(db_path)
        try:
            cur = conn.execute(
                f"UPDATE order_catalog_items SET {', '.join(sets)} WHERE id=?", vals)
            conn.commit()
            if cur.rowcount == 0:
                return jsonify({'error': 'Riviä ei löydy'}), 404
            row = conn.execute('SELECT * FROM order_catalog_items WHERE id=?',
                               (item_id,)).fetchone()
        finally:
            conn.close()
        return jsonify(dict(row))

    # ---------------------------------------------------- manual add / delete
    @app.route('/api/catalog/item', methods=['POST'])
    def catalog_item_add():
        d = request.get_json(force=True)
        name = (d.get('name') or '').strip()
        if not name or not d.get('catalog_id'):
            return jsonify({'error': 'catalog_id ja nimi vaaditaan'}), 400
        try:
            qty = round(float(d.get('qty', 0)), 3)
        except (TypeError, ValueError):
            return jsonify({'error': 'Virheellinen määrä'}), 400
        conn = _conn(db_path)
        try:
            cur = conn.execute(
                '''INSERT INTO order_catalog_items
                       (catalog_id, name, unit, adjusted_qty, manual)
                   VALUES (?,?,?,?,1)''',
                (d['catalog_id'], name, (d.get('unit') or '').strip(), qty))
            conn.commit()
            row = conn.execute('SELECT * FROM order_catalog_items WHERE id=?',
                               (cur.lastrowid,)).fetchone()
        except sqlite3.IntegrityError:
            return jsonify({'error': 'Tuote on jo listalla'}), 409
        finally:
            conn.close()
        return jsonify(dict(row))

    @app.route('/api/catalog/item/<int:item_id>', methods=['DELETE'])
    def catalog_item_delete(item_id):
        conn = _conn(db_path)
        try:
            conn.execute('DELETE FROM order_catalog_items WHERE id=?', (item_id,))
            conn.commit()
        finally:
            conn.close()
        return jsonify({'ok': True})

    # ---------------------------------------------------- Excel export
    @app.route('/api/catalog/export')
    def catalog_export():
        plan_id = request.args.get('plan_id', type=int)
        week = request.args.get('week', type=int)
        conn = _conn(db_path)
        try:
            cat = conn.execute(
                'SELECT * FROM order_catalogs WHERE meal_plan_id=? AND week_number=?',
                (plan_id, week)).fetchone()
            if not cat:
                return jsonify({'error': 'Tilauslistaa ei ole luotu tälle viikolle'}), 404
            items = conn.execute(
                '''SELECT * FROM order_catalog_items WHERE catalog_id=?
                   ORDER BY name COLLATE NOCASE''', (cat['id'],)).fetchall()
        finally:
            conn.close()

        wb = Workbook()
        ws = wb.active
        ws.title = f'Viikko {week}'
        ws['A1'] = f'TILAUSLISTA — Viikko {week}'
        ws['A1'].font = Font(bold=True, size=13)
        ws['A2'] = f'Luotu: {datetime.now().strftime("%d.%m.%Y %H:%M")}'

        headers = ['Tuote', 'Yksikkö', 'Laskettu määrä', 'Tilattava määrä',
                   'Huomiot', 'Kerätty']
        for col, h in enumerate(headers, 1):
            c = ws.cell(4, col, h)
            c.font = Font(bold=True)

        r = 5
        for it in items:
            qty = it['adjusted_qty'] if it['adjusted_qty'] is not None \
                else it['calculated_qty']
            ws.cell(r, 1, it['name'])
            ws.cell(r, 2, it['unit'])
            ws.cell(r, 3, it['calculated_qty'])
            ws.cell(r, 4, qty)
            ws.cell(r, 5, it['note'] or '')
            ws.cell(r, 6, 'x' if it['checked'] else '')
            if it['adjusted_qty'] is not None:
                ws.cell(r, 4).font = Font(bold=True)  # korostetaan säädetyt
            r += 1

        for col, width in zip('ABCDEF', (34, 10, 15, 16, 30, 9)):
            ws.column_dimensions[col].width = width

        buf = io.BytesIO()
        wb.save(buf)
        wb.close()
        buf.seek(0)
        return send_file(
            buf, as_attachment=True,
            download_name=f'tilauslista_vko{week}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # ---------------------------------------------------- PDF export (print)
    @app.route('/api/catalog/export/pdf')
    def catalog_export_pdf():
        plan_id = request.args.get('plan_id', type=int)
        week = request.args.get('week', type=int)
        conn = _conn(db_path)
        try:
            cat = conn.execute(
                'SELECT * FROM order_catalogs WHERE meal_plan_id=? AND week_number=?',
                (plan_id, week)).fetchone()
            if not cat:
                return jsonify({'error': 'Tilauslistaa ei ole luotu tälle viikolle'}), 404
            items = conn.execute(
                '''SELECT * FROM order_catalog_items WHERE catalog_id=?
                   ORDER BY name COLLATE NOCASE''', (cat['id'],)).fetchall()
        finally:
            conn.close()

        buf = io.BytesIO()
        doc = SimpleDocTemplate(
            buf, pagesize=A4,
            topMargin=15 * mm, bottomMargin=15 * mm,
            leftMargin=15 * mm, rightMargin=15 * mm,
            title=f'Tilauslista viikko {week}')

        styles = getSampleStyleSheet()
        elements = [
            Paragraph(f'TILAUSLISTA — Viikko {week}', styles['Title']),
            Paragraph(f'Luotu: {datetime.now().strftime("%d.%m.%Y %H:%M")}',
                      styles['Normal']),
            Spacer(1, 8 * mm),
        ]

        # Selkeät solutyylit — Paragraph rivittää pitkän tekstin sarakkeen
        # sisään sen sijaan, että se valuisi viereisen sarakkeen päälle.
        cell_style = ParagraphStyle('cell', parent=styles['Normal'],
                                    fontName='Helvetica', fontSize=9, leading=11)
        cell_style_right = ParagraphStyle('cellRight', parent=cell_style,
                                          alignment=2)  # TA_RIGHT

        header = ['Tuote', 'Yksikkö', 'Laskettu', 'Tilattava', 'Huomiot', 'Kerätty']
        rows = [header]
        for it in items:
            qty = it['adjusted_qty'] if it['adjusted_qty'] is not None \
                else it['calculated_qty']
            rows.append([
                Paragraph(it['name'] or '', cell_style),
                Paragraph(it['unit'] or '', cell_style),
                Paragraph('' if it['calculated_qty'] is None else str(it['calculated_qty']),
                          cell_style_right),
                Paragraph('' if qty is None else str(qty), cell_style_right),
                Paragraph(it['note'] or '', cell_style),
                'x' if it['checked'] else '',
            ])

        # Sarakeleveydet mahtuvat A4:n käytettävään leveyteen
        # (210mm - 15mm*2 marginaalit = 180mm).
        col_widths = [50 * mm, 15 * mm, 20 * mm, 20 * mm, 58 * mm, 15 * mm]
        table = Table(rows, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#5CB85C')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (5, 0), (5, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f5f5f5')]),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(table)
        doc.build(elements)
        buf.seek(0)
        return send_file(
            buf, as_attachment=False,
            download_name=f'tilauslista_vko{week}.pdf',
            mimetype='application/pdf')


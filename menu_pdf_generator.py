"""
menu_pdf_generator.py — Brändätty viikkoruokalista (Kesti).

Uses the restaurant's own menu PDF as the visual template:
  1. Strips dish TEXT everywhere (redaction pass 1)
  2. Strips the vector-outline title and day labels (redaction pass 2,
     graphics removal limited to those boxes)
  3. Renders the result once to branded_background.png (cached)
  4. Overlays fresh title, day labels and dish lines at the EXACT
     coordinates measured from the original file, using the same
     typeface (Amaranth, OFL-licensed).

Layout measurements (PDF points, from the original template):
  - Dish lines: Amaranth 12 pt, centred on x=326.5
  - Day blocks start at y=143.2 (top), step 110.2, line gap 23.4
  - Day labels: x=57, vector height ~15 pt -> 20 pt Amaranth Bold
  - Title:      x=41, baseline y=121, cap height ~24 -> 30 pt bold
"""

import os
import re
from datetime import date, timedelta

import fitz  # pymupdf
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PDF = os.path.join(BASE_DIR, 'templates_src', 'menu_template.pdf')
BACKGROUND_PNG = os.path.join(BASE_DIR, 'templates_src', 'branded_background.png')
FONT_REG = os.path.join(BASE_DIR, 'templates_src', 'Amaranth-Regular.ttf')
FONT_BOLD = os.path.join(BASE_DIR, 'templates_src', 'Amaranth-Bold.ttf')

DAY_NAMES = ['Ma', 'Ti', 'Ke', 'To', 'Pe']
CODE_MAP = {'G': 'GL', 'VG': 'VEG'}
MENU_CODES = ['L', 'GL', 'M', 'VEG']

PAGE_W, PAGE_H = A4  # 595.28 x 841.89 (template is 595.8x842.4; close enough)

# ---- Measured layout constants (top-based y coordinates) -------------
TITLE_X = 41
TITLE_BASELINE = 121          # bottom of 'Lounaslista 24/26' glyphs
TITLE_SIZE = 30

DAY_LABEL_X = 57
DAY_LABEL_SIZE = 20
# Baselines of the five original day labels (bottom of glyph boxes)
DAY_LABEL_BASELINES = [192, 297, 412, 516, 606]

DISH_CENTER_X = 326.5
DISH_SIZE = 12
DISH_MIN_SIZE = 9
BLOCK_TOPS = [143.2, 253.1, 365.4, 474.0, 583.9]  # top of first line per day
LINE_GAP = 23.4
GLYPH_TOP_TO_BASELINE = 11.6  # Amaranth 12pt: baseline offset from span top

# Redaction boxes for stale vector title + day labels (x0, y0, x1, y1)
TITLE_BOX = (35, 94, 285, 130)
DAY_LABEL_BOXES = [(48, 170, 140, 200), (48, 274, 140, 304),
                   (48, 390, 140, 419), (48, 493, 140, 523),
                   (48, 584, 140, 613)]


def _register_fonts():
    """Register Amaranth if present; fall back to Helvetica."""
    try:
        if 'Amaranth' not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont('Amaranth', FONT_REG))
            pdfmetrics.registerFont(TTFont('Amaranth-Bold', FONT_BOLD))
        return 'Amaranth', 'Amaranth-Bold'
    except Exception:
        return 'Helvetica', 'Helvetica-Bold'


def ensure_background(force=False):
    """
    Create the branded background (cached).

    Two renders are composited:
      A) text-only redaction  -> keeps ALL graphics (pill, footer band,
         legend, disclaimers) but stale title/day labels remain
      B) + graphics redaction -> label areas clean, but shared vector
         layers (pill, footer band) can be lost
    Final image = A, with only the title/day-label rectangles copied
    from B. Brand graphics stay, stale labels go.
    """
    if os.path.exists(BACKGROUND_PNG) and not force:
        return BACKGROUND_PNG
    if not os.path.exists(TEMPLATE_PDF):
        return None

    DPI = 200
    scale = DPI / 72.0

    # Render A: text removed, ALL graphics explicitly kept
    # (newer PyMuPDF removes covered line art by DEFAULT, which would
    #  delete the pill, footer band, legend and disclaimers)
    doc = fitz.open(TEMPLATE_PDF)
    page = doc[0]
    page.add_redact_annot(page.rect)
    try:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                              graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    except TypeError:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    pix_a = page.get_pixmap(dpi=DPI)

    # Render B: additionally remove vector label/title outlines
    page.add_redact_annot(fitz.Rect(*TITLE_BOX))
    for box in DAY_LABEL_BOXES:
        page.add_redact_annot(fitz.Rect(*box))
    try:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                              graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
    except TypeError:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    pix_b = page.get_pixmap(dpi=DPI)

    # Composite with PIL: A everywhere, B inside the label rectangles
    from PIL import Image
    img_a = Image.frombytes('RGB', (pix_a.width, pix_a.height), pix_a.samples)
    img_b = Image.frombytes('RGB', (pix_b.width, pix_b.height), pix_b.samples)
    for (x0, y0, x1, y1) in [TITLE_BOX] + DAY_LABEL_BOXES:
        px = (int(x0 * scale), int(y0 * scale),
              int(x1 * scale) + 1, int(y1 * scale) + 1)
        img_a.paste(img_b.crop(px), px)

    os.makedirs(os.path.dirname(BACKGROUND_PNG), exist_ok=True)
    img_a.save(BACKGROUND_PNG)
    return BACKGROUND_PNG


def extract_menu_codes(notes):
    if not notes:
        return ''
    m = re.search(r'Ruokavaliot:\s*([A-ZÄÖÅ,\-\s]+)', notes)
    if not m:
        return ''
    raw = [c.strip().upper() for c in m.group(1).split(',') if c.strip()]
    mapped = [CODE_MAP.get(c, c) for c in raw]
    shown = [c for c in MENU_CODES if c in mapped]
    return f"({', '.join(shown)})" if shown else ''


def monday_of_week(week_number, year):
    try:
        return date.fromisocalendar(year, min(week_number, 52), 1)
    except ValueError:
        return date.fromisocalendar(year, 1, 1)


def _fmt_line(meal):
    """'JUUSTOINEN SAVUKALAKEITTO (L, GL)' from a meal dict."""
    if not meal:
        return None
    name = meal['name'].upper().strip()
    codes = extract_menu_codes(meal.get('notes'))
    return f'{name} {codes}'.strip()


def _split_days(meals):
    """
    Adapt a flat meal list into 5 day dicts {soup, salad, main, sides}.
    Heuristics: 'keitto' -> soup, 'salaatti' -> salad, meal.get('sides')
    or names with ', ' listing side dishes -> sides; rest -> main.
    A meal dict may also carry 'day' (0-4) and 'slot' for exact placement.
    """
    days = [{'soup': None, 'salad': None, 'main': None, 'sides': None}
            for _ in range(5)]

    # Exact placement first
    rest = []
    for m in meals:
        if m.get('day') is not None and m.get('slot') in days[0]:
            days[m['day'] % 5][m['slot']] = m
        else:
            rest.append(m)

    soups = [m for m in rest if 'keitto' in m['name'].lower()]
    salads = [m for m in rest if 'salaatti' in m['name'].lower()
              and 'keitto' not in m['name'].lower()]
    sides = [m for m in rest if m.get('slot') == 'sides']
    mains = [m for m in rest if m not in soups and m not in salads
             and m not in sides]

    for i in range(5):
        if not days[i]['soup'] and soups:
            days[i]['soup'] = soups[i % len(soups)]
        if not days[i]['salad'] and salads:
            days[i]['salad'] = salads[i % len(salads)]
        if not days[i]['main'] and mains:
            days[i]['main'] = mains[i % len(mains)]
        if not days[i]['sides'] and sides:
            days[i]['sides'] = sides[i % len(sides)]
    return days


def build_week_menu_pdf(meals, week_number, year, output_path, config=None):
    """
    Render a branded weekly menu PDF.

    meals: EITHER a flat list of {'name','notes',...} (heuristic split)
           OR a list of 5 dicts {'soup','salad','main','sides'} each a
           meal dict or None.
    Returns (output_path, warnings).
    """
    warnings = []
    config = config or {}
    font, bold = _register_fonts()

    if meals and isinstance(meals[0], dict) and 'soup' in meals[0]:
        days = meals
    else:
        days = _split_days(meals or [])

    for i, d in enumerate(days):
        if not d['soup']:
            warnings.append(f'{DAY_NAMES[i]}: keitto puuttuu')
        if not d['main']:
            warnings.append(f'{DAY_NAMES[i]}: pääruoka puuttuu')

    monday = monday_of_week(week_number, year)
    bg = ensure_background()

    c = canvas.Canvas(output_path, pagesize=A4)

    if bg and os.path.exists(bg):
        c.drawImage(ImageReader(bg), 0, 0, width=PAGE_W, height=PAGE_H,
                    preserveAspectRatio=False, mask='auto')

    # Title, e.g. 'Lounaslista 28/26'
    subtitle = config.get('subtitle', 'Lounaslista {week}/{yy}').format(
        week=week_number, yy=str(year)[-2:])
    c.setFillColorRGB(0.05, 0.05, 0.05)
    c.setFont(bold, TITLE_SIZE)
    c.drawString(TITLE_X, PAGE_H - TITLE_BASELINE, subtitle)

    # Day labels, e.g. 'Ma 6.7'
    for i, day in enumerate(DAY_NAMES):
        d = monday + timedelta(days=i)
        c.setFont(bold, DAY_LABEL_SIZE)
        c.drawString(DAY_LABEL_X, PAGE_H - DAY_LABEL_BASELINES[i],
                     f'{day} {d.day}.{d.month}')

    # Dish lines: soup / salad & buffet / main (& 2nd main, Friday only in
    # practice — real Kesti menus join two mains with "&", e.g.
    # "AURAJUUSTOPOSSUA & BROILERIN NUIJAT") / sides
    for i, d in enumerate(days):
        salad_line = _fmt_line(d['salad'])
        if salad_line and 'SALAATTIBUFFET' not in salad_line:
            salad_line += ' & SALAATTIBUFFET'
        elif not salad_line:
            salad_line = 'PÄIVÄN SALAATTI & SALAATTIBUFFET'
        main_line = _fmt_line(d['main'])
        main2_line = _fmt_line(d.get('main2'))
        if main_line and main2_line:
            main_line = f"{main_line} & {main2_line}"
        lines = [_fmt_line(d['soup']), salad_line, main_line, _fmt_line(d['sides'])]
        lines = [ln for ln in lines if ln]
        for j, line in enumerate(lines):
            y_top = BLOCK_TOPS[i] + j * LINE_GAP
            size = DISH_SIZE
            while stringWidth(line, font, size) > 460 and size > DISH_MIN_SIZE:
                size -= 0.5
            w = stringWidth(line, font, size)
            c.setFont(font, size)
            c.drawString(DISH_CENTER_X - w / 2,
                         PAGE_H - (y_top + GLYPH_TOP_TO_BASELINE), line)

    c.showPage()
    c.save()
    return output_path, warnings


def build_single_recipe_pdf(recipe, ingredients, output_path):
    """Yksittäisen reseptin tulostettava PDF keittiökäyttöön.

    recipe: dict with name_fi, recipe_type, season, servings, instructions
    ingredients: list of {'name_fi', 'quantity', 'unit'} (recipe_ingredients
      JOIN ingredients — EI recipes.ingredients-saraketta, jota mikään
      oikea tallennuspolku ei koskaan täytä)
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    font_name, font_bold = 'Helvetica', 'Helvetica-Bold'
    try:
        pdfmetrics.registerFont(TTFont('Amaranth', FONT_REG))
        pdfmetrics.registerFont(TTFont('Amaranth-Bold', FONT_BOLD))
        font_name, font_bold = 'Amaranth', 'Amaranth-Bold'
    except Exception:
        pass  # fonttitiedostot puuttuvat — kelpaa oletusfontilla

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                             topMargin=2*cm, bottomMargin=2*cm,
                             leftMargin=2*cm, rightMargin=2*cm)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('RecipeTitle', parent=styles['Heading1'],
                                  fontName=font_bold, fontSize=22, textColor='#1a5c3a',
                                  spaceAfter=6)
    meta_style = ParagraphStyle('RecipeMeta', parent=styles['Normal'],
                                 fontName=font_name, fontSize=11, textColor='#555',
                                 spaceAfter=16)
    heading_style = ParagraphStyle('RecipeHeading', parent=styles['Heading2'],
                                    fontName=font_bold, fontSize=14, textColor='#1a5c3a',
                                    spaceBefore=12, spaceAfter=8)
    body_style = ParagraphStyle('RecipeBody', parent=styles['Normal'],
                                 fontName=font_name, fontSize=11, leading=15)

    content = [Paragraph(recipe['name_fi'], title_style)]

    meta_parts = []
    if recipe.get('recipe_type'):
        meta_parts.append(recipe['recipe_type'].capitalize())
    if recipe.get('season'):
        meta_parts.append(recipe['season'].capitalize())
    if recipe.get('servings'):
        meta_parts.append(f"{recipe['servings']} annosta")
    if meta_parts:
        content.append(Paragraph(' · '.join(meta_parts), meta_style))

    content.append(Paragraph('Raaka-aineet', heading_style))
    if ingredients:
        items = []
        for ing in ingredients:
            qty = ing.get('quantity')
            unit = ing.get('unit') or ''
            qty_text = f"{qty:g} {unit}".strip() if qty is not None else unit
            line = f"{qty_text} {ing['name_fi']}".strip() if qty_text else ing['name_fi']
            items.append(ListItem(Paragraph(line, body_style), leftIndent=6))
        content.append(ListFlowable(items, bulletType='bullet', start='•'))
    else:
        content.append(Paragraph('Raaka-aineita ei ole vielä tallennettu tälle reseptille.',
                                  body_style))

    content.append(Paragraph('Valmistusohje', heading_style))
    if recipe.get('instructions'):
        for para in recipe['instructions'].split('\n'):
            if para.strip():
                content.append(Paragraph(para.strip(), body_style))
                content.append(Spacer(1, 4))
    else:
        content.append(Paragraph('Valmistusohjetta ei ole vielä tallennettu tälle reseptille.',
                                  body_style))

    doc.build(content)
    return output_path


# ----------------------------------------------------------------------
# Hoiva-branded weekly menus (Tammikoti & Jyränköläkoti / Palvelutalon
# asukas) — same redact-and-overlay technique as the Kesti template
# above, but for two real templates supplied by the restaurant, each a
# fixed 7-day (Ma-Su) layout with only 2 dish lines per day (main + soup,
# no separate salad/sides line — this template just doesn't show one).
#
# Only the parts that change week to week are redacted: the subtitle's
# week number, the 7 day-date labels (vector outlines, not real text —
# same as Kesti's day labels), and the dish text (real text). The static
# branding (title, facility name, footer, ordering instructions) is left
# untouched in the cached background image.
# ----------------------------------------------------------------------

HOIVA_TEMPLATES = {
    'tammikoti': {
        'pdf': os.path.join(BASE_DIR, 'templates_src', 'hoiva_template_tammikoti.pdf'),
        'background': os.path.join(BASE_DIR, 'templates_src', 'hoiva_background_tammikoti.png'),
        'subtitle_format': 'Lounaslista {week}/{yy}',
        'subtitle_box': (28, 100, 252, 135),
        'subtitle_x': 32.8,
        'subtitle_baseline': 130.9,
        'subtitle_size': 22,
        'day_label_boxes': [
            (47, 186, 115, 209), (52, 265, 109, 289), (50, 333, 111, 356),
            (52, 396, 113, 419), (52, 470, 113, 493), (52, 550, 111, 573),
            (49, 628, 111, 651),
        ],
        'day_label_baselines': [204.7, 284.2, 351.6, 414.6, 488.7, 568.3, 646.8],
        'dish_boxes': [
            [(267, 190, 401, 206), (259, 214, 409, 230)],
            [(262, 262, 405, 278), (241, 286, 424, 302)],
            [(251, 330, 392, 346), (240, 353, 401, 369)],
            [(281, 396, 385, 412), (202, 419, 464, 435)],
            [(263, 470, 404, 486), (274, 494, 393, 509)],
            [(277, 543, 359, 559), (266, 566, 371, 582)],
            [(279, 615, 382, 631), (282, 638, 379, 654)],
        ],
        'dish_tops': [192.4, 264.1, 332.5, 398.4, 471.9, 545.4, 617.4],
    },
    'asukas': {
        'pdf': os.path.join(BASE_DIR, 'templates_src', 'hoiva_template_asukas.pdf'),
        'background': os.path.join(BASE_DIR, 'templates_src', 'hoiva_background_asukas.png'),
        'subtitle_format': 'Lounaslista vko {week} / {yy}',
        'subtitle_box': (37, 222, 322, 256),
        'subtitle_x': 42.0,
        'subtitle_baseline': 251.0,
        'subtitle_size': 22,
        'day_label_boxes': [
            (50, 262, 118, 285), (55, 324, 111, 347), (50, 388, 111, 412),
            (50, 456, 111, 479), (50, 519, 111, 543), (50, 596, 109, 620),
            (48, 671, 111, 695),
        ],
        'day_label_baselines': [281.0, 342.7, 407.2, 474.8, 538.0, 615.0, 690.6],
        'dish_boxes': [
            [(270, 259, 404, 275), (261, 283, 412, 299)],
            [(267, 320, 410, 336), (245, 344, 428, 360)],
            [(271, 382, 409, 398), (260, 406, 420, 422)],
            [(292, 448, 394, 464), (212, 472, 474, 488)],
            [(275, 521, 413, 537), (285, 545, 403, 561)],
            [(284, 585, 390, 601), (285, 608, 389, 624)],
            [(285, 654, 402, 670), (280, 677, 407, 693), (310, 699, 378, 716)],
        ],
        'dish_tops': [261.6, 321.9, 384.4, 450.0, 523.5, 587.3, 656.1],
    },
}

HOIVA_DAY_NAMES = ['Ma', 'Ti', 'Ke', 'To', 'Pe', 'La', 'Su']
HOIVA_LINE_GAP = 23.4
HOIVA_GLYPH_TOP_TO_BASELINE = 11.6
HOIVA_DISH_CENTER_X = 330
HOIVA_DAY_LABEL_SIZE = 18


def _ensure_hoiva_background(variant, force=False):
    cfg = HOIVA_TEMPLATES[variant]
    if os.path.exists(cfg['background']) and not force:
        return cfg['background']
    if not os.path.exists(cfg['pdf']):
        return None

    DPI = 200
    scale = DPI / 72.0

    doc = fitz.open(cfg['pdf'])
    page = doc[0]

    # Pass A: remove only the real dish-text lines (their own bboxes),
    # keep every graphic — this is real text, so no graphics mode needed.
    for day_boxes in cfg['dish_boxes']:
        for box in day_boxes:
            page.add_redact_annot(fitz.Rect(*box))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                          graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    pix_a = page.get_pixmap(dpi=DPI)

    # Pass B: additionally strip the vector-outline subtitle number and
    # day-date labels (not real text — same situation as Kesti's labels).
    page.add_redact_annot(fitz.Rect(*cfg['subtitle_box']))
    for box in cfg['day_label_boxes']:
        page.add_redact_annot(fitz.Rect(*box))
    try:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                              graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
    except TypeError:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    pix_b = page.get_pixmap(dpi=DPI)

    from PIL import Image
    img_a = Image.frombytes('RGB', (pix_a.width, pix_a.height), pix_a.samples)
    img_b = Image.frombytes('RGB', (pix_b.width, pix_b.height), pix_b.samples)
    for (x0, y0, x1, y1) in [cfg['subtitle_box']] + cfg['day_label_boxes']:
        px = (int(x0 * scale), int(y0 * scale), int(x1 * scale) + 1, int(y1 * scale) + 1)
        img_a.paste(img_b.crop(px), px)

    os.makedirs(os.path.dirname(cfg['background']), exist_ok=True)
    img_a.save(cfg['background'])
    return cfg['background']


def build_hoiva_menu_pdf(days, week_number, year, output_path, variant='tammikoti',
                          dessert_day=None, dessert_text=None):
    """Render one of the two real Hoiva-facing weekly menus.

    days: list of 7 dicts {'main': meal-or-None, 'main2': meal-or-None,
      'main3': meal-or-None, 'soup': meal-or-None}, Ma..Su in order — the
      generator's real per-weekday assignment. main2/main3 (optional extra
      main-course picks, added manually per day) are appended to the
      main-course line when present. main3 in practice only ever appears
      on Friday — it's Hoiva-exclusive content that happens to be stored
      on Kesti's Friday row (since Hoiva's ma-pe mirrors Kesti live), which
      is why Kesti's own menu never shows it even though this one does.
    dessert_day: 0-6 (Ma=0..Su=6) — which weekday gets the 3rd
      ('JÄLKIRUOKA') line. Only rendered on the 'asukas' variant, since
      the 'tammikoti' template has no dessert line on any day.
    """
    if variant not in HOIVA_TEMPLATES:
        raise ValueError(f'Tuntematon hoiva-pohja: {variant}')
    cfg = HOIVA_TEMPLATES[variant]
    font, bold = _register_fonts()

    monday = monday_of_week(week_number, year)
    bg = _ensure_hoiva_background(variant)

    c = canvas.Canvas(output_path, pagesize=A4)
    if bg and os.path.exists(bg):
        c.drawImage(ImageReader(bg), 0, 0, width=PAGE_W, height=PAGE_H,
                    preserveAspectRatio=False, mask='auto')

    c.setFillColorRGB(0.05, 0.05, 0.05)
    c.setFont(bold, cfg['subtitle_size'])
    subtitle = cfg['subtitle_format'].format(week=week_number, yy=str(year)[-2:])
    c.drawString(cfg['subtitle_x'], PAGE_H - cfg['subtitle_baseline'], subtitle)

    for i, day in enumerate(HOIVA_DAY_NAMES):
        d = monday + timedelta(days=i)
        label = f'{day} {d.day}.{d.month}'
        c.setFont(bold, HOIVA_DAY_LABEL_SIZE)
        w = stringWidth(label, bold, HOIVA_DAY_LABEL_SIZE)
        box = cfg['day_label_boxes'][i]
        center_x = (box[0] + box[2]) / 2
        c.drawString(center_x - w / 2, PAGE_H - cfg['day_label_baselines'][i], label)

    for i in range(7):
        day_info = days[i] if i < len(days) else {}
        main_line = _fmt_line(day_info.get('main')) or '(PÄÄRUOKA PUUTTUU)'
        for extra_key in ('main2', 'main3'):
            extra = day_info.get(extra_key)
            if extra:
                main_line = f"{main_line}, {_fmt_line(extra)}"
        soup_line = _fmt_line(day_info.get('soup')) or '(KEITTO PUUTTUU)'
        lines = [main_line, soup_line]
        if variant == 'asukas' and dessert_day == i:
            lines.append(f'JÄLKIRUOKA: {dessert_text}'.upper() if dessert_text else 'JÄLKIRUOKA')

        top = cfg['dish_tops'][i]
        for j, line in enumerate(lines):
            y_top = top + j * HOIVA_LINE_GAP
            size = DISH_SIZE
            while stringWidth(line, font, size) > 190 and size > DISH_MIN_SIZE:
                size -= 0.5
            w = stringWidth(line, font, size)
            c.setFont(font, size)
            c.drawString(HOIVA_DISH_CENTER_X - w / 2,
                         PAGE_H - (y_top + HOIVA_GLYPH_TOP_TO_BASELINE), line)

    c.showPage()
    c.save()
    return output_path


# ----------------------------------------------------------------------
# Kymenkartano-branded weekly menu — unlike Kesti/Hoiva, almost all of
# this template's body text is REAL selectable text (Calibri), not vector
# outlines, so the day names/dates/dish lines only need a plain text
# redaction (no graphics-removal pass). Only the header's "Viikko NN/YYYY"
# is a vector outline (same situation as the other two templates' day
# labels) — the "Ruokalista" title itself never changes and stays in the
# background untouched.
# ----------------------------------------------------------------------

KYMENKARTANO_PDF = os.path.join(BASE_DIR, 'templates_src', 'kymenkartano_template.pdf')
KYMENKARTANO_BACKGROUND = os.path.join(BASE_DIR, 'templates_src', 'kymenkartano_background.png')

KYMENKARTANO_HEADER_WEEK_BOX = (25, 60, 220, 93)
KYMENKARTANO_HEADER_X = 29.9
KYMENKARTANO_HEADER_BASELINE = 88.8
KYMENKARTANO_HEADER_SIZE = 24

KYMENKARTANO_DAY_NAMES = ['MA', 'TI', 'KE', 'TO', 'PE', 'LA', 'SU']
KYMENKARTANO_DAY_LABEL_X = 51.1
KYMENKARTANO_DISH_X = 116.3
KYMENKARTANO_BLOCK_TOPS = [124.6, 193.0, 261.3, 329.7, 398.1, 466.4, 534.8]
KYMENKARTANO_LINE_GAP = 17.1
KYMENKARTANO_FONT_SIZE = 12.5
KYMENKARTANO_GLYPH_TOP_TO_BASELINE = 10.2
# Generous per-day redaction box (both columns, up to 4 rows for Sunday's
# extra dessert line) — safe to over-cover since this area only ever
# holds text that we are about to overlay fresh anyway.
KYMENKARTANO_DAY_BOX_HEIGHT = 4 * KYMENKARTANO_LINE_GAP + 10


def _ensure_kymenkartano_background(force=False):
    if os.path.exists(KYMENKARTANO_BACKGROUND) and not force:
        return KYMENKARTANO_BACKGROUND
    if not os.path.exists(KYMENKARTANO_PDF):
        return None

    DPI = 200
    scale = DPI / 72.0

    doc = fitz.open(KYMENKARTANO_PDF)
    page = doc[0]

    # Pass A: remove the real body text (day names/dates/dish lines) —
    # plain text redaction, all graphics/images kept.
    for top in KYMENKARTANO_BLOCK_TOPS:
        box = (45, top - 4, 530, top - 4 + KYMENKARTANO_DAY_BOX_HEIGHT)
        page.add_redact_annot(fitz.Rect(*box))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                          graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    pix_a = page.get_pixmap(dpi=DPI)

    # Pass B: additionally strip the vector-outline week number in the header.
    page.add_redact_annot(fitz.Rect(*KYMENKARTANO_HEADER_WEEK_BOX))
    try:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                              graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
    except TypeError:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    pix_b = page.get_pixmap(dpi=DPI)

    from PIL import Image
    img_a = Image.frombytes('RGB', (pix_a.width, pix_a.height), pix_a.samples)
    img_b = Image.frombytes('RGB', (pix_b.width, pix_b.height), pix_b.samples)
    x0, y0, x1, y1 = KYMENKARTANO_HEADER_WEEK_BOX
    px = (int(x0 * scale), int(y0 * scale), int(x1 * scale) + 1, int(y1 * scale) + 1)
    img_a.paste(img_b.crop(px), px)

    os.makedirs(os.path.dirname(KYMENKARTANO_BACKGROUND), exist_ok=True)
    img_a.save(KYMENKARTANO_BACKGROUND)
    return KYMENKARTANO_BACKGROUND


def build_kymenkartano_menu_pdf(days, week_number, year, output_path,
                                dessert_day=None, dessert_text=None):
    """Render the Kymenkartano weekly menu.

    days: list of 7 dicts {'main', 'main2', 'soup', 'salad'} (meal-or-None),
      Ma..Su in order.
    dessert_day: 0-6 — which weekday gets the extra dessert line (no
      'JÄLKIRUOKA:' prefix here, the template just prints the name itself,
      e.g. 'Jälkiruokakimara'). Line is only rendered when dessert_text is set.
    """
    font, bold = _register_fonts()
    monday = monday_of_week(week_number, year)
    bg = _ensure_kymenkartano_background()

    c = canvas.Canvas(output_path, pagesize=A4)
    if bg and os.path.exists(bg):
        c.drawImage(ImageReader(bg), 0, 0, width=PAGE_W, height=PAGE_H,
                    preserveAspectRatio=False, mask='auto')

    c.setFillColorRGB(1, 1, 1)  # header bar is black — week text is white
    c.setFont(bold, KYMENKARTANO_HEADER_SIZE)
    c.drawString(KYMENKARTANO_HEADER_X, PAGE_H - KYMENKARTANO_HEADER_BASELINE,
                 f'Viikko {week_number}/{year}')

    c.setFillColorRGB(0.05, 0.05, 0.05)  # body text is black on a light page
    for i in range(7):
        d = monday + timedelta(days=i)
        top = KYMENKARTANO_BLOCK_TOPS[i]
        day_info = days[i] if i < len(days) else {}

        c.setFont(bold, KYMENKARTANO_FONT_SIZE)
        c.drawString(KYMENKARTANO_DAY_LABEL_X,
                     PAGE_H - (top + KYMENKARTANO_GLYPH_TOP_TO_BASELINE),
                     KYMENKARTANO_DAY_NAMES[i])
        c.drawString(KYMENKARTANO_DAY_LABEL_X,
                     PAGE_H - (top + KYMENKARTANO_LINE_GAP + KYMENKARTANO_GLYPH_TOP_TO_BASELINE),
                     f'{d.day}.{d.month}.')

        def _line(meal):
            if not meal:
                return None
            codes = extract_menu_codes(meal.get('notes'))
            return f"{meal['name']} {codes}".strip()

        soup_line = _line(day_info.get('soup')) or '(keitto puuttuu)'
        salad = day_info.get('salad')
        if salad:
            salad_line = _line(salad)
            if 'buffet' not in salad_line.lower():
                salad_line += ' & salaattibuffet'
        else:
            salad_line = 'salaattibuffet'
        main_line = _line(day_info.get('main')) or '(pääruoka puuttuu)'
        main2 = day_info.get('main2')
        if main2:
            main_line = f"{main_line}, {_line(main2)}"
        lines = [soup_line, salad_line, main_line]
        if dessert_day == i and dessert_text:
            lines.append(dessert_text)

        c.setFont(font, KYMENKARTANO_FONT_SIZE)
        for j, line in enumerate(lines):
            y_top = top + j * KYMENKARTANO_LINE_GAP
            c.drawString(KYMENKARTANO_DISH_X,
                         PAGE_H - (y_top + KYMENKARTANO_GLYPH_TOP_TO_BASELINE), line)

    c.showPage()
    c.save()
    return output_path


def build_kitchen_instructions_pdf(week_number, day_names, days_data, output_path):
    """Keittiön tulostettava valmistusohjekooste yhdelle viikolle.

    day_names: weekday labels for this facility (5 or 7 entries).
    days_data: one list per day_names entry, each a list of
      {'role_label', 'name_fi', 'instructions', 'ingredients'} dicts (role
      order as assigned by the generator: keitto/salaatti/lounas).
      ingredients: list of {'name_fi', 'quantity', 'unit'} — always shown,
      with the instructions (if any) printed below them.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    font_name, font_bold = 'Helvetica', 'Helvetica-Bold'
    try:
        pdfmetrics.registerFont(TTFont('Amaranth', FONT_REG))
        pdfmetrics.registerFont(TTFont('Amaranth-Bold', FONT_BOLD))
        font_name, font_bold = 'Amaranth', 'Amaranth-Bold'
    except Exception:
        pass

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                             topMargin=2*cm, bottomMargin=2*cm,
                             leftMargin=2*cm, rightMargin=2*cm)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('KitchenTitle', parent=styles['Heading1'],
                                  fontName=font_bold, fontSize=20, textColor='#1a5c3a',
                                  spaceAfter=14)
    day_style = ParagraphStyle('KitchenDay', parent=styles['Heading2'],
                                fontName=font_bold, fontSize=15, textColor='#1a5c3a',
                                spaceBefore=14, spaceAfter=6)
    dish_style = ParagraphStyle('KitchenDish', parent=styles['Heading3'],
                                 fontName=font_bold, fontSize=12, spaceBefore=8, spaceAfter=3)
    label_style = ParagraphStyle('KitchenLabel', parent=styles['Normal'],
                                  fontName=font_bold, fontSize=10, spaceBefore=4, spaceAfter=2)
    body_style = ParagraphStyle('KitchenBody', parent=styles['Normal'],
                                 fontName=font_name, fontSize=10.5, leading=14)

    content = [Paragraph(f'Valmistusohjeet — viikko {week_number}', title_style)]

    for day_name, dishes in zip(day_names, days_data):
        content.append(Paragraph(day_name, day_style))
        if not dishes:
            content.append(Paragraph('Ei aterioita tälle päivälle.', body_style))
            continue
        for dish in dishes:
            label = f"{dish['role_label']}: {dish['name_fi']}" if dish.get('role_label') else dish['name_fi']
            content.append(Paragraph(label, dish_style))

            content.append(Paragraph('Raaka-aineet', label_style))
            ingredients = dish.get('ingredients') or []
            if ingredients:
                items = []
                for ing in ingredients:
                    qty = ing.get('quantity')
                    unit = ing.get('unit') or ''
                    qty_text = f"{qty:g} {unit}".strip() if qty is not None else unit
                    line = f"{qty_text} {ing['name_fi']}".strip() if qty_text else ing['name_fi']
                    items.append(ListItem(Paragraph(line, body_style), leftIndent=6))
                content.append(ListFlowable(items, bulletType='bullet', start='•'))
            else:
                content.append(Paragraph('Raaka-aineita ei ole vielä tallennettu.', body_style))

            content.append(Paragraph('Valmistusohje', label_style))
            if dish.get('instructions'):
                for para in dish['instructions'].split('\n'):
                    if para.strip():
                        content.append(Paragraph(para.strip(), body_style))
            else:
                content.append(Paragraph('Valmistusohjetta ei ole vielä tallennettu.', body_style))
        content.append(Spacer(1, 6))

    doc.build(content)
    return output_path

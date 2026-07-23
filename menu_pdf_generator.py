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

    # Dish lines: soup / salad & buffet / main / sides
    for i, d in enumerate(days):
        salad_line = _fmt_line(d['salad'])
        if salad_line and 'SALAATTIBUFFET' not in salad_line:
            salad_line += ' & SALAATTIBUFFET'
        elif not salad_line:
            salad_line = 'PÄIVÄN SALAATTI & SALAATTIBUFFET'
        lines = [_fmt_line(d['soup']), salad_line,
                 _fmt_line(d['main']), _fmt_line(d['sides'])]
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

"""
OCR import: extract meals from scanned/photographed weekly menu tables.

Works fully offline with Tesseract + Finnish language data.
Handles the printed weekly table format:
    columns = weekdays (Maanantai..Sunnuntai with dates)
    sections per column = Aamiainen / Lounas / Päiväkahvi / Päivällinen / Iltapala

Robust to phone photos: auto-detects rotation by scoring Finnish
menu keywords at 0/90/180/270 degrees, and re-joins dish names that
wrap onto two lines inside a table cell.

System requirements (documented in OHJEET):
    Windows: Tesseract installer (UB Mannheim) + Finnish language pack
    macOS:   brew install tesseract tesseract-lang
    Linux:   apt install tesseract-ocr tesseract-ocr-fin
"""

import re
from collections import defaultdict

DAYS = ['maanantai', 'tiistai', 'keskiviikko', 'torstai', 'perjantai',
        'lauantai', 'sunnuntai']
MEALS = ['aamiainen', 'lounas', 'päiväkahvi', 'päivällinen', 'iltapala']

# Non-recipe staples that appear on every menu — never worth importing
BASICS = {
    'juomapaketti', 'leipä ja margariini', 'leipä', 'margariini',
    'hillo', 'hedelmäsose', 'mehukeitto', 'goudajuusto', 'paprika',
    'tomaatti', 'kurkku', 'maustekurkku', 'keittokinkku', 'kinkkumakkara',
    'balkanmakkara', 'kahvi', 'tee', 'maito',
}

KEYWORDS = DAYS + MEALS + ['juomapaketti', 'margariini', 'vko']


def _tesseract_available():
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        langs = pytesseract.get_languages(config='')
        return 'fin' in langs
    except Exception:
        return False


def _best_rotation(img):
    """Try 4 rotations on a small copy; pick the one with most Finnish keywords."""
    import pytesseract
    small = img.resize((max(1, img.width // 4), max(1, img.height // 4)))
    best_angle, best_score = 0, -1
    for angle in (0, 90, 180, 270):
        text = pytesseract.image_to_string(
            small.rotate(angle, expand=True), lang='fin').lower()
        score = sum(text.count(k) for k in KEYWORDS)
        if score > best_score:
            best_angle, best_score = angle, score
    return best_angle, best_score


def _join_wrapped(dishes):
    """Re-join dish names that wrapped across lines in a table cell.
    A fragment starting lowercase (or 1-3 chars) continues the previous dish."""
    out = []
    for d in dishes:
        d = d.strip()
        if not d:
            continue
        if out and (d[0].islower() or len(d) <= 3):
            out[-1] = out[-1] + d if len(d) <= 3 else out[-1] + ' ' + d
        else:
            out.append(d)
    return out


def extract_week_menu(image_path):
    """
    Parse one scanned weekly menu table.
    Returns: {'week_label': str|None, 'days': {day: {meal: [dishes]}}, 'rotation': int}
    Raises RuntimeError with a clear message if Tesseract/fin is missing.
    """
    if not _tesseract_available():
        raise RuntimeError(
            'Tesseract OCR (suomen kielipaketilla) puuttuu koneelta. '
            'Windows: asenna UB Mannheim -asennuspaketti ja valitse Finnish. '
            'macOS: brew install tesseract tesseract-lang. '
            'Linux: apt install tesseract-ocr tesseract-ocr-fin.')

    import pytesseract
    from PIL import Image, ImageOps

    img = Image.open(image_path)
    img = ImageOps.exif_transpose(img)  # honour EXIF orientation first
    angle, score = _best_rotation(img)
    if angle:
        img = img.rotate(angle, expand=True)
    if score <= 0:
        raise RuntimeError('Kuvasta ei tunnistettu ruokalistataulukkoa. '
                           'Tarkista kuvan laatu ja rajaus.')

    # Half resolution is enough for print and 4x faster on big phone photos
    if img.width > 3000:
        img = img.resize((img.width // 2, img.height // 2))

    data = pytesseract.image_to_data(img, lang='fin',
                                     output_type=pytesseract.Output.DICT)

    # Locate weekday headers -> column boundaries
    day_cols = []
    for i, w in enumerate(data['text']):
        if w.strip().lower() in DAYS and int(data['conf'][i]) > 50:
            day_cols.append((w.strip().capitalize(), data['left'][i]))
    day_cols.sort(key=lambda d: d[1])
    if len(day_cols) < 3:
        raise RuntimeError(f'Vain {len(day_cols)} viikonpäiväotsikkoa löytyi — '
                           'taulukkoa ei voitu jäsentää. Kokeile tarkempaa kuvaa.')

    bounds = []
    for i, (day, x) in enumerate(day_cols):
        left = 0 if i == 0 else (day_cols[i - 1][1] + x) // 2
        right = img.width if i == len(day_cols) - 1 else (x + day_cols[i + 1][1]) // 2
        bounds.append((day, left, right))

    # Week label (e.g. "Vko 36") if present anywhere
    full_text = ' '.join(w for w in data['text'] if w.strip())
    m = re.search(r'[Vv]ko\s*(\d{1,2})', full_text)
    week_label = f'Vko {m.group(1)}' if m else None

    # Assign words to columns
    col_words = defaultdict(list)
    for i, w in enumerate(data['text']):
        if not w.strip() or int(data['conf'][i]) < 30:
            continue
        x, y = data['left'][i], data['top'][i]
        for day, l, r in bounds:
            if l <= x < r:
                col_words[day].append((y, x, w.strip()))
                break

    # Per column: cluster words into lines, split into meal sections
    days_out = {}
    line_tol = max(8, img.height // 250)
    for day, l, r in bounds:
        words = sorted(col_words[day])
        lines, cur, cur_y = [], [], None
        for y, x, w in words:
            if cur_y is None or abs(y - cur_y) <= line_tol:
                cur.append((x, w))
                cur_y = y if cur_y is None else cur_y
            else:
                lines.append(' '.join(t for _, t in sorted(cur)))
                cur, cur_y = [(x, w)], y
        if cur:
            lines.append(' '.join(t for _, t in sorted(cur)))

        sections, current = {}, None
        for line in lines:
            key = line.strip().lower().rstrip(':')
            if key in MEALS:
                current = key
                sections[current] = []
            elif current is not None:
                t = line.strip()
                if (t.lower() not in DAYS and
                        not re.match(r'^\d{2}\.\d{2}\.\d{4}$', t) and
                        not re.match(r'^[Vv]ko\b', t)):
                    sections[current].append(t)
        days_out[day] = {meal: _join_wrapped(d) for meal, d in sections.items()}

    return {'week_label': week_label, 'days': days_out, 'rotation': angle}


def menu_to_queue_items(parsed, source_name, main_meals_only=True):
    """
    Convert a parsed week menu into review-queue items (unique dishes).
    main_meals_only: keep only Lounas + Päivällinen (skip breakfast/coffee/snack
    staples). Basics like 'Juomapaketti' are always skipped.
    """
    from recipe_scraper_v2 import guess_category, guess_season
    from datetime import datetime

    wanted = {'lounas', 'päivällinen'} if main_meals_only else set(MEALS)
    seen, items = set(), []
    for day, meals in parsed['days'].items():
        for meal, dishes in meals.items():
            if meal not in wanted:
                continue
            for dish in dishes:
                name = dish.strip()
                key = name.lower()
                if not name or key in BASICS or key in seen or len(name) < 4:
                    continue
                seen.add(key)
                items.append({
                    'name_fi': name,
                    'source_url': '',
                    'source_site': f'OCR-tuonti ({source_name})',
                    'ingredients_raw': [],
                    'instructions_raw': '',
                    'season': guess_season(name, '') or 'kaikki',
                    'meal_type': 'lounas' if meal == 'lounas' else 'päivällinen',
                    'dish_category': guess_category(name, ''),
                    'prep_time_min': None,
                    'difficulty': None,
                    'servings': None,
                    'notes': f'{parsed.get("week_label") or ""} {day} {meal}'.strip(),
                    'scraped_at': datetime.now().isoformat(),
                })
    return items

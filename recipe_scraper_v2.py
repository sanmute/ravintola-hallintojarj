"""
Recipe Scraper v2 - JSON-LD based (production version)

Instead of guessing CSS selectors, this parses schema.org Recipe markup
(JSON-LD), which Ruoka.fi, K-Ruoka, Valio, Kotikokki and most modern
recipe sites embed in every recipe page. Far more robust against
site redesigns.

Usage on your Windows machine (needs internet access):
    python recipe_scraper_v2.py <url1> <url2> ...
    python recipe_scraper_v2.py --file urls.txt

Output: recipes_for_review.json -> handled in the web app's
"Tarkistusjono" (Review queue) tab.
"""

import sys
import json
import re
import time
from datetime import datetime

import requests
from bs4 import BeautifulSoup

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0 Safari/537.36',
    'Accept-Language': 'fi-FI,fi;q=0.9,en;q=0.8',
}

REVIEW_FILE = 'recipes_for_review.json'

# Sites do not always expose recipeIngredient in their JSON-LD.  These are
# deliberately Finnish-first, with English/Swedish variants for recipes that
# are published in more than one language.
INGREDIENT_HEADINGS = {
    # _normalise_heading() strips spaces AND hyphens, so entries here must
    # already be in that stripped form (a bug fix — 'raaka-aineet', an
    # extremely common Finnish heading, never matched before this because
    # the heading text got its hyphen stripped while this literal didn't).
    'ainekset', 'ainesosat', 'raakaaineet', 'tarvitset', 'tarvikkeet',
    'ingredients', 'ingredienter', 'ingredienser',
}

# Same idea for the instructions section — sites vary in wording. Also
# pre-stripped of spaces/hyphens to match what _normalise_heading() produces.
INSTRUCTION_HEADINGS = {
    'ohje', 'ohjeet', 'valmistusohje', 'valmistusohjeet', 'valmistus',
    'vaiheet', 'tekoohje', 'näinteetsen', 'näinvalmistat', 'valmistaminen',
    'instructions', 'directions', 'method',
}


def _normalise_heading(value):
    """Make a section heading suitable for a small, language-aware lookup."""
    return re.sub(r'[^a-zåäö]', '', value.casefold())


def extract_ingredients_from_html(soup):
    """Read an ingredient list that is presented as normal HTML.

    JSON-LD remains the preferred source.  This fallback covers recipe pages
    whose ingredients only appear below headings such as ``Ainekset`` or
    ``Ainesosat``.
    """
    for heading in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'strong', 'b']):
        if _normalise_heading(heading.get_text(' ', strip=True)) not in INGREDIENT_HEADINGS:
            continue

        # The usual markup is "<h2>Ainekset</h2><ul><li>...</li>...".
        # Looking at following siblings prevents accidentally collecting page
        # navigation or the instructions section further down the page.
        for sibling in heading.find_all_next(['ul', 'ol', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'], limit=8):
            if sibling is not heading and sibling.name.startswith('h'):
                break
            if sibling.name in ('ul', 'ol'):
                lines = [li.get_text(' ', strip=True) for li in sibling.find_all('li')]
                lines = [line for line in lines if line]
                if lines:
                    return lines
            if sibling.name == 'p':
                text = sibling.get_text(' ', strip=True)
                if text and re.search(r'\d|\bkg\b|\bdl\b|\brkl\b', text, re.I):
                    return [text]

    # Some sites use a semantic class but no visible heading.
    for container in soup.select('[class*="ingredient" i], [id*="ingredient" i], '
                                 '[class*="aines" i], [id*="aines" i]'):
        lines = [li.get_text(' ', strip=True) for li in container.find_all('li')]
        lines = [line for line in lines if line]
        if lines:
            return lines
    return []


def extract_instructions_from_html(soup):
    """Read cooking instructions presented as normal HTML, below a heading
    like 'Ohje', 'Valmistusohje' or 'Vaiheet'.

    JSON-LD (recipeInstructions) remains the preferred source — this is a
    fallback for pages that either lack Recipe JSON-LD entirely or have it
    without the instructions field populated.
    """
    for heading in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'strong', 'b']):
        if _normalise_heading(heading.get_text(' ', strip=True)) not in INSTRUCTION_HEADINGS:
            continue

        for sibling in heading.find_all_next(['ul', 'ol', 'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'], limit=12):
            if sibling is not heading and sibling.name.startswith('h'):
                break
            if sibling.name in ('ul', 'ol'):
                lines = [li.get_text(' ', strip=True) for li in sibling.find_all('li')]
                lines = [line for line in lines if line]
                if lines:
                    return lines
            if sibling.name in ('p', 'div'):
                text = sibling.get_text(' ', strip=True)
                # A real instruction step reads as a sentence — skip short
                # fragments (nav labels, empty spacer divs, etc.)
                if text and len(text) > 15:
                    return [text]

    # Some sites use a semantic class but no visible heading.
    for container in soup.select('[class*="instruction" i], [id*="instruction" i], '
                                 '[class*="valmistus" i], [id*="valmistus" i], '
                                 '[class*="ohje" i], [id*="ohje" i]'):
        lines = [li.get_text(' ', strip=True) for li in container.find_all('li')]
        lines = [line for line in lines if line]
        if lines:
            return lines
        text = container.get_text(' ', strip=True)
        if text and len(text) > 15:
            return [text]
    return []


def _parse_amount(value):
    """Parse common Finnish decimal and fraction notations into a float."""
    value = value.strip().replace(',', '.')
    fractions = {'½': .5, '¼': .25, '¾': .75, '⅓': 1 / 3, '⅔': 2 / 3}
    if value in fractions:
        return fractions[value]
    for char, amount in fractions.items():
        if value.endswith(char) and value[:-1].isdigit():
            return float(value[:-1]) + amount
    if '/' in value:
        try:
            if ' ' in value:
                whole, fraction = value.split(None, 1)
                numerator, denominator = fraction.split('/', 1)
                return float(whole) + float(numerator) / float(denominator)
            numerator, denominator = value.split('/', 1)
            return float(numerator) / float(denominator)
        except (ValueError, ZeroDivisionError):
            return None
    try:
        return float(value)
    except ValueError:
        return None


def structure_ingredients(lines):
    """Turn scraped ingredient text into the format used by the database."""
    structured = []
    pattern = re.compile(
        r'^\s*(?P<qty>(?:\d+(?:[.,]\d+)?|\d+\s+\d+/\d+|\d+/\d+|[½¼¾⅓⅔]|\d+[½¼¾⅓⅔]))'
        r'\s*(?P<unit>kg|g|mg|l|dl|cl|ml|tl|rkl|kpl|pkt|pss|prk|rs|purkki|rasia|nippu|ripaus)?\.?'
        r'\s+(?P<name>.+?)\s*$', re.I)
    for line in lines:
        text = re.sub(r'^\s*[-•*]\s*', '', str(line)).strip()
        if not text:
            continue
        match = pattern.match(text)
        if match:
            quantity = _parse_amount(match.group('qty'))
            name = match.group('name').strip(' ,-')
            unit = (match.group('unit') or 'kpl').lower()
        else:
            # Retain an unquantified ingredient instead of silently dropping it.
            quantity, unit, name = None, 'kpl', text
        if name:
            structured.append({'name': name, 'qty': quantity, 'unit': unit})
    return structured


SERVINGS_PATTERNS = [
    # "5 annosta", "5 annosta", "4 hengelle" (Finnish)
    re.compile(r'\b(\d+)\s*(?:annos(?:ta|ia)?|hengelle)\b', re.I),
    # "Annosta: 5", "Annosmäärä: 5", "Annokset: 5"
    re.compile(r'\bannos(?:ta|ia|määrä|et)?\s*[:\-]?\s*(\d+)\b', re.I),
    # "Serves 5", "5 servings" (English) / "4 portioner" (Swedish) — bonus,
    # recipe sites are occasionally published in more than one language
    re.compile(r'\bserves?\s*[:\-]?\s*(\d+)\b', re.I),
    re.compile(r'\b(\d+)\s*(?:servings|portioner)\b', re.I),
]


def extract_servings_from_html(soup):
    """Fallback for pages whose JSON-LD lacks recipeYield: look for a plain
    portions statement like "5 Annosta" anywhere on the page."""
    text = soup.get_text(' ', strip=True)
    for pattern in SERVINGS_PATTERNS:
        m = pattern.search(text)
        if m:
            try:
                n = int(m.group(1))
                if 0 < n <= 200:  # sanity bound — rules out stray numbers (phone, year, etc.)
                    return n
            except ValueError:
                continue
    return None


def parse_iso_duration(value):
    """PT45M / PT1H30M -> minutes"""
    if not value or not isinstance(value, str):
        return None
    m = re.match(r'PT(?:(\d+)H)?(?:(\d+)M)?', value)
    if not m:
        return None
    hours = int(m.group(1) or 0)
    mins = int(m.group(2) or 0)
    total = hours * 60 + mins
    return total or None


def find_recipe_jsonld(soup):
    """Find a schema.org Recipe object in any JSON-LD block on the page."""
    for script in soup.find_all('script', type='application/ld+json'):
        try:
            data = json.loads(script.string or '')
        except (json.JSONDecodeError, TypeError):
            continue
        candidates = data if isinstance(data, list) else [data]
        # @graph wrapper is common
        expanded = []
        for c in candidates:
            if isinstance(c, dict) and '@graph' in c:
                expanded.extend(c['@graph'])
            else:
                expanded.append(c)
        for obj in expanded:
            if not isinstance(obj, dict):
                continue
            obj_type = obj.get('@type', '')
            types = obj_type if isinstance(obj_type, list) else [obj_type]
            if 'Recipe' in types:
                return obj
    return None


def guess_category(name, ingredients_text):
    """Best-effort auto-categorization; reviewed by a human anyway."""
    text = (name + ' ' + ingredients_text).lower()
    fish_words = ['kala', 'lohi', 'silakka', 'kuha', 'ahven', 'siika', 'tonnikala',
                  'sei', 'muikku', 'katkarapu', 'silli']
    chicken_words = ['kana', 'broileri', 'kalkkuna']
    veg_markers = ['kasvis', 'vegaani', 'vege', 'linssi', 'soija', 'tofu',
                   'härkis', 'nyhtökaura', 'quorn', 'halloumi']
    meat_words = ['jauheliha', 'nauta', 'porsa', 'possu', 'sika', 'liha',
                  'makkara', 'kinkku', 'pekoni', 'nakk', 'maksa', 'lammas',
                  'poro', 'riista', 'karjalanpaisti']
    if any(w in text for w in fish_words):
        return 'kala'
    if any(w in text for w in chicken_words):
        return 'kana'
    if any(w in text for w in meat_words):
        return 'naudanliha'
    if any(w in text for w in veg_markers):
        return 'kasvis'
    return None



def guess_season(name, ingredients_text):
    """
    Guess season from seasonal Finnish ingredients/keywords.
    Returns None when unsure (most dishes are year-round) -
    a human confirms in the review queue.
    """
    text = (name + ' ' + ingredients_text).lower()
    spring = ['parsa', 'raparperi', 'nokkonen', 'varhaiskaali', 'kevät',
              'pääsiäis', 'mämmi', 'villiyrt']
    summer = ['uudet perunat', 'uusi peruna', 'mansikka', 'kesäkurpitsa',
              'varhaisperuna', 'grilli', 'grillattu', 'kesä', 'juhannus',
              'tuore herne', 'mustikka', 'vadelma', 'rapu']
    autumn = ['kurpitsa', 'sieni', 'kantarelli', 'suppilovahvero', 'herkkutatti',
              'riista', 'hirvi', 'poro', 'syksy', 'puolukka', 'omena',
              'luumu', 'juures', 'lanttu', 'palsternakka']
    winter = ['joulu', 'kinkku', 'rosolli', 'lipeäkala', 'talvi', 'laskiais',
              'hernekeitto', 'uuni juures', 'kaalikääryle']
    scores = {
        'kevät': sum(w in text for w in spring),
        'kesä': sum(w in text for w in summer),
        'syksy': sum(w in text for w in autumn),
        'talvi': sum(w in text for w in winter),
    }
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else None


def scrape_recipe(url):
    """Scrape a single recipe page via JSON-LD."""
    resp = requests.get(url, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.content, 'html.parser')

    recipe = find_recipe_jsonld(soup)
    if not recipe:
        print(f"  [skip] Ei Recipe-merkintää (JSON-LD): {url}")
        return None

    name = recipe.get('name', '').strip()
    ingredients = recipe.get('recipeIngredient') or recipe.get('ingredients') or []
    if isinstance(ingredients, str):
        ingredients = [ingredients]
    ingredients = [str(item).strip() for item in ingredients if str(item).strip()]
    if not ingredients:
        ingredients = extract_ingredients_from_html(soup)

    instructions = recipe.get('recipeInstructions', '')
    if isinstance(instructions, list):
        parts = []
        for step in instructions:
            if isinstance(step, dict):
                parts.append(step.get('text', ''))
            else:
                parts.append(str(step))
        instructions = ' '.join(p for p in parts if p)
    if not instructions:
        # JSON-LD had no recipeInstructions (common — many sites only put
        # ingredients in structured data) — look for a heading like
        # "Ohje"/"Valmistusohje"/"Vaiheet" and take what follows it.
        instructions = ' '.join(extract_instructions_from_html(soup))

    prep = parse_iso_duration(recipe.get('totalTime')) or \
           parse_iso_duration(recipe.get('cookTime')) or \
           parse_iso_duration(recipe.get('prepTime'))

    servings = recipe.get('recipeYield')
    if isinstance(servings, list):
        servings = servings[0] if servings else None
    if isinstance(servings, str):
        m = re.search(r'(\d+)', servings)
        servings = int(m.group(1)) if m else None
    if not servings:
        # JSON-LD didn't have it — look for a plain-text statement like "5 Annosta"
        servings = extract_servings_from_html(soup)

    return {
        'name_fi': name,
        'source_url': url,
        'source_site': url.split('/')[2].replace('www.', ''),
        'ingredients_raw': ingredients,
        'ingredients_struct': structure_ingredients(ingredients),
        'instructions_raw': instructions[:2000],
        'season': guess_season(name, ' '.join(ingredients)) or 'kaikki',
        'meal_type': 'lounas',
        'dish_category': guess_category(name, ' '.join(ingredients)) or 'kasvis',
        'prep_time_min': prep,
        'difficulty': None,
        'servings': servings,
        'notes': '',
        'scraped_at': datetime.now().isoformat(),
    }


def collect_recipe_links(listing_url, limit=30):
    """From a category/search page, collect links that look like recipe pages."""
    resp = requests.get(listing_url, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.content, 'html.parser')
    base = '/'.join(listing_url.split('/')[:3])
    links = set()
    for a in soup.find_all('a', href=True):
        href = a['href']
        if any(token in href for token in ('/resepti', '/reseptit/', '/recipe')):
            full = href if href.startswith('http') else base + href
            # skip listing pages themselves
            if full.rstrip('/') != listing_url.rstrip('/'):
                links.add(full.split('?')[0])
        if len(links) >= limit:
            break
    return sorted(links)


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return

    urls = []
    if args[0] == '--file':
        with open(args[1], encoding='utf-8') as f:
            urls = [line.strip() for line in f if line.strip() and not line.startswith('#')]
    elif args[0] == '--listing':
        # e.g. python recipe_scraper_v2.py --listing https://www.ruoka.fi/reseptit/keitot
        print(f"Haetaan reseptilinkit sivulta: {args[1]}")
        urls = collect_recipe_links(args[1], limit=int(args[2]) if len(args) > 2 else 30)
        print(f"Löytyi {len(urls)} linkkiä")
    else:
        urls = args

    # Load existing review queue, append new
    try:
        with open(REVIEW_FILE, encoding='utf-8') as f:
            results = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        results = []

    existing_urls = {r.get('source_url') for r in results}
    ok, failed = 0, 0

    for url in urls:
        if url in existing_urls:
            print(f"  [skip] Jo jonossa: {url}")
            continue
        try:
            recipe = scrape_recipe(url)
            if recipe and recipe['name_fi']:
                results.append(recipe)
                cat = recipe['dish_category'] or '?'
                sea = recipe['season'] or 'kaikki?'
                print(f"  [ok]   {recipe['name_fi']}  (kategoria: {cat}, kausi: {sea})")
                ok += 1
            else:
                failed += 1
        except Exception as e:
            print(f"  [fail] {url}: {e}")
            failed += 1
        time.sleep(1.0)  # be polite to the server

    with open(REVIEW_FILE, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"\nValmis: {ok} uutta reseptiä, {failed} epäonnistui.")
    print(f"Jonossa nyt yhteensä: {len(results)} reseptiä -> {REVIEW_FILE}")
    print("Avaa sovelluksen Tarkistusjono-välilehti luokitellaksesi ja viedäksesi ne tietokantaan.")


if __name__ == '__main__':
    main()

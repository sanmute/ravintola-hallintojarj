"""
recipe_classifier.py — Tunnista PoweResta-tuonnissa mukaan tulleet
ruoka-aineet, jotka eivät kuulu pääruoka-, keitto- tai salaattikiertoon
(jälkiruoat, juomat, välipalat, pelkät lisäkkeet). Käytetään vain
esihenkilön luku-/poisto-tarkistukseen — ei vaikuta generaattoriin.

Keitto- ja salaattinimiset reseptit jätetään kokonaan tarkistuksen
ulkopuolelle: ne ovat jo omassa, toimivaksi todetussa kierrossaan
(ks. meal_plan_generator.py SOUP_KEYWORD/SALAD_KEYWORD), joten niitä ei
pidä koskaan merkitä poistettavaksi riippumatta siitä, mitä muita sanoja
niiden nimessä sattuu olemaan (esim. "Karpalomehukeitto" sisältää "mehu",
mutta on keitto).
"""

import re

SOUP_KEYWORD = 'keitto'
SALAD_KEYWORD = 'salaa'

DESSERT_KEYWORDS = [
    'kiisseli', 'mousse', 'muffini', 'keksi', 'kuorrute',
    'leivonnainen', 'suklaabiskvii', 'panna cotta', 'pannacotta',
    'jäätelö', 'rahka', 'voipulla', 'marmelaadi', 'hilloke',
]

SNACK_KEYWORDS = [
    'pulla', 'leivos', 'mutakeksi', 'pipari',
    'popcorn', 'sipsit', 'pähkinät', 'karamelli', 'suklaatanko',
]

# Whole-word matches only (\b...\b) — many of these are short enough that a
# plain substring check would false-positive inside unrelated compound words
# (e.g. 'tee' inside 'gluteeniton', 'mehu' inside 'karpalomehukeitto' — the
# soup/salad exclusion above already handles the latter case too).
DRINK_WORDS = [
    'mehu', 'juoma', 'kahvi', 'tee', 'maito', 'smoothie',
    'limonaadi', 'viina', 'olut', 'viini', 'glögi',
]

# Exact (whole, trimmed) name matches only — NOT substring. Words like
# 'pasta', 'riisi', 'perunasose' appear constantly inside real main-course
# casserole names ("Jauheliha Perunasosevuoka", "Lohi-pastavuoka",
# "Kana-pastasalaatti"), so substring-matching them would flag real mains
# (and already-classified salads) for deletion.
SIDE_ONLY_EXACT_NAMES = {
    'perunasose', 'perunamuusi', 'riisi', 'pasta', 'leipä',
    'kaurapuuro', 'mysli', 'puuro',
}


def classify_recipe(name_fi):
    """Classify a recipe name for the non-main-dish audit.

    Returns one of: 'dessert', 'drink', 'snack', 'side_only', 'unknown',
    or None if the recipe is a soup/salad already in active rotation and
    should be excluded from the audit entirely.
    """
    name_lower = name_fi.lower()

    if SOUP_KEYWORD in name_lower or SALAD_KEYWORD in name_lower:
        return None

    for kw in DESSERT_KEYWORDS:
        if kw in name_lower:
            return 'dessert'

    for kw in SNACK_KEYWORDS:
        if kw in name_lower:
            return 'snack'

    for word in DRINK_WORDS:
        if re.search(rf'\b{re.escape(word)}\b', name_lower):
            return 'drink'

    if name_lower.strip() in SIDE_ONLY_EXACT_NAMES:
        return 'side_only'

    return 'unknown'

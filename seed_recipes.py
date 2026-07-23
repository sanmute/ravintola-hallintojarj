"""
Seed the database with real recipes from the existing meal plan.
These are the dishes visible in the current rotation - gives the system
a working recipe base from day one, before any scraping.
"""

from meal_plan_db import MealPlanDB

# (name, dish_category, default season list)
# Categories: kala=fish(blue), kana=chicken(yellow), naudanliha=meat(red), kasvis=vegetable(green)
SEED_RECIPES = [
    # Soups (keitto)
    ("Kirkas kalakeitto", "kala"),
    ("Kasvis-juustokeitto", "kasvis"),
    ("Borssikeitto", "naudanliha"),
    ("Jauhelihakeitto", "naudanliha"),
    ("Bataattisosekeitto", "kasvis"),
    ("Porsaanlihakeitto", "naudanliha"),
    ("Kasvissuikalekeitto", "kasvis"),
    ("Paprika kanakeitto", "kana"),
    ("Mausteinen kalaseljanka", "kala"),
    ("Mustajuurisosekeitto", "kasvis"),
    ("Curry kanakeitto", "kana"),
    ("Chorizo-papukeitto", "naudanliha"),
    ("Nakkikeitto", "naudanliha"),
    ("Lihakeitto", "naudanliha"),
    ("Tomaattinen jauhelihakeitto", "naudanliha"),
    ("Ribollita italialainen kasviskeitto", "kasvis"),
    ("Pinaattikeitto", "kasvis"),
    ("Hapankaalikeitto", "naudanliha"),
    ("Tomaatti-vuohenjuustokeitto", "kasvis"),
    ("Jauheliha-papukeitto", "naudanliha"),
    ("Kanakeitto", "kana"),
    ("Juustoinen savukalakeitto", "kala"),
    ("Irlantilainen lihakeitto (beef stew)", "naudanliha"),
    ("Hernekeitto", "naudanliha"),
    ("Juustoinen kanakeitto", "kana"),
    ("Samettinen kanakeitto", "kana"),
    ("Kirkas kasviskeitto", "kasvis"),
    ("Itämainen kanakeitto", "kana"),
    ("Kukkakaalisosekeitto", "kasvis"),
    ("Gulassikeitto", "naudanliha"),
    ("Kahden kalan keitto", "kala"),
    ("Kermainen kalakeitto", "kala"),
    ("Casablankan kanakeitto", "kana"),
    ("Linssikeitto", "kasvis"),
    ("Siskonmakkarakeitto", "naudanliha"),
    ("Purjo-perunasosekeitto", "kasvis"),
    ("Juustoinen riistakeitto", "naudanliha"),
    ("Parsakaalikeitto", "kasvis"),
    ("Tex mex jauhelihakeitto", "naudanliha"),
    ("Tomaattikeitto", "kasvis"),
    ("Makkarakeitto", "naudanliha"),
    ("Curry-jauhelihakeitto", "naudanliha"),
    # Main dishes
    ("Lihapullat", "naudanliha"),
    ("Stroganoff", "naudanliha"),
    ("Paistettua kalaa", "kala"),
    ("Kalapyörykät", "kala"),
    ("Kermainen (Igor) kanapata", "kana"),
    ("Vuohenjuusto-broileri & pekoninakit", "kana"),
    ("Broilerikiusaus", "kana"),
    ("Savulohikiusaus", "kala"),
    ("Jauheliha-perunalaatikko", "naudanliha"),
    ("Lihamureke", "naudanliha"),
    ("Kanapaella ja kasvispihvit", "kana"),
    ("Kanapaella", "kana"),
    ("Kalamurekepihvit", "kala"),
    ("Paistettu maksa", "naudanliha"),
    ("Maksapihvit", "naudanliha"),
    ("Tortillat", "naudanliha"),
    ("Jauhelihakastike", "naudanliha"),
    ("Italianpata", "naudanliha"),
    ("Kasvis-pastavuoka", "kasvis"),
    ("Broileripyörykät", "kana"),
    ("Uunimakkara", "naudanliha"),
    ("Silakkapihvit", "kala"),
    ("Mary me chicken", "kana"),
    ("Aurajuustopossu ja broilerin nuijat", "kana"),
    ("Aurajuustopossu", "naudanliha"),
    ("Quorn kasviskiusaus", "kasvis"),
    ("Merimiespihvit", "naudanliha"),
    ("Jauhelihakastike (bolognese)", "naudanliha"),
    ("Broileri wok", "kana"),
    ("Kalaleike", "kala"),
    ("Kalapuikot", "kala"),
    ("Stifado kreikkalainen lihapata", "naudanliha"),
    ("Leikepäivä (kana ja possu)", "kana"),
    ("Broilerinleike", "kana"),
    ("Nakkikastike", "naudanliha"),
    ("Lasagnette", "naudanliha"),
    ("Makkara pannu", "naudanliha"),
    ("Makkarakastike", "naudanliha"),
    ("Karjalanpaisti", "naudanliha"),
    ("Tandoorin puna-ahven", "kala"),
    ("Kananpoikaa viinissä", "kana"),
    ("Porsaanposkea ja riistapyörykät", "naudanliha"),
    ("Riistapyörykät", "naudanliha"),
    ("Kaalilaatikko", "naudanliha"),
    ("Janssonin kiusaus", "kala"),
    ("Broilerimureke", "kana"),
    ("Lasagne", "naudanliha"),
    ("Uunikala", "kala"),
    ("Mantelikala", "kala"),
    ("Burgundinpata", "naudanliha"),
    ("Pizza", "naudanliha"),
    ("Makaroni-laatikko", "naudanliha"),
    ("Broileri pastavuoka", "kana"),
    ("Kinkkukiusaus", "naudanliha"),
    ("Hampurilaiset", "naudanliha"),
]


def seed_database(db_path="meal_plans.db", season="kaikki"):
    """Insert seed recipes for a given season. Safe to re-run (duplicates skipped)."""
    db = MealPlanDB(db_path)
    added, skipped = 0, 0
    for name, category in SEED_RECIPES:
        recipe_id = db.add_recipe(
            name_fi=name,
            season=season,
            meal_type='lounas',
            dish_category=category,
            notes='Seeded from existing meal plan'
        )
        if recipe_id:
            added += 1
        else:
            skipped += 1
    print(f"Season '{season}': {added} added, {skipped} skipped (already exist)")
    return added, skipped


if __name__ == '__main__':
    import sys
    season = sys.argv[1] if len(sys.argv) > 1 else 'kaikki'
    seed_database(season=season)

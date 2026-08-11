"""
Meal Plan Generation Algorithm
Generates balanced meal plans with controlled repetition
"""

from meal_plan_db import MealPlanDB, Season, MealType, DishCategory
from collections import defaultdict
from datetime import datetime, timedelta
import random

class MealPlanGenerator:
    """Generates meal plans with constraints"""
    
    SOUP_MIN_GAP_WEEKS = 2   # don't repeat the same soup within 2 calendar weeks
    DAYS_PER_WEEK = 5        # Ma-Pe — each weekday gets its own soup+main+salad
    MAIN_CATEGORY_CAP_PER_WEEK = 2  # no protein category more than 2x among the week's 5 mains
    CYCLE_WEEKS = 6          # a year plan generates this many unique weeks per season,
                             # then repeats them for the rest of that season's calendar
                             # weeks — a fixed repeating menu cycle, as real institutional
                             # catering typically runs, instead of 13 fully unique weeks.

    # Kesti's Friday salad is always this fixed buffet line (confirmed from
    # the restaurant's real weekly menu, e.g. "BATAATTISOSEKEITTO / VÄRIKÄS
    # JA RAIKAS SALAATTIBUFFET / AURAJUUSTOPOSSUA...") rather than whatever
    # the normal salad rotation would have picked. Hoiva inherits this for
    # free since its ma-pe mirrors Kesti live; Kymenkartano is unaffected
    # (fully independent plan, not in scope here).
    FRIDAY_SALAD_NAME = 'Värikäs ja raikas salaattibuffet'

    def __init__(self, db_path='meal_plans.db', days_per_week=None, facility='kesti'):
        """days_per_week: overrides the class default (5, Ma-Pe) — kesti stays
        5, hoiva and kymenkartano use 7 (Ma-Su), each generated separately
        from the same shared recipe pool.
        facility: only used to decide whether the fixed Friday-salad rule
        (see FRIDAY_SALAD_NAME) applies — that rule is Kesti-only."""
        self.db = MealPlanDB(db_path)
        self.facility = facility
        if days_per_week is not None:
            self.DAYS_PER_WEEK = days_per_week
        self._friday_salad = None  # lazily created/looked up, cached per instance
        self.max_repetition_per_cycle = 2  # Max 2x per 4-week cycle
        self.cycle_length = 7  # 7 days per week
        self.weeks_per_cycle = 4
        # Rotation cursor per season, so salad assignment cycles evenly through
        # that season's salad pool instead of being picked at random alongside
        # the main dishes (where a handful of salads would otherwise get
        # over-selected purely by chance). Persists across calls on the same
        # generator instance, so a year plan's talvi segments (which wrap
        # around the calendar year) continue the same rotation rather than
        # restarting it.
        self._salad_rotation = defaultdict(int)
        # Last absolute calendar week each soup was used, so a soup isn't
        # repeated within SOUP_MIN_GAP_WEEKS even across a year plan's
        # separate seasonal segments.
        self._soup_last_used = {}

    # recipe_type values that are never eligible for a weekday soup/main/salad
    # slot (bakery, dessert) — kept out of the mains pool at every split step,
    # since "mains" here really means "everything not yet claimed by a more
    # specific split" (soups are peeled off separately, in _split_soups).
    _NON_MEAL_SLOT_TYPES = ('leivonta', 'jälkiruoka', 'kastike', 'kasvislisäke', 'energialisäke')

    def _split_salads(self, recipe_pool):
        """Split a recipe pool into (mains, salads), using the explicit
        recipes.recipe_type column (r[3]) rather than name matching, and
        sorted by id for a stable rotation order."""
        salads = [r for r in recipe_pool if r[3] == 'salaatti']
        mains = [r for r in recipe_pool
                 if r[3] != 'salaatti' and r[3] not in self._NON_MEAL_SLOT_TYPES]
        salads.sort(key=lambda r: r[0])
        return mains, salads

    def _select_salad(self, salads, season):
        """Round-robin through the season's salad pool — one salad per week,
        never the same one twice in a row, evenly distributed over time."""
        idx = self._salad_rotation[season] % len(salads)
        self._salad_rotation[season] += 1
        return salads[idx]

    def _get_friday_salad(self):
        """Look up (or create, on first use) the fixed Friday-buffet recipe.
        Cached on the instance since it's the same row for every Friday of
        every week generated in one call."""
        if self._friday_salad is not None:
            return self._friday_salad
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        row = conn.execute('SELECT id, name_fi, dish_category, recipe_type FROM recipes WHERE name_fi = ?',
                          (self.FRIDAY_SALAD_NAME,)).fetchone()
        conn.close()
        if row:
            self._friday_salad = tuple(row)
            return self._friday_salad
        recipe_id = self.db.add_recipe(
            name_fi=self.FRIDAY_SALAD_NAME, season='kaikki', meal_type='lounas',
            dish_category='kasvis', recipe_type='salaatti')
        if recipe_id is None:
            # Name already existed after all (e.g. inserted moments ago) —
            # re-fetch instead of trusting a None id.
            conn = sqlite3.connect(self.db.db_path)
            row = conn.execute('SELECT id, name_fi, dish_category, recipe_type FROM recipes WHERE name_fi = ?',
                              (self.FRIDAY_SALAD_NAME,)).fetchone()
            conn.close()
            recipe_id = row[0]
        self._friday_salad = (recipe_id, self.FRIDAY_SALAD_NAME, 'kasvis', 'salaatti')
        return self._friday_salad

    def _split_soups(self, recipe_pool):
        """Split a recipe pool into (mains, soups), using the explicit
        recipes.recipe_type column (r[3]) rather than name matching."""
        soups = [r for r in recipe_pool if r[3] == 'keitto']
        mains = [r for r in recipe_pool if r[3] != 'keitto']
        return mains, soups

    def _select_soup(self, soups, absolute_week):
        """Pick a soup not used within the last SOUP_MIN_GAP_WEEKS calendar
        weeks. Falls back to the full pool if every soup was recently used
        (e.g. a very small seasonal pool)."""
        candidates = [
            s for s in soups
            if absolute_week - self._soup_last_used.get(s[0], -999) >= self.SOUP_MIN_GAP_WEEKS
        ]
        if not candidates:
            candidates = soups
        recipe = random.choice(candidates)
        self._soup_last_used[recipe[0]] = absolute_week
        return recipe
    
    def generate_meal_plan(self, season, num_weeks=52, only_new=False, facility='kesti'):
        """
        Generate a complete meal plan: each of the DAYS_PER_WEEK weekdays gets
        its own main dish + soup + salad (independent rotations).

        Args:
            season: Season name (e.g., 'talvi', 'kevät', 'kesä', 'syksy')
            num_weeks: Total number of weeks (default 52)
            only_new: restrict to hand-added/edited recipes only
            facility: 'kesti', 'hoiva' or 'kymenkartano' — each generated and
                stored as its own independent meal plan, from the same shared
                recipe pool (only DAYS_PER_WEEK differs by facility).

        Returns:
            meal_plan_id if successful, None if failed
        """

        # Get available recipes for this season (lunch only)
        available_recipes = self._get_season_pool(season, only_new=only_new)

        if not available_recipes:
            print(f"❌ No recipes found for season: {season}")
            return None

        print(f"📋 Found {len(available_recipes)} recipes for {season}")
        print(f"📅 Generating {num_weeks}-week meal plan ({num_weeks // 4} cycles)")

        # Create meal plan
        plan_name = f"{facility.upper()} {season.upper()} Meal Plan {datetime.now().strftime('%Y-%m-%d')}"
        meal_plan_id = self.db.create_meal_plan(plan_name, season, num_weeks, facility=facility)

        # Generate meal assignments
        meal_assignments = self._generate_assignments(
            available_recipes,
            num_weeks,
            season
        )

        if not meal_assignments:
            print("❌ Failed to generate valid meal plan")
            return None

        # Add meals to the plan
        for week_num, day_num, recipe_id, meal_type in meal_assignments:
            self.db.add_meal_to_plan(
                meal_plan_id,
                week_num,
                day_num,
                recipe_id,
                meal_type
            )

        print(f"✅ Meal plan created with ID: {meal_plan_id}")
        return meal_plan_id
    
    def _generate_assignments(self, available_recipes, num_weeks, season, week_offset=0,
                              day_indices=None):
        """
        Generate valid recipe assignments for the meal plan. Each weekday
        (Ma-Pe, DAYS_PER_WEEK of them) gets its own main dish, soup, and
        salad — three independent picks per day, not just one per week.

        Constraints:
        - Main dish: max repetition 2x per 4-week cycle, category variety
          capped within the week, season-appropriate
        - Soup: not repeated within SOUP_MIN_GAP_WEEKS calendar weeks
          (naturally forces a distinct soup each weekday too)
        - Salad: round-robin through the season's pool (also naturally
          distinct across the week when the pool has >= DAYS_PER_WEEK items)

        week_offset: absolute calendar week of week_num=1 minus 1. Used so the
        soup min-gap constraint stays continuous across a year plan's separate
        seasonal segments (e.g. the two talvi segments that wrap the year).

        day_indices: which weekday indices to generate for (defaults to
        range(DAYS_PER_WEEK), i.e. every day). Used by Hoiva's weekend
        extension to generate ONLY day_of_week 5/6 (la/su), since its
        ma-pe is always a live mirror of Kesti's plan, never its own data.
        """

        recipe_list = list(available_recipes)
        mains_pool, salad_pool = self._split_salads(recipe_list)
        mains_pool, soup_pool = self._split_soups(mains_pool)
        assignments = []

        if salad_pool:
            print(f"🥗 {len(salad_pool)} salaattia kierrossa kaudelle '{season}' "
                  f"(1 salaatti/arkipäivä)")
        else:
            print(f"⚠️  Ei salaatteja kaudelle '{season}'")

        if soup_pool:
            print(f"🍲 {len(soup_pool)} keittoa kierrossa kaudelle '{season}' "
                  f"(1 keitto/arkipäivä, min. {self.SOUP_MIN_GAP_WEEKS} vk väli)")
        else:
            print(f"⚠️  Ei keittoja kaudelle '{season}'")

        # Group main-dish recipes by category (soups/salads rotate
        # independently, so they don't count against the weekly
        # category-variety cap)
        by_category = defaultdict(list)
        for recipe in mains_pool:
            category = recipe[2]  # dish_category (query returns id, name, category, recipe_type)
            by_category[category].append(recipe)

        print(f"\n📊 Pääruokien jakauma kategorioittain:")
        for category, recipes in by_category.items():
            print(f"   {category}: {len(recipes)} recipes")

        # Track recipe usage per cycle for repetition constraint
        usage_per_cycle = defaultdict(lambda: defaultdict(int))

        weekdays = day_indices if day_indices is not None else range(self.DAYS_PER_WEEK)

        # Generate assignments week by week, weekday by weekday
        for week_num in range(1, num_weeks + 1):
            cycle_num = (week_num - 1) // self.weeks_per_cycle + 1
            absolute_week = week_offset + week_num

            week_main_ids = []

            for weekday in weekdays:
                # Main dish
                main = self._select_recipe(
                    mains_pool,
                    by_category,
                    usage_per_cycle[cycle_num],
                    week_main_ids,
                    week_num,
                    weekday
                )
                if not main:
                    print(f"⚠️  Warning: Could not find valid main dish for week {week_num}, day {weekday}")
                    main = random.choice(mains_pool) if mains_pool else random.choice(recipe_list)

                week_main_ids.append(main[0])
                assignments.append((week_num, weekday, main[0], MealType.LUNCH.value))
                usage_per_cycle[cycle_num][main[0]] += 1

                # Soup — independent pick, own min-gap constraint
                if soup_pool:
                    soup = self._select_soup(soup_pool, absolute_week)
                    assignments.append((week_num, weekday, soup[0], 'keitto'))

                # Salad — independent pick, own rotation. Kesti's Friday is
                # always the fixed buffet line instead of the rotated pick
                # (doesn't consume a turn from the rotation — Mon-Thu keep
                # cycling through the real pool at their own pace).
                if self.facility == 'kesti' and weekday == 4:
                    salad = self._get_friday_salad()
                    assignments.append((week_num, weekday, salad[0], 'salaatti'))
                elif salad_pool:
                    salad = self._select_salad(salad_pool, season)
                    assignments.append((week_num, weekday, salad[0], 'salaatti'))

        print(f"✅ Generated {len(assignments)} meal assignments")

        # Validate constraints
        self._validate_assignments(assignments, num_weeks)

        return assignments
    
    def _select_recipe(self, recipe_list, by_category, cycle_usage, 
                      week_recipes, week_num, day_num):
        """
        Select a recipe that satisfies all constraints
        
        Returns None if no valid recipe found
        """
        
        # Shuffle to add variety
        candidates = recipe_list.copy()
        random.shuffle(candidates)
        
        for recipe in candidates:
            recipe_id = recipe[0]
            
            # Constraint 1: Not already used in this cycle more than max allowed
            if cycle_usage[recipe_id] >= self.max_repetition_per_cycle:
                continue
            
            # Constraint 2: Not already in this week
            if recipe_id in week_recipes:
                continue
            
            # Constraint 3: Prefer category variety
            recipe_category = recipe[2]
            id_to_category = {r[0]: r[2] for r in recipe_list}
            week_categories = [id_to_category[rid] for rid in week_recipes if rid in id_to_category]

            # If this category is overrepresented in the week, deprioritize
            # (5 mains/week across 4 categories -> allow up to MAIN_CATEGORY_CAP_PER_WEEK of one category)
            category_count = week_categories.count(recipe_category)
            if category_count >= self.MAIN_CATEGORY_CAP_PER_WEEK:
                continue

            return recipe

        return None

    def _validate_assignments(self, assignments, num_weeks):
        """Validate that main-dish assignments meet the repetition constraint.
        Soup/salad have their own independent rules (min-gap / rotation) and
        are exempt from this check."""

        usage_per_cycle = defaultdict(lambda: defaultdict(int))

        for week_num, day_num, recipe_id, meal_type in assignments:
            if meal_type != MealType.LUNCH.value:
                continue
            cycle_num = (week_num - 1) // self.weeks_per_cycle + 1
            usage_per_cycle[cycle_num][recipe_id] += 1
        
        violations = 0
        for cycle_num, recipes in usage_per_cycle.items():
            for recipe_id, count in recipes.items():
                if count > self.max_repetition_per_cycle:
                    violations += 1
                    print(f"⚠️  Recipe {recipe_id} used {count}x in cycle {cycle_num}")
        
        if violations == 0:
            print("✅ All constraints satisfied")
        else:
            print(f"⚠️  {violations} constraint violations found")
    
    def _get_season_pool(self, season, only_new=False):
        """Recipes for a season, plus year-round recipes tagged 'kaikki'.
        Excludes recipes flagged manual_only (e.g. added sauces) — those are
        never auto-selected, only added by hand in the UI.
        only_new: restrict to recipes added after the overhaul cutover
        (recipes.id > app_settings.recipe_overhaul_cutover_id) — i.e. added
        via ANY channel (manual, PoweResta, OCR, scrape) since that point,
        regardless of import method. Recipes already in the database at
        cutover stay "legacy" forever, even if later edited."""
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()
        cols = [r[1] for r in c.execute('PRAGMA table_info(recipes)').fetchall()]
        manual_only_filter = 'AND (manual_only IS NULL OR manual_only = 0)' if 'manual_only' in cols else ''
        new_filter = ''
        params = (season,)
        if only_new:
            row = c.execute(
                "SELECT value FROM app_settings WHERE key = 'recipe_overhaul_cutover_id'").fetchone()
            cutover_id = int(row[0]) if row else 0
            new_filter = 'AND id > ?'
            params = (season, cutover_id)
        c.execute(f"""SELECT id, name_fi, dish_category, recipe_type FROM recipes
                     WHERE meal_type = 'lounas' AND (season = ? OR season = 'kaikki')
                     {manual_only_filter} {new_filter}""",
                  params)
        recipes = c.fetchall()
        conn.close()
        return recipes

    # Calendar mapping: which 13-week segment belongs to which season.
    # Week 1 = first week of January.
    YEAR_SEGMENTS = [
        ('talvi', 1, 9),     # Jan-Feb
        ('kevät', 10, 22),   # Mar-May
        ('kesä', 23, 35),    # Jun-Aug
        ('syksy', 36, 48),   # Sep-Nov
        ('talvi', 49, 52),   # Dec
    ]

    def _season_calendar_weeks(self):
        """Map each season to its full chronological list of calendar week
        numbers, merging talvi's two YEAR_SEGMENTS chunks (1-9 and 49-52)
        into one continuous 13-week run (talvi wraps the calendar year but
        is conceptually a single season)."""
        season_weeks = {}
        for season, start_week, end_week in self.YEAR_SEGMENTS:
            season_weeks.setdefault(season, []).extend(range(start_week, end_week + 1))
        return season_weeks

    def _generate_cycling_year_assignments(self, only_new=False, day_indices=None):
        """Yield (actual_week, day_num, recipe_id, meal_type) for a full
        52-week calendar year, one season at a time: generates CYCLE_WEEKS
        (6) unique weeks from that season's recipe pool, then repeats that
        block to fill however many calendar weeks the season actually spans
        (13 for each of the 4 seasons) — e.g. weeks 1-6 get fresh content,
        week 7 repeats week 1, week 8 repeats week 2, and so on. This is a
        fixed repeating menu cycle, same as real institutional catering
        typically runs, rather than 13 fully unique weeks per season.
        """
        for season, week_numbers in self._season_calendar_weeks().items():
            pool = self._get_season_pool(season, only_new=only_new)
            print(f"\n🍂 Kausi {season}: {len(week_numbers)} kalenteriviikkoa, "
                  f"{self.CYCLE_WEEKS} viikon kiertävä ruokalista ({len(pool)} reseptiä)")

            assignments = self._generate_assignments(pool, self.CYCLE_WEEKS, season,
                                                      day_indices=day_indices)
            by_cycle_week = defaultdict(list)
            for week_num, day_num, recipe_id, meal_type in assignments:
                by_cycle_week[week_num].append((day_num, recipe_id, meal_type))

            for i, actual_week in enumerate(week_numbers):
                cycle_week = (i % self.CYCLE_WEEKS) + 1
                for day_num, recipe_id, meal_type in by_cycle_week.get(cycle_week, []):
                    yield actual_week, day_num, recipe_id, meal_type

    def generate_year_plan(self, start_date=None, only_new=False, facility='kesti'):
        """
        Generate a full 52-week plan split into 4 seasonal themes
        following the calendar: talvi (1-9, 49-52), kevät (10-22),
        kesä (23-35), syksy (36-48).

        Each season generates only CYCLE_WEEKS (6) unique weeks from that
        season's recipe pool (plus year-round 'kaikki' recipes), then
        repeats that block for the rest of the season's calendar weeks —
        a fixed repeating menu cycle. Every weekday gets its own main dish
        + soup + salad.

        only_new: restrict to hand-added/edited recipes only, excluding
        bulk PoweResta/OCR/scrape imports.
        facility: 'kesti', 'hoiva' or 'kymenkartano' — independent plan per
        facility from the same shared recipe pool (DAYS_PER_WEEK differs).
        """
        from datetime import datetime

        # Verify all seasons have enough recipes
        shortages = []
        for season in ('talvi', 'kevät', 'kesä', 'syksy'):
            n = len(self._get_season_pool(season, only_new=only_new))
            if n < 25:
                shortages.append((season, n))
        if shortages:
            msg = ', '.join(f"{s}: {n} reseptiä" for s, n in shortages)
            print(f"❌ Liian vähän reseptejä kausille: {msg} (tarvitaan vähintään 25/kausi)")
            return None

        plan_name = f"{facility.upper()} VUOSISUUNNITELMA {datetime.now().strftime('%Y-%m-%d')}"
        meal_plan_id = self.db.create_meal_plan(plan_name, 'vuosi', 52, facility=facility)

        total = 0
        for actual_week, day_num, recipe_id, meal_type in self._generate_cycling_year_assignments(only_new):
            self.db.add_meal_to_plan(meal_plan_id, actual_week, day_num, recipe_id, meal_type)
            total += 1

        print(f"\n✅ Vuosisuunnitelma valmis ({self.CYCLE_WEEKS} viikon kiertävä ruokalista per kausi): "
              f"{total} ateriaa, ID {meal_plan_id}")
        return meal_plan_id

    WEEKEND_DAYS = (5, 6)  # La, Su — the only days Hoiva ever stores itself

    def generate_weekend_extension(self, kesti_plan_id, only_new=False):
        """
        Generate Hoiva's La/Su extension for an existing Kesti plan.

        Hoiva's ma-pe is a live mirror of the linked Kesti plan (read
        directly from it at query time, never copied/stored) — this method
        only ever generates and stores La/Su, and links the new plan back
        to kesti_plan_id via meal_plans.mirrors_plan_id so the mirror can
        be resolved later.

        Returns the new Hoiva meal_plan_id, or None if kesti_plan_id
        doesn't exist or isn't a 'kesti' plan.
        """
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        conn.row_factory = sqlite3.Row
        source = conn.execute(
            'SELECT id, name, season, num_weeks, facility FROM meal_plans WHERE id = ?',
            (kesti_plan_id,)).fetchone()
        conn.close()
        if not source or (source['facility'] or 'kesti') != 'kesti':
            print(f"❌ Plan {kesti_plan_id} is not a valid Kesti plan to mirror")
            return None

        season, num_weeks = source['season'], source['num_weeks']
        plan_name = f"HOIVA {season.upper()} (la-su, peilaa: {source['name']})"
        hoiva_plan_id = self.db.create_meal_plan(plan_name, season, num_weeks, facility='hoiva')

        conn = sqlite3.connect(self.db.db_path)
        conn.execute('UPDATE meal_plans SET mirrors_plan_id = ? WHERE id = ?',
                    (kesti_plan_id, hoiva_plan_id))
        conn.commit()
        conn.close()

        total = 0
        if season == 'vuosi':
            for actual_week, day_num, recipe_id, meal_type in \
                    self._generate_cycling_year_assignments(only_new, day_indices=self.WEEKEND_DAYS):
                self.db.add_meal_to_plan(hoiva_plan_id, actual_week, day_num, recipe_id, meal_type)
                total += 1
        else:
            pool = self._get_season_pool(season, only_new=only_new)
            assignments = self._generate_assignments(
                pool, num_weeks, season, day_indices=self.WEEKEND_DAYS)
            for week_num, day_num, recipe_id, meal_type in assignments:
                self.db.add_meal_to_plan(hoiva_plan_id, week_num, day_num, recipe_id, meal_type)
            total = len(assignments)

        print(f"✅ Hoiva la-su valmis: {total} ateriaa, ID {hoiva_plan_id} (peilaa suunnitelmaa {kesti_plan_id})")
        return hoiva_plan_id

    def get_meal_plan_stats(self, meal_plan_id):
        """Get statistics about a meal plan"""
        
        meal_plan = self.db.get_meal_plan(meal_plan_id)
        if not meal_plan:
            return None
        
        # Count recipe usage
        recipe_usage = defaultdict(int)
        category_usage = defaultdict(int)  # main dishes only — soup/salad would dilute the protein-balance reading

        for week_num, day_num, recipe_id, meal_type in meal_plan['days']:
            recipe_usage[recipe_id] += 1

            if meal_type == MealType.LUNCH.value:
                recipe = self.db.get_recipe_details(recipe_id)
                if recipe:
                    category_usage[recipe['dish_category']] += 1
        
        return {
            'plan_name': meal_plan['name'],
            'season': meal_plan['season'],
            'weeks': meal_plan['num_weeks'],
            'total_meals': len(meal_plan['days']),
            'unique_recipes': len(recipe_usage),
            'recipe_usage': dict(recipe_usage),
            'category_distribution': dict(category_usage)
        }


if __name__ == '__main__':
    # Example usage
    generator = MealPlanGenerator()
    
    # This will work once we have recipes in the database
    # meal_plan_id = generator.generate_meal_plan('talvi', num_weeks=52)
    # 
    # if meal_plan_id:
    #     stats = generator.get_meal_plan_stats(meal_plan_id)
    #     print("\n📊 Meal Plan Statistics:")
    #     print(f"   Total meals: {stats['total_meals']}")
    #     print(f"   Unique recipes: {stats['unique_recipes']}")
    #     print(f"   Categories: {stats['category_distribution']}")
    
    print("✅ Meal Plan Generator initialized")

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
    
    SALAD_KEYWORD = 'salaa'  # matches salaatti/salaatit/salaattia etc.
    SOUP_KEYWORD = 'keitto'  # matches keitto/keittoa/sosekeitto etc.
    SOUP_MIN_GAP_WEEKS = 2   # don't repeat the same soup within 2 calendar weeks
    DAYS_PER_WEEK = 5        # Ma-Pe — each weekday gets its own soup+main+salad
    MAIN_CATEGORY_CAP_PER_WEEK = 2  # no protein category more than 2x among the week's 5 mains

    def __init__(self, db_path='meal_plans.db'):
        self.db = MealPlanDB(db_path)
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

    def _split_salads(self, recipe_pool):
        """Split a recipe pool into (mains, salads). Salads are identified by
        name and sorted by id for a stable rotation order."""
        salads = [r for r in recipe_pool if self.SALAD_KEYWORD in r[1].lower()]
        mains = [r for r in recipe_pool if self.SALAD_KEYWORD not in r[1].lower()]
        salads.sort(key=lambda r: r[0])
        return mains, salads

    def _select_salad(self, salads, season):
        """Round-robin through the season's salad pool — one salad per week,
        never the same one twice in a row, evenly distributed over time."""
        idx = self._salad_rotation[season] % len(salads)
        self._salad_rotation[season] += 1
        return salads[idx]

    def _split_soups(self, recipe_pool):
        """Split a recipe pool into (mains, soups). Soups are identified by name."""
        soups = [r for r in recipe_pool if self.SOUP_KEYWORD in r[1].lower()]
        mains = [r for r in recipe_pool if self.SOUP_KEYWORD not in r[1].lower()]
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
    
    def generate_meal_plan(self, season, num_weeks=52, only_new=False):
        """
        Generate a complete meal plan: each of the DAYS_PER_WEEK weekdays gets
        its own main dish + soup + salad (independent rotations).

        Args:
            season: Season name (e.g., 'talvi', 'kevät', 'kesä', 'syksy')
            num_weeks: Total number of weeks (default 52)
            only_new: restrict to hand-added/edited recipes only

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
        plan_name = f"{season.upper()} Meal Plan {datetime.now().strftime('%Y-%m-%d')}"
        meal_plan_id = self.db.create_meal_plan(plan_name, season, num_weeks)

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
    
    def _generate_assignments(self, available_recipes, num_weeks, season, week_offset=0):
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
            category = recipe[2]  # dish_category (query returns id, name, category)
            by_category[category].append(recipe)

        print(f"\n📊 Pääruokien jakauma kategorioittain:")
        for category, recipes in by_category.items():
            print(f"   {category}: {len(recipes)} recipes")

        # Track recipe usage per cycle for repetition constraint
        usage_per_cycle = defaultdict(lambda: defaultdict(int))

        # Generate assignments week by week, weekday by weekday
        for week_num in range(1, num_weeks + 1):
            cycle_num = (week_num - 1) // self.weeks_per_cycle + 1
            absolute_week = week_offset + week_num

            week_main_ids = []

            for weekday in range(self.DAYS_PER_WEEK):
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

                # Salad — independent pick, own rotation
                if salad_pool:
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
        c.execute(f"""SELECT id, name_fi, dish_category FROM recipes
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

    def generate_year_plan(self, start_date=None, only_new=False):
        """
        Generate a full 52-week plan split into 4 seasonal themes
        following the calendar: talvi (1-9, 49-52), kevät (10-22),
        kesä (23-35), syksy (36-48).

        Each segment draws only from that season's recipe pool
        (plus year-round 'kaikki' recipes). Repetition constraint
        applies within each 4-week cycle as usual. Every weekday gets
        its own main dish + soup + salad.

        only_new: restrict to hand-added/edited recipes only, excluding
        bulk PoweResta/OCR/scrape imports.
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

        plan_name = f"VUOSISUUNNITELMA {datetime.now().strftime('%Y-%m-%d')}"
        meal_plan_id = self.db.create_meal_plan(plan_name, 'vuosi', 52)

        total = 0
        for season, start_week, end_week in self.YEAR_SEGMENTS:
            pool = self._get_season_pool(season, only_new=only_new)
            seg_weeks = end_week - start_week + 1
            print(f"\n🍂 Kausi {season}: viikot {start_week}-{end_week} ({len(pool)} reseptiä)")

            assignments = self._generate_assignments(
                pool, seg_weeks, season,
                week_offset=start_week - 1
            )
            for week_num, day_num, recipe_id, meal_type in assignments:
                self.db.add_meal_to_plan(
                    meal_plan_id,
                    start_week + week_num - 1,   # offset into the year
                    day_num, recipe_id,
                    meal_type
                )
            total += len(assignments)

        print(f"\n✅ Vuosisuunnitelma valmis: {total} ateriaa, ID {meal_plan_id}")
        return meal_plan_id

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

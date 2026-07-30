"""
Finnish Seasonal Meal Plan Generator - Database Models
"""

import sqlite3
from datetime import datetime
from enum import Enum
import json

class Season(Enum):
    WINTER = "talvi"      # December-February
    SPRING = "kevät"      # March-May
    SUMMER = "kesä"       # June-August
    AUTUMN = "syksy"      # September-November

class MealType(Enum):
    BREAKFAST = "aamiainen"
    LUNCH = "lounas"
    DINNER = "päivällinen"
    SNACK = "välipala"

class DishCategory(Enum):
    FISH = "kala"           # Blue
    CHICKEN = "kana"        # Yellow
    BEEF = "naudanliha"     # Red
    VEGETABLE = "kasvis"    # Green
    CARB = "hiilihydraatti" # Legend items (rice, pasta, potatoes, etc.)

class MealPlanDB:
    """SQLite database for meal plan management"""
    
    def __init__(self, db_path='meal_plans.db'):
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """Create database tables"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        # Recipes table
        c.execute('''CREATE TABLE IF NOT EXISTS recipes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name_fi TEXT NOT NULL UNIQUE,
            season TEXT NOT NULL,
            meal_type TEXT NOT NULL,
            dish_category TEXT NOT NULL,
            ingredients TEXT,
            nutrition TEXT,
            prep_time_min INTEGER,
            difficulty INTEGER,
            source_url TEXT,
            scraped_date TEXT,
            notes TEXT,
            servings INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )''')
        
        # Ingredients table
        c.execute('''CREATE TABLE IF NOT EXISTS ingredients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name_fi TEXT NOT NULL UNIQUE,
            category TEXT,
            unit TEXT,
            seasonal_availability TEXT,
            notes TEXT
        )''')
        
        # Recipe-Ingredient junction table
        c.execute('''CREATE TABLE IF NOT EXISTS recipe_ingredients (
            recipe_id INTEGER,
            ingredient_id INTEGER,
            quantity REAL,
            unit TEXT,
            FOREIGN KEY(recipe_id) REFERENCES recipes(id),
            FOREIGN KEY(ingredient_id) REFERENCES ingredients(id),
            PRIMARY KEY(recipe_id, ingredient_id)
        )''')
        
        # Meal plans table
        c.execute('''CREATE TABLE IF NOT EXISTS meal_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            season TEXT NOT NULL,
            start_date DATE,
            num_weeks INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )''')
        
        # Meal plan days table
        c.execute('''CREATE TABLE IF NOT EXISTS meal_plan_days (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meal_plan_id INTEGER,
            week_number INTEGER,
            day_of_week INTEGER,
            recipe_id INTEGER,
            meal_type TEXT,
            FOREIGN KEY(meal_plan_id) REFERENCES meal_plans(id),
            FOREIGN KEY(recipe_id) REFERENCES recipes(id)
        )''')
        
        conn.commit()
        conn.close()
    
    def add_recipe(self, name_fi, season, meal_type, dish_category,
                   ingredients=None, nutrition=None, prep_time_min=None,
                   difficulty=None, source_url=None, notes=None, servings=None,
                   recipe_type=None, instructions=None):
        """Add a recipe to the database"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        cols = [r[1] for r in c.execute('PRAGMA table_info(recipes)').fetchall()]
        has_servings = 'servings' in cols
        has_recipe_type = 'recipe_type' in cols
        has_instructions = 'instructions' in cols

        try:
            if has_servings:
                c.execute('''INSERT INTO recipes
                            (name_fi, season, meal_type, dish_category,
                             ingredients, nutrition, prep_time_min, difficulty,
                             source_url, notes, scraped_date, servings)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                         (name_fi, season, meal_type, dish_category,
                          json.dumps(ingredients) if ingredients else None,
                          json.dumps(nutrition) if nutrition else None,
                          prep_time_min, difficulty, source_url, notes,
                          datetime.now().isoformat(), servings))
            else:
                c.execute('''INSERT INTO recipes
                            (name_fi, season, meal_type, dish_category,
                             ingredients, nutrition, prep_time_min, difficulty,
                             source_url, notes, scraped_date)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                         (name_fi, season, meal_type, dish_category,
                          json.dumps(ingredients) if ingredients else None,
                          json.dumps(nutrition) if nutrition else None,
                          prep_time_min, difficulty, source_url, notes,
                          datetime.now().isoformat()))
            recipe_id = c.lastrowid
            if has_recipe_type and recipe_type:
                c.execute('UPDATE recipes SET recipe_type = ? WHERE id = ?',
                          (recipe_type, recipe_id))
            if has_instructions and instructions:
                c.execute('UPDATE recipes SET instructions = ? WHERE id = ?',
                          (instructions, recipe_id))
            conn.commit()
            return recipe_id
        except sqlite3.IntegrityError:
            print(f"Recipe '{name_fi}' already exists")
            return None
        finally:
            conn.close()
    
    def add_ingredient(self, name_fi, category=None, unit=None, 
                      seasonal_availability=None, notes=None):
        """Add an ingredient to the database"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        try:
            c.execute('''INSERT INTO ingredients 
                        (name_fi, category, unit, seasonal_availability, notes)
                        VALUES (?, ?, ?, ?, ?)''',
                     (name_fi, category, unit, seasonal_availability, notes))
            conn.commit()
            return c.lastrowid
        except sqlite3.IntegrityError:
            print(f"Ingredient '{name_fi}' already exists")
            return None
        finally:
            conn.close()
    
    def link_ingredient_to_recipe(self, recipe_id, ingredient_id, quantity, unit):
        """Link an ingredient to a recipe with quantity"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('''INSERT OR REPLACE INTO recipe_ingredients 
                    (recipe_id, ingredient_id, quantity, unit)
                    VALUES (?, ?, ?, ?)''',
                 (recipe_id, ingredient_id, quantity, unit))
        conn.commit()
        conn.close()
    
    def get_recipes_by_season_and_type(self, season, meal_type):
        """Get all recipes for a specific season and meal type"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('''SELECT id, name_fi, dish_category FROM recipes 
                    WHERE season = ? AND meal_type = ?''',
                 (season, meal_type))
        recipes = c.fetchall()
        conn.close()
        return recipes
    
    def get_recipe_details(self, recipe_id):
        """Get full recipe details including ingredients"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        # Get recipe info
        c.execute('SELECT * FROM recipes WHERE id = ?', (recipe_id,))
        recipe = c.fetchone()
        
        if recipe:
            # Get ingredients
            c.execute('''SELECT i.name_fi, ri.quantity, ri.unit 
                        FROM recipe_ingredients ri
                        JOIN ingredients i ON ri.ingredient_id = i.id
                        WHERE ri.recipe_id = ?''', (recipe_id,))
            ingredients = c.fetchall()
            conn.close()
            return {
                'id': recipe[0],
                'name': recipe[1],
                'season': recipe[2],
                'meal_type': recipe[3],
                'dish_category': recipe[4],
                'ingredients': ingredients,
                'nutrition': json.loads(recipe[6]) if recipe[6] else None,
                'prep_time': recipe[7],
                'difficulty': recipe[8]
            }
        
        conn.close()
        return None
    
    def get_all_recipes(self, season=None, meal_type=None):
        """Get recipes with optional filtering"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        query = 'SELECT id, name_fi, season, meal_type, dish_category FROM recipes WHERE 1=1'
        params = []
        
        if season:
            query += ' AND season = ?'
            params.append(season)
        if meal_type:
            query += ' AND meal_type = ?'
            params.append(meal_type)
        
        c.execute(query, params)
        recipes = c.fetchall()
        conn.close()
        return recipes
    
    def create_meal_plan(self, name, season, num_weeks=52, facility='kesti'):
        """Create a new meal plan"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        cols = [r[1] for r in c.execute('PRAGMA table_info(meal_plans)').fetchall()]
        if 'facility' in cols:
            c.execute('''INSERT INTO meal_plans (name, season, start_date, num_weeks, facility)
                        VALUES (?, ?, ?, ?, ?)''',
                     (name, season, datetime.now().date(), num_weeks, facility))
        else:
            c.execute('''INSERT INTO meal_plans (name, season, start_date, num_weeks)
                        VALUES (?, ?, ?, ?)''',
                     (name, season, datetime.now().date(), num_weeks))
        conn.commit()
        meal_plan_id = c.lastrowid
        conn.close()
        return meal_plan_id
    
    def add_meal_to_plan(self, meal_plan_id, week_number, day_of_week, 
                        recipe_id, meal_type):
        """Add a recipe to a specific day in the meal plan"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('''INSERT INTO meal_plan_days 
                    (meal_plan_id, week_number, day_of_week, recipe_id, meal_type)
                    VALUES (?, ?, ?, ?, ?)''',
                 (meal_plan_id, week_number, day_of_week, recipe_id, meal_type))
        conn.commit()
        conn.close()
    
    def get_meal_plan(self, meal_plan_id):
        """Get full meal plan with all recipes"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('SELECT * FROM meal_plans WHERE id = ?', (meal_plan_id,))
        plan = c.fetchone()
        
        if plan:
            c.execute('''SELECT week_number, day_of_week, recipe_id, meal_type 
                        FROM meal_plan_days 
                        WHERE meal_plan_id = ?
                        ORDER BY week_number, day_of_week''',
                     (meal_plan_id,))
            days = c.fetchall()
            conn.close()
            return {
                'id': plan[0],
                'name': plan[1],
                'season': plan[2],
                'num_weeks': plan[4],
                'days': days
            }
        
        conn.close()
        return None

if __name__ == '__main__':
    # Initialize database
    db = MealPlanDB()
    print("✅ Database initialized at meal_plans.db")

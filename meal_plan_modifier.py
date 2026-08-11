"""
Meal Plan Modification System
Handles last-minute changes, supply shortages, and meal substitutions
"""

from meal_plan_db import MealPlanDB
from datetime import datetime
from collections import defaultdict
import json

class MealModifier:
    """Manages modifications to meal plans with audit trail"""
    
    def __init__(self, db_path='meal_plans.db'):
        self.db = MealPlanDB(db_path)
        self.init_modification_tables()
    
    def init_modification_tables(self):
        """Create tables for tracking modifications"""
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()
        
        # Modifications log (audit trail)
        c.execute('''CREATE TABLE IF NOT EXISTS meal_modifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meal_plan_id INTEGER,
            week_number INTEGER,
            day_of_week INTEGER,
            original_recipe_id INTEGER,
            new_recipe_id INTEGER,
            reason TEXT,
            modified_by TEXT,
            modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(meal_plan_id) REFERENCES meal_plans(id),
            FOREIGN KEY(original_recipe_id) REFERENCES recipes(id),
            FOREIGN KEY(new_recipe_id) REFERENCES recipes(id)
        )''')
        
        # Meal exclusions (recipes to avoid in a week)
        c.execute('''CREATE TABLE IF NOT EXISTS recipe_exclusions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meal_plan_id INTEGER,
            recipe_id INTEGER,
            reason TEXT,
            excluded_from_date DATE,
            excluded_to_date DATE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(meal_plan_id) REFERENCES meal_plans(id),
            FOREIGN KEY(recipe_id) REFERENCES recipes(id)
        )''')
        
        # Supply shortage alerts
        c.execute('''CREATE TABLE IF NOT EXISTS supply_shortages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ingredient_id INTEGER,
            reason TEXT,
            shortage_from_date DATE,
            shortage_to_date DATE,
            severity TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(ingredient_id) REFERENCES ingredients(id)
        )''')
        
        conn.commit()
        conn.close()
    
    def change_meal(self, meal_plan_id, week_number, day_of_week,
                   new_recipe_id, meal_type='lounas', reason="Manual change", modified_by="admin"):
        """
        Change a specific meal in the plan

        Args:
            meal_plan_id: ID of the meal plan
            week_number: Week number (1-52)
            day_of_week: Weekday (0-4, Ma-Pe)
            new_recipe_id: ID of the new recipe
            meal_type: Which dish role on that day — 'lounas' (main), 'keitto'
                (soup) or 'salaatti' (salad). A day has one row per role, so
                this is required to identify the right row.
            reason: Reason for change (supply shortage, preference, etc.)
            modified_by: Who made the change (for audit trail)

        Returns:
            True if successful, False otherwise
        """

        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()

        try:
            # Get current meal
            c.execute('''SELECT recipe_id FROM meal_plan_days
                        WHERE meal_plan_id = ? AND week_number = ? AND day_of_week = ? AND meal_type = ?''',
                     (meal_plan_id, week_number, day_of_week, meal_type))
            current = c.fetchone()

            if not current:
                print(f"❌ No meal found for plan {meal_plan_id}, week {week_number}, day {day_of_week}, type {meal_type}")
                return False

            original_recipe_id = current[0]

            # Update the meal
            c.execute('''UPDATE meal_plan_days
                        SET recipe_id = ?
                        WHERE meal_plan_id = ? AND week_number = ? AND day_of_week = ? AND meal_type = ?''',
                     (new_recipe_id, meal_plan_id, week_number, day_of_week, meal_type))
            
            # Log the modification
            c.execute('''INSERT INTO meal_modifications 
                        (meal_plan_id, week_number, day_of_week, original_recipe_id, 
                         new_recipe_id, reason, modified_by)
                        VALUES (?, ?, ?, ?, ?, ?, ?)''',
                     (meal_plan_id, week_number, day_of_week, original_recipe_id, 
                      new_recipe_id, reason, modified_by))
            
            conn.commit()
            
            # Get recipe names for feedback
            old_recipe = self.db.get_recipe_details(original_recipe_id)
            new_recipe = self.db.get_recipe_details(new_recipe_id)
            
            print(f"✅ Changed: '{old_recipe['name']}' → '{new_recipe['name']}'")
            print(f"   Week {week_number}, Meal {day_of_week}")
            print(f"   Reason: {reason}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error changing meal: {e}")
            return False
        finally:
            conn.close()
    
    def exclude_recipe(self, meal_plan_id, recipe_id, reason, 
                      from_date=None, to_date=None):
        """
        Exclude a recipe from being used (supply shortage, etc.)
        
        Args:
            meal_plan_id: ID of the meal plan
            recipe_id: ID of recipe to exclude
            reason: Why it's excluded (e.g., "salmon not available")
            from_date: Start date of exclusion
            to_date: End date of exclusion
        
        Returns:
            Number of meals changed
        """
        
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()
        
        try:
            # Log the exclusion
            c.execute('''INSERT INTO recipe_exclusions 
                        (meal_plan_id, recipe_id, reason, excluded_from_date, excluded_to_date)
                        VALUES (?, ?, ?, ?, ?)''',
                     (meal_plan_id, recipe_id, reason, from_date, to_date))
            
            # Find all instances of this recipe in the plan
            c.execute('''SELECT week_number, day_of_week FROM meal_plan_days 
                        WHERE meal_plan_id = ? AND recipe_id = ?''',
                     (meal_plan_id, recipe_id))
            occurrences = c.fetchall()
            
            if occurrences:
                print(f"⚠️  Found {len(occurrences)} instances of this recipe in meal plan")
                print(f"   Reason: {reason}")
                print(f"\n   These meals need to be replaced:")
                for week, day in occurrences:
                    print(f"   - Week {week}, Meal {day}")
            else:
                print(f"ℹ️  Recipe not used in this meal plan")
            
            conn.commit()
            return len(occurrences)
            
        except Exception as e:
            print(f"❌ Error excluding recipe: {e}")
            return 0
        finally:
            conn.close()
    
    def report_supply_shortage(self, ingredient_id, reason, 
                              from_date=None, to_date=None, severity="medium"):
        """
        Report a supply shortage for an ingredient
        Helps identify affected recipes
        
        Args:
            ingredient_id: ID of ingredient
            reason: Description of shortage
            from_date: When shortage starts
            to_date: When shortage ends
            severity: "low", "medium", or "high"
        
        Returns:
            List of affected recipes
        """
        
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()
        
        try:
            # Log the shortage
            c.execute('''INSERT INTO supply_shortages 
                        (ingredient_id, reason, shortage_from_date, shortage_to_date, severity)
                        VALUES (?, ?, ?, ?, ?)''',
                     (ingredient_id, reason, from_date, to_date, severity))
            
            # Find recipes using this ingredient
            c.execute('''SELECT DISTINCT r.id, r.name_fi 
                        FROM recipes r
                        JOIN recipe_ingredients ri ON r.id = ri.recipe_id
                        WHERE ri.ingredient_id = ?''',
                     (ingredient_id,))
            
            affected_recipes = c.fetchall()
            
            # Get ingredient name
            c.execute('SELECT name_fi FROM ingredients WHERE id = ?', (ingredient_id,))
            ingredient = c.fetchone()
            ingredient_name = ingredient[0] if ingredient else "Unknown"
            
            print(f"⚠️  Supply Shortage Alert: {ingredient_name}")
            print(f"   Reason: {reason}")
            print(f"   Severity: {severity.upper()}")
            print(f"   Affects {len(affected_recipes)} recipes:")
            for recipe_id, recipe_name in affected_recipes:
                print(f"   - {recipe_name}")
            
            conn.commit()
            return affected_recipes
            
        except Exception as e:
            print(f"❌ Error reporting shortage: {e}")
            return []
        finally:
            conn.close()
    
    def suggest_recipe_replacement(self, meal_plan_id, week_number,
                                  day_of_week, meal_type='lounas', exclude_recipes=None):
        """
        Suggest alternative recipes for a specific meal

        Args:
            meal_plan_id: ID of the meal plan
            week_number: Week number
            day_of_week: Weekday (0-4)
            meal_type: Dish role — 'lounas'/'lounas2' (main), 'keitto' (soup)
                or 'salaatti' (salad). Suggestions stay within the same role
                (e.g. a soup slot only suggests other soups) but, on a
                manual short-notice swap, are NOT restricted to the current
                dish's protein category (kala/kana/naudanliha/kasvis) —
                the manager may need to swap a fish main for a chicken one
                on short notice, same as soups/salads already allow any
                category. The generator's own automatic weekly category
                balance is unaffected; this only concerns manual swaps.
            exclude_recipes: List of recipe IDs to avoid

        Returns:
            List of suitable replacement recipes
        """

        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()

        try:
            # Get current meal
            c.execute('''SELECT recipe_id FROM meal_plan_days
                        WHERE meal_plan_id = ? AND week_number = ? AND day_of_week = ? AND meal_type = ?''',
                     (meal_plan_id, week_number, day_of_week, meal_type))
            current = c.fetchone()

            if not current:
                return []

            current_recipe_id = current[0]

            if meal_type in ('keitto', 'salaatti', 'kastike', 'kasvislisäke', 'energialisäke'):
                role_recipe_type = meal_type
            else:
                # 'lounas', 'lounas2' or 'lounas3' — all main-course slots.
                role_recipe_type = 'pääruoka'
            c.execute('''SELECT id, name_fi FROM recipes
                        WHERE recipe_type = ? AND id != ?
                        ORDER BY name_fi''', (role_recipe_type, current_recipe_id))

            suggestions = c.fetchall()

            # Filter out excluded recipes
            if exclude_recipes:
                suggestions = [r for r in suggestions if r[0] not in exclude_recipes]

            return suggestions

        except Exception as e:
            print(f"❌ Error getting suggestions: {e}")
            return []
        finally:
            conn.close()
    
    def get_modification_history(self, meal_plan_id):
        """Get all modifications made to a meal plan"""
        
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()
        
        c.execute('''SELECT week_number, day_of_week, original_recipe_id, new_recipe_id, 
                           reason, modified_by, modified_at
                    FROM meal_modifications 
                    WHERE meal_plan_id = ?
                    ORDER BY modified_at DESC''',
                 (meal_plan_id,))
        
        modifications = []
        for week, day, old_id, new_id, reason, who, when in c.fetchall():
            old_recipe = self.db.get_recipe_details(old_id)
            new_recipe = self.db.get_recipe_details(new_id)
            
            modifications.append({
                'week': week,
                'day': day,
                'from': old_recipe['name'] if old_recipe else 'Unknown',
                'to': new_recipe['name'] if new_recipe else 'Unknown',
                'reason': reason,
                'by': who,
                'at': when
            })
        
        conn.close()
        return modifications
    
    def get_exclusions_report(self, meal_plan_id):
        """Get report of all excluded recipes and their status"""
        
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()
        
        # Get exclusions
        c.execute('''SELECT re.recipe_id, r.name_fi, re.reason, 
                           re.excluded_from_date, re.excluded_to_date
                    FROM recipe_exclusions re
                    JOIN recipes r ON re.recipe_id = r.id
                    WHERE re.meal_plan_id = ?''',
                 (meal_plan_id,))
        
        exclusions = []
        for recipe_id, name, reason, from_date, to_date in c.fetchall():
            # Find remaining instances
            c.execute('''SELECT COUNT(*) FROM meal_plan_days 
                        WHERE meal_plan_id = ? AND recipe_id = ?''',
                     (meal_plan_id, recipe_id))
            remaining = c.fetchone()[0]
            
            exclusions.append({
                'recipe': name,
                'reason': reason,
                'from': from_date,
                'to': to_date,
                'remaining_instances': remaining
            })
        
        conn.close()
        return exclusions
    
    def get_current_status(self, meal_plan_id):
        """Get current status of meal plan with all modifications and exclusions"""
        
        modifications = self.get_modification_history(meal_plan_id)
        exclusions = self.get_exclusions_report(meal_plan_id)
        
        return {
            'modifications_count': len(modifications),
            'exclusions_count': len(exclusions),
            'recent_changes': modifications[:5],
            'excluded_recipes': exclusions
        }


if __name__ == '__main__':
    modifier = MealModifier()
    print("✅ Meal Modifier initialized")
    
    # Example usage (would need recipes in database first):
    # 1. Change a meal:
    #    modifier.change_meal(meal_plan_id=1, week_number=2, day_of_week=3, 
    #                        new_recipe_id=5, reason="Salmon not available", 
    #                        modified_by="cook_name")
    #
    # 2. Exclude a recipe due to shortage:
    #    affected = modifier.exclude_recipe(meal_plan_id=1, recipe_id=10,
    #                                      reason="Supplier ran out of beef")
    #
    # 3. Report ingredient shortage:
    #    recipes = modifier.report_supply_shortage(ingredient_id=5,
    #                                             reason="Tomatoes out of season",
    #                                             severity="high")
    #
    # 4. Get suggestions:
    #    suggestions = modifier.suggest_recipe_replacement(meal_plan_id=1,
    #                                                    week_number=2,
    #                                                    day_of_week=3)

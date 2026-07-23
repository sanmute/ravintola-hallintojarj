import sqlite3

conn = sqlite3.connect('meal_plans.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Get the latest meal plan
cursor.execute("SELECT id FROM meal_plans ORDER BY id DESC LIMIT 1")
latest_plan = cursor.fetchone()

if latest_plan:
    meal_plan_id = latest_plan['id']
    print(f"Meal Plan ID: {meal_plan_id}\n")
    
    # Check the structure
    cursor.execute("""
        SELECT * FROM meal_plan_days 
        WHERE meal_plan_id = ? AND week_number = 1
        LIMIT 5
    """, (meal_plan_id,))
    
    days = cursor.fetchall()
    print(f"Found {len(days)} days in week 1\n")
    
    for day in days:
        print(f"Raw day data: {dict(day)}")
        recipe = conn.execute("SELECT name_fi FROM recipes WHERE id = ?", (day['recipe_id'],)).fetchone()
        print(f"  Recipe: {recipe['name_fi'] if recipe else 'NOT FOUND'}\n")

conn.close()
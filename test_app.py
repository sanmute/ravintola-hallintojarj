"""
test_app.py — Diagnostic launcher to see real errors.

Run this instead of the exe to see what's actually breaking.
    python test_app.py

Then open http://localhost:5001 in browser and try to reproduce the error.
The terminal will show the actual exception, not just "<" JSON error.
"""

import os
import sys

# Add current dir to path so imports work
sys.path.insert(0, os.path.dirname(__file__))

from app import app, DB_PATH
from meal_plan_db import MealPlanDB
from meal_plan_modifier import MealModifier

if __name__ == '__main__':
    print('=' * 60)
    print('  Ruokalistasuunnittelija — DIAGNOSTINEN KÄYNNISTYS')
    print('  (näyttää todellisia virheitä)')
    print('=' * 60)
    print()
    
    # Init databases
    try:
        MealPlanDB(DB_PATH)
        MealModifier(DB_PATH)
        print('✓ Tietokannat alustettu')
    except Exception as e:
        print(f'✗ Tietokannan alustus epäonnistui: {e}')
        sys.exit(1)
    
    print('✓ Käynnistetään palvelin')
    print()
    print('Avaa selaimessa: http://localhost:5001')
    print('Sulkeminen: Ctrl+C')
    print()
    
    # Run with debug=True so we see actual exceptions
    try:
        app.run(debug=True, host='127.0.0.1', port=5001, use_reloader=False)
    except KeyboardInterrupt:
        print('\nSuljettu.')
        sys.exit(0)

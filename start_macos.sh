#!/bin/bash
# Ruokalistasuunnittelija — macOS Quick Start
# Run this script to set up and launch the meal planner

set -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo ""
echo "╔════════════════════════════════════════════════╗"
echo "║   Ruokalistasuunnittelija — macOS Launcher     ║"
echo "╚════════════════════════════════════════════════╝"
echo ""

# Check Python
echo "✓ Checking Python..."
if ! command -v python3 &> /dev/null; then
    echo ""
    echo "❌ Python 3 not found!"
    echo ""
    echo "Install Python via Homebrew:"
    echo "  brew install python@3.11"
    echo ""
    echo "Or download from: https://www.python.org/downloads/"
    echo ""
    exit 1
fi

PYTHON_VERSION=$(python3 -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')
echo "   Found Python $PYTHON_VERSION"

# Create virtual environment if needed
if [ ! -d "venv" ]; then
    echo ""
    echo "✓ Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "✓ Activating virtual environment..."
source venv/bin/activate

# Install/update dependencies
echo "✓ Installing dependencies..."
pip install -q --upgrade pip
pip install -q -r requirements.txt 2>&1 | grep -v "already satisfied" || true

# Seed data if needed
if [ ! -f "meal_plans.db" ]; then
    echo "✓ Loading initial recipes..."
    python3 seed_recipes.py > /dev/null
fi

echo ""
echo "╔════════════════════════════════════════════════╗"
echo "║         🚀 Starting Meal Planner               ║"
echo "╚════════════════════════════════════════════════╝"
echo ""
echo "📱 Open your browser:"
echo "   http://localhost:5001"
echo ""
echo "📋 Tabs available:"
echo "   • Reseptit (Recipes)"
echo "   • Tarkistusjono (Review Queue)"
echo "   • Ruokalistat (Meal Plans)"
echo "   • Selaa & muokkaa (Browse & Edit)"
echo ""
echo "⏹️  To stop: Press Ctrl+C"
echo ""
echo "────────────────────────────────────────────────"
echo ""

# Start the app
python3 app.py

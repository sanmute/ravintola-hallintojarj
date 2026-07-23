"""Quick OCR diagnostic."""
import sys
import os

print("=" * 60)
print("  OCR DIAGNOSTIIKKA")
print("=" * 60)
print()

# 1. Check pytesseract
print("[1] Tarkistetaan pytesseract...")
try:
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\santeri.mutanen\AppData\Local\Tesseract-OCR\tesseract.exe'
    print(f"  ✓ pytesseract importti OK")
    print(f"    tesseract_cmd: {pytesseract.pytesseract.tesseract_cmd}")
except Exception as e:
    print(f"  ✗ pytesseract virhe: {e}")
    sys.exit(1)

# 2. Check if tesseract.exe exists
print()
print("[2] Tarkistetaan tesseract.exe olemassaolo...")
cmd_path = pytesseract.pytesseract.tesseract_cmd
if not cmd_path:
    # If not set, try to set it
    pytesseract.tesseract_cmd = r'C:\Users\santeri.mutanen\AppData\Local\Tesseract-OCR\tesseract.exe'
    cmd_path = pytesseract.pytesseract_cmd
if cmd_path and os.path.exists(cmd_path):
    print(f"  ✓ Löydetty: {cmd_path}")
else:
    print(f"  ✗ EI LÖYDETTY: {cmd_path}")
    print()
    print("  Etsi oikea polku:")
    for p in ["C:\\Program Files\\Tesseract-OCR", "C:\\Program Files (x86)\\Tesseract-OCR"]:
        if os.path.exists(os.path.join(p, "tesseract.exe")):
            print(f"    Löytyi: {p}\\tesseract.exe")
    sys.exit(1)

# 3. Check if pytesseract can call tesseract
print()
print("[3] Testataan pytesseract.get_tesseract_version()...")
try:
    version = pytesseract.get_tesseract_version()
    print(f"  ✓ Tesseract versio: {version}")
except Exception as e:
    print(f"  ✗ Virhe: {e}")
    sys.exit(1)

# 4. Check ocr_menu_import module
print()
print("[4] Tarkistetaan ocr_menu_import moduuli...")
try:
    from ocr_menu_import import _tesseract_available
    print(f"  ✓ ocr_menu_import importti OK")
    available = _tesseract_available()
    print(f"    _tesseract_available() = {available}")
    if not available:
        print(f"  ⚠ Tesseract merkitty saatamattomuudeksi!")
except Exception as e:
    print(f"  ✗ Virhe: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print()
print("✓ KAIKKI OK — OCR pitäisi toimia")
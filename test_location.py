print("=" * 70)
print("TESTING RUBAVU LOCATION ENGINE")
print("=" * 70)

from pathlib import Path
import pandas as pd

from location_engine import (
    normalize_columns,
    create_master_lookup,
    normalize_locations,
    validate_structure
)

BASE_DIR = Path(__file__).resolve().parent

MASTER_FILE = (
    BASE_DIR
    / "data"
    / "baseline"
    / "RubavunewAugust.xlsx"
)

print()
print("Baseline file:")
print(MASTER_FILE)

if not MASTER_FILE.exists():
    print()
    print("❌ BASELINE FILE NOT FOUND")
    raise SystemExit()

print()
print("Loading baseline...")

master = pd.read_excel(MASTER_FILE)

print(f"Rows loaded: {len(master):,}")

master = normalize_columns(master)

print()
print("Columns:")
for col in master.columns:
    print(f" - {col}")

print()
print("Creating location lookup...")

lookup = create_master_lookup(master)

print(f"Official sectors: {len(lookup['sector'])}")
print(f"Official cells: {len(lookup['cell'])}")
print(f"Official villages: {len(lookup['village'])}")

print()
print("Normalizing locations...")

normalized = normalize_locations(master, lookup)

print()
print("Validating structure...")

validation = validate_structure(normalized)

print()
print("=" * 70)
print("FINAL RESULT")
print("=" * 70)

print(f"Sectors: {validation['sectors']}")
print(f"Cells: {validation['cells']}")
print(f"Villages: {validation['villages']}")

if validation["valid"]:
    print()
    print("✅ LOCATION ENGINE TEST PASSED")
    print("✓ 12 Sectors")
    print("✓ 80 Cells")
    print("✓ 519 Villages")
else:
    print()
    print("❌ LOCATION ENGINE TEST FAILED")

print("=" * 70)
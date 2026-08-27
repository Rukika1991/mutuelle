import pandas as pd

from config.settings import MASTER_FILE

from services.location_service import (
    normalize_columns,
    create_master_lookup,
    load_location_mapping,
    normalize_locations,
    validate_rubavu_structure,
    structure_summary,
)


# Load official baseline
master = pd.read_excel(MASTER_FILE)

master = normalize_columns(master)

print("\nOFFICIAL BASELINE COLUMNS:")
print(master.columns.tolist())


# Create official lookup
lookup = create_master_lookup(master)


# Load location mapping
mapping = load_location_mapping()


# Normalize the master itself
master = normalize_locations(
    master,
    lookup,
    mapping
)


# Validate
validation = validate_rubavu_structure(master)

print("\nRUBAVU STRUCTURE")
print("=" * 50)

print(
    structure_summary(validation)
)

print("\nVALID:", validation["valid"])
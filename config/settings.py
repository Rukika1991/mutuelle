from pathlib import Path


# ============================================================
# PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data"

BASELINE_DIR = DATA_DIR / "baseline"
UPLOADS_DIR = DATA_DIR / "uploads"
PROCESSED_DIR = DATA_DIR / "processed"

OUTPUT_DIR = BASE_DIR / "output"


# ============================================================
# IMPORTANT FILES
# ============================================================

MASTER_FILE = BASELINE_DIR / "RubavunewAugust.xlsx"

LOCATION_MAPPING_FILE = BASE_DIR / "location_mapping.xlsx"


# ============================================================
# OFFICIAL RUBAVU STRUCTURE
# ============================================================

EXPECTED_SECTORS = 12
EXPECTED_CELLS = 80
EXPECTED_VILLAGES = 519


# ============================================================
# CONTRIBUTION LEVELS
# ============================================================

LEVEL_CONTRIBUTIONS = {
    1: 0,
    2: 4000,
    3: 5000,
    4: 8000,
    5: 20000,
}


# ============================================================
# CREATE REQUIRED DIRECTORIES
# ============================================================

for directory in [
    DATA_DIR,
    BASELINE_DIR,
    UPLOADS_DIR,
    PROCESSED_DIR,
    OUTPUT_DIR,
]:
    directory.mkdir(parents=True, exist_ok=True)
    
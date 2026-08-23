from pathlib import Path


# ============================================================
# PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data"

BASELINE_DIR = DATA_DIR / "baseline"
PROGRESS_DIR = DATA_DIR / "progress"
PROCESSED_DIR = DATA_DIR / "processed"

OUTPUT_DIR = BASE_DIR / "output"


# ============================================================
# CREATE DIRECTORIES
# ============================================================

for directory in [
    BASELINE_DIR,
    PROGRESS_DIR,
    PROCESSED_DIR,
    OUTPUT_DIR
]:
    directory.mkdir(
        parents=True,
        exist_ok=True
    )


# ============================================================
# MUTUELLE CONTRIBUTION RULES
# ============================================================

LEVEL_CONTRIBUTIONS = {
    1: 0,
    2: 4000,
    3: 5000,
    4: 8000,
    5: 20000
}
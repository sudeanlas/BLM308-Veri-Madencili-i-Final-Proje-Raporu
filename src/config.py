from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "veri"
FIGURES_DIR = PROJECT_ROOT / "figures"
MODELS_DIR = PROJECT_ROOT / "models"
OUTPUT_DIR = PROJECT_ROOT / "teslim"

RAW_DATA_PATH = DATA_DIR / "bank-additional-full.csv"
MODEL_PATH = MODELS_DIR / "random_forest_duration_haric.joblib"
METRICS_PATH = OUTPUT_DIR / "metrics.csv"

RANDOM_STATE = 42
TARGET_COLUMN = "y"
LEAKAGE_COLUMN = "duration"

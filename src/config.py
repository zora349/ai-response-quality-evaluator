from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_PATH = BASE_DIR / "data" / "auto_eval_samples.csv"
OUTPUT_PATH = BASE_DIR / "output" / "auto_eval_result.xlsx"
REPORT_PATH = BASE_DIR / "output" / "auto_eval_report.md"
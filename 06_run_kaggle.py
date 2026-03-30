"""
Kaggle Dataset — Data Quality Pipeline
=======================================
Downloads the 'loans-and-liability' dataset from Kaggle using kagglehub,
then runs the full profiling + validation pipeline automatically.

Usage:
    python run_kaggle.py

Requirements:
    pip install kagglehub pandas numpy openpyxl jinja2 scipy matplotlib

Kaggle credentials:
    Set KAGGLE_USERNAME and KAGGLE_KEY as environment variables, or place
    your kaggle.json in ~/.kaggle/kaggle.json before running.
    See: https://www.kaggle.com/settings → API → Create New Token
"""

import os
import glob
import time
import kagglehub
import pandas as pd

from logger import get_logger, log_pipeline_start, log_pipeline_end
from data_profiler  import run as run_profile
from data_validator import run as run_validate


# ── DATASET-SPECIFIC VALIDATION CONFIG ───────────────────────────────────────
# Tailored to the loans-and-liability dataset schema.
# Adjust allowed values and ranges based on what you find after profiling.

LOANS_CONFIG = {
    "required_columns": [
        "credit.policy", "purpose", "int.rate", "installment",
        "log.annual.inc", "dti", "fico", "days.with.cr.line",
        "revol.bal", "revol.util", "inq.last.6mths",
        "delinq.2yrs", "pub.rec", "not.fully.paid",
    ],
    "not_null_columns": [
        "credit.policy", "int.rate", "installment",
        "fico", "dti", "not.fully.paid",
    ],
    "unique_columns": [],   # No natural unique key in this dataset
    "numeric_columns": {
        "int.rate":         {"min": 0.0,  "max": 1.0},    # stored as decimal e.g. 0.12
        "installment":      {"min": 0,    "max": 5000},
        "dti":              {"min": 0,    "max": 100},
        "fico":             {"min": 300,  "max": 850},
        "revol.util":       {"min": 0,    "max": 150},    # can exceed 100 in edge cases
        "inq.last.6mths":   {"min": 0,    "max": 30},
        "delinq.2yrs":      {"min": 0,    "max": 30},
        "pub.rec":          {"min": 0,    "max": 20},
        "not.fully.paid":   {"min": 0,    "max": 1},
        "credit.policy":    {"min": 0,    "max": 1},
    },
    "categorical_columns": {
        "purpose": [
            "credit_card", "debt_consolidation", "educational",
            "home_improvement", "major_purchase", "small_business",
            "all_other",
        ],
    },
    "date_columns":  [],   # No date columns in this dataset
    "date_range":    {},
    "consistency_rules": [
        # FICO score should be plausible given credit policy:
        # credit.policy=1 means borrower meets underwriting criteria (typically FICO >= 660)
        # We flag rows where policy=1 but fico < 600 as suspicious
    ],
    "warn_threshold": 0.5,
    "fail_threshold": 2.0,
}


# ── HELPERS ───────────────────────────────────────────────────────────────────

def find_csv(path: str) -> str:
    """Find the first CSV file in the downloaded dataset directory."""
    csvs = glob.glob(os.path.join(path, "**", "*.csv"), recursive=True)
    if not csvs:
        raise FileNotFoundError(f"No CSV files found in {path}")
    # Prefer the largest file if there are multiple
    csvs.sort(key=lambda p: os.path.getsize(p), reverse=True)
    return csvs[0]


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    """
    Light preprocessing before profiling/validation:
      - Strip whitespace from column names
      - Convert percentage strings to float if needed
      - Standardise boolean-like integer columns
    """
    df.columns = df.columns.str.strip()

    # int.rate may be stored as "12.00%" or 0.12 depending on version
    if "int.rate" in df.columns and df["int.rate"].dtype == object:
        df["int.rate"] = df["int.rate"].str.replace("%", "").astype(float) / 100

    return df


def print_dataset_summary(df: pd.DataFrame, csv_path: str):
    print(f"\n── Dataset snapshot ─────────────────────────────────")
    print(f"   File    : {os.path.basename(csv_path)}")
    print(f"   Rows    : {len(df):,}")
    print(f"   Columns : {len(df.columns)}")
    print(f"   Columns : {list(df.columns)}")
    print(f"\n   First 3 rows:")
    print(df.head(3).to_string(index=False))
    print()


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    log          = get_logger()
    output_dir   = "reports_loans"
    dataset_name = "Loans and Liability"
    t0           = time.time()

    # ── Step 1: Download from Kaggle ──────────────────────────────────────────
    log.info("Downloading dataset from Kaggle ...")
    path = kagglehub.dataset_download("matinmahmoudi/loans-and-liability")
    log.info(f"Downloaded to: {path}")

    csv_path = find_csv(path)
    log.info(f"CSV file: {csv_path}")

    # ── Step 2: Load and preprocess ───────────────────────────────────────────
    df = pd.read_csv(csv_path)
    df = preprocess(df)
    log_pipeline_start(log, "Kaggle Loans Pipeline", csv_path, len(df), len(df.columns))
    log.info(f"Columns: {list(df.columns)}")

    os.makedirs(output_dir, exist_ok=True)
    local_csv = os.path.join(output_dir, "loans_preprocessed.csv")
    df.to_csv(local_csv, index=False)
    log.info(f"Preprocessed CSV saved → {local_csv}")

    prefix = os.path.join(output_dir, "loans")

    # ── Step 3: Profile ───────────────────────────────────────────────────────
    run_profile(local_csv, f"{prefix}_profile", dataset_name=dataset_name)

    # ── Step 4: Validate ──────────────────────────────────────────────────────
    run_validate(local_csv, f"{prefix}_validation",
                 dataset_name=dataset_name,
                 config=LOANS_CONFIG)

    # ── Final summary ─────────────────────────────────────────────────────────
    log_pipeline_end(log, "Kaggle Loans Pipeline", time.time() - t0, output_dir)
    log.info("Output files:")
    for f in sorted(os.listdir(output_dir)):
        if f == "loans_preprocessed.csv":
            continue
        size = os.path.getsize(os.path.join(output_dir, f))
        log.info(f"  {f:<50} {size/1024:.1f} KB")


if __name__ == "__main__":
    main()

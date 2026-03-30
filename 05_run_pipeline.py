"""
Data Quality Orchestrator
==========================
Runs the full data quality workflow in sequence:
    1. Profile  — understand the dataset structure and content
    2. Validate — assert quality rules and score each dimension

Usage:
    python run_pipeline.py --input transactions.csv --output_dir reports/
"""

import argparse
import os
import time
from logger import get_logger, log_pipeline_start, log_pipeline_end
from data_profiler  import run as run_profile
from data_validator import run as run_validate


def run(input_path: str, output_dir: str, dataset_name: str = None):
    log    = get_logger()
    t0     = time.time()
    name   = dataset_name or os.path.basename(input_path).replace(".csv", "")
    prefix = os.path.join(output_dir, name.replace(" ", "_"))
    os.makedirs(output_dir, exist_ok=True)

    import pandas as pd
    df = pd.read_csv(input_path)
    log_pipeline_start(log, "Data Quality Orchestrator", input_path, len(df), len(df.columns))

    log.info("── Step 1 of 2: Profiling ──────────────────────────")
    run_profile(input_path, f"{prefix}_profile", dataset_name=name)

    log.info("── Step 2 of 2: Validation ─────────────────────────")
    run_validate(input_path, f"{prefix}_validation", dataset_name=name)

    log_pipeline_end(log, "Data Quality Orchestrator", time.time() - t0, output_dir)
    log.info("Output files:")
    for f in sorted(os.listdir(output_dir)):
        size = os.path.getsize(os.path.join(output_dir, f))
        log.info(f"  {f:<50} {size/1024:.1f} KB")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Data Quality Orchestrator")
    parser.add_argument("--input",      default="transactions.csv")
    parser.add_argument("--output_dir", default="reports")
    parser.add_argument("--name",       default=None)
    args = parser.parse_args()
    run(args.input, args.output_dir, args.name)

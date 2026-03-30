# Data Quality Pipeline — Profiling + Validation

A Python pipeline for understanding and validating any tabular dataset. The pipeline runs in two sequential steps:

1. **Profiling** — explore the dataset: structure, distributions, missing values, correlations, and anomalies
2. **Validation** — assert quality rules across six dimensions and score the result

---

## Two ways to run this

### Option A — Synthetic data (no external accounts needed)

Generates a realistic financial transaction dataset with deliberately injected quality issues (nulls, negatives, duplicates, date inconsistencies, outliers), then runs the full profiling and validation pipeline on it.

```bash
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib

python 02_generate_data.py          # generates transactions.csv
python 05_run_pipeline.py           # runs profiling + validation on it
```

Use this option to explore the pipeline, understand the outputs, or test changes without needing any external data.

---

### Option B — Kaggle dataset (loans and liability)

Downloads the [`matinmahmoudi/loans-and-liability`](https://www.kaggle.com/datasets/matinmahmoudi/loans-and-liability) dataset directly from Kaggle, preprocesses it, and runs the full pipeline automatically.

```bash
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib kagglehub

python 06_run_kaggle.py
```

**Kaggle credentials required.** Before running, place your `kaggle.json` in `~/.kaggle/kaggle.json`. Get it from [kaggle.com](https://www.kaggle.com/settings) → API → Create New Token.

---

## Files

The files are numbered in execution order:

```
01_logger.py          Shared logging module — imported by all others, never run directly
02_generate_data.py   Synthetic data generator — Option A only
03_data_profiler.py   Profiling module — can also be used standalone
04_data_validator.py  Validation module — can also be used standalone
05_run_pipeline.py    Orchestrator for Option A — runs profiling + validation on any local CSV
06_run_kaggle.py      Entry point for Option B — downloads Kaggle data then runs everything
logs/                 Auto-created — one timestamped log file per pipeline run
reports/              Auto-created — profiling and validation outputs (json, xlsx, html)
```

---

## The six validation dimensions

The framework is structured around the **DQOps six-dimension model**:

| Dimension | What it checks |
|---|---|
| **Completeness** | Required columns present, mandatory fields not null, row-level completeness |
| **Validity** | Values within allowed domains, numeric ranges, correct date formats, no negatives where prohibited |
| **Uniqueness** | Key fields are free of duplicates |
| **Consistency** | Relationships between fields are logically coherent (e.g. value date ≥ transaction date) |
| **Timeliness** | Dates fall within expected ranges, data is not stale |
| **Accuracy** | Numeric values are within statistically plausible bounds (4σ outlier detection) |

Each check produces a **score from 0–100** and a **status (PASS / WARN / FAIL)**. Thresholds are configurable.

---

## Architecture

```
Input data
    │
    │   Option A                      Option B
    ├── 02_generate_data.py  ──►      06_run_kaggle.py
    │   (synthetic CSV)               (kagglehub download + preprocess)
    │                                         │
    └────────────────────────────────────────►│
                                              │
                                    05_run_pipeline.py
                                              │
                          ┌───────────────────┴──────────────────┐
                          ▼                                       ▼
               03_data_profiler.py                   04_data_validator.py
                          │                                       │
               ┌──────────┴──────────┐             ┌─────────────┴─────────────┐
               │  Per-column stats   │             │  Completeness             │
               │  Missing values     │             │  Validity                 │
               │  Distributions      │             │  Uniqueness               │
               │  Correlations       │             │  Consistency              │
               │  Inline charts      │             │  Timeliness               │
               └──────────┬──────────┘             │  Accuracy                 │
                          │                        └─────────────┬─────────────┘
                          │                                       │
                          └───────────────────┬───────────────────┘
                                              │
                              ┌───────────────┼───────────────┐
                              ▼               ▼               ▼
                           .json           .xlsx           .html
                                              │
                                          01_logger.py
                                    logs/dq_pipeline_<timestamp>.log
```

---

## Output formats

Both profiling and validation produce three output files each:

**JSON** — machine-readable, useful for integrating into broader pipelines or triggering alerts downstream.

**Excel** — human-readable formatted workbook. The profiling workbook has four sheets (Overview, Numeric Profiles, Categorical Profiles, Correlations). The validation workbook has three sheets (Dashboard, All Checks, Failures & Warnings).

**HTML** — self-contained browser-viewable report, shareable without any tooling. Includes inline charts for profiling and colour-coded status badges for validation.

---

## Logging

Every pipeline run writes a timestamped log file to `logs/dq_pipeline_YYYYMMDD_HHMMSS.log`.

The console shows INFO and above in colour. The log file captures DEBUG and above in plain text — every column profiled, every check result, every file written.

Log levels map directly to check status:

```
DEBUG    → PASS checks (visible in file, not on console)
WARNING  → WARN checks (visible on console in yellow)
ERROR    → FAIL checks (visible on console in red)
```

To see only failures from any log file:

```bash
grep ERROR logs/dq_pipeline_*.log
```

---

## Data profiling — what it measures per column

| Column type | Metrics produced |
|---|---|
| Numeric | Mean, std, min/max, P25/P50/P75/P95, skewness, kurtosis, zero count, negative count, IQR outlier count |
| Categorical | Cardinality, top-N value frequencies, % coverage |
| Datetime | Min, max, range in days |
| Text | Average and max string length, top values |

Dataset-level: overall completeness score, duplicate row detection, memory footprint, correlation matrix with automatic high-correlation warnings (|r| ≥ 0.7).

---

## Configuration

The validator ships with a default config for financial transaction data. Pass a custom config dict to `DataValidator` to adapt it to any dataset:

```python
from data_validator import DataValidator
import pandas as pd

df = pd.read_csv("my_data.csv")

config = {
    "required_columns":    ["id", "date", "amount", "status"],
    "not_null_columns":    ["id", "date", "amount"],
    "unique_columns":      ["id"],
    "numeric_columns":     {"amount": {"min": 0, "max": 1_000_000}},
    "categorical_columns": {"status": ["ACTIVE", "INACTIVE", "PENDING"]},
    "date_columns":        ["date"],
    "date_range":          {"date": {"min": "2020-01-01", "max": "today"}},
    "consistency_rules": [
        {"name": "end >= start", "col_a": "end_date", "col_b": "start_date", "op": ">="}
    ],
    "warn_threshold": 0.5,   # % issues → WARN
    "fail_threshold": 2.0,   # % issues → FAIL
}

validator = DataValidator(df, config=config, dataset_name="My Dataset")
report = validator.run()
```

The `06_run_kaggle.py` file shows a complete example of a dataset-specific config for the loans dataset.

---

## Extending the pipeline

| Extension | How |
|---|---|
| Add custom checks | Subclass `DataValidator` and add `check_*` methods |
| Connect to a database | Replace CSV ingestion with DuckDB or SQLAlchemy query |
| Schedule daily runs | Wrap `05_run_pipeline.py` in an Airflow DAG or cron job |
| Email on failure | Add `smtplib` notification when `overall_status == "FAIL"` |
| Compare runs over time | Load two JSON outputs and diff the dimension scores |
| Add a new Kaggle dataset | Copy `06_run_kaggle.py`, change the dataset slug and config |

---

## Dependencies

```bash
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib   # Option A
pip install kagglehub                                              # Option B only
```

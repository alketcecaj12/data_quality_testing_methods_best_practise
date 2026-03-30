# Data Quality Pipeline — Profiling + Validation

A Python pipeline for understanding and validating any tabular dataset.
The pipeline runs in three sequential steps:

1. **Profiling** (`03_data_profiler.py`) — explore structure, distributions, missing values, correlations
2. **Validation** (`04_data_validator.py`) — assert quality rules across six standard dimensions
3. **Advanced checks** (`05_advanced_checks.py`) — statistical depth, cross-column rules, patterns, temporal, schema drift

---

## Project structure

Files are numbered in execution order. The prefix is for human navigation only —
Python imports use the base name (e.g. `from data_validator import ...`).

```
data_quality_pipeline/
│
├── 01_logger.py              Shared logging — never run directly, imported by all others
├── 02_generate_data.py       Synthetic data generator — Option A only
├── 03_data_profiler.py       Profiling module — can be used standalone
├── 04_data_validator.py      Base validation — six DQOps dimensions
├── 05_advanced_checks.py     Advanced checks — five additional dimensions
├── 06_run_pipeline.py        Orchestrator for Option A (local / synthetic CSV)
├── 07_run_kaggle.py          Entry point for Option B (Kaggle dataset)
│
├── logs/                     Auto-created — one timestamped .log file per run
│   └── dq_pipeline_YYYYMMDD_HHMMSS.log
│
├── reports/                  Auto-created by Option A
│   ├── <dataset>_profile.json / .xlsx / .html
│   └── <dataset>_validation.json / .xlsx / .html
│
└── reports_loans/            Auto-created by Option B
    ├── loans_preprocessed.csv
    ├── loans_reference_schema.json    (saved on first run, used for drift detection)
    ├── loans_profile.json / .xlsx / .html
    ├── loans_validation.json / .xlsx / .html
    └── loans_advanced_validation.json / .xlsx / .html
```

---

## Two ways to run this

### Option A — Synthetic data (no external accounts needed)

Generates a realistic financial transaction dataset with deliberately injected
quality issues (nulls, negatives, duplicates, date inconsistencies, outliers),
then runs the full pipeline.

```bash
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib

python 02_generate_data.py      # generates transactions.csv
python 06_run_pipeline.py       # profiling + validation on it
```

### Option B — Kaggle dataset (loans and liability)

Downloads [`matinmahmoudi/loans-and-liability`](https://www.kaggle.com/datasets/matinmahmoudi/loans-and-liability)
directly, preprocesses it, and runs profiling + advanced validation automatically.

```bash
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib kagglehub

python 07_run_kaggle.py
```

**Kaggle credentials required.** Place your `kaggle.json` in `~/.kaggle/kaggle.json`
before running. Get it from [kaggle.com](https://www.kaggle.com/settings) → API → Create New Token.

---

## What each file does

### `01_logger.py`
Centralised logging module. Creates one shared logger instance used by all other
modules. Outputs to console (coloured, INFO+) and to a rotating file (plain text, DEBUG+).
Each run produces a new timestamped file in `logs/`.

```
DEBUG    → every PASS check, column profiled, file written
WARNING  → every WARN check
ERROR    → every FAIL check
```

Filter failures across all runs:
```bash
grep ERROR logs/dq_pipeline_*.log
```

---

### `02_generate_data.py`
Generates a synthetic financial transaction CSV (`transactions.csv`) with
3,000 rows and deliberately injected quality issues across all dimensions:
nulls, negatives, duplicates, date violations, future dates, statistical outliers,
invalid categorical values. Used for testing without needing real data.

---

### `03_data_profiler.py`
Explores the dataset before validation. Answers: *what do I actually have?*

**Per-column metrics:**

| Type | Metrics |
|---|---|
| Numeric | mean, std, min/max, P25/P50/P75/P95, skewness, kurtosis, zeros, negatives, IQR outliers |
| Categorical | cardinality, top-N frequencies, % coverage |
| Datetime | min, max, range in days |
| Text | avg/max length, top values |

**Dataset-level:** overall completeness, duplicate rows, memory,
correlation matrix, automatic warnings (high missing, outlier-heavy columns,
unexpected unique identifiers, high correlations).

**Outputs:** `.json` (machine-readable), `.xlsx` (4 sheets: Overview / Numeric /
Categorical / Correlations), `.html` (self-contained with inline charts).

Standalone usage:
```bash
python 03_data_profiler.py --input my_data.csv --output reports/my_profile
```

---

### `04_data_validator.py`
Validates the dataset against the **DQOps six-dimension model**.

| Dimension | What it checks |
|---|---|
| Completeness | Required columns present, mandatory fields not null, row completeness |
| Validity | Allowed domains, numeric ranges, date formats, no negatives |
| Uniqueness | Duplicate key detection |
| Consistency | Cross-field logical rules (e.g. value_date ≥ transaction_date) |
| Timeliness | Date range bounds, data freshness (staleness check) |
| Accuracy | Statistical outlier detection (4σ) |

Each check returns a score (0–100) and status (PASS / WARN / FAIL).
Thresholds are configurable (`warn_threshold`, `fail_threshold`).

**Outputs:** `.json`, `.xlsx` (3 sheets: Dashboard / All Checks / Failures & Warnings), `.html`.

Standalone usage:
```bash
python 04_data_validator.py --input my_data.csv --output reports/my_validation
```

Custom config:
```python
from data_validator import DataValidator
import pandas as pd

df = pd.read_csv("my_data.csv")
config = {
    "required_columns":    ["id", "date", "amount"],
    "not_null_columns":    ["id", "amount"],
    "unique_columns":      ["id"],
    "numeric_columns":     {"amount": {"min": 0, "max": 1_000_000}},
    "categorical_columns": {"status": ["ACTIVE", "INACTIVE"]},
    "date_columns":        ["date"],
    "date_range":          {"date": {"min": "2020-01-01", "max": "today"}},
    "consistency_rules":   [{"name": "end >= start", "col_a": "end_date",
                             "col_b": "start_date", "op": ">="}],
    "warn_threshold": 0.5,
    "fail_threshold": 2.0,
}
report = DataValidator(df, config=config, dataset_name="My Dataset").run()
```

---

### `05_advanced_checks.py`
Extends `DataValidator` with five additional check dimensions.
Drop-in replacement — same interface, same outputs.

| Dimension | Checks added |
|---|---|
| **Statistical** | Benford's Law (chi-squared), Shapiro-Wilk normality, categorical dominance |
| **Cross-column** | Conditional rules (if A=x then B not null), ratio bounds, mutual exclusivity |
| **Pattern** | Regex format validation (email, IBAN, phone, postcode), encoding/control chars, whitespace |
| **Temporal** | Date gap detection, monotonicity, outside-business-hours flagging |
| **Schema Drift** | Added/removed columns, dtype changes, cardinality drift, mean drift vs reference |

Usage:
```python
from advanced_checks import AdvancedDataValidator

advanced_config = {
    "benford_columns":       ["amount", "installment"],
    "normality_columns":     ["int.rate", "dti"],
    "categorical_balance":   {"status": 95.0},
    "conditional_rules": [
        {"name": "Approved must have FICO", "when_col": "credit.policy",
         "when_val": 1, "then_col": "fico", "then_op": "not_null"}
    ],
    "ratio_checks": [
        {"name": "installment/income", "numerator": "installment",
         "denominator": "log.annual.inc", "min": 0.0, "max": 50.0}
    ],
    "pattern_checks":             {"email_col": "email"},
    "whitespace_check_columns":   ["purpose"],
    "date_sequence_columns":      [{"col": "transaction_date", "max_gap_days": 7}],
    "reference_schema_path":      "reports/reference_schema.json",
}

validator = AdvancedDataValidator(df, config=base_config,
                                  advanced_config=advanced_config)

# Save reference schema on first run (subsequent runs detect drift)
validator.save_reference_schema("reports/reference_schema.json")

report = validator.run()
```

---

### `06_run_pipeline.py`
Orchestrator for Option A. Runs profiling then base validation on any local CSV.

```bash
python 06_run_pipeline.py --input my_data.csv --output_dir reports/ --name "Q4 Extract"
```

---

### `07_run_kaggle.py`
Entry point for Option B. Downloads the loans dataset, preprocesses it,
runs profiling, base validation, and advanced validation in sequence.
Saves a reference schema on first run; detects schema drift on subsequent runs.

```bash
python 07_run_kaggle.py
```

The loans-specific config (validation rules + advanced checks) lives inside
this file and serves as a concrete example of how to adapt the pipeline
to a real dataset.

---

## Validation dimensions — full summary

| # | Dimension | Module | Key checks |
|---|---|---|---|
| 1 | Completeness | `04` | Null checks, required columns, row completeness |
| 2 | Validity | `04` | Domain, range, date format, negatives |
| 3 | Uniqueness | `04` | Duplicate keys |
| 4 | Consistency | `04` | Cross-field logical rules |
| 5 | Timeliness | `04` | Date range, staleness |
| 6 | Accuracy | `04` | 4σ statistical outliers |
| 7 | Statistical | `05` | Benford's Law, normality, categorical balance |
| 8 | Cross-column | `05` | Conditional rules, ratios, mutual exclusivity |
| 9 | Pattern | `05` | Regex, encoding, whitespace |
| 10 | Temporal | `05` | Date gaps, monotonicity, business hours |
| 11 | Schema Drift | `05` | Column set, dtype, cardinality, mean drift |

---

## Extending the pipeline

| Goal | How |
|---|---|
| Add a custom base check | Subclass `DataValidator`, add a `check_*` method |
| Add a custom advanced check | Subclass `AdvancedDataValidator`, add a `check_*` method |
| Connect to a database | Replace CSV ingestion with DuckDB / SQLAlchemy |
| Schedule daily runs | Wrap `06_run_pipeline.py` in an Airflow DAG |
| Alert on failure | Add `smtplib` / Slack webhook when `overall_status == "FAIL"` |
| Add another Kaggle dataset | Copy `07_run_kaggle.py`, update the slug and config |
| Track quality over time | Load two `.json` report files and diff the dimension scores |

---

## Dependencies

```bash
# Core (required for all options)
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib

# Option B only
pip install kagglehub
```

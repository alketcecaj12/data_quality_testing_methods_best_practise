# 🔍 Data Quality Pipeline — Profiling + Validation

> **Understand your data before you trust it.**
> A modular Python pipeline that profiles, validates, and monitors any tabular dataset — with structured logging, Excel & HTML reports, and Kaggle integration out of the box.

---

## ✨ What it does

| Step | Module | Description |
|------|--------|-------------|
| 1️⃣ | `03_data_profiler.py` | Explore structure, distributions, missing values, correlations |
| 2️⃣ | `04_data_validator.py` | Assert quality rules across 6 standard DQOps dimensions |
| 3️⃣ | `05_advanced_checks.py` | Statistical depth, cross-column rules, patterns, temporal, schema drift |

---

## 📁 Project structure

> ⚠️ Files are numbered for human navigation only. Python imports use the base name (e.g. `from data_validator import ...`).

```
data_quality_pipeline/
│
├── 📄 01_logger.py              Shared logging — imported by all, never run directly
├── 📄 02_generate_data.py       Synthetic data generator — Option A only
├── 📄 03_data_profiler.py       Profiling module — standalone or via orchestrator
├── 📄 04_data_validator.py      Base validation — 6 DQOps dimensions
├── 📄 05_advanced_checks.py     Advanced checks — 5 additional dimensions
├── 📄 06_run_pipeline.py        🚀 Option A entry point  (local / synthetic CSV)
├── 📄 07_run_kaggle.py          🚀 Option B entry point  (Kaggle dataset)
│
├── 📂 logs/                     Auto-created — one .log file per run
│   └── dq_pipeline_YYYYMMDD_HHMMSS.log
│
├── 📂 reports/                  Auto-created by Option A
│   ├── <dataset>_profile.json / .xlsx / .html
│   └── <dataset>_validation.json / .xlsx / .html
│
└── 📂 reports_loans/            Auto-created by Option B
    ├── loans_preprocessed.csv
    ├── loans_reference_schema.json   ← saved on first run, used for drift detection
    ├── loans_profile.json / .xlsx / .html
    ├── loans_validation.json / .xlsx / .html
    └── loans_advanced_validation.json / .xlsx / .html
```

---

## 🚀 Two ways to run

### 🅰️ Option A — Synthetic data *(no accounts needed)*

Generates a realistic financial transaction dataset with deliberately injected quality issues, then runs the full pipeline on it.

```bash
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib

python 02_generate_data.py      # generates transactions.csv (3,000 rows, injected issues)
python 06_run_pipeline.py       # profiling + validation in one shot
```

### 🅱️ Option B — Kaggle dataset *(loans & liability)*

Downloads [`matinmahmoudi/loans-and-liability`](https://www.kaggle.com/datasets/matinmahmoudi/loans-and-liability) directly, preprocesses it, and runs profiling + advanced validation automatically.

```bash
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib kagglehub

python 07_run_kaggle.py
```

> 🔑 **Kaggle credentials required.** Place your `kaggle.json` in `~/.kaggle/kaggle.json` before running.
> Get it from [kaggle.com](https://www.kaggle.com/settings) → API → **Create New Token**.

---

## 📦 File-by-file breakdown

### 🪵 `01_logger.py`
Centralised logging module shared across all files. Creates one logger instance per run — console output is coloured, file output is plain text with timestamps.

```
🟦 DEBUG    → every PASS check, column profiled, file written
🟨 WARNING  → every WARN check
🟥 ERROR    → every FAIL check
```

📌 Filter failures across all historical runs:
```bash
grep ERROR logs/dq_pipeline_*.log
```

---

### 🏭 `02_generate_data.py`
Generates `transactions.csv` — 3,000 rows of synthetic financial transaction data with deliberately injected quality problems across all dimensions: nulls, negatives, duplicates, future dates, invalid categoricals, statistical outliers. Used for testing without real data.

---

### 🔬 `03_data_profiler.py`
Answers: *what do I actually have?* Runs before validation to give a full picture of the dataset.

**Per-column metrics:**

| 📊 Type | 📐 Metrics |
|---------|-----------|
| Numeric | mean, std, min/max, P25/P50/P75/P95, skewness, kurtosis, zeros, negatives, IQR outliers |
| Categorical | cardinality, top-N frequencies, % coverage |
| Datetime | min, max, range in days |
| Text | avg/max length, top values |

**Dataset-level:** overall completeness, duplicate rows, memory footprint, correlation matrix, automatic warnings.

**📤 Outputs:** `.json` · `.xlsx` (Overview / Numeric / Categorical / Correlations) · `.html` (inline charts)

```bash
# Standalone usage
python 03_data_profiler.py --input my_data.csv --output reports/my_profile
```

---

### ✅ `04_data_validator.py`
Validates against the **DQOps six-dimension model**. Each check returns a score (0–100) and a status badge.

| 🏷️ Dimension | 🔍 What it checks |
|---|---|
| 🟩 Completeness | Required columns present, mandatory fields not null, row completeness |
| 🟩 Validity | Allowed domains, numeric ranges, date formats, no negatives |
| 🟩 Uniqueness | Duplicate key detection |
| 🟩 Consistency | Cross-field logical rules (e.g. value_date ≥ transaction_date) |
| 🟩 Timeliness | Date range bounds, staleness check |
| 🟩 Accuracy | Statistical outlier detection (4σ) |

**📤 Outputs:** `.json` · `.xlsx` (Dashboard / All Checks / Failures & Warnings) · `.html`

```bash
# Standalone usage
python 04_data_validator.py --input my_data.csv --output reports/my_validation
```

Custom config example:
```python
from data_validator import DataValidator
import pandas as pd

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
    "warn_threshold": 0.5,   # % issues → ⚠️ WARN
    "fail_threshold": 2.0,   # % issues → ❌ FAIL
}
report = DataValidator(pd.read_csv("my_data.csv"), config=config).run()
```

---

### 🧪 `05_advanced_checks.py`
Drop-in replacement for `DataValidator` — adds **5 additional dimensions** on top of the base 6.

| 🏷️ Dimension | 🔍 Checks added |
|---|---|
| 📈 Statistical | Benford's Law (chi-squared), Shapiro-Wilk normality, categorical dominance |
| 🔗 Cross-column | Conditional rules (if A=x then B not null), ratio bounds, mutual exclusivity |
| 🔤 Pattern | Regex: email, IBAN, phone, postcode · encoding/control chars · whitespace |
| 🕐 Temporal | Date gap detection, monotonicity, outside-business-hours flagging |
| 🔄 Schema Drift | Added/removed columns, dtype changes, cardinality drift, mean drift vs reference |

```python
from advanced_checks import AdvancedDataValidator

advanced_config = {
    "benford_columns":       ["amount"],
    "normality_columns":     ["dti", "fico"],
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
    "date_sequence_columns":      [{"col": "date", "max_gap_days": 7}],
    "reference_schema_path":      "reports/reference_schema.json",
}

validator = AdvancedDataValidator(df, config=base_config, advanced_config=advanced_config)
validator.save_reference_schema("reports/reference_schema.json")  # first run only
report = validator.run()
```

---

### 🎛️ `06_run_pipeline.py`
Orchestrator for **Option A**. Runs profiling → base validation on any local CSV in one command.

```bash
python 06_run_pipeline.py --input my_data.csv --output_dir reports/ --name "Q4 Extract"
```

---

### 🌐 `07_run_kaggle.py`
Entry point for **Option B**. Downloads the loans dataset, preprocesses it, runs profiling → base validation → advanced validation in sequence. Saves a reference schema on the first run and detects schema drift on all subsequent runs.

```bash
python 07_run_kaggle.py
```

The loans-specific config (validation rules + advanced checks config) lives inside this file and serves as a concrete, ready-to-use example of how to adapt the pipeline to a real dataset.

---

## 📊 All 11 validation dimensions at a glance

| # | 🏷️ Dimension | 📄 Module | 🔍 Key checks |
|---|---|---|---|
| 1 | ✅ Completeness | `04` | Nulls, required columns, row completeness |
| 2 | ✅ Validity | `04` | Domain, range, date format, negatives |
| 3 | ✅ Uniqueness | `04` | Duplicate keys |
| 4 | ✅ Consistency | `04` | Cross-field logical rules |
| 5 | ✅ Timeliness | `04` | Date range, staleness |
| 6 | ✅ Accuracy | `04` | 4σ statistical outliers |
| 7 | 🧪 Statistical | `05` | Benford's Law, normality, categorical balance |
| 8 | 🔗 Cross-column | `05` | Conditional rules, ratios, mutual exclusivity |
| 9 | 🔤 Pattern | `05` | Regex, encoding, whitespace |
| 10 | 🕐 Temporal | `05` | Date gaps, monotonicity, business hours |
| 11 | 🔄 Schema Drift | `05` | Column set, dtype, cardinality, mean drift |

---

## 🔧 Extending the pipeline

| 🎯 Goal | 💡 How |
|---|---|
| Add a custom base check | Subclass `DataValidator`, add a `check_*` method |
| Add a custom advanced check | Subclass `AdvancedDataValidator`, add a `check_*` method |
| Connect to a database | Replace CSV ingestion with DuckDB / SQLAlchemy |
| ⏰ Schedule daily runs | Wrap `06_run_pipeline.py` in an Airflow DAG |
| 📬 Alert on failure | Add `smtplib` / Slack webhook when `overall_status == "FAIL"` |
| 📊 Track quality over time | Load two `.json` report files and diff the dimension scores |
| ➕ Add another Kaggle dataset | Copy `07_run_kaggle.py`, update slug and config |

---

## 📦 Dependencies

```bash
# 🅰️ Core — required for Option A and B
pip install pandas numpy openpyxl jinja2 faker scipy matplotlib

# 🅱️ Kaggle — required for Option B only
pip install kagglehub
```

---

> 💡 **Tip:** Run `grep ERROR logs/dq_pipeline_*.log` at any time to see a history of every data quality failure across all pipeline runs.

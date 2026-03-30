"""
Advanced Data Quality Checks
==============================
Extends the base DataValidator with five additional check categories:

    1. Statistical depth   — normality, Benford's Law, chi-squared categorical balance
    2. Cross-column        — conditional rules, ratio checks, mutual exclusivity
    3. Pattern / format    — regex validation for emails, IBANs, phones, postcodes, encoding
    4. Temporal            — date gap detection, monotonicity, weekend/business-day clustering
    5. Schema drift        — compare current dataset against a saved reference schema

Usage:
    from data_validator import DataValidator
    from advanced_checks import AdvancedDataValidator

    validator = AdvancedDataValidator(df, config=config, dataset_name="My Dataset")
    report = validator.run()          # runs base + all advanced checks
"""

import json
import re
import warnings
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import scipy.stats as stats

from data_validator import DataValidator, CheckResult
from logger import get_logger

warnings.filterwarnings("ignore")


# ── ADVANCED VALIDATOR ────────────────────────────────────────────────────────

class AdvancedDataValidator(DataValidator):
    """
    Drop-in replacement for DataValidator.
    Adds five new check categories on top of the base six.
    All new checks follow the same CheckResult pattern and appear
    in the same Excel / HTML / JSON outputs.
    """

    ADVANCED_CONFIG_DEFAULTS = {
        # Statistical
        "benford_columns":       [],     # Numeric cols to test against Benford's Law
        "normality_columns":     [],     # Numeric cols to test for normality (Shapiro-Wilk)
        "categorical_balance":   {},     # {col: max_dominant_pct} e.g. {"status": 95.0}

        # Cross-column
        "ratio_checks": [],
        # Each entry: {"name": "...", "numerator": col, "denominator": col,
        #              "min": float, "max": float}

        "conditional_rules": [],
        # Each entry: {"name": "...", "when_col": col, "when_val": value,
        #              "then_col": col, "then_op": "not_null"|"equals"|"in",
        #              "then_val": value_or_list}

        "mutual_exclusivity": [],
        # Each entry: {"name": "...", "cols": [col_a, col_b],
        #              "rule": "not_both_null" | "not_both_filled"}

        # Pattern / format
        "pattern_checks": {},
        # {col: "email" | "iban" | "phone" | "postcode_dk" | "postcode_se" | regex_string}

        "encoding_check_columns": [],   # Cols to scan for non-ASCII / control chars
        "whitespace_check_columns": [], # Cols to flag leading/trailing whitespace

        # Temporal
        "date_sequence_columns": [],
        # Each entry: {"col": col, "freq": "D"|"M"|"Y", "max_gap_days": int}

        "monotonic_columns": [],
        # Each entry: {"col": col, "direction": "increasing"|"decreasing"}

        "business_hours_columns": [],
        # Flags rows posted outside 07:00–18:00 Mon–Fri
        # Each entry: {"date_col": col, "time_col": col_or_None}

        # Schema drift
        "reference_schema_path": None,  # Path to a saved reference schema JSON
    }

    # Built-in regex patterns
    PATTERNS = {
        "email":       r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$",
        "iban":        r"^[A-Z]{2}\d{2}[A-Z0-9]{4,30}$",
        "phone":       r"^\+?[\d\s\-\(\)]{7,20}$",
        "postcode_dk": r"^\d{4}$",
        "postcode_se": r"^\d{3}\s?\d{2}$",
        "postcode_no": r"^\d{4}$",
        "url":         r"^https?://[^\s/$.?#].[^\s]*$",
    }

    def __init__(self, df: pd.DataFrame, config: dict = None,
                 advanced_config: dict = None, dataset_name: str = "dataset"):
        super().__init__(df, config=config, dataset_name=dataset_name)
        self.acfg = {**self.ADVANCED_CONFIG_DEFAULTS, **(advanced_config or {})}
        self.log  = get_logger()

    # ── 1. STATISTICAL DEPTH ─────────────────────────────────────────────────

    def check_benford(self):
        """
        Benford's Law: first digit of naturally occurring numbers should follow
        a known logarithmic distribution. Significant deviation can indicate
        data manipulation or extraction errors.
        """
        benford_expected = {
            str(d): np.log10(1 + 1/d) for d in range(1, 10)
        }

        for col in self.acfg["benford_columns"]:
            if col not in self.df.columns:
                continue
            numeric = pd.to_numeric(self.df[col], errors="coerce").dropna()
            numeric = numeric[numeric > 0]
            if len(numeric) < 100:
                self.log.debug(f"Benford skip '{col}': fewer than 100 positive values")
                continue

            first_digits = numeric.astype(str).str.lstrip("0.").str[0]
            first_digits = first_digits[first_digits.str.match(r"[1-9]")]
            observed     = first_digits.value_counts(normalize=True).sort_index()
            expected     = pd.Series(benford_expected)

            observed_counts = first_digits.value_counts().reindex(
                [str(i) for i in range(1, 10)], fill_value=0
            )
            _, p_value = stats.chisquare(
                f_obs=observed_counts.values,
                f_exp=[benford_expected[str(i)] * len(first_digits) for i in range(1, 10)],
            )

            deviates = p_value < 0.05
            n_issues = int(deviates)
            pct      = 100.0 if deviates else 0.0
            detail   = (
                f"Chi-squared p={p_value:.4f} — "
                f"{'DEVIATES from Benford distribution (p<0.05)' if deviates else 'conforms to Benford distribution'}"
            )
            self.results.append(CheckResult(
                dimension   = "Statistical",
                check_name  = f"Benford's Law: {col}",
                column      = col,
                status      = "WARN" if deviates else "PASS",
                score       = 50.0 if deviates else 100.0,
                n_issues    = n_issues,
                n_total     = 1,
                pct_issues  = pct,
                detail      = detail,
                sample_bad  = [],
            ))

    def check_normality(self):
        """
        Shapiro-Wilk test for normality on numeric columns.
        Highly skewed or non-normal distributions can indicate
        data quality issues or the need for transformation.
        """
        for col in self.acfg["normality_columns"]:
            if col not in self.df.columns:
                continue
            numeric = pd.to_numeric(self.df[col], errors="coerce").dropna()
            if len(numeric) < 8:
                continue
            # Shapiro-Wilk is unreliable above 5000 samples — subsample
            sample = numeric.sample(min(500, len(numeric)), random_state=42)
            _, p_value = stats.shapiro(sample)
            skew = float(stats.skew(numeric))
            kurt = float(stats.kurtosis(numeric))

            non_normal = p_value < 0.05
            detail = (
                f"Shapiro-Wilk p={p_value:.4f}, skewness={skew:.2f}, kurtosis={kurt:.2f} — "
                f"{'non-normal distribution' if non_normal else 'approximately normal'}"
            )
            self.results.append(CheckResult(
                dimension  = "Statistical",
                check_name = f"Normality: {col}",
                column     = col,
                status     = "WARN" if (non_normal and abs(skew) > 2) else "PASS",
                score      = 70.0 if (non_normal and abs(skew) > 2) else 100.0,
                n_issues   = 1 if non_normal else 0,
                n_total    = 1,
                pct_issues = 100.0 if non_normal else 0.0,
                detail     = detail,
                sample_bad = [],
            ))

    def check_categorical_balance(self):
        """
        Flags categorical columns where a single value dominates beyond
        a configured threshold — may indicate a default-fill issue or
        data extraction problem.
        """
        for col, max_pct in self.acfg["categorical_balance"].items():
            if col not in self.df.columns:
                continue
            vc      = self.df[col].value_counts(normalize=True) * 100
            top_val = vc.index[0]
            top_pct = float(vc.iloc[0])
            dominated = top_pct > max_pct
            detail = (
                f"Top value '{top_val}' = {top_pct:.1f}% of non-null rows "
                f"(threshold: {max_pct}%)"
            )
            self.results.append(CheckResult(
                dimension  = "Statistical",
                check_name = f"Categorical balance: {col}",
                column     = col,
                status     = "WARN" if dominated else "PASS",
                score      = 60.0 if dominated else 100.0,
                n_issues   = int(dominated),
                n_total    = 1,
                pct_issues = 100.0 if dominated else 0.0,
                detail     = detail,
                sample_bad = [str(top_val)] if dominated else [],
            ))

    # ── 2. CROSS-COLUMN CHECKS ────────────────────────────────────────────────

    def check_ratio_rules(self):
        """
        Validates that the ratio of two numeric columns stays within expected bounds.
        E.g. installment / loan_amnt should be between 0.01 and 0.05.
        """
        for rule in self.acfg["ratio_checks"]:
            num_col = rule["numerator"]
            den_col = rule["denominator"]
            if num_col not in self.df.columns or den_col not in self.df.columns:
                continue
            num = pd.to_numeric(self.df[num_col], errors="coerce")
            den = pd.to_numeric(self.df[den_col], errors="coerce")
            both_valid = ~num.isna() & ~den.isna() & (den != 0)
            ratio      = num / den.replace(0, np.nan)
            violation  = both_valid & (
                (ratio < rule.get("min", -np.inf)) |
                (ratio > rule.get("max",  np.inf))
            )
            self._add(
                "Cross-column",
                f"Ratio check: {rule['name']}",
                num_col,
                violation,
                f"{violation.sum()} rows where {num_col}/{den_col} outside "
                f"[{rule.get('min')}, {rule.get('max')}]",
            )

    def check_conditional_rules(self):
        """
        Conditional logic: when column A = value X, column B must satisfy condition Y.
        E.g. when credit.policy = 1 (approved), fico must not be null.
        """
        for rule in self.acfg["conditional_rules"]:
            when_col = rule["when_col"]
            then_col = rule["then_col"]
            if when_col not in self.df.columns or then_col not in self.df.columns:
                continue
            condition = self.df[when_col] == rule["when_val"]
            op        = rule["then_op"]

            if op == "not_null":
                violation = condition & self.df[then_col].isna()
            elif op == "equals":
                violation = condition & (self.df[then_col] != rule["then_val"])
            elif op == "in":
                violation = condition & ~self.df[then_col].isin(rule["then_val"])
            else:
                continue

            self._add(
                "Cross-column",
                f"Conditional: {rule['name']}",
                then_col,
                violation,
                f"When {when_col}={rule['when_val']}, {then_col} violates {op} rule",
            )

    def check_mutual_exclusivity(self):
        """
        Checks that two columns are not both null or not both filled,
        depending on the rule. E.g. a record cannot have both a
        refund_amount and a charge_amount populated.
        """
        for rule in self.acfg["mutual_exclusivity"]:
            cols = rule.get("cols", [])
            if len(cols) != 2 or not all(c in self.df.columns for c in cols):
                continue
            a_null = self.df[cols[0]].isna()
            b_null = self.df[cols[1]].isna()

            if rule["rule"] == "not_both_null":
                violation = a_null & b_null
                detail    = f"Both '{cols[0]}' and '{cols[1]}' are null"
            elif rule["rule"] == "not_both_filled":
                violation = ~a_null & ~b_null
                detail    = f"Both '{cols[0]}' and '{cols[1]}' are filled (mutually exclusive)"
            else:
                continue

            self._add(
                "Cross-column",
                f"Mutual exclusivity: {rule['name']}",
                cols[0],
                violation,
                f"{violation.sum()} rows — {detail}",
            )

    # ── 3. PATTERN / FORMAT CHECKS ────────────────────────────────────────────

    def check_patterns(self):
        """
        Regex-based format validation. Supports named patterns (email, iban,
        phone, postcode_dk/se/no, url) or any custom regex string.
        """
        for col, pattern_key in self.acfg["pattern_checks"].items():
            if col not in self.df.columns:
                continue
            regex   = self.PATTERNS.get(pattern_key, pattern_key)
            notnull = self.df[col].notna()
            bad     = notnull & ~self.df[col].astype(str).str.match(regex, na=False)
            self._add(
                "Pattern",
                f"Format check ({pattern_key}): {col}",
                col,
                bad,
                f"{bad.sum()} values do not match {pattern_key} pattern",
            )

    def check_encoding(self):
        """
        Flags non-ASCII characters, null bytes, and control characters
        in string columns — common artefacts from bad CSV exports or
        encoding mismatches.
        """
        control_char_re = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
        for col in self.acfg["encoding_check_columns"]:
            if col not in self.df.columns:
                continue
            has_control = self.df[col].astype(str).str.contains(
                control_char_re, na=False
            )
            self._add(
                "Pattern",
                f"Encoding / control chars: {col}",
                col,
                has_control,
                f"{has_control.sum()} values contain control characters or null bytes",
            )

    def check_whitespace(self):
        """
        Flags leading or trailing whitespace in string columns —
        a common source of failed joins and lookup mismatches.
        """
        for col in self.acfg["whitespace_check_columns"]:
            if col not in self.df.columns:
                continue
            notnull   = self.df[col].notna()
            has_space = notnull & (
                self.df[col].astype(str).str.startswith(" ") |
                self.df[col].astype(str).str.endswith(" ")
            )
            self._add(
                "Pattern",
                f"Whitespace: {col}",
                col,
                has_space,
                f"{has_space.sum()} values have leading or trailing whitespace",
            )

    # ── 4. TEMPORAL CHECKS ────────────────────────────────────────────────────

    def check_date_gaps(self):
        """
        Detects unexpected gaps in date sequences. E.g. a daily transaction
        feed should not be missing entire days.
        """
        for rule in self.acfg.get("date_sequence_columns", []):
            col          = rule if isinstance(rule, str) else rule.get("col")
            max_gap_days = rule.get("max_gap_days", 7) if isinstance(rule, dict) else 7
            if col not in self.df.columns:
                continue
            dates = pd.to_datetime(self.df[col], errors="coerce").dropna().sort_values()
            if len(dates) < 2:
                continue
            gaps = dates.diff().dt.days.dropna()
            large_gaps = gaps[gaps > max_gap_days]
            n_issues   = len(large_gaps)
            pct        = round(n_issues / max(len(gaps), 1) * 100, 2)
            detail     = (
                f"{n_issues} gap(s) exceeding {max_gap_days} days detected. "
                f"Max gap: {int(gaps.max())} days"
            )
            self.results.append(CheckResult(
                dimension  = "Temporal",
                check_name = f"Date gaps: {col}",
                column     = col,
                status     = self._status(pct),
                score      = self._score(pct),
                n_issues   = n_issues,
                n_total    = len(gaps),
                pct_issues = pct,
                detail     = detail,
                sample_bad = [],
            ))

    def check_monotonicity(self):
        """
        Verifies that a column is strictly increasing or decreasing
        (e.g. auto-increment IDs, cumulative balances).
        """
        for rule in self.acfg["monotonic_columns"]:
            col       = rule["col"]
            direction = rule.get("direction", "increasing")
            if col not in self.df.columns:
                continue
            numeric = pd.to_numeric(self.df[col], errors="coerce")
            diff    = numeric.diff().dropna()
            if direction == "increasing":
                violation_mask = diff < 0
            else:
                violation_mask = diff > 0
            bad = violation_mask.reindex(self.df.index, fill_value=False)
            self._add(
                "Temporal",
                f"Monotonicity ({direction}): {col}",
                col,
                bad,
                f"{bad.sum()} rows where sequence is not {direction}",
            )

    def check_business_hours(self):
        """
        Flags transactions or events posted outside business hours
        (07:00–18:00, Monday–Friday). Useful for detecting automated
        or unusual activity.
        """
        for rule in self.acfg["business_hours_columns"]:
            date_col = rule.get("date_col") if isinstance(rule, dict) else rule
            if date_col not in self.df.columns:
                continue
            dt = pd.to_datetime(self.df[date_col], errors="coerce")
            outside_hours   = ~dt.isna() & (
                (dt.dt.hour < 7) | (dt.dt.hour >= 18)
            )
            weekend         = ~dt.isna() & (dt.dt.dayofweek >= 5)
            outside_business = outside_hours | weekend
            self._add(
                "Temporal",
                f"Outside business hours: {date_col}",
                date_col,
                outside_business,
                f"{outside_business.sum()} events outside Mon–Fri 07:00–18:00",
            )

    # ── 5. SCHEMA DRIFT ───────────────────────────────────────────────────────

    def save_reference_schema(self, path: str):
        """
        Saves the current dataset schema as a reference JSON.
        Run once on a known-good dataset, then compare future runs against it.
        """
        schema = {
            "saved_at":   datetime.now().isoformat(),
            "n_columns":  len(self.df.columns),
            "columns":    {
                col: {
                    "dtype":       str(self.df[col].dtype),
                    "n_unique":    int(self.df[col].nunique()),
                    "pct_missing": round(self.df[col].isna().mean() * 100, 2),
                    "mean": round(float(self.df[col].mean()), 4)
                             if pd.api.types.is_numeric_dtype(self.df[col]) else None,
                    "std":  round(float(self.df[col].std()), 4)
                             if pd.api.types.is_numeric_dtype(self.df[col]) else None,
                }
                for col in self.df.columns
            }
        }
        with open(path, "w") as f:
            json.dump(schema, f, indent=2)
        self.log.info(f"Reference schema saved → {path}")

    def check_schema_drift(self):
        """
        Compares current dataset against a saved reference schema.
        Detects: added/removed columns, dtype changes, cardinality drift,
        missing-rate drift, and mean drift on numeric columns.
        """
        path = self.acfg.get("reference_schema_path")
        if not path:
            return
        try:
            with open(path) as f:
                ref = json.load(f)
        except FileNotFoundError:
            self.log.warning(f"Reference schema not found at '{path}' — skipping drift checks")
            return

        ref_cols = set(ref["columns"].keys())
        cur_cols = set(self.df.columns)

        # Added / removed columns
        added   = cur_cols - ref_cols
        removed = ref_cols - cur_cols
        if added or removed:
            self.results.append(CheckResult(
                dimension  = "Schema Drift",
                check_name = "Column set changed",
                column     = "(schema)",
                status     = "FAIL" if removed else "WARN",
                score      = 0.0 if removed else 60.0,
                n_issues   = len(added) + len(removed),
                n_total    = len(ref_cols),
                pct_issues = round((len(added) + len(removed)) / max(len(ref_cols), 1) * 100, 1),
                detail     = (
                    f"Added: {sorted(added) or '—'}  |  Removed: {sorted(removed) or '—'}"
                ),
                sample_bad = sorted(removed),
            ))

        # Per-column drift for columns present in both
        for col in ref_cols & cur_cols:
            ref_col = ref["columns"][col]
            issues  = []

            # Dtype change
            cur_dtype = str(self.df[col].dtype)
            if cur_dtype != ref_col["dtype"]:
                issues.append(f"dtype {ref_col['dtype']} → {cur_dtype}")

            # Missing rate drift (>5pp change)
            cur_missing = round(self.df[col].isna().mean() * 100, 2)
            delta_miss  = abs(cur_missing - ref_col["pct_missing"])
            if delta_miss > 5:
                issues.append(
                    f"missing rate {ref_col['pct_missing']}% → {cur_missing}% "
                    f"(Δ {delta_miss:.1f}pp)"
                )

            # Cardinality drift (>20% relative change)
            cur_unique  = self.df[col].nunique()
            ref_unique  = ref_col["n_unique"]
            if ref_unique > 0:
                card_drift = abs(cur_unique - ref_unique) / ref_unique * 100
                if card_drift > 20:
                    issues.append(
                        f"cardinality {ref_unique} → {cur_unique} "
                        f"({card_drift:.0f}% change)"
                    )

            # Mean drift on numerics (>10% relative change)
            if pd.api.types.is_numeric_dtype(self.df[col]) and ref_col.get("mean") is not None:
                cur_mean = self.df[col].mean()
                ref_mean = ref_col["mean"]
                if ref_mean != 0:
                    mean_drift = abs(cur_mean - ref_mean) / abs(ref_mean) * 100
                    if mean_drift > 10:
                        issues.append(
                            f"mean {ref_mean:.2f} → {cur_mean:.2f} "
                            f"({mean_drift:.0f}% drift)"
                        )

            if issues:
                self.results.append(CheckResult(
                    dimension  = "Schema Drift",
                    check_name = f"Column drift: {col}",
                    column     = col,
                    status     = "WARN",
                    score      = 70.0,
                    n_issues   = len(issues),
                    n_total    = 4,
                    pct_issues = round(len(issues) / 4 * 100, 1),
                    detail     = "  |  ".join(issues),
                    sample_bad = [],
                ))

    # ── RUN ALL (base + advanced) ─────────────────────────────────────────────

    def run(self):
        report = super().run()   # runs all six base checks

        self.log.info("Running advanced checks ...")

        # 1. Statistical
        self.check_benford()
        self.check_normality()
        self.check_categorical_balance()

        # 2. Cross-column
        self.check_ratio_rules()
        self.check_conditional_rules()
        self.check_mutual_exclusivity()

        # 3. Pattern / format
        self.check_patterns()
        self.check_encoding()
        self.check_whitespace()

        # 4. Temporal
        self.check_date_gaps()
        self.check_monotonicity()
        self.check_business_hours()

        # 5. Schema drift
        self.check_schema_drift()

        # Rebuild report with advanced checks included
        import numpy as np
        all_dims = list(report.dimension_scores.keys()) + [
            "Statistical", "Cross-column", "Pattern", "Temporal", "Schema Drift"
        ]
        dim_scores = {}
        for dim in all_dims:
            dim_checks = [r for r in self.results if r.dimension == dim]
            dim_scores[dim] = round(
                np.mean([r.score for r in dim_checks]) if dim_checks else 100.0, 1
            )

        overall = round(np.mean(list(dim_scores.values())), 1)
        report.dimension_scores = dim_scores
        report.overall_score    = overall
        report.overall_status   = "PASS" if overall >= 90 else "WARN" if overall >= 75 else "FAIL"
        report.checks           = self.results

        adv_checks = [r for r in self.results if r.dimension in
                      ("Statistical", "Cross-column", "Pattern", "Temporal", "Schema Drift")]
        n_adv_fail = sum(1 for r in adv_checks if r.status == "FAIL")
        n_adv_warn = sum(1 for r in adv_checks if r.status == "WARN")
        self.log.info(
            f"Advanced checks complete  │  {len(adv_checks)} checks: "
            f"{len(adv_checks)-n_adv_fail-n_adv_warn} PASS / "
            f"{n_adv_warn} WARN / {n_adv_fail} FAIL"
        )
        return report

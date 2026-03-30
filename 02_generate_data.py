"""
Synthetic financial transaction data generator.
Produces a realistic dataset with deliberately injected data quality issues
across all six DQOps dimensions for validation pipeline testing.
"""

import pandas as pd
import numpy as np
from faker import Faker
from datetime import datetime, timedelta

fake = Faker()
np.random.seed(42)

CURRENCIES = ["DKK", "EUR", "USD", "GBP", "SEK", "NOK"]
TRANSACTION_TYPES = ["PAYMENT", "TRANSFER", "REFUND", "FEE", "ADJUSTMENT"]
STATUSES = ["COMPLETED", "PENDING", "FAILED", "REVERSED"]
COUNTRIES = ["DK", "SE", "NO", "FI", "DE", "GB", "US", "FR"]
ACCOUNT_PREFIXES = ["DK", "SE", "NO", "FI"]


def random_account():
    prefix = np.random.choice(ACCOUNT_PREFIXES)
    return f"{prefix}{np.random.randint(10000000, 99999999)}"


def random_date(start="2024-01-01", end="2024-12-31"):
    s = datetime.strptime(start, "%Y-%m-%d")
    e = datetime.strptime(end, "%Y-%m-%d")
    return s + timedelta(days=np.random.randint(0, (e - s).days))


def generate_transactions(n=3000):
    records = []

    for i in range(n):
        amount = abs(np.random.lognormal(mean=6, sigma=1.8))
        tx_date = random_date()

        records.append({
            "transaction_id":   f"TXN{i+1:07d}",
            "transaction_date": tx_date.strftime("%Y-%m-%d"),
            "value_date":       (tx_date + timedelta(days=np.random.randint(0, 3))).strftime("%Y-%m-%d"),
            "amount":           round(amount, 2),
            "currency":         np.random.choice(CURRENCIES, p=[0.5, 0.2, 0.15, 0.05, 0.05, 0.05]),
            "transaction_type": np.random.choice(TRANSACTION_TYPES),
            "status":           np.random.choice(STATUSES, p=[0.85, 0.08, 0.04, 0.03]),
            "sender_account":   random_account(),
            "receiver_account": random_account(),
            "country_code":     np.random.choice(COUNTRIES),
            "description":      fake.sentence(nb_words=5),
            "created_by":       f"user_{np.random.randint(1, 50):03d}",
            "created_at":       tx_date.strftime("%Y-%m-%d %H:%M:%S"),
        })

    df = pd.DataFrame(records)

    # ── Inject data quality issues ────────────────────────────────────────────

    idx = df.index.tolist()
    np.random.shuffle(idx)

    # Completeness: null amounts (~2%)
    df.loc[idx[:60], "amount"] = np.nan

    # Completeness: missing currency (~1%)
    df.loc[idx[60:90], "currency"] = np.nan

    # Completeness: missing description (~3%)
    df.loc[idx[90:180], "description"] = np.nan

    # Validity: negative amounts (~1%)
    df.loc[idx[180:210], "amount"] = -abs(df.loc[idx[180:210], "amount"])

    # Validity: invalid currency codes (~0.5%)
    df.loc[idx[210:225], "currency"] = np.random.choice(["XX", "ZZ", "??"], 15)

    # Validity: invalid country codes (~0.5%)
    df.loc[idx[225:240], "country_code"] = np.random.choice(["XX", "ZZZ", "123"], 15)

    # Validity: invalid status values (~0.5%)
    df.loc[idx[240:255], "status"] = np.random.choice(["UNKNOWN", "NULL", "ERROR"], 15)

    # Uniqueness: duplicate transaction IDs (~1%)
    dup_idx = idx[255:285]
    df.loc[dup_idx, "transaction_id"] = df.loc[idx[:30], "transaction_id"].values

    # Consistency: value_date before transaction_date (~1%)
    for i in idx[285:315]:
        tx = datetime.strptime(df.at[i, "transaction_date"], "%Y-%m-%d")
        df.at[i, "value_date"] = (tx - timedelta(days=np.random.randint(1, 5))).strftime("%Y-%m-%d")

    # Timeliness: transactions with future dates (~0.5%)
    future = datetime.today() + timedelta(days=np.random.randint(30, 365))
    df.loc[idx[315:330], "transaction_date"] = future.strftime("%Y-%m-%d")

    # Accuracy: extreme outlier amounts (~0.3%)
    df.loc[idx[330:340], "amount"] = np.random.uniform(1e9, 1e10, 10)

    return df


if __name__ == "__main__":
    import os
    os.makedirs("/home/claude/data_validation", exist_ok=True)
    df = generate_transactions()
    df.to_csv("/home/claude/data_validation/transactions.csv", index=False)
    total_issues = (
        df["amount"].isna().sum() +
        df["currency"].isna().sum() +
        (df["amount"] < 0).sum()
    )
    print(f"Generated {len(df):,} transactions with quality issues injected.")
    print(f"Sample issue count (nulls + negatives): {total_issues}")

"""
Data Quality Pipeline — Logger
================================
Centralised logging for the profiling, validation, and Kaggle pipelines.

Features:
    - Structured log levels: DEBUG / INFO / WARNING / ERROR / CRITICAL
    - Simultaneous output to console (coloured) and rotating log file
    - Per-run log file named with timestamp: logs/dq_pipeline_YYYYMMDD_HHMMSS.log
    - One shared logger instance imported by all pipeline modules
    - Event-specific helpers: log_check_result(), log_dimension_summary(),
      log_profile_column(), log_pipeline_start(), log_pipeline_end()

Usage:
    from logger import get_logger
    log = get_logger()
    log.info("Starting pipeline")
    log.warning("Column 'amount' has 5.2% missing values")
"""

import logging
import os
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler

# ── ANSI colour codes for console output ─────────────────────────────────────

RESET  = "\033[0m"
BOLD   = "\033[1m"
GREY   = "\033[90m"
CYAN   = "\033[96m"
GREEN  = "\033[92m"
YELLOW = "\033[93m"
RED    = "\033[91m"
BRED   = "\033[1;91m"

LEVEL_COLOURS = {
    "DEBUG":    GREY,
    "INFO":     CYAN,
    "WARNING":  YELLOW,
    "ERROR":    RED,
    "CRITICAL": BRED,
}

STATUS_COLOURS = {
    "PASS": GREEN,
    "WARN": YELLOW,
    "FAIL": RED,
}


# ── Coloured console formatter ────────────────────────────────────────────────

class ColouredFormatter(logging.Formatter):
    """Coloured formatter for console output only."""

    FMT = "{colour}{level:<8}{reset} {grey}{time}{reset}  {msg}"

    def format(self, record: logging.LogRecord) -> str:
        colour = LEVEL_COLOURS.get(record.levelname, RESET)
        time   = datetime.fromtimestamp(record.created).strftime("%H:%M:%S")
        return self.FMT.format(
            colour=colour,
            level=record.levelname,
            reset=RESET,
            grey=GREY,
            time=time,
            msg=record.getMessage(),
        )


# ── Plain file formatter ──────────────────────────────────────────────────────

class PlainFormatter(logging.Formatter):
    """Plain text formatter for log files."""
    def format(self, record: logging.LogRecord) -> str:
        time = datetime.fromtimestamp(record.created).strftime("%Y-%m-%d %H:%M:%S")
        return f"{time}  {record.levelname:<8}  {record.getMessage()}"


# ── Logger factory ────────────────────────────────────────────────────────────

_logger_instance: logging.Logger = None


def get_logger(
    name:      str  = "dq_pipeline",
    log_dir:   str  = "logs",
    level:     int  = logging.DEBUG,
    max_bytes: int  = 5 * 1024 * 1024,   # 5 MB per log file
    backup_count: int = 3,
) -> logging.Logger:
    """
    Return the shared logger instance, creating it on first call.
    Subsequent calls return the same instance regardless of arguments.
    """
    global _logger_instance
    if _logger_instance is not None:
        return _logger_instance

    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.propagate = False

    # Console handler — coloured, INFO and above
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    console.setFormatter(ColouredFormatter())
    logger.addHandler(console)

    # File handler — plain text, DEBUG and above, rotating
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path  = os.path.join(log_dir, f"dq_pipeline_{timestamp}.log")
    file_handler = RotatingFileHandler(
        log_path, maxBytes=max_bytes, backupCount=backup_count, encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(PlainFormatter())
    logger.addHandler(file_handler)

    logger.info(f"Logger initialised → {log_path}")
    _logger_instance = logger
    return logger


def reset_logger():
    """Force a new logger instance on next get_logger() call (useful for tests)."""
    global _logger_instance
    _logger_instance = None


# ── Structured event helpers ──────────────────────────────────────────────────

def log_pipeline_start(log: logging.Logger, pipeline: str, source: str, n_rows: int, n_cols: int):
    log.info("─" * 55)
    log.info(f"PIPELINE START  │  {pipeline}")
    log.info(f"Source          │  {source}")
    log.info(f"Dataset size    │  {n_rows:,} rows × {n_cols} columns")
    log.info("─" * 55)


def log_pipeline_end(log: logging.Logger, pipeline: str, elapsed_sec: float, output_dir: str):
    log.info("─" * 55)
    log.info(f"PIPELINE END    │  {pipeline}")
    log.info(f"Elapsed         │  {elapsed_sec:.2f}s")
    log.info(f"Outputs in      │  {os.path.abspath(output_dir)}")
    log.info("─" * 55)


def log_check_result(log: logging.Logger, dimension: str, check: str,
                     column: str, status: str, n_issues: int,
                     n_total: int, pct: float, detail: str):
    """Log a single validation check result at the appropriate level."""
    colour = STATUS_COLOURS.get(status, RESET)
    msg = (
        f"[{dimension:<13}] [{colour}{status}{RESET}]  "
        f"{check}  │  col={column}  "
        f"issues={n_issues:,}/{n_total:,} ({pct:.1f}%)  │  {detail}"
    )
    if status == "FAIL":
        log.error(msg)
    elif status == "WARN":
        log.warning(msg)
    else:
        log.debug(msg)


def log_dimension_summary(log: logging.Logger, scores: dict):
    """Log the per-dimension score table."""
    log.info("Dimension scores:")
    for dim, score in scores.items():
        status = "PASS" if score >= 90 else "WARN" if score >= 75 else "FAIL"
        colour = STATUS_COLOURS.get(status, RESET)
        bar    = "█" * int(score // 5) + "░" * (20 - int(score // 5))
        log.info(f"  {dim:<15}  {bar}  {score:>5.1f}  [{colour}{status}{RESET}]")


def log_profile_column(log: logging.Logger, name: str, inferred: str,
                       pct_missing: float, n_unique: int, extra: str = ""):
    """Log a single column profile summary at DEBUG level."""
    miss_flag = f"  ⚠ {pct_missing:.1f}% missing" if pct_missing > 5 else ""
    log.debug(
        f"  col={name:<28}  type={inferred:<12}  "
        f"unique={n_unique:>6,}{miss_flag}  {extra}"
    )


def log_warnings(log: logging.Logger, warnings_list: list):
    if not warnings_list:
        log.info("No dataset-level warnings.")
        return
    log.warning(f"{len(warnings_list)} dataset warning(s):")
    for w in warnings_list:
        log.warning(f"  ⚠  {w}")

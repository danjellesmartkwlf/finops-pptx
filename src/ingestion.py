"""Redshift data ingestion module.

Handles all data loading from Redshift for the FinOps Report Generator.
Reads bucket definitions from config.yaml, builds SQL queries dynamically,
and returns aggregated monthly cost data per bucket.
"""

from __future__ import annotations

import logging
import os
from typing import Any

import psycopg2
import yaml
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Redshift connection
# ---------------------------------------------------------------------------

def get_redshift_connection() -> psycopg2.extensions.connection:
    """Return a new psycopg2 connection to Redshift.

    Connection parameters are read from environment variables:
        REDSHIFT_HOST, REDSHIFT_PORT, REDSHIFT_DATABASE,
        REDSHIFT_USER, REDSHIFT_PASSWORD, REDSHIFT_SCHEMA.

    Returns:
        A psycopg2 connection object.

    Raises:
        psycopg2.OperationalError: If the connection cannot be established.
        EnvironmentError: If required environment variables are missing.
    """
    required_vars = [
        "REDSHIFT_HOST",
        "REDSHIFT_DATABASE",
        "REDSHIFT_USER",
        "REDSHIFT_PASSWORD",
    ]
    missing = [v for v in required_vars if not os.getenv(v)]
    if missing:
        raise EnvironmentError(
            f"Missing required environment variables: {', '.join(missing)}"
        )

    conn = psycopg2.connect(
        host=os.getenv("REDSHIFT_HOST"),
        port=int(os.getenv("REDSHIFT_PORT", "5439")),
        dbname=os.getenv("REDSHIFT_DATABASE"),
        user=os.getenv("REDSHIFT_USER"),
        password=os.getenv("REDSHIFT_PASSWORD"),
        options=f"-c search_path={os.getenv('REDSHIFT_SCHEMA', 'public')}",
    )
    return conn


def check_redshift_connection() -> tuple[bool, str]:
    """Test Redshift connectivity by executing a simple query.

    Returns:
        A tuple of (success: bool, message: str).
    """
    try:
        conn = get_redshift_connection()
        with conn.cursor() as cur:
            cur.execute("SELECT 1")
        conn.close()
        return True, "Redshift connection successful."
    except Exception as exc:
        return False, f"Redshift connection failed: {exc}"


# ---------------------------------------------------------------------------
# Shared connection management
# ---------------------------------------------------------------------------

_connection: psycopg2.extensions.connection | None = None


def get_shared_connection() -> psycopg2.extensions.connection:
    """Return a shared Redshift connection, creating one if needed.

    Reuses an existing connection if it's still open. Creates a new one
    if the connection is None or closed.
    """
    global _connection
    if _connection is None or _connection.closed:
        _connection = get_redshift_connection()
    return _connection


def close_shared_connection() -> None:
    """Close the shared connection if open."""
    global _connection
    if _connection is not None and not _connection.closed:
        _connection.close()
        _connection = None


# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------

def load_config(config_path: str = "config.yaml") -> dict[str, Any]:
    """Read and return the YAML configuration file.

    Args:
        config_path: Path to the YAML configuration file.

    Returns:
        A dictionary representing the parsed YAML content.

    Raises:
        FileNotFoundError: If the config file does not exist.
        yaml.YAMLError: If the file contains invalid YAML.
    """
    with open(config_path, "r") as fh:
        config: dict[str, Any] = yaml.safe_load(fh)
    return config


# ---------------------------------------------------------------------------
# Internal query builders
# ---------------------------------------------------------------------------

def _build_month_filter(date_column: str, month: str, year: int) -> str:
    """Return a SQL clause restricting *date_column* to a calendar month.

    Uses DATE_TRUNC so that any timestamp within the month matches.

    Args:
        date_column: The column name holding the date/timestamp.
        month: Two-digit month string (e.g. '01', '12').
        year: Four-digit year.

    Returns:
        A SQL WHERE-clause fragment (without leading AND).
    """
    return f"DATE_TRUNC('month', {date_column}) = '{year}-{month}-01'"


def _build_simple_query(
    source: dict[str, Any],
    month: str,
    year: int,
) -> str:
    """Build a SUM query for a single source from config.yaml.

    Args:
        source: A single source dict from config.yaml.
        month: Two-digit month string.
        year: Four-digit year.

    Returns:
        A complete SQL SELECT statement.
    """
    table = source["table"]
    cost_column = source["cost_column"]
    date_column = source["date_column"]
    filters: list[str] = source.get("filters", [])

    where_parts = [_build_month_filter(date_column, month, year)]
    where_parts.extend(filters)

    where_clause = " AND ".join(where_parts)
    return f"SELECT COALESCE(SUM({cost_column}), 0) AS total_cost FROM {table} WHERE {where_clause}"


# ---------------------------------------------------------------------------
# Query execution helper
# ---------------------------------------------------------------------------

def _execute_scalar_query(
    conn: psycopg2.extensions.connection,
    sql: str,
) -> float:
    """Execute a SQL query and return the first column of the first row as a float.

    Args:
        conn: An open psycopg2 connection.
        sql: The SQL statement to execute (should return a single scalar).

    Returns:
        The query result as a float, or 0.0 if NULL.
    """
    logger.debug("Executing SQL:\n%s", sql)
    with conn.cursor() as cur:
        cur.execute(sql)
        row = cur.fetchone()
    value = row[0] if row and row[0] is not None else 0.0
    return float(value)


def _fetch_source_cost(
    conn: psycopg2.extensions.connection,
    source: dict[str, Any],
    month: str,
    year: int,
) -> float:
    """Compute cost for a single source definition from config.yaml.

    Args:
        conn: An open psycopg2 connection.
        source: A single source dict from a bucket definition.
        month: Two-digit month string.
        year: Four-digit year.

    Returns:
        The total cost as a float for this source/month.
    """
    sql = _build_simple_query(source, month, year)
    return _execute_scalar_query(conn, sql)


# ---------------------------------------------------------------------------
# Previous month helper
# ---------------------------------------------------------------------------

def _previous_month(month: str, year: int) -> tuple[str, int]:
    """Return the (month, year) tuple for the month before the given one.

    Args:
        month: Two-digit month string (e.g. '01', '12').
        year: Four-digit year.

    Returns:
        A tuple of (month_str, year_int) for the preceding month.
    """
    m = int(month)
    if m == 1:
        return "12", year - 1
    return f"{m - 1:02d}", year


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def fetch_bucket_actuals(
    month: str,
    year: int,
    config_path: str = "config.yaml",
) -> dict[str, dict[str, float]]:
    """Fetch current and previous month actuals for every configured bucket.

    Reads bucket definitions from *config_path*, queries Redshift for each
    bucket's sources, and aggregates monthly totals.

    Args:
        month: Two-digit month string for the reporting month (e.g. '01').
        year: Four-digit reporting year (e.g. 2026).
        config_path: Path to the YAML config file.

    Returns:
        A dict keyed by bucket name. Each value is a dict with keys:
            - ``current_month`` (float): Total cost for the given month.
            - ``previous_month`` (float): Total cost for the prior month.

    Raises:
        psycopg2.OperationalError: On Redshift connection failure.
        FileNotFoundError: If the config file is missing.

    Example::

        >>> actuals = fetch_bucket_actuals("01", 2026)
        >>> actuals["Total"]["current_month"]
        1234567.89
    """
    config = load_config(config_path)
    buckets: list[dict[str, Any]] = config.get("buckets", [])

    prev_month, prev_year = _previous_month(month, year)

    results: dict[str, dict[str, float]] = {}

    conn = get_shared_connection()
    for bucket in buckets:
        bucket_name: str = bucket["name"]
        sources: list[dict[str, Any]] = bucket.get("sources", [])

        current_total = 0.0
        previous_total = 0.0

        for source in sources:
            current_total += _fetch_source_cost(conn, source, month, year)
            previous_total += _fetch_source_cost(conn, source, prev_month, prev_year)

        results[bucket_name] = {
            "current_month": current_total,
            "previous_month": previous_total,
        }

        logger.info(
            "Bucket '%s': current=%.2f, previous=%.2f",
            bucket_name,
            current_total,
            previous_total,
        )

    return results

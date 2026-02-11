"""Redshift data ingestion module.

Handles all data loading from Redshift for the FinOps Report Generator.
Reads bucket definitions from config.yaml, builds SQL queries dynamically,
and returns aggregated monthly cost data per bucket.
"""

from __future__ import annotations

import calendar
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


# ---------------------------------------------------------------------------
# App-level COGS actuals (Phase 2)
# ---------------------------------------------------------------------------

def fetch_app_actuals(
    month: str,
    year: int,
    config_path: str = "config.yaml",
) -> dict[str, dict[str, float]]:
    """Fetch current and previous month COGS spend grouped by awn_app.

    Queries ``daily_cur_summary`` with the same common filters used for the
    COGS bucket, but groups results by ``awn_app`` instead of summing
    everything into a single total.

    Args:
        month: Two-digit month string for the reporting month (e.g. '01').
        year: Four-digit reporting year (e.g. 2026).
        config_path: Path to the YAML config file.

    Returns:
        A dict keyed by awn_app name. Each value is a dict with keys:
            - ``current_month`` (float): COGS spend for the given month.
            - ``previous_month`` (float): COGS spend for the prior month.
    """
    config = load_config(config_path)

    # Find the COGS bucket to reuse its source definition
    cogs_bucket = None
    for bucket in config.get("buckets", []):
        if bucket["name"] == "COGS":
            cogs_bucket = bucket
            break
    if cogs_bucket is None:
        raise ValueError("No 'COGS' bucket found in config.")

    source = cogs_bucket["sources"][0]
    table = source["table"]
    cost_column = source["cost_column"]
    date_column = source["date_column"]
    filters: list[str] = source.get("filters", [])

    prev_month, prev_year = _previous_month(month, year)
    two_months_ago_m, two_months_ago_y = _previous_month(prev_month, prev_year)

    conn = get_shared_connection()
    results: dict[str, dict[str, float]] = {}

    for m, y, period in [
        (month, year, "current_month"),
        (prev_month, prev_year, "previous_month"),
        (two_months_ago_m, two_months_ago_y, "two_months_ago"),
    ]:
        where_parts = [_build_month_filter(date_column, m, y)] + list(filters)
        where_clause = " AND ".join(where_parts)

        sql = (
            f"SELECT COALESCE(awn_app, 'other') AS app_name, "
            f"COALESCE(SUM({cost_column}), 0) AS total_cost "
            f"FROM {table} "
            f"WHERE {where_clause} "
            f"GROUP BY COALESCE(awn_app, 'other') "
            f"ORDER BY total_cost DESC"
        )
        logger.debug("App actuals SQL:\n%s", sql)

        with conn.cursor() as cur:
            cur.execute(sql)
            rows = cur.fetchall()

        for row in rows:
            app_name = str(row[0]) if row[0] else "other"
            cost_val = float(row[1]) if row[1] is not None else 0.0
            if app_name not in results:
                results[app_name] = {"current_month": 0.0, "previous_month": 0.0, "two_months_ago": 0.0}
            results[app_name][period] = cost_val

    logger.info("Fetched app-level actuals for %d apps", len(results))
    return results


# ---------------------------------------------------------------------------
# Historical data (6-month trend)
# ---------------------------------------------------------------------------

def _walk_back_months(
    month: str,
    year: int,
    count: int,
) -> list[tuple[str, int]]:
    """Return a chronological list of (month_str, year) tuples going back *count* months.

    The list starts with the oldest month and ends with the given month/year.

    Args:
        month: Two-digit month string for the most recent month.
        year: Four-digit year for the most recent month.
        count: How many months to include (including the given month).

    Returns:
        A list of (month_str, year) tuples in chronological order.
    """
    pairs: list[tuple[str, int]] = []
    m, y = int(month), year
    for _ in range(count):
        pairs.append((f"{m:02d}", y))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    pairs.reverse()
    return pairs


def _build_history_query(
    source: dict[str, Any],
    months: list[tuple[str, int]],
) -> str:
    """Build a single query that returns monthly aggregates for multiple months.

    Uses ``GROUP BY DATE_TRUNC('month', date_column)`` to batch all requested
    months into one round-trip per source.

    Args:
        source: A single source dict from config.yaml.
        months: List of (month_str, year) tuples to include.

    Returns:
        A complete SQL SELECT statement.
    """
    table = source["table"]
    cost_column = source["cost_column"]
    date_column = source["date_column"]
    filters: list[str] = source.get("filters", [])

    # Build month predicates: DATE_TRUNC('month', col) IN ('2026-01-01', ...)
    date_literals = ", ".join(
        f"'{y}-{m}-01'" for m, y in months
    )
    month_clause = f"DATE_TRUNC('month', {date_column}) IN ({date_literals})"

    where_parts = [month_clause] + list(filters)
    where_clause = " AND ".join(where_parts)

    return (
        f"SELECT DATE_TRUNC('month', {date_column}) AS month_start, "
        f"COALESCE(SUM({cost_column}), 0) AS total_cost "
        f"FROM {table} "
        f"WHERE {where_clause} "
        f"GROUP BY DATE_TRUNC('month', {date_column}) "
        f"ORDER BY month_start"
    )


def fetch_bucket_history(
    month: str,
    year: int,
    num_months: int = 6,
    config_path: str = "config.yaml",
) -> dict[str, list[dict[str, Any]]]:
    """Fetch *num_months* of historical monthly actuals for every configured bucket.

    Args:
        month: Two-digit month string for the most recent month.
        year: Four-digit year for the most recent month.
        num_months: Number of months to include (default 6).
        config_path: Path to the YAML config file.

    Returns:
        A dict keyed by bucket name. Each value is a chronological list of dicts
        with keys ``month_start`` (str, e.g. "Jan 2026") and ``actual`` (float).
    """
    config = load_config(config_path)
    buckets: list[dict[str, Any]] = config.get("buckets", [])
    months = _walk_back_months(month, year, num_months)

    # Build month_label lookup: "YYYY-MM-01" -> "Jan 2026"
    label_lookup: dict[str, str] = {}
    for m_str, y in months:
        m_int = int(m_str)
        abbr = calendar.month_abbr[m_int]
        label_lookup[f"{y}-{m_str}-01"] = f"{abbr} {y}"

    conn = get_shared_connection()
    results: dict[str, list[dict[str, Any]]] = {}

    for bucket in buckets:
        bucket_name: str = bucket["name"]
        sources: list[dict[str, Any]] = bucket.get("sources", [])

        # Accumulator keyed by date string
        monthly_totals: dict[str, float] = {
            f"{y}-{m}-01": 0.0 for m, y in months
        }

        for source in sources:
            sql = _build_history_query(source, months)
            logger.debug("History SQL:\n%s", sql)
            with conn.cursor() as cur:
                cur.execute(sql)
                rows = cur.fetchall()
            for row in rows:
                date_val = row[0]
                cost_val = float(row[1]) if row[1] is not None else 0.0
                # Normalise the date to "YYYY-MM-01" string
                if hasattr(date_val, "strftime"):
                    key = date_val.strftime("%Y-%m-%d")
                else:
                    key = str(date_val)[:10]
                if key in monthly_totals:
                    monthly_totals[key] += cost_val

        # Build chronological list with human-readable labels
        history: list[dict[str, Any]] = []
        for m_str, y in months:
            date_key = f"{y}-{m_str}-01"
            history.append({
                "month_start": label_lookup[date_key],
                "actual": monthly_totals[date_key],
            })

        results[bucket_name] = history
        logger.info(
            "Bucket '%s' history: %d months loaded",
            bucket_name,
            len(history),
        )

    return results

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

    two_months_ago_m, two_months_ago_y = _previous_month(prev_month, prev_year)

    results: dict[str, dict[str, float]] = {}

    conn = get_shared_connection()
    for bucket in buckets:
        bucket_name: str = bucket["name"]
        sources: list[dict[str, Any]] = bucket.get("sources", [])

        current_total = 0.0
        previous_total = 0.0
        two_months_ago_total = 0.0

        for source in sources:
            current_total += _fetch_source_cost(conn, source, month, year)
            previous_total += _fetch_source_cost(conn, source, prev_month, prev_year)
            two_months_ago_total += _fetch_source_cost(conn, source, two_months_ago_m, two_months_ago_y)

        results[bucket_name] = {
            "current_month": current_total,
            "previous_month": previous_total,
            "two_months_ago": two_months_ago_total,
        }

        logger.info(
            "Bucket '%s': current=%.2f, previous=%.2f, two_months_ago=%.2f",
            bucket_name,
            current_total,
            previous_total,
            two_months_ago_total,
        )

    return results


# ---------------------------------------------------------------------------
# Cylance actuals (cylance_cost_allocation)
# ---------------------------------------------------------------------------

def fetch_cylance_actuals(
    month: str,
    year: int,
) -> dict[str, dict[str, float]]:
    """Fetch current and previous month COGS/OpEx from cylance_cost_allocation.

    Uses the pre-computed allocation table rather than the raw CUR, since
    Cylance costs are split by percentage-based allocation rules.

    Args:
        month: Two-digit month string for the reporting month (e.g. '03').
        year: Four-digit reporting year.

    Returns:
        A dict with keys ``"COGS"`` and ``"OpEx"``, each containing
        ``current_month`` and ``previous_month`` float values.
    """
    prev_month, prev_year = _previous_month(month, year)

    curr_date = f"{year}-{month}-01"
    prev_date = f"{prev_year}-{prev_month}-01"

    sql = f"""
        SELECT
            cost_month,
            COALESCE(SUM(cogs_adjusted_cost), 0) AS cogs_total,
            COALESCE(SUM(opex_adjusted_cost), 0) AS opex_total
        FROM public.cylance_cost_allocation
        WHERE cost_month IN ('{curr_date}', '{prev_date}')
        GROUP BY cost_month
    """

    logger.debug("Cylance actuals SQL:\n%s", sql)
    conn = get_shared_connection()
    with conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    results: dict[str, dict[str, float]] = {
        "COGS": {"current_month": 0.0, "previous_month": 0.0},
        "OpEx": {"current_month": 0.0, "previous_month": 0.0},
    }

    for row in rows:
        date_val = row[0]
        month_key = date_val.strftime("%Y-%m-%d") if hasattr(date_val, "strftime") else str(date_val)[:10]
        cogs = float(row[1]) if row[1] is not None else 0.0
        opex = float(row[2]) if row[2] is not None else 0.0

        if month_key == curr_date:
            results["COGS"]["current_month"] = cogs
            results["OpEx"]["current_month"] = opex
        elif month_key == prev_date:
            results["COGS"]["previous_month"] = cogs
            results["OpEx"]["previous_month"] = opex

    logger.info(
        "Cylance actuals: COGS current=%.2f previous=%.2f | OpEx current=%.2f previous=%.2f",
        results["COGS"]["current_month"], results["COGS"]["previous_month"],
        results["OpEx"]["current_month"], results["OpEx"]["previous_month"],
    )
    return results


def fetch_cylance_dbx_summary(
    month: str,
    year: int,
) -> dict[str, dict[str, float]]:
    """Fetch current and previous month Databricks COGS/OpEx for Cylance.

    Uses the pre-computed ``public.cylance_dbx_cost_allocation`` view which
    has columns ``usage_month``, ``cost_center``, and ``contract_cost``.
    Any ``cost_center`` value other than ``'COGS'`` is treated as OpEx.

    Args:
        month: Two-digit month string for the reporting month (e.g. ``'03'``).
        year: Four-digit reporting year.

    Returns:
        A dict with keys ``"COGS"`` and ``"OpEx"``, each containing
        ``current_month`` and ``previous_month`` float values.
    """
    prev_month, prev_year = _previous_month(month, year)

    curr_date = f"{year}-{month}-01"
    prev_date = f"{prev_year}-{prev_month}-01"

    sql = f"""
        SELECT
            usage_month,
            CASE WHEN cost_center = 'COGS' THEN 'COGS' ELSE 'OpEx' END AS bucket,
            COALESCE(SUM(contract_cost), 0) AS total_cost
        FROM public.cylance_dbx_cost_allocation
        WHERE usage_month IN ('{curr_date}', '{prev_date}')
        GROUP BY 1, 2
    """

    logger.debug("Cylance DBX summary SQL:\n%s", sql)
    conn = get_shared_connection()
    with conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    results: dict[str, dict[str, float]] = {
        "COGS": {"current_month": 0.0, "previous_month": 0.0},
        "OpEx": {"current_month": 0.0, "previous_month": 0.0},
    }

    for row in rows:
        date_val = row[0]
        month_key = date_val.strftime("%Y-%m-%d") if hasattr(date_val, "strftime") else str(date_val)[:10]
        bucket = str(row[1])  # 'COGS' or 'OpEx'
        cost = float(row[2]) if row[2] is not None else 0.0

        if bucket not in results:
            bucket = "OpEx"

        if month_key == curr_date:
            results[bucket]["current_month"] += cost
        elif month_key == prev_date:
            results[bucket]["previous_month"] += cost

    logger.info(
        "Cylance DBX summary: COGS current=%.2f previous=%.2f | OpEx current=%.2f previous=%.2f",
        results["COGS"]["current_month"], results["COGS"]["previous_month"],
        results["OpEx"]["current_month"], results["OpEx"]["previous_month"],
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
# COGS drill-down (Phase 3)
# ---------------------------------------------------------------------------

def fetch_cogs_drilldown(
    month: str,
    year: int,
    config_path: str = "config.yaml",
) -> list[dict[str, Any]]:
    """Fetch top COGS MoM movers broken down by awn_app, product_name, operation.

    Uses a single-pass conditional SUM query for current + previous month,
    filters to rows with ABS(delta) > noise_floor, ranks by ABS(delta),
    and collapses the long tail into an "All Other" row.

    Args:
        month: Two-digit month string for the reporting month.
        year: Four-digit reporting year.
        config_path: Path to the YAML config file.

    Returns:
        A list of dicts sorted by ABS(delta_cost) DESC, each with:
            - ``awn_app``, ``product_name``, ``operation``
            - ``current_month``, ``previous_month``, ``delta_cost``
    """
    config = load_config(config_path)
    dd_cfg = config.get("drilldown", {})
    noise_floor = dd_cfg.get("noise_floor", 1000)
    top_n = dd_cfg.get("top_n", 15)

    source = dd_cfg.get("source", {})
    table = source.get("table", "public.daily_cur_summary")
    cost_column = source.get("cost_column", "cogs_adjusted_cost")
    date_column = source.get("date_column", "usage_date")
    filters: list[str] = source.get("filters", [])

    prev_month, prev_year = _previous_month(month, year)

    curr_filter = _build_month_filter(date_column, month, year)
    prev_filter = _build_month_filter(date_column, prev_month, prev_year)

    # Build the combined month filter
    combined_month = (
        f"({curr_filter} OR {prev_filter})"
    )
    where_parts = [combined_month] + list(filters)
    where_clause = " AND ".join(where_parts)

    sql = f"""
WITH grouped AS (
    SELECT
        COALESCE(awn_app, 'Untagged') AS awn_app,
        product_name,
        operation,
        SUM(CASE WHEN {curr_filter} THEN {cost_column} ELSE 0 END) AS current_month,
        SUM(CASE WHEN {prev_filter} THEN {cost_column} ELSE 0 END) AS previous_month
    FROM {table}
    WHERE {where_clause}
    GROUP BY COALESCE(awn_app, 'Untagged'), product_name, operation
),
filtered AS (
    SELECT *,
        current_month - previous_month AS delta_cost
    FROM grouped
    WHERE ABS(current_month - previous_month) > {noise_floor}
),
ranked AS (
    SELECT *,
        ROW_NUMBER() OVER (ORDER BY ABS(delta_cost) DESC) AS rn
    FROM filtered
),
combined AS (
    SELECT awn_app, product_name, operation, current_month, previous_month, delta_cost
    FROM ranked
    WHERE rn <= {top_n}

    UNION ALL

    SELECT
        'All Other' AS awn_app,
        '' AS product_name,
        '' AS operation,
        SUM(current_month) AS current_month,
        SUM(previous_month) AS previous_month,
        SUM(delta_cost) AS delta_cost
    FROM ranked
    WHERE rn > {top_n}
    HAVING COUNT(*) > 0
)
SELECT awn_app, product_name, operation, current_month, previous_month, delta_cost
FROM combined
ORDER BY ABS(delta_cost) DESC
"""

    logger.debug("COGS drilldown SQL:\n%s", sql)
    conn = get_shared_connection()
    with conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    results: list[dict[str, Any]] = []
    for row in rows:
        results.append({
            "awn_app": str(row[0]) if row[0] else "Untagged",
            "product_name": str(row[1]) if row[1] else "",
            "operation": str(row[2]) if row[2] else "",
            "current_month": float(row[3]) if row[3] is not None else 0.0,
            "previous_month": float(row[4]) if row[4] is not None else 0.0,
            "delta_cost": float(row[5]) if row[5] is not None else 0.0,
        })

    logger.info("Fetched %d COGS drilldown rows", len(results))
    return results


def fetch_ec2_purchase_breakdown(
    month: str,
    year: int,
    config_path: str = "config.yaml",
) -> list[dict[str, Any]]:
    """Fetch EC2 RunInstances MoM breakdown by purchase_option and region.

    Same CTE pattern as :func:`fetch_cogs_drilldown` but filtered to the
    configured EC2 product_name and operation, grouped by purchase_option
    and region.

    Args:
        month: Two-digit month string for the reporting month.
        year: Four-digit reporting year.
        config_path: Path to the YAML config file.

    Returns:
        A list of dicts sorted by ABS(delta_cost) DESC, each with:
            - ``purchase_option``, ``region``
            - ``current_month``, ``previous_month``, ``delta_cost``
    """
    config = load_config(config_path)
    dd_cfg = config.get("drilldown", {})
    ec2_product = dd_cfg.get("ec2_product_name", "Elastic Compute Cloud")
    ec2_operation = dd_cfg.get("ec2_operation", "RunInstances")
    top_n = dd_cfg.get("top_n", 15)
    ec2_noise_floor = 100  # Lower threshold for this filtered subset

    source = dd_cfg.get("source", {})
    table = source.get("table", "public.daily_cur_summary")
    cost_column = source.get("cost_column", "cogs_adjusted_cost")
    date_column = source.get("date_column", "usage_date")
    filters: list[str] = source.get("filters", [])

    prev_month, prev_year = _previous_month(month, year)

    curr_filter = _build_month_filter(date_column, month, year)
    prev_filter = _build_month_filter(date_column, prev_month, prev_year)

    combined_month = f"({curr_filter} OR {prev_filter})"
    where_parts = (
        [combined_month]
        + list(filters)
        + [f"product_name = '{ec2_product}'", f"operation = '{ec2_operation}'"]
    )
    where_clause = " AND ".join(where_parts)

    sql = f"""
WITH grouped AS (
    SELECT
        purchase_option,
        region,
        SUM(CASE WHEN {curr_filter} THEN {cost_column} ELSE 0 END) AS current_month,
        SUM(CASE WHEN {prev_filter} THEN {cost_column} ELSE 0 END) AS previous_month
    FROM {table}
    WHERE {where_clause}
    GROUP BY purchase_option, region
),
filtered AS (
    SELECT *,
        current_month - previous_month AS delta_cost
    FROM grouped
    WHERE ABS(current_month - previous_month) > {ec2_noise_floor}
),
ranked AS (
    SELECT *,
        ROW_NUMBER() OVER (ORDER BY ABS(delta_cost) DESC) AS rn
    FROM filtered
),
combined AS (
    SELECT purchase_option, region, current_month, previous_month, delta_cost
    FROM ranked
    WHERE rn <= {top_n}

    UNION ALL

    SELECT
        'All Other' AS purchase_option,
        '' AS region,
        SUM(current_month) AS current_month,
        SUM(previous_month) AS previous_month,
        SUM(delta_cost) AS delta_cost
    FROM ranked
    WHERE rn > {top_n}
    HAVING COUNT(*) > 0
)
SELECT purchase_option, region, current_month, previous_month, delta_cost
FROM combined
ORDER BY ABS(delta_cost) DESC
"""

    logger.debug("EC2 purchase breakdown SQL:\n%s", sql)
    conn = get_shared_connection()
    with conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    results: list[dict[str, Any]] = []
    for row in rows:
        results.append({
            "purchase_option": str(row[0]) if row[0] else "",
            "region": str(row[1]) if row[1] else "",
            "current_month": float(row[2]) if row[2] is not None else 0.0,
            "previous_month": float(row[3]) if row[3] is not None else 0.0,
            "delta_cost": float(row[4]) if row[4] is not None else 0.0,
        })

    logger.info("Fetched %d EC2 purchase breakdown rows", len(results))
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


def fetch_unit_cost_data(
    month: str,
    year: int,
    num_months: int = 5,
) -> dict[str, Any]:
    """Fetch unit cost data from public_bronze.v_pod_monthly_unit_costs.

    Queries the pre-built view for the last *num_months* months, returning
    both org-level aggregates and per-pod breakdowns for trend charts and
    MoM comparison tables.

    Args:
        month: Two-digit month string for the reporting month.
        year: Four-digit reporting year.
        num_months: Number of months of history.

    Returns:
        Dict with keys:
            - ``org_history``: chronological list of org-level monthly dicts
              (``month_label``, ``total_cogs``, ``total_analyzed_obs``,
              ``cogs_per_1m_analyzed``).
            - ``pod_history``: chronological list of per-pod monthly dicts
              (``pod``, ``month_label``, ``awn_cogs``,
              ``total_analyzed_obs``, ``cogs_per_1m_analyzed``).
            - ``pod_mom``: list of per-pod MoM comparison dicts for the
              reporting month vs the previous month.
    """
    from collections import defaultdict

    months = _walk_back_months(month, year, num_months)

    # Build month-label lookup: "YYYY-MM-01" -> "Jan 2026"
    label_lookup: dict[str, str] = {}
    for m_str, y in months:
        m_int = int(m_str)
        abbr = calendar.month_abbr[m_int]
        label_lookup[f"{y}-{m_str}-01"] = f"{abbr} {y}"

    date_literals = ", ".join(f"'{y}-{m}-01'" for m, y in months)

    sql = f"""
        SELECT month, pod, awn_cogs, total_analyzed_obs, cogs_per_1m_analyzed
        FROM public_bronze.v_pod_monthly_unit_costs
        WHERE DATE_TRUNC('month', month) IN ({date_literals})
        ORDER BY month, pod
    """

    logger.debug("Unit cost SQL:\n%s", sql)
    conn = get_shared_connection()
    with conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    # Parse into pod -> date_key -> data dict
    pod_month_data: dict[str, dict[str, dict[str, Any]]] = defaultdict(dict)

    for row in rows:
        date_val = row[0]
        pod = str(row[1]) if row[1] else "unknown"
        cogs = float(row[2]) if row[2] is not None else 0.0
        analyzed_obs = int(row[3]) if row[3] is not None else 0
        cogs_per_1m = float(row[4]) if row[4] is not None else 0.0

        date_key = date_val.strftime("%Y-%m-%d") if hasattr(date_val, "strftime") else str(date_val)[:10]

        pod_month_data[pod][date_key] = {
            "month_label": label_lookup.get(date_key, date_key),
            "awn_cogs": cogs,
            "total_analyzed_obs": analyzed_obs,
            "cogs_per_1m_analyzed": cogs_per_1m,
        }

    # Org-level aggregation per month
    org_history: list[dict[str, Any]] = []
    for m_str, y in months:
        date_key = f"{y}-{m_str}-01"
        total_cogs = 0.0
        total_obs = 0
        for pd in pod_month_data.values():
            if date_key in pd:
                total_cogs += pd[date_key]["awn_cogs"]
                total_obs += pd[date_key]["total_analyzed_obs"]
        cogs_per_1m = (total_cogs / (total_obs / 1_000_000)) if total_obs > 0 else 0.0
        org_history.append({
            "month_label": label_lookup[date_key],
            "total_cogs": total_cogs,
            "total_analyzed_obs": total_obs,
            "cogs_per_1m_analyzed": cogs_per_1m,
        })

    # Pod-level chronological history
    pod_history: list[dict[str, Any]] = []
    for m_str, y in months:
        date_key = f"{y}-{m_str}-01"
        for pod in sorted(pod_month_data.keys()):
            if date_key in pod_month_data[pod]:
                entry = pod_month_data[pod][date_key].copy()
                entry["pod"] = pod
                pod_history.append(entry)

    # Pod MoM comparison (current vs previous month)
    curr_key = f"{year}-{month}-01"
    prev_m, prev_y = _previous_month(month, year)
    prev_key = f"{prev_y}-{prev_m}-01"

    def _pct(new: float, old: float) -> float:
        return ((new - old) / old) * 100.0 if old else 0.0

    pod_mom: list[dict[str, Any]] = []
    for pod in sorted(pod_month_data.keys()):
        curr = pod_month_data[pod].get(curr_key, {})
        prev = pod_month_data[pod].get(prev_key, {})

        cc = curr.get("awn_cogs", 0.0)
        pc = prev.get("awn_cogs", 0.0)
        co = curr.get("total_analyzed_obs", 0)
        po = prev.get("total_analyzed_obs", 0)
        cu = curr.get("cogs_per_1m_analyzed", 0.0)
        pu = prev.get("cogs_per_1m_analyzed", 0.0)

        pod_mom.append({
            "pod": pod.upper(),
            "curr_cogs": cc,
            "cogs_pct": _pct(cc, pc),
            "curr_obs": co,
            "obs_pct": _pct(float(co), float(po)),
            "curr_unit_cost": cu,
            "unit_cost_pct": _pct(cu, pu),
        })

    # Add org-level totals row
    org_curr = next((h for h in org_history if h["month_label"] == label_lookup.get(curr_key)), None)
    org_prev = next((h for h in org_history if h["month_label"] == label_lookup.get(prev_key)), None)
    if org_curr and org_prev:
        pod_mom.append({
            "pod": "Total",
            "curr_cogs": org_curr["total_cogs"],
            "cogs_pct": _pct(org_curr["total_cogs"], org_prev["total_cogs"]),
            "curr_obs": org_curr["total_analyzed_obs"],
            "obs_pct": _pct(float(org_curr["total_analyzed_obs"]), float(org_prev["total_analyzed_obs"])),
            "curr_unit_cost": org_curr["cogs_per_1m_analyzed"],
            "unit_cost_pct": _pct(org_curr["cogs_per_1m_analyzed"], org_prev["cogs_per_1m_analyzed"]),
        })

    logger.info(
        "Fetched unit cost data: %d months, %d pods",
        num_months, len(pod_month_data),
    )

    return {
        "org_history": org_history,
        "pod_history": pod_history,
        "pod_mom": pod_mom,
    }


# ---------------------------------------------------------------------------
# AWN Databricks spend (dbx_cur)
# ---------------------------------------------------------------------------

def _build_dbx_cogs_cost_expr() -> str:
    """Return a SQL CASE expression that computes AWN Databricks COGS cost.

    Encodes the full COGS attribution logic from DATA.md:
    - Only workspace ``prod-observations-workspace`` after 2025-05-15
    - Excludes INTERACTIVE sessions, specific cluster IDs, experiment
      lakeflow jobs, and rows tagged ``tag_opex_in_prod = 'true'``
    - Query COGS (sku_group = 'AWS Serverless SQL Compute'): full contract_cost
    - Non-Query COGS: 15% of contract_cost before 2025-11-01, full after
    """
    return """
CASE
  WHEN usage_start_time < '2025-05-15' THEN 0
  WHEN workspace_name = 'prod-observations-workspace'
    AND (billing_origin_product IS NULL OR billing_origin_product != 'INTERACTIVE')
    AND (usage_cluster_id IS NULL
         OR usage_cluster_id NOT IN ('0603-203550-6nkhexq9', '0617-191106-u5z8uo9'))
    AND (lakeflow_job_name IS NULL OR (
        lakeflow_job_name NOT ILIKE '%experiment%'
        AND lakeflow_job_name NOT LIKE '%New Job%'
        AND lakeflow_job_name NOT LIKE '%optimize-2%'
        AND lakeflow_job_name NOT LIKE '%Refill_Large_By_Day%'
        AND lakeflow_job_name NOT LIKE '%Optimize_large_by_day%'
    ))
    AND (tag_opex_in_prod IS NULL OR tag_opex_in_prod != 'true')
  THEN
    CASE
      WHEN sku_group = 'AWS Serverless SQL Compute' THEN contract_cost
      WHEN usage_start_time < '2025-11-01' THEN contract_cost * 0.15
      ELSE contract_cost
    END
  ELSE 0
END
""".strip()


def fetch_dbx_awn_summary(
    month: str,
    year: int,
) -> dict[str, dict[str, float]]:
    """Fetch current and previous month Databricks COGS/OpEx for Arctic Wolf.

    Uses the COGS attribution logic to split ``contract_cost`` into COGS and
    OpEx (total minus COGS) from ``public.dbx_cur``.

    Args:
        month: Two-digit month string for the reporting month.
        year: Four-digit reporting year.

    Returns:
        A dict with keys ``"COGS"`` and ``"OpEx"``, each containing
        ``current_month`` and ``previous_month`` float values.
    """
    prev_month, prev_year = _previous_month(month, year)
    cogs_expr = _build_dbx_cogs_cost_expr()

    curr_date = f"{year}-{month}-01"
    prev_date = f"{prev_year}-{prev_month}-01"

    sql = f"""
        SELECT
            DATE_TRUNC('month', usage_start_time) AS cost_month,
            COALESCE(SUM({cogs_expr}), 0) AS cogs_total,
            COALESCE(SUM(contract_cost), 0) - COALESCE(SUM({cogs_expr}), 0) AS opex_total
        FROM public.dbx_cur
        WHERE organization = 'Arctic Wolf'
          AND DATE_TRUNC('month', usage_start_time) IN ('{curr_date}', '{prev_date}')
        GROUP BY 1
    """

    logger.debug("DBX AWN summary SQL:\n%s", sql)
    conn = get_shared_connection()
    with conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    results: dict[str, dict[str, float]] = {
        "COGS": {"current_month": 0.0, "previous_month": 0.0},
        "OpEx": {"current_month": 0.0, "previous_month": 0.0},
    }

    for row in rows:
        date_val = row[0]
        month_key = date_val.strftime("%Y-%m-%d") if hasattr(date_val, "strftime") else str(date_val)[:10]
        cogs = float(row[1]) if row[1] is not None else 0.0
        opex = float(row[2]) if row[2] is not None else 0.0

        if month_key == curr_date:
            results["COGS"]["current_month"] = cogs
            results["OpEx"]["current_month"] = opex
        elif month_key == prev_date:
            results["COGS"]["previous_month"] = cogs
            results["OpEx"]["previous_month"] = opex

    logger.info(
        "DBX AWN summary: COGS current=%.2f previous=%.2f | OpEx current=%.2f previous=%.2f",
        results["COGS"]["current_month"], results["COGS"]["previous_month"],
        results["OpEx"]["current_month"], results["OpEx"]["previous_month"],
    )
    return results


def fetch_dbx_awn_breakdown(
    month: str,
    year: int,
    dimension: str,
    cost_type: str,
    top_n: int = 15,
) -> list[dict[str, Any]]:
    """Fetch AWN Databricks spend breakdown by a given dimension.

    Args:
        month: Two-digit month string for the reporting month.
        year: Four-digit reporting year.
        dimension: Column to group by (``"workspace_name"`` or ``"sku_name"``).
        cost_type: ``"cogs"`` or ``"opex"``.
        top_n: Maximum number of rows before collapsing into "All Other".

    Returns:
        A list of dicts sorted by current_month DESC, each with:
            - ``dimension_value``, ``current_month``, ``previous_month``,
              ``delta_cost``
    """
    prev_month, prev_year = _previous_month(month, year)
    cogs_expr = _build_dbx_cogs_cost_expr()

    if cost_type == "cogs":
        cost_expr = cogs_expr
    else:
        cost_expr = f"(contract_cost - ({cogs_expr}))"

    curr_filter = _build_month_filter("usage_start_time", month, year)
    prev_filter = _build_month_filter("usage_start_time", prev_month, prev_year)

    sql = f"""
WITH grouped AS (
    SELECT
        COALESCE({dimension}, 'Unknown') AS dimension_value,
        SUM(CASE WHEN {curr_filter} THEN {cost_expr} ELSE 0 END) AS current_month,
        SUM(CASE WHEN {prev_filter} THEN {cost_expr} ELSE 0 END) AS previous_month
    FROM public.dbx_cur
    WHERE organization = 'Arctic Wolf'
      AND ({curr_filter} OR {prev_filter})
    GROUP BY 1
),
with_delta AS (
    SELECT *,
        current_month - previous_month AS delta_cost
    FROM grouped
    WHERE current_month > 0 OR previous_month > 0
),
ranked AS (
    SELECT *,
        ROW_NUMBER() OVER (ORDER BY current_month DESC) AS rn
    FROM with_delta
),
combined AS (
    SELECT dimension_value, current_month, previous_month, delta_cost
    FROM ranked WHERE rn <= {top_n}

    UNION ALL

    SELECT
        'All Other' AS dimension_value,
        SUM(current_month) AS current_month,
        SUM(previous_month) AS previous_month,
        SUM(delta_cost) AS delta_cost
    FROM ranked WHERE rn > {top_n}
    HAVING COUNT(*) > 0
)
SELECT dimension_value, current_month, previous_month, delta_cost
FROM combined
ORDER BY current_month DESC
"""

    logger.debug("DBX AWN breakdown (%s, %s) SQL:\n%s", dimension, cost_type, sql)
    conn = get_shared_connection()
    with conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    results: list[dict[str, Any]] = []
    for row in rows:
        results.append({
            "dimension_value": str(row[0]) if row[0] else "Unknown",
            "current_month": float(row[1]) if row[1] is not None else 0.0,
            "previous_month": float(row[2]) if row[2] is not None else 0.0,
            "delta_cost": float(row[3]) if row[3] is not None else 0.0,
        })

    logger.info("Fetched %d DBX AWN breakdown rows (%s, %s)", len(results), dimension, cost_type)
    return results


def fetch_app_cogs_history(
    month: str,
    year: int,
    num_months: int = 6,
    config_path: str = "config.yaml",
) -> list[dict[str, Any]]:
    """Fetch *num_months* of historical COGS spend grouped by awn_app.

    Queries the COGS source table with the standard COGS filters, grouping
    by ``COALESCE(awn_app, 'Untagged')`` and ``DATE_TRUNC('month', ...)``.

    Args:
        month: Two-digit month string for the most recent month.
        year: Four-digit year for the most recent month.
        num_months: Number of months to include (default 6).
        config_path: Path to the YAML config file.

    Returns:
        A flat chronological list of dicts with keys:
            - ``app_name`` (str): The awn_app value or ``'Untagged'``.
            - ``month_label`` (str): Human-readable label (e.g. ``'Jan 2026'``).
            - ``monthly_cost`` (float): Total COGS cost for this app/month.

    Raises:
        ValueError: If no ``COGS`` bucket is found in the config.
        psycopg2.OperationalError: On Redshift connection failure.
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

    months = _walk_back_months(month, year, num_months)

    # Build month_label lookup: "YYYY-MM-01" -> "Jan 2026"
    label_lookup: dict[str, str] = {}
    for m_str, y in months:
        m_int = int(m_str)
        abbr = calendar.month_abbr[m_int]
        label_lookup[f"{y}-{m_str}-01"] = f"{abbr} {y}"

    # Build month predicates: DATE_TRUNC('month', col) IN ('2026-01-01', ...)
    date_literals = ", ".join(
        f"'{y}-{m}-01'" for m, y in months
    )
    month_clause = f"DATE_TRUNC('month', {date_column}) IN ({date_literals})"

    where_parts = [month_clause] + list(filters)
    where_clause = " AND ".join(where_parts)

    sql = (
        f"SELECT COALESCE(awn_app, 'Untagged') AS app_name, "
        f"DATE_TRUNC('month', {date_column}) AS month_start, "
        f"COALESCE(SUM({cost_column}), 0) AS monthly_cost "
        f"FROM {table} "
        f"WHERE {where_clause} "
        f"GROUP BY 1, 2 "
        f"ORDER BY app_name, month_start"
    )

    logger.debug("App COGS history SQL:\n%s", sql)
    conn = get_shared_connection()
    with conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    results: list[dict[str, Any]] = []
    for row in rows:
        app_name = str(row[0]) if row[0] else "Untagged"
        date_val = row[1]
        cost_val = float(row[2]) if row[2] is not None else 0.0

        # Normalise the date to "YYYY-MM-01" string
        if hasattr(date_val, "strftime"):
            date_key = date_val.strftime("%Y-%m-%d")
        else:
            date_key = str(date_val)[:10]

        month_label = label_lookup.get(date_key, date_key)

        results.append({
            "app_name": app_name,
            "month_label": month_label,
            "monthly_cost": cost_val,
        })

    logger.info("Fetched %d app COGS history rows", len(results))
    return results

"""MoM and Forecast Variance calculation engine for the FinOps Report Generator.

Provides functions to compute Month-over-Month deltas and Forecast Variance
metrics for each cost bucket defined in config.yaml.
"""


def calculate_metrics(
    bucket_name: str,
    current_month_actual: float,
    previous_month_actual: float,
    forecast: float | None = None,
) -> dict:
    """Calculate MoM and optional Forecast Variance metrics for a single bucket.

    Args:
        bucket_name: Display name for the cost bucket (e.g. "Total", "COGS").
        current_month_actual: Actual spend for the current reporting month.
        previous_month_actual: Actual spend for the previous month.
        forecast: Forecasted spend for the current month. None when no
            forecast baseline is defined for this bucket.

    Returns:
        A dict containing the metric name, actuals, MoM delta/percentage/
        direction, and -- when a forecast is provided -- variance
        delta/percentage/direction.
    """
    mom_delta: float = current_month_actual - previous_month_actual
    mom_pct: float = (
        (mom_delta / previous_month_actual * 100.0)
        if previous_month_actual != 0
        else 0.0
    )

    if mom_delta > 0:
        mom_dir = "increase"
    elif mom_delta < 0:
        mom_dir = "decrease"
    else:
        mom_dir = "no change"

    result: dict = {
        "metric": bucket_name,
        "actual": current_month_actual,
        "previous_month": previous_month_actual,
        "mom_delta": mom_delta,
        "mom_pct": mom_pct,
        "mom_dir": mom_dir,
        "has_forecast": forecast is not None,
    }

    if forecast is not None:
        var_delta: float = current_month_actual - forecast
        var_pct: float = (
            (var_delta / forecast * 100.0) if forecast != 0 else 0.0
        )

        if var_delta > 0:
            var_dir = "over"
        elif var_delta < 0:
            var_dir = "under"
        else:
            var_dir = "on target"

        result["forecast"] = forecast
        result["var_delta"] = var_delta
        result["var_pct"] = var_pct
        result["var_dir"] = var_dir

    return result


def calculate_all_buckets(
    actuals: dict[str, dict],
    forecasts: dict[str, float | None],
    bucket_configs: list[dict],
) -> list[dict]:
    """Calculate metrics for every bucket defined in the configuration.

    Args:
        actuals: Keyed by bucket name. Each value is a dict with
            ``current_month`` (float) and ``previous_month`` (float).
        forecasts: Keyed by the forecast_mapping_key defined in config.yaml.
            Values are the forecasted dollar amount, or None when no forecast
            exists for that key.
        bucket_configs: The ``buckets`` list from config.yaml. Each entry must
            contain ``name`` (str) and ``forecast_mapping_key`` (str).

    Returns:
        A list of metric dicts (one per bucket), each produced by
        :func:`calculate_metrics`.
    """
    results: list[dict] = []

    for bucket in bucket_configs:
        name: str = bucket["name"]
        forecast_key: str = bucket["forecast_mapping_key"]

        bucket_actuals = actuals[name]
        forecast_value: float | None = forecasts.get(forecast_key)

        metrics = calculate_metrics(
            bucket_name=name,
            current_month_actual=bucket_actuals["current_month"],
            previous_month_actual=bucket_actuals["previous_month"],
            forecast=forecast_value,
        )
        results.append(metrics)

    return results

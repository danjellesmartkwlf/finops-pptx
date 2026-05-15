CREATE OR REPLACE VIEW public_bronze.v_pod_monthly_unit_costs AS
SELECT
    c.month,
    c.pod,
    c.awn_cogs,
    o.total_analyzed_obs,
    CASE
        WHEN o.total_analyzed_obs > 0
        THEN (c.awn_cogs / o.total_analyzed_obs) * 1000000.0
        ELSE 0
    END AS cogs_per_1m_analyzed
FROM (
    SELECT
        DATE_TRUNC('month', usage_date) AS month,
        pod,
        SUM(cogs_adjusted_cost) AS awn_cogs
    FROM public.daily_cur_summary
    WHERE pod != 'Other'
    GROUP BY DATE_TRUNC('month', usage_date), pod
) c
JOIN (
    SELECT
        DATE_TRUNC('month', event_date) AS month,
        pod,
        SUM(analyzed_observations) AS total_analyzed_obs
    FROM public_bronze.observations_with_customer_info_v2
    GROUP BY DATE_TRUNC('month', event_date), pod
) o ON c.month = o.month AND c.pod = o.pod
WITH NO SCHEMA BINDING;

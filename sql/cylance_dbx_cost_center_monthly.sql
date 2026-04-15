CREATE OR REPLACE VIEW cylance_dbx_cost_center_monthly AS
SELECT
    DATE_TRUNC('month', usage_start_time)::DATE AS usage_date
    , CASE
        WHEN workspace_id IN (
            '60890024704537531'
            , '14719583199753553'
            , '20712006485233323'
            , '33511330847169474'
            , '14719583199753553'
            , '40590727749956612'
            , '44798595593087930'
            , '77085780843432162'
        ) THEN 'OPEX'
        WHEN workspace_id IN (
            '66708573951389864'
            , '23404685876768229'
            , '47823637885058062'
            , '76660026459584182'
            , '65591468776087962'
            , '29565807635441610'
        ) THEN 'COGS'
        ELSE 'UNKNOWN'
    END AS cost_center
    , SUM(contract_cost) AS total_cost
FROM dbx_cur
WHERE pod = 'Cylance'
GROUP BY 1, 2;

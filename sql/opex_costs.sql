SELECT 
    CASE
        WHEN product_name IN (
                'Config','Security Hub','Inspector','GuardDuty','CloudTrail',
                'Backup','Cost Explorer','CloudWatch Events','CloudWatch','Systems Manager'
            )
        
        THEN 'AWS ' || product_name::text

        WHEN product_name LIKE '%Bedrock%'
        THEN 'AWS Bedrock'



        WHEN product_name = 'Support (Enterprise)'
        THEN 'AWS Enterprise Support'

        -- WHEN tag_eks_cluster_name = 'infra-elastic-stream' THEN tag_eks_cluster_name

        WHEN account_name ILIKE '%dev%'
        THEN 'Dev AWS Accounts'

        WHEN account_name ILIKE '%test%'
        THEN 'Test AWS Accounts'

        WHEN account_name LIKE '%AFT%'
        THEN 'AFT Accounts'
        WHEN account_name ILIKE '%payer%'
        THEN 'Payer Account'
        WHEN account_name LIKE '%Audit%'
        THEN 'Audit Accounts'
        WHEN account_name LIKE '%artifacts%'
        THEN 'AWN Artifacts'
        WHEN account_name ILIKE '%global-product-intelligence%'
        THEN 'FinOps Account'

        WHEN account_name IN (
            'awn-prod-vxintel',
            'awn-prod-tip',
            'awn-prod-thras',
            'awn-prod-shared-lab-services'
        )
        THEN 'Labs AWS Accounts'

        WHEN account_cost_center = 'cogs'
            AND tag_opex_in_prod = 'true'
        THEN 'OpexInProd Other'

        WHEN account_cost_center = 'cogs'
            AND tag_opex_in_prod != 'true'
        THEN 'COGS Other'

        ELSE 'Other'

    END AS "awn_app",
    
    date_trunc('month', usage_date) AS "Date",
    ROUND(SUM(opex_adjusted_cost),2) AS "OPEX (SUM)"
FROM {table}
WHERE usage_date >= '{start_date}'
    AND usage_date < '{end_date}'
    {exclusion_filters}
    GROUP BY 1, 2
    ORDER BY 3 DESC;
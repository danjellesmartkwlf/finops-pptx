# Data 

## Redshift

Redshift contains many tables, but we will use a few for this project. All are in the `public` database.

You can use the `redshift` mcp server as needed to explore data, inventory tables, etc.


There are 4 Cost and Usage Reports we are concerned with. These are their table names and if needed, filters

- Arctic Wolf AWS CUR
  - public.daily_cur_summary
- Arctic Wolf Databricks CUR
  - public.dbx_cur
  - Filter: Organization = "Arctic Wolf"
- Cylance AWS CUR
  - public.cylance_daily_cur_summary
- Cylance Databricks CUR
  - Filter: Organization = "Cylance"


Always Opex Globally
The following services are always considered OPEX, regardless of which account the spend is in:

```
product_name IN 'Config','Security Hub','Inspector','GuardDuty','CloudTrail','Backup','Cost Explorer','CloudWatch Events','CloudWatch','Systems Manager'
```


## public.daily_cur_summary

This table has AWS daily cur data for Arctic Wolf

Notable fields that are unique or special

- cogs_adjusted_cost - COGS spend
- opex_adjusted_cost - Opex spend
- net_amortized_cost - Total cost (Net Amortized)
- usage_date - Date field
- awn_app - Arctic Wolf Application
  - Only applicable to COGS
- financial_cost_category 
  - Only applicable to OPEX
- account_name - AWS Account Name
- account_cost_center - Default cost allocation for spend in the account
- tag_opex_in_prod - Tagged resources that run in a prod/cogs account but are attributed to opex

Tags:
- tag_domain
- tag_system
- tag_component
- tag_eks_cluster_name
- tag_cell
- tag_owner

Common Filters
- charge_type != 'Credit
- service does not contain 'Databrick%'
- service does not contain 'Tackle%


## public.cylance_daily_cur_summary

This table has AWS daily cur data for Cylance


## public.dbx_cur

This is for Databricks spend for both Arctic Wolf and Cylance
The organization is delineated with the `organization` field

Arctic Wolf - Cogs - Query

```
ifelse(
    {usage_start_time} < '2025-05-15',
    0,

    ifelse(
        organization = 'Arctic Wolf'
        AND {workspace_name} = 'prod-observations-workspace'
        AND {usage_start_time} >= '2025-05-15'

        // new logic
        // AND NOT in({usage_cluster_id}, ['0603-203550-6nkhexq9', '0617-191106-u5z8uo9'])
        // AND {billing_origin_product} <> 'INTERACTIVE'
        // end new logic
        AND (
            isNull({usage_cluster_id})
            OR {usage_cluster_id} <> '0603-203550-6nkhexq9' 
            OR {usage_cluster_id} <> '0617-191106-u5z8uo9'
           
        )
        AND (
            isNull({billing_origin_product})
             OR {billing_origin_product} <> 'INTERACTIVE'
        )
        
        AND {sku_group} = 'AWS Serverless SQL Compute'
        AND (
                isNull({lakeflow_job_name})
                OR (
                NOT contains({lakeflow_job_name}, 'experiment')
                AND NOT contains({lakeflow_job_name}, 'Experiment')
                AND NOT contains({lakeflow_job_name}, 'New Job')
                AND NOT contains({lakeflow_job_name}, 'optimize-2')
                AND NOT contains({lakeflow_job_name}, 'Refill_Large_By_Day')
                AND NOT contains({lakeflow_job_name}, 'Optimize_large_by_day')
                )
        )
        AND (
            isNull({tag_opex_in_prod})
            OR (
                {tag_opex_in_prod} <> 'true'
            )
        
        ),
        {Cost Invoice},
        0
    )
)
```


Arctic Wolf - Cogs - Non Query
```
ifelse(
    {usage_start_time} < '2025-05-15',
    0,

    ifelse(
        organization = 'Arctic Wolf'
        AND {workspace_name} = 'prod-observations-workspace'
        AND {usage_start_time} >= '2025-05-15'
        AND {sku_group} <> 'AWS Serverless SQL Compute'

        // new logic
        AND NOT in({usage_cluster_id}, ['0603-203550-6nkhexq9', '0617-191106-u5z8uo9'])
        AND {billing_origin_product} <> 'INTERACTIVE'
        // end new logic

        AND (
                isNull({lakeflow_job_name})
                OR (
                NOT contains({lakeflow_job_name}, 'experiment')
                AND NOT contains({lakeflow_job_name}, 'Experiment')
                AND NOT contains({lakeflow_job_name}, 'New Job')
                AND NOT contains({lakeflow_job_name}, 'optimize-2')
                AND NOT contains({lakeflow_job_name}, 'Refill_Large_By_Day')
                AND NOT contains({lakeflow_job_name}, 'Optimize_large_by_day')
                )
        )
        AND (
            isNull({tag_opex_in_prod})
            OR (
                {tag_opex_in_prod} <> 'true'
            )
        
        ),

        ifelse(
            {usage_start_time} < '2025-05-15', 0,
            {usage_start_time} >= '2025-05-15' AND {usage_start_time} < '2025-11-01',{Cost Invoice}*.15,
            {Cost Invoice}
        ),
        0
    )
)
```

Arctic Wolf COGS total = Query + NonQuery COGS
The remaining Arctic Wolf Databricks costs are OPEX

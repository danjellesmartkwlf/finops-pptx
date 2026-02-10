# PRD: FinOps Report Generator (Phase 1)

**Version:** 1.0
**Status:** Draft
**Owner:** FinOps Engineering
**Date:** February 10, 2026

## 1. Executive Summary

The goal of Phase 1 is to automate the extraction and formatting of the monthly executive cloud cost presentation. We will replace manual data gathering with a Python-based application that:

1. Ingests actuals from Redshift and forecasts from Excel.
2. Calculates Month-over-Month (MoM) and Forecast Variance for three key buckets (Total, COGS, OpEx).
3. Generates a narrative using a strict, pre-approved text template.
4. Outputs a standard PowerPoint deck.

**Out of Scope for Phase 1:**

* Automated root cause analysis (The "Variance Hunter").
* Complex Savings Plan/RI coverage analysis logic.
* Direct Jira integration for context injection (manual entry for now).

---

## 2. User Stories

### 2.1 The Setup

* **As a FinOps Engineer**, I want to define my reporting hierarchy (Total, COGS, OpEx) in a YAML configuration file so that I can change definitions without rewriting code.
* **As a FinOps Engineer**, I want to map specific rows in my Excel forecast to these buckets so that comparisons are accurate.

### 2.2 The Execution

* **As a User**, I want to run a command (or click a button) to generate the report for "January 2026."
* **As a User**, I want the system to calculate:
* Actual vs. Previous Month (MoM $ and %).
* Actual vs. Forecast (Variance $ and %).


* **As a User**, I want to handle metrics that *have no forecast* gracefully (i.e., hide the forecast variance section in the text).

### 2.3 The Review ("The Cockpit")

* **As a User**, I want a UI (Streamlit) where I can see the generated narrative text side-by-side with the charts.
* **As a User**, I want the ability to *edit* the generated text to add context (e.g., "This spike was due to the Load Test") before final export.
* **As a User**, I want to use an **MCP-powered "Ask AI" widget** to query Redshift ad-hoc if I need to explain a number during my review.

### 2.4 The Output

* **As an Executive**, I want a PowerPoint deck that follows the corporate template exactly (fonts, colors, logos).
* **As an Executive**, I want the text description to follow a consistent standard format every month so I can scan it quickly.

---

## 3. Functional Requirements

### 3.1 Configuration (`config.yaml`)

The system must be driven by a configuration file containing:

* **Database Connection:** Credentials will reside in .env. An .env.example file will be committed to the repo 
* **Buckets Definition:** List of metrics to track (Total, COGS, OpEx).
* *Attributes:* Name, SQL Logic (or Tag Filter), Forecast Mapping Key.


* **Templates:** The specific Python f-strings for the narrative generation.

### 3.2 Data Ingestion Layer

* **Redshift:** Must query the CUR (Cost & Usage Report) tables.
* *Aggregation:* Sum cost by `Month` and `CostCategory` (or defined Tags).


* **Excel:** Must ingest the Forecast file.
* *Validation:* Fail if the specified "Forecast Key" (e.g., Row 12) is missing.



### 3.3 The Logic Engine (Narrative Builder)

The system must support two template modes based on data availability:

**Mode A: Standard Variance (Forecast Exists)**

> *Template:* "{Metric} spend for {Month} was ${Actual}. This is a {MoM_Dir} of {MoM_Pct}% from last month. Against the forecast of ${Forecast}, we are {Var_Dir} by {Var_Pct}%."

**Mode B: MoM Only (No Forecast)**

> *Template:* "{Metric} spend for {Month} was ${Actual}. This is a {MoM_Dir} of {MoM_Pct}% from last month. (No forecast baseline defined)."

### 3.4 MCP Integration (The "Sidecar")

* The Streamlit UI must include an "Investigate" component.
* It must allow natural language queries (e.g., "What service drove the COGS increase?").
* It must route these queries via the MCP Server to Redshift.
* **Constraint:** The MCP server *read-only* access to Redshift. It cannot modify data.

### 3.5 PowerPoint Generation

* **Input:** A `.pptx` master template.
* **Mapping:** The config must map internal metric names to specific **Slide IDs** and **Shape Placeholder IDs** in the template.
* **Output:** A downloadable `.pptx` file.

---

## 4. Technical Architecture

### 4.1 Tech Stack

* **Language:** Python 3.10+
* **UI Framework:** Streamlit
* **Database:** `psycopg2` or `redshift_connector`
* **Presentation:** `python-pptx`
* **AI/Agent:** MCP Client (Model Context Protocol) connecting to internal LLM.

### 4.2 Directory Structure

```text
/finops-reporter
  ├── config.yaml            # The Source of Truth
  ├── app.py                 # Streamlit UI Entry point
  ├── src/
  │   ├── ingestion.py       # Redshift & Excel logic
  │   ├── narrative.py       # Template filling logic
  │   ├── pptx_gen.py        # Slide builder
  │   └── mcp_client.py      # AI Sidecar logic
  └── assets/
      └── template.pptx      # Corporate Master Deck

```

---

## 5. Success Metrics (Phase 1)

* **Accuracy:** The "Total Spend" number in the generated deck matches the Redshift query result exactly (to the penny).
* **Speed:** Generating the draft report takes < 2 minutes (vs. hours manually).
* **Consistency:** 100% of generated slides follow the defined text template format.
* **Adoption:** The FinOps team uses the "Cockpit" to generate the February month-end report.

---

## 6. Milestones

| Milestone | Deliverable | Target Date |
| --- | --- | --- |
| **M1: Foundation** | `config.yaml` defined, Redshift/Excel connection working. | Week 1 |
| **M2: Logic & Text** | Narrative engine generating correct "Mad Libs" text. | Week 2 |
| **M3: The Cockpit** | Streamlit UI working with charts + text editor + MCP sidecar. | Week 3 |
| **M4: The Artifact** | PowerPoint export working with corporate template. | Week 4 |

---

## 7. Open Questions / Risks

* **Excel Stability:** If the forecast Excel format changes (e.g., someone inserts a row), the mapping will break. *Mitigation: Add strict schema validation in the ingestion step.*
* **Tag Consistency:** Does "COGS" rely on a tag that might be missing for some resources? *Mitigation: Add a "Unallocated" bucket to the report to catch untagged spend.*
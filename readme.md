# Health MLR 2024 Simplified Interpretation Project

## Overview

This project is a simplified public-data analysis of 2024 Medical Loss Ratio (MLR) results using CMS Public Use File data.

The goal is to turn a difficult CMS filing-template export into a readable issuer-state-market table, then summarize MLR and rebate exposure across the core commercial medical markets:

- Individual
- Small Group
- Large Group

This project is intentionally scoped as a public-data interpretation project. It is not a pricing model, rate filing model, IBNR model, claim-level model, or official CMS MLR calculator.

## Project question

For 2024, how did CMS-reported MLR experience differ across the Individual, Small Group, and Large Group commercial medical markets, and where did rebate exposure appear?

## Data sources

Primary references:

- CMS Medical Loss Ratio overview: https://www.cms.gov/marketplace/private-health-insurance/medical-loss-ratio
- CMS Medical Loss Ratio Data and System Resources: https://www.cms.gov/marketplace/resources/data/medical-loss-ratio-data-systems-resources
- eCFR 45 CFR Part 158: https://www.ecfr.gov/current/title-45/subtitle-A/subchapter-B/part-158
- Microsoft Power Query overview: https://learn.microsoft.com/en-us/power-query/power-query-what-is-power-query
- Microsoft Power Query merge queries overview: https://learn.microsoft.com/en-us/power-query/merge-queries-overview
- Microsoft Power Query unpivot columns: https://learn.microsoft.com/en-us/power-query/unpivot-column
- Microsoft Power Query pivot columns: https://learn.microsoft.com/en-us/power-query/pivot-columns

Main CMS files used from the 2024 MLR Public Use File:

- `Part3_MLR_Rebate_Calculation.csv`
- `MR_Submission_Template_Header.csv`

## Repository structure

```text
.
├── README.md
├── workbook/
│   └── Simplified Health MLR Interpretation Study - 2024.xlsx
├── data/
│   └── raw CMS files are not redistributed here
└── output/
    └── summary tables, if exported later
```

Note: The workbook is the primary project artifact. The raw CMS files can be downloaded from CMS directly.

## What MLR means

At a high level, MLR compares medical claims and qualifying quality-improvement expenses to adjusted premium revenue.

Simplified memory line:

```text
MLR ≈ medical and quality spending / adjusted premium revenue
```

The official CMS calculation is more specific than simple claims divided by gross premium. CMS reports MLR numerator and denominator fields in the public-use file. This project uses those CMS-reported fields instead of recreating the full statutory formula from scratch.

## Unit of analysis

The cleaned unit of analysis is an issuer-state-market-year block.

One block means:

```text
one issuer + one state + one market + one reporting year
```

Example:

```text
Celtic Insurance Company + Texas + Individual + 2024
```

A block is not a ZIP code, claim, member, policy, or plan.

## Scope

Included markets:

- Individual
- Small Group
- Large Group

Included CMS market columns:

- `CMM_INDIVIDUAL_TOTAL`
- `CMM_SMALL_GROUP_TOTAL`
- `CMM_LARGE_GROUP_TOTAL`

Excluded from this simplified version:

- student health
- mini-med
- expatriate
- non-core or out-of-scope market columns
- member-level or claim-level detail

The simplified scope keeps the project focused on core commercial comprehensive major medical markets.

## Why the CMS file needs transformation

The CMS Part3 file is not structured like a normal analysis table.

A normal analysis table would look like this:

```text
one row = issuer + state + market + year
columns = denominator, numerator, standard, MLR, rebate
```

The CMS Part3 file is closer to a filing-template export:

```text
one row = submission ID + calculation line item
columns = market buckets
```

The key interpretation is:

```text
ROW_LOOKUP_CODE = metric or line item
CMM market column = market bucket
cell value = reported value
```

Example:

```text
ROW_LOOKUP_CODE = MLR_STANDARD
CMM_INDIVIDUAL_TOTAL = 0.800
```

This means the MLR standard for that submission ID's Individual market block is 80%.

## Fields selected

The simplified project uses only the fields needed to interpret MLR and rebate exposure:

- `MLR_DENOMINATOR`
- `MLR_NUMERATOR`
- `MLR_STANDARD`
- `REBATE_AMT_LIMITED`
- `REBATE_AMT_CREDIBILITY_ADJ_MLR`

For display and ranking, the workbook uses a selected rebate amount:

```text
selected rebate amount = REBATE_AMT_LIMITED if available,
otherwise REBATE_AMT_CREDIBILITY_ADJ_MLR
```

## Cleaning and transformation workflow

The workflow is designed to be reproducible in Excel without manual copy/paste.

Recommended tool: Excel Power Query.

Power Query is used for extract-transform-load work: importing CSVs, filtering rows, reshaping columns, merging tables, and loading the cleaned result back to Excel.

### Step 1: Import the CSV files

Use Excel:

```text
Data > Get Data > From Text/CSV
```

Import both source files:

```text
MR_Submission_Template_Header.csv
Part3_MLR_Rebate_Calculation.csv
```

Load them into Power Query as separate queries.

### Step 2: Filter Part3 to the required row lookup codes

Keep only these `ROW_LOOKUP_CODE` values:

```text
MLR_DENOMINATOR
MLR_NUMERATOR
MLR_STANDARD
REBATE_AMT_LIMITED
REBATE_AMT_CREDIBILITY_ADJ_MLR
```

This reduces the source to the metrics needed for the simplified project.

### Step 3: Keep only the required market columns

Keep these Part3 columns:

```text
MR_SUBMISSION_TEMPLATE_ID
ROW_LOOKUP_CODE
CMM_INDIVIDUAL_TOTAL
CMM_SMALL_GROUP_TOTAL
CMM_LARGE_GROUP_TOTAL
```

These represent the submission identifier, metric label, and three commercial medical market buckets.

### Step 4: Unpivot the market columns

The CMS file stores markets sideways across columns. Power Query's unpivot operation turns those columns into attribute-value rows.

Unpivot:

```text
CMM_INDIVIDUAL_TOTAL
CMM_SMALL_GROUP_TOTAL
CMM_LARGE_GROUP_TOTAL
```

Conceptual result:

```text
MR_SUBMISSION_TEMPLATE_ID | ROW_LOOKUP_CODE | Market_Raw | Value
```

Then map market names:

```text
CMM_INDIVIDUAL_TOTAL -> Individual
CMM_SMALL_GROUP_TOTAL -> Small Group
CMM_LARGE_GROUP_TOTAL -> Large Group
```

### Step 5: Pivot the row lookup codes into metric columns

After unpivoting, each row has a metric label and a value. Pivot `ROW_LOOKUP_CODE` so each selected metric becomes a column.

Before pivot:

```text
submission_id | market | row_lookup_code | value
```

After pivot:

```text
submission_id | market | MLR_DENOMINATOR | MLR_NUMERATOR | MLR_STANDARD | REBATE_AMT_LIMITED | REBATE_AMT_CREDIBILITY_ADJ_MLR
```

This is the main transformation from a filing-template layout into an analysis table.

### Step 6: Join Header to the cleaned Part3 table

Use Power Query Merge Queries.

```text
Left table: cleaned Part3 table
Right table: Header table
Join key: MR_SUBMISSION_TEMPLATE_ID
Join type: Left Outer
```

Expand the Header fields needed for analysis:

```text
COMPANY_NAME
BUSINESS_STATE
HIOS_ISSUER_ID
NAIC_COMPANY_CODE
GROUP_AFFILIATION
```

This is the Excel/Power Query equivalent of a SQL left join:

```sql
SELECT *
FROM cleaned_part3 p
LEFT JOIN header h
ON p.MR_SUBMISSION_TEMPLATE_ID = h.MR_SUBMISSION_TEMPLATE_ID;
```

### Step 7: Add calculated columns

Selected rebate amount:

```excel
=IF([@REBATE_AMT_LIMITED]<>"",[@REBATE_AMT_LIMITED],[@REBATE_AMT_CREDIBILITY_ADJ_MLR])
```

Calculated MLR:

```excel
=IFERROR([@MLR_NUMERATOR]/[@MLR_DENOMINATOR],"")
```

Rebate-positive flag:

```excel
=IF([@[Selected Rebate Amount]]>0,"Y","N")
```

### Step 8: Apply final filters

Keep rows where:

```text
MLR_DENOMINATOR > 0
MLR_STANDARD is not blank
BUSINESS_STATE is not Grand Total
Market is Individual, Small Group, or Large Group
```

These filters remove unusable rows and avoid double-counting aggregate `Grand Total` rows.

## Formula-only alternative

A formula-only workflow is possible, but less maintainable. A practical compromise is to use Power Query to create a long table, then use Excel formulas to pull fields.

Create a helper key in the long Part3 table:

```excel
=[@MR_SUBMISSION_TEMPLATE_ID]&"|"&[@Market]&"|"&[@ROW_LOOKUP_CODE]
```

Pull metrics with `XLOOKUP`:

```excel
=XLOOKUP([@MR_SUBMISSION_TEMPLATE_ID]&"|"&[@Market]&"|MLR_DENOMINATOR",Part3_Long[Key],Part3_Long[Value],"")
```

```excel
=XLOOKUP([@MR_SUBMISSION_TEMPLATE_ID]&"|"&[@Market]&"|MLR_NUMERATOR",Part3_Long[Key],Part3_Long[Value],"")
```

```excel
=XLOOKUP([@MR_SUBMISSION_TEMPLATE_ID]&"|"&[@Market]&"|MLR_STANDARD",Part3_Long[Key],Part3_Long[Value],"")
```

Attach Header fields with `XLOOKUP`:

```excel
=XLOOKUP([@MR_SUBMISSION_TEMPLATE_ID],Header[MR_SUBMISSION_TEMPLATE_ID],Header[COMPANY_NAME],"")
```

```excel
=XLOOKUP([@MR_SUBMISSION_TEMPLATE_ID],Header[MR_SUBMISSION_TEMPLATE_ID],Header[BUSINESS_STATE],"")
```

## Summary formulas

Weighted MLR by market:

```excel
=SUMIFS(Clean[MLR_NUMERATOR],Clean[Market],A2)/SUMIFS(Clean[MLR_DENOMINATOR],Clean[Market],A2)
```

Block count by market:

```excel
=COUNTIFS(Clean[Market],A2)
```

Rebate-positive block count:

```excel
=COUNTIFS(Clean[Market],A2,Clean[Selected Rebate Amount],">0")
```

Total selected rebate amount:

```excel
=SUMIFS(Clean[Selected Rebate Amount],Clean[Market],A2)
```

Top 10 rebate blocks:

```excel
=TAKE(SORT(FILTER(Clean,Clean[Selected Rebate Amount]>0),MATCH("Selected Rebate Amount",Clean[#Headers],0),-1),10)
```

## Optional VBA

VBA is not required for this project. If used, VBA should only refresh queries and formulas.

```vba
Sub RefreshMLRWorkbook()
    ThisWorkbook.RefreshAll
    Application.CalculateFullRebuild
End Sub
```

The core data-cleaning logic should stay in Power Query and documented formulas, not hidden in VBA.

## Python equivalent

The same workflow can be reproduced in Python/pandas using `read_csv`, `melt`, `pivot_table`, `merge`, and `groupby`.

```python
import pandas as pd

header = pd.read_csv("MR_Submission_Template_Header.csv")
part3 = pd.read_csv("Part3_MLR_Rebate_Calculation.csv")

metrics = [
    "MLR_DENOMINATOR",
    "MLR_NUMERATOR",
    "MLR_STANDARD",
    "REBATE_AMT_LIMITED",
    "REBATE_AMT_CREDIBILITY_ADJ_MLR",
]

market_cols = [
    "CMM_INDIVIDUAL_TOTAL",
    "CMM_SMALL_GROUP_TOTAL",
    "CMM_LARGE_GROUP_TOTAL",
]

part3_filtered = part3.loc[
    part3["ROW_LOOKUP_CODE"].isin(metrics),
    ["MR_SUBMISSION_TEMPLATE_ID", "ROW_LOOKUP_CODE", *market_cols]
]

long = part3_filtered.melt(
    id_vars=["MR_SUBMISSION_TEMPLATE_ID", "ROW_LOOKUP_CODE"],
    value_vars=market_cols,
    var_name="market_raw",
    value_name="value",
)

market_map = {
    "CMM_INDIVIDUAL_TOTAL": "Individual",
    "CMM_SMALL_GROUP_TOTAL": "Small Group",
    "CMM_LARGE_GROUP_TOTAL": "Large Group",
}
long["market"] = long["market_raw"].map(market_map)

clean = long.pivot_table(
    index=["MR_SUBMISSION_TEMPLATE_ID", "market"],
    columns="ROW_LOOKUP_CODE",
    values="value",
    aggfunc="first",
).reset_index()

clean = clean.merge(header, on="MR_SUBMISSION_TEMPLATE_ID", how="left")
clean["selected_rebate_amount"] = clean["REBATE_AMT_LIMITED"].fillna(
    clean["REBATE_AMT_CREDIBILITY_ADJ_MLR"]
)
clean["calculated_mlr"] = clean["MLR_NUMERATOR"] / clean["MLR_DENOMINATOR"]
```

## Metrics calculated

Calculated MLR:

```text
Calculated MLR = MLR Numerator / MLR Denominator
```

Weighted MLR by market:

```text
Weighted MLR = sum(MLR Numerator) / sum(MLR Denominator)
```

Weighted MLR is used instead of a simple average so that very small issuer-state blocks do not receive the same weight as large blocks.

Rebate-positive block:

```text
selected rebate amount > 0
```

## 2024 summary results

| Market | Blocks | Rebate-positive blocks | Total selected rebate | Weighted MLR | Benchmark standard |
|---|---:|---:|---:|---:|---:|
| Individual | 675 | 72 | $1.18B | 86.9% | 80.0% |
| Small Group | 530 | 59 | $274.0M | 86.5% | 80.0% |
| Large Group | 702 | 82 | $186.1M | 90.7% | 85.0% |

Total cleaned 2024 blocks analyzed: 1,907.

Total selected rebate amount: approximately $1.64B.

Overall weighted MLR across the included 2024 blocks: 89.0%.

## Interpretation

The aggregate 2024 weighted MLRs for the selected markets were above the broad benchmark standards. However, positive rebate amounts still appeared in specific issuer-state-market blocks.

Main takeaway:

```text
Aggregate market-level experience can look adequate,
while issuer-state-market rebate exposure can still exist.
```

## Future work

This project is intentionally Excel-first. Future versions could make the workflow more reproducible by adding a small Python script or notebook that performs the same steps shown in the workbook.

Recommended future repository additions:

```text
.
├── README.md
├── workbook/
│   └── Simplified Health MLR Interpretation Study - 2024.xlsx
├── src/
│   └── clean_mlr_2024.py
├── notebooks/
│   └── mlr_2024_cleaning_walkthrough.ipynb
├── output/
│   ├── clean_mlr_blocks_2024.csv
│   ├── market_summary_2024.csv
│   └── top_rebate_blocks_2024.csv
└── data/
    └── raw CMS files are not redistributed here
```

### Future Python script outline

The Python version would follow the same logic as the Excel Power Query workflow:

1. Read Header and Part3 CSV files.
2. Filter Part3 to selected `ROW_LOOKUP_CODE` values.
3. Keep only Individual, Small Group, and Large Group CMM total columns.
4. Reshape the three market columns from wide to long format.
5. Pivot selected row lookup codes into metric columns.
6. Merge Header fields by `MR_SUBMISSION_TEMPLATE_ID`.
7. Apply final filters.
8. Calculate selected rebate amount and calculated MLR.
9. Export clean block, market summary, and top rebate output files.

Starter script:

```python
from pathlib import Path
import pandas as pd

RAW_DIR = Path("data/raw/2024")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

HEADER_FILE = RAW_DIR / "MR_Submission_Template_Header.csv"
PART3_FILE = RAW_DIR / "Part3_MLR_Rebate_Calculation.csv"

METRICS = [
    "MLR_DENOMINATOR",
    "MLR_NUMERATOR",
    "MLR_STANDARD",
    "REBATE_AMT_LIMITED",
    "REBATE_AMT_CREDIBILITY_ADJ_MLR",
]

MARKET_COLS = {
    "CMM_INDIVIDUAL_TOTAL": "Individual",
    "CMM_SMALL_GROUP_TOTAL": "Small Group",
    "CMM_LARGE_GROUP_TOTAL": "Large Group",
}

HEADER_FIELDS = [
    "MR_SUBMISSION_TEMPLATE_ID",
    "COMPANY_NAME",
    "BUSINESS_STATE",
    "HIOS_ISSUER_ID",
    "NAIC_COMPANY_CODE",
    "GROUP_AFFILIATION",
]


def read_source_files() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Read the CMS Header and Part3 CSV files."""
    header = pd.read_csv(HEADER_FILE, low_memory=False)
    part3 = pd.read_csv(PART3_FILE, low_memory=False)
    return header, part3


def clean_part3(part3: pd.DataFrame) -> pd.DataFrame:
    """Reshape Part3 from filing-template layout into block-level metrics."""
    keep_cols = ["MR_SUBMISSION_TEMPLATE_ID", "ROW_LOOKUP_CODE", *MARKET_COLS.keys()]

    filtered = part3.loc[
        part3["ROW_LOOKUP_CODE"].isin(METRICS),
        keep_cols,
    ].copy()

    long = filtered.melt(
        id_vars=["MR_SUBMISSION_TEMPLATE_ID", "ROW_LOOKUP_CODE"],
        value_vars=list(MARKET_COLS.keys()),
        var_name="market_raw",
        value_name="value",
    )

    long["market"] = long["market_raw"].map(MARKET_COLS)
    long["value"] = pd.to_numeric(long["value"], errors="coerce")

    clean = long.pivot_table(
        index=["MR_SUBMISSION_TEMPLATE_ID", "market"],
        columns="ROW_LOOKUP_CODE",
        values="value",
        aggfunc="first",
    ).reset_index()

    clean.columns.name = None
    return clean


def attach_header(clean: pd.DataFrame, header: pd.DataFrame) -> pd.DataFrame:
    """Attach issuer and state fields from Header using MR_SUBMISSION_TEMPLATE_ID."""
    header_small = header[[c for c in HEADER_FIELDS if c in header.columns]].drop_duplicates()
    return clean.merge(header_small, on="MR_SUBMISSION_TEMPLATE_ID", how="left")


def add_calculated_fields(clean: pd.DataFrame) -> pd.DataFrame:
    """Add selected rebate, calculated MLR, and rebate-positive flag."""
    clean = clean.copy()

    clean["selected_rebate_amount"] = clean["REBATE_AMT_LIMITED"].fillna(
        clean["REBATE_AMT_CREDIBILITY_ADJ_MLR"]
    )

    clean["calculated_mlr"] = clean["MLR_NUMERATOR"] / clean["MLR_DENOMINATOR"]
    clean["rebate_positive"] = clean["selected_rebate_amount"].fillna(0) > 0

    return clean


def apply_filters(clean: pd.DataFrame) -> pd.DataFrame:
    """Apply final project-scope filters."""
    return clean.loc[
        (clean["MLR_DENOMINATOR"] > 0)
        & (clean["MLR_STANDARD"].notna())
        & (clean["BUSINESS_STATE"].ne("Grand Total"))
        & (clean["market"].isin(["Individual", "Small Group", "Large Group"]))
    ].copy()


def build_market_summary(clean: pd.DataFrame) -> pd.DataFrame:
    """Summarize weighted MLR and rebate exposure by market."""
    summary = clean.groupby("market", as_index=False).agg(
        block_count=("MR_SUBMISSION_TEMPLATE_ID", "count"),
        rebate_positive_blocks=("rebate_positive", "sum"),
        total_selected_rebate=("selected_rebate_amount", "sum"),
        mlr_denominator=("MLR_DENOMINATOR", "sum"),
        mlr_numerator=("MLR_NUMERATOR", "sum"),
    )

    summary["weighted_mlr"] = summary["mlr_numerator"] / summary["mlr_denominator"]
    summary["benchmark_standard"] = summary["market"].map(
        {"Individual": 0.80, "Small Group": 0.80, "Large Group": 0.85}
    )
    summary["margin_vs_benchmark"] = summary["weighted_mlr"] - summary["benchmark_standard"]

    return summary


def main() -> None:
    header, part3 = read_source_files()
    clean = clean_part3(part3)
    clean = attach_header(clean, header)
    clean = add_calculated_fields(clean)
    clean = apply_filters(clean)

    market_summary = build_market_summary(clean)
    top_rebates = clean.loc[clean["selected_rebate_amount"].fillna(0) > 0].sort_values(
        "selected_rebate_amount", ascending=False
    ).head(10)

    clean.to_csv(OUTPUT_DIR / "clean_mlr_blocks_2024.csv", index=False)
    market_summary.to_csv(OUTPUT_DIR / "market_summary_2024.csv", index=False)
    top_rebates.to_csv(OUTPUT_DIR / "top_rebate_blocks_2024.csv", index=False)


if __name__ == "__main__":
    main()
```

### Future validation checks

A later version could add validation checks such as:

```python
assert clean["MLR_DENOMINATOR"].gt(0).all()
assert clean["MLR_STANDARD"].notna().all()
assert clean["market"].isin(["Individual", "Small Group", "Large Group"]).all()
```

Other useful checks:

- Compare market-level block counts before and after filtering.
- Confirm no `Grand Total` rows remain.
- Confirm weighted MLR equals `sum(numerator) / sum(denominator)`.
- Confirm top rebate rows are sorted descending by selected rebate amount.

### Possible future analysis extensions

- Expand the workflow from 2024 only to 2020–2024.
- Add state-level summary tables.
- Add issuer group rollups.
- Add charts comparing weighted MLR by market.
- Add a short notebook walkthrough explaining each transformation step.
- Compare MLR interpretation with a separate public premium trend project.

## Limitations

- Uses public aggregate regulatory data, not claim-level data.
- Uses CMS-reported numerator, denominator, standard, and rebate fields.
- Does not recreate every CMS regulatory adjustment.
- Focuses only on 2024.
- Focuses only on Individual, Small Group, and Large Group comprehensive major medical markets.
- Should not be treated as an official CMS calculator.

## Skills demonstrated

- Public insurance data interpretation
- Excel Power Query workflow design
- Joining two CSV files by a shared ID
- Reshaping wide filing-template data into a long/clean format
- Pivoting row labels into analysis columns
- Weighted ratio calculation
- Rebate exposure flagging
- Documentation of assumptions and limitations




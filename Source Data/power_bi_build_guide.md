# Power BI Build Guide

This file is the beginner-friendly handoff for building the report now that the source files, curated model data, and first-pass semantic model are in place.

## Recommended Dashboard Scope

Keep the project focused on the UK labour market, with unemployment as the anchor topic and pay as the second pillar.

Use these four pages:

1. `Executive Summary`
2. `Labour Market Overview`
3. `Local Area Explorer`
4. `Pay And Demographics`

## Data Sources Selected

These are the public sources the project is being built around:

1. `ONS / EMP`
   Average weekly earnings time series.
2. `ONS / UNEM`
   UK claimant count and vacancies time series.
3. `ONS / CC01`
   Latest claimant count by local and unitary authority.
4. `ONS / M01`
   Modelled unemployment for local and unitary authorities.
5. `ONS / ASHE Table 8`
   Local authority earnings by place of residence.

## Power BI Import Order

The raw source files are now transformed into model-ready CSVs inside `Source Data\Model Data`.

If you want to refresh the source data outside Power BI, use:

1. Double-click `Source Data\refresh_data.bat`

That batch file runs both scripts in the correct order:

1. `download_uk_labour_market_data.py`
2. `prepare_model_data.py`

After opening the PBIP, do this:

1. Open Power BI Desktop and load the PBIP project.
2. Let the semantic model load from the TMDL definition.
3. Refresh the model.
4. Confirm the following tables appear:
   `Dim Date`
   `Dim Area`
   `Fact National Monthly`
   `Fact Region Monthly`
   `Fact Local Claimant`
   `Fact Local Unemployment History`
   `Fact Local Pay`
   `Outlook Snapshot`

## Core Model Shape

Use a star schema, not a flat single-table report.

### Dimensions

1. `Dim Date`
   One row per month from the earliest claimant data to the latest month.
2. `Dim Area`
   One row per area code.
   Include area name, geography type, and a manual `Area Level` column.
### Facts

1. `Fact National Monthly`
2. `Fact Region Monthly`
3. `Fact Local Claimant`
4. `Fact Local Unemployment History`
5. `Fact Local Pay`
6. `Outlook Snapshot`

## DAX Measures

The semantic model already includes the first pass of measures. Use the existing measures before creating more.

### Measures already available

1. `Outlook Direction`
2. `Outlook Headline`
3. `Outlook Detail`
4. `Latest Claimant Count (K)`
5. `Latest Vacancies (K)`
6. `Latest Real Pay YoY %`
7. `Claimant 3M Delta (K)`
8. `Vacancies 3M Delta (K)`
9. `National Claimant Count (K)`
10. `National Claimant Rate %`
11. `Vacancies (K)`
12. `Vacancy Rate Per 100 Jobs`
13. `Regional Claimant Count (K)`
14. `Regional Claimant Rate %`
15. `Local Claimant Count`
16. `Local Claimant Rate %`
17. `Local Claimant YoY Change`
18. `Modelled Unemployment Count`
19. `Modelled Unemployment Rate %`
20. `Median Weekly Pay GBP`
21. `Mean Weekly Pay GBP`
22. `Median Pay YoY %`
23. `Pay Gender Gap GBP`
24. `Selected Area`

## Page 1: Executive Summary

This is the main landing page and should feel like a portfolio-quality control room.

Use these visuals:

1. Card: `Outlook Direction`
2. Multi-row card: `Outlook Headline`, `Outlook Detail`
3. Card: `Latest Common Signal Date`
4. Card: `Latest Claimant Count (K)`
5. Card: `Latest Vacancies (K)`
6. Card: `Latest Real Pay YoY %`
7. Card: `Claimant 3M Delta (K)`
8. Card: `Vacancies 3M Delta (K)`
9. Line chart: `National Claimant Count (K)` by `Dim Date[YearMonth]`
10. Line chart: `Vacancies (K)` by `Dim Date[YearMonth]`

### Fields to drop

1. Summary direction card
   Values: `Outlook Snapshot[Outlook Direction]` measure
2. Narrative card
   Values: `Outlook Snapshot[Outlook Headline]`, `Outlook Snapshot[Outlook Detail]` measures
3. National claimant trend
   Axis: `Dim Date[YearMonth]`
   Values: `Fact National Monthly[National Claimant Count (K)]` measure
4. Vacancies trend
   Axis: `Dim Date[YearMonth]`
   Values: `Fact National Monthly[Vacancies (K)]` measure

## Page 2: Labour Market Overview

Use these visuals:

1. Card: `National Claimant Count (K)`
2. Card: `National Claimant Rate %`
3. Card: `Vacancy Rate Per 100 Jobs`
4. Line chart: monthly claimant count trend
5. Line chart: national average weekly earnings trend
6. Clustered bar chart: latest claimant count by region/country
7. Matrix: latest claimant count and claimant rate by region/country

### Fields to drop

1. Line chart for claimant trend
   Axis: `Dim Date[YearMonth]`
   Values: `Fact National Monthly[National Claimant Count (K)]`
2. Vacancies trend
   Axis: `Dim Date[YearMonth]`
   Values: `Fact National Monthly[Vacancies (K)]`
3. AWE trend
   Axis: `Dim Date[YearMonth]`
   Values: `Fact National Monthly[Real Regular Pay Level GBP]`
4. Regional comparison
   Axis: `Dim Area[AreaName]`
   Values: `Fact Region Monthly[Regional Claimant Count (K)]`
   Filter: latest visible month only

## Page 3: Local Area Explorer

Use these visuals:

1. Map or filled map by local authority
2. Bar chart of top 15 areas by claimant count
3. Line chart for selected area unemployment history
4. Matrix with area, claimant count, unemployment rate, rank
5. Slicers: area, area level, date

### Fields to drop

1. Map
   Location: `Dim Area[AreaName]`
   Tooltip: `Fact Local Claimant[Local Claimant Count]`, `Fact Local Unemployment History[Modelled Unemployment Rate %]`
2. Selected area trend
   Axis: `Fact Local Unemployment History[PeriodLabel]`
   Values: `Fact Local Unemployment History[Modelled Unemployment Count]`
3. Ranking bar chart
   Axis: `Dim Area[AreaName]`
   Values: `Fact Local Claimant[Local Claimant Count]`
   Filter: Top N = 15 by `Fact Local Claimant[Local Claimant Count]`

## Page 4: Pay And Demographics

Use these visuals:

1. Bar chart: median weekly pay by selected areas
2. Line or column chart: AWE trend and YoY change
3. Column chart: local claimant count comparison
4. Table: selected area pay and unemployment detail
5. Scatter chart: unemployment rate vs median weekly pay
6. Detail table for selected area

### Fields to drop

1. Pay comparison
   Axis: `Dim Area[AreaName]`
   Values: `Fact Local Pay[Median Weekly Pay GBP]`
2. Local claimant comparison
   Axis: `Dim Area[AreaName]`
   Values: `Fact Local Claimant[Local Claimant Count]`
3. Scatter
   X-axis: unemployment rate measure
   Y-axis: `Fact Local Pay[Median Weekly Pay GBP]`
   Details: `Dim Area[AreaName]`

## Manual Work I Expect You To Do In Power BI

These are the parts that are still easier to do by hand in Desktop:

1. Dragging fields into visuals
2. Choosing exact colors, fonts, and spacing
3. Final map visual choice
4. Tooltip polishing
5. Final sorting and formatting

## Suggested Visual Style

Use a clean portfolio style:

1. Background: off-white or very light grey
2. Accent color: deep blue or muted teal
3. Highlight color for worsening unemployment: rust or dark orange
4. Use consistent card sizes and left-aligned titles

## Next Execution Step

If you need to rebuild the curated model files, run:

```powershell
python ".\Source Data\download_uk_labour_market_data.py"
python ".\Source Data\prepare_model_data.py"
```

The current script now uses direct ONS published downloads and writes the final source files straight into `Source Data`. Do not keep rebuilding the model before the data is stable.

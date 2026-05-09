---
<div align="center">

<img src="./repo-card.png" alt="UK Labour Market Dashboard project card" width="100%" />
<br /><br />

<p><strong>Labour market dashboard tracking UK unemployment trends, claimant counts, vacancies, pay movement, and regional variation using official ONS data. Live interactive version published on Power BI service.</strong></p>

<p>Built for analysts, recruiters, policy watchers, and business users who need a clear view of labour market pressure across England without manually collecting and comparing ONS datasets.</p>

<p>
  <a href="#overview">Overview</a> |
  <a href="#what-problem-it-solves">What It Solves</a> |
  <a href="#feature-highlights">Features</a> |
  <a href="#screenshots">Screenshots</a> |
  <a href="#quick-start">Quick Start</a> |
  <a href="#tech-stack">Tech Stack</a>
</p>

<h3><strong>Made by Naadir | May 2026</strong></h3>

</div>

---

## Overview

This Power BI dashboard turns official ONS labour market data into an interactive report for tracking unemployment pressure across England and the wider UK labour market. It brings together claimant count, claimant rate, vacancies, vacancies per claimant, real pay movement, unemployment rate, and local authority-level variation.

The report supports both national monitoring and local area exploration. Users can move from an executive summary into detailed regional, local authority, pay, and demographic pages to understand where unemployment pressure is rising, where vacancies are falling, and how pay levels differ across areas.

The practical outcome is a single Power BI report that replaces scattered manual checks across raw data tables. It gives users a fast way to monitor labour market direction, compare places, identify pressure points, and communicate trends using clean visuals and calculated metrics.

## What Problem It Solves

- Removes the need to manually inspect separate labour market tables for claimant counts, vacancies, pay, and unemployment indicators
- Replaces spreadsheet-based comparison with an interactive Power BI report using Power Query transformations and DAX measures
- Makes regional and local unemployment pressure easier to see through maps, trend charts, bar charts, KPI cards, and tables
- Gives a clearer view than raw ONS downloads by combining national trends, local variation, pay movement, and vacancy pressure in one report

### At a glance

| Track | Analyse | Compare |
|---|---|---|
| Claimant counts, claimant rates, vacancies, pay levels, and unemployment rates | Labour market direction, year-on-year pay movement, vacancy pressure, and claimant changes | National, regional, country, and local authority results side-by-side |
| ONS labour market datasets loaded through Power Query | DAX measures for rates, deltas, ratios, and latest-period KPIs | Current year vs prior year, national vs local, and regular pay vs total pay |
| Interactive report navigation, slicers, and area selection | KPI cards, time-series charts, treemaps, scatter plots, and detail tables | Claimant pressure, vacancy movement, and pay distribution across areas |

## Feature Highlights

- **Executive summary**, gives a fast status view of the labour market with headline KPIs, short-term claimant change, vacancy change, and a clear takeaway
- **Labour market overview**, tracks national claimant count, vacancy trends, real pay movement, vacancy rate, vacancies per claimant, and 12-month claimant change
- **Local area explorer**, lets users filter by area and inspect claimant count, claimant rate, modelled unemployment rate, top local areas, and local authority tables
- **Pay and demographics page**, compares median weekly pay, mean pay, gender pay gap, pay distribution, and pay versus unemployment by local area
- **Regional comparison**, shows claimant count and claimant rate by region and country so users can identify the highest-pressure areas quickly
- **Power BI modelling**, uses Power Query for data preparation and DAX for reusable measures, calculated KPIs, ratios, deltas, and report-level metrics

### Core capabilities

| Area | What it gives you |
|---|---|
| **National labour market monitoring** | A high-level view of claimant count, vacancy levels, real pay growth, and forward labour market pressure |
| **Regional and local comparison** | A way to compare claimant rates and claimant counts across England, UK regions, and local authorities |
| **Pay analysis** | Visibility into weekly pay levels, pay distribution, gender pay gap, and pay movement over time |
| **Interactive reporting** | Slicers, navigation pages, KPI cards, charts, and tables for exploring the data without editing the model |

## Screenshots

<details>
<summary><strong>Open screenshot gallery</strong></summary>

<br />

<div align="center">
  <img src="./portfolio/Screen1.png" alt="Labour market overview page showing vacancy rate, vacancies per claimant, real pay growth, claimant movement, national trend charts, regional claimant counts, and regional metrics table" width="88%" />
  <br /><br />
  <img src="./portfolio/Screen2.png" alt="Executive summary page showing labour market status, headline claimant and vacancy KPIs, year-on-year comparison charts, and pay level trends" width="88%" />
  <br /><br />
  <img src="./portfolio/Screen3.png" alt="Local area explorer page showing selected areas, claimant rate, local claimant count, unemployment rate, treemap, top local areas, history chart, and local authority metrics table" width="88%" />
</div>

</details>

## Quick Start

~~~bash
# Clone the repo
git clone https://github.com/Naadir-Dev-Portfolio/powerbi-uk-Labour-Market-Dashboard.git
cd powerbi-uk-Labour-Market-Dashboard

# Install dependencies
No install required

# Run
Open the .pbix file in Power BI Desktop
~~~

No API keys are required. The report is built in Power BI using public ONS data, Power Query transformations, and DAX measures. To refresh the report locally, open the `.pbix` file in Power BI Desktop and use the built-in refresh option. The live version can be published and viewed through Power BI service.

## Tech Stack

<details>
<summary><strong>Open tech stack</strong></summary>

<br />

| Category | Tools |
|---|---|
| **Primary stack** | `DAX` | `Power Query` |
| **UI / App layer** | `Power BI Desktop` | `Power BI Service` |
| **Data / Storage** | `ONS public data` | `Power BI data model` | `Imported tables` |
| **Automation / Integration** | `Power Query refresh` | `DAX measures` | `Power BI report publishing` |
| **Platform** | `Windows` | `Web` |

</details>

## Architecture & Data

<details>
<summary><strong>Open architecture and data details</strong></summary>

<br />

### Application model

The report starts with public ONS labour market data. Power Query loads, cleans, reshapes, and standardises the source tables before they are added to the Power BI data model. DAX measures calculate claimant rates, vacancy ratios, short-term changes, year-on-year pay movement, latest-period KPIs, and totals used across the report pages.

The finished model feeds multiple Power BI pages: a labour market overview, an executive summary, a local area explorer, and a pay and demographics view. Users interact with slicers, area filters, navigation icons, KPI cards, charts, and tables to move from headline trends into local detail.

### Project structure

~~~text
powerbi-uk-Labour-Market-Dashboard/
+-- UK Labour Market Dashboard.pbip
+-- UK Labour Market Dashboard.Report/
+-- UK Labour Market Dashboard.SemanticModel/
+-- README.md
+-- repo-card.png
+-- portfolio/
    +-- powerbi-uk-labour-market-dashboard.json
    +-- powerbi-uk-labour-market-dashboard.png
    +-- powerbi-uk-labour-market-dashboard-featured.png
    +-- powerbi-uk-labour-market-dashboard-full.png
    +-- Screen1.png
    +-- Screen2.png
    +-- Screen3.png
~~~

### Data / system notes

- Data is sourced from public ONS labour market datasets and transformed inside Power BI using Power Query
- No API keys, secrets, or private credentials are required to use the report
- The report can be refreshed in Power BI Desktop and published to Power BI service for interactive web access

</details>

## Contact

Questions, feedback, or collaboration: `naadir.dev.mail@gmail.com`

<sub>DAX | Power Query</sub>

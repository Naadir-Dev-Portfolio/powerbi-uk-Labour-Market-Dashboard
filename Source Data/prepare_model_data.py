from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import win32com.client as win32


ROOT = Path(__file__).resolve().parent
MODEL_DIR = ROOT / "Model Data"


def classify_gss_code(code: str) -> str:
    code = (code or "").strip()
    if not code:
        return "Unknown"
    if code.startswith(("E06", "E07", "E08", "E09", "W06", "S12", "N09")):
        return "Local authority"
    if code.startswith(("E10", "E11")):
        return "County or combined area"
    if code.startswith(("E12",)):
        return "Region"
    if code.startswith(("E92", "W92", "S92", "N92")):
        return "Country"
    if code.startswith(("K02", "K03", "K04")):
        return "National aggregate"
    return "Other"


def month_start_from_title(series: pd.Series) -> pd.Series:
    titles = series.astype(str).str.strip()
    return pd.to_datetime(titles, format="%Y %b", errors="coerce").dt.to_period("M").dt.to_timestamp()


def load_excel_used_range(workbook_path: Path, sheet_name: str) -> list[list[object]]:
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = excel.Workbooks.Open(str(workbook_path))
    try:
        worksheet = workbook.Worksheets(sheet_name)
        values = worksheet.UsedRange.Value
        return [list(row) for row in values]
    finally:
        workbook.Close(False)
        excel.Quit()


def sanitize_header(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def find_column(columns: list[str], required_parts: list[str]) -> str:
    for column in columns:
        normalized = column.lower()
        if all(part.lower() in normalized for part in required_parts):
            return column
    raise KeyError(f"Could not find column with parts: {required_parts}")


def prepare_national_and_region_tables() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    claimant = pd.read_csv(ROOT / "ons_claimant_and_vacancies.csv")
    claimant = claimant[claimant["Title"].astype(str).str.match(r"^\d{4} [A-Z]{3}$")].copy()
    claimant["Date"] = month_start_from_title(claimant["Title"])

    national = pd.DataFrame(
        {
            "Date": claimant["Date"],
            "NationalClaimantCountK": claimant[
                "Claimant Count : K02000001 UK : People : SA : Thousands"
            ],
            "NationalClaimantRatePct": claimant[
                "Claimant Count : K02000001 UK : People : SA : Percentage (%)"
            ],
            "VacanciesK": claimant["UK Vacancies (thousands) - Total"],
            "VacancyRatePer100Jobs": claimant["UK Job Vacancies ratio per 100 emp jobs - Total"],
        }
    )
    for col in [
        "NationalClaimantCountK",
        "NationalClaimantRatePct",
        "VacanciesK",
        "VacancyRatePer100Jobs",
    ]:
        national[col] = pd.to_numeric(national[col], errors="coerce")

    awe = pd.read_csv(ROOT / "ons_average_weekly_earnings.csv")
    awe = awe[awe["Title"].astype(str).str.match(r"^\d{4} [A-Z]{3}$")].copy()
    awe["Date"] = month_start_from_title(awe["Title"])

    keep_map = {
        "RealRegularPayIndex": find_column(
            list(awe.columns),
            ["whole economy", "real terms index", "seasonally adjusted regular pay"],
        ),
        "RealRegularPayYoYSinglePct": find_column(
            list(awe.columns),
            ["whole economy", "real terms year on year single month growth", "seasonally adjusted regular pay"],
        ),
        "RealRegularPayLevelGBP": find_column(
            list(awe.columns),
            ["whole economy", "real terms level", "seasonally adjusted regular pay"],
        ),
        "RealTotalPayIndex": find_column(
            list(awe.columns),
            ["whole economy", "real terms index", "seasonally adjusted total pay"],
        ),
        "RealTotalPayYoYSinglePct": find_column(
            list(awe.columns),
            ["whole economy", "real terms year on year single month growth", "seasonally adjusted total pay"],
        ),
        "RealTotalPayLevelGBP": find_column(
            list(awe.columns),
            ["whole economy", "real terms level", "seasonally adjusted total pay"],
        ),
    }
    awe_small = awe[["Date", *keep_map.values()]].rename(columns={v: k for k, v in keep_map.items()})
    for col in [col for col in awe_small.columns if col != "Date"]:
        awe_small[col] = pd.to_numeric(awe_small[col], errors="coerce")

    national = national.merge(awe_small, on="Date", how="left").sort_values("Date").reset_index(drop=True)
    national["VacanciesPerClaimant"] = national["VacanciesK"] / national["NationalClaimantCountK"]

    for col in [
        "NationalClaimantCountK",
        "NationalClaimantRatePct",
        "VacanciesK",
        "VacancyRatePer100Jobs",
        "RealRegularPayYoYSinglePct",
        "RealTotalPayYoYSinglePct",
        "VacanciesPerClaimant",
    ]:
        national[f"{col}_3MDelta"] = national[col] - national[col].shift(3)
        national[f"{col}_12MDelta"] = national[col] - national[col].shift(12)

    region_records: list[dict[str, object]] = []
    count_pattern = re.compile(
        r"^Claimant Count : (?P<code>[A-Z0-9]+) (?P<name>.+?) : People : SA : Thousands$"
    )

    for column in claimant.columns:
        match = count_pattern.match(column)
        if not match:
            continue
        code = match.group("code")
        name = match.group("name").strip()
        rate_column = f"Claimant Count : {code} {name} : People : SA : Percentage (%)"
        if rate_column not in claimant.columns:
            continue
        level = classify_gss_code(code)
        if level not in {"National aggregate", "Country", "Region"}:
            continue
        temp = pd.DataFrame(
            {
                "Date": claimant["Date"],
                "AreaCode": code,
                "AreaName": name,
                "AreaLevel": level,
                "ClaimantCountK": claimant[column],
                "ClaimantRatePct": claimant[rate_column],
            }
        )
        temp["ClaimantCountK"] = pd.to_numeric(temp["ClaimantCountK"], errors="coerce")
        temp["ClaimantRatePct"] = pd.to_numeric(temp["ClaimantRatePct"], errors="coerce")
        region_records.extend(temp.to_dict("records"))

    region = pd.DataFrame(region_records).sort_values(["AreaLevel", "AreaName", "Date"]).reset_index(drop=True)

    latest = national.dropna(
        subset=["NationalClaimantCountK", "VacanciesK", "RealRegularPayYoYSinglePct"]
    ).iloc[-1].copy()
    signal_score = 0.0
    signal_score += -1 if latest["NationalClaimantCountK_3MDelta"] > 0 else 1
    signal_score += 1 if latest["VacanciesK_3MDelta"] > 0 else -1
    signal_score += 0.5 if latest["RealRegularPayYoYSinglePct"] > 0 else -0.5

    if signal_score >= 1:
        direction = "Improving"
        headline = "Unemployment pressure is easing and forward indicators are still supportive."
    elif signal_score <= -1:
        direction = "Deteriorating"
        headline = "Unemployment pressure is worsening and forward indicators suggest more softening ahead."
    else:
        direction = "Steady to softening"
        headline = "The labour market is broadly steady, but vacancies and claimant trends need close monitoring."

    detail = (
        f"Latest month: {latest['Date'].strftime('%B %Y')}. "
        f"Claimant count moved {latest['NationalClaimantCountK_3MDelta']:.1f}k over 3 months, "
        f"vacancies moved {latest['VacanciesK_3MDelta']:.1f}k, and real regular pay growth is "
        f"{latest['RealRegularPayYoYSinglePct']:.1f}% year on year."
    )

    snapshot = pd.DataFrame(
        [
            {
                "SnapshotDate": latest["Date"],
                "Direction": direction,
                "Headline": headline,
                "Detail": detail,
                "NationalClaimantCountK": latest["NationalClaimantCountK"],
                "NationalClaimantRatePct": latest["NationalClaimantRatePct"],
                "VacanciesK": latest["VacanciesK"],
                "VacancyRatePer100Jobs": latest["VacancyRatePer100Jobs"],
                "RealRegularPayYoYSinglePct": latest["RealRegularPayYoYSinglePct"],
                "VacanciesPerClaimant": latest["VacanciesPerClaimant"],
                "Claimant3MDeltaK": latest["NationalClaimantCountK_3MDelta"],
                "Vacancies3MDeltaK": latest["VacanciesK_3MDelta"],
                "Claimant12MDeltaK": latest["NationalClaimantCountK_12MDelta"],
                "Vacancies12MDeltaK": latest["VacanciesK_12MDelta"],
                "SignalScore": signal_score,
            }
        ]
    )

    return national, region, snapshot


def prepare_cc01_local_claimant() -> pd.DataFrame:
    rows = load_excel_used_range(ROOT / "ons_cc01_local_claimant_count.xls", "CC01")
    header = [sanitize_header(x) for x in rows[4]]
    df = pd.DataFrame(rows[5:], columns=header)
    df = df[df["Geography code"].notna()].copy()

    month_match = re.search(r"([A-Za-z]+ \d{4})", str(rows[0][0]))
    snapshot_date = pd.to_datetime(month_match.group(1), format="%B %Y") if month_match else pd.NaT

    df = df.rename(
        columns={
            "Geography": "AreaName",
            "Geography code": "AreaCode",
            "Number of men1": "MenCount",
            "Number of women1": "WomenCount",
            "Number of people1": "PeopleCount",
            "Proportion of men2": "MenRatePct",
            "Proportion of women2": "WomenRatePct",
            "Proportion of people2": "PeopleRatePct",
            "Men: change on year": "MenYoYChange",
            "Women: change on year": "WomenYoYChange",
            "People: change on year": "PeopleYoYChange",
            "Men: proportion change on year2": "MenRateYoYChange",
        }
    )
    df["SnapshotDate"] = snapshot_date
    df["AreaLevel"] = df["AreaCode"].astype(str).map(classify_gss_code)
    df["IsLocalArea"] = df["AreaLevel"].eq("Local authority")
    numeric_cols = [
        "MenCount",
        "WomenCount",
        "PeopleCount",
        "MenRatePct",
        "WomenRatePct",
        "PeopleRatePct",
        "MenYoYChange",
        "WomenYoYChange",
        "PeopleYoYChange",
        "MenRateYoYChange",
    ]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df[
        [
            "SnapshotDate",
            "AreaName",
            "AreaCode",
            "AreaLevel",
            "IsLocalArea",
            "MenCount",
            "WomenCount",
            "PeopleCount",
            "MenRatePct",
            "WomenRatePct",
            "PeopleRatePct",
            "MenYoYChange",
            "WomenYoYChange",
            "PeopleYoYChange",
            "MenRateYoYChange",
        ]
    ].reset_index(drop=True)


def prepare_m01_local_history() -> pd.DataFrame:
    levels = load_excel_used_range(ROOT / "ons_m01_modelled_unemployment.xls", "LA,UA Levels")
    rates = load_excel_used_range(ROOT / "ons_m01_modelled_unemployment.xls", "LA,UA Rates")

    period_columns: list[tuple[int, str]] = []
    for idx, value in enumerate(levels[2]):
        if value and re.match(r"^(\d{4}/\d{2,4}|[A-Z][a-z]{2} \d{4} to [A-Z][a-z]{2} \d{4})$", str(value)):
            period_columns.append((idx, str(value)))

    level_records: list[dict[str, object]] = []
    for row in levels[5:]:
        area_name, area_code = row[0], row[1]
        if not area_code:
            continue
        for idx, period_label in period_columns:
            level_value = row[idx]
            ci_value = row[idx + 1] if idx + 1 < len(row) else None
            if level_value in (None, ":"):
                continue
            level_records.append(
                {
                    "AreaName": area_name,
                    "AreaCode": area_code,
                    "PeriodLabel": period_label,
                    "UnemploymentCount": level_value,
                    "UnemploymentCountCI": ci_value,
                }
            )

    rate_records: list[dict[str, object]] = []
    for row in rates[5:]:
        area_name, area_code = row[0], row[1]
        if not area_code:
            continue
        for idx, period_label in period_columns:
            rate_value = row[idx]
            ci_value = row[idx + 1] if idx + 1 < len(row) else None
            if rate_value in (None, ":"):
                continue
            rate_records.append(
                {
                    "AreaName": area_name,
                    "AreaCode": area_code,
                    "PeriodLabel": period_label,
                    "UnemploymentRatePct": rate_value,
                    "UnemploymentRateCI": ci_value,
                }
            )

    levels_df = pd.DataFrame(level_records)
    rates_df = pd.DataFrame(rate_records)
    merged = levels_df.merge(rates_df, on=["AreaName", "AreaCode", "PeriodLabel"], how="left")
    period_order = {label: i + 1 for i, (_, label) in enumerate(period_columns)}
    merged["PeriodSort"] = merged["PeriodLabel"].map(period_order)
    merged["AreaLevel"] = merged["AreaCode"].astype(str).map(classify_gss_code)
    for col in [
        "UnemploymentCount",
        "UnemploymentCountCI",
        "UnemploymentRatePct",
        "UnemploymentRateCI",
        "PeriodSort",
    ]:
        merged[col] = pd.to_numeric(merged[col], errors="coerce")
    return merged.reset_index(drop=True)


def prepare_ashe_local_pay() -> pd.DataFrame:
    workbook = ROOT / "ons_ashe_table8_weekly_pay_gross_2025.xlsx"
    records: list[pd.DataFrame] = []
    for sheet in ["All", "Male", "Female"]:
        df = pd.read_excel(workbook, sheet_name=sheet, header=4)
        df.columns = [sanitize_header(col) for col in df.columns]
        df = df[df["Code"].notna()].copy()
        df = df[df["Code"].astype(str).str.match(r"^[A-Z][0-9A-Z]{8}$")]
        df["SexGroup"] = sheet
        records.append(
            df[
                [
                    "Description",
                    "Code",
                    "(thousand)",
                    "Median",
                    "change",
                    "Mean",
                    "change.1",
                    "10",
                    "20",
                    "25",
                    "30",
                    "40",
                    "60",
                    "70",
                    "75",
                    "80",
                    "90",
                    "SexGroup",
                ]
            ]
        )

    pay = pd.concat(records, ignore_index=True)
    pay = pay.rename(
        columns={
            "Description": "AreaName",
            "Code": "AreaCode",
            "(thousand)": "JobsThousand",
            "Median": "MedianWeeklyPayGBP",
            "change": "MedianYoYGrowthPct",
            "Mean": "MeanWeeklyPayGBP",
            "change.1": "MeanYoYGrowthPct",
            "10": "P10",
            "20": "P20",
            "25": "P25",
            "30": "P30",
            "40": "P40",
            "60": "P60",
            "70": "P70",
            "75": "P75",
            "80": "P80",
            "90": "P90",
        }
    )
    pay["AreaLevel"] = pay["AreaCode"].astype(str).map(classify_gss_code)
    pay["PayYear"] = 2025
    for col in [
        "JobsThousand",
        "MedianWeeklyPayGBP",
        "MedianYoYGrowthPct",
        "MeanWeeklyPayGBP",
        "MeanYoYGrowthPct",
        "P10",
        "P20",
        "P25",
        "P30",
        "P40",
        "P60",
        "P70",
        "P75",
        "P80",
        "P90",
        "PayYear",
    ]:
        pay[col] = pd.to_numeric(pay[col], errors="coerce")
    return pay.reset_index(drop=True)


def prepare_dim_area(*frames: pd.DataFrame) -> pd.DataFrame:
    priority_frames = []
    for priority, frame in enumerate(frames, start=1):
        subset = frame[[col for col in frame.columns if col in {"AreaName", "AreaCode", "AreaLevel"}]].copy()
        subset = subset.dropna(subset=["AreaCode"]).copy()
        subset["AreaName"] = subset["AreaName"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
        subset["AreaLevel"] = subset["AreaLevel"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
        subset["SourcePriority"] = priority
        priority_frames.append(subset)

    area = pd.concat(priority_frames, ignore_index=True)
    area = area.sort_values(["SourcePriority", "AreaLevel", "AreaName"])
    area = area.drop_duplicates(subset=["AreaCode"], keep="first").copy()
    area = area.drop(columns=["SourcePriority"])
    area = area.sort_values(["AreaLevel", "AreaName"]).reset_index(drop=True)
    area["AreaSort"] = area.groupby("AreaLevel").cumcount() + 1
    area["IsLocalArea"] = area["AreaLevel"].eq("Local authority")
    return area


def prepare_dim_date(national: pd.DataFrame, claimant_local: pd.DataFrame) -> pd.DataFrame:
    min_date = min(national["Date"].min(), claimant_local["SnapshotDate"].min())
    max_date = max(national["Date"].max(), claimant_local["SnapshotDate"].max())
    dates = pd.date_range(min_date, max_date, freq="MS")
    dim_date = pd.DataFrame({"Date": dates})
    dim_date["Year"] = dim_date["Date"].dt.year
    dim_date["MonthNo"] = dim_date["Date"].dt.month
    dim_date["MonthName"] = dim_date["Date"].dt.strftime("%B")
    dim_date["Quarter"] = "Q" + dim_date["Date"].dt.quarter.astype(str)
    dim_date["YearMonth"] = dim_date["Date"].dt.strftime("%Y-%m")
    dim_date["MonthSort"] = dim_date["Year"] * 100 + dim_date["MonthNo"]
    max_month = dim_date["Date"].max()
    dim_date["MonthsFromLatest"] = (
        (max_month.year - dim_date["Date"].dt.year) * 12
        + (max_month.month - dim_date["Date"].dt.month)
    )
    dim_date["IsLast24Months"] = dim_date["MonthsFromLatest"].between(0, 23)
    dim_date["IsLast60Months"] = dim_date["MonthsFromLatest"].between(0, 59)
    dim_date["IsLast120Months"] = dim_date["MonthsFromLatest"].between(0, 119)
    return dim_date


def main() -> None:
    MODEL_DIR.mkdir(parents=True, exist_ok=True)

    national, region, snapshot = prepare_national_and_region_tables()
    claimant_local = prepare_cc01_local_claimant()
    local_history = prepare_m01_local_history()
    local_pay = prepare_ashe_local_pay()
    dim_area = prepare_dim_area(claimant_local, region, local_history, local_pay)
    dim_date = prepare_dim_date(national, claimant_local)

    outputs = {
        "dim_date.csv": dim_date,
        "dim_area.csv": dim_area,
        "fact_national_monthly.csv": national,
        "fact_region_monthly.csv": region,
        "fact_local_claimant_latest.csv": claimant_local,
        "fact_local_unemployment_history.csv": local_history,
        "fact_local_pay_2025.csv": local_pay,
        "outlook_snapshot.csv": snapshot,
    }

    for file_name, frame in outputs.items():
        frame.to_csv(MODEL_DIR / file_name, index=False)
        print(f"Created {file_name} ({len(frame):,} rows)")


if __name__ == "__main__":
    main()

from __future__ import annotations

import csv
import io
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from collections import Counter
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from urllib.error import HTTPError
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT

NOMIS_BASE = "https://www.nomisweb.co.uk/api/v01/dataset"
USER_AGENT = "Mozilla/5.0 (compatible; powerbi-uk-Labour-Market-Dashboard/1.0)"


class LinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        for key, value in attrs:
            if key.lower() == "href" and value:
                self.links.append(value)


@dataclass
class DownloadResult:
    name: str
    status: str
    output: str | None = None
    notes: list[str] | None = None


def log(message: str) -> None:
    print(message)


def ensure_dirs() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)


def fetch_bytes(url: str) -> bytes:
    delays = [2, 5, 10, 20]
    last_error: Exception | None = None
    for attempt, delay in enumerate(delays, start=1):
        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                return response.read()
        except HTTPError as exc:
            last_error = exc
            if exc.code != 429 or attempt == len(delays):
                raise
            time.sleep(delay)
    if last_error:
        raise last_error
    raise RuntimeError(f"Failed to download {url}")


def fetch_text(url: str) -> str:
    return fetch_bytes(url).decode("utf-8", errors="replace")


def write_bytes(path: Path, content: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)


def first_matching_link(page_url: str, extensions: tuple[str, ...]) -> str:
    html = fetch_text(page_url)
    parser = LinkParser()
    parser.feed(html)

    candidates: list[str] = []
    for link in parser.links:
        lowered = link.lower()
        if not any(ext in lowered for ext in extensions):
            continue
        if link.startswith("http://") or link.startswith("https://"):
            candidates.append(link)
        else:
            candidates.append(urllib.parse.urljoin(page_url, link))

    if not candidates:
        raise RuntimeError(f"No matching download link found on {page_url}")
    return candidates[0]


def csv_rows_from_url(url: str) -> list[dict[str, str]]:
    content = fetch_text(url)
    reader = csv.DictReader(io.StringIO(content))
    return list(reader)


def discover_geography_types(dataset_id: str) -> list[dict[str, str]]:
    params = {
        "time": "latest",
        "RecordLimit": "2000",
        "select": "GEOGRAPHY_NAME,GEOGRAPHY_TYPE,GEOGRAPHY_TYPECODE",
    }
    url = f"{NOMIS_BASE}/{dataset_id}.data.csv?{urllib.parse.urlencode(params)}"
    rows = csv_rows_from_url(url)

    seen: dict[tuple[str, str], dict[str, str]] = {}
    for row in rows:
        key = (row.get("GEOGRAPHY_TYPE", ""), row.get("GEOGRAPHY_TYPECODE", ""))
        if key not in seen:
            seen[key] = {
                "geography_type": row.get("GEOGRAPHY_TYPE", ""),
                "geography_typecode": row.get("GEOGRAPHY_TYPECODE", ""),
            }
    return list(seen.values())


def find_local_authority_type(dataset_id: str) -> str | None:
    candidates = discover_geography_types(dataset_id)
    ranked_keywords = (
        "local authority",
        "unitary authority",
        "london borough",
        "metropolitan district",
        "council area",
        "district",
    )
    for candidate in candidates:
        geography_type = candidate["geography_type"].lower()
        if any(keyword in geography_type for keyword in ranked_keywords):
            return candidate["geography_typecode"]
    return None


def paged_nomis_download(
    dataset_id: str,
    name: str,
    out_path: Path,
    params: dict[str, str],
    page_size: int = 1000,
    pause_seconds: float = 0.25,
) -> DownloadResult:
    rows_accumulated: list[dict[str, str]] = []
    offset = 0
    record_count: int | None = None

    while True:
        query = dict(params)
        query["RecordLimit"] = str(page_size)
        query["RecordOffset"] = str(offset)
        url = f"{NOMIS_BASE}/{dataset_id}.data.csv?{urllib.parse.urlencode(query)}"
        rows = csv_rows_from_url(url)

        if not rows:
            break

        rows_accumulated.extend(rows)
        if record_count is None:
            try:
                record_count = int(rows[0].get("RECORD_COUNT", "0") or "0")
            except ValueError:
                record_count = 0

        offset += len(rows)
        if len(rows) < page_size:
            break
        if record_count and offset >= record_count:
            break
        time.sleep(pause_seconds)

    if not rows_accumulated:
        return DownloadResult(name=name, status="no-data", notes=["No rows returned"])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows_accumulated[0].keys()))
        writer.writeheader()
        writer.writerows(rows_accumulated)

    note = f"{len(rows_accumulated):,} rows"
    if record_count:
        note = f"{note}; record_count={record_count:,}"
    return DownloadResult(name=name, status="ok", output=str(out_path.relative_to(ROOT)), notes=[note])


def download_binary_from_dataset_page(
    name: str,
    page_url: str,
    out_name: str,
    extensions: tuple[str, ...],
) -> DownloadResult:
    asset_url = first_matching_link(page_url, extensions)
    out_path = RAW_DIR / out_name
    write_bytes(out_path, fetch_bytes(asset_url))
    return DownloadResult(
        name=name,
        status="ok",
        output=str(out_path.relative_to(ROOT)),
        notes=[asset_url],
    )


def download_ons_awe() -> DownloadResult:
    page_url = (
        "https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/"
        "earningsandworkinghours/datasets/averageweeklyearnings/current"
    )
    return download_binary_from_dataset_page(
        name="ons_average_weekly_earnings",
        page_url=page_url,
        out_name="ons_average_weekly_earnings.csv",
        extensions=(".csv",),
    )


def download_ons_unem_claimant_vacancies() -> DownloadResult:
    page_url = (
        "https://www.ons.gov.uk/employmentandlabourmarket/peoplenotinwork/"
        "unemployment/datasets/claimantcountandvacanciesdataset/current"
    )
    return download_binary_from_dataset_page(
        name="ons_claimant_and_vacancies",
        page_url=page_url,
        out_name="ons_claimant_and_vacancies.csv",
        extensions=(".csv",),
    )


def download_ons_cc01_local_claimant() -> DownloadResult:
    page_url = (
        "https://www.ons.gov.uk/employmentandlabourmarket/peoplenotinwork/"
        "unemployment/datasets/claimantcountbyunitaryandlocalauthorityexperimental/current"
    )
    return download_binary_from_dataset_page(
        name="ons_cc01_local_claimant_count",
        page_url=page_url,
        out_name="ons_cc01_local_claimant_count.xls",
        extensions=(".xls", ".xlsx"),
    )


def download_ons_m01_modelled_unemployment() -> DownloadResult:
    page_url = (
        "https://www.ons.gov.uk/employmentandlabourmarket/peoplenotinwork/"
        "unemployment/datasets/modelledunemploymentforlocalandunitaryauthoritiesm01/current"
    )
    return download_binary_from_dataset_page(
        name="ons_m01_modelled_unemployment",
        page_url=page_url,
        out_name="ons_m01_modelled_unemployment.xls",
        extensions=(".xls", ".xlsx"),
    )


def download_ons_ashe_table8() -> DownloadResult:
    page_url = (
        "https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/"
        "earningsandworkinghours/datasets/placeofresidencebylocalauthorityashetable8"
    )
    asset_url = first_matching_link(page_url, (".zip",))
    zip_bytes = fetch_bytes(asset_url)
    target_name_fragment = "Table 8.1a   Weekly pay - Gross"
    output_path = RAW_DIR / "ons_ashe_table8_weekly_pay_gross_2025.xlsx"

    with ZipFile(io.BytesIO(zip_bytes), "r") as archive:
        chosen_member = next(
            (
                name
                for name in archive.namelist()
                if target_name_fragment.lower() in name.lower() and name.lower().endswith(".xlsx")
            ),
            None,
        )
        if not chosen_member:
            raise RuntimeError("Could not find the weekly pay gross workbook inside the ASHE zip")
        write_bytes(output_path, archive.read(chosen_member))

    return DownloadResult(
        name="ons_ashe_table8_local_earnings",
        status="ok",
        output=str(output_path.relative_to(ROOT)),
        notes=[asset_url, chosen_member],
    )


def main() -> int:
    ensure_dirs()

    tasks = [
        download_ons_awe,
        download_ons_unem_claimant_vacancies,
        download_ons_cc01_local_claimant,
        download_ons_m01_modelled_unemployment,
        download_ons_ashe_table8,
    ]

    results: list[DownloadResult] = []
    for task in tasks:
        try:
            result = task()
        except Exception as exc:  # noqa: BLE001
            result = DownloadResult(
                name=task.__name__,
                status="error",
                notes=[str(exc)],
            )
        results.append(result)
        log(f"{result.name}: {result.status}")
        if result.output:
            log(f"  output: {result.output}")
        if result.notes:
            for note in result.notes:
                log(f"  note: {note}")

    failures = [result for result in results if result.status in {"error", "skipped"}]
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())

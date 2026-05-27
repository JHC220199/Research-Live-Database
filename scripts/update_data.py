"""
NRLA Data Dashboard — Automated Data Updater
=============================================
Fetches the latest statistics from official UK sources and updates the JSON
data files in /data/. Designed to run daily via GitHub Actions.
 
Sources:
  - ONS website JSON endpoint  (CPI, CPIH, GDP, earnings, PIPR)
  - Bank of England IADB       (base rate, mortgage advances)
  - MHCLG Live Table 213       (dwelling completions — England)
  - Stats Wales                (dwelling completions — Wales)
  - MoJ / HMCTS CSV zip        (landlord possession statistics)
  - EHS / FRS collection pages (annual survey release detection)
"""
 
import csv
import io
import json
import os
import re
import sys
import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path
 
import requests
from bs4 import BeautifulSoup
 
# ── Paths ─────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
 
# ── Shared HTTP session ───────────────────────────────────────────────────────
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "NRLA-Data-Dashboard/1.0 (github.com/nrla; data@nrla.org.uk)"
})
 
# ── Logging helpers ───────────────────────────────────────────────────────────
def log(msg):   print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
def warn(msg):  print(f"[{datetime.now().strftime('%H:%M:%S')}] ⚠  {msg}", file=sys.stderr)
def error(msg): print(f"[{datetime.now().strftime('%H:%M:%S')}] ✖  {msg}", file=sys.stderr)
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  ONS TIME SERIES
# ═══════════════════════════════════════════════════════════════════════════════
 
# Full ONS website paths for each CDID code.
# The data endpoint is: https://www.ons.gov.uk/{path}/data  → returns JSON
ONS_PATHS = {
    "D7G7": "economy/inflationandpriceindices/timeseries/d7g7",
    "D7BT": "economy/inflationandpriceindices/timeseries/d7bt",
    "L55O": "economy/inflationandpriceindices/timeseries/l55o",
    "L522": "economy/inflationandpriceindices/timeseries/l522",
    "IHYQ": "economy/grossdomesticproductgdp/timeseries/ihyq",
    "IHYR": "economy/grossdomesticproductgdp/timeseries/ihyr",
    "K54U": "employmentandlabourmarket/peopleinwork/earningsandworkinghours/timeseries/k54u",
    "A2FD": "employmentandlabourmarket/peopleinwork/earningsandworkinghours/timeseries/a2fd",
    "A3WV": "employmentandlabourmarket/peopleinwork/earningsandworkinghours/timeseries/a3wv",
}
 
 
def fetch_ons_timeseries(cdid: str) -> dict | None:
    """
    Fetch the latest value for an ONS time series using the website's JSON endpoint.
    Returns {"value": "...", "period": "..."} or None on failure.
    """
    path = ONS_PATHS.get(cdid.upper())
    if not path:
        warn(f"ONS {cdid}: no path configured")
        return None
 
    url = f"https://www.ons.gov.uk/{path}/data"
    try:
        r = SESSION.get(url, timeout=20)
        r.raise_for_status()
        data = r.json()
 
        # Try months first (most recent for monthly series), then quarters, then years
        for period_type in ("months", "quarters", "years"):
            entries = data.get(period_type, [])
            if entries:
                latest = entries[-1]
                raw_date = latest.get("date", latest.get("label", ""))
                # Format monthly dates: "2026 JAN" → "Jan 2026"
                formatted = _format_ons_date(raw_date)
                return {"value": latest.get("value"), "period": formatted}
 
        warn(f"ONS {cdid}: no period data in response")
        return None
 
    except Exception as e:
        error(f"ONS {cdid}: {e}")
        return None
 
 
def _format_ons_date(raw: str) -> str:
    """Convert ONS date strings like '2026 JAN' or '2025 Q3' to readable form."""
    raw = raw.strip()
    # Monthly: "2026 JAN"
    m = re.match(r"(\d{4})\s+([A-Z]{3})$", raw)
    if m:
        try:
            dt = datetime.strptime(f"{m.group(2)} {m.group(1)}", "%b %Y")
            return dt.strftime("%b %Y")
        except ValueError:
            pass
    # Quarterly: "2025 Q3"
    m = re.match(r"(\d{4})\s+(Q[1-4])$", raw)
    if m:
        return f"{m.group(2)} {m.group(1)}"
    return raw
 
 
def fetch_ons_pipr_uk() -> dict | None:
    """
    Fetch the latest PIPR (Price Index of Private Rents) annual rate for the UK.
    Scrapes the headline figure from the ONS bulletin HTML page.
    """
    bulletin_url = (
        "https://www.ons.gov.uk/economy/inflationandpriceindices/bulletins/"
        "privaterentalpricegreatbritain/latest"
    )
    try:
        r = SESSION.get(bulletin_url, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
 
        # ONS bulletin pages embed key stats in <p> or headline sections
        # Look for text like "annual percentage change of X%" or "rose by X%"
        full_text = soup.get_text(" ", strip=True)
 
        # Try to find the UK annual rate — typically stated as "X%" in opening paragraph
        patterns = [
            r"United Kingdom.*?(\d+\.?\d*)\s*%",
            r"UK.*?annual.*?(\d+\.?\d*)\s*%",
            r"(\d+\.?\d*)\s*%.*?annual.*?(?:change|rate)",
            r"rose by (\d+\.?\d*)\s*%",
            r"increased by (\d+\.?\d*)\s*%",
            r"(\d+\.?\d*)\s*%\s+in the 12 months",
        ]
        for pattern in patterns:
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                val = match.group(1)
                # Get the period from the page title or URL
                title = soup.find("title")
                period = "Latest"
                if title:
                    date_m = re.search(r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})", title.text)
                    if date_m:
                        period = f"{date_m.group(1)[:3]} {date_m.group(2)}"
                return {"value": val, "period": period}
 
        warn("PIPR: could not extract headline rate from bulletin")
        return None
 
    except Exception as e:
        error(f"PIPR: {e}")
        return None
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  BANK OF ENGLAND
# ═══════════════════════════════════════════════════════════════════════════════
 
def fetch_boe_series(series_code: str, label: str) -> dict | None:
    """
    Fetch the latest value for a Bank of England statistical series via IADB CSV.
    Handles both monthly and quarterly date formats.
    """
    url = (
        f"https://www.bankofengland.co.uk/boeapps/iadb/fromshowcolumns.asp"
        f"?csv.x=yes&Datefrom=01/Jan/2022&Dateto=now"
        f"&SeriesCodes={series_code}&CSVF=TT&UsingCodes=Y"
    )
    try:
        r = SESSION.get(url, timeout=25)
        r.raise_for_status()
        lines = [l.strip() for l in r.text.splitlines() if l.strip()]
 
        data_lines = []
        for line in lines:
            parts = [p.strip() for p in line.split(",")]
            if len(parts) >= 2:
                # Match both "01 Jan 2024" (monthly) and "01 Jan 2024" style dates
                if re.match(r"\d{2}\s+\w+\s+\d{4}", parts[0]):
                    data_lines.append(parts)
 
        if not data_lines:
            warn(f"BoE {series_code}: no data rows found")
            return None
 
        latest = data_lines[-1]
        date_str = latest[0]
        value = latest[1] if len(latest) > 1 else None
 
        if not value or value in ("", ".", ".."):
            # Sometimes last row is incomplete — try second-to-last
            if len(data_lines) > 1:
                latest = data_lines[-2]
                date_str = latest[0]
                value = latest[1] if len(latest) > 1 else None
 
        try:
            dt = datetime.strptime(date_str, "%d %b %Y")
            period = dt.strftime("%b %Y")
        except ValueError:
            period = date_str
 
        return {"value": value, "period": period}
 
    except Exception as e:
        error(f"BoE {series_code} ({label}): {e}")
        return None
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  MHCLG LIVE TABLE 213
# ═══════════════════════════════════════════════════════════════════════════════
 
def fetch_mhclg_table213() -> dict | None:
    """
    Download MHCLG Live Table 213 (permanent dwellings completed, England).
    Scrapes the gov.uk page to find the current Excel download URL.
    """
    page_url = "https://www.gov.uk/government/statistical-data-sets/live-tables-on-house-building"
    try:
        import openpyxl
 
        r = SESSION.get(page_url, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
 
        # Find the Table 213 Excel download link
        file_url = None
        for link in soup.find_all("a", href=True):
            href = link["href"]
            text = link.get_text(strip=True)
            if (re.search(r"table.{0,5}213", text, re.I) or
                    re.search(r"table.{0,5}213|LiveTable213", href, re.I)):
                if re.search(r"\.(xlsx?)$", href, re.I):
                    file_url = href if href.startswith("http") else "https://www.gov.uk" + href
                    break
 
        if not file_url:
            # Broader search in assets URLs
            for link in soup.find_all("a", href=re.compile(r"LiveTable213|table213|Table213", re.I)):
                file_url = link["href"]
                if not file_url.startswith("http"):
                    file_url = "https://www.gov.uk" + file_url
                break
 
        if not file_url:
            warn("MHCLG Table 213: download link not found on page")
            return None
 
        log(f"  MHCLG Table 213: {file_url}")
        xr = SESSION.get(file_url, timeout=60)
        xr.raise_for_status()
 
        wb = openpyxl.load_workbook(io.BytesIO(xr.content), data_only=True)
 
        # Find the "213" worksheet
        ws = None
        for name in wb.sheetnames:
            if "213" in name:
                ws = wb[name]
                break
        if ws is None:
            ws = wb.active
 
        rows = list(ws.iter_rows(values_only=True))
 
        # Find header row containing tenure column labels
        col_map = {}
        header_idx = None
        for i, row in enumerate(rows):
            row_text = " ".join(str(c).lower() for c in row if c)
            if "private" in row_text and ("total" in row_text or "housing" in row_text):
                header_idx = i
                for j, cell in enumerate(row):
                    if not cell:
                        continue
                    s = str(cell).lower()
                    if "private" in s and "enterprise" in s:
                        col_map["private"] = j
                    elif "housing assoc" in s or "registered" in s or " rsl" in s:
                        col_map["ha"] = j
                    elif "local auth" in s:
                        col_map["la"] = j
                    elif s.strip() == "total":
                        col_map["total"] = j
                if col_map:
                    break
 
        if not header_idx or not col_map:
            warn("MHCLG Table 213: could not identify header columns")
            return None
 
        # Find the last quarterly data row
        result = {}
        for row in reversed(rows[header_idx + 1:]):
            if not row[0]:
                continue
            period_str = str(row[0]).strip()
            # Quarterly format: "2025 Q3" or "2025Q3" or "Q3 2025"
            if not re.search(r"Q[1-4]", period_str, re.I):
                continue
            for key, col_idx in col_map.items():
                if col_idx < len(row) and row[col_idx] is not None:
                    try:
                        result[key] = {
                            "value": str(int(float(str(row[col_idx]).replace(",", "")))),
                            "period": period_str
                        }
                    except (ValueError, TypeError):
                        pass
            if result:
                break
 
        return result if result else None
 
    except Exception as e:
        error(f"MHCLG Table 213: {e}")
        return None
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  STATS WALES
# ═══════════════════════════════════════════════════════════════════════════════
 
def fetch_stats_wales_completions() -> dict | None:
    """
    Fetch new dwelling completions from Stats Wales.
    """
    # Try the Stats Wales OData API first
    api_url = (
        "https://statswales.gov.wales/api/v1/dataset/"
        "hous0302/data?$top=10&$orderby=Year_ItemName_ENG desc"
    )
    try:
        r = SESSION.get(api_url, timeout=20)
        if r.status_code == 200:
            data = r.json()
            rows = data.get("value", [])
            if rows:
                # Find total completions (all tenures)
                for row in rows:
                    item = row.get("Tenure_ItemName_ENG", "").lower()
                    if "all" in item or "total" in item:
                        return {
                            "total": {
                                "value": str(int(row.get("Data", 0))),
                                "period": row.get("Year_ItemName_ENG", "")
                            }
                        }
    except Exception:
        pass
 
    # Fallback: scrape the Stats Wales catalogue page
    try:
        r = SESSION.get(
            "https://statswales.gov.wales/Catalogue/Housing/New-House-Building",
            timeout=20
        )
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        tables = soup.find_all("table")
        for table in tables:
            for row in reversed(table.find_all("tr")):
                cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                if len(cells) >= 2 and re.search(r"\d{4}", cells[0]):
                    try:
                        val = int(cells[-1].replace(",", ""))
                        if val > 0:
                            return {"total": {"value": str(val), "period": cells[0]}}
                    except (ValueError, IndexError):
                        continue
    except Exception as e:
        warn(f"Stats Wales fallback: {e}")
 
    return None
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  MoJ LANDLORD POSSESSION STATISTICS
# ═══════════════════════════════════════════════════════════════════════════════
 
def fetch_moj_possession() -> dict | None:
    """
    Fetch landlord possession statistics from MoJ / HMCTS.
 
    Strategy:
    1. Fetch the collection page to find the most recent publication URL
    2. Visit that publication page to find the CSV zip download link
    3. Download, extract, and parse the relevant CSV files
    """
    collection_url = (
        "https://www.gov.uk/government/collections/"
        "mortgage-and-landlord-possession-statistics"
    )
    try:
        r = SESSION.get(collection_url, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
 
        # Find the first (most recent) publication link in the Documents section
        pub_url = None
        for link in soup.find_all("a", href=re.compile(
            r"/government/statistics/mortgage-and-landlord-possession-statistics-(?!earlier|guide)"
        )):
            href = link["href"]
            if not href.startswith("http"):
                href = "https://www.gov.uk" + href
            pub_url = href
            break
 
        if not pub_url:
            warn("MoJ: could not find latest publication link")
            return None
 
        log(f"  MoJ: latest publication → {pub_url}")
 
        # Visit the publication page and find the CSV zip
        pr = SESSION.get(pub_url, timeout=20)
        pr.raise_for_status()
        psoup = BeautifulSoup(pr.text, "html.parser")
 
        csv_zip_url = None
        for link in psoup.find_all("a", href=re.compile(r"CSVs\.zip", re.I)):
            csv_zip_url = link["href"]
            break
 
        if not csv_zip_url:
            # Fallback: look for any zip file
            for link in psoup.find_all("a", href=re.compile(r"\.zip$", re.I)):
                csv_zip_url = link["href"]
                break
 
        if not csv_zip_url:
            warn("MoJ: CSV zip download link not found on publication page")
            return None
 
        log(f"  MoJ: downloading CSV zip from {csv_zip_url}")
        zr = SESSION.get(csv_zip_url, timeout=60)
        zr.raise_for_status()
 
        return _parse_moj_csv_zip(zr.content)
 
    except Exception as e:
        error(f"MoJ possession: {e}")
        return None
 
 
def _parse_moj_csv_zip(zip_content: bytes) -> dict | None:
    """Parse the MoJ CSV zip and extract possession statistics."""
    result = {}
 
    try:
        z = zipfile.ZipFile(io.BytesIO(zip_content))
        names = z.namelist()
        log(f"  MoJ zip contents: {names}")
 
        for name in names:
            name_lower = name.lower()
 
            # Table 4: Claims issued and bailiff repossessions
            if "table_4" in name_lower or "table4" in name_lower or "_4_" in name_lower:
                rows = _read_csv_from_zip(z, name)
                r4 = _moj_latest_quarterly_row(rows)
                if r4:
                    period = r4.get("period", "")
                    numerics = r4.get("values", [])
                    if len(numerics) >= 1:
                        result["claims_issued"] = {"value": numerics[0], "period": period}
                    if len(numerics) >= 4:
                        result["repossessions_bailiffs"] = {"value": numerics[-1], "period": period}
 
            # Table 6a: Mean/median time
            elif "table_6" in name_lower or "table6" in name_lower or "_6a" in name_lower:
                rows = _read_csv_from_zip(z, name)
                r6 = _moj_latest_quarterly_row(rows)
                if r6:
                    period = r6.get("period", "")
                    numerics = r6.get("values", [])
                    if len(numerics) >= 1:
                        result["mean_time_all"] = {"value": numerics[0], "period": period}
                    if len(numerics) >= 2:
                        result["median_time_all"] = {"value": numerics[1], "period": period}
 
            # Table 7: Accelerated/private/social claims
            elif "table_7" in name_lower or "table7" in name_lower or "_7_" in name_lower:
                rows = _read_csv_from_zip(z, name)
                r7 = _moj_latest_quarterly_row(rows)
                if r7:
                    period = r7.get("period", "")
                    numerics = r7.get("values", [])
                    if len(numerics) >= 1:
                        result["claims_prs"] = {"value": numerics[0], "period": period}
                    if len(numerics) >= 2:
                        result["claims_accelerated"] = {"value": numerics[1], "period": period}
 
    except Exception as e:
        error(f"MoJ CSV parse: {e}")
        return None
 
    return result if result else None
 
 
def _read_csv_from_zip(z: zipfile.ZipFile, name: str) -> list[list[str]]:
    """Read a CSV file from a zip archive, returning a list of rows."""
    with z.open(name) as f:
        content = f.read().decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(content))
        return list(reader)
 
 
def _moj_latest_quarterly_row(rows: list) -> dict | None:
    """
    Find the most recent quarterly row in a MoJ CSV table.
    Returns {"period": "...", "values": [numeric strings]} or None.
    """
    for row in reversed(rows):
        if not row:
            continue
        period_cell = row[0].strip()
        # Quarter format: "2026 Q1", "2025 Q4", "Q1 2026", etc.
        if re.search(r"Q[1-4]", period_cell, re.I):
            numerics = []
            for cell in row[1:]:
                try:
                    val = str(int(float(cell.replace(",", "").strip())))
                    numerics.append(val)
                except (ValueError, AttributeError):
                    pass
            if numerics:
                return {"period": period_cell, "values": numerics}
    return None
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  ANNUAL RELEASE DETECTION (EHS / FRS)
# ═══════════════════════════════════════════════════════════════════════════════
 
def find_latest_annual_release(collection_url: str, publication_pattern: str,
                               label: str) -> tuple[str | None, str | None]:
    """
    Scrape a gov.uk collection page to find the most recent annual release year.
    Returns (publication_url, year_string) or (None, None).
    """
    try:
        r = SESSION.get(collection_url, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
 
        year_pattern = re.compile(r"20\d\d")
        pub_links = soup.find_all("a", href=re.compile(publication_pattern, re.I))
 
        best_link = None
        best_year = 0
        for link in pub_links:
            years = year_pattern.findall(link.get("href", "") + link.get_text())
            for y in years:
                if int(y) > best_year:
                    best_year = int(y)
                    best_link = link
 
        if best_link:
            href = best_link["href"]
            if not href.startswith("http"):
                href = "https://www.gov.uk" + href
            log(f"  {label}: latest release → {href} ({best_year})")
            return href, str(best_year)
 
        warn(f"{label}: no release link found at {collection_url}")
        return None, None
 
    except Exception as e:
        error(f"{label} collection page: {e}")
        return None, None
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  JSON FILE HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
 
def load_json(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)
 
 
def save_json(path: Path, data: dict) -> bool:
    """Write JSON. Returns True if content changed."""
    new_content = json.dumps(data, indent=2, ensure_ascii=False)
    if path.exists():
        old_content = path.read_text(encoding="utf-8")
        if old_content == new_content:
            return False
    path.write_text(new_content, encoding="utf-8")
    return True
 
 
def now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
 
 
def update_dataset(datasets: list, dataset_id: str, fetched: dict | None) -> bool:
    """Update a dataset entry in-place. Returns True if changed."""
    if not fetched or fetched.get("value") is None:
        return False
    for ds in datasets:
        if ds["id"] == dataset_id:
            old = ds.get("latest", {})
            new_val = {
                "value": str(fetched["value"]),
                "period": fetched.get("period", old.get("period")),
                "fetched_at": now_iso()
            }
            if old.get("value") != new_val["value"] or old.get("period") != new_val["period"]:
                ds["latest"] = new_val
                return True
    return False
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION RUNNERS
# ═══════════════════════════════════════════════════════════════════════════════
 
def update_macro():
    log("=== Updating Macro-Economic data ===")
    path = DATA_DIR / "macro.json"
    data = load_json(path)
    ds = data["datasets"]
    changed = False
 
    ons_cdid_map = {
        "cpi_annual_rate":      "D7G7",
        "cpi_index":            "D7BT",
        "cpih_annual_rate":     "L55O",
        "cpih_index":           "L522",
        "gdp_qoq":              "IHYQ",
        "gdp_yoy":              "IHYR",
        "awe_total_pay":        "K54U",
        "real_earnings_index":  "A2FD",
        "real_earnings_growth": "A3WV",
    }
 
    for dataset_id, cdid in ons_cdid_map.items():
        log(f"  ONS {cdid} → {dataset_id}")
        result = fetch_ons_timeseries(cdid)
        if update_dataset(ds, dataset_id, result):
            log(f"    Updated: {result}")
            changed = True
        else:
            log(f"    No change (fetched: {result})")
        time.sleep(0.3)
 
    # PIPR (rental price index)
    log("  PIPR annual rate")
    pipr = fetch_ons_pipr_uk()
    if update_dataset(ds, "pipr_annual_rate_uk", pipr):
        log(f"    Updated: {pipr}")
        changed = True
    if update_dataset(ds, "pipr_index_uk", pipr):
        changed = True
 
    # Bank of England base rate
    log("  BoE base rate (IUMABEDR)")
    boe_rate = fetch_boe_series("IUMABEDR", "base rate")
    if update_dataset(ds, "boe_base_rate", boe_rate):
        log(f"    Updated: {boe_rate}")
        changed = True
    else:
        log(f"    No change (fetched: {boe_rate})")
 
    # BoE total gross mortgage advances (Table 1.21)
    # LPMVQZHD = total gross advances, all lenders (quarterly)
    log("  BoE gross advances (LPMVQZHD)")
    boe_adv = fetch_boe_series("LPMVQZHD", "gross advances")
    if update_dataset(ds, "boe_gross_advances", boe_adv):
        log(f"    Updated: {boe_adv}")
        changed = True
    else:
        log(f"    No change (fetched: {boe_adv})")
 
    if changed:
        data["last_fetched"] = now_iso()
        save_json(path, data)
        log("  macro.json saved ✓")
    else:
        log("  No changes to macro.json")
 
 
def update_housing_tenure():
    log("=== Updating Housing Trends (FRS) ===")
    path = DATA_DIR / "housing_tenure.json"
    data = load_json(path)
 
    pub_url, year = find_latest_annual_release(
        "https://www.gov.uk/government/collections/family-resources-survey--2",
        r"family-resources-survey",
        "FRS"
    )
 
    if not year:
        log("  FRS: no new release detected")
        return
 
    changed = False
    for ds in data["datasets"]:
        if ds["status"] == "active":
            latest = ds.get("latest", {})
            stored_period = latest.get("period", "")
            expected_period = f"{int(year)-1}/{str(year)[2:]}"
            if year not in stored_period:
                ds["latest"] = {
                    "value": latest.get("value"),
                    "period": expected_period,
                    "fetched_at": now_iso()
                }
                changed = True
 
    if changed:
        data["last_fetched"] = now_iso()
        save_json(path, data)
        log("  housing_tenure.json updated ✓")
    else:
        log("  No changes to housing_tenure.json")
 
 
def update_dwellings():
    log("=== Updating Dwellings data ===")
    path = DATA_DIR / "dwellings.json"
    data = load_json(path)
    ds = data["datasets"]
    changed = False
 
    log("  MHCLG Table 213 (England completions)")
    mhclg = fetch_mhclg_table213()
    if mhclg:
        for key, dst_id in [
            ("private", "mhclg_completions_private"),
            ("ha",      "mhclg_completions_ha"),
            ("la",      "mhclg_completions_la"),
            ("total",   "mhclg_completions_total"),
        ]:
            if key in mhclg and update_dataset(ds, dst_id, mhclg[key]):
                log(f"    {dst_id}: {mhclg[key]}")
                changed = True
 
    log("  Stats Wales (Wales completions)")
    wales = fetch_stats_wales_completions()
    if wales:
        if "total" in wales and update_dataset(ds, "stats_wales_completions_total", wales["total"]):
            log(f"    stats_wales_completions_total: {wales['total']}")
            changed = True
 
    if changed:
        data["last_fetched"] = now_iso()
        save_json(path, data)
        log("  dwellings.json saved ✓")
    else:
        log("  No changes to dwellings.json")
 
 
def update_possession():
    log("=== Updating Claims & Bailiffs (MoJ) ===")
    path = DATA_DIR / "possession.json"
    data = load_json(path)
    ds = data["datasets"]
 
    moj = fetch_moj_possession()
    if not moj:
        log("  No MoJ data retrieved")
        return
 
    changed = False
    mapping = {
        "claims_issued":          "mlp_claims_issued",
        "repossessions_bailiffs": "mlp_repossessions_bailiffs",
        "claims_prs":             "mlp_claims_prs",
        "claims_accelerated":     "mlp_claims_accelerated",
        "mean_time_all":          "mlp_mean_time_all",
        "median_time_all":        "mlp_median_time_all",
    }
    for moj_key, ds_id in mapping.items():
        if moj_key in moj and update_dataset(ds, ds_id, moj[moj_key]):
            log(f"  {ds_id}: {moj[moj_key]}")
            changed = True
 
    if changed:
        data["last_fetched"] = now_iso()
        save_json(path, data)
        log("  possession.json saved ✓")
    else:
        log("  No changes to possession.json")
 
 
def update_ehs():
    log("=== Updating English Housing Survey ===")
    path = DATA_DIR / "ehs.json"
    data = load_json(path)
 
    pub_url, year = find_latest_annual_release(
        "https://www.gov.uk/government/collections/english-housing-survey",
        r"english-housing-survey",
        "EHS"
    )
 
    if not year:
        log("  EHS: no new release detected")
        return
 
    changed = False
    for ds in data["datasets"]:
        if ds["status"] == "active":
            latest = ds.get("latest", {})
            if year not in str(latest.get("period", "")):
                ds["latest"] = {
                    "value": latest.get("value"),
                    "period": year,
                    "fetched_at": now_iso()
                }
                changed = True
 
    if changed:
        data["last_fetched"] = now_iso()
        save_json(path, data)
        log("  ehs.json updated ✓")
    else:
        log("  No changes to ehs.json")
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════
 
def main():
    log(f"NRLA Data Dashboard updater — {datetime.now().strftime('%Y-%m-%d %H:%M UTC')}")
    log(f"Data directory: {DATA_DIR}")
 
    if not DATA_DIR.exists():
        error(f"Data directory not found: {DATA_DIR}")
        sys.exit(1)
 
    errors = []
    for label, fn in [
        ("Macro-Economic",  update_macro),
        ("Housing Tenure",  update_housing_tenure),
        ("Dwellings",       update_dwellings),
        ("Claims/Bailiffs", update_possession),
        ("EHS",             update_ehs),
    ]:
        try:
            fn()
        except Exception as e:
            msg = f"{label}: unexpected error — {e}"
            error(msg)
            errors.append(msg)
 
    if errors:
        log(f"\nCompleted with {len(errors)} error(s):")
        for e in errors:
            log(f"  ✖ {e}")
        sys.exit(1)
    else:
        log("\nAll sections updated successfully ✓")
 
 
if __name__ == "__main__":
    main()
 

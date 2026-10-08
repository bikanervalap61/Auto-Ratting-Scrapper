import os
import re
import sys
import json
import time
import random
import calendar
import urllib.parse
import urllib.request
from datetime import datetime, date, timezone, timedelta
from pathlib import Path

IST = timezone(timedelta(hours=5, minutes=30))

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent.resolve()
EXCEL_PATH = BASE_DIR / "Outlet name and Link.xlsx" if (BASE_DIR / "Outlet name and Link.xlsx").exists() else BASE_DIR / "Bvgg GMB Rating Sheet (2).xlsx"
CACHE_PATH = BASE_DIR / "discovered_links.json"

# Header styling
date_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
rate_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
rev_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
closed_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
closed_font = Font(name="Calibri", size=10, italic=True, color="7F7F7F")


# ── Link Cache Manager ───────────────────────────────────────────
def load_cached_links():
    if CACHE_PATH.exists():
        try:
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_cached_links(cache):
    try:
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"[Warning] Could not save link cache: {e}")


def resolve_shortlink(url):
    """
    Expands share.google, goo.gl, maps.app.goo.gl shortlinks via HTTP redirect.
    Converts Google Search redirects (kgmid) into direct Google Maps queries
    to completely avoid Google Web Search bot captchas.
    """
    if not url or not str(url).startswith("http"):
        return None

    if "share.google" in url or "goo.gl" in url:
        try:
            req = urllib.request.Request(
                url,
                headers={
                    "User-Agent": (
                        "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
                        "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1"
                    )
                }
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                resolved = resp.geturl()
                if "/search" in resolved:
                    q_m = re.search(r"[?&]q=([^&]+)", resolved)
                    if q_m:
                        query_str = urllib.parse.unquote_plus(q_m.group(1))
                        return f"https://www.google.com/maps/search/{urllib.parse.quote(query_str)}?hl=en"
                return resolved
        except Exception:
            pass

    return url


# ── GMB Data Extractor ───────────────────────────────────────────
def extract_gmb_data(page, brand="Bikanervala"):
    """
    Extracts live rating and total reviews from Google Maps.
    Uses mobile / compact layout selectors for 100% reliable review capture.
    """
    rating, reviews = None, None

    try:
        html = page.content()

        # Method 1: Combined aria-label like '4.2 stars 765 Reviews' or '3.9 stars 3,843 reviews'
        m_comb = re.search(r'aria-label="([\d.]+)\s*stars?\s*([\d,]+)\s*reviews?"', html, re.I)
        if m_comb:
            rating = float(m_comb.group(1))
            reviews = int(m_comb.group(2).replace(",", ""))

        # Method 2: Separate aria-label like '3.9 stars ' and '1,003 reviews'
        if rating is None:
            m_r = re.search(r'aria-label="([\d.]+)\s*stars?[^"]*"', html, re.I)
            if m_r:
                rating = float(m_r.group(1))

        if reviews is None:
            m_rev = re.search(r'aria-label="([\d,]+)\s*reviews?[^"]*"', html, re.I)
            if m_rev:
                reviews = int(m_rev.group(1).replace(",", ""))

        # Method 3: Inner text pattern like '3.9\n(1,003)' or '4.2(765)'
        if rating is None or reviews is None:
            m_txt = re.search(r'([\d.]+)\s*(?:stars?|★)?\s*\(([\d,]+)\)', html)
            if m_txt:
                if rating is None:
                    rating = float(m_txt.group(1))
                if reviews is None:
                    reviews = int(m_txt.group(2).replace(",", ""))

        # Method 4: Check general stars & reviews in text
        if rating is None or reviews is None:
            m_txt2 = re.search(r'([\d.]+)\s*stars?.*?([\d,]+)\s*reviews?', html, re.I | re.DOTALL)
            if m_txt2:
                if rating is None:
                    rating = float(m_txt2.group(1))
                if reviews is None:
                    reviews = int(m_txt2.group(2).replace(",", ""))

        # Method 5: Check K format like '3.8K reviews' or '3.8 K'
        if reviews is None:
            m_k = re.search(r'([\d.]+)\s*[Kk]\s*reviews?', html, re.I)
            if m_k:
                reviews = int(float(m_k.group(1)) * 1000)

    except Exception:
        pass

    return rating, reviews


def scrape_store(page, store_code, brand, address, direct_url=None, ignore_store_code=False):
    clean_addr = str(address).strip()
    if "(closed)" in clean_addr.lower() or "(closed)" in str(store_code).lower():
        print(f"  [{store_code}] Outlet is marked as Closed in sheet.")
        return {
            "store_code": store_code,
            "address": address,
            "rating": "Closed",
            "reviews": "Closed",
            "status": "closed",
            "is_closed": True
        }

    cached_links = load_cached_links()

    # In Comparison Sheet, ignore cluster store_code to prevent Haldiram/Om from copying Bikanervala
    url = direct_url
    if not url:
        url = cached_links.get(clean_addr.lower())
    if not url:
        clean_norm = re.sub(r"\(.*?\)", "", clean_addr).strip().lower()
        clean_norm = re.sub(r"\s+", " ", clean_norm)
        url = cached_links.get(clean_norm)
    if not url and not ignore_store_code and store_code and not str(store_code).startswith("#"):
        url = cached_links.get(str(store_code).strip().upper())

    discovered_url = None
    rating, reviews = None, None

    target_url = resolve_shortlink(url)
    if target_url and target_url.startswith("http"):
        try:
            page.goto(target_url, timeout=30000, wait_until="domcontentloaded")
            page.wait_for_timeout(3500)
            rating, reviews = extract_gmb_data(page, brand=brand)
            if "/maps/place/" in page.url:
                discovered_url = page.url
        except Exception as e:
            print(f"    [Direct URL warning] {e}")

    # Fallback to search query if not retrieved
    if rating is None or reviews is None:
        clean_name = re.sub(r"[,\-_/]+", " ", clean_addr).strip()
        name_no_bv = re.sub(r"^bv\s+", "", clean_name, flags=re.I).strip()
        q_brand = "" if brand.lower() in name_no_bv.lower() else f"{brand} "
        search_query = urllib.parse.quote(f"{q_brand}{name_no_bv}")
        search_url = f"https://www.google.com/maps/search/{search_query}?hl=en"
        try:
            page.goto(search_url, timeout=30000, wait_until="domcontentloaded")
            page.wait_for_timeout(4000)
            f_rating, f_reviews = extract_gmb_data(page, brand=brand)
            if f_rating is not None:
                rating = f_rating
            if f_reviews is not None:
                reviews = f_reviews
            if "/maps/place/" in page.url:
                discovered_url = page.url
        except Exception as e:
            print(f"    [Search warning] {e}")

    if rating is not None and reviews is not None:
        status = "success"
    elif rating is not None or reviews is not None:
        status = "partial"
    else:
        status = "failed"

    # Save newly discovered link permanently under store name
    if discovered_url:
        cached_links[clean_addr.lower()] = discovered_url
        if not ignore_store_code and store_code and not str(store_code).startswith("#"):
            cached_links[str(store_code).strip().upper()] = discovered_url
        save_cached_links(cached_links)

    print(f"  -> Result: {rating} | Reviews: {reviews}")
    return {
        "store_code": store_code,
        "address": address,
        "rating": rating,
        "reviews": reviews,
        "status": status,
        "is_closed": False
    }


# ── Monthly Gurgaon Sheet Manager ────────────────────────────────
def get_or_create_monthly_sheet(wb, today):
    """
    Finds the active monthly Gurgaon sheet (e.g. 'Gurgaon  sept').
    If today belongs to a new month (e.g. October), it automatically clones the sheet,
    sets days (1st to 30th/31st), clears daily cells, and starts from 1st of the month!
    """
    month_abbr = today.strftime("%b").lower()  # 'sep', 'oct', etc.
    target_name = f"Gurgaon  {month_abbr}"

    # Check if target sheet already exists
    for sname in wb.sheetnames:
        if month_abbr in sname.lower() and "gurgaon" in sname.lower():
            return wb[sname]

    # Look for any recent Gurgaon sheet to clone
    source_sheet = None
    for sname in reversed(wb.sheetnames):
        if "gurgaon" in sname.lower() or "sept" in sname.lower():
            source_sheet = wb[sname]
            break

    if not source_sheet:
        source_sheet = wb.active

    print(f"\n[Month Rollover] Creating new monthly sheet '{target_name}' by cloning '{source_sheet.title}'...")
    new_sheet = wb.copy_worksheet(source_sheet)
    new_sheet.title = target_name

    _, num_days = calendar.monthrange(today.year, today.month)

    # Locate Summary header column
    summary_col = None
    for c in range(9, 55):
        val = str(new_sheet.cell(row=1, column=c).value or "")
        if "summary" in val.lower():
            summary_col = c
            break
    if not summary_col:
        summary_col = 40

    # Clear previous month daily data
    for r in range(3, 27):
        for c in range(9, summary_col):
            new_sheet.cell(row=r, column=c).value = None

    # Write Day 1 to num_days in Row 1
    col = 9
    for day in range(1, num_days + 1):
        dt = datetime(today.year, today.month, day)
        new_sheet.cell(row=1, column=col).value = dt
        new_sheet.cell(row=2, column=col).value = "Review Count"
        col += 1

    # Update Summary headers
    new_sheet.cell(row=1, column=col).value = "Summary"
    new_sheet.cell(row=2, column=col).value = f"{today.strftime('%b').title()} Review Count"
    new_sheet.cell(row=2, column=col + 1).value = "Average GMB Live Rating"

    print(f"[Month Rollover] Created sheet '{target_name}' with days 1 to {num_days}.")
    return new_sheet


def update_gurgaon_sheet(ws, page, today, all_results, end_row=26):
    print(f"\n==================================================")
    print(f"  1. Processing Gurgaon Monthly Sheet: {ws.title}")
    print(f"==================================================")

    # Find today's date column
    today_col = None
    target_date = datetime(today.year, today.month, today.day).date()

    for c in range(8, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if isinstance(v, (datetime, date)) and v.date() == target_date:
            today_col = c
            break

    if not today_col:
        for c in range(8, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if isinstance(v, (datetime, date)) and v.day == today.day and v.month == today.month:
                today_col = c
                break

    if not today_col:
        print(f"  [Warning] Today's column for {target_date} not found. Using Col 38.")
        today_col = 38

    print(f"  Target column for {target_date}: Col {today_col} ({openpyxl.utils.get_column_letter(today_col)})")

    results = []
    for r in range(3, end_row + 1):
        sno = ws.cell(row=r, column=1).value
        code = str(ws.cell(row=r, column=2).value or "").strip()
        name = str(ws.cell(row=r, column=3).value or "").strip()

        if not name:
            continue

        is_explicit_closed = "(closed)" in name.lower() or "(closed)" in code.lower()
        print(f"[{sno}] {code} - {name}")
        if is_explicit_closed:
            print(f"  -> Marked as Closed in Excel.")
            res = {
                "sno": sno,
                "store_code": code or f"#{sno}",
                "address": name,
                "rating": "Closed",
                "reviews": "Closed",
                "status": "closed",
                "is_closed": True
            }
        else:
            res = scrape_store(page, code, "Bikanervala", name, ignore_store_code=False)
            res["sno"] = sno

        results.append(res)
        all_results.append(res)

        # Write to today's column (Reviews)
        c_cell = ws.cell(row=r, column=today_col)
        if res.get("is_closed"):
            c_cell.value = "Closed"
            c_cell.fill = closed_fill
            c_cell.font = closed_font
        elif res["reviews"] is not None:
            c_cell.value = res["reviews"]
        c_cell.alignment = Alignment(horizontal="center", vertical="center")

        if not is_explicit_closed:
            time.sleep(random.uniform(1.5, 3.0))

    return results


def find_last_data_row(ws, start_row=3):
    last = start_row
    for r in range(start_row, ws.max_row + 1):
        sno = ws.cell(row=r, column=1).value
        # Check col 2, 3, 4 for content
        c2 = ws.cell(row=r, column=2).value
        c3 = ws.cell(row=r, column=3).value
        c4 = ws.cell(row=r, column=4).value
        if sno is not None or c2 or c3 or c4:
            last = r
    return last


def update_dated_pair_sheet(ws, page, title, brand, name_col, code_col, start_row, end_row, all_results, test_mode=False):
    if end_row is None:
        end_row = (start_row + 1) if test_mode else find_last_data_row(ws, start_row)

    print(f"\n==================================================")
    print(f"  Processing: {title} ({end_row - start_row + 1} stores: Rows {start_row} to {end_row})")
    print(f"==================================================")

    today = datetime.now()
    today_date = datetime(today.year, today.month, today.day)

    # Dynamic column detection from Row 2 headers
    actual_name_col = name_col
    actual_code_col = code_col
    actual_link_col = None
    for col_idx in range(1, min(ws.max_column + 1, 10)):
        hdr = str(ws.cell(row=2, column=col_idx).value or "").strip().lower()
        if any(k in hdr for k in ["store address", "outlet name", "address", "outlet"]):
            actual_name_col = col_idx
        elif any(k in hdr for k in ["store code", "code"]) and "brand" not in hdr and "pair" not in hdr:
            actual_code_col = col_idx
        elif any(k in hdr for k in ["link", "url"]):
            actual_link_col = col_idx

    # Find today's column pair by scanning every column from col 5 onwards
    rate_col, rev_col = None, None
    for c in range(5, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if isinstance(v, (datetime, date)) and v.date() == today_date.date():
            hdr = str(ws.cell(row=2, column=c).value or "").lower()
            if "rating" in hdr:
                rate_col = c
                rev_col = c + 1
                break
            elif "review" in hdr:
                rev_col = c
                rate_col = c - 1
                break

    # Add new column pair if not present
    if not rate_col:
        rate_col = ws.max_column + 1
        rev_col = rate_col + 1
        print(f"  Adding today's column pair at Col {rate_col} & Col {rev_col}")

        ws.cell(row=1, column=rate_col, value=today_date).font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=1, column=rate_col).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=1, column=rate_col).fill = date_fill
        ws.cell(row=1, column=rev_col, value=today_date).font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=1, column=rev_col).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=1, column=rev_col).fill = date_fill

        c_r = ws.cell(row=2, column=rate_col, value="GMB Live Rating")
        c_r.font = Font(name="Calibri", size=9, bold=True)
        c_r.alignment = Alignment(horizontal="center", vertical="center")
        c_r.fill = rate_fill

        c_rev = ws.cell(row=2, column=rev_col, value="Review Count")
        c_rev.font = Font(name="Calibri", size=9, bold=True)
        c_rev.alignment = Alignment(horizontal="center", vertical="center")
        c_rev.fill = rev_fill

    results = []
    for r in range(start_row, end_row + 1):
        sno = ws.cell(row=r, column=1).value
        code = str(ws.cell(row=r, column=actual_code_col).value or "").strip() if actual_code_col else ""
        name = str(ws.cell(row=r, column=actual_name_col).value or "").strip()
        row_link = str(ws.cell(row=r, column=actual_link_col).value or "").strip() if actual_link_col else None
        if not row_link:
            row_link = None

        if not name:
            continue

        # Special closed check for India BV Row 81 S.No 79 or '(Closed)' in name
        is_explicit_closed = "(closed)" in name.lower() or (title == "India BV" and str(sno) == "79")

        # Dynamic brand detection from Column C or header
        row_brand = str(ws.cell(row=r, column=3).value or "").strip()
        if row_brand and "brand" not in row_brand.lower():
            store_brand = row_brand
        else:
            store_brand = "Haldiram" if name.upper().startswith("HR") else ("Om Sweets" if name.upper().startswith(("OM", "OM,")) else "Bikanervala")

        print(f"[{sno}] {name}")
        if is_explicit_closed:
            print(f"  -> Marked as Closed.")
            res = {
                "sno": sno,
                "store_code": code or f"#{sno}",
                "address": name,
                "rating": "Closed",
                "reviews": "Closed",
                "status": "closed",
                "is_closed": True
            }
            for c_past in [7, 8]:
                if ws.cell(row=r, column=c_past).value in [4.9, 283, "4.9", "283"]:
                    ws.cell(row=r, column=c_past).value = "Closed"
                    ws.cell(row=r, column=c_past).fill = closed_fill
                    ws.cell(row=r, column=c_past).font = closed_font
        else:
            # Always ignore cluster/index codes so every outlet uses its unique store name
            ignore_code = True
            res = scrape_store(page, code or f"#{sno}", store_brand, name, direct_url=row_link, ignore_store_code=ignore_code)
            res["sno"] = sno
            res["brand"] = store_brand

        results.append(res)
        all_results.append(res)

        c1 = ws.cell(row=r, column=rate_col)
        c2 = ws.cell(row=r, column=rev_col)

        if res.get("is_closed"):
            c1.value = "Closed"
            c2.value = "Closed"
            c1.fill = closed_fill
            c2.fill = closed_fill
            c1.font = closed_font
            c2.font = closed_font
        else:
            c1.value = res["rating"]
            c2.value = res["reviews"]

        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")

        if not is_explicit_closed:
            time.sleep(random.uniform(1.5, 3.0))

    return results


def safe_save_workbook(wb, path):
    try:
        wb.save(path)
        print(f"  [Auto-Saved] Successfully saved to {path.name}")
        return path
    except PermissionError:
        fallback = path.parent / f"{path.stem} - Updated{path.suffix}"
        print(f"\n  [NOTE] '{path.name}' is open in Excel.")
        print(f"  Saved to '{fallback.name}' instead so no data was lost.")
        wb.save(fallback)
        return fallback


def main():
    test_mode = "--test" in sys.argv
    if test_mode:
        print("[TEST MODE] Running 2 stores per sheet for fast verification...")

    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Master workbook not found at {EXCEL_PATH}")

    today = datetime.now()
    today_str = today.strftime("%d/%m/%Y")
    print(f"=== Starting GMB Scraper for Bvgg GMB Rating Sheet (2) | Date: {today_str} ===")

    wb = openpyxl.load_workbook(EXCEL_PATH)
    all_results = []
    cat_summaries = {}

    with sync_playwright() as pw:
        # Use Mobile Safari / iPhone device context for 100% reliable review count badges & zero captchas
        iphone = pw.devices["iPhone 14"]
        browser = pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(**iphone)
        page = context.new_page()

        # 1. Bikanervala GGN Sheet
        if "Bikanervala GGN" in wb.sheetnames:
            ws_ggn = wb["Bikanervala GGN"]
            res_ggn = update_dated_pair_sheet(
                ws_ggn, page, "Bikanervala GGN", brand="Bikanervala",
                name_col=4, code_col=2, start_row=3, end_row=4 if test_mode else None, all_results=all_results, test_mode=test_mode
            )
            safe_save_workbook(wb, EXCEL_PATH)
            cat_summaries["Bikanervala GGN"] = {"total": len(res_ggn), "results": res_ggn}
        else:
            ws_monthly = get_or_create_monthly_sheet(wb, today)
            res_ggn = update_gurgaon_sheet(ws_monthly, page, today, all_results, end_row=4 if test_mode else 26)
            safe_save_workbook(wb, EXCEL_PATH)
            cat_summaries["Bikanervala GGN"] = {"total": len(res_ggn), "results": res_ggn}

        # 2. Competitors Sheet
        comp_sheet_name = next((s for s in ["Competitors", "Comparison Sheet", "Gg Comparison Sheet"] if s in wb.sheetnames), None)
        if comp_sheet_name:
            ws_comp = wb[comp_sheet_name]
            res_comp = update_dated_pair_sheet(
                ws_comp, page, comp_sheet_name, brand="Auto",
                name_col=4, code_col=2, start_row=3, end_row=4 if test_mode else None, all_results=all_results, test_mode=test_mode
            )
            safe_save_workbook(wb, EXCEL_PATH)
            cat_summaries["Competitors Comparison"] = {"total": len(res_comp), "results": res_comp}

        # 3. India BV
        if "India BV" in wb.sheetnames:
            ws_ind = wb["India BV"]
            res_ind = update_dated_pair_sheet(
                ws_ind, page, "India BV", brand="Bikanervala",
                name_col=4, code_col=2, start_row=3, end_row=4 if test_mode else None, all_results=all_results, test_mode=test_mode
            )
            safe_save_workbook(wb, EXCEL_PATH)
            cat_summaries["India BV"] = {"total": len(res_ind), "results": res_ind}

        # 4. International BV (53 stores)
        if "International BV" in wb.sheetnames:
            ws_int = wb["International BV"]
            res_int = update_dated_pair_sheet(
                ws_int, page, "International BV", brand="Bikanervala",
                name_col=4, code_col=2, start_row=3, end_row=4 if test_mode else None, all_results=all_results, test_mode=test_mode
            )
            safe_save_workbook(wb, EXCEL_PATH)
            cat_summaries["International BV"] = {"total": len(res_int), "results": res_int}

        # 5. Haldiram India (111 stores)
        if "Haldiram India" in wb.sheetnames:
            ws_hr = wb["Haldiram India"]
            res_hr = update_dated_pair_sheet(
                ws_hr, page, "Haldiram India", brand="Haldiram",
                name_col=4, code_col=2, start_row=3, end_row=4 if test_mode else None, all_results=all_results, test_mode=test_mode
            )
            safe_save_workbook(wb, EXCEL_PATH)
            cat_summaries["Haldiram India"] = {"total": len(res_hr), "results": res_hr}

        # 6. OM Sweets India (22 stores)
        if "OM Sweets India" in wb.sheetnames:
            ws_om = wb["OM Sweets India"]
            res_om = update_dated_pair_sheet(
                ws_om, page, "OM Sweets India", brand="Om Sweets",
                name_col=4, code_col=2, start_row=3, end_row=4 if test_mode else None, all_results=all_results, test_mode=test_mode
            )
            safe_save_workbook(wb, EXCEL_PATH)
            cat_summaries["OM Sweets India"] = {"total": len(res_om), "results": res_om}

        # 7. Other Outlets (59 competitor stores)
        if "Other Outlets" in wb.sheetnames:
            ws_other = wb["Other Outlets"]
            res_other = update_dated_pair_sheet(
                ws_other, page, "Other Outlets", brand="Auto",
                name_col=4, code_col=2, start_row=3, end_row=4 if test_mode else None, all_results=all_results, test_mode=test_mode
            )
            safe_save_workbook(wb, EXCEL_PATH)
            cat_summaries["Other Outlets"] = {"total": len(res_other), "results": res_other}

        context.close()
        browser.close()

    saved_final = safe_save_workbook(wb, EXCEL_PATH)
    wb.close()
    print(f"\n[Completed] Excel file updated: {saved_final}")

    # Generate run summary JSON
    summary_data = {
        "date": today_str,
        "total_stores": len(all_results),
        "successful": sum(1 for x in all_results if x["status"] == "success"),
        "closed": sum(1 for x in all_results if x.get("is_closed")),
        "failed": sum(1 for x in all_results if x["status"] == "failed"),
        "categories": cat_summaries
    }
    with open(BASE_DIR / "run_summary.json", "w", encoding="utf-8") as f:
        json.dump(summary_data, f, indent=2, ensure_ascii=False)

    # Automatically sync to Google Sheet Webhook if configured
    webhook_url = os.getenv("GOOGLE_SHEET_WEBHOOK_URL")
    if webhook_url:
        try:
            import requests
            print(f"\n[Google Sheets] Syncing live data for {today_str} to Google Sheet Webhook...")
            g_payload = {
                "date": today_str,
                "categories": {k: v["results"] for k, v in cat_summaries.items() if "results" in v}
            }
            g_res = requests.post(webhook_url, json=g_payload, timeout=90)
            print(f"[Google Sheets] HTTP Status: {g_res.status_code}")
            print(f"[Google Sheets] Response: {g_res.text}")
        except Exception as gs_err:
            print(f"[Google Sheets ERROR] Could not sync to Google Sheet: {gs_err}")
    else:
        print("\n[Google Sheets WARNING] GOOGLE_SHEET_WEBHOOK_URL secret / environment variable is NOT set!")
        print("  Google Sheet was NOT updated because GOOGLE_SHEET_WEBHOOK_URL is missing in GitHub Repository Secrets.")

    # Automatically send email
    try:
        import send_email
        print("\n[Email] Automatically triggering send_email...")
        send_email.main()
    except Exception as em_err:
        print(f"[Email warning] {em_err}")


if __name__ == "__main__":
    main()

"""
divert_stock_check.py — C&S (divert.cssourcing.com) stock-check automation
for Westin Trading.

Semi-automated Python + Playwright script. It drives the C&S sourcing portal
using the user's OWN logged-in session (headed browser, manual login step),
searches each distinct front5 from a research spreadsheet, scrapes the Product
List results, matches them back to the research rows on back5 == CsUPC, and
writes a matches-only copy of the research workbook with a "No Buy" column.

This mirrors the architecture of nettrade_batch_fill.py:
  headed browser → manual login → search loop → scrape → match → write output.

────────────────────────────────────────────────────────────────────────────
CORE PHILOSOPHY: NEVER GUESS.
The divert.cssourcing.com DOM has not been inspected. Every site selector in
the SELECTORS block below is UNSET. The script will NOT run a search job on
guessed selectors:
  1. Run `--inspect` after logging in to dump the real elements on the page.
  2. Fill in the SELECTORS block from what you see.
  3. The pre-flight self-check hard-stops if any required selector is unset or
     does not resolve on the live page.
If the live DOM does not match what is assumed here, STOP and ask the user —
do not invent a plausible selector.
────────────────────────────────────────────────────────────────────────────

Usage:
    # Step 1 — discover the real DOM (login manually when the browser opens):
    python divert_stock_check.py --input RESEARCH.xlsx --inspect

    # Step 2 — validate matching on the first chunk only, then eyeball it:
    python divert_stock_check.py --input RESEARCH.xlsx --stop-after-chunk 1

    # Step 3 — resume and finish all remaining chunks:
    python divert_stock_check.py --input RESEARCH.xlsx

Output:
    <input basename>_STOCKED.xlsx   (matches only, original + "No Buy" column)
    <input basename>_progress.json  (resume state; safe to delete when done)

Dependencies:
    pip install playwright openpyxl
    playwright install chromium
"""

import argparse
import json
import os
import random
import re
import sys
import time
from datetime import datetime

import openpyxl


# ═══════════════════════════════════════════════════════════════════════════
# SELECTORS — UNVERIFIED. Confirm against the live DOM before any real run.
# ═══════════════════════════════════════════════════════════════════════════
# The C&S portal has NOT been inspected. Every value below is None on purpose.
# Run this script with `--inspect` (after logging in manually) to dump the real
# inputs, selects, buttons, and tables on the Product Search page, then fill
# these in with confirmed CSS selectors. verify_selectors() will refuse to run
# a search job until the REQUIRED selectors resolve on the live page.
#
# Use any valid CSS selector, e.g. "#UpcSearch", "input[name='Upc']",
# "table.product-list". Prefer id/name attributes over positional selectors.
SELECTORS = {
    # REQUIRED — the "Upc" search text box on the Product Search page.
    "upc_input": None,
    # REQUIRED — the Search button that submits the query.
    "search_button": None,
    # REQUIRED — the Product List results <table> element.
    "results_table": None,
    # OPTIONAL — the DC dropdown <select>. Leave None to accept the site's
    # default (spec says DC = "All DC"). Set it only if the default is not
    # already "All DC" and you must select it explicitly.
    "dc_dropdown": None,
    # OPTIONAL — a link/button that navigates to Product Search after login.
    # Leave None to instead pause and let the user navigate there manually.
    "product_search_nav": None,
    # OPTIONAL — an element that only exists once logged in. Leave None to use
    # the console "press Enter once logged in" prompt (the safe default, since
    # the login flow is unverified).
    "logged_in_marker": None,
}

# The value to select in the DC dropdown, only used if dc_dropdown is set.
DC_VALUE = "All DC"

LOGIN_URL = "https://divert.cssourcing.com/login"

# Expected Product List column headers (normalized key -> canonical name).
# Header text is matched by normalizing (lowercase, whitespace removed) so
# minor spacing differences don't break the mapping. If the live table's
# headers don't contain the CsUPC and No Buy columns, the scrape hard-stops
# rather than guessing which column is which.
EXPECTED_HEADERS = {
    "dc": "DC",
    "dcname": "DcName",
    "itemcode": "ItemCode",
    "upc": "UPC",
    "csupc": "CsUPC",
    "description": "Description",
    "pk/sz": "Pk/Sz",
    "type": "Type",
    "qcdays": "QC Days",
    "nobuy": "No Buy",
}
REQUIRED_RESULT_COLS = {"CsUPC", "No Buy"}


# ═══════════════════════════════════════════════════════════════════════════
# UPC helpers — the CONFIRMED matching logic. Do not change without asking.
# ═══════════════════════════════════════════════════════════════════════════

def to_digits(val) -> str:
    """Strip everything except digits. Handles Excel floats and None/NaN."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none"):
        return ""
    if "." in s:  # Excel may hand us e.g. '850031180208.0'
        try:
            s = str(int(float(s)))
        except (ValueError, OverflowError):
            pass
    return re.sub(r"\D", "", s)


def pad12(val) -> str:
    """Zero-pad the UPC to a 12-character string (left-padded with '0')."""
    d = to_digits(val)
    if not d:
        return ""
    if len(d) > 12:
        # Not expected for this file shape; caller flags it. Keep last 12.
        return d[-12:]
    return d.zfill(12)


def front5(val) -> str:
    """Characters at index [1:6] of the zero-padded 12-digit UPC.
    Example: 850031180208 -> '50031'."""
    p = pad12(val)
    return p[1:6] if len(p) == 12 else ""


def back5(val) -> str:
    """Last 5 characters of the zero-padded 12-digit UPC."""
    p = pad12(val)
    return p[-5:] if len(p) == 12 else ""


def csupc5(val) -> str:
    """Site CsUPC normalized to a 5-char zero-padded string for comparison.
    CsUPC may lose leading zeros as an Excel/JS number, so we zero-pad."""
    d = to_digits(val)
    if not d:
        return ""
    if len(d) > 5:
        return d[-5:]
    return d.zfill(5)


# ═══════════════════════════════════════════════════════════════════════════
# Research file
# ═══════════════════════════════════════════════════════════════════════════

def is_blank_row(values) -> bool:
    """A fully blank spacer row: every cell is None or empty/whitespace."""
    for v in values:
        if v is not None and str(v).strip() != "":
            return False
    return True


def read_research(path: str):
    """Read the research workbook.

    Returns (headers, rows, upc_idx) where:
      headers  — list of original column header strings (order preserved)
      rows     — list of dicts: {"values": [...], "front5": s, "back5": s}
                 for every non-blank data row (values aligned to headers)
      upc_idx  — index of the UPC column within headers
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    row_iter = ws.iter_rows(values_only=True)
    try:
        header_row = next(row_iter)
    except StopIteration:
        raise SystemExit(f"ERROR: {path} has no rows.")

    headers = [("" if h is None else str(h).strip()) for h in header_row]

    # Locate the UPC column by header name (case-insensitive, trimmed).
    upc_idx = None
    for i, h in enumerate(headers):
        if h.strip().lower() == "upc":
            upc_idx = i
            break
    if upc_idx is None:
        raise SystemExit(
            "ERROR: could not find a 'UPC' column in the research file. "
            f"Headers seen: {headers}"
        )

    rows = []
    blank_count = 0
    for values in row_iter:
        values = list(values)
        # normalize row length to header length
        if len(values) < len(headers):
            values += [None] * (len(headers) - len(values))
        if is_blank_row(values):
            blank_count += 1
            continue
        upc_val = values[upc_idx]
        p12 = pad12(upc_val)
        rows.append({
            "values": values,
            "pad12": p12,
            "front5": p12[1:6] if len(p12) == 12 else "",
            "back5": p12[-5:] if len(p12) == 12 else "",
        })

    print(f"  Read {len(rows)} data rows, skipped {blank_count} blank spacer rows.")
    return headers, rows, upc_idx


def build_front5_index(rows):
    """Map each distinct front5 -> list of research-row indices sharing it."""
    index = {}
    for i, r in enumerate(rows):
        f5 = r["front5"]
        if not f5:
            continue
        index.setdefault(f5, []).append(i)
    return index


def chunk_list(items, n):
    """Split a list into n roughly-equal contiguous chunks."""
    items = list(items)
    if n <= 0:
        return [items]
    k, m = divmod(len(items), n)
    chunks, start = [], 0
    for i in range(n):
        size = k + (1 if i < m else 0)
        chunks.append(items[start:start + size])
        start += size
    return [c for c in chunks if c]  # drop empties when items < n


# ═══════════════════════════════════════════════════════════════════════════
# Progress (resume) file
# ═══════════════════════════════════════════════════════════════════════════

def load_progress(path, expected_chunks, input_path):
    """Load progress JSON if valid for this input, else start fresh."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r") as f:
            prog = json.load(f)
    except (json.JSONDecodeError, OSError):
        print(f"  Progress file {path} unreadable — starting fresh.")
        return None
    if prog.get("input_file") != os.path.abspath(input_path):
        print("  Progress file is for a different input — starting fresh.")
        return None
    if prog.get("chunks") != expected_chunks:
        print("  Front5 set changed since last run — starting fresh.")
        return None
    done = len(prog.get("completed_chunks", []))
    print(f"  Resuming: {done}/{len(expected_chunks)} chunks already completed.")
    return prog


def save_progress(path, prog):
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(prog, f, indent=2)
    os.replace(tmp, path)  # atomic-ish: never leaves a half-written file


# ═══════════════════════════════════════════════════════════════════════════
# Browser / scraping
# ═══════════════════════════════════════════════════════════════════════════

def wait_for_login(page):
    """Block until the user has logged in manually."""
    marker = SELECTORS.get("logged_in_marker")
    if marker:
        print(f"  Waiting for logged-in marker: {marker}")
        page.wait_for_selector(marker, timeout=0)  # no timeout; user-paced
        print("  Detected logged-in marker.")
    else:
        input(
            "\n>>> Log in to divert.cssourcing.com in the opened browser window, "
            "then press Enter here to continue... "
        )


def goto_product_search(page):
    """Navigate to Product Search, either by a confirmed nav selector or a
    manual pause (never guess the nav link)."""
    nav = SELECTORS.get("product_search_nav")
    if nav:
        page.click(nav)
        page.wait_for_load_state("networkidle")
    else:
        input(
            "\n>>> Navigate to the Product Search page in the browser, "
            "then press Enter here to continue... "
        )


def verify_selectors(page):
    """Pre-flight self-check. HARD STOP if any REQUIRED selector is unset or
    does not resolve on the current page. This is the gate that prevents a
    full run on guessed selectors."""
    required = ["upc_input", "search_button", "results_table"]
    problems = []
    for key in required:
        sel = SELECTORS.get(key)
        if not sel:
            problems.append(f"  - '{key}' is UNSET in the SELECTORS block.")
            continue
        try:
            el = page.query_selector(sel)
        except Exception as e:  # invalid selector syntax, etc.
            problems.append(f"  - '{key}' selector '{sel}' errored: {e}")
            continue
        if el is None:
            problems.append(
                f"  - '{key}' selector '{sel}' did not match any element "
                "on the current page."
            )

    # Optional selectors: if set, they must resolve too (a set-but-broken
    # selector is a silent bug we won't tolerate).
    for key in ("dc_dropdown", "product_search_nav", "logged_in_marker"):
        sel = SELECTORS.get(key)
        if sel:
            try:
                if page.query_selector(sel) is None:
                    problems.append(
                        f"  - optional '{key}' selector '{sel}' is set but "
                        "matched no element on this page."
                    )
            except Exception as e:
                problems.append(f"  - optional '{key}' selector '{sel}' errored: {e}")

    if problems:
        print("\n" + "=" * 70)
        print("SELECTOR SELF-CHECK FAILED — refusing to run on guessed selectors.")
        print("=" * 70)
        print("\n".join(problems))
        print(
            "\nFix: run with --inspect (after logging in and reaching the "
            "Product Search page with visible results) to dump the real DOM, "
            "fill in the SELECTORS block, and try again.\n"
            "NEVER guess a selector — confirm it against the live page first."
        )
        raise SystemExit(2)

    print("  Selector self-check passed.")


def inspect_page(page):
    """Dump the interactive elements and tables on the current page so the
    real selectors can be identified. Run after login + reaching Product
    Search. Read-only; changes nothing."""
    print("\n" + "=" * 70)
    print(f"INSPECT — current URL: {page.url}")
    print("=" * 70)

    def attrs(el):
        return el.evaluate(
            "e => ({tag:e.tagName, id:e.id, name:e.getAttribute('name'), "
            "type:e.getAttribute('type'), placeholder:e.getAttribute('placeholder'), "
            "cls:e.getAttribute('class'), text:(e.innerText||'').trim().slice(0,40)})"
        )

    print("\n--- <input> elements ---")
    for el in page.query_selector_all("input"):
        print("  ", attrs(el))

    print("\n--- <select> elements ---")
    for el in page.query_selector_all("select"):
        a = attrs(el)
        opts = [o.inner_text().strip() for o in el.query_selector_all("option")]
        print("  ", a, "options:", opts[:12])

    print("\n--- <button> / [type=submit] / <a> (clickables) ---")
    for el in page.query_selector_all("button, input[type=submit], a"):
        a = attrs(el)
        if a.get("text") or a.get("id") or a.get("name"):
            print("  ", a)

    print("\n--- <table> elements (with first-row header cells) ---")
    for idx, el in enumerate(page.query_selector_all("table")):
        a = el.evaluate("e => ({id:e.id, cls:e.getAttribute('class')})")
        first = el.query_selector("tr")
        heads = []
        if first:
            heads = [c.inner_text().strip()
                     for c in first.query_selector_all("th, td")]
        print(f"  table[{idx}]", a, "headers:", heads)

    print("\n" + "=" * 70)
    print("Copy the confirmed selectors into the SELECTORS block, then re-run "
          "without --inspect.")
    print("=" * 70)


def scrape_results(page):
    """Scrape the Product List table on the current page.

    Returns a list of row dicts keyed by canonical header names. Returns []
    when the site shows zero results (skip silently, per spec).

    Hard-stops if the table exists but its headers don't contain the required
    columns — we never guess which column is CsUPC / No Buy.
    """
    table = page.query_selector(SELECTORS["results_table"])
    if table is None:
        return []  # zero results / no table rendered — skip silently

    trs = table.query_selector_all("tr")
    if not trs:
        return []

    header_cells = trs[0].query_selector_all("th, td")
    raw_headers = [c.inner_text().strip() for c in header_cells]

    # Map column position -> canonical name via normalized header text.
    col_map = {}  # position index -> canonical name
    for pos, h in enumerate(raw_headers):
        norm = re.sub(r"\s+", "", h.strip().lower())
        if norm in EXPECTED_HEADERS:
            col_map[pos] = EXPECTED_HEADERS[norm]

    present = set(col_map.values())
    missing = REQUIRED_RESULT_COLS - present
    if missing:
        raise SystemExit(
            "ERROR: Product List headers are missing required column(s) "
            f"{sorted(missing)}. Headers seen: {raw_headers}. "
            "Confirm the correct results table selector with --inspect — "
            "do not guess the column mapping."
        )

    rows = []
    for tr in trs[1:]:
        cells = tr.query_selector_all("th, td")
        if not cells:
            continue
        rec = {}
        for pos, cell in enumerate(cells):
            if pos in col_map:
                rec[col_map[pos]] = cell.inner_text().strip()
        if rec:
            rows.append(rec)
    return rows


def run_searches(page, remaining_chunks, chunk_offset, prog, progress_path,
                 stop_after):
    """Run the search loop over the remaining chunks, saving progress after
    each chunk. Mutates prog in place."""
    total_chunks = len(prog["chunks"])
    processed_this_run = 0

    for local_i, chunk in enumerate(remaining_chunks):
        chunk_idx = chunk_offset + local_i
        print(f"\n--- Chunk {chunk_idx + 1}/{total_chunks} "
              f"({len(chunk)} front5 values) ---")

        for f5 in chunk:
            if f5 in prog["searched_front5"]:
                continue  # already done (e.g. mid-chunk crash last time)

            # a. enter front5, leave DC = All DC (default), click Search
            page.fill(SELECTORS["upc_input"], f5)
            if SELECTORS.get("dc_dropdown"):
                try:
                    page.select_option(SELECTORS["dc_dropdown"], label=DC_VALUE)
                except Exception:
                    pass  # default is already All DC; not fatal
            page.click(SELECTORS["search_button"])
            # Let results render. networkidle is best-effort; fall back to a
            # short settle if the site keeps a long-poll open.
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                page.wait_for_timeout(1500)

            # b. scrape the Product List table
            results = scrape_results(page)
            prog["results"][f5] = results
            prog["searched_front5"].append(f5)
            if results:
                print(f"    {f5}: {len(results)} result rows")
            # c. zero rows -> skip silently (no log line)

            # 3. randomized 1-3s courtesy delay between searches
            time.sleep(random.uniform(1.0, 3.0))

        # 4. save progress after EACH chunk
        if chunk_idx not in prog["completed_chunks"]:
            prog["completed_chunks"].append(chunk_idx)
        save_progress(progress_path, prog)
        print(f"  Chunk {chunk_idx + 1} complete — progress saved.")

        processed_this_run += 1
        if stop_after and processed_this_run >= stop_after:
            print(f"\nStopping after {stop_after} chunk(s) this run "
                  "(--stop-after-chunk). Re-run to resume.")
            return False  # not fully finished
    return True  # all chunks done


# ═══════════════════════════════════════════════════════════════════════════
# Matching + output
# ═══════════════════════════════════════════════════════════════════════════

def match_and_report(rows, front5_index, prog, verbose_sample=8):
    """Match scraped results back to research rows.

    For each scraped row: front5 F and CsUPC -> csupc5 C5. A research row
    sharing F whose back5 == C5 is a HIT. Dedup across DCs: one output row per
    matched research row; "No Buy" = distinct values seen across all matching
    scraped rows, comma-joined.

    Returns dict: research_row_index -> "No Buy" string.
    """
    matched = {}      # row_idx -> set of No Buy values
    samples = []      # for validation printout

    for f5, results in prog["results"].items():
        candidates = front5_index.get(f5, [])
        if not candidates or not results:
            continue
        # back5 -> [research row indices] for this front5
        by_back5 = {}
        for ridx in candidates:
            by_back5.setdefault(rows[ridx]["back5"], []).append(ridx)

        for res in results:
            c5 = csupc5(res.get("CsUPC", ""))
            if not c5:
                continue
            hit_rows = by_back5.get(c5)
            if not hit_rows:
                continue
            no_buy = (res.get("No Buy") or "").strip()
            for ridx in hit_rows:
                matched.setdefault(ridx, set())
                if no_buy != "":
                    matched[ridx].add(no_buy)
                if len(samples) < verbose_sample:
                    samples.append((f5, rows[ridx], res, c5))

    # Print a validation sample so the user can eyeball the derivation before
    # trusting a full run (front5/back5 vs site CsUPC).
    if samples:
        print("\n--- Match validation sample (verify before trusting a full run) ---")
        print("  front5 | pad12 -> back5 | site CsUPC (raw -> pad5) | No Buy")
        for f5, rrow, res, c5 in samples:
            print(f"    {f5} | pad12={rrow['pad12']} back5={rrow['back5']} | "
                  f"CsUPC='{res.get('CsUPC','')}' -> {c5} | "
                  f"NoBuy='{res.get('No Buy','')}'")
        print("  (If CsUPC raw already shows leading zeros, no padding was needed; "
              "if it lost them, the zero-pad above recovered the match.)")

    # Collapse No Buy sets to comma-joined strings (distinct, sorted).
    out = {}
    for ridx, vals in matched.items():
        if vals:
            out[ridx] = ", ".join(sorted(vals))
        else:
            out[ridx] = ""  # matched but No Buy blank across all DCs
    return out


def write_output(headers, rows, no_buy_map, out_path):
    """Write a matches-only copy: original columns + a 'No Buy' column."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stocked"

    ws.append(list(headers) + ["No Buy"])

    n = 0
    # Preserve original row order among matched rows.
    for ridx in sorted(no_buy_map.keys()):
        ws.append(list(rows[ridx]["values"]) + [no_buy_map[ridx]])
        n += 1

    wb.save(out_path)
    return n


# ═══════════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════════

def default_out_path(input_path):
    base, ext = os.path.splitext(input_path)
    return f"{base}_STOCKED{ext or '.xlsx'}"


def default_progress_path(input_path):
    base, _ = os.path.splitext(input_path)
    return f"{base}_progress.json"


def main():
    ap = argparse.ArgumentParser(
        description="C&S (divert.cssourcing.com) stock-check automation.")
    ap.add_argument("--input", required=True, help="Research .xlsx file.")
    ap.add_argument("--out", default=None,
                    help="Output workbook (default: <input>_STOCKED.xlsx).")
    ap.add_argument("--progress", default=None,
                    help="Progress JSON path (default: <input>_progress.json).")
    ap.add_argument("--chunks", type=int, default=5,
                    help="Number of chunks to split front5 values into (default 5).")
    ap.add_argument("--stop-after-chunk", type=int, default=0,
                    help="Stop after N chunks this run (0 = run to completion). "
                         "Use 1 to validate the first chunk, then re-run to resume.")
    ap.add_argument("--inspect", action="store_true",
                    help="Log in, open Product Search, dump the DOM, and exit. "
                         "Use this to discover the real selectors.")
    ap.add_argument("--headless", action="store_true",
                    help="Run headless (NOT recommended — login is manual).")
    ap.add_argument("--rebuild-output", action="store_true",
                    help="Skip searching; just rebuild the output from an "
                         "existing complete progress file.")
    args = ap.parse_args()

    out_path = args.out or default_out_path(args.input)
    progress_path = args.progress or default_progress_path(args.input)

    print("=" * 70)
    print("C&S DIVERT STOCK CHECK")
    print("=" * 70)

    # ---- Read research file, build front5 chunks ----
    print("\nReading research file...")
    headers, rows, _ = read_research(args.input)
    front5_index = build_front5_index(rows)
    unique_front5 = sorted(front5_index.keys())
    print(f"  {len(unique_front5)} distinct front5 values across "
          f"{len(rows)} research rows.")

    chunks = chunk_list(unique_front5, args.chunks)
    print(f"  Split into {len(chunks)} chunks.")

    # ---- Rebuild-output-only path (no browser) ----
    if args.rebuild_output:
        prog = load_progress(progress_path, chunks, args.input)
        if not prog:
            raise SystemExit("ERROR: no valid progress file to rebuild from.")
        no_buy_map = match_and_report(rows, front5_index, prog)
        n = write_output(headers, rows, no_buy_map, out_path)
        print(f"\nWrote {n} matched rows to {out_path}")
        return

    # ---- Load / init progress ----
    prog = load_progress(progress_path, chunks, args.input)
    if prog is None:
        prog = {
            "input_file": os.path.abspath(args.input),
            "created": datetime.now().isoformat(timespec="seconds"),
            "chunks": chunks,
            "completed_chunks": [],
            "searched_front5": [],
            "results": {},
        }

    # ---- Browser session ----
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        page = browser.new_page()
        print(f"\nOpening {LOGIN_URL} ...")
        page.goto(LOGIN_URL)

        wait_for_login(page)
        goto_product_search(page)

        if args.inspect:
            inspect_page(page)
            browser.close()
            return

        # Pre-flight gate: never run on guessed/unresolved selectors.
        verify_selectors(page)

        # Determine which chunks remain.
        completed = set(prog["completed_chunks"])
        chunk_offset = 0
        while chunk_offset in completed:
            chunk_offset += 1
        remaining = chunks[chunk_offset:]

        if not remaining:
            print("\nAll chunks already completed — building output.")
            finished = True
        else:
            finished = run_searches(
                page, remaining, chunk_offset, prog, progress_path,
                args.stop_after_chunk)

        browser.close()

    # ---- Match + write output (always, so partial progress is usable) ----
    no_buy_map = match_and_report(rows, front5_index, prog)
    n = write_output(headers, rows, no_buy_map, out_path)
    print(f"\nWrote {n} matched rows to {out_path}")
    if not finished:
        print("Run was stopped early / partial — re-run the same command to resume.")
    else:
        print("Done. All chunks processed.")


if __name__ == "__main__":
    main()

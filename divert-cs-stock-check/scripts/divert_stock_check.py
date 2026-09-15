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
    <input basename>_STOCKED.xlsx      (internal: Stocked/Review Queue/Stocked
                                        Vendor Lines/Lookfor tabs)
    <input basename>_CS_STOCK_SUMMARY.xlsx  (account-manager summary, On
                                        Offer + Ask Source)
    <input basename>_FINAL_REVIEW.xlsx (account-manager-ready: source columns,
                                        No Buy == "No" only, bold = case-code
                                        match, plain = item-code match)
    <input basename>_progress.json     (resume state; safe to delete when done)

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
from openpyxl.styles import Font


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
    # REQUIRED — confirmed 2026-08-24 via --inspect on the real login/search
    # page (this ASP.NET site swaps search content into the same /login URL
    # via postback rather than navigating to a new URL).
    "upc_input": "#MainContent_upc",
    # REQUIRED — confirmed 2026-08-24 via --inspect.
    "search_button": "#MainContent_search_btn",
    # REQUIRED — confirmed 2026-08-24 via --inspect with a real search
    # ("10095") executed first. table[60] in the dump: id='MainContent_product_item',
    # headers exactly ['DC', 'DcName', 'ItemCode', 'UPC', 'CsUPC', 'Description',
    # 'Pk/Sz', 'Type', 'QC Days', 'No Buy'].
    "results_table": "#MainContent_product_item",
    # OPTIONAL — CONFIRMED 2026-08-24. The DC dropdown is NOT a plain <select>
    # on this site (it's a custom Infragistics widget — no <select> elements
    # exist at all). A real search for "10095" returned results from 20+
    # distinct DCs (01024, 01A01, 06066, 07067, 15602, 15603, 15606, 15607,
    # 15610, 15614, 15630, ...) without touching this control, proving "All DC"
    # is already in effect by default. Leave unset.
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

# Confirmed 2026-08-24 via manual --inspect: searching "10095" reliably
# returns many Product List rows on the live site. Used only as a probe
# search during the pre-flight self-check, since the results table does not
# exist in the DOM until at least one search has been run — not used for
# any real front5 lookup.
PROBE_FRONT5 = "10095"

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
    """Characters at index [6:11] of the zero-padded 12-digit UPC — the item
    code that sits between the manufacturer code (front5) and the trailing
    GS1 check digit at index 11.

    CORRECTED 2026-08-24: originally taken as the literal last 5 characters
    (index [7:12]), which silently swaps in the check digit for the item
    code's true last digit. Caught only because real confirmed-buy UPCs gave
    ground truth to check against: e.g. 0-72310-00041-4 decodes to
    system=0, mfr(front5)=72310, item=00041, check=4 — the live site's CsUPC
    for that exact product is '00041' (index [6:11] of the padded UPC), not
    '00414' (the old index [7:12] result). Verified against three independent
    confirmed BIGELOW matches, all exact under this formula, all wrong under
    the old one. front5 is unaffected — this only shifts the back5 window.
    """
    p = pad12(val)
    return p[6:11] if len(p) == 12 else ""


def gtin14_front5(val) -> str:
    """Front5 from a 14-digit case GTIN: characters [3:8].

    Confirmed 2026-09-15 by the user as the authoritative rule for reading
    a GTIN-14: "drop first three numbers and last number", leaving
    front5 + case code. e.g. '10070277000055' -> front5 '70277',
    case code '00005'. Verified against 90 rows of the Emmi-Roth file that
    carry both a GTIN-12 and a GTIN-14; the 10 that differ do so because
    the case code and item code are genuinely different numbers for the
    same product, not because the rule is wrong."""
    d = to_digits(val)
    return d[3:8] if len(d) == 14 else ""


def gtin14_case_back5(val) -> str:
    """Case code from a 14-digit case GTIN: characters [8:13]. This is the
    value C&S shows in its CsUPC column — the case-level identifier — as
    distinct from the item code carried by a GTIN-12."""
    d = to_digits(val)
    return d[8:13] if len(d) == 14 else ""


def format_upc_display(pad12_val: str, raw_val) -> str:
    """Dashed UPC-A for readability: system-mfr5-item5-check, e.g.
    '072310000414' -> '0-72310-00041-4'. Falls back to the original raw
    research value (unmodified) if pad12 isn't a clean 12-digit string —
    display-only, never used for matching."""
    if pad12_val and len(pad12_val) == 12 and pad12_val.isdigit():
        return f"{pad12_val[0]}-{pad12_val[1:6]}-{pad12_val[6:11]}-{pad12_val[11]}"
    return raw_val


def csupc5(val) -> str:
    """Site CsUPC normalized to a 5-char zero-padded string for comparison.
    CsUPC may lose leading zeros as an Excel/JS number, so we zero-pad."""
    d = to_digits(val)
    if not d:
        return ""
    if len(d) > 5:
        return d[-5:]
    return d.zfill(5)


def site_upc_back5(val) -> str:
    """Back5 portion of the site's own scraped 'UPC' column, e.g.
    '50003-79769' -> '79769'. This is a DIFFERENT real field than CsUPC —
    confirmed 2026-08-24 that the two can genuinely disagree on the same
    scraped row (e.g. UPC '50003-79769' with CsUPC '79774'). Used only for
    the Review Queue (a second, still-exact match against a different real
    field), never for the primary CsUPC-based match — do not conflate them.
    """
    s = str(val or "").strip()
    if "-" in s:
        tail = s.split("-")[-1]
        if tail.isdigit() and len(tail) == 5:
            return tail
    d = to_digits(val)
    # Undashed fallbacks. HARDENED 2026-09-14: this used to return d[-5:]
    # unconditionally, which on a full 12-digit UPC is the disproved last-5
    # [7:12] window (check digit in place of the item code's last digit) —
    # the very bug corrected in back5(). It mattered little while this field
    # only fed the Review Queue; now that a site-UPC match confirms an item
    # as stocked, it has to use the same window as back5().
    if len(d) in (11, 12):
        return pad12(d)[6:11]
    if len(d) == 10:          # front5 + item code, dash simply absent
        return d[5:10]
    return d[-5:] if len(d) >= 5 else ""


# Generic filler words common across grocery descriptions regardless of
# vendor — excluded so they can't create a false relevance match between
# two genuinely unrelated product lines that happen to share a front5.
_DESCRIPTION_STOPWORDS = {
    "AND", "THE", "FOR", "WITH", "NEW", "PER", "ALL", "ORIG", "ORIGINAL",
    "FRESH", "PACK", "SIZE", "CASE", "EACH",
}


def significant_tokens(text) -> set:
    """Meaningful uppercase word tokens from a description, for the Lookfor
    relevance filter — NOT used for identity matching (that stays strictly
    exact-digit based). Drops punctuation, pure numbers, short tokens
    (<3 chars), and common cross-vendor filler words."""
    s = str(text or "").upper()
    words = re.findall(r"[A-Z]+", s)
    return {w for w in words if len(w) >= 3 and w not in _DESCRIPTION_STOPWORDS}


# Corporate-suffix noise in the research file's BRAND column — these carry
# no vendor identity, so they must never anchor a brand match.
_BRAND_STOPWORDS = {"LLC", "INC", "CORP", "CO", "LTD", "COMPANY", "THE",
                    "AND", "BRANDS"}


def consonant_skeleton(word) -> str:
    """First letter, then all consonants — e.g. BIGELOW -> 'BGLW'.

    C&S abbreviates descriptions by dropping vowels, at varying depth for
    the same vendor (BIGELOW -> BIGLOW -> BGLW; BOTTICELLI -> BTLLI;
    CLEANSE -> CLNCSE). Reducing both sides to a consonant skeleton
    normalizes across that, so one brand's many spellings collapse together
    while genuinely different brands stay apart.
    """
    w = re.sub(r"[^A-Z]", "", str(word or "").upper())
    if not w:
        return ""
    return w[0] + "".join(c for c in w[1:] if c not in "AEIOU")


def _is_subsequence(needle: str, haystack: str) -> bool:
    """True if every char of needle appears in haystack in order."""
    it = iter(haystack)
    return all(c in it for c in needle)


def brand_token_matches(brand_word, token) -> bool:
    """True if a site-description token plausibly IS this brand word, via
    consonant-skeleton subsequence. Requires the same first letter, so
    'DM'/'DELMONTE' can never match 'JOYBA'. Two-letter skeletons use a
    stricter prefix test, since they're too short to be safe as a
    subsequence."""
    sa = consonant_skeleton(brand_word)
    sb = consonant_skeleton(token)
    if len(sa) < 2 or len(sb) < 2 or sa[0] != sb[0]:
        return False
    short, long_ = (sa, sb) if len(sa) <= len(sb) else (sb, sa)
    if len(short) == 2:
        return long_.startswith(short)
    return _is_subsequence(short, long_)


def matched_brand_word(description, brands):
    """Return the brand word a description's leading tokens identify it as,
    or None. Confirmed 2026-08-24 against every real scraped description
    collected: C&S descriptions lead with a brand abbreviation
    ('BIGLOW RED RSPBRRY', 'BTLLI EXTRA VIRGIN OLIVE OIL',
    'CLNCSE ORG MNT HNY YRB MT T', '*JOYBA BBL RASP...',
    'DM CUT GREEN BEANS'). That token is the one field that separates two
    unrelated vendors sharing a front5 — JOYBA is a Del Monte brand, so
    their manufacturer prefix and much of their fruit/beverage vocabulary
    genuinely overlap, and only the brand token tells them apart."""
    tokens = re.findall(r"[A-Z]+", str(description or "").upper())
    for brand in brands:
        for bw in re.findall(r"[A-Z]+", str(brand or "").upper()):
            if len(bw) < 2 or bw in _BRAND_STOPWORDS:
                continue
            for t in tokens:
                if brand_token_matches(bw, t):
                    return bw
    return None


# ═══════════════════════════════════════════════════════════════════════════
# Research file
# ═══════════════════════════════════════════════════════════════════════════

def is_blank_row(values) -> bool:
    """A fully blank spacer row: every cell is None or empty/whitespace."""
    for v in values:
        if v is not None and str(v).strip() != "":
            return False
    return True


def find_header_index(headers, name):
    """Locate a column by header name (case-insensitive, trimmed). Returns
    None if not found — callers decide whether that's fatal."""
    target = name.strip().lower()
    for i, h in enumerate(headers):
        if h.strip().lower() == target:
            return i
    return None


def read_research(path: str):
    """Read the research workbook.

    Returns (headers, rows, upc_idx) where:
      headers  — list of original column header strings (order preserved)
      rows     — list of dicts: {"values": [...], "front5": s, "back5": s}
                 for every non-blank data row (values aligned to headers)
      upc_idx  — index of the UPC column within headers
      prep_audit — dict describing any pre-computed UPC12/FRONT5/BACK5
                 helper columns found in the file and whether they agree
                 with this script's own confirmed derivation

    A "SEARCH PREP" file may ship with UPC12/FRONT5/BACK5 already computed
    upstream. Those columns are NEVER used for matching — front5/back5 are
    always recomputed here from the raw UPC via the confirmed rule — but
    they ARE audited, because a helper column computed with the old,
    pre-2026-08-24 last-5 formula is a silent trap for anyone reading the
    sheet by eye.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    row_iter = ws.iter_rows(values_only=True)
    try:
        header_row = next(row_iter)
    except StopIteration:
        raise SystemExit(f"ERROR: {path} has no rows.")

    headers = [("" if h is None else str(h).strip()) for h in header_row]

    # OPTIONAL case-level GTIN-14 column. When present, each row carries
    # BOTH identifiers: the item code (from the GTIN-12 "UPC" column, which
    # matches the site's UPC column) and the case code (from the GTIN-14,
    # which matches the site's CsUPC column). Added 2026-09-15 for vendor
    # files that publish both, e.g. the Emmi-Roth price list. Without this
    # column the script behaves exactly as before: one back5 per row,
    # compared against both site fields.
    case_idx = None
    for cand in ("CASE UPC", "GTIN-14", "GTIN-14 UPC CODE", "CASE GTIN"):
        case_idx = find_header_index(headers, cand)
        if case_idx is not None:
            break

    upc_idx = find_header_index(headers, "UPC")
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
        item_front5 = p12[1:6] if len(p12) == 12 else ""
        item_back5 = p12[6:11] if len(p12) == 12 else ""

        case_front5 = case_back5 = ""
        if case_idx is not None:
            case_val = values[case_idx]
            case_front5 = gtin14_front5(case_val)
            case_back5 = gtin14_case_back5(case_val)

        # A row may legitimately carry only one of the two codes (a random-
        # weight item often has no consumer GTIN-12 at all). Fall back so the
        # row is still searchable and matchable on whichever code it has.
        rows.append({
            "values": values,
            "pad12": p12,
            "front5": item_front5 or case_front5,
            "back5": item_back5 or case_back5,
            "item_back5": item_back5,
            "case_back5": case_back5,
        })

    print(f"  Read {len(rows)} data rows, skipped {blank_count} blank spacer rows.")

    # ---- Audit pre-computed helper columns (never used for matching) ----
    prep_audit = {}
    for col, key in (("UPC12", "pad12"), ("FRONT5", "front5"), ("BACK5", "back5")):
        idx = find_header_index(headers, col)
        if idx is None:
            continue
        width = 12 if key == "pad12" else 5
        disagree = 0
        example = None
        for r in rows:
            got = str(r["values"][idx]).strip() if r["values"][idx] is not None else ""
            got = got.split(".")[0].zfill(width) if got else ""
            if got != r[key]:
                disagree += 1
                if example is None:
                    example = (r["pad12"], got, r[key])
        prep_audit[col] = {"rows": len(rows), "disagree": disagree, "example": example}
        if disagree:
            legacy = sum(1 for r in rows
                         if (str(r["values"][idx]).strip().split(".")[0].zfill(width)
                             if r["values"][idx] is not None else "") == r["pad12"][7:12])
            print(f"  !! WARNING: the file's '{col}' column disagrees with the "
                  f"confirmed derivation on {disagree}/{len(rows)} rows.")
            if example:
                print(f"     e.g. UPC {example[0]}: file={example[1]!r} "
                      f"confirmed={example[2]!r}")
            if col == "BACK5" and legacy == len(rows):
                print("     Every row matches the OLD last-5 [7:12] formula, which "
                      "swaps the GS1 check digit in for the item code's last digit.")
                print("     That formula was disproved against real confirmed-buy "
                      "ground truth on 2026-08-24. This script IGNORES the column "
                      "and recomputes back5 as [6:11]; matching is unaffected.")
            prep_audit[col]["all_legacy_last5"] = (legacy == len(rows))

    return headers, rows, upc_idx, prep_audit


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

def load_progress(path, expected_chunks, input_path, require_chunk_layout=True):
    """Load progress JSON if valid for this input, else start fresh.

    require_chunk_layout=True (the searching path) also requires the exact
    chunk PARTITION to match --chunks from the prior run, since resuming
    relies on completed_chunks indexing into that same partition.

    require_chunk_layout=False (the --rebuild-output / --diagnose paths,
    which never search) only needs the underlying front5 SET to match —
    chunk count is irrelevant when nothing is being resumed. FIXED
    2026-09-14: previously used the same strict chunks == check for both,
    so a fully-searched progress file from a `--full-speed` (--chunks 1)
    run was discarded by `--rebuild-output`'s default --chunks 5, with the
    misleading message "Front5 set changed" even though it hadn't."""
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
    expected_set = sorted(f5 for chunk in expected_chunks for f5 in chunk)
    saved_chunks = prog.get("chunks") or []
    saved_set = sorted(f5 for chunk in saved_chunks for f5 in chunk)
    if saved_set != expected_set:
        print("  Front5 set changed since last run (research file content "
              "differs) — starting fresh.")
        return None
    if require_chunk_layout and saved_chunks != expected_chunks:
        print("  Progress file uses a different --chunks layout than this "
              "run — starting fresh. Pass the same --chunks value used "
              "originally to resume, or use --rebuild-output (chunk layout "
              "doesn't matter there).")
        return None
    if require_chunk_layout:
        done = len(prog.get("completed_chunks", []))
        print(f"  Resuming: {done}/{len(expected_chunks)} chunks already "
              "completed.")
    else:
        n_searched = len(prog.get("searched_front5", []))
        print(f"  Loaded progress: {n_searched}/{len(expected_set)} front5 "
              "values searched (chunk layout ignored for this read-only "
              "path).")
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


def _check_selector(page, key, problems):
    """Check one selector; append a problem message and return False if it
    doesn't resolve. Returns True if it resolves cleanly."""
    sel = SELECTORS.get(key)
    if not sel:
        problems.append(f"  - '{key}' is UNSET in the SELECTORS block.")
        return False
    try:
        el = page.query_selector(sel)
    except Exception as e:  # invalid selector syntax, etc.
        problems.append(f"  - '{key}' selector '{sel}' errored: {e}")
        return False
    if el is None:
        problems.append(
            f"  - '{key}' selector '{sel}' did not match any element "
            "on the current page."
        )
        return False
    return True


def verify_selectors(page):
    """Pre-flight self-check. HARD STOP if any REQUIRED selector is unset or
    does not resolve on the current page. This is the gate that prevents a
    full run on guessed selectors.

    results_table only exists in the DOM after a search has run (confirmed
    via --inspect: the bare Product Search page has no such table). So this
    runs one real probe search — using the already-verified upc_input and
    search_button — before checking for results_table, rather than checking
    for it on a page where it could never legitimately be present yet.
    """
    problems = []

    upc_ok = _check_selector(page, "upc_input", problems)
    search_ok = _check_selector(page, "search_button", problems)

    if upc_ok and search_ok:
        page.fill(SELECTORS["upc_input"], PROBE_FRONT5)
        page.click(SELECTORS["search_button"])
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            page.wait_for_timeout(2000)
        _check_selector(page, "results_table", problems)
    else:
        # Can't probe without a working search box/button — report
        # results_table as unverifiable rather than silently skipping it.
        problems.append(
            "  - 'results_table' could not be checked: upc_input/"
            "search_button must resolve first to run the probe search."
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
    # Let any in-flight navigation/render finish before querying the DOM.
    # If the user just clicked Search right before pressing Enter, the page
    # may still be mid-navigation — querying too early destroys the execution
    # context and crashes Playwright. Settle first, and retry once on that
    # specific error rather than guessing anything about the page.
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        page.wait_for_timeout(1500)

    print("\n" + "=" * 70)
    print(f"INSPECT — current URL: {page.url}")
    print("=" * 70)

    def attrs(el):
        return el.evaluate(
            "e => ({tag:e.tagName, id:e.id, name:e.getAttribute('name'), "
            "type:e.getAttribute('type'), placeholder:e.getAttribute('placeholder'), "
            "cls:e.getAttribute('class'), text:(e.innerText||'').trim().slice(0,40)})"
        )

    def query_all(selector):
        """query_selector_all with one retry if navigation raced us."""
        try:
            return page.query_selector_all(selector)
        except Exception as e:
            if "Execution context was destroyed" in str(e) or "navigat" in str(e).lower():
                page.wait_for_timeout(2000)
                return page.query_selector_all(selector)
            raise

    print("\n--- <input> elements ---")
    for el in query_all("input"):
        print("  ", attrs(el))

    print("\n--- <select> elements ---")
    for el in query_all("select"):
        a = attrs(el)
        opts = [o.inner_text().strip() for o in el.query_selector_all("option")]
        print("  ", a, "options:", opts[:12])

    print("\n--- <button> / [type=submit] / <a> (clickables) ---")
    for el in query_all("button, input[type=submit], a"):
        a = attrs(el)
        if a.get("text") or a.get("id") or a.get("name"):
            print("  ", a)

    print("\n--- <table> elements (with first-row header cells) ---")
    found_dc_upc_table = False
    for idx, el in enumerate(query_all("table")):
        a = el.evaluate("e => ({id:e.id, cls:e.getAttribute('class')})")
        first = el.query_selector("tr")
        heads = []
        if first:
            heads = [c.inner_text().strip()
                     for c in first.query_selector_all("th, td")]
        if any("csupc" in h.lower().replace(" ", "") for h in heads):
            found_dc_upc_table = True
        print(f"  table[{idx}]", a, "headers:", heads)

    # Fallback: if no <table> contained the Product List headers, the grid
    # may be rendered as non-semantic <div>-based markup instead (some ASP.NET
    # grid widgets do this). Broaden the search rather than assume defeat.
    if not found_dc_upc_table:
        print("\n--- No <table> matched Product List headers — scanning for "
              "text 'CsUPC' anywhere on the page (fallback for div-based "
              "grids) ---")
        try:
            matches = page.evaluate(
                """() => {
                    const out = [];
                    const walker = document.createTreeWalker(
                        document.body, NodeFilter.SHOW_ELEMENT);
                    let node;
                    while ((node = walker.nextNode())) {
                        const txt = (node.textContent || '');
                        if (txt.includes('CsUPC') &&
                            node.children.length > 0 &&
                            node.children.length < 30) {
                            out.push({tag: node.tagName, id: node.id,
                                      cls: node.getAttribute('class'),
                                      childCount: node.children.length});
                        }
                    }
                    return out.slice(0, 15);
                }"""
            )
            for m in matches:
                print("  ", m)
            if not matches:
                print("  No element containing 'CsUPC' text found at all — "
                      "results likely did not load before this dump. Re-run "
                      "and confirm rows are visible on screen before "
                      "pressing Enter.")
        except Exception as e:
            print(f"  Fallback scan failed: {e}")

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
                 stop_after, delay_min=1.0, delay_max=3.0, save_every=25,
                 max_retries=3):
    """Run the search loop over the remaining chunks, saving progress after
    each chunk and every `save_every` searches within a chunk.

    Every front5 attempted is accounted for in prog["run_log"] with an
    explicit status — "hit", "zero", or "error". A silent zero-result and a
    failed request must never look the same, so a search that raises is
    retried with exponential backoff and, if it still fails, recorded as an
    error rather than being allowed to masquerade as a zero-result.

    delay_min/delay_max of 0 run at full speed with no courtesy delay; the
    backoff on a failed request still applies, so a real rate limit is
    absorbed and logged rather than hammered.
    """
    total_chunks = len(prog["chunks"])
    processed_this_run = 0
    since_save = 0
    run_log = prog.setdefault("run_log", {})

    for local_i, chunk in enumerate(remaining_chunks):
        chunk_idx = chunk_offset + local_i
        print(f"\n--- Chunk {chunk_idx + 1}/{total_chunks} "
              f"({len(chunk)} front5 values) ---")

        for f5 in chunk:
            # Already done, unless it previously errored — errors are retried
            # on resume so a run can actually reconcile to zero errors.
            if (f5 in prog["searched_front5"]
                    and run_log.get(f5, {}).get("status") != "error"):
                continue

            results = None
            last_err = None
            attempts = 0
            for attempt in range(max_retries):
                attempts = attempt + 1
                try:
                    page.fill(SELECTORS["upc_input"], f5)
                    if SELECTORS.get("dc_dropdown"):
                        try:
                            page.select_option(SELECTORS["dc_dropdown"],
                                               label=DC_VALUE)
                        except Exception:
                            pass  # default is already All DC; not fatal
                    page.click(SELECTORS["search_button"])
                    # Let results render. networkidle is best-effort; fall
                    # back to a short settle if the site keeps a poll open.
                    try:
                        page.wait_for_load_state("networkidle", timeout=10000)
                    except Exception:
                        page.wait_for_timeout(1500)
                    results = scrape_results(page)
                    last_err = None
                    break
                except SystemExit:
                    # A header-mapping hard-stop is a "never guess" stop.
                    # It must propagate, not be retried into a logged error.
                    raise
                except Exception as e:
                    last_err = f"{type(e).__name__}: {e}"
                    if attempt < max_retries - 1:
                        backoff = 2.0 * (2 ** attempt)
                        print(f"    {f5}: request failed ({last_err}) — backing "
                              f"off {backoff:.0f}s, retry {attempt + 2}/"
                              f"{max_retries}")
                        time.sleep(backoff)

            if last_err is not None:
                run_log[f5] = {"status": "error", "rows": 0,
                               "attempts": attempts, "error": last_err}
                print(f"    {f5}: ERROR after {attempts} attempt(s) — {last_err}")
            else:
                prog["results"][f5] = results
                run_log[f5] = {"status": "hit" if results else "zero",
                               "rows": len(results), "attempts": attempts,
                               "error": ""}
                print(f"    {f5}: {len(results)} result rows"
                      if results else f"    {f5}: 0 results (zero-result)")

            if f5 not in prog["searched_front5"]:
                prog["searched_front5"].append(f5)

            since_save += 1
            if save_every and since_save >= save_every:
                save_progress(progress_path, prog)
                since_save = 0

            if delay_max > 0:
                time.sleep(random.uniform(delay_min, delay_max))

        # save progress after EACH chunk
        if chunk_idx not in prog["completed_chunks"]:
            prog["completed_chunks"].append(chunk_idx)
        save_progress(progress_path, prog)
        since_save = 0
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

def match_and_report(rows, front5_index, prog, brand_idx=None, desc_idx=None,
                     verbose_sample=8, legacy_csupc_only=False):
    """Match scraped results back to research rows.

    For each scraped row: front5 F and CsUPC -> csupc5 C5. A research row
    sharing F whose back5 == C5 is a HIT (primary match). Dedup across DCs:
    one output row per matched research row; "No Buy" = distinct values seen
    across all matching scraped rows, comma-joined.

    SECOND, SEPARATE pass — the Review Queue: confirmed 2026-08-24 that a
    scraped row's own 'UPC' column back5 can genuinely disagree with that
    same row's CsUPC (e.g. UPC '50003-79769' but CsUPC '79774'). A research
    row whose back5 matches the site UPC's back5, but never got a primary
    CsUPC hit anywhere, is NOT silently dropped — it's flagged for review.
    This is still an exact match, just against a different real field, so it
    never touches the primary (CsUPC-only) matched set — no fuzzy matching.

    THIRD, SEPARATE pass — Lookfor: for any front5 with at least one
    confirmed primary match (a proven vendor relationship), scan every
    scraped row under that front5 for items whose CsUPC has NO
    representation at all in the research file for that front5 — neither
    as a back5 match nor a site-UPC-column-back5 match (so already-Stocked
    and already-Review-Queued items are excluded, never double-counted).
    These are real items C&S carries under a vendor already proven to be
    stocked, but that never made it onto the research offer sheet at all —
    e.g. "ITO EN shows 2 matches, but C&S actually carries 7" is exactly
    what this surfaces. brand_idx (optional) pulls the vendor's own BRAND
    label from any research row sharing that front5, for display only.

    RELEVANCE FILTER, revised 2026-08-24: a front5 (manufacturer prefix) is
    not always exclusive to one vendor. JOYBA is a Del Monte brand, so its
    front5 also carries Del Monte's canned-goods line (sliced beets, green
    beans, canned peaches) — and the two lines share plenty of generic
    fruit/beverage vocabulary (FRT, BBL, MANGO, LMND), so a word-overlap
    test let that noise through into a client-facing list.

    The rule is now brand-anchored: every real C&S description observed
    leads with a brand abbreviation ('BIGLOW RED RSPBRRY', 'BTLLI EXTRA
    VIRGIN OLIVE OIL', '*JOYBA BBL RASP...', 'DM CUT GREEN BEANS'), and
    that token is what actually separates two vendors sharing a front5. A
    candidate is kept only if some token in its description identifies it
    as one of the brands the research file lists under that front5, via
    consonant-skeleton subsequence matching (see brand_token_matches) so
    C&S's vowel-dropping abbreviations still match. Word-overlap survives
    only as a fallback when the research file has no BRAND column at all.

    Nothing is silently discarded: every excluded candidate is returned in
    lookfor_rejected with the reason, surfaced in the run summary, and
    writable to an audit tab.

    Returns (no_buy_map, review_map, lookfor_list, lookfor_rejected):
      no_buy_map — research_row_index -> "No Buy" string (primary matches).
      review_map — research_row_index -> list of dicts (one per contributing
        scraped row) with front5, dc, site_upc, csupc_raw, description, pk_sz.
      lookfor_list — list of dicts, one per distinct (front5, CsUPC,
        description) item not present anywhere in the research file:
        front5, brand, description, pk_sz, type, csupc, item_codes,
        site_upcs, dcs, dcnames, no_buys, relevance.
      lookfor_rejected — same shape, for candidates the relevance filter
        excluded: front5, brand, description, csupc, reason, pk_sz, dcs,
        no_buys.
    """
    matched = {}      # row_idx -> {"no_buy": set, "pk_sz": set, "description": set}
    reviewed = {}      # row_idx -> list of contributing scraped-row dicts
    samples = []      # for validation printout

    for f5, results in prog["results"].items():
        candidates = front5_index.get(f5, [])
        if not candidates or not results:
            continue
        # Two maps, one per site field. When a row carries only a single
        # code both maps get the same value, so behavior is unchanged for
        # every file that doesn't publish a separate case GTIN.
        #   by_case_back5 -> compared against the site's CsUPC   (case code)
        #   by_item_back5 -> compared against the site's UPC col (item code)
        by_back5 = {}        # retained: Lookfor/diagnose still use it
        by_case_back5 = {}
        by_item_back5 = {}
        for ridx in candidates:
            r = rows[ridx]
            by_back5.setdefault(r["back5"], []).append(ridx)
            cb = r.get("case_back5") or r["back5"]
            ib = r.get("item_back5") or r["back5"]
            if cb:
                by_case_back5.setdefault(cb, []).append(ridx)
            if ib:
                by_item_back5.setdefault(ib, []).append(ridx)

        for res in results:
            c5 = csupc5(res.get("CsUPC", ""))
            hit_rows = by_case_back5.get(c5) if c5 else None
            if hit_rows:
                no_buy = (res.get("No Buy") or "").strip()
                pk_sz = (res.get("Pk/Sz") or "").strip()
                desc = (res.get("Description") or "").strip()
                for ridx in hit_rows:
                    m = matched.setdefault(
                        ridx, {"no_buy": set(), "pk_sz": set(), "description": set()})
                    if no_buy != "":
                        m["no_buy"].add(no_buy)
                    if pk_sz != "":
                        m["pk_sz"].add(pk_sz)
                    if desc != "":
                        m["description"].add(desc)
                    if len(samples) < verbose_sample:
                        samples.append((f5, rows[ridx], res, c5))
                continue

            # No primary CsUPC hit for this row — check the site's own UPC
            # column back5 as a second, still-exact match against a
            # different real field.
            site_b5 = site_upc_back5(res.get("UPC", ""))
            if not site_b5:
                continue
            near_hit_rows = by_item_back5.get(site_b5)
            if not near_hit_rows:
                continue
            for ridx in near_hit_rows:
                reviewed.setdefault(ridx, []).append({
                    "front5": f5,
                    "dc": res.get("DC", ""),
                    "site_upc": res.get("UPC", ""),
                    "csupc_raw": res.get("CsUPC", ""),
                    "description": res.get("Description", ""),
                    "pk_sz": res.get("Pk/Sz", ""),
                    "no_buy": res.get("No Buy", ""),
                })

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

    # ---- Merge the two exact-match tiers into one confirmed set ----
    # CORRECTED 2026-09-14, against real NEAR EAST ground truth.
    #
    # CsUPC is NOT universally the manufacturer item code. For some vendors
    # it mirrors it (ANCIENT HARVEST, front5 89125: site UPC '89125-12000',
    # CsUPC '12000'); for others it is C&S's own internally assigned case
    # code with no relation to the UPC at all (NEAR EAST, front5 72251:
    # site UPC '72251-00030', CsUPC '02044'; the CsUPCs on that vendor run
    # 02044, 02045, 02048, 02049, 02051 ... a C&S sequence, not item codes).
    #
    # The site's UPC column, by contrast, is ALWAYS manufacturer prefix +
    # item code. A back5 match against it is a full 10-digit manufacturer
    # code match — a definitive product identification, and if anything
    # STRONGER evidence than CsUPC, not weaker. The original design had the
    # hierarchy backwards.
    #
    # Treating CsUPC as the only primary rule meant that for any
    # NEAR-EAST-shaped vendor the script:
    #   1. reported 0 rows in Stocked (9 real stocked items for 72251),
    #   2. dropped every one of them from the account-manager summary,
    #      which only ever received the CsUPC-matched set, and
    #   3. failed to qualify the front5 for Lookfor at all, so the
    #      "what else does C&S carry from this vendor" list came back
    #      empty for exactly the vendors most worth asking about.
    #
    # Both tiers are now confirmed stocked. Which field matched is recorded
    # in match_type so the distinction stays visible everywhere it matters.
    review_only = set() if legacy_csupc_only else {
        ridx for ridx in reviewed if ridx not in matched}

    match_type = {ridx: "case code (CsUPC)" for ridx in matched}
    for ridx in review_only:
        match_type[ridx] = "item code (site UPC)"

    out = {}
    site_detail = {}   # ridx -> {"description": str, "pk_sz": str} — the
                       # site's OWN description/pack-size, distinct from the
                       # research row's. Exposed on the Stocked tab so a
                       # case-code match can be pack/size-verified the same
                       # way an item-code match already can via Review Queue.
    for ridx, vals in matched.items():
        out[ridx] = ", ".join(sorted(vals["no_buy"])) if vals["no_buy"] else ""
        site_detail[ridx] = {
            "description": ", ".join(sorted(vals["description"])),
            "pk_sz": ", ".join(sorted(vals["pk_sz"])),
        }
    for ridx in review_only:
        entries = reviewed[ridx]
        nb = {str(e.get("no_buy", "")).strip() for e in entries}
        nb = {v for v in nb if v}
        out[ridx] = ", ".join(sorted(nb)) if nb else ""
        descs = {str(e.get("description", "")).strip() for e in entries}
        pksz = {str(e.get("pk_sz", "")).strip() for e in entries}
        site_detail[ridx] = {
            "description": ", ".join(sorted(d for d in descs if d)),
            "pk_sz": ", ".join(sorted(p for p in pksz if p)),
        }

    # Review Queue keeps the per-DC site detail for every item-code match,
    # so they can still be eyeballed row by row. They are no longer
    # EXCLUDED from the deliverables — only labelled.
    review_out = {ridx: entries for ridx, entries in reviewed.items()
                  if ridx not in matched}

    if review_only:
        print(f"  Confirmed {len(matched)} item(s) on case code (CsUPC) and "
              f"{len(review_only)} on item code (site UPC column). Both are "
              "exact matches against real site fields and both count as "
              "stocked; --legacy-csupc-only restores the old CsUPC-only rule.")

    # ---- Third pass: Lookfor — items C&S stocks that aren't on the offer ----
    # Qualify a vendor on EITHER match tier — a front5 proven stocked only
    # by item-code matches is just as proven as one matched on CsUPC.
    qualifying_front5s = {rows[ridx]["front5"] for ridx in out.keys()}
    # Keyed by (front5, csupc) ONLY, not description — CsUPC is C&S's own
    # item code within that front5, so it's the real identity. Confirmed
    # 2026-08-24: the same physical item scraped from different DCs can
    # carry cosmetic description drift (a stray leading/trailing "*", a
    # truncated ending) that looked like duplicate rows when description
    # was part of the key. Descriptions now aggregate under the CsUPC like
    # every other cross-DC field (Pk/Sz, No Buy, etc).
    lookfor = {}  # (front5, csupc) -> aggregated dict

    rejected = {}  # (front5, csupc) -> rejection dict

    for f5 in qualifying_front5s:
        candidates = front5_index.get(f5, [])
        # Every back5 already present in the research file for this front5 —
        # regardless of match status — counts as "already on the offer."
        # Every code the offer carries under this front5 — item codes AND
        # case codes. Updated 2026-09-15: with a dual-code file, an item
        # already matched on its case code would otherwise reappear in
        # Lookfor as "C&S carries this and you don't", because only the
        # item-code set was being checked.
        offer_back5s = set()
        for ridx in candidates:
            r = rows[ridx]
            for key in ("back5", "item_back5", "case_back5"):
                v = r.get(key)
                if v:
                    offer_back5s.add(v)

        # ALL distinct brands the research file lists under this front5 —
        # not just the first. A research file can itself carry more than one
        # brand under a shared manufacturer prefix.
        brands = set()
        if brand_idx is not None:
            for ridx in candidates:
                b = str(rows[ridx]["values"][brand_idx] or "").strip()
                if b:
                    brands.add(b)
        brand = sorted(brands)[0] if brands else ""

        # Fallback vocabulary, used only when there's no BRAND column to
        # anchor on. Weaker (generic grocery words overlap across unrelated
        # lines), so it is never the primary rule.
        reference_vocab = set()
        if desc_idx is not None:
            for ridx in candidates:
                reference_vocab |= significant_tokens(rows[ridx]["values"][desc_idx])

        for res in prog["results"].get(f5, []):
            c5 = csupc5(res.get("CsUPC", ""))
            if not c5 or c5 in offer_back5s:
                continue
            site_b5 = site_upc_back5(res.get("UPC", ""))
            if site_b5 and site_b5 in offer_back5s:
                continue  # already surfaced via Stocked or Review Queue
            description = (res.get("Description") or "").strip()

            # Relevance gate. Primary rule is the brand token that leads
            # every real C&S description; vocabulary is only a fallback.
            if brands:
                hit = matched_brand_word(description, brands)
                relevant = hit is not None
                reason = (f"brand:{hit}" if relevant
                          else f"no brand match ({'/'.join(sorted(brands))})")
            else:
                shared = significant_tokens(description) & reference_vocab
                relevant = bool(shared)
                reason = ("vocab:" + ",".join(sorted(shared)) if relevant
                          else "no shared vocabulary (no BRAND column)")

            if not relevant:
                rkey = (f5, c5)
                rej = rejected.setdefault(rkey, {
                    "front5": f5, "brand": brand, "descriptions": set(),
                    "csupc": c5, "reasons": set(), "pk_sz": set(),
                    "dcs": set(), "no_buys": set(),
                })
                if description:
                    rej["descriptions"].add(description)
                rej["reasons"].add(reason)
                if res.get("Pk/Sz"):
                    rej["pk_sz"].add(res["Pk/Sz"])
                if res.get("DC"):
                    rej["dcs"].add(res["DC"])
                if res.get("No Buy"):
                    rej["no_buys"].add(res["No Buy"])
                continue

            key = (f5, c5)
            entry = lookfor.setdefault(key, {
                "front5": f5, "brand": brand, "descriptions": set(),
                "pk_sz": set(), "type": set(), "csupc": c5,
                "item_codes": set(), "site_upcs": set(), "dcs": set(),
                "dcnames": set(), "no_buys": set(), "relevance": set(),
            })
            if description:
                entry["descriptions"].add(description)
            entry["relevance"].add(reason)
            if res.get("Pk/Sz"):
                entry["pk_sz"].add(res["Pk/Sz"])
            if res.get("Type"):
                entry["type"].add(res["Type"])
            if res.get("ItemCode"):
                entry["item_codes"].add(str(res["ItemCode"]))
            if res.get("UPC"):
                entry["site_upcs"].add(res["UPC"])
            if res.get("DC"):
                entry["dcs"].add(res["DC"])
            if res.get("DcName"):
                entry["dcnames"].add(res["DcName"])
            if res.get("No Buy"):
                entry["no_buys"].add(res["No Buy"])

    # Pick one canonical description per item for display (the longest —
    # avoids showing a truncated variant like "4P" when "4PK" was also
    # seen), while keeping the full set available for transparency.
    for entry in lookfor.values():
        entry["description"] = (max(entry["descriptions"], key=len)
                                if entry["descriptions"] else "")
    for entry in rejected.values():
        entry["description"] = (max(entry["descriptions"], key=len)
                                if entry["descriptions"] else "")

    lookfor_list = sorted(lookfor.values(), key=lambda e: (e["front5"], e["csupc"]))
    lookfor_rejected = sorted(rejected.values(),
                              key=lambda e: (e["front5"], e["csupc"]))

    if lookfor_rejected:
        print(f"\n  Lookfor relevance filter: kept {len(lookfor_list)}, "
              f"excluded {len(lookfor_rejected)} item(s) whose description "
              "did not identify them as the front5's own brand "
              "(a front5 can be shared by unrelated vendors).")
        print("  Re-run with --lookfor-audit to see every excluded item and why.")

    return out, review_out, lookfor_list, lookfor_rejected, match_type, site_detail


def diagnose_matching(rows, front5_index, prog):
    """Root-cause a zero- (or low-) match run without guessing.

    For each searched front5 that returned scraped rows, prints the
    candidate research back5 values side by side with the scraped CsUPC
    values (raw as scraped, and normalized via csupc5), plus the overlap
    count. This makes it visible whether the two sets are simply disjoint
    (genuinely not stocked) or whether they look shifted/misaligned (a real
    bug), instead of assuming either without evidence.
    """
    print("\n" + "=" * 70)
    print("DIAGNOSE — front5 candidates vs. scraped CsUPC")
    print("=" * 70)

    total_candidates = 0
    total_scraped_rows = 0
    total_overlap = 0
    front5_with_results = 0

    for f5 in sorted(prog["results"].keys()):
        results = prog["results"][f5]
        if not results:
            continue
        front5_with_results += 1
        candidates = front5_index.get(f5, [])
        cand_back5 = sorted({rows[i]["back5"] for i in candidates})
        raw_csupc = sorted({r.get("CsUPC", "") for r in results})
        norm_csupc = sorted({csupc5(r.get("CsUPC", "")) for r in results})
        overlap = sorted(set(cand_back5) & set(norm_csupc))

        total_candidates += len(candidates)
        total_scraped_rows += len(results)
        total_overlap += len(overlap)

        print(f"\nfront5 {f5}: {len(candidates)} research candidates, "
              f"{len(results)} scraped rows, {len(raw_csupc)} distinct CsUPC")
        print(f"  research back5 (sample):   {cand_back5[:8]}")
        print(f"  scraped CsUPC raw (sample): {raw_csupc[:8]}")
        print(f"  scraped CsUPC normalized:   {norm_csupc[:8]}")
        print(f"  overlap: {len(overlap)} {overlap[:8]}")

    print("\n" + "=" * 70)
    print(f"TOTALS: {front5_with_results} front5 had scraped results, "
          f"{total_candidates} total research candidates across them, "
          f"{total_scraped_rows} total scraped rows, "
          f"{total_overlap} total back5/CsUPC overlaps found.")
    print("=" * 70)
    if total_overlap == 0 and total_scraped_rows > 0 and total_candidates > 0:
        print(
            "\nZero overlap despite real candidates AND real scraped data on "
            "both sides — the back5/CsUPC formats likely don't line up "
            "(e.g. a digit-count or offset mismatch). Compare a 'research "
            "back5' sample above against its 'scraped CsUPC raw' sample by "
            "eye. Do NOT guess a fix — report exactly what you see back."
        )


def helper_col_fixups(headers):
    """Indices of any pre-computed UPC12/FRONT5/BACK5 helper columns, mapped
    to the row key holding this script's own confirmed value.

    These columns are carried through to the output because the deliverable
    preserves the input's column order — but a BACK5 computed with the old
    last-5 formula is wrong, and shipping a known-wrong number into an
    account-manager document is worse than shipping one that differs from
    the input. Output tabs therefore show the confirmed derivation. The raw
    input file is never modified. Pass keep_source_helpers=True (CLI:
    --keep-source-back5) to carry the original values through untouched.
    """
    out = {}
    for col, key in (("UPC12", "pad12"), ("FRONT5", "front5"), ("BACK5", "back5")):
        idx = find_header_index(headers, col)
        if idx is not None:
            out[idx] = key
    return out


def write_output(headers, rows, no_buy_map, review_map, lookfor_list, out_path,
                 lookfor_rejected=None, run_log=None, unique_front5=None,
                 prep_audit=None, keep_source_helpers=False, match_type=None,
                 site_detail=None):
    """Write the matches-only 'Stocked' tab, a 'Review Queue' tab, a
    'Stocked Vendor Lines' tab, and a 'Lookfor' tab.

    Stocked: original columns + 'No Buy', matches-only, primary CsUPC-exact
    hits only — unaffected by anything in the review queue.

    Review Queue: research rows where the site's own UPC-column back5 (a
    different real field than CsUPC) matched, but no scraped row's CsUPC
    ever did — so they'd otherwise be silently dropped. Never merged into
    Stocked; the primary match rule stays exact-CsUPC-only.

    Stocked Vendor Lines: every original research row (matched or not, in
    original file order) whose BRAND has at least one confirmed Stocked
    match — the "load the whole line" view, since a vendor C&S carries even
    one item from is worth considering in full for other clients. Based on
    the confirmed Stocked tab only, not the Review Queue (still unconfirmed).
    Each row is marked whether it was itself one of the confirmed hits.

    Lookfor: items C&S actually carries under a vendor already proven
    stocked, that never made it onto the research offer sheet at all —
    computed in match_and_report(), not from research rows (there are none
    to pull from — that's the point). No original headers apply here; its
    own column set is used instead.
    """
    upc_idx = find_header_index(headers, "UPC")
    fixups = {} if keep_source_helpers else helper_col_fixups(headers)

    def display_values(row):
        """Row values with the UPC cell rendered as a readable dashed UPC-A
        (e.g. '0-72310-00041-4') instead of a raw digit string. Display-only
        — never used for matching, and never touches the source file.

        Any pre-computed UPC12/FRONT5/BACK5 helper column is also refreshed
        from this script's own confirmed derivation, so the output never
        carries a stale last-5 BACK5 forward."""
        vals = list(row["values"])
        if upc_idx is not None:
            vals[upc_idx] = format_upc_display(row.get("pad12", ""), vals[upc_idx])
        for idx, key in fixups.items():
            vals[idx] = row.get(key, "")
        return vals

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stocked"
    # "Matched On" says which real site field confirmed the item: the case
    # code (CsUPC) or the item code (the site's own UPC column). Both are
    # exact matches; CsUPC is simply not the item code for every vendor.
    mt = match_type or {}
    sd = site_detail or {}
    ws.append(list(headers) + ["No Buy", "Matched On", "Site Description",
                                "Site Pk/Sz"])

    n = 0
    for ridx in sorted(no_buy_map.keys()):
        detail = sd.get(ridx, {})
        ws.append(display_values(rows[ridx])
                  + [no_buy_map[ridx], mt.get(ridx, "case code (CsUPC)"),
                     detail.get("description", ""), detail.get("pk_sz", "")])
        n += 1

    ws2 = wb.create_sheet("Review Queue")
    ws2.append(list(headers) + ["Front5", "DC(s) seen", "Site UPC(s) seen",
                                 "CsUPC(s) seen", "Site Description(s)",
                                 "Site Pk/Sz(s)", "No Buy(s) seen"])
    for ridx in sorted(review_map.keys()):
        entries = review_map[ridx]
        front5s = ", ".join(sorted({e["front5"] for e in entries}))
        dcs = ", ".join(sorted({e["dc"] for e in entries if e["dc"]}))
        site_upcs = ", ".join(sorted({e["site_upc"] for e in entries if e["site_upc"]}))
        csupcs = ", ".join(sorted({str(e["csupc_raw"]) for e in entries if e["csupc_raw"]}))
        descs = ", ".join(sorted({e["description"] for e in entries if e["description"]}))
        pksz = ", ".join(sorted({e["pk_sz"] for e in entries if e["pk_sz"]}))
        no_buys = ", ".join(sorted({e["no_buy"] for e in entries if e["no_buy"]}))
        ws2.append(display_values(rows[ridx]) +
                   [front5s, dcs, site_upcs, csupcs, descs, pksz, no_buys])

    n_vendor_lines = 0
    brand_idx = find_header_index(headers, "BRAND")
    if brand_idx is None:
        print("  WARNING: no 'BRAND' column found — skipping 'Stocked Vendor "
              f"Lines' tab. Headers seen: {headers}")
    else:
        qualifying_brands = {
            str(rows[ridx]["values"][brand_idx] or "").strip().upper()
            for ridx in no_buy_map.keys()
        }
        qualifying_brands.discard("")

        ws3 = wb.create_sheet("Stocked Vendor Lines")
        ws3.append(list(headers) + ["Confirmed Stocked", "No Buy"])
        for ridx, row in enumerate(rows):
            brand = str(row["values"][brand_idx] or "").strip().upper()
            if brand not in qualifying_brands:
                continue
            is_hit = ridx in no_buy_map
            ws3.append(display_values(row) +
                       ["Yes" if is_hit else "", no_buy_map.get(ridx, "")])
            n_vendor_lines += 1

    ws4 = wb.create_sheet("Lookfor")
    ws4.append(["Front5", "Brand", "Description", "UPC = CsUPC", "Pk/Sz",
                "Type", "CsUPC", "Site ItemCode(s)", "Site UPC(s) raw",
                "DC(s)", "DC Name(s)", "No Buy(s)", "Matched On"])
    for entry in lookfor_list:
        ws4.append([
            entry["front5"], entry["brand"], entry["description"],
            f"{entry['front5']}-{entry['csupc']}",
            ", ".join(sorted(entry["pk_sz"])),
            ", ".join(sorted(entry["type"])),
            entry["csupc"],
            ", ".join(sorted(entry["item_codes"])),
            ", ".join(sorted(entry["site_upcs"])),
            ", ".join(sorted(entry["dcs"])),
            ", ".join(sorted(entry["dcnames"])),
            ", ".join(sorted(entry["no_buys"])),
            ", ".join(sorted(entry["relevance"])),
        ])

    # Optional audit tab — every candidate the relevance filter excluded,
    # with the reason. Keeps the filter verifiable instead of a black box.
    if lookfor_rejected:
        ws5 = wb.create_sheet("Lookfor Audit (excluded)")
        ws5.append(["Front5", "Research Brand", "Site Description", "Pk/Sz",
                    "CsUPC", "DC(s)", "No Buy(s)", "Excluded Because"])
        for entry in lookfor_rejected:
            ws5.append([
                entry["front5"], entry["brand"], entry["description"],
                ", ".join(sorted(entry["pk_sz"])),
                entry["csupc"],
                ", ".join(sorted(entry["dcs"])),
                ", ".join(sorted(entry["no_buys"])),
                ", ".join(sorted(entry["reasons"])),
            ])

    # Run Log — every front5 accounted for, with hit / zero-result / error
    # kept distinguishable. A search that returned nothing and a search that
    # failed must never be indistinguishable in the record.
    if unique_front5 is not None:
        write_run_log_tab(wb, run_log or {}, unique_front5, prep_audit)

    wb.save(out_path)
    return n, len(review_map), n_vendor_lines, len(lookfor_list)


def run_log_tally(run_log, unique_front5):
    """Reconcile the run log against the full front5 search list.

    Returns (tally, per_front5) where tally counts hit / zero / error /
    not-attempted and sums to len(unique_front5) by construction."""
    tally = {"hit": 0, "zero": 0, "error": 0, "not attempted": 0}
    per = []
    for f5 in unique_front5:
        entry = (run_log or {}).get(f5)
        status = entry.get("status") if entry else "not attempted"
        if status not in tally:
            status = "not attempted"
        tally[status] += 1
        per.append((f5, status, entry or {}))
    return tally, per


def write_run_log_tab(wb, run_log, unique_front5, prep_audit=None):
    """One row per front5 in the search list, plus a reconciliation block.

    Ordered errors first, then zero-results, then hits, so the things that
    need attention are at the top instead of buried in 200+ successful rows.
    """
    ws = wb.create_sheet("Run Log")
    tally, per = run_log_tally(run_log, unique_front5)

    ws.append(["RECONCILIATION"])
    ws["A1"].font = Font(bold=True)
    for label in ("hit", "zero", "error", "not attempted"):
        ws.append([label, tally[label]])
    ws.append(["TOTAL", sum(tally.values())])
    ws.append(["FRONT5 values in search list", len(unique_front5)])
    ws.append([])

    if prep_audit:
        ws.append(["INPUT FILE HELPER-COLUMN AUDIT"])
        ws.cell(ws.max_row, 1).font = Font(bold=True)
        ws.append(["Column", "Rows", "Disagree with confirmed rule", "Note"])
        for col, info in prep_audit.items():
            note = ""
            if info["disagree"] and info.get("all_legacy_last5"):
                note = ("Uses the disproved last-5 [7:12] formula; ignored — "
                        "back5 recomputed as [6:11] for all matching.")
            elif info["disagree"]:
                note = "Disagrees with the confirmed rule; ignored for matching."
            ws.append([col, info["rows"], info["disagree"], note])
        ws.append([])

    ws.append(["FRONT5", "Status", "Result Rows", "Attempts", "Error"])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    order = {"error": 0, "not attempted": 1, "zero": 2, "hit": 3}
    for f5, status, entry in sorted(per, key=lambda t: (order[t[1]], t[0])):
        ws.append([f5, status, entry.get("rows", 0),
                   entry.get("attempts", 0), entry.get("error", "")])
    return tally


def format_pack_size(pack, size, uos) -> str:
    """Combine the research file's separate PACK/SIZE/UOS columns into one
    readable field, e.g. '12/15.50 FO' — matching the site's own Pk/Sz
    display convention. Blank pieces are dropped rather than shown as
    'None'."""
    pack_s = str(pack).strip() if pack not in (None, "") else ""
    size_s = str(size).strip() if size not in (None, "") else ""
    uos_s = str(uos).strip() if uos not in (None, "") else ""
    left = f"{pack_s}/{size_s}" if pack_s and size_s else (pack_s or size_s)
    return f"{left} {uos_s}".strip() if uos_s else left


def write_share_summary(headers, rows, no_buy_map, lookfor_list, out_path,
                        match_type=None):
    """Write a clean, standalone, account-manager-ready summary — a
    SEPARATE file from the internal 4-tab workbook, since the internal
    matching/comparison tabs (Review Queue, Stocked Vendor Lines, the raw
    'Lookfor' diagnostics) aren't meant for external sharing.

    Rows: Stocked (confirmed "On Offer") + Lookfor (confirmed "Ask Source").

    CORRECTED 2026-09-14: item-code (site UPC column) matches used to be
    excluded from this file entirely, because they lived only in the Review
    Queue and this function was never passed it. That silently dropped every
    stocked item from any vendor whose CsUPC is a C&S-assigned case code
    rather than the manufacturer item code (NEAR EAST being the case that
    exposed it) — exactly the items a client quote needs. Both tiers now
    flow through, with "Matched On" naming the field that confirmed each.

    Columns: Brand, Description, Pack/Size, Your Cost, List Price,
    % Spread, Status. Your Cost/List Price/% Spread are blank for "Ask
    Source" rows — there's no pricing for items that were never on the
    offer. Sorted by Brand then Description for easy scanning.
    """
    brand_idx = find_header_index(headers, "BRAND")
    desc_idx = find_header_index(headers, "DESCRIPTION")
    upc_idx = find_header_index(headers, "UPC")
    pack_idx = find_header_index(headers, "PACK")
    size_idx = find_header_index(headers, "SIZE")
    uos_idx = find_header_index(headers, "UOS")
    cost_idx = find_header_index(headers, "YOUR COST")
    list_idx = find_header_index(headers, "LIST PRICE")

    missing = [name for name, idx in [
        ("BRAND", brand_idx), ("DESCRIPTION", desc_idx), ("UPC", upc_idx),
        ("PACK", pack_idx), ("SIZE", size_idx), ("YOUR COST", cost_idx),
        ("LIST PRICE", list_idx),
    ] if idx is None]
    if missing:
        print(f"  WARNING: research file is missing column(s) {missing} — "
              "skipping the account-manager summary file. "
              f"Headers seen: {headers}")
        return 0

    # (brand, description, upc, pack_size, cost, list_price, spread_pct, status)
    out_rows = []

    for ridx in no_buy_map.keys():
        vals = rows[ridx]["values"]
        cost = vals[cost_idx]
        list_price = vals[list_idx]
        spread_pct = None
        try:
            cost_f = float(cost)
            list_f = float(list_price)
            if list_f:
                spread_pct = round((list_f - cost_f) / list_f * 100, 2)
        except (TypeError, ValueError):
            pass
        upc_display = format_upc_display(rows[ridx].get("pad12", ""), vals[upc_idx])
        out_rows.append((
            vals[brand_idx], vals[desc_idx], upc_display,
            format_pack_size(vals[pack_idx], vals[size_idx],
                              vals[uos_idx] if uos_idx is not None else ""),
            cost, list_price, spread_pct, "On Offer",
            (match_type or {}).get(ridx, "case code (CsUPC)"),
        ))

    for entry in lookfor_list:
        # No research-file UPC exists for a Lookfor item (that's the point —
        # it was never on the offer). Use the site's own front5-CsUPC pair,
        # the same identifying number C&S itself displays, so the source
        # has something concrete to look the item up by.
        out_rows.append((
            entry["brand"], entry["description"],
            f"{entry['front5']}-{entry['csupc']}",
            ", ".join(sorted(entry["pk_sz"])),
            "", "", "", "Ask Source", "",
        ))

    out_rows.sort(key=lambda r: (str(r[0] or "").upper(), str(r[1] or "").upper()))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CS Stock Summary"
    ws.append(["Brand", "Description", "UPC = CsUPC", "Pack/Size", "Your Cost",
               "List Price", "% Spread", "Status", "Matched On"])
    for row in out_rows:
        ws.append(list(row))

    wb.save(out_path)
    return len(out_rows)


def _no_buy_is_clean_no(no_buy_str, strict=False) -> bool:
    """True if the item is confirmed sellable — Final Review's inclusion
    rule for No Buy.

    CORRECTED 2026-09-14, at the user's explicit direction: default is now
    'No' at ANY DC that matched is sufficient — 'a DC saying No Buy = No
    means I can sell this item there', regardless of what other DCs said.
    The original 2026-08-24 rule required EVERY DC's value to be exactly
    'No' and dropped the entire item on a single stray 'Yes' elsewhere —
    stricter than the user's stated criteria (UPC matches, pack/size
    match, No Buy = No) and it was silently discarding genuinely sellable
    items from the account-manager document over an unrelated DC's policy.
    A blank/no-value row is still excluded either way — no evidence isn't
    evidence of 'No'. Pass strict=True (CLI: --strict-no-buy) to restore
    the original unanimous-No-only rule."""
    vals = {v.strip().lower() for v in str(no_buy_str or "").split(",") if v.strip()}
    if not vals:
        return False
    return vals == {"no"} if strict else "no" in vals


def _review_no_buy_is_clean_no(entries, strict=False) -> bool:
    """Same rule as _no_buy_is_clean_no, applied to a Review Queue row's
    list of contributing scraped-row dicts instead of a pre-joined string."""
    vals = {e.get("no_buy", "").strip().lower() for e in entries if e.get("no_buy", "").strip()}
    if not vals:
        return False
    return vals == {"no"} if strict else "no" in vals


def write_final_review(headers, rows, no_buy_map, review_map, out_path,
                       keep_source_helpers=False, match_type=None,
                       strict_no_buy=False):
    """Write the account-manager-ready Final Review workbook — the shape
    proven out by hand on the real Coffees & Teas run and confirmed
    2026-08-24: same columns as the research file (no added columns, since
    the account manager is used to seeing the source document and extra
    analytical columns confuse rather than help), grouped by Brand with a
    blank row between brand groups, one bold/plain visual marker in place
    of a Status column.

    CORRECTED 2026-09-14, at the user's explicit direction, superseding the
    earlier same-day "bold = case code, plain = item code" mapping below.
    The user's own rule: bold = "populates on the Product List" (i.e. was
    actually matched — either tier, doesn't matter which field proved it);
    plain = "brand is proven stocked, but this exact item never appeared in
    a search result." Since Final Review stays MATCHES-ONLY (the user
    explicitly declined to widen it to the full vendor line — that stays on
    the separate Stocked Vendor Lines tab), every row that remains in scope
    here was, by definition, scraped from a real Product List response.
    There is therefore no "plain" case left to render: every included row
    is bold. bold/plain no longer carries a which-field or confidence
    signal at all — that distinction still lives in the Stocked tab's
    "Matched On" column for anyone who wants it.

    Prior (superseded 2026-09-14) mapping, kept for history:
      - "case code" match (definite)  -> the primary Stocked match, CsUPC
        exact (back5 == CsUPC). Was rendered BOLD.
      - "item code" match (secondary) -> the Review Queue match, site UPC
        column back5 only. Was rendered PLAIN.

    Both tiers are filtered to rows where No Buy is unambiguously "No"
    (see _no_buy_is_clean_no) — a row with a Yes/No discrepancy across DCs,
    or no No Buy value seen at all, is excluded rather than guessed at.

    This is a SEPARATE, narrower document from the internal Stocked/Review
    Queue/Lookfor workbook and the CS_STOCK_SUMMARY file — it never includes
    Lookfor ("ask for it") items, which the user treats as a distinct
    second-layer document, not part of this one.
    """
    brand_idx = find_header_index(headers, "BRAND")
    desc_idx = find_header_index(headers, "DESCRIPTION")
    upc_idx = find_header_index(headers, "UPC")
    fixups = {} if keep_source_helpers else helper_col_fixups(headers)

    def display_values(row):
        vals = list(row["values"])
        if upc_idx is not None:
            vals[upc_idx] = format_upc_display(row.get("pad12", ""), vals[upc_idx])
        for idx, key in fixups.items():
            vals[idx] = row.get(key, "")
        return vals

    # CORRECTED 2026-09-14: every row remaining in Final Review (matches-
    # only, per the user's explicit choice) was scraped from a real Product
    # List response regardless of tier — bold now means exactly that, so
    # every included row is bold. match_type / "Matched On" (still on the
    # Stocked tab) is the place to see which field actually matched.
    included = []  # (row_idx, bold) — bold is always True now; kept as a
                   # tuple for the render loop below, which still branches
                   # on it.
    for ridx in range(len(rows)):
        if ridx in no_buy_map:
            if _no_buy_is_clean_no(no_buy_map[ridx], strict=strict_no_buy):
                included.append((ridx, True))
        elif ridx in review_map:
            if _review_no_buy_is_clean_no(review_map[ridx], strict=strict_no_buy):
                included.append((ridx, True))

    if brand_idx is None:
        print("  WARNING: no 'BRAND' column found — Final Review rows will "
              "not be grouped by brand.")

    def sort_key(item):
        ridx, _ = item
        vals = rows[ridx]["values"]
        brand = str(vals[brand_idx] or "").strip().upper() if brand_idx is not None else ""
        desc = str(vals[desc_idx] or "").strip().upper() if desc_idx is not None else ""
        return (brand, desc)

    included.sort(key=sort_key)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Final Review"
    ws.append(list(headers))
    bold_font = Font(bold=True)

    last_brand = None
    n_bold = 0
    n_plain = 0
    for ridx, bold in included:
        brand = (str(rows[ridx]["values"][brand_idx] or "").strip()
                 if brand_idx is not None else "")
        if brand_idx is not None and last_brand is not None and brand != last_brand:
            ws.append([None] * len(headers))
        last_brand = brand

        ws.append(display_values(rows[ridx]))
        if bold:
            for cell in ws[ws.max_row]:
                cell.font = bold_font
            n_bold += 1
        else:
            n_plain += 1

    wb.save(out_path)
    return n_bold, n_plain


# ═══════════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════════

def default_out_path(input_path):
    base, ext = os.path.splitext(input_path)
    return f"{base}_STOCKED{ext or '.xlsx'}"


def default_progress_path(input_path):
    base, _ = os.path.splitext(input_path)
    return f"{base}_progress.json"


def default_share_path(input_path):
    base, ext = os.path.splitext(input_path)
    return f"{base}_CS_STOCK_SUMMARY{ext or '.xlsx'}"


def default_final_review_path(input_path):
    base, ext = os.path.splitext(input_path)
    return f"{base}_FINAL_REVIEW{ext or '.xlsx'}"


def main():
    ap = argparse.ArgumentParser(
        description="C&S (divert.cssourcing.com) stock-check automation.")
    ap.add_argument("--input", required=True, help="Research .xlsx file.")
    ap.add_argument("--out", default=None,
                    help="Internal output workbook (default: <input>_STOCKED.xlsx).")
    ap.add_argument("--share-out", default=None,
                    help="Account-manager-ready summary file (default: "
                         "<input>_CS_STOCK_SUMMARY.xlsx).")
    ap.add_argument("--final-out", default=None,
                    help="Final Review workbook: source columns, filtered to "
                         "No Buy == 'No' with no cross-DC discrepancy, bold "
                         "for case-code (CsUPC) matches, plain for item-code "
                         "(Review Queue) matches (default: "
                         "<input>_FINAL_REVIEW.xlsx).")
    ap.add_argument("--progress", default=None,
                    help="Progress JSON path (default: <input>_progress.json).")
    ap.add_argument("--chunks", type=int, default=5,
                    help="Number of chunks to split front5 values into (default 5).")
    ap.add_argument("--delay-min", type=float, default=1.0,
                    help="Minimum courtesy delay between searches, seconds "
                         "(default 1.0). Use 0 with --delay-max 0 for full speed.")
    ap.add_argument("--delay-max", type=float, default=3.0,
                    help="Maximum courtesy delay between searches, seconds "
                         "(default 3.0). Set to 0 to disable the delay entirely; "
                         "backoff on a failed request still applies.")
    ap.add_argument("--full-speed", action="store_true",
                    help="Shorthand for --delay-min 0 --delay-max 0 --chunks 1: "
                         "no courtesy delay and a single chunk. Progress is "
                         "still saved every --save-every searches, so the run "
                         "stays resumable.")
    ap.add_argument("--strict-no-buy", action="store_true",
                    help="Restore the pre-2026-09-14 Final Review rule: "
                         "include an item only if No Buy is 'No' at EVERY "
                         "DC it matched, dropping it entirely on a single "
                         "stray 'Yes' elsewhere. Default (recommended) is "
                         "'No' at ANY matching DC confirms it's sellable "
                         "there, per the user's stated criteria.")
    ap.add_argument("--legacy-csupc-only", action="store_true",
                    help="Restore the pre-2026-09-14 rule: confirm a match "
                         "ONLY on CsUPC, and treat a site-UPC-column match as "
                         "an unconfirmed Review Queue item excluded from the "
                         "account-manager summary. This misses every stocked "
                         "item from a vendor whose CsUPC is a C&S-assigned "
                         "case code rather than the manufacturer item code "
                         "(e.g. NEAR EAST). Provided only to reproduce an "
                         "earlier run.")
    ap.add_argument("--keep-source-back5", action="store_true",
                    help="Carry any pre-computed UPC12/FRONT5/BACK5 helper "
                         "columns from the input through to the output "
                         "unchanged. Default is to refresh them from this "
                         "script's own confirmed derivation, so a stale "
                         "last-5 BACK5 is not shipped onward. Matching never "
                         "reads these columns either way.")
    ap.add_argument("--save-every", type=int, default=25,
                    help="Save progress every N searches within a chunk "
                         "(default 25). 0 = only save at chunk boundaries.")
    ap.add_argument("--max-retries", type=int, default=3,
                    help="Attempts per front5 before it is logged as an error "
                         "(default 3), with 2s/4s exponential backoff between.")
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
    ap.add_argument("--lookfor-audit", action="store_true",
                    help="Add a 'Lookfor Audit (excluded)' tab listing every "
                         "candidate the relevance filter excluded and why — "
                         "use it to verify nothing legitimate was dropped.")
    ap.add_argument("--diagnose", action="store_true",
                    help="No browser, no output file. Reads the saved progress "
                         "file and prints, per searched front5, the candidate "
                         "research back5 values vs. the scraped CsUPC values "
                         "side by side, so a zero-match run can be root-caused "
                         "instead of guessed at.")
    args = ap.parse_args()

    if args.full_speed:
        args.delay_min = 0.0
        args.delay_max = 0.0
        args.chunks = 1

    out_path = args.out or default_out_path(args.input)
    share_path = args.share_out or default_share_path(args.input)
    final_path = args.final_out or default_final_review_path(args.input)
    progress_path = args.progress or default_progress_path(args.input)

    print("=" * 70)
    print("C&S DIVERT STOCK CHECK")
    print("=" * 70)

    # ---- Read research file, build front5 chunks ----
    print("\nReading research file...")
    headers, rows, _, prep_audit = read_research(args.input)
    front5_index = build_front5_index(rows)
    unique_front5 = sorted(front5_index.keys())
    print(f"  {len(unique_front5)} distinct front5 values across "
          f"{len(rows)} research rows.")

    chunks = chunk_list(unique_front5, args.chunks)
    print(f"  Split into {len(chunks)} chunk(s); courtesy delay "
          f"{args.delay_min:.1f}-{args.delay_max:.1f}s"
          f"{' (full speed)' if args.delay_max <= 0 else ''}.")
    brand_idx = find_header_index(headers, "BRAND")
    desc_idx = find_header_index(headers, "DESCRIPTION")

    # ---- Diagnose-only path (no browser, no output file) ----
    if args.diagnose:
        prog = load_progress(progress_path, chunks, args.input,
                             require_chunk_layout=False)
        if not prog:
            raise SystemExit("ERROR: no valid progress file to diagnose.")
        diagnose_matching(rows, front5_index, prog)
        return

    # ---- Rebuild-output-only path (no browser) ----
    if args.rebuild_output:
        prog = load_progress(progress_path, chunks, args.input,
                             require_chunk_layout=False)
        if not prog:
            raise SystemExit("ERROR: no valid progress file to rebuild from.")
        (no_buy_map, review_map, lookfor_list, lookfor_rejected,
         match_type, site_detail) = match_and_report(
            rows, front5_index, prog, brand_idx=brand_idx, desc_idx=desc_idx,
            legacy_csupc_only=args.legacy_csupc_only)
        n, n_review, n_vendor, n_lookfor = write_output(
            headers, rows, no_buy_map, review_map, lookfor_list, out_path,
            lookfor_rejected=lookfor_rejected if args.lookfor_audit else None,
            run_log=prog.get("run_log"), unique_front5=unique_front5,
            prep_audit=prep_audit,
            keep_source_helpers=args.keep_source_back5,
            match_type=match_type, site_detail=site_detail)
        print(f"\nWrote {n} matched rows, {n_review} Review Queue rows, "
              f"{n_vendor} Stocked Vendor Lines rows, and {n_lookfor} "
              f"Lookfor rows to {out_path}")
        n_share = write_share_summary(headers, rows, no_buy_map, lookfor_list,
                                      share_path, match_type=match_type)
        if n_share:
            print(f"Wrote {n_share}-row account-manager summary to {share_path}")
        n_bold, n_plain = write_final_review(
            headers, rows, no_buy_map, review_map, final_path,
            keep_source_helpers=args.keep_source_back5,
            match_type=match_type, strict_no_buy=args.strict_no_buy)
        print(f"Wrote Final Review to {final_path}: {n_bold + n_plain} row(s) "
              "(every row bold — all populate on the Product List; "
              "'Matched On' on the Stocked tab still shows which field "
              "proved each one).")
        print_run_summary(prog.get("run_log"), unique_front5, n, n_review,
                          finished=None)
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
            "run_log": {},
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
            # results_table can only be discovered once real search results
            # are on screen — make that an explicit, separate, impossible-to-
            # skip step instead of folding it into "navigate to Product
            # Search", which people (reasonably) read as just "get to the
            # page", not "run a search on it."
            input(
                "\n>>> INSPECT MODE — before continuing, RUN AN ACTUAL SEARCH "
                "in the browser window now:\n"
                "      1. Click the 'Upc' box and type a UPC or partial UPC "
                "(e.g. 10095)\n"
                "      2. Click the 'Search' button\n"
                "      3. WAIT until the Product List grid shows rows on "
                "screen\n"
                "    Only once you can SEE result rows in the browser, come "
                "back here and press Enter... "
            )
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
                args.stop_after_chunk, delay_min=args.delay_min,
                delay_max=args.delay_max, save_every=args.save_every,
                max_retries=args.max_retries)

        browser.close()

    # ---- Match + write output (always, so partial progress is usable) ----
    (no_buy_map, review_map, lookfor_list, lookfor_rejected,
     match_type, site_detail) = match_and_report(
        rows, front5_index, prog, brand_idx=brand_idx, desc_idx=desc_idx,
        legacy_csupc_only=args.legacy_csupc_only)
    n, n_review, n_vendor, n_lookfor = write_output(
        headers, rows, no_buy_map, review_map, lookfor_list, out_path,
        lookfor_rejected=lookfor_rejected if args.lookfor_audit else None,
        run_log=prog.get("run_log"), unique_front5=unique_front5,
        prep_audit=prep_audit,
        keep_source_helpers=args.keep_source_back5,
        match_type=match_type, site_detail=site_detail)
    print(f"\nWrote {n} matched rows, {n_review} Review Queue rows, "
          f"{n_vendor} Stocked Vendor Lines rows, and {n_lookfor} "
          f"Lookfor rows to {out_path}")
    n_share = write_share_summary(headers, rows, no_buy_map, lookfor_list,
                                  share_path, match_type=match_type)
    if n_share:
        print(f"Wrote {n_share}-row account-manager summary to {share_path}")
    n_bold, n_plain = write_final_review(
        headers, rows, no_buy_map, review_map, final_path,
        keep_source_helpers=args.keep_source_back5,
        match_type=match_type, strict_no_buy=args.strict_no_buy)
    print(f"Wrote Final Review to {final_path}: {n_bold + n_plain} row(s) "
          "(every row bold — all populate on the Product List; "
          "'Matched On' on the Stocked tab still shows which field "
          "proved each one).")
    if n_review:
        print(f"  {n_review} item(s) need a quick manual look in the "
              "'Review Queue' tab — the site's UPC column matched but "
              "CsUPC never did, so they weren't auto-confirmed as stocked.")
    # ---- Final reconciliation — must add up to the full search list ----
    print_run_summary(prog.get("run_log"), unique_front5, n, n_review, finished)


def print_run_summary(run_log, unique_front5, n_matched, n_review, finished):
    """Print the end-of-run reconciliation. Never reports success while any
    front5 errored or was never attempted."""
    tally, _ = run_log_tally(run_log, unique_front5)
    n_err = tally["error"]
    n_missing = tally["not attempted"]
    print("\n" + "=" * 70)
    print("RUN SUMMARY")
    print("=" * 70)
    print(f"  FRONT5 values in search list : {len(unique_front5)}")
    print(f"  Searches with results (hit)  : {tally['hit']}")
    print(f"  Searches with zero results   : {tally['zero']}")
    print(f"  Searches that errored        : {n_err}")
    print(f"  Never attempted              : {n_missing}")
    print(f"  Reconciles to                : {sum(tally.values())} "
          f"({'OK' if sum(tally.values()) == len(unique_front5) else 'MISMATCH'})")
    print(f"  Items matched (Stocked)      : {n_matched}")
    print(f"  Items in Review Queue        : {n_review}")
    print("  Full per-front5 detail is in the 'Run Log' tab.")
    if n_err or n_missing:
        print("\n  RESULT: NOT CLEAN — "
              f"{n_err} error(s), {n_missing} never attempted. The matched "
              "set is INCOMPLETE; re-run the same command to retry the "
              "failed front5 values before trusting the output.")
    elif finished is False:
        print("\n  RESULT: PARTIAL — run was stopped early. Re-run the same "
              "command to resume.")
    else:
        print("\n  RESULT: CLEAN — all searches accounted for, no errors.")


if __name__ == "__main__":
    main()

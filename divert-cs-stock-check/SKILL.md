---
name: divert-cs-stock-check
description: >
  Check whether C&S Wholesale Grocers stocks the items on a research
  spreadsheet, by driving the C&S sourcing portal
  (https://divert.cssourcing.com) with the user's own logged-in browser
  session. Searches each distinct UPC front5 from the research file, scrapes
  the Product List results, matches them back to the research rows on
  back5 == CsUPC, and produces a matches-only copy of the workbook with a
  "No Buy" column. Use this skill WHENEVER the user has a research spreadsheet
  (Coffees_and_Teas_Brands_-_RESEARCH.xlsx or any file in the same shape) and
  wants to know which of those items C&S / divert carries. Trigger on phrases
  like "does C&S stock these", "check divert for this list", "run the C&S
  stock check", "which of these does divert carry", "check cssourcing", "run
  the divert stock check on this file". This is scoped to divert.cssourcing.com
  / C&S only — do not use it for Net Trade research (new-offer-research-agent),
  floor comparisons (source-floor-price-compare), or bulk uploads
  (trading-floor-bulk-upload-builder).
---

# Divert C&S Stock Check

Take a research spreadsheet and find out which of its items C&S Wholesale
Grocers carries, by searching the C&S sourcing portal at
`https://divert.cssourcing.com`. The deliverable is a **new copy** of the
research workbook containing only the matched rows, with all original columns
preserved plus a single new **"No Buy"** column.

This is a semi-automated Python + Playwright script, built on the same pattern
as `nettrade_batch_fill.py`: a **headed** browser, a **manual login step**
(the user logs in with their own C&S credentials), then an automated
search → scrape → match → write loop. The user is present at each run to log
in — the script never handles credentials.

## Hard rules — read these first

1. **NEVER guess a selector or a mapping.** The C&S portal DOM is not baked
   into this skill. Every site selector starts UNSET. If the live DOM doesn't
   match what the script expects, the script STOPS and tells you to inspect —
   it does not invent a plausible selector and press on. This mirrors the
   back5-collision lesson from the Net Trade build: a wrong result is worse
   than a missing one.

2. **The original upload is never touched.** Output is always a new file
   (`..._STOCKED.xlsx`). The research file is read-only input.

3. **Matches only.** Rows the site did not return are dropped from the output
   entirely. This is deliberate and confirmed with the user, not an oversight.

4. **The matching logic is confirmed — do not re-derive it.** front5, back5,
   and the CsUPC comparison are specified below exactly. Do not change them
   without asking the user.

5. **Validate before a full run.** Confirm front5/back5 matching against a
   handful of real site results (the `--stop-after-chunk 1` step) before
   trusting a full ~370-search run.

## Lessons learned (2026-08-24 retrospective) — read before adding a tab or column

Real usage exposed a gap between what this skill produces and what the user
actually needed to hand the account manager, plus a genuinely new business
rule that was never implemented as a filter. Captured here so the next
change starts from the real requirement instead of re-discovering it.

**What happened.** The internal workbook grew to 4-5 tabs (Stocked, Review
Queue, Stocked Vendor Lines, Lookfor, optional Lookfor Audit) plus a
separate `_CS_STOCK_SUMMARY.xlsx` file with its own column layout — each
added in response to a real, well-evidenced need in the moment. But the
document the user actually built and used with the account manager
(`..._FINAL.xlsx`) was simpler than all of that: the **same 23-column
schema as the original research file** (no custom columns), a subset of
the original rows kept, the original blank spacer rows between vendor
sections left in place, and a plain **bold vs. not-bold** marker on the
BRAND cell instead of a Status text column (bold = on the original offer
sheet **and** stocked; not bold = C&S stocks it but it wasn't on the offer).
The user assembled this by hand from the richer output — "used the doc you
put together to review the master doc, remove lines that CS didn't stock,
and bold lines that match the offer sheet and CS stock" — and described the
process as scope creep they had to recenter away from.

**New rule this surfaces.** "No Buy" was only ever displayed/aggregated by
this skill, never used as an exclusion filter. The user's actual final
document also drops every row where No Buy isn't specifically `"No"` — a
stocked-but-No-Buy-isn't-"No" item is not something C&S will actually sell,
so it doesn't belong in an account-manager-facing deliverable even though
it's legitimately "stocked." No existing tab applies that filter.

**Revision for next time.**
1. Before adding a new tab, column, or output file, confirm the exact final
   shape the user needs — ask what they'll do with it and who sees it,
   rather than inferring the shape from whatever problem was just surfaced.
   Each artifact added here was justified in isolation; the sum was more
   than the user needed to act on.
2. Treat **Review Queue**, **Stocked Vendor Lines**, and **Lookfor Audit**
   as internal diagnostics for validating a run, not deliverables to build
   toward by default — useful for trusting the matching logic, not for a
   customer-facing account manager.
3. The account-manager-ready shape that was actually proven out by hand:
   original research columns, filtered to Stocked with **No Buy == "No"**
   exactly (plus Lookfor "ask source" items reformatted into the same
   columns), with bold/not-bold in place of a Status column.
4. **Automated 2026-08-24 as the "Final Review" output**, once the user
   confirmed the two open questions with real detail:
   - "Bold" is the **whole row**, not just the BRAND cell — confirmed by
     checking `font.bold` across multiple columns of the same rows in the
     user's real final file.
   - The not-bold rows carry full original pricing/spec data (same as bold
     rows), which rules out them being reformatted Lookfor items — Lookfor
     items have no research-file pricing by definition. They are Review
     Queue rows (site UPC-column match without a CsUPC match) the user
     manually vetted and kept.
   - The user's own vocabulary maps directly onto this skill's two exact
     match tiers: **"case code" match = CsUPC = the primary Stocked match
     (bold, definite)**; **"item code" match = the site's UPC column = the
     Review Queue match (plain, still needs the user's own discrepancy
     check)** — confirming the two-tier design was already correct.
   - A new, previously unimplemented rule: exclude any row — Stocked or
     Review Queue — whose aggregated No Buy value isn't unambiguously
     `"No"` (a Yes/No discrepancy across DCs, or no value seen at all).
     There are too many of these to adjudicate automatically, and an
     inconsistent row isn't action-ready.
   - The Lookfor / "ask source" concept stays a **separate, second-layer
     document** (the existing Lookfor tab / CS_STOCK_SUMMARY "Ask Source"
     rows) — never merged into Final Review.

   See "Final Review" under Output below for the resulting format.

## Scope

This skill is **C&S / divert.cssourcing.com only**. There is intentionally no
generic multi-portal abstraction — the login flow, DOM, and search form are
specific to this one site. Keep it that way.

## Input file schema

`Coffees_and_Teas_Brands_-_RESEARCH.xlsx` and future files in the same shape.
Columns (order preserved in output):

> UPC, BRAND, DESCRIPTION, SIZE, UOS, PACK, SHELF, LIST PRICE, YOUR COST,
> ITEM WGT, ITEM HGT, ITEM LNG, ITM WIDTH, CASE HGT, CASE LNGT, CASE WIDTH,
> ITEM CUBE, BLOCK, TIER, PALLET, CATEGORY 1, CATEGORY 2, CATEGORY 3

Notes the script handles automatically:
- **UPC is stored as an Excel number**, so leading zeros are sometimes
  stripped (in the sample file: 2,144 rows at 12 digits, 629 at 11 digits,
  2 at 10 digits). The script zero-pads every UPC to 12 characters.
- **Fully blank spacer rows** (334 in the sample file) are skipped.
- The UPC column is located by header name (`UPC`), not by position.

## Confirmed matching logic (do not change without asking)

1. Zero-pad every UPC to a 12-character string, left-padded with `"0"`.
2. `front5` = characters at index `[1:6]` (0-indexed) of the padded UPC.
   Example: `850031180208` → front5 `"50031"`.
3. `back5` = characters at index `[6:11]` of the padded 12-digit UPC — the
   item code sitting between the manufacturer code (front5) and the trailing
   GS1 check digit at index 11. **Corrected 2026-08-24** — originally taken
   as the literal last 5 characters (index `[7:12]`), which swaps the check
   digit in for the item code's true last digit. Caught against real
   confirmed-buy ground truth: `0-72310-00041-4` decodes to front5 `72310`,
   item `00041`, check `4` — the live site's CsUPC for that exact product is
   `00041` (index `[6:11]`), not `00414` (the old, wrong `[7:12]` result).
   Verified exact against three independent confirmed BIGELOW matches.
4. Build the set of **distinct** front5 values across all UPC rows — dedupe,
   and search **once per unique front5**, not once per row.

### Matching a scraped result back to the research file
For each scraped result row, take the site's **CsUPC**, normalize it to a
5-character zero-padded string (`csupc5` — CsUPC may lose leading zeros as a
number), and compare it to the `back5` of the research rows that share that
front5. **Exact match = hit.** No fuzzy or partial matches — a wrong match is
worse than a miss.

> The `--stop-after-chunk 1` validation run prints the full derivation
> (pad12 → back5 vs. CsUPC raw → csupc5) for a sample of matches, so you can
> confirm whether CsUPC needs the zero-padding on the real site before
> trusting the full run.

## Site workflow

1. Launch a **headed** Chromium browser and open
   `https://divert.cssourcing.com/login`. **Wait for the user to log in
   manually.** By default the script prints a prompt and blocks on Enter
   (`press Enter once logged in`) — the login flow is unverified, so this is
   the safe default. If a post-login marker element is confirmed, set
   `SELECTORS["logged_in_marker"]` and the script will poll for it instead.
2. Reach **Product Search** (via a confirmed nav selector, or a manual pause).
   For each unique front5:
   - Enter the front5 in the **"Upc"** search box, leave **DC = "All DC"**
     (the default), click **Search**.
   - Scrape the **Product List** table: DC, DcName, ItemCode, UPC, CsUPC,
     Description, Pk/Sz, Type, QC Days, No Buy.
   - Zero rows → skip silently (no output, no log for that front5).
3. Add a **randomized 1–3 second delay** between searches (rate-limit
   courtesy).
4. Split the unique front5 values into **5 roughly-equal chunks**, processed
   sequentially. **Save progress to disk after each chunk** so the run is
   resumable if the browser session times out or the script crashes partway.
   On restart the script detects the progress file and resumes from the next
   unprocessed chunk.

## Dedup across DCs

A matched item can appear under multiple DCs in one search. Collapse these into
**one output row per matched research item** (ignore which DC carried it). If
**"No Buy"** differs across DCs for the same item, report all distinct values
seen, comma-separated (e.g. `"No, Yes"`); if uniform, just the one value
(e.g. `"No"`).

## Output

A new workbook, `<input basename>_STOCKED.xlsx` (e.g.
`Coffees_and_Teas_Brands_-_RESEARCH_STOCKED.xlsx`), with four tabs. The UPC
column is rendered as a readable dashed UPC-A (`0-72310-00041-4`) in every
tab that has one — display-only, matching always runs on the raw digits.

- **Stocked** — matches-only, primary CsUPC-exact hits, all original columns
  preserved plus one new **"No Buy"** column populated per the dedup rule
  above. This is the confirmed list; the matching rule here is never relaxed.
- **Review Queue** — confirmed 2026-08-24: a scraped row's own **UPC** column
  back5 can genuinely disagree with that same row's **CsUPC** (e.g. UPC
  `50003-79769` but CsUPC `79774`). A research row whose back5 matches the
  site UPC column but never got a primary CsUPC hit on any scraped row would
  otherwise be silently dropped — instead it lands here with the front5,
  DC(s), site UPC(s), CsUPC(s), description(s), Pk/Sz(s), and **No Buy(s)**
  seen, so it can be resolved in a quick, pre-narrowed pass instead of
  manually scanning the full research file. Still an exact match, just
  against a different real field — never fuzzy, and never merged into Stocked.
- **Stocked Vendor Lines** — every original research row (matched or not, in
  original file order) whose BRAND has at least one confirmed Stocked match —
  the "load the whole line" view, since a vendor C&S carries even one item
  from is worth considering in full for other clients. Based on the Stocked
  tab only, not the Review Queue (still unconfirmed). Each row marked
  **"Confirmed Stocked"** Yes/blank plus its **No Buy** value where applicable.
- **Lookfor** — requested 2026-08-24: for any front5 with at least one
  confirmed Stocked match (a proven vendor relationship), items C&S actually
  carries under that front5 with **no representation at all** in the research
  file — e.g. "ITO EN shows 2 matches, but C&S actually carries 7." Checked
  against both back5 and the site UPC-column back5, so nothing already in
  Stocked or Review Queue is double-counted. Deduped by (front5, CsUPC,
  description); columns: Front5, Brand, Description, Pk/Sz, Type, CsUPC,
  site ItemCode(s), site UPC(s), DC(s), DC Name(s), No Buy(s).
  **Relevance filter — brand-anchored, revised 2026-08-24.** A front5 is
  not always exclusive to one vendor: JOYBA is a Del Monte brand, so its
  front5 also carries Del Monte's canned goods. Word-overlap filtering
  failed here because the two lines share generic fruit/beverage vocabulary
  (`DM BBL APL FRT WTRMLN GEL` shares BBL and FRT with `JOYBA BBL … DRGN
  FRT`; `DM DICE MANGO LT SYRP` shares MANGO).

  The working signal is the **brand token every real C&S description leads
  with** — `BIGLOW RED RSPBRRY`, `BTLLI EXTRA VIRGIN OLIVE OIL`,
  `*JOYBA BBL RASP…`, `DM CUT GREEN BEANS`. A candidate is kept only if
  some token identifies it as one of the brands the research file lists
  under that front5. Since C&S abbreviates by dropping vowels at varying
  depth, matching uses **consonant skeletons** with a subsequence test
  (`consonant_skeleton` + `brand_token_matches`): BIGELOW/BIGLOW/BGLW all
  reduce to `BGLW`, BOTTICELLI accepts `BTLLI`, CLEANSE accepts `CLNCSE`.
  A shared first letter is required and 2-char skeletons use a stricter
  prefix test, so `DM`/`DELMONTE`/`DLMNT` can never match `JOYBA`.

  Nothing is silently dropped: kept rows carry a **"Matched On"** column,
  the run prints how many were excluded, and `--lookfor-audit` adds a
  **"Lookfor Audit (excluded)"** tab listing every exclusion with its
  reason. Word-overlap remains only as a fallback when the research file
  has no BRAND column at all.

  **Deduping, fixed 2026-08-24:** keyed strictly on `(front5, CsUPC)` —
  CsUPC is C&S's own item code within that front5, the real identity — not
  on description. The same physical item scraped from different DCs can
  carry cosmetic description drift (a stray leading/trailing `*`, a
  truncated ending), which fragmented into apparent duplicate rows when
  description was part of the key. Descriptions now aggregate into a set
  like every other cross-DC field, with the longest one shown (avoids a
  truncated `4P` when `4PK` was also seen).

  **UPC = CsUPC column, added 2026-08-24:** every Lookfor row (and the
  account-manager summary) carries an identifying `front5-CsUPC` pair —
  the same identifier C&S itself displays — since there's nothing to hand
  the source without one.

Unmatched research rows not appearing in Stocked or Review Queue are
genuinely dropped — neither CsUPC nor the site UPC column matched anything.
The original research file is left untouched.

### Account-manager summary — a SEPARATE, shareable file

Requested 2026-08-24: the account manager doesn't need the matching detail,
just what C&S can stock and buy. `<input basename>_CS_STOCK_SUMMARY.xlsx`
(`--share-out` to override) is written as its own file — never a tab in the
internal workbook — with one sheet, one row per item, sorted by Brand then
Description:

| Brand | Description | UPC = CsUPC | Pack/Size | Your Cost | List Price | % Spread | Status |

- Rows come from **Stocked** (`Status = "On Offer"`, full pricing from the
  research file) and **Lookfor** (`Status = "Ask Source"`, pricing left
  blank — there's none for an item that was never on the offer) only.
  Review Queue rows are excluded — still unconfirmed, not for external eyes.
- **UPC = CsUPC**, added 2026-08-24: the research file's own dashed UPC for
  "On Offer" rows; the site's own `front5-CsUPC` pair for "Ask Source" rows,
  since those have no research-file UPC by definition — needed so there's
  something concrete to hand the source when asking for an item.
- **Pack/Size** combines the research file's separate PACK/SIZE/UOS columns
  (e.g. `12/15.50 FO`), matching the site's own Pk/Sz display convention.
- **% Spread** = `(List Price − Your Cost) / List Price × 100`.

### Final Review — the proven account-manager-ready shape

Added 2026-08-24 after the user's real Coffees & Teas run showed the shape
they actually needed: `<input basename>_FINAL_REVIEW.xlsx` (`--final-out` to
override). No custom columns — the account manager is used to seeing the
source document, and extra analytical columns confuse rather than help.

- **Same columns as the research file**, in the same order. No added
  columns — bold/plain carries the confidence signal instead of a Status
  column.
- **Rows included:** Stocked (case-code / CsUPC match) rows and Review
  Queue (item-code / site UPC-column match) rows, **both filtered to a
  clean, unambiguous `No Buy == "No"`** (`_no_buy_is_clean_no` /
  `_review_no_buy_is_clean_no`) — any row with a Yes/No discrepancy across
  DCs, or no No Buy value at all, is dropped. Never includes Lookfor items;
  that stays a separate "ask the source" document.
- **Bold = case-code (CsUPC) match — definite.** Plain = item-code (site
  UPC column) match — the user still does a final manual pass over these
  for discrepancies before sending, same as their existing process. Bold
  applies to the whole row (confirmed against the user's real file, not
  just the BRAND cell).
- **Grouped by Brand, then Description**, with a blank row between brand
  groups — matching the visual sectioning of the source file, without
  needing to preserve the original file's own blank-row positions.

## What is UNVERIFIED — confirm live, do not guess

The C&S DOM has not been inspected. These must be confirmed on the real site
via `--inspect` before a full run, and the script hard-stops until they are:

- Exact selectors for: the **"Upc"** search input, the **DC** dropdown, the
  **Search** button, and the **Product List** results table.
- The login flow / what signals "logged in" (URL change? an element
  appearing?). Default is the manual Enter prompt; set `logged_in_marker`
  only if confirmed.
- Whether **CsUPC** in the scraped table ever needs its own zero-padding
  (confirm against a few real matched examples in the validation run before
  trusting the comparison — the script already zero-pads defensively and
  prints the raw vs. padded value so you can verify).

## Workflow (run it in this order)

### Step 0 — Confirm the file and dependencies
Make sure the research file is present and dependencies are installed:
```
pip install playwright openpyxl
playwright install chromium
```

### Step 1 — Discover the real DOM (`--inspect`)
```
python scripts/divert_stock_check.py --input RESEARCH.xlsx --inspect
```
A headed browser opens on the login page. **Log in manually.** Navigate to the
Product Search page (run a search that returns some results so the results
table is present), press Enter, and the script dumps every input, select,
button, and table on the page. Use that output to fill in the `SELECTORS`
block at the top of `scripts/divert_stock_check.py` with confirmed CSS
selectors. **Do not guess — copy from what `--inspect` shows.**

### Step 2 — Validate matching on the first chunk
```
python scripts/divert_stock_check.py --input RESEARCH.xlsx --stop-after-chunk 1
```
The pre-flight self-check verifies your selectors resolve on the live page
(and refuses to run if any required one is unset or broken). It then processes
only the first chunk, saves progress, writes a partial `..._STOCKED.xlsx`, and
prints a **match validation sample** showing the front5/back5 → CsUPC
derivation. Eyeball it: are the matches real? Does CsUPC need padding?

### Step 3 — Resume and finish
```
python scripts/divert_stock_check.py --input RESEARCH.xlsx
```
Re-running the same command detects the progress file and resumes from the
next unprocessed chunk — already-searched front5 values are not re-searched,
and no duplicate output rows are produced. When all chunks are done it writes
the final matches-only workbook.

### Rebuild output without searching
If searching is complete but you want to regenerate the workbook:
```
python scripts/divert_stock_check.py --input RESEARCH.xlsx --rebuild-output
```

## Success criteria

- Running end-to-end against the real site with the real research file
  produces the filtered `..._STOCKED.xlsx`.
- front5/back5 matching is validated against a handful of real site results
  (Step 2) before the full ~370-search run.
- A chunk can be interrupted and resumed without re-searching completed front5
  values or producing duplicate output rows.
- No guessed selectors or invented business logic anywhere. If the live DOM
  doesn't match what's assumed, the script stops and you ask the user.

## Script

### `scripts/divert_stock_check.py`
The full implementation: research-file reader, front5 chunking, the resumable
Playwright search loop, the Product List scraper (header-mapped, never
positionally guessed), the back5==CsUPC matcher with cross-DC dedup, and the
matches-only workbook writer. The `SELECTORS` block at the top is the only
part that must be filled in per the live DOM; everything else is confirmed
logic. Read the module docstring before the first run.

## Operational notes

- `~370` unique front5 searches is typical for a file this size. At a 1–3s
  courtesy delay plus page load, budget roughly 15–25 minutes of wall time,
  which is why chunked, resumable progress matters — a session timeout mid-run
  costs only the current chunk.
- The browser is **headed** and login is **manual** by design. Do not run
  `--headless` for a real run; the user must be able to log in.
- Progress lives in `<input basename>_progress.json`. Delete it to force a
  clean re-run from scratch. It is invalidated automatically if the input file
  or the front5 set changes.
- `openpyxl` reads and writes the `.xlsx`; `playwright` (Chromium) drives the
  site.

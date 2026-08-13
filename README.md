# Trendyol Marketplace QC Tool

Checks a Trustana product export (CSV) against Trendyol's Seller Information
Center listing rules and produces an Excel QC report. Built from
`Trendyol_MP_QC_Rules.md` (Trendyol's public Platform Rules, upload field
formats, rejection reasons, and image guide).

## Just want to use it?

If someone shared a link with you (`https://something.streamlit.app`), you
don't need anything below this section. Open the link, upload your Trustana
CSV export, click **Run QC**, and download the Excel report. No install, no
command line. Everything else in this README is for running the tool
locally instead of through that web page, or for whoever is maintaining it.

## What you get

Running the tool produces one Excel file with five sheets:

- **Summary** — totals, a pass/fail count per rule, and any file-level data
  gaps (e.g. "Price/Stock columns are empty in this export").
- **Issues** — a flat list, one row per problem found: SKU, rule, severity
  (HARD = rejection-risk, SOFT = quality flag), and the specific message.
- **Products** — one row per product with an overall status
  (`REJECTION RISK` / `NEEDS REVIEW` / `OK`) and all its issues listed.
- **Content Checks** — the actual evidence behind the category-match and
  title/description-match checks (which keywords matched, which didn't) so
  you can see *why* something passed or failed, not just that it did.
- **Rules Reference** — every rule the tool checks, with severity and which
  section of the source rules doc it comes from.

*(The web app version shows the same information inline on the page, plus a downloadable copy of this same Excel file.)*

## Running it locally instead of via the web link

The rest of this README covers running the tool yourself from a terminal —
useful if you want to script it, run it offline, or the web app isn't
available for some reason. If the shared link works for you, you can skip
straight to "Known limitations" below.

## One-time setup

You need Python 3.9 or newer installed. Check with:

```
python3 --version      # Mac/Linux
python --version       # Windows
```

If that's missing, install it from [python.org](https://www.python.org/downloads/)
(on Windows, tick "Add Python to PATH" during install).

Nothing else to install by hand — the wrapper scripts below set up
everything else (a private Python environment inside this folder,
`openpyxl`/`requests`/`Pillow`) automatically the first time you run them.

Keep all the files in this folder together — the wrapper scripts assume
`trendyol_qc_tool.py` and `requirements.txt` are next to them.

## Running it

**Mac/Linux:**
```
./run_qc.sh /path/to/Products_export.csv
```
(First time only: `chmod +x run_qc.sh` if you get a "permission denied".)

**Windows:** open Command Prompt in this folder and run:
```
run_qc.bat C:\path\to\Products_export.csv
```
Or just drag-and-drop the CSV file onto `run_qc.bat` in File Explorer.

Either way, the report is saved right next to your input file, named
`<yourfile>_QC_Report.xlsx`.

### Optional: real image verification

By default the tool checks image URLs and counts (1–8, valid link format)
but doesn't download the actual image files. To also download every image
and verify real format/file size/resolution against Trendyol's spec
(JPEG/PNG/WEBP, 1KB–10MB, 860×574–2000×2000px), add `--verify-images`:

```
./run_qc.sh Products_export.csv --verify-images
```

This needs your machine to have normal internet access to wherever the
images are hosted (fine on a normal laptop; it will NOT work from a
locked-down/sandboxed environment with restricted outbound network — the
tool detects that automatically and reports "not evaluated" rather than a
false pass, so you'll know if it happened). It's also slower on large
exports — there's a `--max-image-checks N` option to cap how many images it
downloads if you want to keep it quick.

## Running it without the wrapper scripts (any OS)

If you'd rather call Python directly (e.g. you already have a venv you use
for other things):

```
pip install -r requirements.txt
python3 trendyol_qc_tool.py Products_export.csv QC_Report.xlsx [--verify-images]
```

## When Trustana's export format changes

If Trustana renames a column (e.g. "Title [EN]/Marketing" becomes something
else), the tool will silently treat that field as empty rather than error —
so check the Summary sheet's "Not evaluated / N/A" counts if numbers look
off. Fix it by opening `trendyol_qc_tool.py` and editing the `COLUMN_MAP`
dictionary near the top of the file to point at the new column name. No
other code changes needed for a rename.

## When Trendyol's rules change

Every rule lives in its own `check_...()` function in `trendyol_qc_tool.py`,
and is registered once in the `RULE_CATALOGUE` list near the top (which also
drives the "Rules Reference" sheet, so update the description there too).
Search for the rule's name to find it. The banned-word lists
(`HEALTH_CLAIM_TERMS`, `OFF_PLATFORM_KEYWORDS`, `PRICE_SHIPPING_PROMO_TERMS`)
and the category-synonym list (`_CATEGORY_SYNONYM_GROUPS`) are the parts most
likely to need extending over time as real-world false positives/negatives
turn up — they're plain Python lists near the top of the file.

If the tool is also deployed as a web app (`app.py`), it imports its checks
directly from `trendyol_qc_tool.py` rather than duplicating anything — a fix
made here applies to both the command line and the web app once it's
redeployed. See `DEPLOY.md` for how the web app is deployed/updated.

## Known limitations (by design, not bugs)

- **Category-correctness and title/description-relevance checks are
  keyword-overlap heuristics**, not true semantic understanding. They can
  miss a wrong category when a generic word happens to be shared (e.g. an
  "Ice Maker" filed under "Ice Trays" — both mention "ice", so it won't be
  flagged even though they're different products). Treat SOFT flags as
  "worth a human look," not a verdict.
- **Google Category is shown for reference only** (Products sheet), not
  used in any automated pass/fail — testing showed it's a much broader,
  more generic taxonomy than Trendyol's, so comparing them directly produced
  mostly false positives.
- **Image content (does the photo match the product?) is not checked.**
  `--verify-images` verifies file format/size/resolution, not what's in the
  picture. Checking that the *content* of the image matches the title would
  need a vision model, which isn't part of this tool.
- **Price/Stock/Model Code checks depend on your export actually having
  those columns populated.** If a whole column is empty across the entire
  file, the tool treats that as a file-level data gap (noted once in
  Summary) rather than flagging every single product — check the Summary
  notes to see if that happened on your export.

## Questions / issues

This tool was built inside a Claude Cowork session for Sathish (Graas). If
something looks wrong or a rule needs adjusting, the fastest path is
usually: open `trendyol_qc_tool.py`, find the relevant `check_...()`
function (they're short and commented), and adjust — or bring it back to
Claude with the specific CSV/row that's behaving unexpectedly.

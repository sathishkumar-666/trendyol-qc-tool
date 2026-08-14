# Trustana AI Content Verifier for MP

Checks a Trustana product export (CSV) against a marketplace's listing
rules, produces an Excel QC report, and — for every product that passes —
can fill that marketplace's own upload template so it's ready to submit.
Currently supports **Trendyol** and **noon**, with more marketplaces
addable without touching the shared checking logic (see "When a
marketplace's rules change" below). Built from `Trendyol_MP_QC_Rules.md`
(Trendyol's public Platform Rules, upload field formats, rejection reasons,
and image guide) and `Noon_MP_QC_Rules.md` (noon's Seller Help Center
listing/image/QC policy pages), plus noon's own Classification Directory
(`noon_template.xlsx`) for real category validation.

**Important: Trendyol and noon exports from Trustana have different
column layouts.** Configure your Trustana export template to match the
marketplace you're about to check — e.g. noon's export needs "Product
Title EN/Marketing", "Long Description EN/Marketing", and per-bullet
Feature Bullet columns, which Trendyol's export doesn't have (see "Trustana
export shape per marketplace" below). Pick the matching marketplace in the
tool for whichever shape of file you're uploading.

## Trustana export shape per marketplace

Trendyol and noon exports from the same Trustana account use genuinely
different column names for the same kind of content — this isn't a
tool-side quirk, it reflects how each marketplace's Trustana template is
configured. Make sure the export you upload actually matches the
marketplace you pick in the tool:

| Field | Trendyol column | noon column |
|---|---|---|
| Title | `Title [EN]/Marketing` | `Product Title EN/Marketing` (+ `Product Title AR/Marketing`) |
| Description | `Description [EN]/Marketing` | `Long Description EN/Marketing` (+ `Long Description AR/Marketing`) |
| Feature bullets | *(none)* | `Feature Bullet 1-5 EN/Marketing` and `Feature Bullet 1-5 AR/Marketing` (10 columns) |
| What's in the box | *(none)* | `What's In The Box EN/Marketing` (+ AR) |
| Warranty type | `Warranty Type [EN]/Marketing` | *(no such column in noon's export)* |
| Category | `Product Category` (checked against a keyword heuristic) | `Product Category` (checked against noon's own Family/Product Type/Subtype directory — see below) |

Columns shared by both (SKU, Product Name, Brand, Product Model, Images,
barcode, Price, Stock, Google Category) are mapped the same way for either
marketplace. If you pick the wrong marketplace for your file's shape, the
tool won't error — it'll just treat the mismatched fields as empty, so
double-check the Summary sheet's data-gap notes if the numbers look off.

### noon's category check is directory-backed, not a guess

noon ships its own **Classification Directory** — a ~6,500-row list of every
valid Family / Product Type / Product Subtype combination — inside its NIS
bulk-upload template. This tool ships a copy of that template
(`noon_template.xlsx`) and uses its Classification Directory sheet two ways:

1. **Category validity (HARD)** — noon's `Product Category` column is
   expected to already contain the exact Family/Product Type/Subtype triple
   the seller selected in noon's own upload flow (e.g. `Kitchen & Dining /
   Drinkware / Water Bottle`). The tool checks that this triple actually
   exists in the directory, and flags it if the category has the wrong
   number of levels or doesn't match any real entry.
2. **Category-relevance keywords (SOFT)** — instead of the generic
   2-level path heuristic used for Trendyol, noon's content-relevance check
   pulls its comparison keywords from that category's "Products Usually
   Included" column in the directory, which is a much richer, marketplace-
   specific synonym list than guessing from the category path alone.

If `noon_template.xlsx` is missing or its Classification Directory sheet
can't be read, the tool doesn't fail — it falls back to "unverified" for
category validity and to the old path-heuristic for relevance, and notes
this once on the Summary sheet so you know the check ran in a degraded mode.

noon's export is also checked for **stray HTML markup** (`<div>`, `<p>`,
`<h1>`, `<ul>`, etc. — noon requires plain text in title/description/bullet
fields). Trustana's own rich-text export wraps *every* Arabic field in a
mechanical `<div>...</div><div>...</div>` alignment artifact that isn't
seller-authored formatting, so the tool strips that specific wrapper before
checking — only markup embedded inside the actual text still gets flagged.

## Just want to use it?

If someone shared a link with you (`https://something.streamlit.app`), you
don't need anything below this section. Open the link, pick a marketplace,
upload your Trustana CSV export, click **Run QC**, and download the Excel
report. No install, no command line. Everything else in this README is for
running the tool locally instead of through that web page, or for whoever
is maintaining it.

## What you get

Running the tool produces one Excel file with five sheets, and — once your
products are QC-clean — can also fill that marketplace's own upload
template so you can submit it directly (see "Stage 2" below).

- **Summary** — totals, a pass/fail count per rule, and any file-level data
  gaps (e.g. "Price/Stock columns are empty in this export").
- **Issues** — a flat list, one row per problem found: SKU, rule, severity
  (HARD = rejection-risk, SOFT = quality flag), and the specific message.
- **Products** — one row per product with an overall status
  (`REJECTION RISK` / `NEEDS REVIEW` / `OK`), all its issues listed, and an
  "Image URLs" column listing every image URL for that product (one per
  line in the cell) alongside the "# Images" count, so you can open/check
  the actual images without leaving the report.
- **Content Checks** — the actual evidence behind the category-match and
  title/description-match checks (which keywords matched, which didn't) so
  you can see *why* something passed or failed, not just that it did.
- **Rules Reference** — every rule the tool checks for the marketplace you
  ran it against, with severity and which section of the source rules doc
  it comes from.

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
`qc_engine.py` and `requirements.txt` are next to them.

## Running it

**Mac/Linux:**
```
./run_qc.sh /path/to/Products_export.csv trendyol
./run_qc.sh /path/to/Products_export.csv noon
```
(First time only: `chmod +x run_qc.sh` if you get a "permission denied".
Marketplace defaults to `trendyol` if you leave it off.)

**Windows:** open Command Prompt in this folder and run:
```
run_qc.bat C:\path\to\Products_export.csv trendyol
run_qc.bat C:\path\to\Products_export.csv noon
```
Or just drag-and-drop the CSV file onto `run_qc.bat` in File Explorer (this
runs the Trendyol ruleset by default, since drag-and-drop can't pass the
marketplace argument).

Either way, the report is saved right next to your input file, named
`<yourfile>_<marketplace>_QC_Report.xlsx`.

### Optional: real image verification

By default the tool checks image URLs and counts (valid link format, and
count against the marketplace's limits) but doesn't download the actual
image files. To also download every image and verify real
format/file size/resolution against the selected marketplace's spec (e.g.
Trendyol: JPEG/PNG/WEBP, 1KB–10MB, 860×574–2000×2000px; noon: JPEG/JPG
only, min width 660px, ≥72 PPI, ≤10MB), add `--verify-images`:

```
./run_qc.sh Products_export.csv trendyol --verify-images
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
python3 qc_engine.py Products_export.csv QC_Report.xlsx --marketplace trendyol [--verify-images]
python3 qc_engine.py Products_export.csv QC_Report.xlsx --marketplace noon [--verify-images]
```

## Stage 2: filling the marketplace upload template

Once you have a QC report, the tool can take every product that passed and
write it straight into that marketplace's own upload template — so instead
of hand-copying "clean" rows into noon's or Trendyol's bulk-upload sheet,
you get a ready-to-submit file back.

**You provide the template each time — the tool doesn't reuse a fixed
copy.** This is deliberate: Trendyol's downloadable template only contains
sheets for the categories you selected when you requested it, so a template
downloaded for a different category mix simply won't match your catalog.
noon's template is more stable but can still change. So the flow is:

1. Run QC as usual (web app: **Run QC**; CLI: no extra flags needed for this
   part).
2. Click **Generate export summary** (web app) — this shows a QC summary
   for manual review: how many products are eligible, and for Trendyol,
   which category sheet each group will land in. Nothing is written to a
   file yet at this point.
3. Go download the **current** upload template from noon's or Trendyol's
   seller center for whichever categories you're working with.
4. Upload that file where the app asks for it, then click **Fill template
   & download** — only now does the tool actually open your uploaded
   template and fill it. The web app's file-uploader step is required; the
   button stays disabled until you've uploaded a template.

**Which products count as "clean" is configurable.** By default, stage 2
only exports products with **zero issues of any kind** (HARD or SOFT) — the
strictest reading of "fully QC-verified." Tick "Also include products with
quality (SOFT) flags" (web app) or add `--allow-soft-issues` (CLI) to relax
this to "zero HARD (rejection-risk) issues" — SOFT quality flags like a
brand name appearing in the title won't block export. **This matters in
practice**: on a real 151-product Trendyol export tested with this tool,
the strict setting qualified **0 of 151** products (nearly every title
included the brand name, which trips a SOFT flag), while the relaxed
setting qualified **118**. If your export comes back with an empty stage-2
file, try the relaxed setting before assuming something's broken.

**noon** — fills the `template_data` sheet of whichever template you
upload, starting at its data row, using the machine-readable header row
already in that sheet: family/product_type/subtype (from the product's
category triple), SKU, brand, title/description (EN+AR), what's-in-the-box
(EN+AR), image URLs, feature bullets (EN+AR). Fields the tool can't
populate from your export (parent/variation keys, item condition,
per-country VAT rate) are left blank on purpose — they're listed on a
generated "Export Notes" sheet in the output so you know what still needs a
manual pass before uploading.

```
./run_qc.sh Products_export.csv noon --export-template noon_upload.xlsx --template-file path/to/current_noon_template.xlsx
./run_qc.sh Products_export.csv noon --export-template noon_upload.xlsx --template-file path/to/current_noon_template.xlsx --allow-soft-issues
```

**Trendyol** — fills the matching category sheet inside whichever template
you upload (a real per-category template — sheet names look like
`Pans(911)`, `Lunch Boxes(5638)`, `Bowl(2412)`, `Water Bottle &
Flask(814)`, `Backpack(448)`, `Chalk & Board Pen(1525)`, `School
Bags(974)`), filling the 20 shared required columns (Barcode, ModelCode,
Title, Description, Price, Stock, VatRate, Image1-8, etc.) plus that
sheet's category-specific attribute columns, cross-checking attribute
values against the template's own `Attribute_Value_List_*` dropdown sheets
and the field-level guidance in its `Help` sheet (e.g. barcode must be 8 or
13 numeric digits, brand must match the product's own brand). Required
fields your export doesn't have data for (Price/Stock/VatRate/ModelCode)
are left blank and called out on the "Export Notes" sheet rather than
guessed. Products whose category doesn't match any sheet actually present
in the template you uploaded go into a "Leftover - Needs Template" sheet
instead of being silently dropped — this is exactly what happens if you
upload a template that wasn't requested for that category, so check that
sheet if it's bigger than expected. Any attribute value that doesn't match
the template's own dropdown list is flagged on an "Attribute Value
Warnings" sheet.

```
./run_qc.sh Products_export.csv trendyol --export-template trendyol_upload.xlsx --template-file path/to/current_trendyol_template.xlsx --allow-soft-issues
```

If you run the CLI without `--template-file`, it falls back to the copy of
`noon_template.xlsx` / `trendyol_template.xlsx` bundled in this folder and
prints a warning that it did so — useful for a quick local test, but for a
real export always pass `--template-file` pointing at a template you just
downloaded. These bundled copies are also still used for noon's
Classification Directory category check (see above) — that's a stable
taxonomy the marketplace ships separately from the per-category upload
templates, so it doesn't need refreshing on every run the way the upload
template does.

## When Trustana's export format changes

If Trustana renames a column (e.g. "Title [EN]/Marketing" becomes something
else), the tool will silently treat that field as empty rather than error —
so check the Summary sheet's "Not evaluated / N/A" counts if numbers look
off. Fix it by opening `qc_engine.py` and editing the relevant column-map
dictionary near the top of the file to point at the new column name:
`BASE_COLUMN_MAP` for fields shared by every marketplace, or
`TRENDYOL_COLUMN_MAP` / `NOON_COLUMN_MAP` for fields that differ per
marketplace (title, description, feature bullets, warranty type — see
"Trustana export shape per marketplace" above for why these differ). No
other code changes needed for a rename.

## When a marketplace's rules change

Every rule lives in its own `check_...()` function in `qc_engine.py`, and
each marketplace's thresholds live in its `MarketplaceConfig` (`TRENDYOL_CONFIG`,
`NOON_CONFIG`) near the top of the file — e.g. title/description length
limits, barcode format, image specs, whether the marketplace has feature
bullets or a warranty-type field, which Trustana column names to read
(`column_map`), whether it has a Classification Directory for real category
validation (`has_classification_directory`), and whether it disallows HTML
markup in text fields (`disallows_html_tags`). `build_rule_catalogue(cfg)`
generates the "Rules Reference" sheet automatically from those thresholds,
so most rule changes are just editing a number or regex on the relevant
config, not rewriting a check function.

The banned-word lists (`HEALTH_CLAIM_TERMS`, `OFF_PLATFORM_KEYWORDS`,
`PRICE_SHIPPING_PROMO_TERMS`, `NOON_PROHIBITED_PRODUCT_TERMS`,
`COMPETITOR_MENTION_TERMS`) and the category-synonym list
(`_CATEGORY_SYNONYM_GROUPS`) are the parts most likely to need extending
over time as real-world false positives/negatives turn up — they're plain
Python lists near the top of the file.

### Adding a new marketplace

1. Add a new `MarketplaceConfig(...)` instance (copy `NOON_CONFIG` or
   `TRENDYOL_CONFIG` as a starting point) with that marketplace's actual
   thresholds.
2. Register it in the `MARKETPLACES` dict.
3. Drop a `<Marketplace>_MP_QC_Rules.md` reference doc next to the others
   documenting where each threshold came from.
4. It automatically appears in the web app's marketplace dropdown and the
   CLI's `--marketplace` choices — no other wiring needed unless the new
   marketplace needs a genuinely new check (e.g. a field none of the
   existing marketplaces have), in which case add a new
   `check_...()` function following the pattern of `check_feature_bullets`
   / `check_warranty_type`.

If the tool is also deployed as a web app (`app.py`), it imports its checks
directly from `qc_engine.py` rather than duplicating anything — a fix made
here applies to both the command line and the web app once it's
redeployed. See `DEPLOY.md` for how the web app is deployed/updated.

## Known limitations (by design, not bugs)

- **Category *validity* for noon is directory-backed** (checked against
  noon's real Classification Directory), **but category/title/description
  *relevance* checks for both marketplaces are still keyword-overlap
  heuristics**, not true semantic understanding — for noon the keywords
  come from the directory's "Products Usually Included" column, for
  Trendyol they come from the category path itself. Either way they can
  miss a wrong-but-plausible match (e.g. a "Tart Pan" filed under noon's
  "Pie Dish" subtype passed the validity check in testing since it's a
  valid subtype, but a human would flag it as the wrong one). Treat SOFT
  flags as "worth a human look," not a verdict, and treat noon's HARD
  category-validity check as "is this a real category," not "is this the
  *right* category for this specific product."
- **Duplicate-barcode detection reads the CSV as plain text, with no
  numeric parsing** — it will never itself introduce a rounding/precision
  artifact into a barcode value. If your export was ever opened and
  re-saved in Excel or Google Sheets before reaching this tool, though,
  *that* step can silently corrupt a long numeric-looking barcode column
  (spreadsheet software stores numbers as floating point, and a barcode
  with more digits than that can represent exactly gets rounded, often
  collapsing several different real barcodes into one that ends in a long
  run of zeros). When the tool sees a duplicated barcode matching that
  "long number, many trailing zeros" shape, it adds a note to the
  Issues-sheet message flagging this possibility — but it still reports the
  duplicate either way, since some sellers do genuinely reuse an all-zero
  placeholder barcode on purpose, which is just as real an issue. If you
  see that note, the fix is to re-export the CSV directly from Trustana
  without any Excel/Sheets round-trip in between and re-run QC.
- **Google Category is shown for reference only** (Products sheet), not
  used in any automated pass/fail — testing showed it's a much broader,
  more generic taxonomy than either marketplace's own categories, so
  comparing them directly produced mostly false positives.
- **Image content (does the photo match the product?) is not checked.**
  `--verify-images` verifies file format/size/resolution, not what's in the
  picture. Checking that the *content* of the image matches the title would
  need a vision model, which isn't part of this tool.
- **Price/Stock/Model Code/Feature Bullets/Warranty Type checks depend on
  your export actually having those columns populated.** If a whole column
  is empty across the entire file, the tool treats that as a file-level
  data gap (noted once in Summary) rather than flagging every single
  product — check the Summary notes to see if that happened on your
  export. In the tested Trustana export, feature bullets and warranty type
  aren't present as distinct columns, so noon's `feature_bullet_length`
  check is informational rather than a per-product fail, and
  `warranty_type_check` only evaluates the ~11 rows that do have a value.
- **noon's per-product language requirement (English mandatory for
  UAE/KSA, either for Egypt) can't be enforced per-product** — the tested
  export has no country/market column to know which market a given SKU
  targets. See `Noon_MP_QC_Rules.md` §9 for the full list of noon-specific
  gaps.

## Questions / issues

This tool was built inside a Claude Cowork session for Sathish (Graas). If
something looks wrong or a rule needs adjusting, the fastest path is
usually: open `qc_engine.py`, find the relevant `check_...()` function or
`MarketplaceConfig` (they're short and commented), and adjust — or bring it
back to Claude with the specific CSV/row/marketplace that's behaving
unexpectedly.

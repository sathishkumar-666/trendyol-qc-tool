# Amazon.sa Marketplace — Listing Content & QC Rules Reference

Compiled from Amazon's publicly documented product-detail-page / style-guide
rules (title, bullet points, description, product identifiers, image
requirements) and from the Seller Central "Download Product Spreadsheet"
flow on `sellercentral.amazon.sa` (Catalogue → Add Products via Upload).

**Confidence note.** Amazon's help pages require a Seller Central login, so
the numbers below come from Amazon's widely published style-guide values and
third-party summaries of them, not from a login-gated Amazon.sa page. Limits
vary by category (some categories use shorter title limits, e.g. 80–150
characters, and some bullets/description limits differ). **Treat the limits as
defaults and confirm against the category's own style guide / the Data
Definitions sheet of the downloaded spreadsheet.** Organized like the other
rules docs (rule → condition → severity).

---

## 1. Mandatory fields & format limits

| Field | Rule | QC severity |
|---|---|---|
| Title | Required. Up to **200 characters** including spaces (default; some categories lower). Recommended structure: Brand + product type + key attribute(s) (material, size, colour, quantity). The **brand should be included** (opposite of Trendyol/noon). | HARD: missing / too long (min 5 chars as a sanity floor). SOFT: brand absent, ALL CAPS, emoji, repeated words, just the category name. |
| Title characters | Must not contain `! $ ? _ { } ^` or the broken-bar / negation signs (unless part of the brand). | HARD (`title_forbidden_chars`). |
| Title word repetition | No word more than **2 times** (articles/prepositions/conjunctions excluded). | SOFT (in `title_quality`). |
| Description | Up to **2,000 characters**. Plain text only; **no HTML or JavaScript** except a `<br>` line break. No contact details, alternate ordering info, promotional content, prices, shipping/seller info. | HARD: over length, HTML (other than `<br>`), banned/off-platform/promo language. |
| Feature bullets | Up to **5 bullets**, each ≤ **500 characters** (≈255 recommended). Start with a capital, no end punctuation, no promotional/price/shipping text. | HARD: bullet > 500 chars. SOFT: fewer than 3 bullets (when bullets exist anywhere in the export). |
| Product ID (barcode) | A **GTIN: EAN-8 (8), UPC (12), EAN-13 (13) or GTIN-14 (14) digits**, numeric, with a valid GS1 check digit. Required unless a **GTIN exemption** is approved for the brand/category in Seller Central (this tool can't see exemptions). One product per detail page. | HARD (`barcode_format`): missing, non-numeric, wrong length, bad check digit. |
| Images | Main image + up to 8 additional (**max 9**). JPEG preferred; PNG, GIF, TIFF also accepted. Main image: **pure white background (RGB 255)**, product fills ≥ 85 % of the frame, no text/logos/watermarks/props. **≥ 1000 px on the longest side** enables zoom (500 px is the floor); recommended 2000 × 2000; ≤ 10 MB; sRGB. | HARD: no image, > 9 images, non-https / malformed URL / unrecognised extension. With `--verify-images`: format, size and resolution (500 px floor on each side). White-background / fill-% checks are **not** automated. |
| Category / product type | Every upload spreadsheet is for one **product type** (e.g. *Sauté & Frying Pan*). The Trustana `Product Category` is only used to decide which products go into which downloaded spreadsheet. | HARD: category missing. SOFT: shallow category path (heuristic). |
| Brand | Required in the template. Some brands are gated / require approval or Brand Registry. | Not validated (no brand list available). |
| Price / quantity | Offer data; filled from the export's `Price` / `Stock` columns when present. | HARD only if the whole export carries prices and one is missing. |

### General prohibitions (apply to title, bullets, description)
No offensive/obscene material, no contact details or alternative ordering
information, no promotional/time-limited claims ("sale", "free shipping",
"best seller"), no competitor-marketplace mentions, no unsubstantiated medical
claims, nothing violating Saudi/UAE/Egypt law for the destination store. These
use the same banned-term groups as the other marketplaces (health claims,
off-platform contact/link references, price/shipping/promo language,
competitor marketplaces).

---

## 2. Upload spreadsheet (stage 2)

Seller Central → **Catalogue → Add Products via Upload → Download Product
Spreadsheet**:

1. Spreadsheet language: *English (Amazon.sa)*.
2. Search by keyword/title or browse by category and select the **product
   type(s)** (max 20 per spreadsheet).
3. Select the store(s) — *Amazon.sa*.
4. **Generate Spreadsheet**, fill, then upload (Excel or tab-delimited text).

### Verified against a real download (`SAUTE_FRY_PAN.xlsm`, product type SAUTE_FRY_PAN)
The workbook is macro-enabled (`.xlsm`) with sheets *Changes to the template*,
*Instructions*, *Images*, **Data Definitions**, **Template**, *Browse data*,
*Valid Values*, plus hidden *Conditions List*, *Dropdown Lists*,
*AttributePTDMAP*. In **Template**:

| Row | Content |
|---|---|
| 1 | `settings=…` (URL-encoded; includes `labelRow=4&attributeRow=5&dataRow=7`, `ptds=<base64 product type>`, marketplace id `A17E79C6D8DWNP` = Amazon.sa) |
| 2 | Instruction text ("Use ENGLISH… DO NOT modify or delete the coloured header rows") |
| 3 | Column groups (Listing Identity, Variations, Product Identity, Images, Product Details, Offer, Offer (SA), Shipping, Safety & Compliance) |
| 4 | Human labels |
| 5 | Machine attribute names, e.g. `contribution_sku#1.value`, `product_type#1.value`, `item_name[marketplace_id=…][language_tag=en_AE]#1.value`, `brand[…]`, `amzn1.volt.ca.product_id_type`, `amzn1.volt.ca.product_id_value`, `main_product_image_locator[…]#1.media_location`, `other_product_image_locator_1…8`, `product_description[…]`, `bullet_point[…]#1…5`, `fulfillment_availability#1.quantity`, `purchasable_offer[…][audience=ALL]#1.our_price#1.schedule#1.value_with_tax` |
| 6 | Example row (ignored — data starts at `dataRow`) |
| 7+ | Data |

Whether each column is **Required / Conditionally Required / Recommended /
Optional** is listed in the *Data Definitions* sheet (Field Name + "Required?").
Valid values (e.g. Product Id Type = EAN / GTIN / UPC / ASIN / GTIN Exempt,
Item Condition = New, Country of Origin list) are in *Valid Values*; the
Template column validations are dropdowns driven by hidden sheets.

The exporter reads `labelRow/attributeRow/dataRow` from the settings cell (falling
back to pattern detection for other layouts), finds columns **by attribute
name**, takes the product type from `ptds`, and keeps every other part of the
workbook intact (502 data validations, 251 conditional formats, defined names,
`.xlsm` content type).

**Filled automatically:** SKU, Product Type (e.g. `SAUTE_FRY_PAN`), Item Name,
Brand Name, Product Description, Bullet Points 1–5, Product Id Type (UPC for 12
digits, EAN for 8/13, GTIN for 14) and Product Id, Main + 8 Other Image URLs,
Your Price, Quantity (SA) when the export has Price/Stock.

**Left blank, never guessed:** Listing Action (template default is *Create or
Replace*), Browse Nodes, Item Condition, Country of Origin, Fulfilment Channel,
Handling Time, dimensions/weights, Color/Size/Material and every other
product-type attribute, and all compliance fields. The companion
`…_notes.xlsx` lists the columns the template marks *Required* (and
*Conditionally Required*) that are still empty, plus the "Excluded From Export"
sheet. The upload workbook itself gets no extra sheets, so Amazon's validator
sees only what it generated.

## 3. Known gaps
- The template layout was verified against a real Amazon.sa download
  (SAUTE_FRY_PAN). No Amazon-shaped Trustana export has been seen yet: Trustana
  header names for Amazon are matched from a list of likely candidates
  (`AMAZON_COLUMN_MAP`) — verify on the first real export.
- openpyxl may drop the sample screenshots on the *Images* instruction sheet if
  Pillow is not installed; data, dropdowns and formatting are unaffected.
- Category → product type is **not** guessed; you assign export categories to
  each uploaded spreadsheet.
- Valid-value lists (Valid Values sheet) are not enforced on filled cells.
- Brand gating, restricted-category approvals and GTIN exemptions are not
  visible to this tool.

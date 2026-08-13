# Trendyol MP — Listing Content & QC Rules Reference

Compiled from Trendyol's public Seller Information Center (academy.trendyol.com/seller-information-center), specifically: **Platform Rules**, **Before Product Upload: Fields & Formats**, **Product Rejection Details**, **Product Image Upload Guide**, **Product Listing General**, and **Product Information**. These pages are public; the video-based "Trendyol Academy" course library covering the same ground requires seller login and could not be accessed in this session — the written Seller Information Center pages below turned out to contain the actual field-level rules and are the primary source.

This is organized to feed directly into a QC/gatekeeper tool (rule → condition → severity → source), mirroring the structure already used in the Trustana Gatekeeper tool.

---

## 1. Mandatory fields & format limits (product upload)

These are hard validation rules — violating them blocks approval outright, not just a quality ding.

| Field | Rule | Notes |
|---|---|---|
| Barcode / EAN | Required. 2–40 characters. No special characters (`? / & % + ^ ' * _`). Immutable after approval. | If the barcode already exists in Trendyol's catalog, must use "add from catalog" instead of creating new. |
| Model code | Required. 1–40 characters. Must not contain a URL or email address. Immutable after approval (ticket needed to change). | Same model code groups color/size variants onto one product page — must match across variants. |
| Brand | Required. Must be a brand already active on Trendyol (or a Brand Request submitted with a product photo showing the brand name legibly). Immutable after approval. | If no brand, leave blank (Excel/integration) or select "My product has no brand" (single-product UI). |
| Category | Required. Must match Trendyol's category tree. Immutable after approval (ticket needed to change). | Category tree downloadable via Bulk Actions. Maliciously wrong category (e.g. to dodge commission or restrictions) is itself a violation (12 points). |
| Title | Required. **3–200 characters, excluding brand name.** | See banned-content rules below (§3) and title-quality guidance (§5). |
| Description | Required. **Plain text ≤4000 chars; HTML ≤30000 chars.** HTML must be well-formed. | Must match the field type actually used (plain vs HTML) on the seller's own system. |
| Images | Required, **1–8 images**. | See §2 for full technical spec. |
| Product variants | Required per category (where applicable). | Some categories have no variant axis; contact support if a needed variant attribute is missing. |
| Product attributes | Required/optional per category. | Attribute set can change after a catalog update — a previously-valid value can become invalid and needs re-selection. |
| Original price | Required. Must be ≥ Trendyol sale price. | The "crossed-out" reference price. |
| Trendyol sale price | Required. Cannot exceed original price. | |
| Stock | Required (quantity). | |
| Stock code | Optional — seller's own internal tracking code only. | |
| VAT rate | Required. | Must be correct so invoices show correct VAT. |

**Variant/barcode uniqueness constraints** (all hard rejection rules):
- Only one barcode allowed per unique (category, model code, brand, color/distinguishing attribute, variant attribute like size) combination.
- Max 100 variants per product (counting existing + new combined).
- A duplicate model+barcode combination with a different variant value than what already exists is rejected as "same attributes already exist."

---

## 2. Image rules

### Technical spec (hard gate)
- Formats: **JPEG, JPG, PNG, WEBP**
- File size: **1 KB – 10 MB**
- Resolution: **minimum 860×574px, maximum 2000×2000px** (rejection copy elsewhere says 20×20–2000×2000 for the raw technical floor, but 860×574 is the enforced practical minimum)
- Count: **1–8 images per product**
- URL must start with `https://`, resolve directly to an image file (not a webpage), not be a Drive/Dropbox-style share link, contain no spaces/parentheses/Turkish characters/most special characters, and the hosting server must be reachable by Trendyol (no auth wall, no rate-limiting, no IP block). Trendyol's server IPs that must be allow-listed: `89.32.128.66`, `89.32.128.54-55`, `89.40.131.40-41`, `89.40.131.76-77`.

### Content rules (quality/compliance gate)
- Product must be fully visible and occupy most of the frame; front-facing/straight-angle main image.
- Only the exact variant being sold (color/model/size) shown in the main image — no other variants mixed in.
- Plain white/light background preferred.
- No: links to other sites, seller contact info, shipping/return/price/discount/gift messaging, unrelated logos/watermarks, excessive text overlay, heavy Photoshop/artificial editing, low-res/blurry/dark/cropped images, non-professional backdrops, accessories/contents not actually included in the package (unless explicitly labeled as "not included").
- Image must be consistent with title/description/brand/category — mismatch is an explicit rejection reason and a Content Quality violation (2 violation points + immediate closure).
- No nudity, sexual content, or content endangering child privacy (this also triggers "Objectionable Image Found" and can cause account-level violation points, not just product closure).
- Placeholder/dummy images at creation time are prohibited, and images can't later be swapped to represent a materially different product ("Image Mismatch").

### Full list of image-specific rejection reasons (from Product Rejection Details)
Invalid image link · server connection error / IP not allow-listed · wrong format/size/resolution · image mismatch vs. previously approved product · image & brand incompatible (looks like it belongs to a different registered trademark) · image limit exceeded (>8) · no image found at the link · image link inside HTML body not a direct file link · objectionable image content · image quality issue (blur/dark/text-heavy/no plain background) · banned word visible inside the image itself.

---

## 3. Banned/prohibited content in title, description, attributes, and images

A rejection applies uniformly across **title, description, attributes, and images** for:
- Offensive language, hate speech, discriminatory content.
- Content that references shipping, returns, campaigns, price, discount, or "free gift" mechanics inside the listing itself (these belong in Trendyol's own systems, not seller-authored copy).
- Personal contact info (phone, email, address), links to other websites/sales channels, or requests to contact outside Trendyol.
- Health claims that are unauthorized, misleading, or unsubstantiated (a very large enumerated banned-term list is provided — see §7 below for the pattern; covers most disease names, "detox," "cures," "boosts immunity," "weight loss," sexual-performance claims, etc.). This is explicitly checked against product name, description, image text, and any other listing field.
- Use of a language/alphabet not supported in the sale region (title/description/attributes must be in a supported regional language or Latin alphabet).

---

## 4. Category, brand, and Buy Box integrity rules

- **Buy Box / single-content model**: Trendyol groups all sellers of the "same" product under one content record keyed by barcode. Rules seller must follow:
  - Don't create multiple separate content/barcodes for what is actually the same product ("barcode multiplexing").
  - Don't list a different product than what a pre-existing barcode's content describes.
  - Don't alter shared content in a way that harms other sellers on the same Buy Box.
  - Sending a customer something different from the content/image/description is a Buy Box violation.
- **Category-brand restriction**: some categories only accept specific approved brands; entering a mismatched brand+category pair is rejected.
- **Restricted/global-barcode brands**: some brands require an internationally valid global barcode (EAN/UPC/GTIN) — a non-conforming barcode is rejected outright.
- **Malicious miscategorization** (to dodge commission rate or category restrictions) is a distinct, heavily-weighted violation (12 points) separate from an honest "wrong category" content error (2 points).
- **Passive category/brand**: category or brand no longer accepts new listings — must move to an active equivalent.
- Category, brand, model code, and barcode are all **immutable post-approval** — errors here can't be silently corrected later; they require re-listing or a support ticket.

---

## 5. Title & description quality guidance (soft — not hard-gated, but suppresses sales / risks manual flags)

**Title:**
- No word repetition; capitalize each word; include real customer search keywords.
- Don't restate barcode-driven variant attributes (color/size) already captured as attributes.
- Avoid ALL CAPS, emojis, excess punctuation/decoration, long meaningless phrasing.
- No exaggerated/misleading claims (drives up return rate and negative reviews).
- Don't abbreviate model/attribute terms.
- Title must not equal just the category name, must not contain seller name/brand/barcode/stock info (brand exclusion is separate from and in addition to the 3–200 char limit which itself excludes brand length).

**Description:**
- Prefer bullet points over long paragraphs.
- Naturally include likely search keywords.
- Explain use-case/where-used/attributes without needless complexity.
- Same exaggeration/misleading-claim penalty as title.
- No seller name, no stock info, no emojis/caps/decoration abuse.

---

## 6. Content Quality — violation classes that close a listing outright (Platform Rules §"Content Quality")

Each of these results in **immediate product closure** plus violation points against the seller account (points shown):

| Violation | Points |
|---|---|
| Listing/sending an illegal product | 48 |
| Listing/sending a product posing a health risk (e.g. via category/name/image manipulation to bypass QC) | 24 |
| Listing/sending a product prohibited for online sale | 4 |
| Repeatedly listing a product previously flagged "unsafe" by a market authority | 48 |
| Incorrect/misleading product info (false claims about nature/features/warranty) | 4 |
| **General "Violation of Content Quality"** — wrong/misleading brand, image mismatch, barcode/content multiplexing, wrong category, wrong variant info, or cancel-to-relist-higher | 2 |
| Buy Box violation (wrong product sent vs. listing, barcode misuse, damaging edits to shared content) | 2 |
| Inappropriate image usage (sexual/fantasy content shown explicitly) | 6 |
| Unsubstantiated/misleading health claim | 4 |
| Deliberately obfuscated health claim (misspellings, symbols, foreign language to dodge filters) | 12 |
| Malicious miscategorization to bypass restrictions/commission | 12 |

Separately, **originality/IP violations** (counterfeit, non-original documents, suspected inauthenticity) carry the highest point values (48) and fastest-track to suspension.

---

## 7. Products banned outright from the platform (partial list, high-signal items for QC)

Weapons/ammunition/explosives; live animals/plants/seeds/pesticides; medicines/drugs/medical devices (incl. hearing aids, prosthetics, COVID test kits); alcohol/tobacco/vaping products; unregistered wireless/telecom devices and surveillance/eavesdropping equipment; counterfeit/unlicensed/unauthorized-brand goods; pornographic or nudity content (fantasy/adult items allowed only with +18 label and no live models); gift vouchers/discount coupons/virtual money/stocks; fresh/cold-chain-dependent food (meat, poultry, fish, daily milk); items referencing prohibited symbols/regional-sensitivity content; smart watches with cameras, commercial drones; and (region-specific for CEE/GULF) additional restricted categories per the linked prohibited-products PDFs.

Restricted health-claim vocabulary is extensive (cancer, diabetes, weight-loss, immune-boosting, sexual-performance, organ-specific "treats/cures" language, etc.) — a QC keyword-scan against this list is directly actionable and mirrors what Trendyol itself scans for.

---

## 8. Operational quality signals (not content, but feed the same seller/product risk score)

These don't block a single listing's content approval but affect account standing and product visibility, and are useful as secondary QC/report signals if you're modeling seller risk, not just listing content risk:
- On-time shipping rate, out-of-stock rate, defective/wrong/missing return rate, disputed-return rate, return response time, average product rating, customer question response time, and current violation-point total all roll into a weekly **Seller Score**.
- Seller "operation status" (Green/Yellow/Orange/Red) throttles visibility based on shipping-capacity vs. order-volume mismatches — Red = temporarily hidden from search/favorites even though purchasable.
- Return non-returnability rules by category (hygiene items, cosmetics, food/supplements, digital licenses, furniture sets, large appliances once installed) matter if you want to flag listings that *should* declare non-returnability or warranty terms but don't.

---

## 9. Suggested QC rule table (ready to encode)

A starter rule table in the same shape as the existing Gatekeeper flags (`group`, `weight`, `message`), split into **hard** (block/reject-equivalent) and **soft** (quality/risk) tiers:

**Hard (mirror Trendyol's own rejection reasons — near-certain rejection if violated):**
1. `title_length` — title not 3–200 chars excluding brand.
2. `description_length` — plain text >4000 chars or HTML >30000 chars.
3. `barcode_format` — barcode not 2–40 chars or contains `? / & % + ^ ' * _`.
4. `model_code_format` — model code not 1–40 chars or contains URL/email.
5. `image_count` — 0 images or >8 images.
6. `image_technical_spec` — any image outside JPEG/JPG/PNG/WEBP, outside 1KB–10MB, outside 860×574–2000×2000px, or URL not `https://`/direct-file/public.
7. `banned_words` — banned/health-claim/contact-info/off-platform-redirect terms found in title, description, attributes, or (if OCR'd) images.
8. `price_consistency` — sale price > original price.
9. `category_taxonomy` — category not found in taxonomy or resolves to a non-leaf node (already implemented in the Trustana Gatekeeper; same logic applies here against a Trendyol category export).
10. `variant_duplicate` — duplicate (category, model code, brand, color, size) combination.

**Soft (quality/risk signals, don't block but should surface):**
11. `title_quality` — ALL CAPS ratio, repeated words, emoji presence, category-name-only title.
12. `description_quality` — no bullet points on long text, seller-name/stock-code leakage.
13. `image_content_quality` — non-white background heuristic, multiple variants visible, watermark/logo detection, image/title keyword mismatch.
14. `brand_category_mismatch` — brand not typically associated with the given category in your own product history (proxy for the brand-category restriction Trendyol enforces server-side).
15. `content_relevance` — title/description/image/category keyword overlap (same technique already used in the Gatekeeper's existing content-relevance check).

---

## 10. Gaps / things this pass could not verify

- Per-category **mandatory attribute lists** (which attributes are required for which category) are managed inside Trendyol's own category tree/Excel template, not published as static text — would need to be pulled from the seller panel's live category tree/attribute export, not from Seller Information Center pages.
- Exact list of **banned words/terms** (beyond the health-claims list captured in full above) is enforced server-side and not fully published — Trendyol only says "banned word(s) detected," without a public master list for general profanity/off-platform-redirect terms.
- The video-based Trendyol Academy training modules (Product Rejection Reasons, Product Listing, etc.) could not be opened in this session — every individual course page redirected to a login-gated view. If you can log in with your own seller credentials, those videos may add examples/screenshots but are unlikely to contain rules beyond what's captured here, since the written Seller Information Center pages are Trendyol's canonical reference and were clearly more detailed.

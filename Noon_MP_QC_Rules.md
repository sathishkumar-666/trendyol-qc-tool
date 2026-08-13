# noon Marketplace — Listing Content & QC Rules Reference

Compiled from noon's public/partner Seller Help Center
(helpcenter.noon.partners) and Support Portal (support.noon.partners),
specifically: **Product Listing Policy**, **Product Quality Control
Policy**, **Image Requirements and Rejection Reasons**, **Title
Requirements and Rejection Reasons**, **Seller SKU Content Quality**,
**Feature Bullets & Description**, **Barcode Requirements**, **Prohibited
Products**, **Content Language Requirements**, **Product Attributes**,
**Warranty Requirements**, **How to Categorize Seller SKUs**, and the
**Listing Products on noon FAQ**. These pages were reached via the seller's
own authenticated Partner Catalog session (project PRJ13591) and are
otherwise part of noon's public seller help center.

Organized the same way as `Trendyol_MP_QC_Rules.md` (rule → condition →
severity → source) so both feed the same QC tool structure.

---

## 1. Mandatory fields & format limits (product listing)

| Field | Rule | Notes |
|---|---|---|
| Title | Required. **20–200 characters.** | Capitalize first letter of each word (except prepositions/conjunctions/articles). No special characters (`@ ^ * # & ` etc.). No brand name, promotional phrases, prices, shipping mentions, irrelevant keywords, banned words, misspellings, or repeated words. |
| Description | Required (optional for go-live, but content-quality-scored). **250–4000 characters.** | Must be unique/accurate. No brand history/seller info, no price/shipping/warranty mentions, no ALL CAPS, no bold/italic, no emojis, no hyperlinks, no info contradicting title/images/bullets/specs. |
| Feature bullets | Optional but recommended. **~5 bullets, max 250 characters each.** | Semicolon-separated phrases; capitalize first letter; no end punctuation; no special characters, promo language, competitor-marketplace mentions, emojis, HTML tags, or hyperlinks. |
| Barcode | Required for go-live. **Alphanumeric, up to 16 characters.** No symbols (`+ / - % &` etc.) or spaces. Partner SKU or model number acceptable as a substitute. | Unique per SKU/variant. Immutable once assigned (can be replaced with a new one, but the old one stays in system records). Missing barcode = automatic rejection. |
| Images | Required — **minimum 1 for live status; 3+ recommended.** No stated hard maximum in the published rules. | See §2 for full technical spec. |
| Category | Required. Three-level hierarchy: **family → product type → product subtype.** | Family selection is most consequential — it determines which attributes/product types are even selectable. Miscategorization → Quality Check Team rejection. |
| Brand | Required. Must be registered/whitelisted if gated. | Certain brands are noon-controlled ("NSKU") or gated to approved sellers only. |
| Warranty type | Required. Exactly one of **manufacturer / seller / no_warranty** (lowercase in bulk templates). | Must be "seller" warranty if: international stock lacks brand warranty in destination country, product is renewed/refurbished (without official brand warranty), or product is modified/upgraded without manufacturer confirmation. Duration as a plain number (e.g. `24`), not "24 months". |
| Content language | English **mandatory** for UAE and KSA; Arabic optional but advisable. Egypt accepts **either** English or Arabic. | No auto-translation — sellers must supply each language's content themselves. |
| Product attributes | Mandatory attributes vary by category (star-marked in Seller Lab, green in bulk templates); optional attributes are "good to have." | Values must come from predefined dropdown lists where applicable; must map to the correct attribute field (e.g. don't put product type info into "Model Name"). |

---

## 2. Image rules

### Technical spec (hard gate)
- Format: **JPEG or JPG only** (no PNG/WEBP).
- Minimum width: **660px**, minimum aspect ratio **≥ 0.5**.
- Resolution: **≥ 72 PPI**.
- File size: **must not exceed 10 MB.**
- Color space: **RGB/sRGB.**
- Minimum 1 image for live status; 3+ recommended for full quality.

### Content rules (quality/compliance gate)
- Primary image: front view, pure white background (light grey for fashion), light shadows only (no hard shadows/reflections), product outside its packaging, no boxes/bags/tags visible, no lifestyle shots as the primary image.
- Secondary images: must match title/description, high resolution (no blur/pixelation), photographs only (no CAD/illustrations/thumbnails), product covering 70–80% of the frame, no logos/watermarks/unwanted text/accessories not included, no placeholder ("image coming soon") images.
- No duplicate images across the same product's image set.
- Standard sequencing: front → back → side → other angles → details → usage.
- Images must not be cropped in a way that hides part of the product.

### Rejection reasons captured
Need product image without packaging · cropped/half image · low resolution/blurry · duplicate images · conflicting images (shows a different product than what's for sale) · needs editing (borders, price info, dates, watermarks, hard shadows) · broken image link/expired URL · unclear image (needs another view) · incorrect image sequencing.

---

## 3. Banned/prohibited content in title, description, bullets, and images

Applies across **title, description, feature bullets, and images**:
- Special characters (`@ ^ * # & $` etc.), promotional/price/shipping language, competitor-marketplace mentions (e.g. Amazon, Namshi), emojis/emoticons, HTML tags, hyperlinks (including YouTube links or email addresses), brand history/seller-identifying information.
- Grammatical/spelling errors, word repetition (including across auto-concatenated attribute values), information that contradicts other fields (title vs. description vs. images vs. specs).

---

## 4. Prohibited products (banned outright from the platform)

General: narcotic drugs, telecom jamming equipment, GPS trackers, 2G network devices, SIM boxes, wireless scanners, goods bearing Israeli marks/logos/flag, goods from boycotted countries, endangered-species products, gambling equipment, certain fishing nets, original artwork, used tires, radioactive materials, publications contradicting Islamic teachings, counterfeit currency, human body parts (except hair for wigs), homemade food, firearms/weapons (kitchen knives excepted), prescription drugs, government IDs, tobacco, irradiated food, poppy seeds, materials for illegal activity, adult products, alcohol, livestock, sex toys/libido enhancers/male-enhancement supplements.

Also platform-wide: **Food & Beverage is fully gated** for marketplace sellers (cannot be listed under the marketplace model at all), and **pet food is strictly prohibited**.

Country-specific: UAE — 2G-only telecom equipment restricted per TDRA. KSA — rainbow-colored products, anti-Islamic symbols/signage, camera-equipped everyday devices (watches, key fobs, pens) and drones. Egypt — drones (toy or otherwise, with or without camera).

---

## 5. Content Quality Control Policy — account-level violations

- **Counterfeit products** — strictly prohibited; "fake goods designed to deceive customers into believing they're genuine."
- **IP infringement** — listings must not violate copyrights/trademarks/patents/trade secrets.
- **Listing accuracy** — content must precisely match the actual item sold.
- **Seller Controllable Returns (SCR)** — returns caused by fake/wrong/damaged/expired/empty-shipment products are tracked; sellers/brands must stay at "Fair" or better (Poor/Very Poor trigger penalties).
- **Seller Rating** — minimum **3.0** required; below that triggers required corrective action.
- Penalties for violations: delisting, temporary/permanent privilege revocation, government reporting, legal action, and a stated minimum monetary penalty (AED/SAR 200,000 or EGP 1,600,000) plus case-specific additions. Sellers get 3 calendar days to appeal with authenticity documentation before a penalty becomes final.

---

## 6. Category & brand integrity rules

- Family → product type → product subtype must all be selected consistently; an incorrect family selection cascades into wrong type/subtype/attribute options.
- Miscategorization examples noon itself gives: Bluetooth headphones under "Sports and Outdoors" instead of "Electronics accessories"; a sleep supplement under "Baby Product" instead of "Health and Nutrition"; Apple Watch bands under "Phone Accessories" instead of "Wearables."
- Category can be changed while a SKU is pending/rejected for a categorization reason; **cannot be changed once approved/live.**
- Duplicate SKU creation for the same product is explicitly prohibited (consequences: delisting or catalog-rights restriction).
- Brand/category gating: certain brands and categories (e.g. regulated/sensitive categories) are restricted to whitelisted sellers only, similar in spirit to Trendyol's category-brand restriction.

---

## 7. Operational / timing notes (not content, but useful context)

- Content QC decision: within ~2–3 business days of submission.
- Price/stock updates: reflected on-site within ~30 minutes.
- Monthly new-product creation limit: 25,000 (with size variants under one parent counting as 1, but color variants counting individually).

---

## 8. Suggested QC rule table (mapped onto the shared tool structure)

**Hard (rejection-equivalent):**
1. `title_length` — title not 20–200 characters.
2. `description_length` — description present but outside 250–4000 characters (noon has no separate HTML-length mode).
3. `barcode_format` — barcode not ≤16 alphanumeric characters, or contains a forbidden symbol.
4. `feature_bullet_length` — any feature bullet over 250 characters (optional field — informational if entirely absent from the export).
5. `image_count` — 0 images (no stated hard maximum for noon, unlike Trendyol's 8).
6. `image_technical_spec` — image outside JPEG/JPG, below 660px width / 0.5 aspect ratio / 72 PPI, or over 10MB (requires `--verify-images`).
7. `banned_words` — banned/prohibited-product/competitor-mention/off-platform terms found in title, description, or feature bullets.
8. `category_present` — category missing.
9. `variant_duplicate` — duplicate barcode or duplicate SKU across rows.
10. `warranty_type_format` — warranty type not exactly one of manufacturer/seller/no_warranty (optional field in most Trustana exports — informational if entirely absent).

**Soft (quality/risk signals):**
11. `category_depth_heuristic` — category path looks shallower than the expected 3-level family/type/subtype structure.
12. `title_quality` — ALL CAPS, repeated words, emoji, brand name present, title equals category name, excess punctuation.
13. `description_quality` — no bullet-style structure on long text, internal SKU/stock-code leakage, emoji.
14. `content_relevance` — category keywords vs. product name/title/description keyword overlap.
15. `title_description_relevance` — title vs. description keyword overlap (copy-paste/template-mismatch proxy).

---

## 9. Gaps / things this pass could not verify

- The **exact banned-word master list** (beyond the health/prohibited-product terms captured above) is not fully published — similar gap to Trendyol's own admission that it doesn't publish a complete profanity/off-platform-redirect list.
- **Per-category mandatory attribute lists** live inside noon's own category tree/attribute templates in Seller Lab, not as static published text — same category of gap as Trendyol's per-category attribute requirements.
- **Feature bullets and warranty type are not present as distinct columns** in the Trustana export tested against this tool (`Products_151 products.csv`) — the tool treats these as optional/informational rather than per-product hard failures when the whole column is absent, mirroring how Trendyol's Price/Stock/Model Code gaps were handled.
- Country/market context (which of UAE/KSA/Egypt a given SKU targets) is not present in the tested export, so the **language-requirement rule is not enforced per-product** — flagged as a file-level limitation rather than guessed at.

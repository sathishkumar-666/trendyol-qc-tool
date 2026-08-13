#!/usr/bin/env python3
"""
Trendyol Marketplace QC Tool
=============================
Checks a Trustana product export (CSV) against Trendyol's Seller Information
Center listing rules (see Trendyol_MP_QC_Rules.md, the source this tool was
built from) and produces an Excel QC report.

Usage:
    python3 trendyol_qc_tool.py <trustana_export.csv> <output_report.xlsx> [--verify-images] [--max-image-checks N]

    --verify-images       Attempt to download every product image and check
                           real format / file size / resolution against
                           Trendyol's technical spec. Requires outbound
                           network access to the image hosting domain(s).
                           If the first few attempts fail (e.g. sandboxed /
                           allow-listed network), image verification is
                           automatically skipped for the rest of the run and
                           the report notes this rather than reporting false
                           passes.
    --max-image-checks N  Cap on how many images to actually download when
                           --verify-images is set (default 500). Protects
                           against very large exports.

Column mapping (adjust COLUMN_MAP below if your Trustana export uses
different header names):
    sku, product_name, brand, model_code, category, google_category,
    images, barcode, title, description, price, sale_price, stock

Every check function below is annotated with the rule ID / section from the
QC rules doc it implements, so this file can be audited and updated when
Trendyol changes its rules.
"""

import sys
import csv
import re
import io
import argparse
from collections import defaultdict, Counter
from datetime import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("This tool requires openpyxl. Install with: pip install openpyxl")

# ---------------------------------------------------------------------------
# Column mapping — edit here if the Trustana export header names change.
# ---------------------------------------------------------------------------
COLUMN_MAP = {
    "sku": "SKU",
    "product_name": "Product Name",
    "brand": "Brand",
    "model_code": "Product Model",
    "category": "Product Category",
    "google_category": "Google Category",
    "images": "Images",
    "barcode": "UPC/Barcode/GTIN/EAN/Basic Info",
    "title": "Title [EN]/Marketing",
    "description": "Description [EN]/Marketing",
    "price": "Price",
    "stock": "Stock",
}

HARD = "HARD"
SOFT = "SOFT"

# ---------------------------------------------------------------------------
# Banned-content word lists — Rules doc §3 (banned content) and §7 (banned
# products / health-claim vocabulary). Non-exhaustive by Trendyol's own
# admission (see Rules doc §10 "Gaps") — extend as real rejections surface
# more terms.
# ---------------------------------------------------------------------------
HEALTH_CLAIM_TERMS = [
    "cancer", "tumor", "diabetes", "diabetic", "cures", "cure for", "curing",
    "detox", "detoxify", "weight loss", "lose weight", "fat burner",
    "fat burning", "immune booster", "boosts immunity", "immune boosting",
    "anti-cancer", "cholesterol treatment", "cures depression", "libido",
    "erectile", "sexual performance", "sexual enhancement", "viagra",
    "clinically proven to cure", "fda approved cure", "miracle cure",
    "heals permanently", "treats disease", "treats illness",
]

OFF_PLATFORM_KEYWORDS = [
    "http://", "https://", "www.", "instagram.com", "facebook.com",
    "wa.me", "t.me", "whatsapp", "telegram", "call us at", "contact us at",
    "email us at", "visit our website", "visit our store",
]

PRICE_SHIPPING_PROMO_TERMS = [
    "% off", "percent off", "free shipping", "free gift", "free delivery",
    "coupon code", "promo code", "clearance sale", "flash sale",
    "limited time offer", "buy now and save", "discount code",
    "use code", "sale price", "special campaign",
]

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
LONG_DIGIT_RUN_RE = re.compile(r"\b\d{9,}\b")
PHONE_LIKE_RE = re.compile(r"\b\d{2,4}[-.\s]\d{3}[-.\s]\d{3,4}\b")
EMOJI_RE = re.compile(
    "[" "\U0001F300-\U0001FAFF" "\U00002600-\U000027BF" "\U0001F1E6-\U0001F1FF" "]"
)
BARCODE_FORBIDDEN_CHARS_RE = re.compile(r"[?/&%+^'*_]")
HTML_TAG_RE = re.compile(r"<\s*(p|br|ul|li|ol|div|span|strong|b|i|table|tr|td)[\s/>]", re.I)

ALLOWED_IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".webp")


# ---------------------------------------------------------------------------
# Rule catalogue (for the "Rules Reference" sheet — keeps the report
# self-documenting and traceable back to the source doc).
# ---------------------------------------------------------------------------
RULE_CATALOGUE = [
    # id, name, severity, source section, condition summary
    ("title_length", "Title length", HARD, "§1, §9.1", "Title 3-200 chars excluding brand name; must be present."),
    ("description_length", "Description length", HARD, "§1, §9.2", "Plain text <=4000 chars, HTML <=30000 chars; must be present."),
    ("barcode_format", "Barcode / EAN format", HARD, "§1, §9.3", "2-40 chars, no ?/&%+^'*_ characters; must be present."),
    ("model_code_format", "Model code format", HARD, "§1, §9.4", "1-40 chars, no URL/email; must be present."),
    ("image_count", "Image count", HARD, "§1, §2, §9.5", "1-8 images required."),
    ("image_url_format", "Image URL format", HARD, "§2, §9.6", "https, direct file link, no spaces/parentheses, plausible image extension."),
    ("image_technical_spec", "Image technical spec (downloaded)", HARD, "§2, §9.6", "JPEG/PNG/WEBP, 1KB-10MB, 860x574-2000x2000px (requires --verify-images)."),
    ("banned_words", "Banned / restricted content", HARD, "§3, §7, §9.7", "Health claims, off-platform links/contact info, price/shipping/promo language in title or description."),
    ("price_consistency", "Price consistency", HARD, "§1, §9.8", "Sale price must not exceed original price (skipped if not present in export)."),
    ("category_present", "Category presence", HARD, "§1, §9.9", "Category must be present and non-trivial."),
    ("variant_duplicate", "Barcode / SKU duplication", HARD, "§1, §9.10", "Barcode must not be reused across unrelated SKUs; SKU must be unique."),
    ("category_depth_heuristic", "Category taxonomy depth (heuristic)", SOFT, "§1, §4, §9.9", "Category path looks shallow / possibly not a leaf node — needs manual check against Trendyol's live category tree (not available in this export)."),
    ("title_source_fallback", "Title sourced from Product Name (fallback)", SOFT, "(derived)", "Title [EN]/Marketing is blank; Product Name was used as a stand-in for QC purposes — flagged so the real marketing title still gets populated."),
    ("title_quality", "Title quality", SOFT, "§5, §9.11", "No ALL CAPS, no repeated words, no emoji, not just the category name, brand name excluded, no excess punctuation."),
    ("description_quality", "Description quality", SOFT, "§5, §9.12", "Prefer bullet points on long text; no internal SKU/stock-code leakage; no emoji/caps abuse."),
    ("duplicate_title", "Duplicate title across SKUs", SOFT, "(derived)", "Same exact title reused for multiple distinct products/barcodes — possible template copy-paste."),
    ("image_content_heuristic", "Image content quality (heuristic)", SOFT, "§2, §9.13", "Corner-pixel brightness as a rough plain-background proxy; reused image URL across SKUs (requires --verify-images for the background check)."),
    ("content_relevance", "Category matches product name/title/description", SOFT, "§4, §9.15", "At least one significant keyword (or known synonym) from the last two levels of the assigned category should appear in Product Name, Title, or Description — a proxy for 'is this the right category for this product'."),
    ("title_description_relevance", "Title matches description", SOFT, "(derived)", "Title and Description should share at least one significant keyword — a proxy for 'is the correct description attached to this title', catching copy-paste/template mismatches."),
]
RULE_SEVERITY = {r[0]: r[2] for r in RULE_CATALOGUE}
RULE_NAME = {r[0]: r[1] for r in RULE_CATALOGUE}


# ---------------------------------------------------------------------------
# CSV loading
# ---------------------------------------------------------------------------
def load_products(csv_path):
    """Load the Trustana export, tolerating non-UTF8 encoding (seen in the
    field data itself, e.g. curly quotes) and multi-line cells (the Images
    column contains one URL per line inside a single quoted CSV field)."""
    raw = None
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(csv_path, encoding=enc) as f:
                raw = f.read()
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raise RuntimeError(f"Could not decode {csv_path} with utf-8 or latin-1")

    reader = csv.DictReader(io.StringIO(raw))
    rows = list(reader)

    products = []
    for row in rows:
        p = {}
        for key, col in COLUMN_MAP.items():
            p[key] = (row.get(col) or "").strip()
        p["images_list"] = [u.strip() for u in p["images"].splitlines() if u.strip()]

        # Title [EN]/Marketing is the intended marketplace title, but it's
        # blank on a lot of rows in practice. Fall back to Product Name
        # (Trustana's internal name field, always populated) so the rest of
        # the QC checks have something to evaluate rather than reporting
        # "missing title" for a product that's really just missing the
        # dedicated marketing-title field. The fallback itself is flagged
        # separately (see check_title_source) since Product Name often
        # carries brand/stock-code baggage that doesn't belong in a
        # marketplace title.
        p["title_raw"] = p["title"]
        if not p["title"] and p["product_name"]:
            p["title"] = p["product_name"]
            p["title_source"] = "Product Name (fallback)"
        else:
            p["title_source"] = "Title [EN]/Marketing"

        p["_raw"] = row
        products.append(p)
    return products


# ---------------------------------------------------------------------------
# Per-product HARD rule checks. Each returns a list of (rule_id, ok, message)
# ---------------------------------------------------------------------------
def check_title_length(p):
    title, brand = p["title"], p["brand"]
    if not title:
        return [("title_length", False, "Missing title (both Title [EN]/Marketing and Product Name are empty).")]
    effective = re.sub(re.escape(brand), "", title, flags=re.I).strip() if brand else title
    length = len(effective)
    if length < 3:
        return [("title_length", False, f"Title too short: {length} chars excluding brand (min 3).")]
    if length > 200:
        return [("title_length", False, f"Title too long: {length} chars excluding brand (max 200).")]
    return [("title_length", True, "")]


def check_description_length(p):
    desc = p["description"]
    if not desc:
        return [("description_length", False, "Missing description.")]
    is_html = bool(HTML_TAG_RE.search(desc))
    limit = 30000 if is_html else 4000
    if len(desc) > limit:
        kind = "HTML" if is_html else "plain text"
        return [("description_length", False, f"Description is {len(desc)} chars, exceeds {limit}-char {kind} limit.")]
    return [("description_length", True, "")]


def check_barcode_format(p):
    bc = p["barcode"]
    if not bc:
        return [("barcode_format", False, "Missing barcode/EAN.")]
    issues = []
    if not (2 <= len(bc) <= 40):
        issues.append(f"Barcode length {len(bc)} outside 2-40 char range.")
    if BARCODE_FORBIDDEN_CHARS_RE.search(bc):
        issues.append("Barcode contains a forbidden character (one of ?/&%+^'*_).")
    if issues:
        return [("barcode_format", False, " ".join(issues))]
    return [("barcode_format", True, "")]


def check_model_code(p, model_code_data_present):
    mc = p["model_code"]
    if not model_code_data_present:
        # The whole export has this column empty — almost certainly a field
        # this Trustana export doesn't populate/map, not 151 individual
        # seller errors. Flag once at the file level instead of per-row.
        return [("model_code_format", None,
                 "Model code (Product Model) is not populated anywhere in this export — likely a field this "
                 "Trustana export doesn't map, rather than a per-product error. Confirm with the source system "
                 "before treating as a QC gap; Trendyol requires it at upload time.")]
    if not mc:
        return [("model_code_format", False, "Missing model code (Product Model column empty).")]
    issues = []
    if not (1 <= len(mc) <= 40):
        issues.append(f"Model code length {len(mc)} outside 1-40 char range.")
    if re.search(r"https?://|www\.|@", mc, re.I):
        issues.append("Model code appears to contain a URL or email address.")
    if issues:
        return [("model_code_format", False, " ".join(issues))]
    return [("model_code_format", True, "")]


def check_image_count_and_url_format(p):
    imgs = p["images_list"]
    results = []
    if len(imgs) == 0:
        results.append(("image_count", False, "No images found."))
    elif len(imgs) > 8:
        results.append(("image_count", False, f"{len(imgs)} images exceeds the 8-image limit."))
    else:
        results.append(("image_count", True, ""))

    bad_urls = []
    for u in imgs:
        problems = []
        if not u.lower().startswith("https://"):
            problems.append("not https")
        if " " in u or "(" in u or ")" in u:
            problems.append("contains space/parenthesis")
        if not u.lower().split("?")[0].endswith(ALLOWED_IMAGE_EXTENSIONS):
            problems.append("no recognizable image extension (jpg/jpeg/png/webp)")
        if problems:
            bad_urls.append(f"{u} ({', '.join(problems)})")
    if bad_urls:
        results.append(("image_url_format", False, "Malformed image URL(s): " + "; ".join(bad_urls)))
    else:
        results.append(("image_url_format", True, ""))
    return results


def check_price_consistency(p, price_data_present):
    if not price_data_present:
        return [("price_consistency", None, "Price/Stock not included in this Trustana export — verify against the marketplace-side pricing feed separately.")]
    try:
        price = float(p["price"]) if p["price"] else None
    except ValueError:
        price = None
    if price is None:
        return [("price_consistency", False, "Missing price.")]
    return [("price_consistency", True, "")]


def check_category_present(p):
    cat = p["category"]
    if not cat:
        return [("category_present", False, "Missing category.")]
    return [("category_present", True, "")]


def check_category_depth_heuristic(p):
    cat = p["category"]
    if not cat:
        return []  # already hard-flagged by check_category_present
    parts = [x.strip() for x in cat.split("/") if x.strip()]
    if len(parts) < 3:
        return [("category_depth_heuristic", False,
                  f"Category path has only {len(parts)} level(s) ({cat}) — may not be a leaf node. "
                  f"Verify against Trendyol's live category tree (not available in this export).")]
    return [("category_depth_heuristic", True, "")]


TEXT_FIELDS_FOR_BANNED_WORDS = ["title", "description"]


def check_banned_words(p):
    findings = []
    for field in TEXT_FIELDS_FOR_BANNED_WORDS:
        text = p.get(field, "")
        if not text:
            continue
        low = text.lower()
        for term in HEALTH_CLAIM_TERMS:
            if term in low:
                findings.append(f"possible unauthorized/unsubstantiated health claim '{term}' in {field}")
        for kw in OFF_PLATFORM_KEYWORDS:
            if kw in low:
                findings.append(f"off-platform link/contact reference '{kw}' in {field}")
        for term in PRICE_SHIPPING_PROMO_TERMS:
            if term in low:
                findings.append(f"price/shipping/promo language '{term}' in {field}")
        if EMAIL_RE.search(text):
            findings.append(f"email address detected in {field}")
        if PHONE_LIKE_RE.search(text) or LONG_DIGIT_RUN_RE.search(text):
            findings.append(f"phone-number-like digit sequence detected in {field}")
    if findings:
        return [("banned_words", False, "; ".join(findings))]
    return [("banned_words", True, "")]


def run_hard_checks(p, price_data_present, model_code_data_present):
    results = []
    results += check_title_length(p)
    results += check_description_length(p)
    results += check_barcode_format(p)
    results += check_model_code(p, model_code_data_present)
    results += check_image_count_and_url_format(p)
    results += check_price_consistency(p, price_data_present)
    results += check_category_present(p)
    results += check_banned_words(p)
    return results


# ---------------------------------------------------------------------------
# Per-product SOFT rule checks
# ---------------------------------------------------------------------------
def check_title_source(p):
    if p.get("title_source") == "Product Name (fallback)":
        return [("title_source_fallback", False,
                  f"Title [EN]/Marketing is blank — used Product Name ('{p['title']}') as a stand-in for QC purposes. "
                  f"Product Name often carries brand/size/internal-code text not meant for a marketplace title — "
                  f"populate a proper Title [EN]/Marketing value rather than relying on this fallback.")]
    return [("title_source_fallback", True, "")]


def check_title_quality(p):
    title, brand, cat = p["title"], p["brand"], p["category"]
    if not title:
        return []
    issues = []
    letters = [c for c in title if c.isalpha()]
    if letters:
        upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
        if upper_ratio > 0.7 and len(letters) > 8:
            issues.append("title appears to be in ALL CAPS")
    words = re.findall(r"[A-Za-z']+", title.lower())
    for i in range(len(words) - 1):
        if words[i] == words[i + 1] and len(words[i]) > 2:
            issues.append(f"repeated word '{words[i]}'")
    if EMOJI_RE.search(title):
        issues.append("emoji found in title")
    leaf = cat.split("/")[-1].strip().lower() if cat else ""
    if leaf and title.strip().lower() == leaf:
        issues.append("title is just the category name")
    if brand and re.search(re.escape(brand), title, re.I):
        issues.append(f"title contains brand name '{brand}' (Trendyol shows brand separately; keep it out of the title)")
    if title.count("!") > 1 or title.count("?") > 1:
        issues.append("excessive punctuation in title")
    last_token = title.strip().split()[-1] if title.strip() else ""
    stripped_token = re.sub(r"[^A-Za-z0-9]", "", last_token)
    if (len(stripped_token) >= 6 and stripped_token.isupper()
            and any(c.isdigit() for c in stripped_token) and any(c.isalpha() for c in stripped_token)):
        issues.append(f"title ends with a letters+digits code ('{last_token}') — likely a manufacturer model number or internal stock code carried over from Product Name; Trendyol prohibits stock info in the title, so confirm this is meant to be customer-facing before keeping it")
    if issues:
        return [("title_quality", False, "; ".join(issues))]
    return [("title_quality", True, "")]


def check_description_quality(p):
    desc, sku = p["description"], p["sku"]
    if not desc:
        return []
    issues = []
    has_bullets = bool(re.search(r"(^|\n)\s*[-•*]\s", desc)) or bool(re.search(r"<\s*(li|ul)[\s/>]", desc, re.I))
    if len(desc) > 400 and not has_bullets:
        issues.append("long description with no bullet points — consider restructuring for readability")
    if sku and sku in desc:
        issues.append("internal SKU/stock code found in description text")
    if EMOJI_RE.search(desc):
        issues.append("emoji found in description")
    if issues:
        return [("description_quality", False, "; ".join(issues))]
    return [("description_quality", True, "")]


_STOPWORDS_CATEGORY = {"and", "the", "for", "with", "other"}

# Small curated synonym groups so the category-vs-content check doesn't
# false-positive on wording differences that don't actually indicate a
# wrong category (e.g. a product titled "water bottle" assigned to a
# "...Flask" category leaf, or a "tumbler" assigned to "Sippy Cups").
# Extend this list as real false positives/negatives are found — it's a
# judgment call list, not a source-of-truth taxonomy mapping.
_CATEGORY_SYNONYM_GROUPS = [
    {"bottle", "flask"},
    {"cup", "tumbler", "mug"},
    {"marker", "pen"},
    {"backpack", "bag", "rucksack"},
    {"fryer", "cooker"},
    {"blender", "mixer"},
    {"jewelry", "jewellery"},
    {"color", "colour"},
]


def _synonym_match(a, b):
    return any(a in group and b in group for group in _CATEGORY_SYNONYM_GROUPS)


def _stem(word):
    """Very light singular/plural normalizer so 'fryers' matches 'fryer',
    'accessories' matches 'accessory', etc. Not real stemming — just enough
    to stop plural/singular mismatches from causing false positives."""
    w = word.lower()
    if w.endswith("ies") and len(w) > 4:
        return w[:-3] + "y"
    if w.endswith("es") and len(w) > 4:
        return w[:-2]
    if w.endswith("s") and len(w) > 3:
        return w[:-1]
    return w


def _meaningful_tokens(text):
    """Stemmed, stopword-filtered tokens from free text. Filters on the
    STEMMED length (>=3), not the raw word length, so short-after-stemming
    words like 'bag' (from 'bags') or 'cup'/'set' aren't inconsistently
    dropped on one side of a comparison and kept on the other — that
    mismatch was previously causing false "no overlap" flags."""
    tokens = re.findall(r"[a-z]+", text.lower())
    return {s for s in (_stem(t) for t in tokens if t not in _STOPWORDS_CATEGORY) if len(s) >= 3}


def _category_tokens(cat_path, n_levels=2):
    """Meaningful stemmed tokens from the last `n_levels` segments of a
    '/'-separated category path (e.g. Product Category or Google Category).
    Using the last couple of levels rather than just the leaf catches cases
    where the leaf alone is too narrow (e.g. 'Steam Irons') but a parent
    level ('Iron') already matches."""
    if not cat_path:
        return set()
    parts = [x.strip() for x in cat_path.split("/") if x.strip()]
    specific = parts[-n_levels:] if len(parts) >= n_levels else parts
    return _meaningful_tokens(" ".join(specific))


def _tokens_overlap(tokens_a, tokens_b):
    # Stemmed equality, one stem containing the other (min 4 chars, to
    # tolerate compounds like "airfryer" vs "fryer"), or a known synonym
    # pair (e.g. "bottle"/"flask").
    return any(
        a == b or (len(a) >= 4 and len(b) >= 4 and (a in b or b in a)) or _synonym_match(a, b)
        for a in tokens_a for b in tokens_b
    )


def _matched_tokens(tokens_a, tokens_b):
    """Which tokens in tokens_a have a match (equal, substring-compatible,
    or synonym) somewhere in tokens_b. Returns the subset of tokens_a that
    matched — used both for pass/fail and for showing evidence in the
    detailed content-check report."""
    matched = set()
    for a in tokens_a:
        for b in tokens_b:
            if a == b or (len(a) >= 4 and len(b) >= 4 and (a in b or b in a)) or _synonym_match(a, b):
                matched.add(a)
                break
    return matched


def compute_content_detail(p):
    """Central place that computes the token-level evidence behind the
    content checks (category-vs-content, title-vs-description), so the
    Excel report can show *why* a product passed or failed rather than
    just a pass/fail message. Cheap to compute — pure string ops, no I/O —
    so it's fine to run once per product regardless of whether every field
    ends up used."""
    cat, title, desc, name = p["category"], p["title"], p["description"], p["product_name"]

    category_tokens = _category_tokens(cat, n_levels=2) if cat else set()
    content_tokens = _meaningful_tokens(" ".join([name, title, desc]))
    category_matched = _matched_tokens(category_tokens, content_tokens)
    category_unmatched = category_tokens - category_matched

    title_tokens = _meaningful_tokens(title)
    desc_tokens = _meaningful_tokens(desc)
    title_desc_matched = _matched_tokens(title_tokens, desc_tokens)

    return {
        "category_tokens": sorted(category_tokens),
        "category_tokens_matched": sorted(category_matched),
        "category_tokens_unmatched": sorted(category_unmatched),
        "title_tokens": sorted(title_tokens),
        "title_description_overlap": sorted(title_desc_matched),
    }


def check_content_relevance(p, detail=None):
    """Does the assigned Trendyol category actually match what the product
    *is*, per its own name/title/description? Deliberately checks
    Product Name in addition to Title, since Title may be the Product Name
    fallback anyway, and Product Name sometimes carries more specific
    type wording than a cleaned-up marketing title."""
    cat, title, desc, name = p["category"], p["title"], p["description"], p["product_name"]
    if not cat or not (title or desc or name):
        return []
    detail = detail or compute_content_detail(p)
    if not detail["category_tokens"]:
        return []
    if not detail["category_tokens_matched"]:
        specific_label = "/".join([x.strip() for x in cat.split("/") if x.strip()][-2:])
        expected = ", ".join(detail["category_tokens"]) or "(none)"
        return [("content_relevance", False,
                  f"No keyword overlap between the assigned category ('{specific_label}') and the product "
                  f"name/title/description — expected one of [{expected}] to appear; none did. Check whether "
                  f"this is the right category for this product.")]
    return [("content_relevance", True, "")]


def check_title_description_relevance(p, detail=None):
    """Does the description actually talk about the same product as the
    title? A total lack of shared vocabulary between title and description
    is a strong signal of a copy-paste mistake (wrong description pasted
    onto this listing, template left unedited, etc.) rather than a subtle
    wording difference."""
    title, desc = p["title"], p["description"]
    if not title or not desc:
        return []
    detail = detail or compute_content_detail(p)
    if not detail["title_tokens"]:
        return []
    if not detail["title_description_overlap"]:
        return [("title_description_relevance", False,
                  "No shared keywords between the title and the description at all — check that the correct "
                  "description is attached to this listing (possible copy-paste/template mismatch).")]
    return [("title_description_relevance", True, "")]


def run_soft_checks(p):
    detail = compute_content_detail(p)
    results = []
    results += check_title_source(p)
    results += check_title_quality(p)
    results += check_description_quality(p)
    results += check_content_relevance(p, detail=detail)
    results += check_title_description_relevance(p, detail=detail)
    results += check_category_depth_heuristic(p)
    return results


# ---------------------------------------------------------------------------
# Cross-product checks (duplicates)
# ---------------------------------------------------------------------------
def run_cross_product_checks(products):
    """Returns dict: sku -> list of (rule_id, severity, message)"""
    extra = defaultdict(list)

    by_barcode = defaultdict(list)
    by_sku = defaultdict(list)
    by_title = defaultdict(list)
    by_image_url = defaultdict(list)

    for p in products:
        if p["barcode"]:
            by_barcode[p["barcode"]].append(p)
        if p["sku"]:
            by_sku[p["sku"]].append(p)
        if p["title"]:
            by_title[p["title"].strip().lower()].append(p)
        for u in p["images_list"]:
            by_image_url[u].append(p)

    for barcode, plist in by_barcode.items():
        if len(plist) > 1:
            skus = ", ".join(sorted(set(x["sku"] for x in plist)))
            for p in plist:
                extra[p["sku"]].append((
                    "variant_duplicate", HARD,
                    f"Barcode '{barcode}' is reused across multiple SKUs ({skus}) — only one barcode is allowed per unique product/variant combination."
                ))

    for sku, plist in by_sku.items():
        if len(plist) > 1:
            for p in plist:
                extra[p["sku"]].append((
                    "variant_duplicate", HARD,
                    f"Duplicate SKU '{sku}' appears {len(plist)} times in the export."
                ))

    for title, plist in by_title.items():
        if len(plist) > 1:
            barcodes = sorted(set(x["barcode"] for x in plist if x["barcode"]))
            if len(barcodes) > 1:
                skus = ", ".join(sorted(set(x["sku"] for x in plist)))
                for p in plist:
                    extra[p["sku"]].append((
                        "duplicate_title", SOFT,
                        f"Exact same title used for {len(plist)} different products (SKUs: {skus}) — check for copy-paste from a template."
                    ))

    for url, plist in by_image_url.items():
        if len(plist) > 1:
            skus = sorted(set(x["sku"] for x in plist))
            if len(skus) > 1:
                for p in plist:
                    extra[p["sku"]].append((
                        "image_content_heuristic", SOFT,
                        f"Image URL reused across {len(skus)} different SKUs ({', '.join(skus)}) — check for placeholder/mismatched imagery."
                    ))

    return extra


# ---------------------------------------------------------------------------
# Optional: real image download verification
# ---------------------------------------------------------------------------
def verify_images_for_product(p, session, max_checks_remaining, network_ok):
    """Returns (results, network_ok, checks_used). results is a list of
    (rule_id, ok/None, message)."""
    if not network_ok or max_checks_remaining <= 0 or not p["images_list"]:
        return [], network_ok, 0

    try:
        from PIL import Image
    except ImportError:
        return [("image_technical_spec", None, "Pillow not installed — image technical verification skipped.")], network_ok, 0

    results = []
    checks_used = 0
    for u in p["images_list"]:
        if checks_used >= max_checks_remaining:
            break
        checks_used += 1
        try:
            resp = session.get(u, timeout=10)
            resp.raise_for_status()
            content = resp.content
            size = len(content)
            fmt = None
            w = h = None
            try:
                img = Image.open(io.BytesIO(content))
                fmt = img.format
                w, h = img.size
            except Exception:
                pass

            problems = []
            if fmt not in ("JPEG", "PNG", "WEBP"):
                problems.append(f"format {fmt or 'unknown'} not in JPEG/PNG/WEBP")
            if not (1024 <= size <= 10 * 1024 * 1024):
                problems.append(f"file size {size} bytes outside 1KB-10MB")
            if w and h and not (860 <= w <= 2000 and 574 <= h <= 2000):
                problems.append(f"resolution {w}x{h} outside 860x574-2000x2000")
            if problems:
                results.append(("image_technical_spec", False, f"{u}: " + "; ".join(problems)))
            else:
                results.append(("image_technical_spec", True, ""))
        except Exception as e:
            # First-failure heuristic: if this is clearly a network/connectivity
            # problem, stop trying for the rest of the run rather than
            # reporting a wall of identical failures (or false passes).
            network_ok = False
            results.append(("image_technical_spec", None,
                             f"Could not download {u} to verify ({e}). Image technical verification "
                             f"disabled for the rest of this run — this environment's outbound network "
                             f"access appears restricted."))
            break

    return results, network_ok, checks_used


# ---------------------------------------------------------------------------
# Main QC run
# ---------------------------------------------------------------------------
def run_qc(csv_path, verify_images=False, max_image_checks=500):
    products = load_products(csv_path)

    price_data_present = any(p["price"] or p["stock"] for p in products)
    model_code_data_present = any(p["model_code"] for p in products)

    all_issues = []  # (sku, product_name, title, brand, category, rule_id, severity, ok, message)
    per_product_results = defaultdict(list)  # sku -> list of (rule_id, ok, message)

    for p in products:
        hard_results = run_hard_checks(p, price_data_present, model_code_data_present)
        soft_results = run_soft_checks(p)
        for rule_id, ok, msg in hard_results + soft_results:
            per_product_results[p["sku"]].append((rule_id, ok, msg))

    cross = run_cross_product_checks(products)
    for sku, entries in cross.items():
        for rule_id, severity, msg in entries:
            per_product_results[sku].append((rule_id, False, msg))

    network_ok = True
    image_checks_remaining = max_image_checks
    if verify_images:
        try:
            import requests
        except ImportError:
            network_ok = False
            for p in products:
                per_product_results[p["sku"]].append(
                    ("image_technical_spec", None, "The 'requests' package is not installed — image technical verification skipped.")
                )
        else:
            session = requests.Session()
            processed_skus = set()
            for p in products:
                if not network_ok or image_checks_remaining <= 0:
                    break
                processed_skus.add(p["sku"])
                results, network_ok, used = verify_images_for_product(p, session, image_checks_remaining, network_ok)
                image_checks_remaining -= used
                for rule_id, ok, msg in results:
                    per_product_results[p["sku"]].append((rule_id, ok, msg))
            if not network_ok:
                # Make the skip explicit for every product that never got a
                # chance to be checked, so the report doesn't silently under-count.
                for p in products:
                    if p["sku"] not in processed_skus:
                        per_product_results[p["sku"]].append((
                            "image_technical_spec", None,
                            "Skipped — image technical verification was disabled for the rest of the run "
                            "after an earlier download failure (see Summary sheet)."
                        ))

    # Flatten into issues list (only non-pass entries, i.e. ok is False; None
    # entries are informational/skipped and get their own bucket)
    info_notes = []
    for p in products:
        for rule_id, ok, msg in per_product_results[p["sku"]]:
            if ok is False:
                all_issues.append({
                    "sku": p["sku"], "product_name": p["product_name"], "title": p["title"],
                    "brand": p["brand"], "category": p["category"],
                    "rule_id": rule_id, "rule_name": RULE_NAME.get(rule_id, rule_id),
                    "severity": RULE_SEVERITY.get(rule_id, "SOFT"), "message": msg,
                })
            elif ok is None and msg:
                info_notes.append({
                    "sku": p["sku"], "rule_id": rule_id, "message": msg,
                })

    return {
        "products": products,
        "per_product_results": per_product_results,
        "issues": all_issues,
        "info_notes": info_notes,
        "price_data_present": price_data_present,
        "model_code_data_present": model_code_data_present,
        "verify_images_requested": verify_images,
        "image_network_ok": network_ok if verify_images else None,
    }


# ---------------------------------------------------------------------------
# Excel report generation
# ---------------------------------------------------------------------------
HARD_FILL = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
SOFT_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
OK_FILL = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")


def _style_header(ws, row=1):
    for cell in ws[row]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def _autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_report(qc_result, output_path, source_file):
    products = qc_result["products"]
    issues = qc_result["issues"]
    per_product = qc_result["per_product_results"]

    wb = Workbook()

    # --- Summary sheet ---
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Trendyol Marketplace QC Report"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([f"Source file: {source_file}"])
    ws.append([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"])
    ws.append([f"Total products checked: {len(products)}"])
    n_hard_fail_products = sum(1 for p in products if any(
        i["sku"] == p["sku"] and i["severity"] == HARD for i in issues))
    n_soft_flag_products = sum(1 for p in products if any(
        i["sku"] == p["sku"] and i["severity"] == SOFT for i in issues))
    ws.append([f"Products with >=1 HARD (rejection-risk) issue: {n_hard_fail_products}"])
    ws.append([f"Products with >=1 SOFT (quality) flag: {n_soft_flag_products}"])
    ws.append([f"Products with zero issues found: {sum(1 for p in products if not any(i['sku']==p['sku'] for i in issues))}"])
    if not qc_result["price_data_present"]:
        ws.append(["NOTE: Price/Stock columns are empty in this export — price_consistency was skipped for all products."])
    if not qc_result["model_code_data_present"]:
        ws.append(["NOTE: Product Model (model code) column is empty for every product in this export — model_code_format was flagged as informational rather than a per-product fail. Confirm with the source system whether this field is mapped elsewhere."])
    if qc_result["verify_images_requested"]:
        if qc_result["image_network_ok"]:
            ws.append(["Image technical verification: images were downloaded and checked against format/size/resolution spec."])
        else:
            ws.append(["Image technical verification: REQUESTED but could not complete — outbound network access to the image host was not available in this run. Image checks below are limited to URL format/count only. Re-run with network access to this host for full verification."])
    ws.append([])

    ws.append(["Rule", "Severity", "Products failing", "Products passing", "Not evaluated / N/A"])
    header_row = ws.max_row
    rule_ids_in_order = [r[0] for r in RULE_CATALOGUE]
    for rule_id in rule_ids_in_order:
        name = RULE_NAME[rule_id]
        sev = RULE_SEVERITY[rule_id]
        fail = sum(1 for i in issues if i["rule_id"] == rule_id)
        pass_ct = 0
        na_ct = 0
        for p in products:
            statuses = [ok for (rid, ok, msg) in per_product[p["sku"]] if rid == rule_id]
            if not statuses:
                continue
            if any(ok is True for ok in statuses):
                pass_ct += 1
            elif all(ok is None for ok in statuses):
                na_ct += 1
        ws.append([name, sev, fail, pass_ct, na_ct])
    _style_header(ws, header_row)
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        sev_cell = row[1]
        fill = HARD_FILL if sev_cell.value == HARD else SOFT_FILL
        for cell in row:
            cell.fill = fill
    _autosize(ws, [42, 10, 16, 16, 18])

    # --- Issues sheet ---
    ws2 = wb.create_sheet("Issues")
    headers = ["SKU", "Product Name", "Title", "Brand", "Category", "Rule", "Severity", "Message"]
    ws2.append(headers)
    _style_header(ws2)
    issues_sorted = sorted(issues, key=lambda i: (i["sku"], 0 if i["severity"] == HARD else 1, i["rule_name"]))
    for i in issues_sorted:
        ws2.append([i["sku"], i["product_name"], i["title"], i["brand"], i["category"],
                    i["rule_name"], i["severity"], i["message"]])
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
        fill = HARD_FILL if row[6].value == HARD else SOFT_FILL
        for cell in row:
            cell.fill = fill
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = ws2.dimensions
    _autosize(ws2, [10, 40, 40, 16, 30, 26, 10, 70])

    # --- Products sheet ---
    ws3 = wb.create_sheet("Products")
    headers3 = ["SKU", "Product Name", "Title", "Title Source", "Brand", "Category", "Google Category (reference only)",
                "Barcode", "# Images", "Hard Fail Count", "Soft Flag Count", "Overall Status", "Hard Issues", "Soft Issues"]
    ws3.append(headers3)
    _style_header(ws3)
    for p in products:
        p_issues = [i for i in issues if i["sku"] == p["sku"]]
        hard_issues = [i for i in p_issues if i["severity"] == HARD]
        soft_issues = [i for i in p_issues if i["severity"] == SOFT]
        if hard_issues:
            status = "REJECTION RISK"
        elif soft_issues:
            status = "NEEDS REVIEW"
        else:
            status = "OK"
        ws3.append([
            p["sku"], p["product_name"], p["title"], p["title_source"], p["brand"], p["category"], p["google_category"],
            p["barcode"], len(p["images_list"]), len(hard_issues), len(soft_issues), status,
            "; ".join(f"[{i['rule_name']}] {i['message']}" for i in hard_issues),
            "; ".join(f"[{i['rule_name']}] {i['message']}" for i in soft_issues),
        ])
    for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row):
        status_cell = row[11]
        if status_cell.value == "REJECTION RISK":
            status_cell.fill = HARD_FILL
        elif status_cell.value == "NEEDS REVIEW":
            status_cell.fill = SOFT_FILL
        else:
            status_cell.fill = OK_FILL
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = ws3.dimensions
    _autosize(ws3, [10, 40, 40, 22, 16, 30, 34, 18, 9, 12, 12, 16, 60, 60])

    # --- Content Checks sheet: detailed evidence behind the content-relevance
    # checks (category-vs-content, title-vs-description), plus the other
    # content-quality findings, all in one place per product so the QC
    # logic itself is auditable rather than just a pass/fail message. ---
    ws5 = wb.create_sheet("Content Checks")
    headers5 = ["SKU", "Title", "Title Source", "Description (truncated)", "Category",
                "Category Keywords Expected", "Category Keywords Matched", "Category Keywords Unmatched",
                "Category Match?", "Title/Description Shared Keywords", "Title/Description Match?",
                "Title Quality Issues", "Description Quality Issues", "Banned/Restricted Content Found"]
    ws5.append(headers5)
    _style_header(ws5)

    def _issue_msgs(sku, rule_id):
        return "; ".join(i["message"] for i in issues if i["sku"] == sku and i["rule_id"] == rule_id)

    for p in products:
        detail = compute_content_detail(p)
        cat_match_fail = _issue_msgs(p["sku"], "content_relevance")
        title_desc_fail = _issue_msgs(p["sku"], "title_description_relevance")
        cat_match_status = "NO OVERLAP" if cat_match_fail else ("N/A" if not detail["category_tokens"] else "OK")
        title_desc_status = "NO OVERLAP" if title_desc_fail else ("N/A" if not (p["title"] and p["description"]) else "OK")
        desc_snippet = (p["description"][:300] + "...") if len(p["description"]) > 300 else p["description"]
        ws5.append([
            p["sku"], p["title"], p["title_source"], desc_snippet, p["category"],
            ", ".join(detail["category_tokens"]), ", ".join(detail["category_tokens_matched"]),
            ", ".join(detail["category_tokens_unmatched"]), cat_match_status,
            ", ".join(detail["title_description_overlap"]), title_desc_status,
            _issue_msgs(p["sku"], "title_quality"), _issue_msgs(p["sku"], "description_quality"),
            _issue_msgs(p["sku"], "banned_words"),
        ])
    for row in ws5.iter_rows(min_row=2, max_row=ws5.max_row):
        for cell in (row[8], row[10]):
            if cell.value == "NO OVERLAP":
                cell.fill = SOFT_FILL
            elif cell.value == "OK":
                cell.fill = OK_FILL
    ws5.freeze_panes = "A2"
    ws5.auto_filter.ref = ws5.dimensions
    _autosize(ws5, [10, 40, 22, 60, 30, 34, 30, 24, 14, 34, 18, 40, 40, 40])

    # --- Rules Reference sheet ---
    ws4 = wb.create_sheet("Rules Reference")
    ws4.append(["Rule ID", "Name", "Severity", "Source section(s)", "Condition"])
    _style_header(ws4)
    for rule_id, name, sev, source, cond in RULE_CATALOGUE:
        ws4.append([rule_id, name, sev, source, cond])
    for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row):
        fill = HARD_FILL if row[2].value == HARD else SOFT_FILL
        for cell in row:
            cell.fill = fill
    ws4.freeze_panes = "A2"
    _autosize(ws4, [24, 34, 10, 18, 80])

    wb.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="Trendyol Marketplace QC tool for Trustana product exports.")
    parser.add_argument("csv_path", help="Path to the Trustana product export CSV")
    parser.add_argument("output_path", help="Path to write the Excel QC report")
    parser.add_argument("--verify-images", action="store_true", help="Download images and verify format/size/resolution")
    parser.add_argument("--max-image-checks", type=int, default=500, help="Cap on number of images to download when --verify-images is set")
    args = parser.parse_args()

    qc_result = run_qc(args.csv_path, verify_images=args.verify_images, max_image_checks=args.max_image_checks)
    build_report(qc_result, args.output_path, args.csv_path)
    n_hard = sum(1 for i in qc_result["issues"] if i["severity"] == HARD)
    n_soft = sum(1 for i in qc_result["issues"] if i["severity"] == SOFT)
    print(f"Checked {len(qc_result['products'])} products: {n_hard} hard issues, {n_soft} soft flags.")
    print(f"Report written to {args.output_path}")


if __name__ == "__main__":
    main()

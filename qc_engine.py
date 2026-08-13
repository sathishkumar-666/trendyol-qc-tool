#!/usr/bin/env python3
"""
Trustana AI Content Verifier for MP
====================================
Checks a Trustana product export (CSV) against a marketplace's listing
rules and produces an Excel QC report. Currently supports:

    - trendyol  (see Trendyol_MP_QC_Rules.md)
    - noon      (see Noon_MP_QC_Rules.md)

Usage:
    python3 qc_engine.py <trustana_export.csv> <output_report.xlsx> --marketplace trendyol|noon [--verify-images] [--max-image-checks N]

    --marketplace          Which marketplace's rules to check against.
                           Required.
    --verify-images        Attempt to download every product image and check
                           real format / file size / resolution against the
                           marketplace's technical spec. Requires outbound
                           network access to the image hosting domain(s).
                           If the first few attempts fail (e.g. sandboxed /
                           allow-listed network), image verification is
                           automatically skipped for the rest of the run and
                           the report notes this rather than reporting false
                           passes.
    --max-image-checks N   Cap on how many images to actually download when
                           --verify-images is set (default 500). Protects
                           against very large exports.

Column mapping (adjust COLUMN_MAP below if your Trustana export uses
different header names):
    sku, product_name, brand, model_code, category, google_category,
    images, barcode, title, description, price, stock, warranty_type

Adding a new marketplace: add a new MarketplaceConfig to MARKETPLACES near
the top of this file. Every check function below takes the active config
as a parameter rather than hardcoding thresholds, so a new marketplace
usually needs no changes to the check logic itself — just a new config
entry with that marketplace's numbers plugged in.
"""

import sys
import csv
import re
import io
import argparse
from dataclasses import dataclass, field
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
# Shared across marketplaces: both Trendyol and noon read from the same
# Trustana export shape. Fields a given marketplace doesn't use (e.g.
# warranty_type for Trendyol) are simply ignored for that marketplace.
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
    "warranty_type": "Warranty Type [EN]/Marketing",
}

HARD = "HARD"
SOFT = "SOFT"

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
LONG_DIGIT_RUN_RE = re.compile(r"\b\d{9,}\b")
PHONE_LIKE_RE = re.compile(r"\b\d{2,4}[-.\s]\d{3}[-.\s]\d{3,4}\b")
EMOJI_RE = re.compile(
    "[" "\U0001F300-\U0001FAFF" "\U00002600-\U000027BF" "\U0001F1E6-\U0001F1FF" "]"
)
HTML_TAG_RE = re.compile(r"<\s*(p|br|ul|li|ol|div|span|strong|b|i|table|tr|td)[\s/>]", re.I)


# ---------------------------------------------------------------------------
# Banned-content word lists. Each marketplace's config picks which of these
# groups apply. All non-exhaustive by the marketplaces' own admission (see
# the rules docs' "Gaps" sections) — extend as real rejections surface more
# terms. Highly ambiguous/context-dependent banned categories (e.g. political
# symbolism, religious-sensitivity content) are deliberately NOT keyword-
# scanned here — false positive risk is too high for a naive scanner, and
# they need actual category-level human review. See each rules doc's "Gaps"
# section for what's excluded and why.
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

COMPETITOR_MENTION_TERMS = [
    "amazon", "namshi", "noon.com/", "trendyol.com", "available on ebay",
    "available on shopify", "other marketplaces",
]

# noon Rules doc §4: prohibited-product categories. Limited to reasonably
# unambiguous, high-signal keywords — see the module docstring above for why
# the most context-dependent categories are intentionally excluded.
NOON_PROHIBITED_PRODUCT_TERMS = [
    "narcotic", "cocaine", "heroin", "controlled substance",
    "sim box", "gps tracker", "jamming device", "signal jammer",
    "counterfeit", "replica watch", "replica bag",
    "endangered species", "ivory", "rhino horn",
    "ammunition", "firearm", "handgun", "assault rifle", "explosive device",
    "radioactive material", "human organ",
    "prescription drug", "unregistered medicine",
    "tobacco", "cigarette", "e-cigarette", "vape liquid",
    "poppy seed", "counterfeit currency", "fake currency",
    "gambling equipment", "lottery ticket",
    "sex toy", "vibrator", "libido enhancer", "male enhancement supplement",
    "pet food", "live animal for sale",
]

BARCODE_FORBIDDEN_CHARS_TRENDYOL_RE = re.compile(r"[?/&%+^'*_]")
BARCODE_FORBIDDEN_CHARS_NOON_RE = re.compile(r"[+/\-%&\s]")


# ---------------------------------------------------------------------------
# Marketplace configuration
# ---------------------------------------------------------------------------
@dataclass
class MarketplaceConfig:
    key: str
    display_name: str
    rules_doc: str

    # Title
    title_min_len: int
    title_max_len: int
    title_length_excludes_brand: bool  # Trendyol: yes (rule text says "excluding brand"). noon: no (brand simply banned outright, length is on the raw title).

    # Description
    desc_min_len: int
    desc_max_len: int
    desc_html_max_len: "int | None"  # None = marketplace has no separate HTML-length mode

    # Feature bullets (optional field — currently only noon)
    has_feature_bullets: bool
    bullet_max_len: int = 250

    # Barcode
    barcode_min_len: int = 2
    barcode_max_len: int = 40
    barcode_forbidden_chars_re: "re.Pattern" = BARCODE_FORBIDDEN_CHARS_TRENDYOL_RE

    # Model code (a field distinct from barcode — currently only Trendyol)
    has_model_code: bool = True

    # Warranty (optional field — currently only noon)
    has_warranty_field: bool = False
    warranty_allowed_keywords: tuple = ("manufactur", "seller", "no warranty", "none", "no_warranty")

    # Images
    image_min_count: int = 1
    image_max_count: "int | None" = 8
    image_url_allowed_extensions: tuple = (".jpg", ".jpeg", ".png", ".webp")
    image_allowed_pil_formats: tuple = ("JPEG", "PNG", "WEBP")
    image_min_size_bytes: int = 1024
    image_max_size_bytes: int = 10 * 1024 * 1024
    image_min_w: "int | None" = 860
    image_max_w: "int | None" = 2000
    image_min_h: "int | None" = 574
    image_max_h: "int | None" = 2000
    image_min_aspect_ratio: "float | None" = None  # noon: width/height (or height/width) must be >= this
    image_min_ppi: "int | None" = None  # noon-specific; best-effort, see verify_images_for_product

    # Category
    category_min_levels: int = 3

    # Banned-word groups active for this marketplace: list of (label, terms)
    banned_word_groups: list = field(default_factory=list)

    # Brand-in-title rule wording (differs slightly by marketplace tone)
    brand_in_title_note: str = "brand shown separately; keep it out of the title"


TRENDYOL_CONFIG = MarketplaceConfig(
    key="trendyol",
    display_name="Trendyol",
    rules_doc="Trendyol_MP_QC_Rules.md",
    title_min_len=3,
    title_max_len=200,
    title_length_excludes_brand=True,
    desc_min_len=0,
    desc_max_len=4000,
    desc_html_max_len=30000,
    has_feature_bullets=False,
    barcode_min_len=2,
    barcode_max_len=40,
    barcode_forbidden_chars_re=BARCODE_FORBIDDEN_CHARS_TRENDYOL_RE,
    has_model_code=True,
    has_warranty_field=False,
    image_min_count=1,
    image_max_count=8,
    image_url_allowed_extensions=(".jpg", ".jpeg", ".png", ".webp"),
    image_allowed_pil_formats=("JPEG", "PNG", "WEBP"),
    image_min_size_bytes=1024,
    image_max_size_bytes=10 * 1024 * 1024,
    image_min_w=860, image_max_w=2000, image_min_h=574, image_max_h=2000,
    image_min_aspect_ratio=None,
    image_min_ppi=None,
    category_min_levels=3,
    banned_word_groups=[
        ("unauthorized/unsubstantiated health claim", HEALTH_CLAIM_TERMS),
        ("off-platform link/contact reference", OFF_PLATFORM_KEYWORDS),
        ("price/shipping/promo language", PRICE_SHIPPING_PROMO_TERMS),
    ],
    brand_in_title_note="Trendyol shows brand separately; keep it out of the title",
)

NOON_CONFIG = MarketplaceConfig(
    key="noon",
    display_name="noon",
    rules_doc="Noon_MP_QC_Rules.md",
    title_min_len=20,
    title_max_len=200,
    title_length_excludes_brand=False,
    desc_min_len=250,
    desc_max_len=4000,
    desc_html_max_len=None,
    has_feature_bullets=True,
    bullet_max_len=250,
    barcode_min_len=1,
    barcode_max_len=16,
    barcode_forbidden_chars_re=BARCODE_FORBIDDEN_CHARS_NOON_RE,
    has_model_code=False,
    has_warranty_field=True,
    warranty_allowed_keywords=("manufactur", "seller", "no warranty", "none", "no_warranty"),
    image_min_count=1,
    image_max_count=None,
    image_url_allowed_extensions=(".jpg", ".jpeg"),
    image_allowed_pil_formats=("JPEG",),
    image_min_size_bytes=1024,
    image_max_size_bytes=10 * 1024 * 1024,
    image_min_w=660, image_max_w=None, image_min_h=None, image_max_h=None,
    image_min_aspect_ratio=0.5,
    image_min_ppi=72,
    category_min_levels=3,
    banned_word_groups=[
        ("prohibited product", NOON_PROHIBITED_PRODUCT_TERMS),
        ("off-platform link/contact reference", OFF_PLATFORM_KEYWORDS),
        ("price/shipping/promo language", PRICE_SHIPPING_PROMO_TERMS),
        ("competitor-marketplace mention", COMPETITOR_MENTION_TERMS),
    ],
    brand_in_title_note="noon requires brand to be excluded from the title entirely",
)

MARKETPLACES = {
    "trendyol": TRENDYOL_CONFIG,
    "noon": NOON_CONFIG,
}


def get_marketplace_config(key):
    try:
        return MARKETPLACES[key]
    except KeyError:
        raise ValueError(f"Unknown marketplace '{key}'. Choose one of: {', '.join(MARKETPLACES)}")


# ---------------------------------------------------------------------------
# Rule catalogue — built per marketplace so thresholds shown match the
# active config, and so a marketplace only lists rules that actually apply
# to it (e.g. model_code_format only for Trendyol; feature_bullet_length
# and warranty_type_check only for noon).
# ---------------------------------------------------------------------------
def build_rule_catalogue(cfg):
    catalogue = [
        ("title_length", "Title length", HARD, cfg.rules_doc,
         f"Title {cfg.title_min_len}-{cfg.title_max_len} chars"
         + (" excluding brand name" if cfg.title_length_excludes_brand else "")
         + "; must be present."),
        ("description_length", "Description length", HARD, cfg.rules_doc,
         (f"{cfg.desc_min_len}-{cfg.desc_max_len} chars"
          + (f" plain text, up to {cfg.desc_html_max_len} chars if HTML" if cfg.desc_html_max_len else ""))),
        ("barcode_format", "Barcode format", HARD, cfg.rules_doc,
         f"{cfg.barcode_min_len}-{cfg.barcode_max_len} chars, no forbidden characters; must be present."),
        ("image_count", "Image count", HARD, cfg.rules_doc,
         f"At least {cfg.image_min_count} image(s) required"
         + (f", maximum {cfg.image_max_count}." if cfg.image_max_count else " (no stated maximum).")),
        ("image_url_format", "Image URL format", HARD, cfg.rules_doc,
         "https, direct file link, no spaces/parentheses, plausible image extension for this marketplace."),
        ("image_technical_spec", "Image technical spec (downloaded)", HARD, cfg.rules_doc,
         "Format/size/resolution per marketplace spec (requires --verify-images)."),
        ("banned_words", "Banned / restricted content", HARD, cfg.rules_doc,
         "Prohibited-product, off-platform, and price/promo language in title, description, or feature bullets."),
        ("price_consistency", "Price presence", HARD, cfg.rules_doc,
         "Price must be present and parseable (skipped file-wide if not present in export)."),
        ("category_present", "Category presence", HARD, cfg.rules_doc,
         "Category must be present and non-trivial."),
        ("variant_duplicate", "Barcode / SKU duplication", HARD, cfg.rules_doc,
         "Barcode must not be reused across unrelated SKUs; SKU must be unique."),
        ("category_depth_heuristic", "Category taxonomy depth (heuristic)", SOFT, cfg.rules_doc,
         f"Category path should have at least {cfg.category_min_levels} levels — heuristic, not validated against the live category tree."),
        ("title_source_fallback", "Title sourced from Product Name (fallback)", SOFT, "(derived)",
         "Title field is blank; Product Name was used as a stand-in for QC purposes."),
        ("title_quality", "Title quality", SOFT, cfg.rules_doc,
         "No ALL CAPS, no repeated words, no emoji, not just the category name, brand name excluded, no excess punctuation."),
        ("description_quality", "Description quality", SOFT, cfg.rules_doc,
         "Prefer bullet points on long text; no internal SKU/stock-code leakage; no emoji/caps abuse."),
        ("duplicate_title", "Duplicate title across SKUs", SOFT, "(derived)",
         "Same exact title reused for multiple distinct products/barcodes — possible template copy-paste."),
        ("image_content_heuristic", "Image content quality (heuristic)", SOFT, cfg.rules_doc,
         "Reused image URL across SKUs (requires --verify-images for background/heuristic checks)."),
        ("content_relevance", "Category matches product name/title/description", SOFT, cfg.rules_doc,
         "At least one significant keyword (or known synonym) from the category should appear in Product Name, Title, or Description."),
        ("title_description_relevance", "Title matches description", SOFT, "(derived)",
         "Title and Description should share at least one significant keyword."),
    ]
    if cfg.has_model_code:
        catalogue.insert(3, ("model_code_format", "Model code format", HARD, cfg.rules_doc,
                              "1-40 chars, no URL/email; must be present."))
    if cfg.has_feature_bullets:
        catalogue.append(("feature_bullet_length", "Feature bullet length", HARD, cfg.rules_doc,
                           f"Each feature bullet at most {cfg.bullet_max_len} characters (optional field — informational if entirely absent from the export)."))
    if cfg.has_warranty_field:
        catalogue.append(("warranty_type_check", "Warranty type classification", SOFT, cfg.rules_doc,
                           "Warranty value should clearly map to manufacturer/seller/no-warranty (optional field — informational if entirely absent from the export)."))
    return catalogue


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
def check_title_length(p, cfg):
    title, brand = p["title"], p["brand"]
    if not title:
        return [("title_length", False, "Missing title (both the marketing title field and Product Name are empty).")]
    if cfg.title_length_excludes_brand and brand:
        effective = re.sub(re.escape(brand), "", title, flags=re.I).strip()
    else:
        effective = title
    length = len(effective)
    excl = " excluding brand" if cfg.title_length_excludes_brand else ""
    if length < cfg.title_min_len:
        return [("title_length", False, f"Title too short: {length} chars{excl} (min {cfg.title_min_len}).")]
    if length > cfg.title_max_len:
        return [("title_length", False, f"Title too long: {length} chars{excl} (max {cfg.title_max_len}).")]
    return [("title_length", True, "")]


def check_description_length(p, cfg):
    desc = p["description"]
    if not desc:
        if cfg.desc_min_len > 0:
            return [("description_length", False, "Missing description.")]
        return [("description_length", True, "")]
    is_html = bool(cfg.desc_html_max_len) and bool(HTML_TAG_RE.search(desc))
    limit = cfg.desc_html_max_len if is_html else cfg.desc_max_len
    if len(desc) > limit:
        kind = "HTML" if is_html else "plain text"
        return [("description_length", False, f"Description is {len(desc)} chars, exceeds {limit}-char {kind} limit.")]
    if len(desc) < cfg.desc_min_len:
        return [("description_length", False, f"Description is {len(desc)} chars, below the {cfg.desc_min_len}-char minimum.")]
    return [("description_length", True, "")]


def check_feature_bullets(p, cfg, feature_bullets_data_present):
    if not cfg.has_feature_bullets:
        return []
    bullets_text = p.get("feature_bullets", "")
    if not feature_bullets_data_present:
        return [("feature_bullet_length", None,
                 "Feature bullets are not populated anywhere in this export — likely a field this Trustana "
                 "export doesn't map (Description may be doing double duty), rather than a per-product error.")]
    if not bullets_text:
        return [("feature_bullet_length", True, "")]
    bullets = [b.strip() for b in re.split(r"[\n;]", bullets_text) if b.strip()]
    too_long = [b for b in bullets if len(b) > cfg.bullet_max_len]
    if too_long:
        return [("feature_bullet_length", False,
                 f"{len(too_long)} of {len(bullets)} feature bullet(s) exceed {cfg.bullet_max_len} characters.")]
    return [("feature_bullet_length", True, "")]


def check_warranty_type(p, cfg, warranty_data_present):
    if not cfg.has_warranty_field:
        return []
    value = p.get("warranty_type", "")
    if not warranty_data_present:
        return [("warranty_type_check", None,
                 "Warranty Type is not populated anywhere in this export — confirm with the source system "
                 "whether warranty classification is tracked elsewhere before treating as a gap.")]
    if not value:
        return [("warranty_type_check", False, "Missing warranty type classification.")]
    low = value.lower()
    if not any(kw in low for kw in cfg.warranty_allowed_keywords):
        return [("warranty_type_check", False,
                 f"Warranty value '{value}' doesn't clearly map to manufacturer/seller/no-warranty — "
                 f"verify against noon's bulk-template enum before uploading.")]
    return [("warranty_type_check", True, "")]


def check_barcode_format(p, cfg):
    bc = p["barcode"]
    if not bc:
        return [("barcode_format", False, "Missing barcode/EAN.")]
    issues = []
    if not (cfg.barcode_min_len <= len(bc) <= cfg.barcode_max_len):
        issues.append(f"Barcode length {len(bc)} outside {cfg.barcode_min_len}-{cfg.barcode_max_len} char range.")
    if cfg.barcode_forbidden_chars_re.search(bc):
        issues.append("Barcode contains a forbidden character or space.")
    if issues:
        return [("barcode_format", False, " ".join(issues))]
    return [("barcode_format", True, "")]


def check_model_code(p, cfg, model_code_data_present):
    if not cfg.has_model_code:
        return []
    mc = p["model_code"]
    if not model_code_data_present:
        return [("model_code_format", None,
                 "Model code (Product Model) is not populated anywhere in this export — likely a field this "
                 "Trustana export doesn't map, rather than a per-product error. Confirm with the source system "
                 f"before treating as a QC gap; {cfg.display_name} requires it at upload time.")]
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


def check_image_count_and_url_format(p, cfg):
    imgs = p["images_list"]
    results = []
    if len(imgs) == 0:
        results.append(("image_count", False, "No images found."))
    elif cfg.image_max_count and len(imgs) > cfg.image_max_count:
        results.append(("image_count", False, f"{len(imgs)} images exceeds the {cfg.image_max_count}-image limit."))
    else:
        results.append(("image_count", True, ""))

    bad_urls = []
    for u in imgs:
        problems = []
        if not u.lower().startswith("https://"):
            problems.append("not https")
        if " " in u or "(" in u or ")" in u:
            problems.append("contains space/parenthesis")
        if not u.lower().split("?")[0].endswith(cfg.image_url_allowed_extensions):
            problems.append(f"no recognizable image extension ({'/'.join(e.lstrip('.') for e in cfg.image_url_allowed_extensions)})")
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


def check_category_depth_heuristic(p, cfg):
    cat = p["category"]
    if not cat:
        return []  # already hard-flagged by check_category_present
    parts = [x.strip() for x in cat.split("/") if x.strip()]
    if len(parts) < cfg.category_min_levels:
        return [("category_depth_heuristic", False,
                  f"Category path has only {len(parts)} level(s) ({cat}) — {cfg.display_name} expects at least "
                  f"{cfg.category_min_levels}. Verify against {cfg.display_name}'s live category tree (not available in this export).")]
    return [("category_depth_heuristic", True, "")]


TEXT_FIELDS_FOR_BANNED_WORDS = ["title", "description", "feature_bullets"]


def check_banned_words(p, cfg):
    findings = []
    for field_name in TEXT_FIELDS_FOR_BANNED_WORDS:
        text = p.get(field_name, "")
        if not text:
            continue
        low = text.lower()
        for label, terms in cfg.banned_word_groups:
            for term in terms:
                if term in low:
                    findings.append(f"possible {label} '{term}' in {field_name}")
        if EMAIL_RE.search(text):
            findings.append(f"email address detected in {field_name}")
        if PHONE_LIKE_RE.search(text) or LONG_DIGIT_RUN_RE.search(text):
            findings.append(f"phone-number-like digit sequence detected in {field_name}")
    if findings:
        return [("banned_words", False, "; ".join(findings))]
    return [("banned_words", True, "")]


def run_hard_checks(p, cfg, price_data_present, model_code_data_present,
                     feature_bullets_data_present, warranty_data_present):
    results = []
    results += check_title_length(p, cfg)
    results += check_description_length(p, cfg)
    results += check_barcode_format(p, cfg)
    results += check_model_code(p, cfg, model_code_data_present)
    results += check_feature_bullets(p, cfg, feature_bullets_data_present)
    results += check_image_count_and_url_format(p, cfg)
    results += check_price_consistency(p, price_data_present)
    results += check_category_present(p)
    results += check_banned_words(p, cfg)
    return results


# ---------------------------------------------------------------------------
# Per-product SOFT rule checks
# ---------------------------------------------------------------------------
def check_title_source(p):
    if p.get("title_source") == "Product Name (fallback)":
        return [("title_source_fallback", False,
                  f"The marketing title field is blank — used Product Name ('{p['title']}') as a stand-in for QC "
                  f"purposes. Product Name often carries brand/size/internal-code text not meant for a "
                  f"marketplace title — populate a proper marketing title rather than relying on this fallback.")]
    return [("title_source_fallback", True, "")]


def check_title_quality(p, cfg):
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
        issues.append(f"title contains brand name '{brand}' ({cfg.brand_in_title_note})")
    if title.count("!") > 1 or title.count("?") > 1:
        issues.append("excessive punctuation in title")
    last_token = title.strip().split()[-1] if title.strip() else ""
    stripped_token = re.sub(r"[^A-Za-z0-9]", "", last_token)
    if (len(stripped_token) >= 6 and stripped_token.isupper()
            and any(c.isdigit() for c in stripped_token) and any(c.isalpha() for c in stripped_token)):
        issues.append(f"title ends with a letters+digits code ('{last_token}') — likely a manufacturer model number or internal stock code carried over from Product Name; confirm this is meant to be customer-facing before keeping it")
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
    """Does the assigned category actually match what the product *is*, per
    its own name/title/description? Deliberately checks Product Name in
    addition to Title, since Title may be the Product Name fallback anyway,
    and Product Name sometimes carries more specific type wording than a
    cleaned-up marketing title."""
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


def run_soft_checks(p, cfg, warranty_data_present):
    detail = compute_content_detail(p)
    results = []
    results += check_title_source(p)
    results += check_title_quality(p, cfg)
    results += check_description_quality(p)
    results += check_content_relevance(p, detail=detail)
    results += check_title_description_relevance(p, detail=detail)
    results += check_category_depth_heuristic(p, cfg)
    results += check_warranty_type(p, cfg, warranty_data_present)
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
def verify_images_for_product(p, cfg, session, max_checks_remaining, network_ok):
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
            dpi = None
            try:
                img = Image.open(io.BytesIO(content))
                fmt = img.format
                w, h = img.size
                dpi_info = img.info.get("dpi")
                if dpi_info:
                    dpi = min(dpi_info)
            except Exception:
                pass

            problems = []
            if cfg.image_allowed_pil_formats and fmt not in cfg.image_allowed_pil_formats:
                problems.append(f"format {fmt or 'unknown'} not in {'/'.join(cfg.image_allowed_pil_formats)}")
            if not (cfg.image_min_size_bytes <= size <= cfg.image_max_size_bytes):
                problems.append(f"file size {size} bytes outside {cfg.image_min_size_bytes}-{cfg.image_max_size_bytes} bytes")
            if w and h:
                if cfg.image_min_w and w < cfg.image_min_w:
                    problems.append(f"width {w}px below minimum {cfg.image_min_w}px")
                if cfg.image_max_w and w > cfg.image_max_w:
                    problems.append(f"width {w}px above maximum {cfg.image_max_w}px")
                if cfg.image_min_h and h < cfg.image_min_h:
                    problems.append(f"height {h}px below minimum {cfg.image_min_h}px")
                if cfg.image_max_h and h > cfg.image_max_h:
                    problems.append(f"height {h}px above maximum {cfg.image_max_h}px")
                if cfg.image_min_aspect_ratio:
                    ratio = min(w, h) / max(w, h)
                    if ratio < cfg.image_min_aspect_ratio:
                        problems.append(f"aspect ratio {ratio:.2f} below minimum {cfg.image_min_aspect_ratio}")
            # PPI/DPI metadata is frequently absent even on perfectly good web
            # images — only flag when the metadata IS present and too low,
            # never penalize for missing metadata (that would be a false
            # positive on the majority of ordinary web-optimized images).
            if cfg.image_min_ppi and dpi is not None and dpi < cfg.image_min_ppi:
                problems.append(f"resolution {dpi} PPI below minimum {cfg.image_min_ppi} PPI")
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
def run_qc(csv_path, marketplace="trendyol", verify_images=False, max_image_checks=500):
    cfg = get_marketplace_config(marketplace)
    rule_catalogue = build_rule_catalogue(cfg)
    rule_severity = {r[0]: r[2] for r in rule_catalogue}
    rule_name = {r[0]: r[1] for r in rule_catalogue}

    products = load_products(csv_path)

    price_data_present = any(p["price"] or p["stock"] for p in products)
    model_code_data_present = any(p["model_code"] for p in products) if cfg.has_model_code else False
    feature_bullets_data_present = any(p.get("feature_bullets") for p in products) if cfg.has_feature_bullets else False
    warranty_data_present = any(p.get("warranty_type") for p in products) if cfg.has_warranty_field else False

    all_issues = []
    per_product_results = defaultdict(list)  # sku -> list of (rule_id, ok, message)

    for p in products:
        hard_results = run_hard_checks(p, cfg, price_data_present, model_code_data_present,
                                        feature_bullets_data_present, warranty_data_present)
        soft_results = run_soft_checks(p, cfg, warranty_data_present)
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
                results, network_ok, used = verify_images_for_product(p, cfg, session, image_checks_remaining, network_ok)
                image_checks_remaining -= used
                for rule_id, ok, msg in results:
                    per_product_results[p["sku"]].append((rule_id, ok, msg))
            if not network_ok:
                for p in products:
                    if p["sku"] not in processed_skus:
                        per_product_results[p["sku"]].append((
                            "image_technical_spec", None,
                            "Skipped — image technical verification was disabled for the rest of the run "
                            "after an earlier download failure (see Summary sheet)."
                        ))

    info_notes = []
    for p in products:
        for rule_id, ok, msg in per_product_results[p["sku"]]:
            if ok is False:
                all_issues.append({
                    "sku": p["sku"], "product_name": p["product_name"], "title": p["title"],
                    "brand": p["brand"], "category": p["category"],
                    "rule_id": rule_id, "rule_name": rule_name.get(rule_id, rule_id),
                    "severity": rule_severity.get(rule_id, "SOFT"), "message": msg,
                })
            elif ok is None and msg:
                info_notes.append({
                    "sku": p["sku"], "rule_id": rule_id, "message": msg,
                })

    return {
        "marketplace": cfg.key,
        "marketplace_display_name": cfg.display_name,
        "rule_catalogue": rule_catalogue,
        "products": products,
        "per_product_results": per_product_results,
        "issues": all_issues,
        "info_notes": info_notes,
        "price_data_present": price_data_present,
        "model_code_data_present": model_code_data_present,
        "feature_bullets_data_present": feature_bullets_data_present,
        "warranty_data_present": warranty_data_present,
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
    rule_catalogue = qc_result["rule_catalogue"]
    rule_name = {r[0]: r[1] for r in rule_catalogue}
    marketplace_name = qc_result["marketplace_display_name"]

    wb = Workbook()

    # --- Summary sheet ---
    ws = wb.active
    ws.title = "Summary"
    ws.append([f"{marketplace_name} Marketplace QC Report"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([f"Marketplace: {marketplace_name}"])
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
    if not qc_result["model_code_data_present"] and qc_result.get("marketplace") == "trendyol":
        ws.append(["NOTE: Product Model (model code) column is empty for every product in this export — model_code_format was flagged as informational rather than a per-product fail. Confirm with the source system whether this field is mapped elsewhere."])
    if qc_result.get("marketplace") == "noon" and not qc_result.get("feature_bullets_data_present"):
        ws.append(["NOTE: Feature bullets are not populated anywhere in this export — feature_bullet_length was flagged as informational rather than a per-product fail."])
    if qc_result.get("marketplace") == "noon" and not qc_result.get("warranty_data_present"):
        ws.append(["NOTE: Warranty Type is not populated anywhere in this export — warranty_type_check was flagged as informational rather than a per-product fail."])
    if qc_result["verify_images_requested"]:
        if qc_result["image_network_ok"]:
            ws.append(["Image technical verification: images were downloaded and checked against format/size/resolution spec."])
        else:
            ws.append(["Image technical verification: REQUESTED but could not complete — outbound network access to the image host was not available in this run. Image checks below are limited to URL format/count only. Re-run with network access to this host for full verification."])
    ws.append([])

    ws.append(["Rule", "Severity", "Products failing", "Products passing", "Not evaluated / N/A"])
    header_row = ws.max_row
    rule_ids_in_order = [r[0] for r in rule_catalogue]
    for rule_id in rule_ids_in_order:
        name = rule_name[rule_id]
        sev = {r[0]: r[2] for r in rule_catalogue}[rule_id]
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
    ws4.append([f"Rules for: {marketplace_name}"])
    ws4["A1"].font = Font(bold=True, size=12)
    ws4.append([])
    ws4.append(["Rule ID", "Name", "Severity", "Source section(s)", "Condition"])
    header_row4 = ws4.max_row
    for rule_id, name, sev, source, cond in rule_catalogue:
        ws4.append([rule_id, name, sev, source, cond])
    _style_header(ws4, row=header_row4)
    for row in ws4.iter_rows(min_row=header_row4 + 1, max_row=ws4.max_row):
        fill = HARD_FILL if row[2].value == HARD else SOFT_FILL
        for cell in row:
            cell.fill = fill
    ws4.freeze_panes = f"A{header_row4 + 1}"
    _autosize(ws4, [24, 34, 10, 18, 80])

    wb.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="Trustana AI Content Verifier for MP — checks a Trustana export against a marketplace's listing rules.")
    parser.add_argument("csv_path", help="Path to the Trustana product export CSV")
    parser.add_argument("output_path", help="Path to write the Excel QC report")
    parser.add_argument("--marketplace", required=True, choices=list(MARKETPLACES.keys()), help="Which marketplace's rules to check against")
    parser.add_argument("--verify-images", action="store_true", help="Download images and verify format/size/resolution")
    parser.add_argument("--max-image-checks", type=int, default=500, help="Cap on number of images to download when --verify-images is set")
    args = parser.parse_args()

    qc_result = run_qc(args.csv_path, marketplace=args.marketplace, verify_images=args.verify_images, max_image_checks=args.max_image_checks)
    build_report(qc_result, args.output_path, args.csv_path)
    n_hard = sum(1 for i in qc_result["issues"] if i["severity"] == HARD)
    n_soft = sum(1 for i in qc_result["issues"] if i["severity"] == SOFT)
    print(f"[{qc_result['marketplace_display_name']}] Checked {len(qc_result['products'])} products: {n_hard} hard issues, {n_soft} soft flags.")
    print(f"Report written to {args.output_path}")


if __name__ == "__main__":
    main()

"""
Trustana AI Content Verifier for MP — Streamlit web app.

Wraps qc_engine.py (the underlying multi-marketplace QC engine, unchanged)
in a simple upload-and-download web page, meant to be deployed on Streamlit
Community Cloud so it can be shared as a link — no local Python setup
required for whoever uses it.

Deploy: push this folder to a GitHub repo, connect it at
share.streamlit.io, set the main file path to `app.py`. See DEPLOY.md for
the full walkthrough.
"""

import os
import tempfile

import pandas as pd
import streamlit as st

from qc_engine import (
    HARD,
    MARKETPLACES,
    SOFT,
    build_report,
    build_rule_catalogue,
    clean_skus,
    export_noon_template,
    export_trendyol_template,
    get_marketplace_config,
    preview_export_breakdown,
    run_qc,
)

st.set_page_config(page_title="Trustana AI Content Verifier for MP", page_icon="✅", layout="centered")

st.title("Trustana AI Content Verifier for MP")
st.caption(
    "Upload a Trustana product export (CSV), pick a marketplace, and check it against "
    "that marketplace's listing rules before it goes live."
)

marketplace_options = {cfg.display_name: key for key, cfg in MARKETPLACES.items()}
marketplace_label = st.selectbox(
    "Marketplace",
    options=list(marketplace_options.keys()),
    help="Choose which marketplace's listing rules to check this export against.",
)
marketplace_key = marketplace_options[marketplace_label]
cfg = get_marketplace_config(marketplace_key)

# Rule catalogue is built per-marketplace (thresholds differ), so fetch it fresh here for display.
rule_catalogue = build_rule_catalogue(cfg)

with st.expander("What does this check?"):
    st.markdown(
        f"""
For **{cfg.display_name}**, this tool runs **{len(rule_catalogue)} rules** against every product in your export:

- **HARD rules** ({sum(1 for r in rule_catalogue if r[2] == HARD)}) — things that would get a listing
  rejected outright: title/description length, barcode format, image count, banned/restricted
  content, duplicate barcodes, and more.
- **SOFT rules** ({sum(1 for r in rule_catalogue if r[2] == SOFT)}) — quality signals that don't block
  approval but hurt visibility or indicate a possible mistake: title quality, whether the
  category actually matches the product, whether the title and description are about the
  same product, and more.

See the **Rules reference** section at the bottom of this page for the full list, or check the
"Rules Reference" sheet in the downloaded report. Rules come from `{cfg.rules_doc}`.
"""
    )

uploaded = st.file_uploader("Trustana product export (CSV)", type=["csv"])

col1, col2 = st.columns([1, 1])
with col1:
    verify_images = st.checkbox(
        "Also verify real image files",
        value=False,
        help=(
            "Downloads every product image and checks real format/file size/resolution "
            "against the selected marketplace's spec. Needs this server to have outbound "
            "internet access to wherever the images are hosted, and is much slower on large "
            "exports. Without this, images are only checked for a valid URL/count, not "
            "downloaded."
        ),
    )
with col2:
    max_image_checks = st.number_input(
        "Max images to download",
        min_value=1,
        max_value=5000,
        value=500,
        step=50,
        disabled=not verify_images,
    )

run_clicked = st.button("Run QC", type="primary", disabled=uploaded is None)

if run_clicked and uploaded is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, uploaded.name)
        with open(input_path, "wb") as f:
            f.write(uploaded.getbuffer())
        output_path = os.path.join(tmpdir, "QC_Report.xlsx")

        with st.spinner(f"Running {cfg.display_name} QC checks..."):
            try:
                qc_result = run_qc(
                    input_path,
                    marketplace=marketplace_key,
                    verify_images=verify_images,
                    max_image_checks=int(max_image_checks),
                )
                build_report(qc_result, output_path, uploaded.name)
            except Exception as e:
                st.error(f"Couldn't process this file: {e}")
                st.stop()

        with open(output_path, "rb") as f:
            report_bytes = f.read()

    # Cached in session_state so the Stage 2 export button below (a separate
    # widget interaction) can reuse this run's results without needing the
    # user to click "Run QC" again — Streamlit reruns the whole script on
    # every widget click, and a plain local variable wouldn't survive that.
    st.session_state["qc_result"] = qc_result
    st.session_state["report_bytes"] = report_bytes
    st.session_state["source_name"] = uploaded.name
    st.session_state["marketplace_key"] = marketplace_key
    st.session_state.pop("export_bytes", None)
    st.session_state.pop("export_summary", None)
    st.session_state.pop("export_preview", None)

qc_result = st.session_state.get("qc_result")
if qc_result is not None and st.session_state.get("marketplace_key") == marketplace_key:
    products = qc_result["products"]
    issues = qc_result["issues"]
    source_name = st.session_state["source_name"]
    base_name = os.path.splitext(source_name)[0]
    n_hard_products = len({i["sku"] for i in issues if i["severity"] == HARD})
    n_soft_products = len({i["sku"] for i in issues if i["severity"] == SOFT})

    st.success(f"Checked {len(products)} products against {cfg.display_name} rules.")

    m1, m2, m3 = st.columns(3)
    m1.metric("Products checked", len(products))
    m2.metric("Rejection-risk products", n_hard_products)
    m3.metric("Products with quality flags", n_soft_products)

    if not qc_result["price_data_present"]:
        st.info("Price/Stock columns are empty in this export — price_consistency was skipped file-wide.")
    if not qc_result.get("model_code_data_present", True) and cfg.has_model_code:
        st.info("Product Model (model code) is empty across the whole export — treated as a data gap, not a per-product failure.")
    if cfg.has_feature_bullets and not qc_result.get("feature_bullets_data_present", True):
        st.info("Feature bullets are empty across the whole export — treated as a data gap, not a per-product failure.")
    if cfg.has_warranty_field and not qc_result.get("warranty_data_present", True):
        st.info("Warranty type is empty across the whole export — treated as a data gap, not a per-product failure.")
    if cfg.has_classification_directory and not qc_result.get("classification_dir_loaded"):
        st.warning("noon's Classification Directory reference couldn't be loaded — category validity/relevance checks were skipped as unverified.")
    if qc_result.get("verify_images_requested") and qc_result.get("image_network_ok") is False:
        st.warning(
            "Image technical verification was requested but couldn't complete — this server "
            "doesn't have outbound network access to the image host. Image checks for this run "
            "are limited to URL format/count only; the report notes this too."
        )

    st.download_button(
        label="Download QC Report (Excel)",
        data=st.session_state["report_bytes"],
        file_name=f"{base_name}_{marketplace_key}_QC_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.divider()
    st.subheader("Top issues")
    if issues:
        df = pd.DataFrame(
            [
                {
                    "SKU": i["sku"],
                    "Product": i["product_name"],
                    "Rule": i["rule_name"],
                    "Severity": i["severity"],
                    "Message": i["message"],
                }
                for i in issues
            ]
        ).sort_values(["Severity", "SKU"], ascending=[True, True])
        st.dataframe(df, width="stretch", height=400)
    else:
        st.write("No issues found.")

    st.divider()
    st.subheader("Stage 2: export to marketplace upload template")
    allow_soft_issues = st.checkbox(
        "Also include products with quality (SOFT) flags — only block on rejection-risk (HARD) issues",
        value=False,
        help=(
            "Off (default): only products with zero issues of any kind are exported — the strictest reading "
            "of \"fully QC-verified.\" On: products with SOFT quality flags (e.g. a brand name in the title) "
            "are still exported as long as they have no HARD rejection-risk issues. On a real Trendyol export "
            "tested with this tool, the strict setting qualified 0 of 151 products (nearly every title included "
            "the brand name) while the relaxed setting qualified 118 — so try turning this on if your export "
            "comes back empty and that feels too strict for your catalog."
        ),
    )
    n_eligible = len(clean_skus(qc_result, allow_soft_issues=allow_soft_issues))
    st.caption(f"{n_eligible} of {len(products)} products currently meet this export's threshold.")

    if st.button("Generate export summary", disabled=n_eligible == 0):
        st.session_state["export_preview"] = preview_export_breakdown(qc_result, allow_soft_issues=allow_soft_issues)
        st.session_state.pop("export_bytes", None)
        st.session_state.pop("export_summary", None)

    if "export_preview" in st.session_state:
        preview = st.session_state["export_preview"]
        st.markdown("**QC summary for manual checking, before filling any template:**")
        if marketplace_key == "noon":
            st.write(f"- **{preview['eligible']}** product(s) are ready to be written into noon's `template_data` sheet.")
        else:
            for sheet_name, count in sorted(preview["by_sheet"].items()):
                st.write(f"- **{count}** product(s) → `{sheet_name}`")
            if preview["unmapped"]:
                st.write(
                    f"- **{preview['unmapped']}** product(s) are in a category with no known Trendyol sheet mapping at "
                    "all — these will land in 'Leftover - Needs Template' regardless of which template you upload."
                )
        st.write(f"- **{preview['excluded']}** product(s) are excluded — see the QC report's Issues sheet for why.")
        st.write("Review this against your catalog before continuing. Once it looks right:")

        st.info(
            f"**Download the current upload template from {cfg.display_name}'s seller center** — not a saved copy — "
            "and upload it below. This matters especially for Trendyol, where the template you download only includes "
            "sheets for the categories you selected when requesting it, so a template downloaded for a different "
            "category mix won't match this catalog. noon's template is more stable but can still change, so it's "
            "safest to always use a freshly-downloaded copy."
        )

        template_upload = st.file_uploader(
            f"Current {cfg.display_name} upload template (.xlsx)",
            type=["xlsx"],
            key=f"template_upload_{marketplace_key}",
        )

        if st.button("Fill template & download", disabled=template_upload is None, type="primary"):
            with tempfile.TemporaryDirectory() as tmpdir:
                template_path = os.path.join(tmpdir, template_upload.name)
                with open(template_path, "wb") as f:
                    f.write(template_upload.getbuffer())
                export_path = os.path.join(tmpdir, "export.xlsx")
                try:
                    if marketplace_key == "noon":
                        n_exp, n_exc = export_noon_template(qc_result, export_path, template_path=template_path, allow_soft_issues=allow_soft_issues)
                        summary_text = f"Exported {n_exp} product(s) into noon's template_data sheet ({n_exc} excluded — see 'Excluded From Export')."
                    else:
                        summary = export_trendyol_template(qc_result, export_path, template_path=template_path, allow_soft_issues=allow_soft_issues)
                        summary_text = (
                            f"Exported {summary['exported_to_sheets']} product(s) into their matching Trendyol category sheet(s), "
                            f"{summary['leftover']} into 'Leftover - Needs Template' (no matching sheet in the template you uploaded), "
                            f"{summary['excluded']} excluded."
                        )
                        if summary["attribute_warnings"]:
                            summary_text += f" {summary['attribute_warnings']} attribute value(s) flagged in 'Attribute Value Warnings' — double-check those before uploading."
                    with open(export_path, "rb") as f:
                        st.session_state["export_bytes"] = f.read()
                    st.session_state["export_summary"] = summary_text
                except Exception as e:
                    st.error(f"Couldn't generate the export: {e}. Make sure the file you uploaded is really this marketplace's upload template.")

    if "export_bytes" in st.session_state:
        st.success(st.session_state["export_summary"])
        st.download_button(
            label=f"Download {cfg.display_name} upload file",
            data=st.session_state["export_bytes"],
            file_name=f"{base_name}_{marketplace_key}_marketplace_upload.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

st.divider()
with st.expander("Rules reference"):
    rules_df = pd.DataFrame(
        [
            {"Rule ID": r[0], "Name": r[1], "Severity": r[2], "Source": r[3], "Condition": r[4]}
            for r in rule_catalogue
        ]
    )
    st.dataframe(rules_df, width="stretch", height=400)

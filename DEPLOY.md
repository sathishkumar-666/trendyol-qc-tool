# Deploying Trustana AI Content Verifier for MP on Streamlit Community Cloud

This is for whoever is deploying the app (i.e. you) — not for the teammates
who'll just open the link and use it. Streamlit Community Cloud is free and
gives you a public URL like `https://your-app-name.streamlit.app` in a few
minutes.

## What's in this folder

- `app.py` — the web app (pick a marketplace → upload CSV → run checks →
  download report). This is the file Streamlit Cloud runs.
- `qc_engine.py` — the actual multi-marketplace QC engine. `app.py` imports
  from it directly; don't rename or move it out of this folder.
- `Trendyol_MP_QC_Rules.md` / `Noon_MP_QC_Rules.md` — the source rules docs
  each marketplace's checks are built from. Not imported by code, but keep
  them in the repo so the "Rules Reference" sheet's source citations make
  sense to whoever reads the report.
- `requirements.txt` — everything needed, including `streamlit` and
  `pandas`. Streamlit Cloud installs from this automatically.
- `run_qc.sh` / `run_qc.bat` / the rest of `README.md` — the local
  run-it-yourself option from before. You can keep or drop these; they
  don't affect the web deployment either way.

## Steps

1. **Get this folder into a GitHub repo.** If you don't already have one:
   - Create a new repo on GitHub (private is fine — Streamlit Community
     Cloud can deploy from private repos once you connect your GitHub
     account).
   - Push these files to it:
     ```
     cd trendyol_qc_tool
     git init
     git add app.py qc_engine.py Trendyol_MP_QC_Rules.md Noon_MP_QC_Rules.md requirements.txt README.md run_qc.sh run_qc.bat
     git commit -m "Trustana AI Content Verifier for MP"
     git branch -M main
     git remote add origin https://github.com/<your-org>/<your-repo>.git
     git push -u origin main
     ```
   - **Already have this repo deployed from an earlier, Trendyol-only
     version?** See "Updating an existing deployment" below instead —
     don't repeat the steps above.

2. **Go to [share.streamlit.io](https://share.streamlit.io)** and sign in
   with the GitHub account that has access to that repo.

3. Click **"New app"**, then:
   - Repository: pick the repo you just pushed.
   - Branch: `main`.
   - Main file path: `app.py`.
   - (Optional) pick a custom app URL slug, e.g. `trustana-mp-qc` →
     `https://trustana-mp-qc.streamlit.app`.

4. Click **Deploy**. First deploy takes 1-3 minutes while it installs
   `requirements.txt`. After that you'll have your shareable link.

5. **Share the link** with your team. That's it — no install, no Python,
   just open the URL, pick a marketplace, upload a CSV, download the
   report.

## Updating an existing deployment (e.g. adding noon / the rename)

If you already deployed the earlier Trendyol-only version, you don't need
to create a new Streamlit app or repo — just push the new files to the
same repo/branch it's already watching:

```
cd trendyol_qc_tool
git add app.py qc_engine.py Noon_MP_QC_Rules.md Trendyol_MP_QC_Rules.md README.md DEPLOY.md run_qc.sh run_qc.bat requirements.txt
git rm trendyol_qc_tool.py
git commit -m "Add noon marketplace, marketplace selector, rename to Trustana AI Content Verifier for MP"
git push
```

Streamlit Community Cloud will pick up the push and redeploy automatically
within about a minute — you don't need to touch the share.streamlit.io
dashboard. The app's URL and any viewer restrictions you set stay the same;
only the on-page title changes to "Trustana AI Content Verifier for MP" and
a marketplace dropdown appears.

(The `git rm trendyol_qc_tool.py` removes the old single-marketplace engine
file, which `app.py` no longer imports from — keeping it around would just
be dead code sitting in the repo.)

## If the product data is sensitive

By default a Streamlit Community Cloud app is reachable by anyone with the
link — it's not indexed/searchable, but it's not access-controlled either.
If your product catalog shouldn't be uploadable by just anyone with the
URL, go to the app's **Settings → Sharing** in the Streamlit Cloud
dashboard and restrict viewers to specific email addresses or your Google
Workspace domain. This doesn't cost anything extra.

The app itself doesn't persist uploaded files or reports anywhere — each
run writes to a temporary directory that's discarded as soon as the
session ends, and nothing is logged to a database. Data only lives in that
one browser session's memory while it's being processed.

## Updating the tool later

Whenever you (or Claude, in a future session) change `qc_engine.py` or
`app.py` — including adding another marketplace — just commit and push to
the same branch:
```
git add -A
git commit -m "describe the change"
git push
```
Streamlit Community Cloud watches the repo and redeploys automatically
within about a minute — no need to touch the dashboard again.

## One thing to expect

Free-tier Streamlit apps "sleep" after a period of no visitors and take
~20-30 seconds to wake up on the next visit (you'll see a "waking up" screen
briefly). This is normal and not something to fix — just a heads-up so
nobody thinks it's broken the first time they open the link after a while.

## Image verification (`--verify-images` equivalent) on this deployment

The web app's "Also verify real image files" checkbox will actually work
here, unlike testing inside this sandbox — Streamlit Community Cloud has
normal outbound internet access, so it can reach the S3-hosted product
images. It will just be slower for large exports; the "Max images to
download" field caps how many it fetches per run if that becomes a problem.
This applies to both marketplaces — the technical spec it checks against
(format, size, resolution) automatically switches based on the marketplace
selected in the dropdown.

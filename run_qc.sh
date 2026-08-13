#!/usr/bin/env bash
# Trustana AI Content Verifier for MP — convenience wrapper for Mac/Linux.
#
# Usage:
#   ./run_qc.sh <trustana_export.csv> [trendyol|noon] [--verify-images]
#                                      [--export-template OUTPUT_XLSX]
#                                      [--template-file CURRENT_TEMPLATE_XLSX] [--allow-soft-issues]
#
# Marketplace defaults to "trendyol" if omitted. The report is saved next to
# the input file as <name>_<marketplace>_QC_Report.xlsx. First run sets up a
# local Python environment automatically (one-time, ~30 seconds); every run
# after that is instant.
#
# --export-template OUTPUT_XLSX     Also fill the marketplace's own upload
#                                    template with every QC-clean product.
# --template-file CURRENT_TEMPLATE  The upload template to fill — download the
#                                    current one from the marketplace's seller
#                                    center first (Trendyol's varies by which
#                                    categories you requested it for). Omit to
#                                    fall back to the bundled copy (may be stale).
# --allow-soft-issues                Relax "QC-clean" for --export-template from
#                                    zero issues to zero rejection-risk (HARD)
#                                    issues. See README's "Stage 2" section.
#
# Examples:
#   ./run_qc.sh Products_export.csv noon --export-template noon_upload.xlsx --template-file current_noon_template.xlsx
#   ./run_qc.sh Products_export.csv trendyol --export-template trendyol_upload.xlsx --template-file current_trendyol_template.xlsx --allow-soft-issues

set -e
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [ -z "$1" ]; then
  echo "Usage: ./run_qc.sh <trustana_export.csv> [trendyol|noon] [--verify-images]"
  echo "                                          [--export-template OUTPUT_XLSX] [--allow-soft-issues]"
  echo "Marketplace defaults to 'trendyol' if omitted."
  echo "The report will be saved next to the input file as <name>_<marketplace>_QC_Report.xlsx"
  exit 1
fi

INPUT="$1"
shift

MARKETPLACE="trendyol"
if [ "$1" = "trendyol" ] || [ "$1" = "noon" ]; then
  MARKETPLACE="$1"
  shift
fi

if [ ! -f "$INPUT" ]; then
  echo "Error: can't find file '$INPUT'"
  exit 1
fi

BASENAME=$(basename "$INPUT")
BASENAME="${BASENAME%.*}"
INPUT_DIR=$(cd "$(dirname "$INPUT")" && pwd)
OUTPUT="${INPUT_DIR}/${BASENAME}_${MARKETPLACE}_QC_Report.xlsx"

if [ ! -d "$DIR/.venv" ]; then
  echo "First run: setting up Python environment (one-time, ~30 seconds)..."
  python3 -m venv "$DIR/.venv"
  "$DIR/.venv/bin/pip" install --quiet --upgrade pip
  "$DIR/.venv/bin/pip" install --quiet -r "$DIR/requirements.txt"
fi

echo "Marketplace: $MARKETPLACE"
"$DIR/.venv/bin/python" "$DIR/qc_engine.py" "$INPUT" "$OUTPUT" --marketplace "$MARKETPLACE" "$@"
echo ""
echo "Done. Report saved to: $OUTPUT"

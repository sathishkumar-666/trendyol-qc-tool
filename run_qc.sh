#!/usr/bin/env bash
# Trendyol QC tool — convenience wrapper for Mac/Linux.
#
# Usage:
#   ./run_qc.sh <trustana_export.csv> [--verify-images]
#
# The report is saved next to the input file as <name>_QC_Report.xlsx.
# First run sets up a local Python environment automatically (one-time,
# ~30 seconds); every run after that is instant.

set -e
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [ -z "$1" ]; then
  echo "Usage: ./run_qc.sh <trustana_export.csv> [--verify-images]"
  echo "The report will be saved next to the input file as <name>_QC_Report.xlsx"
  exit 1
fi

INPUT="$1"
shift

if [ ! -f "$INPUT" ]; then
  echo "Error: can't find file '$INPUT'"
  exit 1
fi

BASENAME=$(basename "$INPUT")
BASENAME="${BASENAME%.*}"
INPUT_DIR=$(cd "$(dirname "$INPUT")" && pwd)
OUTPUT="${INPUT_DIR}/${BASENAME}_QC_Report.xlsx"

if [ ! -d "$DIR/.venv" ]; then
  echo "First run: setting up Python environment (one-time, ~30 seconds)..."
  python3 -m venv "$DIR/.venv"
  "$DIR/.venv/bin/pip" install --quiet --upgrade pip
  "$DIR/.venv/bin/pip" install --quiet -r "$DIR/requirements.txt"
fi

"$DIR/.venv/bin/python" "$DIR/trendyol_qc_tool.py" "$INPUT" "$OUTPUT" "$@"
echo ""
echo "Done. Report saved to: $OUTPUT"

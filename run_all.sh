#!/usr/bin/env bash
set -euo pipefail

# Runbook: create venv, install deps, fetch FIRST reports, and convert to CSV
# This script always creates/overwrites the local `.venv` in the project root.

echo "Creating virtualenv and activating (.venv)"
python3 -m venv .venv && source .venv/bin/activate

echo "Upgrading pip and installing Python requirements"
pip install -U pip
pip install -r requirements.txt

# Install Playwright browsers (only needed the first time)
echo "Installing Playwright browsers (no-op if already installed)"
playwright install --with-deps || true

echo "Running scraper (headless)"
python fetch-first-dui-data.py

echo "If you want to run with a visible browser for debugging, run this instead:"
echo "  SHOW_BROWSER=1 python fetch-first-dui-data.py"

echo "Converting downloaded .xlsx files to simplified CSVs"
python excel-sheet-to-csv.py

echo "Done. Output CSVs are in ./output"

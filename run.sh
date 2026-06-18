#!/usr/bin/env bash
# macOS / Linux dev launcher — starts the Streamlit dev server.
# Mirrors `streamlit run app/app.py`; no packaging involved.
set -euo pipefail
cd "$(dirname "$0")"

exec streamlit run app/app.py

#!/usr/bin/env bash
# macOS double-click launcher — Finder opens .command files in Terminal.
# Delegates to run.sh (Streamlit dev server) and opens the browser.
cd "$(dirname "$0")"

# Finder-launched shells have a minimal PATH; pull in common locations.
export PATH="/opt/homebrew/bin:/usr/local/bin:$HOME/.local/bin:$PATH"

# Use the project's virtualenv (streamlit is installed there).
if [ -f venv-mac/bin/activate ]; then
  # shellcheck disable=SC1091
  source venv-mac/bin/activate
fi

( sleep 3 && open "http://localhost:8501" ) &

./run.sh
status=$?

echo
echo "Server stopped (exit $status). Press Return to close."
read -r _

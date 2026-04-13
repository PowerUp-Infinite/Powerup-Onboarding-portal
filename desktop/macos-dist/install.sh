#!/bin/bash
# ─────────────────────────────────────────────────────────────
# PowerUp Portal — first-time setup (macOS)
#
# Run this once after extracting the folder. It:
#   1. Verifies Python 3.11+ is installed
#   2. Creates a virtual environment in .venv/
#   3. Installs all dependencies
# ─────────────────────────────────────────────────────────────
set -e
cd "$(dirname "$0")"

echo
echo "========================================"
echo "  PowerUp Portal - First-time setup"
echo "========================================"
echo

# --- 1. Find a usable Python ---
echo "[1/3] Looking for Python 3.11 or later..."
PY=""
for cand in python3.13 python3.12 python3.11 python3; do
    if command -v "$cand" >/dev/null 2>&1; then
        # Verify version >= 3.11
        if "$cand" -c "import sys; sys.exit(0 if sys.version_info >= (3,11) else 1)" 2>/dev/null; then
            PY="$cand"
            break
        fi
    fi
done

if [[ -z "$PY" ]]; then
    echo
    echo "[ERROR] Python 3.11 or later was not found."
    echo
    echo "Install it from one of these:"
    echo "  - https://www.python.org/downloads/macos/   (download the universal2 .pkg installer)"
    echo "  - or via Homebrew:   brew install python@3.11"
    echo
    echo "Then re-run this script."
    exit 1
fi

PYV=$("$PY" --version 2>&1)
echo "      Found $PYV ($PY)."

# --- 2. Create venv ---
if [[ -f ".venv/bin/python3" ]]; then
    echo "[2/3] Virtual environment already exists, reusing it."
else
    echo "[2/3] Creating virtual environment in .venv/ ..."
    "$PY" -m venv .venv
fi

# --- 3. Install dependencies ---
echo "[3/3] Installing dependencies (takes ~1-2 min on first run)..."
.venv/bin/python -m pip install --upgrade pip --quiet
.venv/bin/python -m pip install -r requirements.txt --quiet

# Make the launcher executable
chmod +x ./run-PowerUp-Portal.command 2>/dev/null || true

echo
echo "========================================"
echo "  Setup complete!"
echo "========================================"
echo
echo "Double-click   run-PowerUp-Portal.command   to start the app."
echo "(First time: right-click -> Open -> Open, to bypass Gatekeeper.)"
echo

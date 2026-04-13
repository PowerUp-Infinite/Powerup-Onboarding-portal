#!/bin/bash
# =============================================================
# PowerUp Portal -- first-time setup [macOS]
#
#  1. If Python 3.9+ is missing, runs the bundled .pkg installer
#     (system Python on macOS Ventura+ is usually fine; we only
#     install when needed).
#  2. Creates a virtual environment in .venv/
#  3. Installs all dependencies into the venv
# =============================================================
set -e
cd "$(dirname "$0")"

echo
echo "========================================"
echo "  PowerUp Portal - First-time setup"
echo "========================================"
echo

# --- 1. Find or install Python ---
echo "[1/4] Looking for Python 3.9 or later..."
PY=""
for cand in python3.13 python3.12 python3.11 python3.10 python3.9 python3; do
    if command -v "$cand" >/dev/null 2>&1; then
        if "$cand" -c "import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)" 2>/dev/null; then
            PY="$cand"
            break
        fi
    fi
done

if [[ -z "$PY" ]]; then
    echo "       No suitable Python found. Installing Python 3.11.9 from bundled installer..."
    if [[ ! -f "python-3.11.9-macos11.pkg" ]]; then
        echo
        echo "[ERROR] python-3.11.9-macos11.pkg is missing from this folder."
        echo "        Re-download PowerUp-Portal-Mac.zip and try again."
        exit 1
    fi
    echo "       This will prompt for your Mac password (admin install)."
    sudo installer -pkg python-3.11.9-macos11.pkg -target / || {
        echo
        echo "[ERROR] Python install was cancelled or failed."
        echo "        You can also install manually by double-clicking"
        echo "        python-3.11.9-macos11.pkg in this folder."
        exit 1
    }
    # Refresh PATH so the freshly-installed python3.11 is findable.
    export PATH="/Library/Frameworks/Python.framework/Versions/3.11/bin:$PATH"
    PY="python3.11"
fi

PYV=$("$PY" --version 2>&1)
echo "       Found $PYV ($PY)."

# --- 2. Create venv ---
if [[ -f ".venv/bin/python3" ]]; then
    echo "[2/4] Virtual environment already exists, reusing it."
else
    echo "[2/4] Creating virtual environment in .venv/ ..."
    "$PY" -m venv .venv
fi

# --- 3. Upgrade pip ---
echo "[3/4] Upgrading pip..."
.venv/bin/python -m pip install --upgrade pip --quiet

# --- 4. Install dependencies ---
echo "[4/4] Installing dependencies. Takes ~1-2 min on first run..."
.venv/bin/python -m pip install -r requirements.txt --quiet

# Make the launcher executable (zip extraction sometimes strips +x)
chmod +x ./run-PowerUp-Portal.command 2>/dev/null || true

# Strip macOS quarantine attribute from the launcher and the rest of the
# folder so daily use does NOT hit the "Apple could not verify..." dialog
# again. We can do this safely now -- the user has already trusted this
# install by running it, so trusting the sibling files is fine.
echo "       Removing Gatekeeper quarantine on bundled files..."
xattr -dr com.apple.quarantine . 2>/dev/null || true

echo
echo "========================================"
echo "  Setup complete!"
echo "========================================"
echo
echo "Daily use: double-click   run-PowerUp-Portal.command"
echo
echo "(No more Gatekeeper dialogs -- this script just removed the"
echo " quarantine flag from every file in this folder.)"
echo

#!/bin/bash
# ─────────────────────────────────────────────────────────────
# PowerUp Portal — macOS build script.
#
# Produces desktop/dist/PowerUp-Portal.app (a .app bundle, ~250 MB).
# Run this ON A MAC (not Windows — .app can only be built on macOS):
#     cd desktop
#     chmod +x build-mac.sh
#     ./build-mac.sh
# ─────────────────────────────────────────────────────────────

set -e
cd "$(dirname "$0")"

echo
echo "==== PowerUp Portal — macOS build ===="
echo

# 1. Ensure Python 3.11+ is available.
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] python3 is not installed."
    echo "        Install via:   brew install python@3.11"
    exit 1
fi

PYV=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
if [[ "$PYV" < "3.11" ]]; then
    echo "[ERROR] Python $PYV found. Need 3.11+."
    echo "        brew install python@3.11"
    exit 1
fi

# 2. Venv so we don't pollute the user's global site-packages.
if [[ ! -f ".venv/bin/python" ]]; then
    echo "[1/4] Creating virtual environment..."
    python3 -m venv .venv
fi

echo "[2/4] Installing dependencies..."
source .venv/bin/activate
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet

echo "[3/5] Verifying credentials.json is bundled..."
if [[ ! -f "resources/credentials.json" ]]; then
    echo "[ERROR] desktop/resources/credentials.json is missing."
    echo "        Copy your service account JSON to that path before building."
    exit 1
fi

echo "[4/5] Copying portal/.env into bundled resources/..."
if [[ ! -f "../portal/.env" ]]; then
    echo "[ERROR] ../portal/.env does not exist."
    echo "        The desktop app reads config values from portal/.env at build time."
    exit 1
fi
cp -f "../portal/.env" "resources/.env"

echo "[5/5] Running PyInstaller..."
pyinstaller PowerUp-Portal.spec --clean --noconfirm

echo
echo "==== Build complete ===="
echo "Output: $(pwd)/dist/PowerUp-Portal.app"
echo
echo "Zip it for cleaner distribution:"
echo "    cd dist && zip -ry PowerUp-Portal-Mac.zip PowerUp-Portal.app"
echo
echo "Next steps:"
echo "  1. Test by double-clicking dist/PowerUp-Portal.app"
echo "     (First open: right-click → Open → 'Open' in the prompt to bypass Gatekeeper)"
echo "  2. Upload PowerUp-Portal-Mac.zip to Google Drive for your teammates"
echo

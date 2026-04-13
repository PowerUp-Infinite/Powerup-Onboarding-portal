#!/bin/bash
# ─────────────────────────────────────────────────────────────
# PowerUp Portal — build the macOS distribution folder.
#
# Produces: desktop/dist/PowerUp-Portal-Mac/
#
# Contents:
#   install.sh                          (from macos-dist/)
#   run-PowerUp-Portal.command          (from macos-dist/)
#   INSTALL.txt                         (from macos-dist/)
#   requirements.txt                    (filtered — no pyinstaller)
#   app/main.py, gui.py, app_config.py, ui_theme.py
#   app/workers/
#   app/portal/                         (m2_engine, m3_engine, sheets, config)
#   app/portal_shims/                   (streamlit-free google_auth)
#   app/resources/
#       credentials.json                (MUST exist in desktop/resources/)
#       .env                            (copied from ../portal/.env)
#
# Zip the result and upload to Google Drive for distribution.
# Runs on either macOS or Windows (via git-bash).
# ─────────────────────────────────────────────────────────────
set -e
cd "$(dirname "$0")"

OUT="dist/PowerUp-Portal-Mac"
APP="$OUT/app"
RES="$APP/resources"

echo
echo "========================================"
echo "  PowerUp Portal - macOS packager"
echo "========================================"
echo

# --- 1. Sanity checks ---
echo "[1/6] Sanity checks..."
if [[ ! -f "../portal/.env" ]]; then
    echo "[ERROR] ../portal/.env is missing. Cannot package."
    exit 1
fi
if [[ ! -f "resources/credentials.json" ]]; then
    echo "[ERROR] desktop/resources/credentials.json is missing."
    echo "        Copy your service account JSON to that path before packaging."
    exit 1
fi
if [[ ! -f "macos-dist/install.command" ]]; then
    echo "[ERROR] macos-dist/install.command is missing. Repo is broken."
    exit 1
fi

# --- 2. Clean output folder ---
echo "[2/6] Clearing old output folder..."
rm -rf "$OUT"
mkdir -p "$APP" "$RES"

# --- 3. Copy template files ---
echo "[3/6] Copying launcher scripts (install.command, run.command, INSTALL.txt)..."
cp -f macos-dist/install.command             "$OUT/"
cp -f macos-dist/run-PowerUp-Portal.command  "$OUT/"
cp -f macos-dist/INSTALL.txt                 "$OUT/"
chmod +x "$OUT/install.command" "$OUT/run-PowerUp-Portal.command"

# --- 4. Filter requirements (drop pyinstaller — users don't build) ---
echo "[4/6] Generating filtered requirements.txt..."
grep -v -i '^pyinstaller' requirements.txt > "$OUT/requirements.txt"

# --- 5. Copy the Python code ---
echo "[5/6] Copying Python code into app/ ..."
cp -f main.py       "$APP/"
cp -f gui.py        "$APP/"
cp -f app_config.py "$APP/"
cp -f ui_theme.py   "$APP/"

cp -r workers      "$APP/workers"
cp -r portal_shims "$APP/portal_shims"

# portal/ — whitelist only what's needed
mkdir -p "$APP/portal"
cp -f ../portal/m2_engine.py        "$APP/portal/"
cp -f ../portal/m3_engine.py        "$APP/portal/"
cp -f ../portal/agreement_engine.py "$APP/portal/"
cp -f ../portal/sheets.py           "$APP/portal/"
cp -f ../portal/config.py           "$APP/portal/"

# Resources
cp -f resources/credentials.json "$RES/"
cp -f ../portal/.env             "$RES/.env"

# Strip __pycache__ dirs that snuck in via cp -r
find "$APP" -type d -name '__pycache__' -exec rm -rf {} + 2>/dev/null || true

# --- 6. Report ---
echo "[6/6] Done."
echo
echo "========================================"
echo "  Package built successfully"
echo "========================================"
echo
echo "Output folder:  $(pwd)/$OUT"
echo
echo "Next:"
echo "  1. Zip the folder for cleaner distribution:"
echo "       cd dist && zip -ry PowerUp-Portal-Mac.zip PowerUp-Portal-Mac"
echo "     (zipping preserves the executable bits on the .sh / .command files)"
echo "  2. Upload PowerUp-Portal-Mac.zip to your shared Google Drive"
echo "  3. Mac teammates download, extract, open INSTALL.txt, follow steps"
echo

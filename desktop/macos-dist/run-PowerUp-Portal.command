#!/bin/bash
# ─────────────────────────────────────────────────────────────
# PowerUp Portal — launcher (macOS)
#
# Double-click this file in Finder to start the app.
# .command extension makes Finder treat it as a runnable shell script.
# If nothing happens, run install.sh first.
# ─────────────────────────────────────────────────────────────
cd "$(dirname "$0")"

if [[ ! -f ".venv/bin/python3" ]]; then
    osascript -e 'display alert "PowerUp Portal" message "Setup not complete.\n\nOpen the install.sh file first (right-click → Open With → Terminal). After it finishes, double-click this file again."'
    exit 1
fi

# Launch the app detached so the terminal window can close once it's running.
exec ./.venv/bin/python3 app/main.py >/dev/null 2>&1 &
disown
exit 0

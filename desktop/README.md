# PowerUp Portal — Desktop App

Standalone local build of the PowerUp Portal for M1, M2, M3 deck generation.
Runs entirely on the user's machine; reads data from a local Excel file; fetches
questionnaire responses + monthly reference data from Google Sheets; uploads the
generated deck to the same Drive folders the cloud portal uses.

## Who should use which version

- **Cloud portal** (`portal/` in this repo, on Streamlit): shared, no install.
- **Desktop app** (this folder): fallback for when the cloud portal is down,
  or for teammates who prefer a native .exe / .app.

Both write to the same Drive folders and the same Main Data spreadsheet, so
output is indistinguishable.

---

## For end users (your teammates)

Download the folder for your OS from the shared Google Drive:

### Windows
1. Download `PowerUp-Portal-Windows/PowerUp-Portal.exe`.
2. Double-click to run. Done.

### macOS
1. Download `PowerUp-Portal-Mac/PowerUp-Portal-Mac.zip`.
2. Unzip it. A `PowerUp-Portal.app` appears.
3. **First-time only**: right-click the app → **Open** → **Open** in the
   prompt. macOS will remember this choice.
4. From then on, double-click to open normally.

### Usage (all OSes)
1. Pick the tab: **M1 Report**, **M2 Deck**, or **M3 Deck**.
2. Click **📂 Upload Excel** and pick the client's file.
3. If the file has more than one client, pick the one you want from the
   dropdown.
4. (M2 only) Optionally pick a questionnaire response to match.
   (M3 only) Enter the client name.
5. Click **⚙️ Generate**. Wait up to 2 minutes.
6. Click **🌐 Open in Google Drive** to view the output.

No setup. No login. The app bundles the service account credentials, so it
authenticates to Google automatically.

---

## For the developer (you, building the distribution)

### One-time setup
- Make sure `portal/.env` exists and contains all the Drive / Sheets IDs. The
  build reads them and bakes them into the binary.
- Make sure `desktop/resources/credentials.json` exists. If not, copy the
  service account JSON from the repo root:
  ```
  cp pre-onboarding-portal-487d94175c10.json desktop/resources/credentials.json
  ```

### Build on Windows
```cmd
cd desktop
build-windows.bat
```
Produces `desktop/dist/PowerUp-Portal.exe` (~250 MB, single file).

### Build on macOS
```bash
cd desktop
chmod +x build-mac.sh
./build-mac.sh
```
Produces `desktop/dist/PowerUp-Portal.app`.

**Cross-compilation is not possible** — you have to build the Mac binary on
a Mac and the Windows binary on a Windows machine. PyInstaller does not
support Windows → Mac or vice versa.

### Distribution
After building on each OS:

1. Create two folders in your shared Drive:
   - `PowerUp-Portal-Windows/` → upload `PowerUp-Portal.exe`
   - `PowerUp-Portal-Mac/` → upload `PowerUp-Portal-Mac.zip`
     (zip the `.app` first: `cd dist && zip -ry PowerUp-Portal-Mac.zip PowerUp-Portal.app`)
2. Share each folder with your team (View-only is fine).
3. Paste a `HOW-TO-OPEN.txt` next to the Mac zip with the right-click instruction.

---

## Security note

The `.exe` / `.app` contains your Google service account private key bundled
inside. Anyone with the binary can extract the key. Keep this in mind:

- **Don't publish the binary publicly** (GitHub releases, public Drive links,
  etc.). Share only inside your team.
- **Rotate the service account key** if a teammate leaves or if the binary
  ever leaks. Regenerate the JSON, re-run the build, replace the binary.
- The service account only has access to the specific Drive folders /
  spreadsheets you've shared with it, so the blast radius is bounded.

---

## Architecture

```
desktop/
├── main.py                ← entry point
├── app_config.py          ← loads portal/.env + adds portal/ to sys.path
├── gui.py                 ← customtkinter 3-tab window
├── workers/
│   ├── common.py          ← Excel parser, client picker, Drive upload
│   ├── m1_worker.py       ← M1: sync to Sheets → call Apps Script
│   ├── m2_worker.py       ← M2: local Excel → generate_deck → upload
│   └── m3_worker.py       ← M3: local Excel + Sheets ref → generate_deck → upload
├── resources/
│   └── credentials.json   ← service account (BUNDLED into .exe/.app)
├── PowerUp-Portal.spec    ← PyInstaller spec (cross-platform)
├── build-windows.bat      ← Windows build script
├── build-mac.sh           ← macOS build script
└── requirements.txt       ← Python deps (includes PyInstaller)
```

**Key design choices:**

- **Reuses `portal/m2_engine.py`, `portal/m3_engine.py`, `portal/sheets.py` directly.**
  No code duplication — the desktop app is a thin shell around the same
  engines the cloud portal uses. Bug fixes in `portal/` automatically flow
  into the next desktop build.

- **M1 still goes via the Apps Script**, because the Apps Script is what
  renders the M1 Google Sheet. Desktop syncs the Excel to Main Data first,
  then triggers the Apps Script, same as the cloud portal.

- **M2 does NOT sync to Sheets.** Reads the uploaded Excel directly into a
  data dict and passes it to `generate_deck(data=...)`. Only the
  questionnaire is fetched from Sheets (because the form submits there).

- **M3 reads uploaded Excel for client sections, fetches monthly reference
  data from Sheets.** Same as the cloud portal.

---

## Troubleshooting

### "This app can't be opened" on macOS
Right-click → **Open** (don't double-click). macOS blocks unsigned apps on
first launch; this bypass only needs to be done once per install.

### `ModuleNotFoundError: No module named 'm2_engine'` at runtime
PyInstaller didn't bundle `portal/`. Check:
- `portal/` folder exists at the repo root
- Your build command was run from `desktop/`, not the repo root
- Try `pyinstaller PowerUp-Portal.spec --clean --noconfirm` again

### "No Google credentials found" error on first launch
`desktop/resources/credentials.json` wasn't bundled. Verify it exists before
building and rebuild.

### Build is huge (>300 MB)
That's normal for a PyInstaller one-file bundle with pandas + numpy +
matplotlib + python-pptx + customtkinter. Can't easily reduce without
dropping features.

### App launches but can't upload to Drive
The service account in `credentials.json` has been rotated or lost permission
on the target folder. Re-share the Drive folders with the service account's
email address (`pre-onboarding-portal@...iam.gserviceaccount.com`), or
regenerate the key and rebuild.

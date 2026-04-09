# M2 Apps Script — Setup Guide

## Overview

This is a test version of the M2 Strategy Deck generator ported to Google Apps Script.
It runs entirely inside Google's infrastructure — no Python, no Streamlit, no external server.

## Key Advantages Over Python Version

| Aspect | Python (Streamlit) | Apps Script |
|--------|-------------------|-------------|
| Sheets read | HTTP API round trips | Native `SpreadsheetApp` (instant) |
| Slides manipulation | PPTX → convert → upload | Native `SlidesApp` (in-place) |
| Template handling | Download PPTX → export | `makeCopy()` → edit directly |
| Chart generation | matplotlib → PNG → insert | Sheets Charts → embed |
| Deployment | Streamlit Cloud | Built-in web app |
| Dependencies | pip install ~10 packages | Zero |

## Setup Steps

### 1. Create the Apps Script project

1. Go to [script.google.com](https://script.google.com)
2. Create a new project → name it "M2 Strategy Deck"
3. Create separate `.gs` files for each file in this folder:
   - `Config.gs`
   - `DataLoader.gs`
   - `Helpers.gs`
   - `SlideBuilder.gs`
   - `Slide13Chart.gs`
   - `Appendix.gs`
   - `RiskReward.gs`
   - `Questionnaire.gs`
   - `Main.gs`
4. Copy-paste the contents of each file

### 2. Update Config.gs

Replace all `YOUR_*` placeholders with actual IDs from your `.env` file:

```
MAIN_SPREADSHEET_ID        → your MAIN_SPREADSHEET_ID
TIMESERIES_SPREADSHEET_ID  → your TIMESERIES_SPREADSHEET_ID
QUESTIONNAIRE_SPREADSHEET_ID → your QUESTIONNAIRE_SPREADSHEET_ID
M2_BASE_DECK_ID            → your M2_TEMPLATE_ID
M2_RISK_REWARD_DECK_ID     → your M2_RISK_REWARD_TEMPLATE_ID
M2_CATEGORIZATION_FILE_ID  → your M2_CATEGORIZATION_FILE_ID
RATING_IMAGE_IDS            → your M2_IMG_*_ID values
M2_OUTPUT_FOLDER_ID        → your M2_OUTPUT_FOLDER_ID
```

### 3. Enable APIs

In the Apps Script editor:
1. Click the gear icon (Project Settings)
2. Check "Show 'appsscript.json' manifest file in editor"
3. Add these OAuth scopes to `appsscript.json`:

```json
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive"
  ]
}
```

### 4. Test

1. Open `Main.gs`
2. Edit `testGenerate()` with a real PF_ID and client name
3. Click Run → `testGenerate`
4. First run will ask for permissions — grant them
5. Check Execution Log for output

### 5. Deploy as Web App (optional)

1. Click Deploy → New deployment
2. Type: Web app
3. Execute as: Me
4. Access: Anyone within your organization
5. Deploy → copy the URL
6. Open the URL to use the web interface

## Architecture Notes

### How it differs from the Python version

**Template handling**: The Python version downloads the Google Slides template as PPTX,
manipulates it with `python-pptx` (low-level XML), then re-uploads as PPTX and converts
back to Slides. The Apps Script version copies the template directly as a Google Slides
presentation and edits it in-place using `SlidesApp`. This eliminates two conversions and
multiple Drive API calls.

**Chart generation**: The Python version uses matplotlib to generate PNG images for the
donut and line charts. The Apps Script version creates temporary spreadsheets with chart
data, builds native Google Charts, exports them as images, and inserts them into the slides.
The temp spreadsheets are auto-deleted.

**Shape matching**: The Python version matches shapes by their internal name (e.g.,
`Google Shape;165;p19`). These names are preserved in native Google Slides, so we could
use them via the Advanced Slides API. However, for simplicity, this test version matches
shapes by text content patterns, which is more robust across template changes.

### Limitations of this test version

1. **Packed scheme slides** (two subcategories on one slide) not yet implemented — each
   subcategory gets its own slide. This can be added if the basic flow works.

2. **Legend manipulation** on the pie chart slide is simplified — the chart image includes
   data labels but the template's static legend text may need manual fine-tuning.

3. **Portfolio Preference formatting** (multi-line styled answer with red/green runs) is
   simplified to plain text. The Slides API can do rich formatting but requires more code.

4. **Execution time limit**: Apps Script has a 6-minute limit (30 min for Workspace paid).
   If the deck generation takes too long, we can split into smaller operations or use
   batch Slides API calls.

## Next Steps

1. Test with a real client to verify shape matching works
2. Compare output with the Python-generated deck
3. Fine-tune positions/formatting as needed
4. If successful, port M3 using the same pattern

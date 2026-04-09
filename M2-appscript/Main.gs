/**
 * Main.gs — Entry point for M2 Strategy Deck generation via Apps Script.
 *
 * Usage:
 *   1. Open Apps Script editor (script.google.com)
 *   2. Create a new project and paste all .gs files
 *   3. Update Config.gs with your actual IDs
 *   4. Run generateM2Deck() or use testGenerate() for testing
 *
 * Flow:
 *   1. Copy the base deck template → new Slides presentation
 *   2. Load all data from Google Sheets
 *   3. Match client's questionnaire response
 *   4. Calculate risk profile
 *   5. Fill slides 1-6 (title, welcome, glance, snapshot, working well)
 *   6. Generate charts (pie chart slide 4, line chart slide 13)
 *   7. Insert risk-reward slides (15-18)
 *   8. Build appendix scheme slides
 *   9. Populate and filter questionnaire slides
 *  10. Move output to the target Drive folder
 *
 * Key advantage over Python:
 *   - Google Sheets reads are native (no HTTP round trips)
 *   - Slide manipulation is native (no PPTX conversion)
 *   - Runs inside Google infra (faster API calls)
 *   - No dependency management (no pip install)
 */


/**
 * Generate the M2 strategy deck for a given client.
 *
 * @param {string} pfId - The client's PF_ID
 * @param {string} customerName - Display name for the client
 * @param {string} [questionnaireName] - Name to match in questionnaire (optional)
 * @returns {object} { url, name, id } of the generated presentation
 */
function generateM2Deck(pfId, customerName, questionnaireName) {
  Logger.log('='.repeat(60));
  Logger.log('Generating deck for: ' + customerName + ' (' + pfId + ')');
  Logger.log('='.repeat(60));

  const startTime = new Date();

  // ── Step 1: Copy base deck template ───────────────────────
  Logger.log('\n[1/10] Copying base deck template...');
  const safeName = customerName.replace(/[^\w\s-]/g, '').trim().replace(/\s+/g, '_');
  const deckName = safeName + '_' + pfId.substring(0, 12) + '_deck';

  const baseDeckFile = DriveApp.getFileById(M2_BASE_DECK_ID);
  const outputFolder = DriveApp.getFolderById(M2_OUTPUT_FOLDER_ID);
  const newFile = baseDeckFile.makeCopy(deckName, outputFolder);
  const newFileId = newFile.getId();

  const presentation = SlidesApp.openById(newFileId);
  Logger.log('  Created: ' + deckName + ' (id=' + newFileId + ')');

  // ── Step 2: Load data ─────────────────────────────────────
  Logger.log('\n[2/10] Loading data from Sheets...');
  const data = loadAllData();

  // ── Step 3: Validate PF_ID ────────────────────────────────
  const pfRow = getPfRow(data, pfId);
  if (!pfRow) {
    throw new Error('PF_ID "' + pfId + '" not found in PF_level sheet.');
  }

  // ── Step 4: Match questionnaire ───────────────────────────
  Logger.log('\n[3/10] Matching questionnaire...');
  const qRow = matchQuestionnaire(data, pfId, customerName, questionnaireName);

  // ── Step 5: Risk profile ──────────────────────────────────
  Logger.log('\n[4/10] Calculating risk profile...');
  const riskProfile = calcRiskProfile(qRow);
  Logger.log('  Risk profile: ' + riskProfile);

  // ── Step 6: Riskgroup aggregation ─────────────────────────
  const rgAgg = getRiskgroupAgg(data, pfId);

  const firstName = customerName.split(/\s+/)[0] || 'Client';

  // ── Step 7: Fill slides ───────────────────────────────────
  Logger.log('\n[5/10] Slide 1 - Title');
  doSlide1(presentation, customerName);

  Logger.log('[6/10] Slide 2 - Welcome');
  doSlide2(presentation, firstName);

  Logger.log('[7/10] Slide 3 - You at a Glance');
  if (qRow) {
    doSlide3(presentation, qRow, riskProfile);
  } else {
    Logger.log('  SKIPPED (no questionnaire data)');
  }

  Logger.log('[8/10] Slide 4 - Portfolio Snapshot');
  doSlide4(presentation, pfRow, rgAgg, riskProfile);

  Logger.log('[8b/10] Slide 6 - What\'s working well');
  doSlide6(presentation, pfRow, riskProfile);

  Logger.log('[9/10] Slide 13 - Portfolio vs Infinite');
  doSlide13(presentation, pfId, riskProfile, data);

  Logger.log('[10/10] Appendix - Scheme Slides');
  const nAppendix = doAppendix(presentation, pfId, data);

  Logger.log('[11/10] Risk Reward Slides');
  const rrGoals = qRow ? parseGoals(qRow['Goals'] || '') : [];
  doRiskRewardSlides(presentation, riskProfile, rrGoals);

  Logger.log('[12/10] Questionnaire Slides');
  const goals = qRow ? parseGoals(qRow['Goals'] || '') : [];
  doQuestionnaire(presentation, goals, qRow);

  // ── Done ──────────────────────────────────────────────────
  presentation.saveAndClose();

  const elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  const url = 'https://docs.google.com/presentation/d/' + newFileId + '/edit';

  Logger.log('\n' + '='.repeat(60));
  Logger.log('DONE in ' + elapsed + 's -> ' + deckName);
  Logger.log('URL: ' + url);
  Logger.log('='.repeat(60));

  return {
    id:   newFileId,
    name: deckName,
    url:  url,
  };
}


// ── Test function ───────────────────────────────────────────

/**
 * Test function — update PF_ID and name, then run from Apps Script editor.
 */
function testGenerate() {
  const result = generateM2Deck(
    'TEST_PF_ID',           // Replace with a real PF_ID
    'Test Client Name',     // Replace with client name
    null                    // or a questionnaire name
  );
  Logger.log('Generated: ' + result.url);
}


// ── Web app interface (optional) ────────────────────────────

/**
 * Serve a simple web UI for triggering generation.
 * Deploy as web app: Run as → Me, Access → Anyone in org.
 */
function doGet(e) {
  return HtmlService.createHtmlOutput(getWebUI());
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const result = generateM2Deck(
      params.pfId,
      params.customerName,
      params.questionnaireName || null
    );
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function getWebUI() {
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; max-width: 500px; margin: 40px auto; padding: 20px; }
    h2 { color: #1a1a2e; }
    label { display: block; margin-top: 16px; font-weight: 600; }
    input { width: 100%; padding: 8px; margin-top: 4px; border: 1px solid #ccc; border-radius: 4px; }
    button { margin-top: 20px; padding: 12px 24px; background: #2E8AE5; color: white;
             border: none; border-radius: 4px; cursor: pointer; font-size: 16px; width: 100%; }
    button:hover { background: #1a6cb5; }
    #result { margin-top: 20px; padding: 16px; background: #f0f8f0; border-radius: 4px; display: none; }
    #error { margin-top: 20px; padding: 16px; background: #fdf0f0; border-radius: 4px; display: none; color: #c00; }
    .spinner { display: none; margin-top: 20px; text-align: center; }
  </style>
</head>
<body>
  <h2>M2 Strategy Deck Generator</h2>
  <p>Generate a personalised strategy deck for a client.</p>

  <label>PF ID</label>
  <input type="text" id="pfId" placeholder="Enter PF_ID">

  <label>Client Name</label>
  <input type="text" id="customerName" placeholder="Full name">

  <label>Questionnaire Name (optional)</label>
  <input type="text" id="questionnaireName" placeholder="Name to match in questionnaire">

  <button onclick="generate()">Generate Deck</button>

  <div class="spinner" id="spinner">Generating deck... this may take a few minutes.</div>
  <div id="result"></div>
  <div id="error"></div>

  <script>
    function generate() {
      const pfId = document.getElementById('pfId').value.trim();
      const name = document.getElementById('customerName').value.trim();
      const qName = document.getElementById('questionnaireName').value.trim();

      if (!pfId || !name) { alert('Enter PF ID and Client Name'); return; }

      document.getElementById('spinner').style.display = 'block';
      document.getElementById('result').style.display = 'none';
      document.getElementById('error').style.display = 'none';

      google.script.run
        .withSuccessHandler(function(res) {
          document.getElementById('spinner').style.display = 'none';
          if (res.error) {
            document.getElementById('error').textContent = 'Error: ' + res.error;
            document.getElementById('error').style.display = 'block';
          } else {
            document.getElementById('result').innerHTML =
              '<strong>Deck generated!</strong><br>' +
              '<a href="' + res.url + '" target="_blank">Open: ' + res.name + '</a>';
            document.getElementById('result').style.display = 'block';
          }
        })
        .withFailureHandler(function(err) {
          document.getElementById('spinner').style.display = 'none';
          document.getElementById('error').textContent = 'Error: ' + err.message;
          document.getElementById('error').style.display = 'block';
        })
        .generateM2Deck(pfId, name, qName || null);
    }
  </script>
</body>
</html>`;
}

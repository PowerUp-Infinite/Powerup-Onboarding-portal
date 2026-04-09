/**
 * DataLoader.gs — Read data from Google Sheets into structured objects.
 *
 * Each read function returns an array of row-objects (column name → value).
 * This is the equivalent of pandas DataFrames in Python.
 */

/**
 * Read a sheet tab into an array of {colName: value} objects.
 * Row 1 = headers. Subsequent rows = data.
 */
function readSheetAsObjects(spreadsheetId, tabName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error(`Sheet tab "${tabName}" not found in ${spreadsheetId}`);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  return rows;
}


/**
 * Load all M2 data from Google Sheets.
 * Returns a data object with all sheets as arrays of row-objects.
 */
function loadAllData() {
  Logger.log('Loading all data from Sheets...');
  const data = {};

  data.pfLevel     = readSheetAsObjects(MAIN_SPREADSHEET_ID, TABS.PF_LEVEL);
  data.schemeLevel = readSheetAsObjects(MAIN_SPREADSHEET_ID, TABS.SCHEME_LEVEL);
  data.riskgroup   = readSheetAsObjects(MAIN_SPREADSHEET_ID, TABS.RISKGROUP_LEVEL);
  data.results     = readSheetAsObjects(MAIN_SPREADSHEET_ID, TABS.RESULTS);
  data.lines       = readSheetAsObjects(TIMESERIES_SPREADSHEET_ID, TABS.LINES);
  data.invested    = readSheetAsObjects(TIMESERIES_SPREADSHEET_ID, TABS.INVESTED_VALUE_LINE);

  // Questionnaire
  const qSS = SpreadsheetApp.openById(QUESTIONNAIRE_SPREADSHEET_ID);
  const qSheet = qSS.getSheets()[0]; // first tab
  const qData = qSheet.getDataRange().getValues();
  if (qData.length >= 2) {
    const qHeaders = qData[0].map(h => String(h).trim());
    data.questionnaire = [];
    for (let i = 1; i < qData.length; i++) {
      const row = {};
      for (let j = 0; j < qHeaders.length; j++) {
        row[qHeaders[j]] = qData[i][j];
      }
      data.questionnaire.push(row);
    }
  } else {
    data.questionnaire = [];
  }

  // Convert numeric columns
  const numericSheets = ['pfLevel', 'riskgroup', 'schemeLevel', 'results', 'lines', 'invested'];
  const skipCols = new Set([
    'PF_ID', 'ISIN', 'NAME', 'FUND_NAME', 'FUND_STANDARD_NAME',
    'FUND_LEGAL_NAME', 'TYPE', 'POWERRATING', 'DISTRIBUTION_STATUS',
    'RISK_GROUP_L0', 'UPDATED_SUBCATEGORY', 'UPDATED_BROAD_CATEGORY_GROUP',
    'BROAD_CATEGORY_GROUP', 'DERIVED_CATEGORY', 'Purchase Mode',
    'BM', 'DIR_ISIN', 'ALT_ISIN_J', 'DATE'
  ]);

  for (const key of numericSheets) {
    for (const row of data[key]) {
      for (const col in row) {
        if (skipCols.has(col)) continue;
        const v = row[col];
        if (v === '' || v === null || v === undefined) {
          row[col] = NaN;
        } else if (typeof v === 'string') {
          const n = Number(v);
          row[col] = isNaN(n) ? v : n;
        }
      }
    }
  }

  Logger.log(`Loaded: pfLevel=${data.pfLevel.length}, scheme=${data.schemeLevel.length}, ` +
             `riskgroup=${data.riskgroup.length}, questionnaire=${data.questionnaire.length}`);
  return data;
}


/**
 * Get PF row by PF_ID.
 */
function getPfRow(data, pfId) {
  return data.pfLevel.find(r => String(r.PF_ID) === String(pfId)) || null;
}


/**
 * Get riskgroup rows for a PF_ID, aggregated by RISK_GROUP_L0.
 */
function getRiskgroupAgg(data, pfId) {
  const rows = data.riskgroup.filter(r => String(r.PF_ID) === String(pfId));
  const agg = {};
  for (const r of rows) {
    const g = r.RISK_GROUP_L0;
    if (!agg[g]) agg[g] = { group: g, pctOfPF: 0, currentValue: 0 };
    agg[g].pctOfPF      += (r['% of PF'] || 0);
    agg[g].currentValue += (r.CURRENT_VALUE || 0);
  }
  return Object.values(agg);
}


/**
 * Match questionnaire row.
 * Priority: 1) PF_ID match, 2) exact name, 3) partial first-name.
 */
function matchQuestionnaire(data, pfId, customerName, questionnaireName) {
  const qdf = data.questionnaire;
  if (!qdf.length) return null;

  // 1. Saved questionnaire name (exact match)
  if (questionnaireName) {
    const match = qdf.find(r =>
      String(r.Name || '').trim().toLowerCase() === questionnaireName.trim().toLowerCase()
    );
    if (match) {
      Logger.log(`Questionnaire: matched by saved name -> "${match.Name}"`);
      return match;
    }
  }

  // 2. PF_ID column
  const pfMatch = qdf.find(r => String(r.PF_ID || '') === String(pfId));
  if (pfMatch) {
    Logger.log(`Questionnaire: matched by PF_ID -> "${pfMatch.Name || '?'}"`);
    return pfMatch;
  }

  // 3. Exact name match
  const matchName = questionnaireName || customerName;
  const exact = qdf.find(r =>
    String(r.Name || '').trim().toLowerCase() === matchName.toLowerCase().trim()
  );
  if (exact) {
    Logger.log(`Questionnaire: exact name match -> "${exact.Name}"`);
    return exact;
  }

  // 4. Partial first-name
  if (matchName) {
    const first = matchName.toLowerCase().split(/\s+/)[0];
    const partial = qdf.find(r =>
      String(r.Name || '').toLowerCase().includes(first)
    );
    if (partial) {
      Logger.log(`Questionnaire: partial name match -> "${partial.Name}"`);
      return partial;
    }
  }

  Logger.log(`Questionnaire: no match for "${customerName}"`);
  return null;
}

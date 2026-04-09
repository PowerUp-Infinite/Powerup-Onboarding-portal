/**
 * Helpers.gs — Formatting helpers and risk profile calculation.
 */

// ── INR formatting ──────────────────────────────────────────

function fmtInrRupee(value, prefix) {
  prefix = prefix || '₹';
  if (value === null || value === undefined || isNaN(value) || value === 0) {
    return prefix + '0';
  }
  const av = Math.abs(value);
  const s = value < 0 ? '-' : '';
  if (av >= 1e7) {
    const cr = av / 1e7;
    return s + prefix + (cr < 10 ? cr.toFixed(1) : cr.toFixed(0)) + 'Cr';
  }
  if (av >= 1e5) {
    const l = av / 1e5;
    return s + prefix + (l < 10 ? l.toFixed(1) : l.toFixed(0)) + 'L';
  }
  if (av >= 1e3) {
    const k = av / 1e3;
    return s + prefix + (k < 10 ? k.toFixed(1) : k.toFixed(0)) + 'K';
  }
  return s + prefix + av.toFixed(0);
}

function fmtInrDisplay(value) {
  if (value === null || value === undefined || isNaN(value) || value === 0) return null;
  const av = Math.abs(value);
  if (av >= 1e7) {
    const cr = av / 1e7;
    return 'INR ' + (cr === Math.floor(cr) ? cr.toFixed(0) : cr.toFixed(1)) + ' Cr';
  }
  if (av >= 1e5) {
    const l = av / 1e5;
    return 'INR ' + (l === Math.floor(l) ? l.toFixed(0) : l.toFixed(1)) + 'L';
  }
  if (av >= 1e3) {
    const k = av / 1e3;
    return 'INR ' + (k === Math.floor(k) ? k.toFixed(0) : k.toFixed(1)) + 'K';
  }
  return 'INR ' + av.toFixed(0);
}

function fmtInr2dp(value, prefix) {
  prefix = prefix || '';
  if (value === null || value === undefined || isNaN(value) || value === 0) {
    return prefix + '0';
  }
  const av = Math.abs(value);
  const s = value < 0 ? '-' : '';
  if (av >= 1e7) return s + prefix + (av / 1e7).toFixed(2) + 'Cr';
  if (av >= 1e5) return s + prefix + (av / 1e5).toFixed(2) + 'L';
  if (av >= 1e3) return s + prefix + (av / 1e3).toFixed(2) + 'K';
  return s + prefix + av.toFixed(0);
}

function fmtSchemeVal(cv, pfPct) {
  return fmtInr2dp(cv) + ' (' + (pfPct * 100).toFixed(1) + '%)';
}

function fmtXirrPair(x, bx) {
  function f(v) { return (v === null || v === undefined || isNaN(v)) ? '-' : (v * 100).toFixed(1) + '%'; }
  return f(x) + ' | ' + f(bx);
}

function fmtMissed(mg) {
  if (mg === null || mg === undefined || isNaN(mg) || mg === 0) return '-';
  return fmtInr2dp(mg);
}


// ── Risk profile calculation ────────────────────────────────

function calcRiskProfile(qRow) {
  if (!qRow) return 'Balanced';

  // Step 1: Base from Portfolio Preference
  const pref = String(qRow['Portfolio Preference'] || '').toLowerCase();
  let idx;
  if (pref.includes('15%'))      idx = 4;
  else if (pref.includes('12%')) idx = 3;
  else if (pref.includes('9%'))  idx = 2;
  else if (pref.includes('6%'))  idx = 1;
  else                           idx = 2;
  const base = RISK_SCALE[idx];

  // Step 2: Horizon
  const horizon = String(qRow['Investment Horizon'] || '').toLowerCase();
  const longKws = ['more than 7', 'more than 8', 'long-term', 'long term', '8+'];
  const isLong = longKws.some(k => horizon.includes(k)) && !horizon.includes('medium');
  const hAdj = isLong ? 0 : -1;
  idx = Math.max(0, Math.min(4, idx + hAdj));

  // Step 3: Fall Reaction
  const fall = String(qRow['Fall Reaction'] || '').toLowerCase();
  let fAdj;
  if (fall.includes('invest more'))     fAdj = 1;
  else if (fall.includes('stay'))       fAdj = 0;
  else                                  fAdj = -1;
  idx = Math.max(0, Math.min(4, idx + fAdj));

  // Step 4: Liability management
  const liab = String(qRow['Liability Followup Answer'] || '').toLowerCase();
  const lAdj = (!liab || liab.includes('yes') || liab.includes('comfort')) ? 0 : -1;
  idx = Math.max(0, Math.min(4, idx + lAdj));

  const profile = RISK_SCALE[idx];
  Logger.log(`Risk: base=${base} h=${hAdj} f=${fAdj} l=${lAdj} -> ${profile}`);
  return profile;
}


function getHorizon(text) {
  if (!text) return '';
  const t = String(text).toLowerCase();
  for (const [k, v] of Object.entries(HORIZON_DISPLAY)) {
    if (t.includes(k)) return v;
  }
  return String(text);
}


function parseGoals(text) {
  if (!text) return [];
  return String(text).split(',').map(g => g.trim()).filter(g => g);
}


function portfolioRisk(sm) {
  if (sm < 15)  return 'Very Conservative';
  if (sm < 20)  return 'Conservative';
  if (sm < 40)  return 'Balanced';
  if (sm < 45)  return 'Aggressive';
  return 'Very Aggressive';
}


/**
 * Pick the best Infinite type for the comparison chart.
 */
function bestInfiniteType(pfId, prefix, resultsData) {
  const cust = resultsData.filter(r => String(r.PF_ID) === String(pfId));
  const v1Lump = prefix + '1 - lumpsum - 24M';
  if (cust.find(r => r.TYPE === v1Lump)) return v1Lump;
  const v1 = cust.find(r => String(r.TYPE || '').startsWith(prefix + '1'));
  if (v1) return v1.TYPE;
  const anyPref = cust.find(r => String(r.TYPE || '').startsWith(prefix));
  return anyPref ? anyPref.TYPE : null;
}

/**
 * Config.gs — Configuration constants for M2 Apps Script automation.
 *
 * All IDs and settings live here. Update these to match your environment.
 */

// ── Google Sheets data sources ──────────────────────────────
const MAIN_SPREADSHEET_ID        = '1PS3dhwtgVtyqw19LJkwNcVDLdQ002_WzHycaXZpnWIg';
const TIMESERIES_SPREADSHEET_ID  = '1D2qtelaD4hw2-5KwgvbhVWHwxhzcgdYTBX0r9tuxOv8';
const QUESTIONNAIRE_SPREADSHEET_ID = '1nGdm1hBRR2uI5YKlpjqiYvzWmGkvzf0YCRq1Cm1STYw';

// ── Sheet tab names ─────────────────────────────────────────
const TABS = {
  PF_LEVEL:            'PF_level',
  SCHEME_LEVEL:        'Scheme_level',
  RISKGROUP_LEVEL:     'Riskgroup_level',
  RESULTS:             'Results',
  LINES:               'Lines',
  INVESTED_VALUE_LINE: 'Invested_Value_Line',
};

// ── Google Slides templates (file IDs) ──────────────────────
const M2_BASE_DECK_ID           = '17_vG8hsm5D542_JPPoxlcV8RUMsK5qhgTUgTUXfADwM';
const M2_RISK_REWARD_DECK_ID    = '1HSWZHekW1gyUoi7yunhnQ6YP0foqmQI_09ZGsdWZwrc';

// ── Categorization file (Excel on Drive) ────────────────────
const M2_CATEGORIZATION_FILE_ID = '15T2toTd2l4zdhzNvuYAWdCxEf8ulvZuJ';

// ── Rating images (Drive file IDs) ──────────────────────────
const RATING_IMAGE_IDS = {
  IN_FORM:      '1ZML1rngTYNLwsnM1BqdFqhT1Yb1eaV6d',
  ON_TRACK:     '10UZ45x_maSMw-2J607TK6rwpdfX14AOH',
  OUT_OF_FORM:  '1Imc8B-vdciIPi5ToHWAK0ft0KVd8vyWW',
  OFF_TRACK:    '1wY0-Fm1Yv7QCt0rzEOmVAv2dEGpcTVdc',
};

// ── Output folder ───────────────────────────────────────────
const M2_OUTPUT_FOLDER_ID = '1jO-Yc031gGSnUpDZZ0QoTsKc3cGLgf3J';

// ── Chart colours ───────────────────────────────────────────
const CHART_COLORS = {
  '1) Aggressive':   '#2E8AE5',
  '2) Balanced':     '#4E9EED',
  '3) Conservative': '#6DB0F2',
  'Hybrid':          '#FFE2BF',
  'Debt Like':       '#EBF2F2',
  'Gold & Silver':   '#F7CB88',
  'Global':          '#FFC7B4',
  'Solution':        '#CABAF3',
};

const CHART_LABELS = {
  '1) Aggressive':   'Aggressive',
  '2) Balanced':     'Balanced',
  '3) Conservative': 'Conservative',
  'Hybrid':          'Hybrid',
  'Gold & Silver':   'Gold & Silver',
  'Debt Like':       'Debt',
  'Solution':        'Solution',
  'Global':          'Global',
};

// ── Risk profile scale ──────────────────────────────────────
const RISK_SCALE = [
  'Very Conservative', 'Conservative', 'Balanced', 'Aggressive', 'Very Aggressive'
];

const HORIZON_DISPLAY = {
  'short':           'Less than 3 Years',
  'less than 3':     'Less than 3 Years',
  '3-5':             '3-5 Years',
  'medium-term':     '3-5 Years',
  '5-7':             '5-7 Years',
  'medium to long':  '5-7 Years',
  'more than 7':     '8+ Years',
  'more than 8':     '8+ Years',
  'long-term':       '8+ Years',
  'long':            '8+ Years',
};

// Risk profile → TYPE prefix in Lines/Results
const RISK_TYPE_PREFIX = {
  'Very Aggressive':  'VA',
  'Aggressive':       'A',
  'Balanced':         'B',
  'Conservative':     'C',
  'Very Conservative':'VC',
};

// Risk reward deck: profile → 0-based start slide index (groups of 4)
const RISK_REWARD_IDX = {
  'Very Aggressive':  0,
  'Aggressive':       4,
  'Balanced':         8,
  'Conservative':     12,
  'Very Conservative': 12,
};

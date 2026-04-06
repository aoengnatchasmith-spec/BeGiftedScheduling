// ============================================================
// 00_Config.gs  –  Shared configuration and utility functions
// ❌ Do NOT run anything from this file directly.
//    All functions here are called by other .gs files.
// ============================================================

const CFG = {

  // ── Source file names (must match Google Drive exactly) ──────────────────
  SETUP_FILE_NAME:     'BeGifted Core Setup Data',
  ANALYTICS_FILE_NAME: 'BeGifted Education Analytics',
  PROFILE_FILE_NAME: 'Tutor Profile',

  // ── Source sheet tab names ────────────────────────────────────────────────
  MATRIX_SHEET:   'Subject-LevelMatrix',
  AVAIL_SHEET:    'Availability',
  PROFILE_SHEET:  'Active',
  SESSIONS_SHEET: 'Upcoming Sessions',

  // ── Output sheet names (auto-created in this spreadsheet) ─────────────────
  MASTER_SHEET:       'Master_Normalized',
  UPCOMING_SHEET:     'Upcoming_Sessions',
  PROFILE_NORM_SHEET: 'Tutor_Profile_Normalized',

  // ── Availability column layout (0-based, col A = index 0) ─────────────────
  AVAIL_NICK_COL:  1,                          // col B = Nickname
  AVAIL_TIME_COLS: [4, 6, 8, 10, 12, 14, 16], // E,G,I,K,M,O,Q → Mon-Sun time
  AVAIL_MODE_COLS: [5, 7, 9, 11, 13, 15, 17], // F,H,J,L,N,P,R → Mon-Sun mode
  DAYS: ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],

  // ── Analytics file: Upcoming Sessions column layout (0-based) ─────────────
  SESSION_COLS: {
    packageProgram:   4,  // col E
    tutorName:        5,  // col F  — "Firstname (Nickname) Lastname"
    sessionType:      6,  // col G
    scheduledDate:    9,  // col J  — read via getValues() for Date object
    scheduledTime:    10, // col K  — read via getDisplayValues() for "HH:MM"
    scheduledEnd:     12, // col M
    scheduledEndTime: 13, // col N  — read via getDisplayValues() for "HH:MM"
    sessionStatus:    15, // col P  — "UPCOMING" | "CANCELLED"
  },

  // ── Inactive teachers — excluded from matrix build ────────────────────────
  INACTIVE: [
    'Aong','Bank','Brook','Film','Kristiina','Gift-Kari',
    'JJ','Kam','Keane','Kim','Kristie','Maii',
    'Nop','Pamai','Pawin','Poom','Poon','Tom'
  ],

  // ── Country detection rules (Bachelor, Master, Doctoral only) ─────────────
  COUNTRY_KEYWORDS: {
    UK: {
      keywords:     [', UK', 'United Kingdom'],
      institutions: [
        'University College London','University of Edinburgh','Imperial College',
        'Brunel University','Cardiff','Cambridge','Oxford','Shrewsbury',
        "King's College",'LSE','London School of Economics','Manchester',
        'Bristol','Warwick','Bath','Durham','Leeds','Sheffield',
        'Nottingham','Birmingham','Liverpool','St Andrews','Exeter','York'
      ]
    },
    USA: {
      keywords:     [', USA', ', US', 'United States'],
      institutions: [
        'University of Michigan','Northwestern','Kellogg','College of New Jersey',
        'MIT','Harvard','Stanford','Yale','Columbia','Cornell',
        'Princeton','Duke','Johns Hopkins','UCLA','Berkeley','NYU',
        'Emory','Georgetown','Vanderbilt','Tufts','Boston University'
      ]
    },
    Australia: {
      keywords:     [', Australia', ', AUS'],
      institutions: [
        'University of Queensland','Griffith University','Monash',
        'UNSW','University of Melbourne','University of Sydney','ANU','Macquarie'
      ]
    },
    Thailand: {
      keywords:     [],
      institutions: [
        'Chulalongkorn','Mahidol','KMUTNB','KMUTT','Assumption University',
        'Ramkhamhaeng','Bangkok Patana','International School Bangkok',
        'Sasin','Thammasat','Kasetsart','NIDA','ABAC','Srinakharinwirot',
        'Silpakorn','Rangsit','Stamford'
      ]
    }
  }
};


// ═══════════════════════════════════════════════════════════════════════════════
// UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

// ─── File / Sheet ─────────────────────────────────────────────────────────────

function openFileByName_(name) {
  const files = DriveApp.getFilesByName(name);
  if (!files.hasNext()) {
    throw new Error(
      `[CFG] File not found in Drive: "${name}"\n` +
      `Check CFG.SETUP_FILE_NAME / CFG.ANALYTICS_FILE_NAME match exactly.`
    );
  }
  return SpreadsheetApp.open(files.next());
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}


// ─── Time ─────────────────────────────────────────────────────────────────────

// "9:00" or "09:30" → integer minutes. Returns null if unparseable.
function toMinutes_(timeStr) {
  const m = String(timeStr || '').trim().match(/^(\d{1,2}):(\d{2})$/);
  return m ? parseInt(m[1]) * 60 + parseInt(m[2]) : null;
}

// "9:00" → "09:00"
function zeroPad_(timeStr) {
  const [h, min] = String(timeStr).split(':');
  return String(parseInt(h)).padStart(2, '0') + ':' + min;
}

// Date object → "Monday" … "Sunday". Returns '' if invalid.
function getDayOfWeek_(dateVal) {
  const D = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  return (dateVal instanceof Date && !isNaN(dateVal)) ? D[dateVal.getDay()] : '';
}


// ─── Parsing ──────────────────────────────────────────────────────────────────

// Availability mode cell → { onsite, online }. Empty or "-" → both true.
function parseMode_(modeStr) {
  const s = String(modeStr || '').trim().toLowerCase();
  if (!s || s === '-') return { onsite: true, online: true };
  return { onsite: s.includes('onsite'), online: s.includes('online') };
}

// Parse one availability time cell into range objects.
// Handles: en-dash/em-dash, dot notation (16.00), "(online only)",
// comma-separated ranges, "and"-separated ranges, non-time notes in ().
function parseTimeRanges_(cellValue, defaultOnsite, defaultOnline) {
  if (!cellValue) return [];
  let text = String(cellValue)
    .replace(/[–—]/g, '-')
    .replace(/(\d+)\.(\d{2})/g, '$1:$2')
    .replace(/\s+/g, ' ')
    .trim();
  if (text === '-' || text === '') return [];

  const results = [];
  for (let part of text.split(/,|\band\b/i)) {
    part = part.trim();
    const onlineOnly = /\(online\s*only\)/i.test(part);
    const onsite = onlineOnly ? false : defaultOnsite;
    const online = onlineOnly ? true  : defaultOnline;
    part = part.replace(/\([^)]*\)/g, '').trim(); // strip notes
    const m = part.match(/(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/);
    if (!m) continue;
    results.push({ start: zeroPad_(m[1]), end: zeroPad_(m[2]), onsite, online });
  }
  return results;
}

// Extract nickname from "Firstname (Nickname) Lastname [Online]"
function extractNickname_(tutorName) {
  const m = String(tutorName || '').match(/\(([^)]+)\)/);
  return m ? m[1].trim() : '';
}

// "Kru Gift" → "Gift"  (case-insensitive, trims whitespace)
function extractProfileNickname_(nameCell) {
  return String(nameCell || '').replace(/^kru\s+/i, '').trim();
}

// Split degree cell (degree on line 1, university on line 2 within cell)
function parseDegreeCell_(cellValue) {
  if (!cellValue || String(cellValue).trim() === '') return { degree: '', university: '' };
  const lines = String(cellValue).split('\n').map(l => l.trim()).filter(Boolean);
  return { degree: lines[0] || '', university: lines[1] || '' };
}

// Detect graduation countries from Bachelor/Master/Doctoral text (NOT High School).
// Returns comma-joined country tags, e.g. "UK, Thailand".
// Returns "Other" if there is text but no country matched.
// Returns '' if all fields are empty.
function extractCountries_(degreeTexts) {
  const combined = degreeTexts.join('\n');
  if (!combined.trim()) return '';
  const found = new Set();
  for (const [country, rules] of Object.entries(CFG.COUNTRY_KEYWORDS)) {
    for (const kw of rules.keywords) {
      if (combined.includes(kw)) { found.add(country); break; }
    }
    if (!found.has(country)) {
      for (const inst of rules.institutions) {
        if (combined.toLowerCase().includes(inst.toLowerCase())) {
          found.add(country); break;
        }
      }
    }
  }
  if (found.size === 0) found.add('Other');
  return [...found].join(', ');
}


// ─── Logging ──────────────────────────────────────────────────────────────────

function logSection_(title) {
  Logger.log('\n' + '═'.repeat(55));
  Logger.log('  ' + title);
  Logger.log('═'.repeat(55));
}

function logLayer_(num, label, detail) {
  Logger.log(`  ▸ Layer ${num} [${label}]: ${detail}`);
}


// ─── Quick connectivity test ──────────────────────────────────────────────────
// Run testCFGAccess() alone first to confirm both files are reachable.

function testCFGAccess() {
  logSection_('testCFGAccess');
  try {
    const f = openFileByName_(CFG.SETUP_FILE_NAME);
    Logger.log('  ✓ Setup file: ' + f.getName());
    Logger.log('    Sheets: ' + f.getSheets().map(s => s.getName()).join(', '));
  } catch(e) { Logger.log('  ✗ ' + e.message); }
  try {
    const f = openFileByName_(CFG.ANALYTICS_FILE_NAME);
    Logger.log('  ✓ Analytics file: ' + f.getName());
  } catch(e) { Logger.log('  ✗ ' + e.message); }
}


/**
 * Read "SYNC EM Worship" XLSX from Google Drive in memory (no file conversion)
 * and return the row for the current worship week (matched on column A date).
 *
 * Setup:
 * 1. Create a standalone Apps Script project (or bind to a Google Sheet).
 * 2. Paste this file into the script editor.
 * 3. Add SheetJS once (see loadSheetJS_ below) — either CDN eval or paste xlsx.full.min.js.
 * 4. Run testGetCurrentWeekRow() and approve Drive access.
 */

const CONFIG = {
  /** From: .../d/1BSuNwm8YoW_UJkLoZuWY9Oqgy0hVpG-9/... */
  fileId: '1BSuNwm8YoW_UJkLoZuWY9Oqgy0hVpG-9',
  sheetName: '2026',
  /** Column A header text used to locate the header row automatically */
  dateHeader: 'Date',
  /** Worship rows use the Sunday of the current calendar week */
  matchStrategy: 'sunday', // 'sunday' | 'same-week'
};

/**
 * Public entry point — returns the current week's worship row as an object.
 * @returns {{
 *   rowNumber: number,
 *   date: Date,
 *   headers: string[],
 *   row: Object.<string, *>
 * }}
 */
function getCurrentWeekWorshipRow() {
  const sheetRows = readXlsxSheet_(CONFIG.fileId, CONFIG.sheetName);
  const headerIndex = findHeaderRowIndex_(sheetRows, CONFIG.dateHeader);
  const headers = sheetRows[headerIndex].map(normalizeHeader_);
  const dataRows = sheetRows.slice(headerIndex + 1);
  const targetDate = getTargetWeekDate_(CONFIG.matchStrategy);

  const match = findRowByDate_(dataRows, targetDate, CONFIG.matchStrategy);
  if (!match) {
    const targetLabel = Utilities.formatDate(
      targetDate,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );
    throw new Error(
      `No row found for ${CONFIG.matchStrategy} week containing ${targetLabel} on sheet "${CONFIG.sheetName}".`
    );
  }

  const rowObject = {};
  headers.forEach((header, i) => {
    if (header) rowObject[header] = match.row[i] ?? '';
  });

  return {
    rowNumber: headerIndex + 2 + match.index, // 1-based sheet row number
    date: match.date,
    headers,
    row: rowObject,
    rawRow: match.row,
  };
}

/** Manual smoke test — check Executions log for output */
function testGetCurrentWeekRow() {
  const result = getCurrentWeekWorshipRow();
  console.log('Matched row %s for %s', result.rowNumber, result.date);
  console.log(JSON.stringify(result.row, null, 2));
}

// ---------------------------------------------------------------------------
// XLSX parsing (in memory via SheetJS)
// ---------------------------------------------------------------------------

/**
 * Loads SheetJS into the global scope.
 *
 * Option A (quick): CDN eval on each run — works but adds ~1–2s latency.
 * Option B (recommended): download https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
 *   and add it as a separate script file named "xlsx.full.min.js" in the same project.
 *   Then delete or comment out the eval block below.
 */
function loadSheetJS_() {
  if (typeof XLSX !== 'undefined') return;

  const cdnUrl =
    'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
  eval(UrlFetchApp.fetch(cdnUrl).getContentText());
}

/**
 * @param {string} fileId
 * @param {string} sheetName
 * @returns {Array<Array<*>>} 2D array of cell values
 */
function readXlsxSheet_(fileId, sheetName) {
  loadSheetJS_();

  const file = DriveApp.getFileById(fileId);
  const mime = file.getMimeType();

  if (mime === MimeType.GOOGLE_SHEETS) {
    throw new Error(
      'File is a native Google Sheet. Either export it as .xlsx or use SpreadsheetApp instead.'
    );
  }

  const bytes = file.getBlob().getBytes();
  const data = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) {
    data[i] = bytes[i] & 0xff;
  }

  const workbook = XLSX.read(data, {
    type: 'array',
    cellDates: true,
    cellNF: true,
    cellText: true,
  });

  if (!workbook.Sheets[sheetName]) {
    throw new Error(
      `Sheet "${sheetName}" not found. Available sheets: ${workbook.SheetNames.join(', ')}`
    );
  }

  const worksheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      row.push(readCellValue_(worksheet, r, c));
    }
    rows.push(row);
  }

  return trimEmptyTrailingRows_(rows);
}

/**
 * Prefer formatted text (cell.w) so DD/MM/YYYY strings stay correct.
 * @param {Object} worksheet
 * @param {number} r
 * @param {number} c
 * @returns {*}
 */
function readCellValue_(worksheet, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = worksheet[addr];
  if (!cell) return '';

  if (cell.t === 'd' || cell.v instanceof Date) {
    return cell.v;
  }

  if (cell.w != null && cell.w !== '') {
    return cell.w;
  }

  return cell.v ?? '';
}

// ---------------------------------------------------------------------------
// Date / row matching
// ---------------------------------------------------------------------------

/**
 * @param {'sunday'|'same-week'} strategy
 * @returns {Date}
 */
function getTargetWeekDate_(strategy) {
  const today = stripTime_(new Date());

  if (strategy === 'same-week') {
    return today;
  }

  // Worship sheet dates are Sundays — use the Sunday of the current week
  return getSundayOfWeek_(today);
}

/**
 * @param {Date} date
 * @returns {Date}
 */
function getSundayOfWeek_(date) {
  const d = stripTime_(new Date(date));
  d.setDate(d.getDate() - d.getDay()); // Sunday = 0
  return d;
}

/**
 * @param {Date} date
 * @returns {Date}
 */
function stripTime_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * @param {Array<Array<*>>} rows
 * @param {Date} targetDate
 * @param {'sunday'|'same-week'} strategy
 * @returns {{index: number, row: Array<*>, date: Date}|null}
 */
function findRowByDate_(rows, targetDate, strategy) {
  for (let i = 0; i < rows.length; i++) {
    const parsed = parseCellDate_(rows[i][0]);
    if (!parsed) continue;

    const matches =
      strategy === 'same-week'
        ? isSameWeek_(parsed, targetDate)
        : isSameDay_(parsed, targetDate);

    if (matches) {
      return { index: i, row: rows[i], date: parsed };
    }
  }
  return null;
}

/**
 * @param {*} value
 * @returns {Date|null}
 */
function parseCellDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return stripTime_(value);
  }

  if (typeof value === 'number' && isFinite(value)) {
    // Excel serial (1900 date system)
    const utcDays = Math.floor(value - 25569);
    const date = new Date(utcDays * 86400 * 1000);
    return stripTime_(date);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;

    // DD/MM/YYYY (your sheet format)
    const dmy = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmy) {
      return stripTime_(new Date(+dmy[3], +dmy[2] - 1, +dmy[1]));
    }

    // ISO fallback
    const iso = new Date(trimmed);
    if (!isNaN(iso.getTime())) return stripTime_(iso);
  }

  return null;
}

/**
 * @param {Date} a
 * @param {Date} b
 * @returns {boolean}
 */
function isSameDay_(a, b) {
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
}

/**
 * Same calendar week (Sunday–Saturday).
 * @param {Date} a
 * @param {Date} b
 * @returns {boolean}
 */
function isSameWeek_(a, b) {
  return isSameDay_(getSundayOfWeek_(a), getSundayOfWeek_(b));
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * @param {Array<Array<*>>} rows
 * @param {string} dateHeader
 * @returns {number}
 */
function findHeaderRowIndex_(rows, dateHeader) {
  const target = dateHeader.trim().toLowerCase();

  for (let i = 0; i < rows.length; i++) {
    const firstCell = String(rows[i][0] ?? '')
      .trim()
      .toLowerCase();
    if (firstCell === target) return i;
  }

  throw new Error(`Header row with "${dateHeader}" in column A was not found.`);
}

/**
 * @param {*} header
 * @returns {string}
 */
function normalizeHeader_(header) {
  return String(header ?? '').trim();
}

/**
 * @param {Array<Array<*>>} rows
 * @returns {Array<Array<*>>}
 */
function trimEmptyTrailingRows_(rows) {
  let lastNonEmpty = rows.length - 1;
  while (lastNonEmpty >= 0 && isRowEmpty_(rows[lastNonEmpty])) {
    lastNonEmpty--;
  }
  return rows.slice(0, lastNonEmpty + 1);
}

/**
 * @param {Array<*>} row
 * @returns {boolean}
 */
function isRowEmpty_(row) {
  return row.every((cell) => String(cell ?? '').trim() === '');
}

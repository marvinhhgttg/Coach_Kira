
// --- KONFIGURATION V10 (Auto-Copy Timeline) ---
const SOURCE_TIMELINE_SHEET = 'timeline'; // NEU: Das Blatt mit Formeln
const TIMELINE_SHEET_NAME = 'KK_TIMELINE'; // Zielblatt für Werte
// --- PlanApp Snapshot (SG Stabilität) ---
const PLANAPP_SNAPSHOT_KEY = 'PLANAPP_SNAPSHOT_V1';
const KK_BUILD_ID = '2026-10-17__kira_chat';


function debugDump(options) {
  const opt = Object.assign({
    maxRowsPerSheet: 20,
    includeSheets: [
      SOURCE_TIMELINE_SHEET,
      TIMELINE_SHEET_NAME,
      WEEK_CONFIG_SHEET_NAME,
      LOAD_FACTOR_SHEET_NAME,
      ELEV_FACTOR_SHEET_NAME,
      BASELINE_SHEET_NAME,
      OUTPUT_HISTORY_SHEET,
      OUTPUT_FORECAST_SHEET,
      OUTPUT_STATUS_SHEET,
      OUTPUT_PLAN_SHEET
    ],
    // dump only first N columns to keep logs readable
    maxCols: 30,
    // if true, includes full values (still capped by maxRows/maxCols)
    includeValues: true
  }, options || {});

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  Logger.log('=== Coach_Kira debugDump ===');
  Logger.log(`Spreadsheet: ${ss.getName()} (${ss.getId()})`);
  Logger.log(`Time: ${Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss')} (${tz})`);
  Logger.log(`Build: ${KK_BUILD_ID}`);

  const sheetNames = ss.getSheets().map(s => s.getName());
  Logger.log(`Sheets (${sheetNames.length}): ${sheetNames.join(', ')}`);

  const seen = new Set();
  opt.includeSheets
    .filter(Boolean)
    .map(String)
    .forEach(name => {
      if (seen.has(name)) return;
      seen.add(name);

      const sh = ss.getSheetByName(name);
      if (!sh) {
        Logger.log(`--- Sheet '${name}': NOT FOUND`);
        return;
      }

      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      Logger.log(`--- Sheet '${name}': rows=${lastRow}, cols=${lastCol}`);

      if (lastRow === 0 || lastCol === 0) {
        Logger.log('(empty)');
        return;
      }

      const numRows = Math.min(lastRow, Math.max(1, opt.maxRowsPerSheet));
      const numCols = Math.min(lastCol, Math.max(1, opt.maxCols));

      // Read values in one go (fast). Note: will include headers + first data rows.
      const rng = sh.getRange(1, 1, numRows, numCols);
      const values = rng.getDisplayValues();

      // Header row always
      const header = values[0] || [];
      Logger.log(`Header[1]: ${JSON.stringify(header)}`);

      if (!opt.includeValues) return;

      // Remaining rows (up to maxRowsPerSheet)
      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        // Skip completely empty rows to reduce noise
        const isEmpty = row.every(c => String(c).trim() === '');
        if (isEmpty) continue;
        Logger.log(`Row[${r + 1}]: ${JSON.stringify(row)}`);
      }
    });

  // Optional: show relevant document properties keys (not values)
  try {
    const props = PropertiesService.getDocumentProperties().getProperties();
    const keys = Object.keys(props).sort();
    Logger.log(`DocProperties keys (${keys.length}): ${keys.join(', ')}`);
  } catch (e) {
    Logger.log(`DocProperties read failed: ${e.message}`);
  }

  Logger.log('=== /debugDump ===');
}

function _readPlanAppSnapshot_() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(PLANAPP_SNAPSHOT_KEY);
  if (!raw) return null;

  try {
    const snap = JSON.parse(raw);
    if (!snap || snap.active !== true) return null;

    // --- NEU: Snapshot ist nur am selben Kalendertag gültig ---
    const tz = Session.getScriptTimeZone();
    const todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // bevorzugt snapshotDateKey, sonst aus createdAt ableiten
    const snapKey =
      (snap.snapshotDateKey && String(snap.snapshotDateKey)) ||
      (snap.createdAt ? Utilities.formatDate(new Date(snap.createdAt), tz, 'yyyy-MM-dd') : null);

    if (snapKey && snapKey !== todayKey) {
      // alt -> wegräumen, damit UI nicht weiter "Snapshot aktiv" meldet
      props.deleteProperty(PLANAPP_SNAPSHOT_KEY);
      return null;
    }

    // Plausibilitätschecks
    if (typeof snap.todayRowIndex !== 'number') return null;
    if (!snap.startDate) return null;
    if (!Array.isArray(snap.ctlHistoryYesterday) || snap.ctlHistoryYesterday.length < 7) return null;

    return snap;
  } catch (e) {
    // kaputtes JSON -> löschen
    props.deleteProperty(PLANAPP_SNAPSHOT_KEY);
    return null;
  }
}


// Löscht Snapshot & erzeugt ihn neu aus dem aktuellen (LIVE) Stand
function refreshPlanAppSnapshot() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(PLANAPP_SNAPSHOT_KEY);

  // LIVE holen (weil Snapshot gerade gelöscht wurde)
  const base = getSimStartValues();

  const snap = {
    active: true,
    createdAt: Date.now(),
    startDate: base.startDate,
    todayRowIndex: base.todayRowIndex,
    startRowIndex: base.startRowIndex,
    todayIsClosed: !!base.todayIsClosed,
    atlYesterday: Number(base.atlYesterday) || 0,
    ctlYesterday: Number(base.ctlYesterday) || 0,
    ctlHistoryYesterday: Array.isArray(base.ctlHistoryYesterday) ? base.ctlHistoryYesterday.slice(-7) : []
  };

  // Robustheit: History auf 7 Werte bringen
  while (snap.ctlHistoryYesterday.length < 7) {
    snap.ctlHistoryYesterday.unshift(snap.ctlHistoryYesterday[0] || snap.ctlYesterday || 0);
  }

  props.setProperty(PLANAPP_SNAPSHOT_KEY, JSON.stringify(snap));
  return snap;
}

function clearPlanAppSnapshot() {
  PropertiesService.getDocumentProperties().deleteProperty(PLANAPP_SNAPSHOT_KEY);
  return true;
}

const BASELINE_SHEET_NAME = 'phys_baseline';
const OUTPUT_STATUS_SHEET = 'AI_REPORT_STATUS'; 
const OUTPUT_PLAN_SHEET = 'AI_REPORT_PLAN';
const LOG_SHEET_NAME = 'AI_LOG'; 
const AI_LOG_ENABLED = false; // (2026-09-28) Schreiben ins Blatt AI_LOG abgeschaltet; Blatt bleibt bestehen
const BACKUP_SHEET_NAME = 'AI_PLAN_BACKUP'; // <-- NEU (V35)
const OUTPUT_HISTORY_SHEET = 'AI_DATA_HISTORY';
const OUTPUT_FORECAST_SHEET = 'AI_DATA_FORECAST';
const LOAD_FACTOR_SHEET_NAME = 'KK_LOAD_FACTORS';
const ELEV_FACTOR_SHEET_NAME = 'KK_ELEV_FACTORS'; 
const WEEK_CONFIG_SHEET_NAME = 'week_config'; // <-- NEU (V26)
const ZUKUNFT_TAGE = 14; 
const VERGANGENHEIT_TAGE = 7;
const HISTORY_CHART_TAGE = 90; 
const TE_LIMIT_HIGH_AEROBIC = 4.0;
const TE_LIMIT_ANAEROBIC = 1.0;
const FIXED_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec";

// --- SPORT-WISSEN & REGELN (ZENTRAL) ---
const CONST_TE_GUIDELINES = `
  **DEINE SPORT-PRÄFERENZEN (WICHTIG):**
  - **KEIN SCHWIMMEN ("Swim").** Der Athlet schwimmt nicht.
  - **STATTDESSEN RUDERN ("Row").** Nutze "Row" als Alternative für Oberkörper/Ganzkörper-Cardio.
  
  **RICHTLINIEN FÜR TRAINING EFFECT (TE) ZIELE:**
  Schätze für jeden geplanten Tag die Ziel-TEs (0.0 - 5.0):
  - **Recovery/Z1:** Aerob 0.0-2.0 | Anaerob 0.0
  - **Endurance/Z2:** Aerob 2.5-3.5 | Anaerob 0.0-0.5
  - **Tempo/Z3:** Aerob 3.0-4.0 | Anaerob 0.0-1.0
  - **Threshold/Z4:** Aerob 3.5-4.8 | Anaerob 0.5-2.0
  - **VO2Max/Intervalle:** Aerob 3.0-4.0 | Anaerob 2.0-3.0 (Anaerob Fokus!)
  - **Sprint:** Aerob < 2.5 | Anaerob 2.5+
  - **Row/Rudern:** Aerob 2.5-3.5 | Anaerob 0.0-2.0 (je nach Intensität)
  - **Off/Pause:** Aerob 0.0 | Anaerob 0.0
`;
// ------------------------------------

/**
 * Runtime contract check for the core timeline sheets.
 *
 * Purpose:
 * - Fail fast with clear errors when sheet headers change or required columns are missing.
 * - Centralize header/column validation to reduce “undefined index” bugs.
 *
 * This function is intentionally **read-only** (no writes), and is safe to call
 * at the top of critical entry points (Supervisor, WebApp JSON endpoints).
 *
 * Usage examples:
 *   const ctx = validateTimelineContract_({ requireTodayRow: true });
 *   // ctx.timelineSheetName, ctx.timelineHeaders, ctx.idx.date, ...
 */
function validateTimelineContract_(options) {
  const opt = Object.assign({
    // If true, require that exactly one row has is_today == 1 (or at least one, depending on strictness)
    requireTodayRow: false,
    // If true, accept multiple is_today markers but use the last one
    allowMultipleToday: true,
    // Prefer KK_TIMELINE for reads if present; fallback to timeline
    preferKKTimeline: true,
    // Validate these columns exist (aliases supported per entry)
    required: {
      date: ['date', 'datum', 'day'],
      is_today: ['is_today'],
      sport: ['Sport_x', 'sport_x'],
      zone: ['Zone', 'zone', 'coach_zone'],
      ess: ['coachE_ESS_day', 'coache_ess_day'],
      teAe: ['Target_Aerobic_TE', 'target_aerobic_te'],
      teAn: ['Target_Anaerobic_TE', 'target_anaerobic_te'],
      fix: ['fix']
    }
  }, options || {});

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const kk = ss.getSheetByName(TIMELINE_SHEET_NAME);
  const src = ss.getSheetByName(SOURCE_TIMELINE_SHEET);

  const chosen = (opt.preferKKTimeline && kk) ? kk : src;
  if (!chosen) {
    const msg = `[Contract] Missing core sheet(s): '${TIMELINE_SHEET_NAME}' and '${SOURCE_TIMELINE_SHEET}' not found.`;
    if (typeof logToSheet === 'function') logToSheet('ERROR', msg);
    throw new Error(msg);
  }

  const lastCol = chosen.getLastColumn();
  if (!lastCol || lastCol < 1) {
    const msg = `[Contract] Sheet '${chosen.getName()}' has no columns (empty or missing header row).`;
    if (typeof logToSheet === 'function') logToSheet('ERROR', msg);
    throw new Error(msg);
  }

  const headers = chosen.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  const findCol = (aliases) => {
    for (const name of aliases) {
      const idx = headers.indexOf(name);
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const idx = {};
  const missing = [];
  Object.keys(opt.required).forEach(key => {
    const aliases = opt.required[key];
    const colIdx = findCol(aliases);
    idx[key] = colIdx;
    if (colIdx === -1) missing.push(`${key} (any of: ${aliases.join(', ')})`);
  });

  if (missing.length) {
    const msg = `[Contract] Missing required columns on '${chosen.getName()}': ${missing.join(' | ')}`;
    if (typeof logToSheet === 'function') logToSheet('ERROR', msg);
    throw new Error(msg);
  }

  // Optional sanity check for is_today
  let todayRow = null; // 1-based row number
  if (opt.requireTodayRow) {
    const lastRow = chosen.getLastRow();
    if (!lastRow || lastRow < 2) {
      const msg = `[Contract] Sheet '${chosen.getName()}' has no data rows (needs at least 2 rows).`;
      if (typeof logToSheet === 'function') logToSheet('ERROR', msg);
      throw new Error(msg);
    }

    const isTodayCol = chosen.getRange(2, idx.is_today + 1, lastRow - 1, 1).getValues();
    const hits = [];
    for (let i = 0; i < isTodayCol.length; i++) {
      const v = Number(String(isTodayCol[i][0]).replace(',', '.'));
      if (v === 1) hits.push(i + 2); // sheet row
    }

    if (!hits.length) {
      const msg = `[Contract] No is_today==1 row found in '${chosen.getName()}'.`;
      if (typeof logToSheet === 'function') logToSheet('ERROR', msg);
      throw new Error(msg);
    }

    if (hits.length > 1 && !opt.allowMultipleToday) {
      const msg = `[Contract] Multiple is_today==1 rows found in '${chosen.getName()}': ${hits.join(', ')}`;
      if (typeof logToSheet === 'function') logToSheet('ERROR', msg);
      throw new Error(msg);
    }

    todayRow = hits[hits.length - 1]; // last wins by default
  }

  return {
    timelineSheetName: chosen.getName(),
    timelineHeaders: headers,
    idx,
    todayRow
  };
}


/**
 * Fügt das Menü "Coach Kira" hinzu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Coach Kira')
    // MODIFIZIERT (V10)
    .addItem('KI-Planung & Charts starten (v10)', 'runKiraGeminiSupervisor')
    .addSeparator() // Trennlinie
    .addItem('Tageswechsel (is_today -> +1 Tag)', 'advanceTodayFlag') // <-- NEU
    .addSeparator() // <-- NEU
    .addItem('KI-PLAN ANWENDEN', 'applyKiraPlanToTimeline') // <-- NEU
    .addToUi();
}

function logToSheet(level, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);

    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME, ss.getSheets().length);
      logSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Level', 'Message']]).setFontWeight('bold');
      logSheet.setColumnWidth(1, 150); logSheet.setColumnWidth(2, 80); logSheet.setColumnWidth(3, 800);
      logSheet.setFrozenRows(1); // Friere die Kopfzeile
    }

    let logMessage = (typeof message === 'object') ? JSON.stringify(message) : String(message);
    if (logMessage.length > 45000) logMessage = logMessage.slice(0, 45000) + ' …(gekürzt)';

    // Neue Zeile unten anhängen; Platz schaffen, falls das Blatt voll ist
    const next = logSheet.getLastRow() + 1;
    if (next > logSheet.getMaxRows()) logSheet.insertRowsAfter(logSheet.getMaxRows(), 100);
    logSheet.getRange(next, 1, 1, 3).setValues([[new Date(), String(level || ''), logMessage]]);

    // CLEANER: ab 600 Einträgen auf die letzten 500 kürzen (in Blöcken, spart Laufzeit)
    const MAX_ROWS = 500;
    const lastRow = logSheet.getLastRow(); // inkl. Header
    if (lastRow > MAX_ROWS + 101) logSheet.deleteRows(2, lastRow - (MAX_ROWS + 1));

  } catch (e) {
    Logger.log(`Konnte nicht ins Log-Sheet schreiben: ${e.message}`);
    try {
      PropertiesService.getScriptProperties().setProperty('CK_LOG_LAST_ERROR',
        new Date().toISOString() + ' ' + String(e && e.message ? e.message : e).slice(0, 400));
    } catch (x) {}
  }
}


function copyTimelineData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Sheet Names (robust, ohne externe Konstanten nötig) ----
  const SOURCE_NAME = "timeline";
  const TARGET_NAME = "KK_TIMELINE";

  // ---- Zeit / Cutoff (HEUTE, 00:00) ----
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, tz, "yyyy-MM-dd");
  const todayMidnight = new Date(todayStr + "T00:00:00");

  const sourceSheet = ss.getSheetByName(SOURCE_NAME);
  const targetSheet = ss.getSheetByName(TARGET_NAME);

  if (!sourceSheet) throw new Error(`Quell-Sheet '${SOURCE_NAME}' nicht gefunden.`);
  if (!targetSheet) throw new Error(`Ziel-Sheet '${TARGET_NAME}' nicht gefunden.`);

  // ---- Read all data (einmal, schnell) ----
  const srcLastRow = sourceSheet.getLastRow();
  const srcLastCol = sourceSheet.getLastColumn();
  if (srcLastRow < 2 || srcLastCol < 1) {
    Logger.log(`[copyTimelineData] '${SOURCE_NAME}' ist leer.`);
    return;
  }

  const trgLastRow = targetSheet.getLastRow();
  const trgLastCol = targetSheet.getLastColumn();
  if (trgLastRow < 2 || trgLastCol < 1) {
    Logger.log(`[copyTimelineData] '${TARGET_NAME}' ist leer oder hat keinen Header.`);
  }

  const srcValues = sourceSheet.getRange(1, 1, srcLastRow, srcLastCol).getValues();
  const trgValues = targetSheet.getRange(1, 1, Math.max(trgLastRow, 1), trgLastCol).getValues();

  const srcHeader = srcValues[0].map(h => String(h || "").trim());
  const trgHeader = trgValues[0].map(h => String(h || "").trim());
  const srcHeaderLC = srcHeader.map(h => String(h || '').toLowerCase().trim());
  const trgHeaderLC = trgHeader.map(h => String(h || '').toLowerCase().trim());

  // ---- is_today Spalte in Source/Target finden (case-insensitive) ----
  const srcIsTodayCol = srcHeaderLC.indexOf('is_today');
  const trgIsTodayCol = trgHeaderLC.indexOf('is_today');

  // ---- Find date column in source & target (alias, case-insensitive) ----
  const findDateCol = (headerLC) => {
    const aliases = ['date','datum','day'];
    for (let i = 0; i < aliases.length; i++) {
      const idx = headerLC.indexOf(aliases[i]);
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const srcDateCol = findDateCol(srcHeaderLC);
  const trgDateCol = findDateCol(trgHeaderLC);

  // ---- Contract: KK_TIMELINE should be a 1:1 mirror of timeline (values only) ----
  // If headers/columns differ, abort (better than silently wrong sync).
  const sameShape = (srcLastCol === trgLastCol) && (srcHeaderLC.length === trgHeaderLC.length);
  const sameHeaders = sameShape && srcHeaderLC.every((h, i) => h === trgHeaderLC[i]);
  if (!sameHeaders) {
    const msg = `[copyTimelineData][Contract] Header mismatch timeline vs KK_TIMELINE (need 1:1). srcCols=${srcLastCol} trgCols=${trgLastCol}`;
    try { logToSheet('ERROR', msg); } catch(_) {}
    throw new Error(msg);
  }
  if (srcDateCol < 0 || trgDateCol < 0) {
    const msg = `[copyTimelineData][Contract] date/datum/day column not found in one of the sheets.`;
    try { logToSheet('ERROR', msg); } catch(_) {}
    throw new Error(msg);
  }

  // ---- Build target index by date (yyyy-MM-dd) for fast row mapping ----
  const toDateKey = (v) => {
    if (!v) return null;

    // If it's a Date object
    if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, tz, "yyyy-MM-dd");
    }

    // If it's a string like "2026-01-20" or "20.01.2026"
    const s = String(v).trim();
    if (!s) return null;

    // Try parse ISO
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

    // Try parse German dd.mm.yyyy
    const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (m) {
      const dd = m[1].padStart(2, "0");
      const mm = m[2].padStart(2, "0");
      const yyyy = m[3];
      return `${yyyy}-${mm}-${dd}`;
    }

    // Last resort: Date(s)
    const d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, "yyyy-MM-dd");

    return null;
  };

  const trgIndexByKey = new Map();
  for (let r = 1; r < trgValues.length; r++) {
    const key = toDateKey(trgValues[r][trgDateCol]);
    if (key) trgIndexByKey.set(key, r + 1); // sheet row number (1-based)
  }

  // Columns are 1:1 (validated above) – no name-based mapping needed.

  // ---- Collect rows >= today from source ----
  const rowsToWrite = []; // { targetRow, rowArrayForTarget }
  let matched = 0;
  let skippedNoTarget = 0;
  let skippedNoDate = 0;

  for (let r = 1; r < srcValues.length; r++) {
    const srcDateVal = srcValues[r][srcDateCol];
    const key = toDateKey(srcDateVal);
    if (!key) { skippedNoDate++; continue; }

    // compare with today
    const d = new Date(key + "T00:00:00");
    if (isNaN(d.getTime()) || d < todayMidnight) continue; // Vergangenheit ignorieren

    const targetRow = trgIndexByKey.get(key);
    if (!targetRow) { skippedNoTarget++; continue; } // Zukunft/Heute muss in KK_TIMELINE existieren

    // 1:1 mirror copy (values only)
    const out = srcValues[r].slice(0, trgLastCol);
    rowsToWrite.push({ targetRow, out });
    matched++;
  }

  // ---- Write back (batch per contiguous blocks) ----
  // Sort by targetRow so we can write in blocks
  rowsToWrite.sort((a, b) => a.targetRow - b.targetRow);

  let blocks = 0;
  let i = 0;
  while (i < rowsToWrite.length) {
    const startRow = rowsToWrite[i].targetRow;
    let endRow = startRow;
    const block = [rowsToWrite[i].out];

    i++;
    while (i < rowsToWrite.length && rowsToWrite[i].targetRow === endRow + 1) {
      endRow++;
      block.push(rowsToWrite[i].out);
      i++;
    }

    targetSheet.getRange(startRow, 1, block.length, trgLastCol).setValues(block);
    blocks++;
  }

    // ---- Fix: is_today darf nicht auf "gestern" stehen bleiben ----
  if (trgIsTodayCol >= 0) {
    // gestern berechnen
    const y = new Date(todayMidnight);
    y.setDate(y.getDate() - 1);
    const yesterdayKey = Utilities.formatDate(y, tz, "yyyy-MM-dd");

    const yesterdayRow = trgIndexByKey.get(yesterdayKey);
    const todayRow     = trgIndexByKey.get(todayStr);

    // Gestern -> 0
    if (yesterdayRow) {
      targetSheet.getRange(yesterdayRow, trgIsTodayCol + 1).setValue(0);
    }

    // Heute -> 1 (falls vorhanden)
    if (todayRow) {
      targetSheet.getRange(todayRow, trgIsTodayCol + 1).setValue(1);
    }
  }


  Logger.log(`[copyTimelineData] Done. today=${todayStr} | copiedRows=${matched} | writeBlocks=${blocks} | skippedNoTarget=${skippedNoTarget} | skippedNoDate=${skippedNoDate}`);
}


function getDateColIndex_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(h => String(h).trim().toLowerCase());

  const idxDate  = headers.indexOf("date");
  if (idxDate !== -1) return idxDate + 1;

  const idxDatum = headers.indexOf("datum");
  if (idxDatum !== -1) return idxDatum + 1;

  // Fallback: 1. Spalte
  return 1;
}

function findRowByDateFast_(sheet, dateCol, dateString) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const vals = sheet.getRange(2, dateCol, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === dateString) return i + 2; // +2 wegen Start ab Zeile 2
  }
  return null;
}


function syncTimelineTodayRow_() {
  logToSheet('INFO', `[V41-DELTA] Sync nur HEUTE: '${SOURCE_TIMELINE_SHEET}' -> '${TIMELINE_SHEET_NAME}'...`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_TIMELINE_SHEET);
  if (!src) throw new Error(`Quellblatt '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);
  const dst = ss.getSheetByName(TIMELINE_SHEET_NAME);
  if (!dst) { copyTimelineData(); return; }

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (dst.getLastRow() !== lastRow || dst.getLastColumn() !== lastCol) {
    // Struktur geändert -> Full-Copy als Fallback
    copyTimelineData(); return;
  }

  const headers = src.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const isTodayIdx = headers.indexOf('is_today');
  if (isTodayIdx < 0) { copyTimelineData(); return; }

  // nur is_today-Spalte lesen (schnell)
  const isTodayCol = src.getRange(2, isTodayIdx + 1, lastRow - 1, 1).getValues();
  let todayRow = -1;
  for (let i = 0; i < isTodayCol.length; i++) {
    if (Number(String(isTodayCol[i][0]).replace(',', '.')) === 1) { todayRow = i + 2; break; }
  }
  if (todayRow < 0) { copyTimelineData(); return; }

  // is_today-Spalte im Ziel aktualisieren (damit der alte "1"-Eintrag verschwindet)
  dst.getRange(2, isTodayIdx + 1, isTodayCol.length, 1).setValues(isTodayCol);

  // komplette HEUTE-Zeile aktualisieren
  const rowVals = src.getRange(todayRow, 1, 1, lastCol).getValues();
  dst.getRange(todayRow, 1, 1, lastCol).setValues(rowVals);

  logToSheet('INFO', `[V41-DELTA] OK. Heute-Zeile ${todayRow} aktualisiert.`);
}


/**
 * V137-LET-AI-SPEAK: Hauptsteuerung.
 * ÄNDERUNG: Der "Force Fix" am Ende überschreibt NICHT mehr den Text (Spalte F).
 * Die KI darf ihre Analyse schreiben. Wir korrigieren nur Zelle D11 (Formatierung) und C11 (Ampel).
 */
function runKiraGeminiSupervisor() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const HB_RUN_ID = Utilities.getUuid().slice(0, 8);
  if (!ckSup_begin_(HB_RUN_ID)) return; // läuft schon – Nachlauf ist vorgemerkt
  hb_(ss, HB_RUN_ID, "START", "Supervisor gestartet");



  // 1. Start-Signal
  try {
    try {
      sendTelegram("🚀 Commander, ich starte jetzt die Analyse-Sequenz...");
    } catch (e) {
      Utilities.sleep(1500);
      sendTelegram("🚀 Analyse läuft (Retry)...");
    }
  } catch(e) { console.warn("Telegram Fehler: " + e.message); }

  try {
    logToSheet('INFO', '🚀 Starte KI-Bewertung (Supervisor V137 - Let AI Speak)...');
    hb_(ss, HB_RUN_ID, "INIT", "Pre-Checks / TE-Balance");


    // ------------------------------------------------------------
    // SCHRITT 0: TE-BALANCE JAGD (FAST PATH)
    // ------------------------------------------------------------
    // Ziel: Keine Full-Sheet Reads. Wir lesen nur Header + is_today Spalte + 1 Zelle.
    // (flush ist hier in der Regel nicht nötig und kostet Zeit)

    let teResult = { val: 0, text: "0,0%", source: "NONE" };

    // VERSUCH A: KK_TIMELINE
    try {
      const sheet = ss.getSheetByName('KK_TIMELINE');
      if (sheet) {
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        if (lastRow >= 2 && lastCol >= 1) {
          const headers = sheet.getRange(1, 1, 1, lastCol)
            .getValues()[0]
            .map(h => String(h).toLowerCase().trim());

          const colIdxTE = headers.findIndex(h => h.includes('te_balance') || h.includes('te balance'));
          const colIdxToday = headers.indexOf('is_today');

          if (colIdxTE > -1 && colIdxToday > -1) {
            // Nur is_today lesen
            const isTodayCol = sheet.getRange(2, colIdxToday + 1, lastRow - 1, 1).getValues();
            let todayRow = -1;
            for (let i = isTodayCol.length - 1; i >= 0; i--) {
              if (Number(String(isTodayCol[i][0]).replace(',', '.')) === 1) { todayRow = i + 2; break; }
            }

            if (todayRow > 0) {
              const raw = sheet.getRange(todayRow, colIdxTE + 1).getValue();
              let val = (typeof raw === 'number') ? raw : parseGermanFloat_(String(raw).replace('%', ''));
              if (Number.isFinite(val)) {
                // Prozent-Strings: 9,3% -> 0.093
                if (String(raw).includes('%')) val = val / 100;
                if (val > 1) val = val / 100;
                teResult.val = val;
                teResult.text = (val * 100).toFixed(1).replace('.', ',') + "%";
                teResult.source = "TIMELINE";
              }
            }
          }
        }
      }
    } catch(e) {
      console.warn("Timeline Scan Error: " + e.message);
    }

    // VERSUCH B: FORECAST (Backup)
    if (teResult.source === "NONE" || teResult.val === 0) {
        try {
            const fcSheet = ss.getSheetByName('AI_DATA_FORECAST');
            if (fcSheet) {
                const headers = fcSheet.getRange(1, 1, 1, fcSheet.getLastColumn()).getValues()[0].map(h => h.toString().toLowerCase());
                const rowData = fcSheet.getRange(2, 1, 1, fcSheet.getLastColumn()).getValues()[0];
                const idx = headers.findIndex(h => h.includes('te balance') || h.includes('intensiv'));
                if (idx > -1 && rowData[idx]) {
                    let val = rowData[idx];
                    teResult.val = (typeof val === 'number') ? val : parseFloat(String(val).replace(',','.').replace('%',''))/100;
                    if (teResult.val > 1) teResult.val = teResult.val / 100;
                    teResult.text = (teResult.val * 100).toFixed(1).replace('.', ',') + "%";
                    teResult.source = "FORECAST";
                }
            }
        } catch(e) { }
    }
    
    logToSheet('INFO', `🔎 TE-Balance gefunden: ${teResult.text} (${teResult.source})`);

    // ------------------------------------------------------------
    // SCHRITT 1: DATEN LADEN
    // ------------------------------------------------------------
    //syncTimelineTodayRow_();
    copyTimelineData();
    const wetterDaten = getWetterdatenCached_();
    const datenPaket = getSheetData(); 
    if (datenPaket === null) return;

    /* ACWR Hotfix
    try {
      const fcSheet = ss.getSheetByName('AI_DATA_FORECAST');
      if (fcSheet) {
        const fcData = fcSheet.getRange(1, 1, 2, fcSheet.getLastColumn()).getValues();
        const vals = fcData[1];
        const hds = fcData[0].map(h => h.toString().toLowerCase());
        const idx = hds.indexOf('acwr prognose');
        if (idx > -1 && vals[idx]) datenPaket.heute['coachE_ACWR_forecast'] = vals[idx];
      }
    } catch (e) {} 
    */
    // ✅ ACWR: Source of Truth = KK_TIMELINE (coachE_ACWR_forecast)
// Nur fallback auf AI_DATA_FORECAST, wenn Timeline-Wert fehlt
if (datenPaket?.heute) {
  const tlRaw = datenPaket.heute['coachE_ACWR_forecast'];
  const tlNum = parseGermanFloat_(tlRaw);

  if (!isFinite(tlNum)) {
    const aiForecast = getSheetData("AI_DATA_FORECAST");
    const idx = aiForecast?.meta?.headerMap?.["acwr prognose"];
    const fb = (idx != null) ? aiForecast.rows?.[0]?.[idx] : null;

    if (fb != null && fb !== "") datenPaket.heute['coachE_ACWR_forecast'] = fb;
  }
}


    // ------------------------------------------------------------
    // SCHRITT 2: INJEKTION (WICHTIG FÜR KI)
    // ------------------------------------------------------------
    const fitnessScores = calculateFitnessMetrics(datenPaket);
    const nutritionScore = calculateNutritionScore(datenPaket);
    
    if (teResult.val > 0) {
        // A) KI Kontext: Wir geben der KI den Text ("10,6%")
        datenPaket.monotony_varianz = teResult.val;
        if (datenPaket.heute) {
            datenPaket.heute['te_balance_trend'] = teResult.text; 
            datenPaket.heute['TE_Balance_Trend'] = teResult.text;
        }

        // B) Score Array: Damit calculateGesamtScore stimmt
        const teMetric = fitnessScores.find(item => item && item.metrik && (item.metrik.includes("TE Balance") || item.metrik.includes("Intensiv")));
        if (teMetric) {
            teMetric.raw_wert = teResult.text; 
            teMetric.value = teResult.val;
            // Ampel setzen
            if (teResult.val >= 0.2) { teMetric.num_score = 100; teMetric.ampel = "LILA"; }
            else if (teResult.val >= 0.1) { teMetric.num_score = 80; teMetric.ampel = "GRÜN"; }
            else { teMetric.num_score = 25; teMetric.ampel = "ROT"; }
        }
    }

    // ------------------------------------------------------------
    // SCHRITT 3: SCORE & KI
    // ------------------------------------------------------------
    const alleScores = fitnessScores.concat([nutritionScore]);
    const subScores = calculateSubScores(fitnessScores);
    const scoreResult = calculateGesamtScore(alleScores);
    
    // KI-frei (Stufe 3): nur Daten – den Prompt baut der Mac (kira_prompts)
    const promptData = ckSup_promptData_(
      datenPaket,
      fitnessScores, 
      nutritionScore,
      scoreResult.num, 
      scoreResult.ampel,
      getLoadFactors(), 
      getElevFactors(), 
      getWeekConfig(), 
      (parseGermanFloat(datenPaket.heute.load_fb_day) > 0),
      wetterDaten
    );

    // Lokal-Modus: Auftrag an den Mac geben und hier beenden (Ergebnis kommt per saveSupervisor)
    const ckSupCtx = { fitnessScores: fitnessScores, nutritionScore: nutritionScore, scoreResult: scoreResult, subScores: subScores, teResult: teResult };
    // KI-frei (2026-10-26): Der Auftrag geht immer an den Mac. Das Ergebnis kommt per saveSupervisor zurück.
    ckSup_enqueue_(promptData, ckSupCtx, HB_RUN_ID);
    hb_(ss, HB_RUN_ID, "AI_WAIT_LOCAL", "Wartet auf Mac (coach-kira)");
    logToSheet('INFO', '[Supervisor] Auftrag für den Mac bereitgestellt.');
    ckSup_end_();
    return;




  } catch (e) {
    hb_(ss, HB_RUN_ID, "END_ERROR", e && e.message ? e.message : String(e));
    logToSheet('ERROR', `🛑 FEHLER im Supervisor: ${e.message}`);
    logToSheet('ERROR', `Stack: ${e.stack}`);
    ckSup_end_();
  }
}



/**
 * Schedules heavy post-processing work to run asynchronously.
 * Keeps `runKiraGeminiSupervisor()` fast and less likely to hit execution limits.
 */
function schedulePostWork_() {
  try {
    // Remove older pending post-work triggers to avoid piling up.
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
      try {
        if (t.getHandlerFunction && t.getHandlerFunction() === 'runPostWork_') {
          ScriptApp.deleteTrigger(t);
        }
      } catch (_) {}
    });

    // Schedule a single run shortly after the supervisor finishes.
    ScriptApp.newTrigger('runPostWork_')
      .timeBased()
      .after(15 * 1000)
      .create();

    logToSheet('INFO', '[PostWork] Trigger scheduled (runPostWork_ in ~15s)');
  } catch (e) {
    // If trigger creation fails, we keep the supervisor successful.
    logToSheet('WARN', `[PostWork] Could not schedule trigger: ${e.message}`);
  }
}

/**
 * Runs heavy post-processing tasks outside the main supervisor execution.
 */
function runPostWork_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    logToSheet('INFO', '[PostWork] Start');

    try {
      logToSheet('INFO', '[PostWork] updateHistoryWithRIS()');
      updateHistoryWithRIS();
    } catch (e) {
      logToSheet('WARN', `[PostWork] updateHistoryWithRIS failed: ${e.message}`);
    }

    try {
      logToSheet('INFO', '[PostWork] exportLookerChartsData()');
      exportLookerChartsData();
    } catch (e) {
      logToSheet('WARN', `[PostWork] exportLookerChartsData failed: ${e.message}`);
    }

    logToSheet('INFO', '[PostWork] Done');
  } finally {
    // Clean up our trigger so it won't repeat.
    try {
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(t => {
        if (t.getHandlerFunction && t.getHandlerFunction() === 'runPostWork_') {
          ScriptApp.deleteTrigger(t);
        }
      });
    } catch (_) {}
  }
}


function normalizeScore(value, optimalValue, badValue) {
  const v  = Number(value);
  const opt = Number(optimalValue);
  const bad = Number(badValue);

  // Wenn irgendwas nicht numerisch ist: Score = 0 (verhindert #NUM!)
  if (!Number.isFinite(v) || !Number.isFinite(opt) || !Number.isFinite(bad)) return 0;

  let score;
  // "niedriger ist besser"
  if (opt < bad) {
    score = ((bad - v) / (bad - opt)) * 100;
  } else {
    // "höher ist besser"
    score = ((v - bad) / (opt - bad)) * 100;
  }

  if (!Number.isFinite(score)) return 0;

  score = Math.max(0, Math.min(100, score));
  return Math.round(score);
}


function getBaselineNumber_(baselineObj, aliases) {
  const b = baselineObj || {};
  const keys = Object.keys(b);

  const norm = (x) => String(x || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^a-z0-9]/g, '');

  const aliasNorm = (aliases || []).map(norm);

  // 1) Exakt über normalisierte Alias-Matches
  for (let i = 0; i < keys.length; i++) {
    const kNorm = norm(keys[i]);
    if (!kNorm) continue;
    if (aliasNorm.indexOf(kNorm) !== -1) {
      const n = parseGermanFloat_(b[keys[i]]);
      if (Number.isFinite(n)) return n;
    }
  }

  // 2) Fuzzy Fallback: enthält Alias oder wird vom Alias enthalten
  for (let i = 0; i < keys.length; i++) {
    const kNorm = norm(keys[i]);
    if (!kNorm) continue;
    for (let a = 0; a < aliasNorm.length; a++) {
      const wanted = aliasNorm[a];
      if (!wanted) continue;
      if (kNorm.indexOf(wanted) !== -1 || wanted.indexOf(kNorm) !== -1) {
        const n = parseGermanFloat_(b[keys[i]]);
        if (Number.isFinite(n)) return n;
      }
    }
  }

  return NaN;
}


/**
 * NEU (V125): Garmin-Standard-Skala für ALLE Scores (0-100).
 * 95-100: VIOLETT (Prime/Höchstform)
 * 75-94:  BLAU    (Gut/Hoch)
 * 50-74:  GRÜN    (Mäßig/Basis)
 * 25-49:  ORANGE  (Niedrig/Vorsicht)
 * 0-24:   ROT     (Schlecht/Kritisch)
 */
function getAmpelFromScore(score) {
  // Sicherheits-Cast auf Zahl
  const val = parseInt(score);
  if (isNaN(val)) return "GRAU";

  if (val >= 95) return "LILA";   // War vorher "PRIME" oder "VIOLETT"
  if (val >= 75) return "BLAU";
  if (val >= 50) return "GRÜN";
  if (val >= 25) return "ORANGE";
  return "ROT";
}

/**
 * NEU (V13): Spezielle Normalisierung für ACWR (Optimalbereich).
 */
function normalizeAcwrScore(acwr) {
  const OPTIMAL_MIN = 0.9;
  const OPTIMAL_MAX = 1.1;
  const BAD_LOW = 0.7;
  const BAD_HIGH = 1.3; // Alter ROT-Wert war 1.2, neuer ist 1.3 für 0 Punkte

  if (acwr >= OPTIMAL_MIN && acwr <= OPTIMAL_MAX) {
    return 100; // Optimal range
  }
  if (acwr > OPTIMAL_MAX) { // Zu hoch
    // Linearer Abfall von 100 bei 1.1 auf 0 bei 1.3
    return normalizeScore(acwr, OPTIMAL_MAX, BAD_HIGH);
  }
  if (acwr < OPTIMAL_MIN) { // Zu niedrig
    // Linearer Anstieg von 0 bei 0.7 auf 100 bei 0.9
    return normalizeScore(acwr, OPTIMAL_MIN, BAD_LOW);
  }
  return 0; // Sollte nicht erreicht werden
}

/**
 * NEU (V13): Spezielle Normalisierung für Ernährung (Optimalbereich).
 * Wert ist das Defizit. Positiv = Defizit, Negativ = Überschuss.
 */
function normalizeNutritionScore(deficit) {
  const OPTIMAL_MIN = -300; // Optimaler Überschuss (alt: -300)
  const OPTIMAL_MAX = 200;  // Optimales Defizit (alt: 200)
  const BAD_LOW = -700;     // Zu viel Überschuss (0 Punkte)
  const BAD_HIGH = 600;     // Zu viel Defizit (alt: 500)

  if (deficit >= OPTIMAL_MIN && deficit <= OPTIMAL_MAX) {
    return 100; // Optimal range
  }
  if (deficit > OPTIMAL_MAX) { // Zu hohes Defizit
    return normalizeScore(deficit, OPTIMAL_MAX, BAD_HIGH);
  }
  if (deficit < OPTIMAL_MIN) { // Zu hoher Überschuss
    return normalizeScore(deficit, OPTIMAL_MIN, BAD_LOW);
  }
  return 0;
}

/**
 * V116-MOD: Score für 28-Tage Prozent-Anteil.
 * UPDATE: Progressive Kurve (Exponent 1.7).
 * - Ziel: 25% bis 50% = 100 Punkte.
 * - 20.5% ergibt jetzt ca. 71 Punkte (statt 82).
 */
function normalizeVarianzScore(ratio) {
  // Ideal-Korridor: 25% - 50% -> Volle Punktzahl
  if (ratio >= 0.25 && ratio <= 0.50) {
      return 100; 
  }
  
  // Zu wenig (<25%)
  if (ratio < 0.25) {
      // Berechnung des Fortschritts (0.0 bis 1.0)
      let progress = ratio / 0.25;
      
      // FORMEL: Fortschritt hoch 1.7 * 100
      // Das erzeugt eine Kurve, die "streng" ist.
      // Beispiel 20.5%: (0.82)^1.7 = 0.71 -> Score 71
      // Beispiel 12.5%: (0.50)^1.7 = 0.30 -> Score 30
      return Math.round(Math.pow(progress, 1.7) * 100);
  }
  
  // Zu viel (>50%)
  if (ratio > 0.50) {
      // Ab 70% ist es Score 0 (Rot/Warnung)
      let score = 100 - ((ratio - 0.50) / 0.20) * 100;
      return Math.max(0, Math.round(score));
  }
  
  return 0;
}

/**
 * V117-MOD: Berechnet den "Belastungsfokus" (28 Tage).
 * Nutzt die globale Konstante TE_LIMIT_HIGH_AEROBIC für die Feinjustierung.
 */
function calculateRollingTEVarianz(allRawData, currentRowIndex, lookbackDays, teIndices) {
    if (teIndices.aerobic_te === -1 || teIndices.anaerobic_te === -1 || teIndices.load_index === -1) {
        return 0.0;
    }

    let sumBaseLoad = 0;
    let sumIntenseLoad = 0;
    const startIndex = Math.max(1, currentRowIndex - (lookbackDays - 1));

    for (let i = startIndex; i <= currentRowIndex; i++) {
        if (!allRawData[i]) continue;
        const rowRaw = allRawData[i];
        
        const aerobicTE = parseGermanFloat(rowRaw[teIndices.aerobic_te]);
        const anaerobicTE = parseGermanFloat(rowRaw[teIndices.anaerobic_te]);
        const dailyLoad = parseGermanFloat(rowRaw[teIndices.load_index]);

        if (!dailyLoad || dailyLoad === 0) continue;
        let isIntenseDay = false;

        // 1. Anaerob ist immer intensiv (>= 2.0)
        if (typeof anaerobicTE === 'number' && anaerobicTE >= TE_LIMIT_ANAEROBIC) { 
     isIntenseDay = true;
}
        // 2. Aerob: Nutzt jetzt die GLOBALE Einstellung von oben
        else if (typeof aerobicTE === 'number' && aerobicTE >= TE_LIMIT_HIGH_AEROBIC) {
             isIntenseDay = true;
        }

        if (isIntenseDay) {
            sumIntenseLoad += dailyLoad;
        } else {
            sumBaseLoad += dailyLoad;
        }
    }

    const totalLoad = sumBaseLoad + sumIntenseLoad;
    if (totalLoad > 0) {
        return sumIntenseLoad / totalLoad;
    } else {
        return 0.0;
    }
}

/**
 * V160-ULTRA-ROBUST: Fix für HRV-Crash & Smart-Gains-Read.
 * FIX 1: HRV Thresholds werden sicher als String behandelt, um .includes() Fehler zu vermeiden.
 * FIX 2: KEI priorisiert den heutigen Timeline-Wert (coache/coachE_smart_gains)
 * und nutzt Forecast nur als datumsvalidierten Fallback.
 */
function calculateFitnessMetrics(datenPaket) {
  const { heute, baseline, monotony_varianz } = datenPaket;
  const scores = [];
  const GRUEN = 100, GELB = 50, ROT = 0;

  // --- Grenzwerte ---
  const RHR_OPTIMAL_DELTA = 2; const RHR_BAD_DELTA = 7;        
  const SLEEP_H_OPTIMAL = 7.5; const SLEEP_H_BAD = 6.0;   
  const SLEEP_HOURS_OPTIMAL = SLEEP_H_OPTIMAL;
const SLEEP_HOURS_BAD = SLEEP_H_BAD;     
  const STRAIN_OPTIMAL = 3000; const STRAIN_BAD = 7000;        
  const MONO_OPTIMAL = 1.4;    const MONO_BAD = 2.1;

  // 1. RHR
  const rhr_heute = parseGermanFloat_(heute['rhr_bpm']);
const rhr_default = getBaselineNumber_(baseline, ['RHR_default (bpm)', 'RHR_default(bpm)', 'rhr_default (bpm)', 'rhr_default_bpm', 'rhr_default']);

const rhr_delta = (Number.isFinite(rhr_heute) && Number.isFinite(rhr_default))
  ? (rhr_heute - rhr_default)
  : 0;
const rhr_score = Number.isFinite(rhr_heute)
  ? normalizeScore(rhr_delta, RHR_OPTIMAL_DELTA, RHR_BAD_DELTA)
  : 0;

const rhr_display = Number.isFinite(rhr_heute) ? String(Math.round(rhr_heute)) : "—";

  scores.push({ metrik: "RHR", raw_wert: `${rhr_display} bpm`, num_score: rhr_score, ampel: getAmpelFromScore(rhr_score) });


  // 2. Schlafdauer
  const sleep_hours_raw = parseGermanFloat_(heute['sleep_hours']);       // echte Zahl aus KK_TIMELINE
const sleep_hours = Number.isFinite(sleep_hours_raw)
  ? Math.floor(sleep_hours_raw * 10) / 10                              // TRUNC auf 1 Dezimalstelle
  : NaN;

const sleep_h_score = Number.isFinite(sleep_hours_raw)
  ? normalizeScore(sleep_hours_raw, SLEEP_H_OPTIMAL, SLEEP_H_BAD)
  : 0;


const sleep_display = Number.isFinite(sleep_hours)
  ? sleep_hours.toFixed(1).replace('.', ',')
  : "—";

  scores.push({ metrik: "Schlafdauer", raw_wert: `${sleep_display}h`, num_score: sleep_h_score, ampel: getAmpelFromScore(sleep_h_score) });

  // 3. Schlafscore
  const sleep_s_score = Math.round(parseGermanFloat(heute['sleep_score_0_100']));
  scores.push({ metrik: "Schlafscore", raw_wert: `${sleep_s_score}/100`, num_score: sleep_s_score, ampel: getAmpelFromScore(sleep_s_score) });

  // --- 4. HRV STATUS (V160 FIX) ---
  let hrv_text = "N/A"; let hrv_score = ROT; let hrv_ampel = "ROT";
  const hrv_val = parseGermanFloat(heute['hrv_status']); 
  const hrv_thresh = String(heute['hrv_threshholds'] || ""); 

  if (!isNaN(hrv_val) && hrv_thresh.includes(';')) {
      const parts = hrv_thresh.split(';');
      const min = parseFloat(parts[0].replace(',', '.'));
      const max = parseFloat(parts[1].replace(',', '.'));
      hrv_text = `${hrv_val} ms`; 
      if (hrv_val >= min && hrv_val <= max) { hrv_score = GRUEN; hrv_ampel = "GRÜN"; } 
      else { 
          const dist = (hrv_val < min) ? (min - hrv_val) : (hrv_val - max);
          if (dist <= 2) { hrv_score = GELB; hrv_ampel = "GELB"; } else { hrv_score = ROT; hrv_ampel = "ROT"; }
      }
  } else {
      hrv_text = (heute['hrv_status'] || "Unbekannt").toString().trim();
      if (hrv_text === "Ausgeglichen") { hrv_score = GRUEN; hrv_ampel = "GRÜN"; }
      else if (hrv_text === "Unausgeglichen") { hrv_score = GELB; hrv_ampel = "GELB"; }
      else { hrv_score = ROT; hrv_ampel = "ROT"; }
  }
  scores.push({ metrik: "HRV Status", raw_wert: hrv_text, num_score: hrv_score, ampel: hrv_ampel });

  // 5. Readiness (Fix: Case-Insensitive & Undefined-Protection)
  // Wir suchen sowohl klein- als auch großgeschriebene Keys
  const rawTR = heute['fb_tr_obs'] || heute['fb_TR_obs'] || 
                heute['garmin_training_readiness'] || heute['Garmin_Training_Readiness'];
  
  const readiness = parseGermanFloat(rawTR);
  
  // Validierung: Wenn keine Zahl da ist, setzen wir Standardwerte statt 'undefined'
  const hasReadiness = (typeof readiness === 'number' && !isNaN(readiness));
  const ready_score = hasReadiness ? Math.round(readiness) : 0;
  const ready_text = hasReadiness ? readiness.toString() : "--";
  const ready_ampel = hasReadiness ? getAmpelFromScore(ready_score) : "GRAU";

  scores.push({ 
    metrik: "Training Readiness", 
    raw_wert: ready_text, 
    num_score: ready_score, 
    ampel: ready_ampel 
  });

  // 6. Strain/Monotony
  const strain7 = parseGermanFloat(heute['strain7']);
  
  // 7. TE Balance (FIX: Type-Safety für Zahlen & Strings)
  let teVal = 0;
  let teText = "0,0%";
  
  // A) Versuch: Wert aus 'te_balance_trend' (aus Schritt 1)
  if (datenPaket.heute && datenPaket.heute['te_balance_trend'] !== undefined) {
      let raw = datenPaket.heute['te_balance_trend'];
      
      if (typeof raw === 'number') {
          // CASE 1: Es ist bereits eine Zahl (z.B. 0.107 oder 10.7)
          // Automatische Erkennung ob Dezimal (0.10) oder Prozent (10.0)
          if (Math.abs(raw) <= 1.0 && raw !== 0) { 
              raw = raw * 100; 
          }
          teVal = raw;
          teText = raw.toFixed(1).replace('.', ',') + "%";
          
      } else {
          // CASE 2: Es ist ein String (z.B. "10,7%" oder "0,107")
          teText = String(raw); // Erzwinge String für .replace
          teVal = parseGermanFloat(teText.replace('%',''));
          
          // Auch hier: Dezimal-Check
          if (Math.abs(teVal) <= 1.0 && teVal !== 0) { 
              teVal = teVal * 100; 
              teText = teVal.toFixed(1).replace('.', ',') + "%";
          }
      }
  } else {
      // B) Fallback: Forecast Sheet direkt lesen (falls Schritt 1 leer war)
      try {
        const fcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AI_DATA_FORECAST');
        if (fcSheet) {
          const fcData = fcSheet.getDataRange().getValues();
          const fcHeaders = fcData[0].map(h => h.toString().trim().toLowerCase());
          const colIdx = fcHeaders.indexOf('te balance');
          
          if (colIdx !== -1 && fcData.length > 1) {
             let rawFC = fcData[1][colIdx]; 
             if (typeof rawFC === 'number') {
                 if (rawFC <= 1.0 && rawFC > 0) { rawFC = rawFC * 100; } 
                 teVal = rawFC;
                 teText = rawFC.toFixed(1).replace('.', ',') + "%";
             } else {
                 teText = String(rawFC);
                 teVal = parseGermanFloat(teText.replace('%',''));
                 if (teVal <= 1.0 && teVal > 0) { teVal = teVal * 100; }
                 teText = teVal.toFixed(1).replace('.', ',') + "%";
             }
          }
        }
      } catch(e) { 
        Logger.log("Fehler TE Balance Read: " + e.message); 
      }
  }

  // Bewertung (Ampel)
  // < 10% = Zu wenig Intensität (oder Base Phase) -> GELB/ROT
  // 10-25% = Optimal -> GRÜN
  // > 25% = Hoch -> GELB
  // > 40% = Zu viel -> ROT
  let teScore = 0;
  let teAmpel = "ROT";
  
  if (teVal > 40) { teScore = 20; teAmpel = "ROT"; }
  else if (teVal > 25) { teScore = 60; teAmpel = "GELB"; }
  else if (teVal >= 10) { teScore = 100; teAmpel = "GRÜN"; } // Optimal
  else if (teVal >= 5) { teScore = 70; teAmpel = "GELB"; }   // Base Phase ok
  else { teScore = 40; teAmpel = "ROT"; }                    // Zu wenig

  scores.push({ 
    metrik: "TE Balance (% Intensiv)", 
    raw_wert: teText, 
    num_score: teScore, 
    ampel: teAmpel 
  });

  // 8. ACWR
  // Wir lesen den Wert robust ein
  const acwrRaw = heute['coachE_ACWR_forecast'] ?? heute['coache_acwr_forecast']; // einzig wahre Quelle
const acwrValNum = parseGermanFloat_(acwrRaw);

if (!Number.isFinite(acwrValNum)) {
  logToSheet('WARN', `[ACWR] coachE_ACWR_forecast ist leer/ungueltig: "${acwrRaw}"`);
}

const acwr_score = normalizeAcwrScore(Number.isFinite(acwrValNum) ? acwrValNum : 0);
const displayAcwrSheet = Number.isFinite(acwrValNum) ? acwrValNum.toFixed(2).replace('.', ',') : "—";

  
  // WICHTIG: Für die Anzeige im Sheet nehmen wir Deutsch (Komma)
  // const displayAcwrSheet = acwrVal.toFixed(2).replace('.', ',');
  
  // TRICK: Wir speichern den Wert auch als rohe Zahl oder englisches Format für die KI,
  // falls die KI direkt auf 'raw_wert' zugreift. 
  // Da 'calculateFitnessMetrics' aber meist für das Sheet ist, tricksen wir hier:
  // Wir lassen 'raw_wert' deutsch, aber stellen sicher, dass im Prompt später (.) genutzt wird.
  
  scores.push({ 
    metrik: "ACWR (Forecast)", 
    raw_wert: displayAcwrSheet, // Anzeige: "1,30"
    num_score: acwr_score, 
    ampel: getAmpelFromScore(acwr_score),
    // Zusatzfeld für KI-Prompt (falls dein Prompt-Generator das nutzen kann):
    value_for_ai: Number.isFinite(acwrValNum) ? acwrValNum.toFixed(2) : ""
  });

  // 9. Training Status
  const statusText = (heute['trainingszustand'] || "Kein Status").toString().trim(); 
  let status_score = 50; 
  if(statusText == "Höchstform") status_score = 100;
  if(statusText == "Formaufbau") status_score = 90;
  if(statusText == "Formerhalt") status_score = 80;
  if(statusText == "Erholung") status_score = 70;
  if(statusText == "Unproduktiv") status_score = 50;
  if(statusText == "Formverlust") status_score = 40;
  if(statusText == "Ermüdet") status_score = 30;
  if(statusText == "Überbelastung") status_score = 0;
  scores.push({ metrik: "Training Status", raw_wert: statusText, num_score: status_score, ampel: (status_score < 50 ? "ROT" : (status_score == 50 ? "NEUTRAL" : "GRÜN")) });

  // 10. Protein
  const proteinRaw = parseGermanFloat(heute['protein_g']) || 0;
  scores.push({ metrik: "Protein-Invest", raw_wert: `${proteinRaw}g`, num_score: normalizeProteinScore(proteinRaw/76), ampel: proteinRaw < 10 ? "OFFEN" : getAmpelFromScore(normalizeProteinScore(proteinRaw/76)) });

  // --- 11. KEI (UPDATED: AGGRESSIVE SKALA) ---
  let smartRaw = NaN;
  let smartSource = "Timeline";

  // Primärquelle: heute / Timeline (robuste Alias-Prüfung)
  const smartPrimaryCandidates = [
    heute['coache_smart_gains'],
    heute['coachE_smart_gains'],
    heute['coache_smart_gains_v3']
  ];

  for (let i = 0; i < smartPrimaryCandidates.length; i++) {
    const parsed = parseGermanFloat(smartPrimaryCandidates[i]);
    if (Number.isFinite(parsed)) {
      smartRaw = parsed;
      break;
    }
  }

  // Sekundärquelle: Forecast nur als Fallback, und nur mit Heute-Validierung
  if (!Number.isFinite(smartRaw)) {
    smartSource = "Forecast";
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const fcSheet = ss.getSheetByName('AI_DATA_FORECAST');
      const kkSheet = ss.getSheetByName(TIMELINE_SHEET_NAME) || ss.getSheetByName(SOURCE_TIMELINE_SHEET);

      const normalizeDateKey = (val) => {
        if (val == null || val === "") return "";
        const tz = Session.getScriptTimeZone();
        const d = (val instanceof Date) ? val : new Date(val);
        if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        const s = String(val).trim();
        const m = s.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$/);
        if (m) {
          const yyyy = m[3].length === 2 ? `20${m[3]}` : m[3];
          return `${yyyy}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
        }
        return s;
      };

      const todayKey = normalizeDateKey(heute['date']);
      let forecastTodayKey = "";

      if (fcSheet && fcSheet.getLastRow() > 1 && fcSheet.getLastColumn() > 0) {
        const fcData = fcSheet.getDataRange().getValues();
        const fcHeaders = fcData[0].map(h => String(h).trim().toLowerCase());
        const idxDate = fcHeaders.indexOf('datum') > -1 ? fcHeaders.indexOf('datum') : fcHeaders.indexOf('date');
        const idxKei = fcHeaders.indexOf('kei forecast') > -1
          ? fcHeaders.indexOf('kei forecast')
          : fcHeaders.indexOf('smart gain forecast');

        let todayForecastRow = null;
        if (idxDate !== -1 && todayKey) {
          todayForecastRow = fcData.slice(1).find(r => normalizeDateKey(r[idxDate]) === todayKey) || null;
          if (todayForecastRow) forecastTodayKey = normalizeDateKey(todayForecastRow[idxDate]);
        }

        if (!todayForecastRow && kkSheet && kkSheet.getLastRow() > 1) {
          const kkData = kkSheet.getDataRange().getValues();
          const kkHeaders = kkData[0].map(h => String(h).trim().toLowerCase());
          const idxIsToday = kkHeaders.indexOf('is_today');
          const idxKkDate = kkHeaders.indexOf('date') > -1 ? kkHeaders.indexOf('date') : kkHeaders.indexOf('datum');
          if (idxIsToday !== -1) {
            const todayRowKk = kkData.slice(1).find(r => Number(r[idxIsToday]) === 1 || String(r[idxIsToday]).trim() === '1');
            if (todayRowKk && idxKkDate !== -1) {
              const kkTodayKey = normalizeDateKey(todayRowKk[idxKkDate]);
              if (idxDate !== -1) {
                todayForecastRow = fcData.slice(1).find(r => normalizeDateKey(r[idxDate]) === kkTodayKey) || null;
              }
              if (todayForecastRow) forecastTodayKey = normalizeDateKey(todayForecastRow[idxDate]);
            }
          }
        }

        if (idxKei !== -1 && todayForecastRow && todayForecastRow[idxKei] !== "") {
          const parsedFallback = parseGermanFloat(todayForecastRow[idxKei]);
          if (Number.isFinite(parsedFallback)) {
            smartRaw = parsedFallback;
          }
        }
      }

      if (!Number.isFinite(smartRaw)) {
        smartRaw = 0;
        smartSource = "Fallback0";
      }

      if (smartSource === "Forecast") {
        logToSheet('INFO', `[KEI] Forecast-Fallback genutzt (heute=${todayKey || 'n/a'}, forecast=${forecastTodayKey || 'n/a'}).`);
      }
    } catch(e) {
      smartRaw = 0;
      smartSource = "Fallback0";
      logToSheet('WARN', `[KEI] Forecast-Fallback fehlgeschlagen: ${e.message}`);
    }
  }
  
  // --- NEUE BEWERTUNG ---
let sgScore = 0;
let sgAmpel = "ROT";

// KEI Legende (neu):
// High Efficiency (>= 8), Productive (3 bis < 8), Maintenance/Transition (0 bis < 3)
// Werte < 0 = ineffizient/Überlast-Risiko
const keiAssessment = getKEIAssessment(smartRaw);
sgScore = keiAssessment.score;
sgAmpel = keiAssessment.ampel;
  
  let smartText = (smartRaw > 0 ? "+" : "") + smartRaw.toFixed(2).replace('.', ',');
  
  scores.push({ 
      metrik: "KEI", 
      raw_wert: smartText, 
      num_score: sgScore, 
      ampel: sgAmpel,
      source: smartSource
  });
  
  return scores;
}

/**
 * (V13 - GRANULAR SCORE): Berechnet den Score für die 7-Tage-Bilanz.
 * Positive Werte sind ein DEFIZIT.
 * Negative Werte sind ein ÜBERSCHUSS.
 */
function calculateNutritionScore(datenPaket) {
  const { ernaehrung } = datenPaket;
  const avg_val = ernaehrung.avg_7d_deficit; 
  
  // NEU: Eigene Helper-Funktion für Optimalbereich
  const ern_score = normalizeNutritionScore(avg_val);
  
  // NEU: Dynamischer Ampel-Text basierend auf dem Rohwert
  let ern_ampel_text;
  if (ern_score >= 80) {
    ern_ampel_text = "OK";
  } else if (avg_val > 200) { // 200 ist der OPTIMAL_MAX Wert
    ern_ampel_text = "DEFIZIT";
  } else if (avg_val < -300) { // -300 ist der OPTIMAL_MIN Wert
    ern_ampel_text = "ÜBERSCHUSS";
  } else {
    ern_ampel_text = getAmpelFromScore(ern_score); // Fallback (GELB/ROT)
  }
  
  return { metrik: "7-Tage-Bilanz", raw_wert: avg_val.toFixed(0), num_score: ern_score, ampel: ern_ampel_text };
}


/**
 * Übersetzt den numerischen Gesamtscore in eine Ampelfarbe.
 */
function getGesamtAmpel(score) {
  if (score < 50) return "ROT"; 
  if (score < 80) return "GELB"; 
  return "GRÜN"; 
}

/**
 * V105-WISSENSSPEICHER:
 * Enthält die wichtigsten Definitionen aus den Firstbeat Whitepapers 
 * als "Quelle der Wahrheit" für die KI-Analysen.
 */
// KI-frei (Stufe 3): WISSENSBLOCK_V105 entfernt – liegt in kira_prompts (Fallback, falls KK_WISSEN fehlt).

// KI-frei (Stufe 3): createMasterGeminiPrompt entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

// KI-frei (Stufe 3): buildGeminiPrompt entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

/**
 * V154-PERCENT-FIX: Speicher-Funktion.
 * FIX: Erzwingt das Prozent-Format für "TE Balance" und andere %-Werte,
 * nachdem die globale Formatierung durchgelaufen ist.
 */
function writeOutputToSheets(ss, aiData, fitnessScores, nutritionScore, gesamtScoreNum, gesamtScoreAmpel, recoveryScore, recoveryAmpel, trainingScore, trainingAmpel) {
  
  logToSheet('INFO', '[V154] Speichere Daten (Percent-Fix + App Sync)...');

  // --- 1. JSON PARSEN ---
  let data = {};
  let inputStr = (typeof aiData === 'string') ? aiData : JSON.stringify(aiData);
  try {
    const cleanJson = inputStr.replace(/```json/g, "").replace(/```/g, "").trim();
    try {
        data = JSON.parse(cleanJson);
    } catch(e1) {
        const firstBrace = cleanJson.indexOf('{');
        const lastBrace = cleanJson.lastIndexOf('}');
        if (firstBrace !== -1 && lastBrace !== -1) {
            data = JSON.parse(cleanJson.substring(firstBrace, lastBrace + 1));
        } else { throw e1; }
    }
  } catch (e) {
    logToSheet('ERROR', "JSON Parse Fehler: " + e.message);
    data = {}; 
  }

  // -------------------------------------------------------
  // BLATT 1: AI_REPORT_STATUS (Smart Update)
  // -------------------------------------------------------
  try {
      let statusSheet = ss.getSheetByName('AI_REPORT_STATUS');
      if (!statusSheet) {
          statusSheet = ss.insertSheet('AI_REPORT_STATUS');
          const headers = ["Kategorie", "Metrik", "Status_Ampel", "Status_Wert_Num", "Status_Score_100", "Text_Info"];
          statusSheet.appendRow(headers);
          statusSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      }

      // A) Alle Daten sammeln
      const updates = [];

      // Gesamtscore
      let finalText = (data.empfehlung_zukunft && data.empfehlung_zukunft.text) ? data.empfehlung_zukunft.text : "Keine Analyse.";
      const safeGesamtNum = (gesamtScoreNum === undefined || gesamtScoreNum === null) ? 0 : gesamtScoreNum;
      updates.push({metrik: "Gesamtscore", cat: "Gesamt", ampel: gesamtScoreAmpel || "GRAU", raw: safeGesamtNum, score: safeGesamtNum, text: finalText});

      // Plan Status
      const finalPlanStatus = (data.empfehlung_zukunft && data.empfehlung_zukunft.plan_status) ? data.empfehlung_zukunft.plan_status : "N/A";
      updates.push({metrik: "Plan Status", cat: "Gesamt", ampel: "GRAU", raw: "", score: "", text: finalPlanStatus});

      // Sub-Scores
      updates.push({metrik: "Recovery Score", cat: "Gesamt", ampel: recoveryAmpel, raw: recoveryScore, score: recoveryScore, text: ""});
      updates.push({metrik: "Training Score", cat: "Gesamt", ampel: trainingAmpel, raw: trainingScore, score: trainingScore, text: ""});

      // Einzelscores & Kommentare
      let commentMap = new Map();
      if (data.einzelscore_kommentare && Array.isArray(data.einzelscore_kommentare)) {
          data.einzelscore_kommentare.forEach(k => {
              if (k.metrik) commentMap.set(k.metrik.toLowerCase().trim(), k.text_info || k.text || "");
          });
      }

      fitnessScores.forEach(s => {
  const key = String(s.metrik).toLowerCase().trim();
  const kiText = commentMap.has(key) ? commentMap.get(key) : "";

  const rawNum = extractNumericValue_(s.raw_wert);
  const scoreNum = Number.isFinite(s.num_score) ? s.num_score : "";

  updates.push({
    metrik: s.metrik,
    cat: "Einzelscore",
    ampel: s.ampel,
    raw: Number.isFinite(rawNum) ? rawNum : "",
    score: scoreNum,
    text: kiText || ""
  });
});


      // Ernährung
      let ernaehrungText = (data.bewertung_ernaehrung_7d && data.bewertung_ernaehrung_7d.text) ? data.bewertung_ernaehrung_7d.text : "";
      updates.push({metrik: "7-Tage-Bilanz", cat: "Ernährung", ampel: nutritionScore.ampel, raw: nutritionScore.raw_wert, score: nutritionScore.num_score, text: ernaehrungText});

      // B) Smart Write
      const range = statusSheet.getDataRange();
      const sheetValues = range.getValues(); 
      const rowMap = new Map();
for (let i = 1; i < sheetValues.length; i++) {
  const rowMetrik = String(sheetValues[i][1]).trim().toLowerCase();
  if (rowMetrik) rowMap.set(rowMetrik, i);
}


      const newRows = [];

      updates.forEach(up => {
          const upKey = String(up.metrik).trim().toLowerCase();
if (rowMap.has(upKey)) {
  const rowIndex = rowMap.get(upKey);

              sheetValues[rowIndex][2] = up.ampel;
              sheetValues[rowIndex][3] = up.raw;
              sheetValues[rowIndex][4] = up.score;
              sheetValues[rowIndex][5] = up.text;
          } else {
              // APPEND
              newRows.push([up.cat, up.metrik, up.ampel, up.raw, up.score, up.text]);
          }
      });

      // C) Rückschreiben
      statusSheet.getRange(1, 1, sheetValues.length, 6).setValues(sheetValues);

      // D) Neue Zeilen anhängen
      if (newRows.length > 0) {
          statusSheet.getRange(sheetValues.length + 1, 1, newRows.length, 6).setValues(newRows);
      }
      
      // --- FORMATIERUNG ---
      const lastRow = statusSheet.getLastRow();
      
      // 1. Alles auf Zahl 0.00 zwingen (Basis-Format)
      statusSheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat("0.00");
      statusSheet.getRange(2, 6, lastRow - 1, 1).setWrap(true);

      // 2. +++ RETTUNGS-LOOP FÜR PROZENTE +++
      // Wir suchen gezielt nach "Balance" oder "%" und reparieren das Format
      const finalData = statusSheet.getDataRange().getValues();
      for (let i = 1; i < finalData.length; i++) {
          const mName = String(finalData[i][1]).toLowerCase();
          // Wenn "balance", "intensiv" oder "%" im Namen vorkommt...
          if (mName.includes("balance") || mName.includes("intensiv") || mName.includes("%")) {
              // ...dann setze Format auf "0.0%" (z.B. 10.7%)
              statusSheet.getRange(i + 1, 4).setNumberFormat("0.0%");
          }
      }

      // Breite anpassen
      try { statusSheet.setColumnWidth(6, 400); } catch(e) {}

  } catch (e) { logToSheet('ERROR', `Status-Update Fehler: ${e.message}`); }


  // -------------------------------------------------------
  // BLATT 3: AI_FUTURE_STATUS (App Sync)
  // -------------------------------------------------------
  try {
     const fSheet = ss.getSheetByName('AI_FUTURE_STATUS');
     if (fSheet) {
         const jsonString = JSON.stringify(data);
         fSheet.getRange(1, 1).setValue(jsonString);
         fSheet.getRange(1, 2).setValue(new Date());
     }
  } catch(e) { logToSheet('WARN', `Future-Status Fehler: ${e.message}`); }
  
  try { logToSheet('INFO', `Score: ${gesamtScoreNum} (${gesamtScoreAmpel})`); } catch(e) {}
}


/**
 * V202-HYBRID: Exportiert Daten für Looker Studio.
 * KOMBINATION: 
 * 1. Liest KEI, ACWR & TE Balance direkt aus KK_TIMELINE (Sheet-Formeln).
 * 2. Berechnet weiterhin Scores (Recovery/Training) für die Historie via Script.
 */
function exportLookerChartsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timelineSheet = ss.getSheetByName('KK_TIMELINE'); 
  const baselineSheet = ss.getSheetByName('phys_baseline'); 

  if (!timelineSheet || !baselineSheet) return;

  const timelineDisplayValues = timelineSheet.getDataRange().getDisplayValues();
  const timelineRawValues = timelineSheet.getDataRange().getValues();
  const timelineHeaders = timelineDisplayValues[0];
  
  // Header normalisieren für sichere Suche
  const headersNorm = timelineHeaders.map(h => h.toString().toLowerCase().trim());

  // Baseline laden
  const baselineValues = baselineSheet.getDataRange().getValues();
  const baselineHeaders = baselineValues[0];
  let baselineData = {};
  for(let i=0; i<baselineHeaders.length; i++) {
      baselineData[baselineHeaders[i]] = baselineValues[1][i];
  }

  const indices = {
    date: headersNorm.indexOf('date'),
    is_today: headersNorm.indexOf('is_today'),
    sleep_hours: headersNorm.indexOf('sleep_hours'),
    load_fb_day: headersNorm.indexOf('load_fb_day'),
    readiness: headersNorm.indexOf('garmin_training_readiness'),
    ctl: headersNorm.indexOf('fbctl_obs'), 
    
    // Forecast & Plan
    ess_day: headersNorm.indexOf('coache_ess_day'),
    atl_forecast: headersNorm.indexOf('coache_atl_forecast'),
    ctl_forecast: headersNorm.indexOf('coache_ctl_forecast'),
    
    // >>> QUELLEN AUS TABELLE (DEINE ANFORDERUNG) <<<
    acwr_forecast: headersNorm.indexOf('coache_acwr_forecast'), // Spalte ACWR Forecast
    smart_gains: headersNorm.indexOf('coache_smart_gains'),   // Spalte KEI
    te_balance: headersNorm.indexOf('te_balance_trend'),      // Spalte TE Balance Trend
    
    // Metriken
    monotony7: headersNorm.indexOf('monotony7'),
    strain7: headersNorm.indexOf('strain7'),
    
    // TE Rohdaten (für Score-Berechnung)
    aerobic_te: headersNorm.indexOf('aerobic_te'), 
    anaerobic_te: headersNorm.indexOf('anaerobic_te')
  };

  let heuteRowIndex = -1; 
  for (let i = 1; i < timelineDisplayValues.length; i++) {
    if (parseFloat(String(timelineDisplayValues[i][indices.is_today]).replace(',', '.')) == 1) {
      heuteRowIndex = i; break;
    }
  }
  if (heuteRowIndex === -1) return;

  let historyData = [];
  let forecastData = [];

  // Helper zum sicheren Lesen von Zahlen
  const getVal = (row, idx) => {
    if (idx === -1) return 0;
    let v = row[idx];
    if (typeof v === 'string') v = v.replace(',', '.');
    return parseFloat(v) || 0;
  };

  const parseDateFromDisplayString = (dateStr) => {
    try {
        if (dateStr instanceof Date) return dateStr;
        if (String(dateStr).includes('GMT') || String(dateStr).includes(':')) return new Date(dateStr);
        return new Date(dateStr); 
    } catch(e) { return new Date(); }
  };

  const GESTERN_INDEX = heuteRowIndex;
  const CHART_DAYS = 90; 
  const REAL_START_INDEX = Math.max(1, GESTERN_INDEX - (CHART_DAYS - 1));
  const CALC_START_INDEX = Math.max(1, REAL_START_INDEX - 7); 
  
  let recBuffer = [];
  let trainBuffer = [];
  
  // Speicher für Übergang
  let lastRecovery = 0;
  let lastTraining = 0;

  // --- 1. HISTORY LOOP ---
  for (let i = CALC_START_INDEX; i <= GESTERN_INDEX; i++) { 
    const rowDisplay = timelineDisplayValues[i];
    const rowRaw = timelineRawValues[i];

    // Helper Objekt für Fitness-Calc bauen (damit calculateFitnessMetrics funktioniert)
    let heuteData_i = {};
    timelineHeaders.forEach((h, idx) => { heuteData_i[h] = rowRaw[idx]; });

    // Wir brauchen monotony_varianz eigentlich nicht mehr für TE Balance, 
    // aber calculateFitnessMetrics erwartet das Objekt. Wir setzen 0 oder lesen es.
    let varianzDummy = 0;

    const datenPaket_i = { heute: heuteData_i, baseline: baselineData, monotony_varianz: varianzDummy };
    const fitnessScores_i = calculateFitnessMetrics(datenPaket_i);
    const subScores_i = calculateSubScores(fitnessScores_i);

    const recVal = subScores_i.recoveryScore;
    const trainVal = subScores_i.trainingScore;

    recBuffer.push(recVal);
    trainBuffer.push(trainVal);
    if (recBuffer.length > 7) recBuffer.shift();
    if (trainBuffer.length > 7) trainBuffer.shift();

    if (i >= REAL_START_INDEX) {
        const avgRec = recBuffer.reduce((a,b) => a+b, 0) / recBuffer.length;
        const avgTrain = trainBuffer.reduce((a,b) => a+b, 0) / trainBuffer.length;

        // CTL & Trend (Für History Tabelle behalten wir die Berechnung bei, oder lesen Trend wenn vorhanden)
        // Hier: Klassisch berechnet für Spalte M
        const currentCTL = getVal(rowRaw, indices.ctl);
        let pastCTL = 0;
        if (i - 7 >= 1) {
             pastCTL = getVal(timelineRawValues[i-7], indices.ctl) || currentCTL;
        }
        const ctlTrend = currentCTL - pastCTL;

        // >>> QUELLEN-WECHSEL: WIR LESEN JETZT DIREKT AUS DER TABELLE <<<
        // TE Balance (Mal 100 für Prozent)
        const teBalanceVal = getVal(rowRaw, indices.te_balance) * 100;
        
        // KEI
        const smartGainsVal = getVal(rowRaw, indices.smart_gains);

        // Speicher für Forecast-Übergang aktualisieren
        lastRecovery = recVal;
        lastTraining = trainVal;

        historyData.push([
          parseDateFromDisplayString(rowDisplay[indices.date]),
          rowRaw[indices.sleep_hours],
          rowRaw[indices.load_fb_day],
          rowRaw[indices.readiness],
          rowRaw[indices.monotony7],
          rowRaw[indices.strain7],
          teBalanceVal,  // <--- GELESEN AUS SPALTE
          recVal,
          trainVal,
          avgRec,
          avgTrain,
          currentCTL,
          ctlTrend,
          smartGainsVal  // <--- GELESEN AUS SPALTE
        ]);
    }
  }

  // --- 2. FORECAST LOOP ---
  for (let i = heuteRowIndex; i < timelineDisplayValues.length; i++) { 
    const rowRaw = timelineRawValues[i];
    const dateStr = timelineDisplayValues[i][indices.date];
    if (!dateStr || dateStr === "") continue;

    // >>> QUELLEN-WECHSEL: DIREKT LESEN <<<
    
    // 1. TE Balance (aus Spalte TE_Balance_Trend)
    // Da die Formel im Sheet ist, vertrauen wir ihr.
    const forecastVarianz = getVal(rowRaw, indices.te_balance) * 100;

    // 2. KEI (aus Spalte coachE_Smart_Gains)
    const smartScoreV2 = getVal(rowRaw, indices.smart_gains);

    // 3. ACWR (aus Spalte coachE_ACWR_forecast)
    const acwrVal = getVal(rowRaw, indices.acwr_forecast);

    // 4. Monotonie & Strain (Fallback Schätzung für Charts, falls leer)
    const plannedLoad = getVal(rowRaw, indices.ess_day);
    let monoEst = 1.5; 
    let strainEst = plannedLoad * 7; 

    const realMono = getVal(rowRaw, indices.monotony7);
    const realStrain = getVal(rowRaw, indices.strain7);
    if (realStrain > 0) strainEst = realStrain;
    if (realMono > 0) monoEst = realMono;

    forecastData.push([
      parseDateFromDisplayString(dateStr),
      plannedLoad,
      rowRaw[indices.atl_forecast],
      rowRaw[indices.ctl_forecast],
      acwrVal,             // <--- GELESEN
      monoEst,
      strainEst,
      forecastVarianz,     // <--- GELESEN
      lastRecovery,        // Übergangswert (da Zukunft unbekannt)
      lastTraining,        // Übergangswert
      smartScoreV2         // <--- GELESEN
    ]);
  }

  // --- SCHREIBEN DER DATEN ---
  const historyHeaders = [
      'Datum', 'Schlaf (h)', 'Tatsächlicher Load', 'Training Readiness', 
      'Monotony7', 'Strain7', 'TE Balance (% Intensiv)', 
      'Recovery Score', 'Training Score', 'Recovery Score (Ø7d)', 'Training Score (Ø7d)',
      'CTL (Fitness)', 'CTL Trend (7d)',
      'KEI Score' 
  ];
  const forecastHeaders = ['Datum', 'Geplanter Load (ESS)', 'ATL Prognose', 'CTL Prognose', 'ACWR Prognose', 'Monotony7 Prognose', 'Strain7 Prognose', 'TE Balance (% Intensiv)', 'Recovery Score', 'Training Score', 'KEI Forecast'];

  let historySheet = ss.getSheetByName('AI_DATA_HISTORY');
  if (!historySheet) historySheet = ss.insertSheet('AI_DATA_HISTORY');
  historySheet.clear();
  historySheet.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]).setFontWeight('bold');
  if (historyData.length > 0) {
    historySheet.getRange(2, 1, historyData.length, historyHeaders.length).setValues(historyData);
    historySheet.getRange(2, 1, historyData.length, 1).setNumberFormat('dd.MM.yyyy');
    historySheet.getRange(2, 10, historyData.length, 4).setNumberFormat('0.0'); 
  }

  let forecastSheet = ss.getSheetByName('AI_DATA_FORECAST');
  if (!forecastSheet) forecastSheet = ss.insertSheet('AI_DATA_FORECAST');
  forecastSheet.clear();
  forecastSheet.getRange(1, 1, 1, forecastHeaders.length).setValues([forecastHeaders]).setFontWeight('bold');
  if (forecastData.length > 0) {
    forecastSheet.getRange(2, 1, forecastData.length, forecastHeaders.length).setValues(forecastData);
    forecastSheet.getRange(2, 1, forecastData.length, 1).setNumberFormat('dd.MM.yyyy');
    forecastSheet.getRange(2, 8, forecastData.length, 1).setNumberFormat('0.0');
  }

  if(typeof logToSheet === 'function') logToSheet('INFO', `[Chart Export V202] Daten erfolgreich synchronisiert (Quellen: Sheet-Formeln).`);
}


/**
 * V149-DATA-FIX: Fügt 'Sport_x' zur History-CSV hinzu.
 * FIX V150: Liest TE_Balance_Trend direkt aus Timeline und schreibt in D11.
 */
function getSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (typeof TIMELINE_SHEET_NAME !== 'undefined') ? TIMELINE_SHEET_NAME : 'KK_TIMELINE';
  const timelineSheet = ss.getSheetByName(sheetName);
  const baselineSheet = ss.getSheetByName(BASELINE_SHEET_NAME); 

  if (!timelineSheet || !baselineSheet) {
    const errorMsg = `Blatt fehlt (${sheetName} oder ${BASELINE_SHEET_NAME}).`;
    logToSheet('ERROR', errorMsg);
    throw new Error(errorMsg);
  }

  const timelineValues = timelineSheet.getDataRange().getValues(); 
  const baselineValues = baselineSheet.getDataRange().getValues();

  // ----------------------------------------------------
// BASELINE -> baselineData (Key/Value aus 2 Spalten)
// Erwartung: Spalte A = Key, Spalte B = Value
// ----------------------------------------------------
const baselineData = {};
for (let r = 0; r < baselineValues.length; r++) {
  const key = String(baselineValues[r][0] ?? "").trim();
  if (!key) continue;

  // Optional: Headerzeile überspringen (falls A1 z.B. "Key" ist)
  if (r === 0 && key.toLowerCase() === "key") continue;

  baselineData[key] = baselineValues[r][1];
}


  const timelineHeadersRaw = timelineValues[0].map(h => h.toString().trim());
const timelineHeaders = timelineHeadersRaw.map(h => h.toLowerCase());

// WICHTIG: timelineValues ist bereits getValues() -> kein zweites getValues()!
const timelineRawValues = timelineValues;

// ----------------------------------------------------
// HEUTE-ROWINDEX FINDEN (is_today bevorzugt, sonst Datum)
// ----------------------------------------------------
const tz = Session.getScriptTimeZone();
const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

// 1) bevorzugt: is_today Spalte
const isTodayIdx = timelineHeaders.indexOf('is_today');
let heuteRowIndex = -1;

if (isTodayIdx !== -1) {
  for (let r = 1; r < timelineRawValues.length; r++) {
    const v = timelineRawValues[r][isTodayIdx];
    const n = parseGermanFloat(v); // Komma/Punkt robust
    if (n === 1) { heuteRowIndex = r; break; }
  }
}

// 2) Fallback: date Spalte = heute
if (heuteRowIndex === -1) {
  const dateIdx = timelineHeaders.indexOf('date');
  if (dateIdx === -1) {
    throw new Error("KK_TIMELINE: Weder 'is_today' noch 'date' Spalte vorhanden -> kann HEUTE nicht finden.");
  }

  for (let r = 1; r < timelineRawValues.length; r++) {
    const d = timelineRawValues[r][dateIdx];
    const dObj = (d instanceof Date) ? d : new Date(d);
    if (dObj instanceof Date && !isNaN(dObj.getTime())) {
      const key = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
      if (key === todayKey) { heuteRowIndex = r; break; }
    }
  }
}

if (heuteRowIndex === -1) {
  throw new Error(`KK_TIMELINE: HEUTE-Zeile nicht gefunden (todayKey=${todayKey}).`);
}


// Heute-Zeile als Objekt: lowercase + RAW-Header zusammenführen
const heuteDataLower = arrayToObject(timelineHeaders, timelineRawValues[heuteRowIndex]);
const heuteDataRaw   = arrayToObject(timelineHeadersRaw, timelineRawValues[heuteRowIndex]);
const heuteData = Object.assign({}, heuteDataLower, heuteDataRaw);

  
  // --- KONFIGURATION ---
  const KI_HISTORY_TAGE = 14;
  const TE_CALC_DAYS = 28; 
  
  // --- 1. ERNÄHRUNG (7 Tage) ---
  let totalDeficit = 0;
  let deficitDays = 0;
  const deficitIndex = timelineHeaders.indexOf('deficit');
  const histDays = (typeof VERGANGENHEIT_TAGE !== 'undefined') ? VERGANGENHEIT_TAGE : 7;
  const nutritionStartIndex = Math.max(1, heuteRowIndex - histDays);
  
  for (let i = nutritionStartIndex; i < heuteRowIndex; i++) {
    const row = timelineValues[i];
    if (row[deficitIndex] !== "") {
      const deficitVal = parseGermanFloat(row[deficitIndex]);
      if (!isNaN(deficitVal)) {
        totalDeficit += deficitVal;
        deficitDays++;
      }
    }
  }
  const avg_7d_deficit = (deficitDays > 0) ? totalDeficit / deficitDays : 0;

  // --- 2. KI-HISTORIE TEXT (MIT SPORT!) ---
  const histIndices = {
    date: timelineHeaders.indexOf('date'),
    load_fb_day: timelineHeaders.indexOf('load_fb_day'),
    Sport_x: timelineHeaders.indexOf('Sport_x'), 
    Aerobic_TE: timelineHeaders.indexOf('Aerobic_TE'), 
    Anaerobic_TE: timelineHeaders.indexOf('Anaerobic_TE'), 
    coachE_ATL_forecast: timelineHeaders.indexOf('coachE_ATL_forecast'),
    coachE_CTL_forecast: timelineHeaders.indexOf('coachE_CTL_forecast'),
    coachE_ACWR_forecast: timelineHeaders.indexOf('coachE_ACWR_forecast'),
    Monotony7: timelineHeaders.indexOf('Monotony7'),
    Strain7: timelineHeaders.indexOf('Strain7')
  };
  
  const kiHeaders = ['date', 'Sport', 'Actual_Load', 'Aerobic_TE', 'Anaerobic_TE', 'Planned_ATL', 'Planned_CTL', 'Planned_ACWR', 'Monotony7', 'Strain7'];
  let historyDataCSV = [kiHeaders.join(',')];

  const kiStartIndex = Math.max(1, heuteRowIndex - KI_HISTORY_TAGE);
  for (let i = kiStartIndex; i < heuteRowIndex; i++) {
    const row = timelineValues[i];
    const rowRaw = timelineRawValues[i]; 
     
    const aerobicTE = parseGermanFloat(rowRaw[histIndices.Aerobic_TE]) || 0;
    const anaerobicTE = parseGermanFloat(rowRaw[histIndices.Anaerobic_TE]) || 0;
    const sportVal = (histIndices.Sport_x !== -1) ? row[histIndices.Sport_x] : "N/A"; 

    const kiRow = [
      `"${row[histIndices.date]}"`,
      `"${sportVal}"`, 
      `"${row[histIndices.load_fb_day]}"`, 
      `"${aerobicTE}"`, 
      `"${anaerobicTE}"`, 
      `"${row[histIndices.coachE_ATL_forecast]}"`,
      `"${row[histIndices.coachE_CTL_forecast]}"`,
      `"${row[histIndices.coachE_ACWR_forecast]}"`,
      `"${row[histIndices.Monotony7]}"`,
      `"${row[histIndices.Strain7]}"`
    ];
    historyDataCSV.push(kiRow.join(','));
  }
  const historyCSVString = historyDataCSV.join('\n');

  // --- 3. TE BALANCE (MODIFIZIERT: Lese aus Sheet + Schreibe D11) ---
  let monotonyVarianz = 0.0;
  try {
    // Spaltenindex suchen
    const teTrendIndex = timelineHeaders.indexOf('te_balance_trend');
    
    if (teTrendIndex !== -1) {
      // Wert aus der heutigen Zeile lesen
      monotonyVarianz = parseGermanFloat(timelineRawValues[heuteRowIndex][teTrendIndex]);
      
      // SCHREIBEN: Direkt in AI_REPORT_STATUS Zelle D11
      const outSheetName = (typeof OUTPUT_STATUS_SHEET !== 'undefined') ? OUTPUT_STATUS_SHEET : 'AI_REPORT_STATUS';
      const statusSheet = ss.getSheetByName(outSheetName);
      if (statusSheet) {
        statusSheet.getRange("D11").setValue(monotonyVarianz);
      }
      
      logToSheet('DEBUG', `[getSheetData] TE Balance (${monotonyVarianz}) aus Timeline gelesen & nach D11 geschrieben.`);
    } else {
      logToSheet('WARN', '[getSheetData] Spalte TE_Balance_Trend nicht in Timeline gefunden.');
    }
  } catch (e) {
    logToSheet('ERROR', `TE Balance Fehler: ${e.message}`);
  }

  // --- 4. ZUKUNFT ---
  let maxAcwr = -999, maxAcwrDate = "N/A", minAcwr = 999, minAcwrDate = "N/A";
  let zukunftsPlan = [];

  const acwrIndex = timelineHeaders.indexOf('coachE_ACWR_forecast');
  const atlIndex = timelineHeaders.indexOf('coachE_ATL_forecast');
  const ctlIndex = timelineHeaders.indexOf('coachE_CTL_forecast');
  const loadIstIndex = timelineHeaders.indexOf('load_fb_day');
  const strainIndex = timelineHeaders.indexOf('Strain7');
  const teAeIndex = timelineHeaders.indexOf('Target_Aerobic_TE');      
  const teAnIndex = timelineHeaders.indexOf('Target_Anaerobic_TE');
  const dateIndex = timelineHeaders.indexOf('date');
  const dayIndex = timelineHeaders.indexOf('date.2');
  const essIndex = timelineHeaders.indexOf('coachE_ESS_day');
  const zoneIndex = timelineHeaders.indexOf('Zone');
  const sportIndex = timelineHeaders.indexOf('Sport_x');
  const monotonyIndex = timelineHeaders.indexOf('Monotony7');
  const fixIndex = timelineHeaders.indexOf('fix');
  const phaseIndex = timelineHeaders.indexOf('Week_Phase');
  
  const futureDays = (typeof ZUKUNFT_TAGE !== 'undefined') ? ZUKUNFT_TAGE : 14;

  for (let i = heuteRowIndex + 1; i < timelineRawValues.length && i <= heuteRowIndex + futureDays; i++) {
      const row = timelineRawValues[i];
      const rowDisplay = timelineValues[i];

      const dateVal = row[dateIndex];
      if (!(dateVal instanceof Date) || isNaN(dateVal.getTime())) continue;

      const acwrVal = parseGermanFloat(row[acwrIndex]);
      
      zukunftsPlan.push({
        datum: Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        tag: rowDisplay[dayIndex],
        original_load_ess: parseGermanFloat(row[essIndex]),
        original_act_load: (loadIstIndex !== -1) ? parseGermanFloat(row[loadIstIndex]) : 0, 
        original_zone: (zoneIndex !== -1) ? rowDisplay[zoneIndex] : "N/A",
        original_sport: (sportIndex !== -1) ? rowDisplay[sportIndex] : "N/A",
        original_acwr: acwrVal,
        original_monotony: (monotonyIndex !== -1) ? rowDisplay[monotonyIndex] : "N/A",
        original_fix_marker: (fixIndex !== -1) ? rowDisplay[fixIndex] : "",
        original_atl: (atlIndex !== -1) ? parseGermanFloat(row[atlIndex]) : 0,
        original_ctl: (ctlIndex !== -1) ? parseGermanFloat(row[ctlIndex]) : 0,
        original_strain: (strainIndex !== -1) ? rowDisplay[strainIndex] : "N/A",
        original_te_ae: (teAeIndex !== -1) ? parseGermanFloat(row[teAeIndex]) : 0, 
        original_te_an: (teAnIndex !== -1) ? parseGermanFloat(row[teAnIndex]) : 0,
        week_phase: (phaseIndex !== -1) ? String(rowDisplay[phaseIndex]).trim().toUpperCase() : "A"
      });

      if (acwrVal > maxAcwr) { maxAcwr = acwrVal; maxAcwrDate = dateVal; }
      if (acwrVal < minAcwr) { minAcwr = acwrVal; minAcwrDate = dateVal; }
  }

  logToSheet('INFO', `Daten für heute (${timelineValues[heuteRowIndex][dateIndex]}) geladen. TE-Balance über ${TE_CALC_DAYS} Tage.`);

  // --- RTP SCAN (V129.1) ---
  let rtpStatus = "HEALTHY"; 
  const sportIdx = timelineHeaders.indexOf('Sport_x');
  
  if (sportIdx !== -1 && heuteRowIndex >= 1) {
      for (let i = 1; i <= 10; i++) {
          let checkIndex = heuteRowIndex - i;
          if (checkIndex < 1) break; 

          let sportLog = String(timelineRawValues[checkIndex][sportIdx]).toLowerCase();
          
          if (sportLog.includes("krank") || sportLog.includes("sick") || sportLog.includes("infekt")) {
              if (i < 10) {
                  rtpStatus = `RTP_PHASE_${i}`;
                  logToSheet('INFO', `[RTP-Scan] Infekt vor ${i} Tagen gefunden. Status: ${rtpStatus}`);
              } else {
                  rtpStatus = "HEALTHY";
                  logToSheet('INFO', `[RTP-Scan] Infekt vor 10 Tagen gefunden. Protokoll beendet: HEALTHY`);
              }
              break; 
          }
      }
  }
  logToSheet('INFO', `[RTP Check] Finaler Status: ${rtpStatus}`);

  return { 
    heute: heuteData,
    baseline: baselineData,
    ernaehrung: { avg_7d_deficit: avg_7d_deficit },
    history: historyCSVString, 
    monotony_varianz: monotonyVarianz, // Gibt jetzt den gelesenen Wert zurück
    rtp_status: rtpStatus,
    zukunft: {
      plan: zukunftsPlan,
      maxAcwr: maxAcwr, maxAcwrDate: maxAcwrDate,
      minAcwr: minAcwr, minAcwrDate: minAcwrDate
    }
  };
}

/**
 * (V7-Logik): Wandelt ein Header-Array und ein Daten-Array in ein Objekt um.
 */
function arrayToObject(headers, data) {
  const obj = {};
  headers.forEach((header, index) => {
    let value = data[index];

    // Robust numeric coercion for strings that *look* like numbers.
    // Avoid touching arbitrary text values.
    if (typeof value === 'string') {
      const s = value.trim();
      const looksNumeric = /\d/.test(s) && /^[\s\d.,%+\-]+$/.test(s);
      if (looksNumeric) {
        const n = extractNumericValue_(s);
        if (Number.isFinite(n)) value = n;
      }
    }

    obj[header] = value;
  });
  return obj;
}

/**
 * (V7-Logik): Hilfsfunktion zum Parsen von Zahlen.
 *
 * Note: Prefer `parseGermanFloat_` / `extractNumericValue_` for robust handling
 * of German decimal comma + thousands separators.
 */
function parseGermanFloat(value) {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const n = parseGermanFloat_(value);
    return Number.isFinite(n) ? n : value;
  }
  return value;
}

function parseGermanFloat_(value) {
  if (value === null || value === undefined) return NaN;
  const s = String(value).trim();
  if (!s) return NaN;

  // Deutsch -> Zahl (1,13 -> 1.13). Tausenderpunkte entfernen.
  const cleaned = s
    .replace(/\s/g, '')
    .replace('%', '')
    .replace(/\.(?=\d{3}(\D|$))/g, '')  // nur Tausenderpunkte
    .replace(',', '.');

  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : NaN;
}

function extractNumericValue_(raw) {
  if (raw === null || raw === undefined) return NaN;
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : NaN;

  const s = String(raw).trim();
  if (!s) return NaN;

  // Prozent-Strings korrekt: "9,3%" -> 0.093
  if (s.includes("%")) {
    const n = parseGermanFloat_(s);
    return Number.isFinite(n) ? (n / 100) : NaN;
  }

  // Erste Zahl aus "45 bpm", "6,9h", "1,13" extrahieren
  const m = s.match(/-?\d[\d\s\.,]*/);
  if (!m) return NaN;

  return parseGermanFloat_(m[0]);
}




/**
 * (V7-Logik): Robuste Funktion, um einen KI-Textwert in eine Zahl zu parsen.
 */
function cleanNumberFromKI(text) {
  if (typeof text === 'number') return text;
  if (typeof text !== 'string') return text;

  // Try to extract a numeric token from the text and parse it robustly.
  // Supports German formats like "1,23", "1.689", "9,3%", "52 bpm".
  const m = String(text).match(/-?\d[\d\s\.,%]*/);
  if (!m) return text;

  const token = m[0];
  const n = extractNumericValue_(token);
  return Number.isFinite(n) ? n : text;
}

/**
 * V6.4.1-SECURE: Robuste Verbindung zu Gemini (POST Request).
 * Holt den API-Key ausschließlich aus den Script Properties.
 */
function callGeminiAPI(promptText) {
  // KI-frei (2026-10-26): Gemini-Aufrufe entfernt.
  throw new Error('callGeminiAPI: KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.');
}

function clearMemory() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_TELEGRAM_DATA');
  console.log("Speicher geleert, Commander!");
}

function doGet(e) {
  e = e || { parameter: {} };
  const p = e.parameter || {};

  // Briefings lokal (Cockpit + Ollama)
  if (p.mode === 'briefingPrompt') return ckDash_briefingPrompt_(p);
  if (p.mode === 'briefingData') return ckDash_briefingData_(p);
  if (p.mode === 'telegramJobs') return ckTg_jobs_();
  if (p.mode === 'briefingSources') return ckDash_briefingSources_();
  if (p.mode === 'activityReviews') return ckDash_activityReviews_(p);
  if (p.mode === 'supervisorJob') return ckSup_getJob_();
  if (p.mode === 'diag') return ckDash_diag_();
  // KI-Extras (Treffsicherheit, Selbsttest, Plausibilität, Ernährung, Alltag-Check)
  if (p.mode === 'extras') return ckX_get_(p);
  if (p.mode === 'calendarDays') return ckX_calendar_(p);
  if (p.mode === 'physBaseline') return ckX_phys_();
  if (p.mode === 'loadFactors') return ckX_loadFactors_();



  // ============================================================
  // Coach Kira Dashboard API – minimalinvasive Zusatzrouten
  // ============================================================
  if (p.mode === 'status') {
    return ckDash_status_();
  }

if (p.mode === 'consolidatedAnalysis') {
  return ckDash_consolidatedAnalysis_();
}

if (p.action === 'runConsolidatedAnalysis') {
  return ckDash_runConsolidatedAnalysis_(p);
}

  if (p.mode === 'planSimulationJson') {
  return ckDash_planSimulationJson_();
}

if (p.mode === 'planBriefing') {
  return ckDash_planBriefing_();
}

if (p.action === 'generatePlanBriefing') {
  return ckDash_generatePlanBriefing_(p);
}

if (p.mode === 'statusAiAnalysis') {
  return ckDash_statusAiAnalysis_();
}

if (p.action === 'saveSimulatedPlan') {
  return ckDash_saveSimulatedPlan_(p);
}

if (p.action === 'saveWellbeing') {
  return ckDash_saveWellbeing_(p);
}

if (p.action === 'submitCockpitEntry') {
  return ckDash_submitCockpitEntry_(p);
}


if (p.action === 'runStatusAiAnalysis') {
  return ckDash_runStatusAiAnalysis_(p);
}

  if (p.mode === 'planJson') {
    return ckDash_planJson_();
  }

  if (p.mode === 'planJsonV2') {
  return ckDash_planJsonV2_();
}

  if (p.mode === 'chartData') {
    return ckDash_chartData_(p);
  }

  if (p.mode === 'runStatus') {
    return ckDash_runStatus_(p);
  }

  if (p.mode === 'logs') {
    return ckDash_logs_(p);
  }

  if (p.action === 'runSupervisor') {
    return ckDash_queueAction_(p, 'runKiraGeminiSupervisor');
  }

  if (p.action === 'generateFuturePlan') {
    return ckDash_queueAction_(p, 'generateOptimized14DayPlan');
  }

  const cache = CacheService.getUserCache();
  const lock  = LockService.getScriptLock();

  // --- WebApp Guardrails ---
  // Protect JSON endpoints (mode=json, mode=timeline, format=json) with a shared key.
  // ScriptProperties:
  // - WEBAPP_API_KEY (required)
  function requireWebappKey_(purpose) {
    const props = PropertiesService.getScriptProperties();
    const expected = String(props.getProperty('WEBAPP_API_KEY') || '').trim();
    const got = String((e && e.parameter && e.parameter.key) || '').trim();

    if (!expected) {
      // Fail closed: better explicit error than public data leak.
      return ContentService
        .createTextOutput(JSON.stringify({ error: true, message: 'WEBAPP_API_KEY fehlt in ScriptProperties' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (got !== expected) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: true, statusCode: 401, message: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return null;
  }

    // --- BUILD-ID (Debug) ---
  // Aufruf: .../exec?mode=build
  if (e && e.parameter && e.parameter.mode === 'build') {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, build: KK_BUILD_ID }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // --- NEU: SCRIPTABLE PLAN WIDGET API (7-Tage Vorschau) ---
  // Aufruf: .../exec?format=json&key=...
  if (e && e.parameter && e.parameter.format === 'json') {
    const auth = requireWebappKey_('format=json');
    if (auth) return auth;
    return getPlanForWidget();
  }

  // --- 1. DASHBOARD API-MODUS (Optimiert für Scriptable & Co.) ---
  // Aufruf: .../exec?mode=json&key=...
  if (e && e.parameter && e.parameter.mode === 'json') {

    const auth = requireWebappKey_('mode=json');
    if (auth) return auth;

    // Zuerst im Cache nachsehen
    const cachedData = cache.get("WEBAPP_FULL_PAYLOAD");
    if (cachedData) {
      console.log("Serviere JSON aus dem Cache 🚀");
      return ContentService
        .createTextOutput(cachedData)
        .setMimeType(ContentService.MimeType.JSON);
    }

    try {
      lock.waitLock(8000); // 8 Sekunden warten, falls gerade gerechnet wird
      const jsonData = getDashboardDataAsStringV76(); // Deine unveränderte Funktion

      // Ergebnis für 5 Minuten merken
      cache.put("WEBAPP_FULL_PAYLOAD", jsonData, 300);

      lock.releaseLock();
      return ContentService
        .createTextOutput(jsonData)
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      if (lock.hasLock()) lock.releaseLock();
      return ContentService
        .createTextOutput(JSON.stringify({ error: true, message: "Timeout: Tabelle blockiert." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

// --- 1b. TIMELINE API-MODUS (Raw KK_TIMELINE als JSON) ---
// Aufruf:
//   .../exec?mode=timeline
//   .../exec?mode=timeline&days=14
//   .../exec?mode=timeline&days=90
//   .../exec?mode=timeline&days=90&future=14   (Historie + Forecast)
//   .../exec?mode=timeline&days=&future=14     (ALL Historie + Forecast)
if (e && e.parameter && e.parameter.mode === 'timeline') {

  const auth = requireWebappKey_('mode=timeline');
  if (auth) return auth;

  // days: null => "ALL" (keine History-Slice)
  const daysRaw = (e.parameter.days !== undefined && e.parameter.days !== null)
    ? String(e.parameter.days).trim()
    : "";

  const days = (daysRaw !== "")
    ? parseInt(daysRaw, 10)
    : null;

  // future: Default 14, wenn NICHT gesetzt
  const futureRaw = (e.parameter.future !== undefined && e.parameter.future !== null)
    ? String(e.parameter.future).trim()
    : "";

  const futureDays = (futureRaw !== "")
    ? parseInt(futureRaw, 10)
    : 14; // <- wichtig: Default Forecast 14

  const dHist = (Number.isFinite(days) && days > 0) ? days : null;
  const dFut  = (Number.isFinite(futureDays) && futureDays > 0) ? futureDays : 0;

  // Cache-Key muss days + future unterscheiden
  const daysKey   = dHist ? String(dHist) : "ALL";
  const futureKey = dFut ? String(dFut) : "0";
  const cacheKey  = "TIMELINE_PAYLOAD_" + daysKey + "_F" + futureKey;

  const cached = cache.get(cacheKey);
  if (cached) {
    console.log("Serviere TIMELINE JSON aus dem Cache 🚀 " + cacheKey);
    return ContentService
      .createTextOutput(cached)
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    lock.waitLock(8000);

    // ✅ Muss existieren:
    // function getTimelinePayload(days, futureDays) { ... }
    const payloadObj = getTimelinePayload(dHist, dFut);
    const json = JSON.stringify(payloadObj);

    cache.put(cacheKey, json, 300);

    lock.releaseLock();
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    if (lock.hasLock()) lock.releaseLock();
    return ContentService
      .createTextOutput(JSON.stringify({ error: true, message: String((err && err.stack) || (err && err.message) || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

  // --- 2. STANDARD WEBAPP (HTML) ---
  let page = e && e.parameter ? e.parameter.page : null;

  var appUrl = ScriptApp.getService().getUrl();
  // Dein Fallback-Link bleibt drin
  if (!appUrl) appUrl = "https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec";

  var template;

  // Die Weiche für deine Apps
  if (page === 'chat') {
    template = HtmlService.createTemplateFromFile('ChatApp');
  }
  else if (page === 'plan') {
    template = HtmlService.createTemplateFromFile('PlanApp');
  }
  // --- NEU: Garmin/Firstbeat PlanApp (parallel zur alten PlanApp) ---
  // Aufruf: .../exec?page=planfb
  else if (page === 'planfb') {
    template = HtmlService.createTemplateFromFile('PlanAppFB');
  }
  // --- NEU: Charts Dashboard ---
  // Aufruf: .../exec?page=charts
  else if (page === 'charts') {
    template = HtmlService.createTemplateFromFile('charts');
    template._debug_loaded = 'charts';
  }
  // --- NEU: Charts V2 Dashboard ---
  // Aufruf: .../exec?page=charts_v2
  else if (page === 'charts_v2') {
    template = HtmlService.createTemplateFromFile('charts_V2');
    template._debug_loaded = 'charts_V2';
  }
  else if (page === 'calc') {
    template = HtmlService.createTemplateFromFile('CalculatorApp');
  }
  else if (page === 'lab') {
    template = HtmlService.createTemplateFromFile('LoadLab');
  }
  else if (page === 'log') {
    template = HtmlService.createTemplateFromFile('Tactical_Log');
  }
  else if (page === 'rtp') {
    template = HtmlService.createTemplateFromFile('RTP_Smoothstep_Simulator');
  }
  else if (page === 'prime') {
    template = HtmlService.createTemplateFromFile('PrimeRangeFinder');
  }
  else {
    template = HtmlService.createTemplateFromFile('WebApp_V2'); // Hauptseite
  }

  template.pubUrl = appUrl;

template._debug_build = KK_BUILD_ID;

let out;
try {
  out = template.evaluate();
} catch (err) {
  // Wichtig: Damit siehst du die echte Ursache (meist mit Hinweis auf charts.html / Zeilennummer)
  return ContentService
    .createTextOutput("TEMPLATE EVAL ERROR:\n" + (err && (err.stack || err.message) || String(err)))
    .setMimeType(ContentService.MimeType.TEXT);
}

// bewusst ASCII-only:
out.setTitle('Coach Kira');
out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
out.addMetaTag('viewport', 'width=device-width, initial-scale=1');

return out;
}



// --- Telegram config (Secrets live in ScriptProperties; never hardcode) ---
// Required ScriptProperties:
// - TELEGRAM_BOT_TOKEN
// - TELEGRAM_CHAT_ID
function getTelegramConfig_() {
  const props = PropertiesService.getScriptProperties();
  const token = String(props.getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  const chatId = String(props.getProperty('TELEGRAM_CHAT_ID') || '').trim();
  return { token, chatId };
}

function doPost(e) {
  // Coach-Kira-Chat aus dem Dashboard (vor dem Telegram-Webhook)
  const ckBody = (e && e.postData && e.postData.contents) || '';
  if ((e && e.parameter && e.parameter.mode === 'saveBriefing') || /"kind"\s*:\s*"saveBriefing"/.test(ckBody)) {
    return ckDash_saveBriefing_(e);
  }
  if ((e && e.parameter && e.parameter.mode === 'saveActivityReview') || /"kind"\s*:\s*"saveActivityReview"/.test(ckBody)) {
    return ckDash_saveActivityReview_(e);
  }
  if ((e && e.parameter && e.parameter.mode === 'saveSupervisor') || /"kind"\s*:\s*"saveSupervisor"/.test(ckBody)) {
    return ckSup_save_(e);
  }
  if ((e && e.parameter && e.parameter.mode === 'saveExtra') || /"kind"\s*:\s*"saveExtra"/.test(ckBody)) {
    return ckX_save_(e);
  }
  if (/"kind"\s*:\s*"saveDayMetrics"/.test(ckBody)) {
    return ckGA_saveDays_(e);
  }

  if (/"kind"\s*:\s*"telegramAnswer"/.test(ckBody)) {
    return ckTg_answer_(e);
  }
  if ((e && e.parameter && e.parameter.mode === 'chatLog') || /"kind"\s*:\s*"chatLog"/.test(ckBody)) {
    return ckDash_chatLog_(e);
  }

  const ckIsChat =
    (e && e.parameter && e.parameter.mode === 'chat') ||
    (/"messages"\s*:/.test(ckBody) && !/"update_id"/.test(ckBody));
  if (ckIsChat) {
    return ckDash_chat_(e);
  }
  try {
    // Optional webhook key guard (only enforced if TELEGRAM_WEBHOOK_KEY is set)

    const props = PropertiesService.getScriptProperties();
    const expectedKey = String(props.getProperty('TELEGRAM_WEBHOOK_KEY') || '').trim();
    if (expectedKey) {
      const gotKey = String((e && e.parameter && e.parameter.tg_key) || '').trim();
      if (gotKey !== expectedKey) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'unauthorized' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    const contents = JSON.parse(e.postData.contents);
    const updateId = contents.update_id;
    const lastId = props.getProperty('LAST_UPDATE_ID');
    
    if (updateId && updateId === lastId) {
      return ContentService.createTextOutput(JSON.stringify({status: "ignored_duplicate"}));
    }
    
    if (contents.message) {
      props.setProperty('LAST_UPDATE_ID', updateId);
      props.setProperty('LAST_TELEGRAM_DATA', JSON.stringify(contents));
      
      // TRIGGER ERSTELLEN (KEIN LÖSCHEN HIER - DAS MACHT DER ARBEITER!)
      ScriptApp.newTrigger('processTelegramBackground')
        .timeBased()
        .after(500) // Schneller starten
        .create();
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "ok"}));
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error"}));
  }
}

function processTelegramBackground() {
  const lock = LockService.getScriptLock();
  try {
    // Warte maximal 10 Sekunden, ob ein anderer Prozess fertig wird
    lock.waitLock(10000); 
    
    const props = PropertiesService.getScriptProperties();
    const dataStr = props.getProperty('LAST_TELEGRAM_DATA');
    
    // Trigger sofort löschen, damit er nicht doppelt feuert
    deleteTriggersForFunction('processTelegramBackground');
    
    if (!dataStr) {
      lock.releaseLock();
      return;
    }

    props.deleteProperty('LAST_TELEGRAM_DATA');
    const contents = JSON.parse(dataStr);
    handleCommand(contents);
    
    lock.releaseLock(); // Tür wieder aufmachen
  } catch (e) {
    console.error("Lock konnte nicht aktiviert werden oder Fehler: " + e);
  }
}

function handleCommand(contents) {
  const chatId = contents.message.chat.id;
  const text = contents.message.text || "";
  const user = contents.message.from.first_name || "Commander";

  // KI-frei (Stufe 3): freie Fragen beantwortet Kira lokal auf dem Mac
  if (text && !text.startsWith('/')) {
    ckTg_enqueue_(chatId, text, user);
    return;
  }

  if (text.startsWith("/status")) {
    sendTelegramMessage(chatId, "Befehl erhalten. System-Scan läuft... 📡");
    
    try {
      const dashboard = getDashboardDataForTelegram();
      if (!dashboard.success) throw new Error(dashboard.error);

      const getEmoji = (ampel) => {
        const a = String(ampel).toUpperCase();
        if (a.includes("GRÜN") || a.includes("GREEN")) return "🟢";
        if (a.includes("GELB") || a.includes("YELLOW")) return "🟡";
        if (a.includes("ROT") || a.includes("RED")) return "🔴";
        if (a.includes("BLAU") || a.includes("BLUE")) return "🔵";
        if (a.includes("LILA") || a.includes("VIOLETT")) return "🟣"; // NEU
        if (a.includes("ORANGE")) return "🟠";                      // NEU
        return "⚪";
      };

      // Metriken finden
      const readiness = dashboard.scores.find(s => s.metrik === "Training Readiness") || { num_score: 0, ampel: "GRAU" };
      const gesamt = dashboard.scores.find(s => s.metrik === "Gesamtscore") || { num_score: 0, ampel: "GRAU" };

      // Markdown-Säuberung für ALLE Texte
      const clean = (t) => String(t).replace(/[_*\[\]()]/g, " ");

      let report = `${getEmoji(readiness.ampel)} *FULL REPORT: ${user.toUpperCase()}*\n` +
                   `------------------------------------\n\n` +
                   `🏆 *HAUPTWERTE*\n` +
                   `• Readiness: *${readiness.num_score}%*\n` +
                   `• Gesamtscore: *${gesamt.num_score}%*\n` +
                   `• Training: *${dashboard.trainingScore}%*\n` +
                   `• Recovery: *${dashboard.recoveryScore}%*\n\n` +
                   `🔍 *DETAIL-ANALYSE*\n`;

      dashboard.scores.forEach(s => {
        if (!["Training Readiness", "Gesamtscore", "Plan Status", "Recovery Score", "Training Score"].includes(s.metrik)) {
          report += `${getEmoji(s.ampel)} ${clean(s.metrik)}: *${s.raw_wert}*\n`;
        }
      });

      report += `\n📊 *NERD-STATS*\n` +
                `• ACWR: *${dashboard.nerdStats.acwr}*\n` +
                `• TSB: *${dashboard.nerdStats.tsb}*\n\n` +
                `💡 *KIRAS FAZIT*\n` +
                `_${clean(dashboard.empfehlungText)}_`;

      sendTelegramMessage(chatId, report);
      
    } catch (e) {
      console.error("Telegram Error: " + e.message);
      sendTelegramMessage(chatId, "⚠️ Abbruch: " + e.message);
    }
  }
}

/**
 * HILFSFUNKTION: Räumt alte Trigger auf
 */
function deleteTriggersForFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function sendTelegramMessage(chatId, text) {
  const cfg = getTelegramConfig_();
  if (!cfg.token) {
    // Fail-fast: better no message than silent success.
    logToSheet('ERROR', '[Telegram] TELEGRAM_BOT_TOKEN fehlt in ScriptProperties');
    return;
  }

  const url = `https://api.telegram.org/bot${cfg.token}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'Markdown'
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    console.error('Telegram Fehler: ' + response.getContentText());
  }
}

// NEU: ZUM SPEICHERN DER SCHLÜSSEL
const STRAVA_CLIENT_ID = '183306'; // ERSETZE DURCH DEINE ID
const STRAVA_CLIENT_SECRET = 'f9e54dc554fa59c054548fbd0791023fde890fe2'; // ERSETZE DURCH DEIN GEHEIMNIS
const STRAVA_SCOPES = 'activity:read_all'; // Erlaubt das Lesen aller Aktivitäten

// ----------------------------------------------------

/**
 * [HILFSFUNKTION] Loggt die benötigte Redirect URI in die Skript-Protokolle.
 */
function logRedirectUri() {
  // Stelle sicher, dass getStravaService() funktioniert und die OAuth2-Bib geladen ist
  const service = getStravaService();
  const redirectUri = service.getRedirectUri();

  Logger.log("--- WICHTIG: STRAVA REDIRECT URI ---");
  Logger.log("Kopiere DIESE URL in das Strava 'Authorization Callback Domain'-Feld:");
  Logger.log(redirectUri);
  SpreadsheetApp.getUi().alert(`Die korrekte Redirect URI wurde in deinen Skript-Protokollen (Logger) hinterlegt. Bitte kopiere sie von dort.`);
}

/**
 * Erstellt den OAuth2-Dienst für Strava (Backend-Funktion).
 */
function getStravaService() {
  // 'OAuth2' ist der Bezeichner der Bibliothek, die du hinzugefügt hast.
  return Strava_OAuth2.createService('Strava')
      .setClientId(STRAVA_CLIENT_ID)
      .setClientSecret(STRAVA_CLIENT_SECRET)
      .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
      .setTokenUrl('https://www.strava.com/oauth/token')
      .setScope(STRAVA_SCOPES)
      .setCallbackFunction('authCallback') // Autorisierungs-Rückruffunktion
      .setPropertyStore(PropertiesService.getUserProperties());
}

/**
 * [Wird durch Strava aufgerufen] Behandelt die Antwort des Autorisierungsservers.
 */
function authCallback(request) {
  const service = getStravaService();
  const authorized = service.handleCallback(request);
  
  if (authorized) {
    return HtmlService.createHtmlOutput('<h2>Autorisierung erfolgreich!</h2><p>Du kannst das Fenster jetzt schließen.</p>');
  } else {
    return HtmlService.createHtmlOutput('<h2>Autorisierung fehlgeschlagen.</h2>');
  }
}

/**
 * [Wird vom Nutzer aufgerufen] Generiert die Autorisierungs-URL.
 */
function showStravaAuthUrl() {
  const service = getStravaService();
  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('Strava ist bereits autorisiert.');
    return;
  }
  const authorizationUrl = service.getAuthorizationUrl();
  Logger.log('Öffne folgende URL in deinem Browser: ' + authorizationUrl);
  // Diese URL muss der Nutzer im Browser öffnen
  SpreadsheetApp.getUi().alert(`Um Strava zu autorisieren, öffne diese URL in deinem Browser: ${authorizationUrl}`);
}

/**
 * Hilfsfunktion zum Finden der Zeile anhand des Datums (aus onFormSubmit übernommen)
 * @returns {number} Die 1-basierte Zeilennummer oder -1, wenn nicht gefunden.
 */
function findRowByDate(sheet, dateHeader, dateString) {
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const dateColIndex = headers.indexOf(dateHeader);
    
    if (dateColIndex === -1) return -1;

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][dateColIndex]).includes(dateString)) {
            return i + 1; // 1-basierte Zeilennummer
        }
    }
    return -1;
}

// --- WICHTIG: Füge Menü-Eintrag hinzu! ---
// Erweitere deine onOpen()-Funktion:
/*
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Coach Kira')
    //... deine anderen Einträge
    .addSeparator()
    .addItem('Strava: Autorisieren', 'showStravaAuthUrl') // Autorisiert das Skript
    .addItem('Strava: Import starten', 'importStravaActivities') // Startet den Import
    .addToUi();
}

/**
 * V57-FIX-3: Startet die KI-basierte Historien-Analyse.
 * UPDATE: Nutzt jetzt den zentralen, sicheren 'callGeminiAPI' Aufruf.
 * Entfernt veraltete URL-Konstruktionen, die Fehler verursachten.
 */
function runHistoricalAnalysis(e) {
  // KI-frei (2026-10-26): Historien-Analyse (Gemini) entfernt.
  logToSheet('INFO', '[History] Historien-Analyse ist entfernt (KI läuft nicht mehr im Apps Script).');
}

// KI-frei (2026-10-26): buildHistoryPrompt entfernt (Prompt-Bau liegt nicht mehr im Apps Script).

// KI-frei (2026-10-26): buildActivityReviewPrompt entfernt (Prompt-Bau liegt nicht mehr im Apps Script).

// --- NEU (V11): GOOGLE KALENDER SYNCHRONISIERUNG ---

/**
 * Holt das Kalender-Objekt anhand der ID aus KK_CONFIG.
 */
function getCalendarService() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("KK_CONFIG");
  if (!configSheet) {
    logToSheet('ERROR', '[Kalender] Config-Blatt "KK_CONFIG" nicht gefunden.');
    return null;
  }
  
  const configData = configSheet.getDataRange().getValues();
  let calendarId = null;
  // Finde die CALENDAR_ID in Spalte A, nimm Wert aus Spalte B
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === "CALENDAR_ID") {
      calendarId = configData[i][1];
      break;
    }
  }

  if (!calendarId) {
    logToSheet('ERROR', '[Kalender] "CALENDAR_ID" nicht in "KK_CONFIG" gefunden.');
    return null;
  }

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    logToSheet('ERROR', `[Kalender] Kalender mit ID "${calendarId}" konnte nicht gefunden/zugegriffen werden.`);
    return null;
  }
  
  logToSheet('DEBUG', `[Kalender] Erfolgreich mit Kalender verbunden: ${calendar.getName()}`);
  return calendar;
}


/**
 * V142-FIX: Bulletproof Sync mit "Noon-Force" Strategie.
 * Setzt alle Datumswerte auf 12:00 Uhr Mittags, um Zeitzonen-Rutscher zu verhindern.
 */
function syncPlanToCalendar() {
  logToSheet('INFO', '[Kalender V142] Starte Sync (Noon-Force Mode)...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Service holen
  const calendarService = getCalendarService();
  if (!calendarService) return;
  const calendarId = calendarService.getId();
  
  // 2. Plan holen
  const planSheet = ss.getSheetByName(OUTPUT_PLAN_SHEET);
  if (!planSheet || planSheet.getLastRow() < 2) return;

  const planData = planSheet.getRange(2, 1, 14, planSheet.getLastColumn()).getValues(); // Max 14 Tage
  const planHeaders = planSheet.getRange(1, 1, 1, planSheet.getLastColumn()).getValues()[0];

  const idx = {
    datum: planHeaders.indexOf('Datum'),
    empfLoad: planHeaders.indexOf('Empfohlener Load (ESS)'),
    zone: planHeaders.indexOf('Empfohlene Zone (KI)'),
    beispiel_1: planHeaders.indexOf('Beispiel-Training 1 (KI)'),
    beispiel_2: planHeaders.indexOf('Beispiel-Training 2 (KI)'),
    sport: planHeaders.indexOf('Prognostizierte Auswirkung (KI)'),
    kcal: planHeaders.indexOf('Empfohlene kcal (KI)'), 
    makros: planHeaders.indexOf('Empfohlene Makros (KI)') 
  };

  if (idx.datum === -1) return;

  const eventPrefix = "[Coach Kira] ";
  let updatedCount = 0;
  let createdCount = 0;
  const timeZone = Session.getScriptTimeZone();

  // 3. Bestehende Events scannen (Wide Search)
  const now = new Date();
  const searchStart = new Date(now.getTime() - 48 * 60 * 60 * 1000); 
  const searchEnd = new Date(now.getTime() + 15 * 24 * 60 * 60 * 1000);
  
  let existingEventsMap = new Map(); 

  try {
    const events = Calendar.Events.list(calendarId, {
      timeMin: searchStart.toISOString(),
      timeMax: searchEnd.toISOString(),
      singleEvents: true
    }).items;

    if (events) {
      events.forEach(ev => {
        if (ev.summary && ev.summary.startsWith(eventPrefix)) {
          // A) Datum aus Kalender extrahieren
          let evDateStr = "";
          if (ev.start.date) {
             evDateStr = ev.start.date; // Ist schon YYYY-MM-DD
          } else if (ev.start.dateTime) {
             evDateStr = ev.start.dateTime.substring(0, 10);
          }
          
          if (evDateStr) {
             // Wir merken uns die ID für diesen Tag
             existingEventsMap.set(evDateStr, ev.id);
          }
        }
      });
    }
  } catch(e) {
    logToSheet('ERROR', '[Kalender] Scan-Fehler: ' + e.message);
    return;
  }

  // 4. Plan abarbeiten
  for (const row of planData) {
    const datumRaw = row[idx.datum];
    if (!datumRaw) continue;

    // --- DER FIX: NOON FORCE ---
    // Wir erstellen ein Date-Objekt und zwingen es auf 12:00 Uhr mittags.
    // Damit verhindern wir, dass 00:00 Uhr durch Zeitzonen auf 23:00 Uhr (Vortag) rutscht.
    let safeDate = new Date(datumRaw);
    safeDate.setHours(12, 0, 0, 0); 
    
    // Jetzt formatieren wir es sicher als String
    let targetDateString;
    try {
        targetDateString = Utilities.formatDate(safeDate, timeZone, "yyyy-MM-dd");
    } catch(e) { continue; }
    // ---------------------------

    const empfLoad = row[idx.empfLoad];
    const zone = (idx.zone !== -1) ? row[idx.zone] : "N/A";
    const beispiel_1 = (idx.beispiel_1 !== -1) ? row[idx.beispiel_1] : "N/A";
    const beispiel_2 = (idx.beispiel_2 !== -1) ? row[idx.beispiel_2] : "N/A";
    const sportArt = row[idx.sport];
    const kcal = (idx.kcal > -1) ? row[idx.kcal] : "-";
    const makros = (idx.makros > -1) ? row[idx.makros] : "-";

    // Titel
    let eventTitle = `${eventPrefix}Training (Load: ${empfLoad} | Zone: ${zone})`;
    if (empfLoad == 0) eventTitle = `${eventPrefix}Ruhetag`;

    const eventDescription = `KI-Vorschläge:\n1) ${beispiel_1}\n2) ${beispiel_2}\n\nDetails:\nLoad: ${empfLoad} | Zone: ${zone}\nStatus: ${sportArt}\n\nErnährung:\nKcal: ${kcal}\nMakros: ${makros}`;

    // Prüfen: Gibt es schon ein Event an diesem "sicheren" Datum?
    let existingEventId = existingEventsMap.get(targetDateString);

    // Sync durchführen (Update oder Create)
    // Wichtig: Wir übergeben das 'safeDate' (12:00 Uhr), damit syncEventToCalendar
    // beim Erstellen eines Ganztages-Events (YYYY-MM-DD) auch den richtigen Tag trifft.
    const syncedId = syncEventToCalendar(calendarId, safeDate, eventTitle, eventDescription, existingEventId);

    if (syncedId) {
      if (existingEventId) updatedCount++;
      else {
          createdCount++;
          // Sofort in Map eintragen, falls der Loop spinnt
          existingEventsMap.set(targetDateString, syncedId);
      }
    }
  }

  logToSheet('INFO', `[Kalender] V142 Fertig: ${updatedCount} aktualisiert, ${createdCount} neu.`);
}

/**
 * (V96-MOD): Erstellt oder aktualisiert ein Kalenderereignis intelligent (via Advanced Service).
 * - NEU: Erstellt ein ganztägiges Event.
 * - UPDATE: Aktualisiert NUR Titel und Beschreibung, um manuell eingetragene
 * Start/End-Uhrzeiten des Benutzers nicht zu überschreiben.
 */
function syncEventToCalendar(calendarId, date, title, description, eventId) {
  const dateString = Utilities.formatDate(date, "UTC", "yyyy-MM-dd");
  let existingEvent;

  if (eventId) {
    // V96-Logik: EVENT EXISTIERT (UPDATE)
    // Wir patchen NUR Titel und Beschreibung.
    logToSheet('DEBUG', `[V96 CAL] Event ${eventId} existiert. Patche nur Titel/Beschreibung.`);
    
    let updatePayload = {
      summary: title,
      description: description
      // WICHTIG: 'start' und 'end' fehlen hier absichtlich!
    };
    
    try {
      // @ts-ignore
      existingEvent = Calendar.Events.patch(updatePayload, calendarId, eventId);
      logToSheet('DEBUG', `[V96 CAL] Patch erfolgreich für ${eventId}.`);
    } catch (e) {
      logToSheet('ERROR', `[V96 CAL] Fehler beim PATCHEN von Event ${eventId}: ${e.message}. Versuche Neuerstellung.`);
      // Fallback: Wenn das Patchen fehlschlägt (z.B. Event wurde manuell gelöscht), 
      // rufen wir die Funktion erneut auf und erzwingen eine Neuerstellung.
      return syncEventToCalendar(calendarId, date, title, description, null); // Erzwingt den 'else'-Block
    }

  } else {
    // V96-Logik: NEUES EVENT (CREATE)
    // Erstelle ein GANZTAGS-Event (wie bisher).
    logToSheet('DEBUG', `[V96 CAL] Event ist neu. Erstelle ganztägigen Eintrag für ${dateString}.`);
    
    let newEvent = {
      summary: title,
      description: description,
      start: { date: dateString },
      end: { date: dateString },
      reminders: {
        useDefault: false
      },
      guestsCanSeeOtherGuests: false,
      guestsCanInviteOthers: false
    };

    try {
      // @ts-ignore
      existingEvent = Calendar.Events.insert(newEvent, calendarId);
    } catch (e) {
      logToSheet('ERROR', `[V96 CAL] Fehler beim ERSTELLEN des Events: ${e.message}`);
      return null; // Konnte nicht erstellt werden
    }
  }

  if (existingEvent) {
    return existingEvent.id;
  } else {
    return null;
  }
}

/**
 * (V11): Startet die Haupt-KI-Analyse (langer Lauf) in 10 Sekunden.
 * (Wird von der Web App aufgerufen, um das 30s-Limit zu umgehen)
 */
function triggerFullAnalysis_WebApp() {
  try {
    // Lösche alte Trigger, falls vorhanden
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "runKiraGeminiSupervisor") {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    
    // Erstelle einen neuen Trigger, der in 10 Sekunden einmalig läuft
    ScriptApp.newTrigger("runKiraGeminiSupervisor")
      .timeBased()
      .after(2 * 1000) // 10 Sekunden
      .create();
      
    logToSheet('INFO', '[WebApp] Asynchroner Trigger für runKiraGeminiSupervisor in 10s erstellt.');
    return "KI-Analyse (voller Report) in 10 Sekunden gestartet. Ergebnisse erscheinen in Kürze in den Sheets.";
      
  } catch (e) {
    logToSheet('ERROR', `[WebApp] Fehler beim Erstellen des Triggers: ${e.message}`);
    return `Fehler beim Starten der Analyse: ${e.message}`;
  }
}

/**
 * NEU (V13): Setzt den 'is_today'-Marker (1) manuell eine Zeile nach unten.
 * Schreibt direkt in das 'timeline' (SOURCE) Blatt, NICHT in KK_TIMELINE.
 * Dies simuliert einen Tageswechsel.
 */
function advanceTodayFlag() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // WICHTIG: Wir schreiben in die Formel-Quelle 'timeline'
  const sheet = ss.getSheetByName(SOURCE_TIMELINE_SHEET); 
  if (!sheet) {
    logToSheet('ERROR', `[advanceTodayFlag] Quellblatt '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);
    SpreadsheetApp.getUi().alert(`Fehler: Blatt '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const isTodayIndex = headers.indexOf('is_today'); // 0-basierter Spalten-Index

  if (isTodayIndex === -1) {
    logToSheet('ERROR', `[advanceTodayFlag] Spalte 'is_today' nicht in '${SOURCE_TIMELINE_SHEET}' gefunden.`);
    SpreadsheetApp.getUi().alert(`Fehler: Spalte 'is_today' nicht in '${SOURCE_TIMELINE_SHEET}' gefunden.`);
    return;
  }

  let currentRowArrayIndex = -1; // 0-basierter Zeilen-Index im Array
  for (let i = 1; i < data.length; i++) {
    if (data[i][isTodayIndex] == 1) {
      currentRowArrayIndex = i;
      break;
    }
  }

  if (currentRowArrayIndex === -1) {
    logToSheet('ERROR', `[advanceTodayFlag] Keine Zeile mit 'is_today=1' gefunden.`);
    SpreadsheetApp.getUi().alert(`Fehler: Keine Zeile mit 'is_today=1' gefunden.`);
    return;
  }

  if (currentRowArrayIndex + 1 >= data.length) {
    logToSheet('ERROR', `[advanceTodayFlag] 'is_today=1' ist bereits in der letzten Zeile. Kann nicht verschoben werden.`);
    SpreadsheetApp.getUi().alert(`Fehler: 'is_today=1' ist bereits in der letzten Zeile.`);
    return;
  }

  // Apps Script Ranges sind 1-basiert.
  const oldSheetRow = currentRowArrayIndex + 1; // 1-basierte Zeile (z.B. 69)
  const newSheetRow = oldSheetRow + 1;          // 1-basierte Zeile (z.B. 70)
  const sheetCol = isTodayIndex + 1;          // 1-basierte Spalte

  try {
    // 1. Alte Zeile auf 0 setzen
    sheet.getRange(oldSheetRow, sheetCol).setValue(0);
    // 2. Nächste Zeile auf 1 setzen
    sheet.getRange(newSheetRow, sheetCol).setValue(1);

    const newDate = sheet.getRange(newSheetRow, headers.indexOf('date') + 1).getDisplayValue();
    logToSheet('INFO', `[advanceTodayFlag] 'is_today' erfolgreich von Zeile ${oldSheetRow} auf ${newSheetRow} (Datum: ${newDate}) verschoben.`);
    SpreadsheetApp.getUi().alert(`'is_today' erfolgreich verschoben auf: ${newDate}`);

  } catch (e) {
    logToSheet('ERROR', `[advanceTodayFlag] Fehler beim Schreiben der Werte: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Fehler beim Verschieben des 'is_today'-Flags: ${e.message}`);
  }
}

/**
 * (V2 2026-09-27) Setzt 'is_today' in 'timeline' auf die Zeile mit dem HEUTIGEN Datum.
 * Idempotent: mehrfacher Aufruf am selben Tag ändert nichts. Wird von onFormSubmit ("Morgens") genutzt.
 * @returns {boolean} true, wenn die heutige Zeile gefunden und markiert wurde.
 */
function advanceTodayFlag_silent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SOURCE_TIMELINE_SHEET); // 'timeline'
  if (!sheet) {
    logToSheet('ERROR', `[advanceTodayFlag_silent] Quellblatt '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);
    return false;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const isTodayIdx = headers.indexOf('is_today');
  const dateIdx = headers.indexOf('date');
  if (isTodayIdx < 0 || dateIdx < 0) {
    logToSheet('ERROR', `[advanceTodayFlag_silent] Spalte 'is_today' oder 'date' fehlt in '${SOURCE_TIMELINE_SHEET}'.`);
    return false;
  }

  const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
  const todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const dates = sheet.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues();
  const flags = sheet.getRange(2, isTodayIdx + 1, lastRow - 1, 1).getValues();

  let todayRow = -1;
  let oldRow = -1;
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i][0];
    const key = (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || '').trim().slice(0, 10);
    if (key === todayKey) todayRow = i + 2;
    if (Number(String(flags[i][0]).replace(',', '.')) === 1 && oldRow < 0) oldRow = i + 2;
  }
  if (todayRow < 0) {
    logToSheet('ERROR', `[advanceTodayFlag_silent] Keine Zeile für heute (${todayKey}) in '${SOURCE_TIMELINE_SHEET}'.`);
    return false;
  }
  if (oldRow === todayRow) {
    logToSheet('INFO', `[advanceTodayFlag_silent] 'is_today' steht bereits auf heute (Zeile ${todayRow}) – keine Änderung.`);
    return true;
  }

  try {
    const out = flags.map((_, i) => [i + 2 === todayRow ? 1 : 0]);
    sheet.getRange(2, isTodayIdx + 1, out.length, 1).setValues(out);
    logToSheet('INFO', `[advanceTodayFlag_silent] 'is_today' von Zeile ${oldRow} auf ${todayRow} (${todayKey}) gesetzt.`);
    return true;
  } catch (e) {
    logToSheet('ERROR', `[advanceTodayFlag_silent] Fehler beim Schreiben: ${e.message}`);
    return false;
  }
}


/**
 * NEU (V13.5): Startet die Historien-Analyse (langer Lauf) in 10 Sekunden.
 * (Wird von der Web App aufgerufen, um das 30s-Limit zu umgehen)
 */
function triggerHistoryAnalysis_WebApp() {
  // KI-frei (2026-10-26): Historien-Analyse (Gemini) entfernt.
  return 'Die Historien-Analyse läuft nicht mehr im Apps Script. Langzeit-Auswertungen findest du im Cockpit (Analyse, Berichte).';
}



/**
 * (V17 - Unverändert zu V14): Lädt die Zonen-Load-Faktoren (Punkte/Stunde).
 * Nutzt den Cache.
 */
function getLoadFactors() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'load_factors_v1';
  
  let factorsJson = cache.get(cacheKey);
  if (factorsJson) {
    // logToSheet('DEBUG', '[getLoadFactors] Lade Zonen-Faktoren aus Cache.');
    return JSON.parse(factorsJson);
  }

  // Fallback: persistenter Cache (überlebt Cache-Evictions)
const props = PropertiesService.getDocumentProperties();
factorsJson = props.getProperty(cacheKey);
if (factorsJson) {
  cache.put(cacheKey, factorsJson, 21600);
  return JSON.parse(factorsJson);
}

  logToSheet('INFO', '[getLoadFactors] Lade Zonen-Faktoren aus Sheet ' + LOAD_FACTOR_SHEET_NAME + '...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOAD_FACTOR_SHEET_NAME);
  if (!sheet) {
    logToSheet('ERROR', `[getLoadFactors] Sheet '${LOAD_FACTOR_SHEET_NAME}' nicht gefunden! Nutze Fallback-Werte.`);
    return {
      "Run": {"Z1": 70, "Z2": 105, "Z3": 160, "Z4": 215, "Z5": 260, "Default": 105},
      "Bike": {"Z1": 55, "Z2": 85, "Z3": 130, "Z4": 170, "Z5": 220, "Default": 85},
      "Other": {"Z1": 40, "Z2": 60, "Default": 50}
    };
  }

  const data = sheet.getDataRange().getValues();
  const factors = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const activityType = String(row[0]).trim(); 
    const zone = String(row[1]).trim();         
    const factor = parseFloat(String(row[2]).replace(',', '.')); 

    if (activityType && zone && !isNaN(factor)) {
      if (!factors[activityType]) {
        factors[activityType] = {};
      }
      factors[activityType][zone] = factor;
    }
  }
  cache.put(cacheKey, JSON.stringify(factors), 21600); // 6h Cache
  props.setProperty(cacheKey, JSON.stringify(factors));
  logToSheet('INFO', '[getLoadFactors] Zonen-Faktoren erfolgreich geladen und im Cache gespeichert.');
  return factors;
}

function getElevFactors() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'elev_factors_v1';
  const props = PropertiesService.getDocumentProperties(); // <-- FIX: props definieren

  // 1) Cache
  let factorsJson = cache.get(cacheKey);
  if (factorsJson) return JSON.parse(factorsJson);

  // 2) Persistenter Fallback (DocumentProperties)
  factorsJson = props.getProperty(cacheKey);
  if (factorsJson) {
    cache.put(cacheKey, factorsJson, 21600); // 6h
    return JSON.parse(factorsJson);
  }

  // 3) Sheet lesen
  logToSheet('INFO', '[getElevFactors] Lade Höhen-Faktoren aus Sheet KK_ELEV_FACTORS...');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ELEV_FACTOR_SHEET_NAME);
  if (!sheet) throw new Error(`Sheet '${ELEV_FACTOR_SHEET_NAME}' nicht gefunden.`);

  const data = sheet.getDataRange().getValues();
  const factors = {};

  for (let i = 1; i < data.length; i++) { // skip header
    const activity = String(data[i][0] || '').trim();
    const elevBand = String(data[i][1] || '').trim();
    const raw = data[i][2];

    const factor = parseFloat(String(raw ?? '').replace(',', '.'));
    if (!activity || !elevBand || isNaN(factor)) continue;

    if (!factors[activity]) factors[activity] = {};
    factors[activity][elevBand] = factor;
  }

  // 4) Cache + Properties schreiben
  const json = JSON.stringify(factors);
  cache.put(cacheKey, json, 21600);      // 6h
  props.setProperty(cacheKey, json);     // persistenter Cache

  logToSheet('INFO', '[getElevFactors] Höhen-Faktoren erfolgreich geladen und im Cache gespeichert.');
  return factors;
}


/**
 * Lädt das Wochen-Template (CSV) aus 'week_config' und cached es.
 * @returns {string}
 */
function getWeekConfig() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'week_config_v1';

  const cached = cache.get(cacheKey);
  if (cached) return cached;

  logToSheet('INFO', '[getWeekConfig] Lade Wochen-Template aus Sheet week_config...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('week_config');
  if (!sheet) throw new Error("Sheet 'week_config' nicht gefunden.");

  const values = sheet.getDataRange().getDisplayValues();
  const csvString = values.map(row =>
    row.map(cell => `"${String(cell ?? '').replace(/"/g, '""')}"` ).join(',')
  ).join('\n');

  cache.put(cacheKey, csvString, 21600); // 6h Cache
  // Optionaler Backup:
  // PropertiesService.getDocumentProperties().setProperty(cacheKey, csvString);

  logToSheet('INFO', '[getWeekConfig] Wochen-Template erfolgreich geladen und im Cache gespeichert.');
  return csvString;
}

/**
 * Schätzt den Trainings-Load (EPOC/ESS) basierend auf Dauer, Zone UND (optional) Höhenmetern.
 * (V17 - Macht 'elevationMeters' optional und setzt Standard auf 0)
 *
 * @param {string} activityType - "Run", "Bike (Commute)", "Hike", etc.
 * @param {number} durationMinutes - z.B. 60
 * @param {string} zone - z.B. "Z2"
 * @param {number} [elevationMeters=0] - z.B. 500 (Optional, Standard ist 0)
 * @returns {number} Der geschätzte Gesamt-Load.
 */
function estimateLoad(activityType, durationMinutes, zone, elevationMeters) {
  // --- NEU (V17): Setze Standardwert für Höhenmeter ---
  const elevM = elevationMeters || 0;
  // --- ENDE NEU ---

  // 1. Lade beide Faktor-Sets (aus Cache oder Sheet)
  const zoneFactors = getLoadFactors();
  const elevFactors = getElevFactors();

  // 2. Bestimme den Aktivitätstyp-Schlüssel (Run, Bike, Other)
  let factorKey = "Other"; // Standard-Fallback
  if (activityType.toLowerCase().includes("run") || activityType.toLowerCase().includes("row") || activityType.toLowerCase().includes("hiit")) {
    factorKey = "Run";
  } else if (activityType.toLowerCase().includes("bike")) {
    factorKey = "Bike";
  }

  // 3. Berechne den Basis-Load (Dauer * Zone)
  let baseFactor = 0;
  if (zoneFactors[factorKey]) {
    if (zoneFactors[factorKey][zone]) {
      baseFactor = zoneFactors[factorKey][zone];
    } else {
      baseFactor = zoneFactors[factorKey]["Default"] || zoneFactors["Other"]["Default"];
    }
  } else {
    baseFactor = zoneFactors["Other"]["Default"];
  }
  const baseLoad = (durationMinutes / 60) * baseFactor;

  // 4. Berechne den Höhen-Load
  let elevFactor = elevFactors[factorKey] || elevFactors["Other"];
  const elevationLoad = (elevM / 100) * elevFactor; // Nutzt den neuen 'elevM'-Wert

  // 5. Gesamter Load
  const totalLoad = Math.round(baseLoad + elevationLoad);
  
  logToSheet('DEBUG', `[estimateLoad V17] ${activityType}, ${durationMinutes}min, ${zone}, ${elevM}hm -> Base: ${Math.round(baseLoad)} (F: ${baseFactor}) + Elev: ${Math.round(elevationLoad)} (F: ${elevFactor}) = ${totalLoad}`);
  return totalLoad;
}

/**
 * NEU (V20): Kopiert die "WAHRHEIT" (IST-Werte) in die SOLL/Forecast-Spalten.
 * Wird nach "Nach Aktivität" aufgerufen, um die Timeline für den nächsten KI-Lauf zu "nullen".
 * SCHREIBT DIREKT IN 'timeline' (SOURCE_TIMELINE_SHEET).
 * @param {number} rowIndex - Der 1-basierte Zeilenindex (z.B. 70), der aktualisiert werden soll.
 */
function copyObservedToForecast(rowIndex) {
  if (!rowIndex || rowIndex < 2) {
    logToSheet('ERROR', `[copyObservedToForecast] Ungültiger Zeilenindex: ${rowIndex}`);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // WICHTIG: Definiere SOURCE_TIMELINE_SHEET oben im Skript, falls noch nicht geschehen, 
  // z.B. const SOURCE_TIMELINE_SHEET = 'timeline';
  const sheet = ss.getSheetByName(SOURCE_TIMELINE_SHEET); // 'timeline'
  if (!sheet) {
    logToSheet('ERROR', `[copyObservedToForecast] Quellblatt '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Spalten finden
  const cols = {
    load_fb_day: headers.indexOf('load_fb_day') + 1,
    fbATL_obs: headers.indexOf('fbATL_obs') + 1,
    fbCTL_obs: headers.indexOf('fbCTL_obs') + 1,
    fbACWR_obs: headers.indexOf('fbACWR_obs') + 1,
    
    coachE_ESS_day: headers.indexOf('coachE_ESS_day') + 1,
    coachE_ATL_forecast: headers.indexOf('coachE_ATL_forecast') + 1,
    coachE_CTL_forecast: headers.indexOf('coachE_CTL_forecast') + 1,
    coachE_ACWR_forecast: headers.indexOf('coachE_ACWR_forecast') + 1
  };

  // Prüfen, ob alle Spalten da sind
  for (const key in cols) {
    if (cols[key] === 0) {
      logToSheet('ERROR', `[copyObservedToForecast] Kritische Spalte '${key}' nicht in '${SOURCE_TIMELINE_SHEET}' gefunden.`);
      return;
    }
  }

  try {
    logToSheet('INFO', `[copyObservedToForecast] Kopiere "WAHRHEIT" in Zeile ${rowIndex}...`);
    
    // Werte lesen (IST)
    const load_ist = sheet.getRange(rowIndex, cols.load_fb_day).getValue();
    const atl_ist = sheet.getRange(rowIndex, cols.fbATL_obs).getValue();
    const ctl_ist = sheet.getRange(rowIndex, cols.fbCTL_obs).getValue();
    const acwr_ist = sheet.getRange(rowIndex, cols.fbACWR_obs).getDisplayValue(); // Wichtig: Nimm DisplayValue (Text mit Komma)

    // Werte schreiben (in SOLL)
    sheet.getRange(rowIndex, cols.coachE_ESS_day).setValue(load_ist);
    sheet.getRange(rowIndex, cols.coachE_ATL_forecast).setValue(atl_ist);
    sheet.getRange(rowIndex, cols.coachE_CTL_forecast).setValue(ctl_ist);
    sheet.getRange(rowIndex, cols.coachE_ACWR_forecast).setValue(acwr_ist);

    logToSheet('INFO', `[copyObservedToForecast] "WAHRHEIT" (Load: ${load_ist}, ATL: ${atl_ist}, CTL: ${ctl_ist}) erfolgreich in Forecast-Spalten (Zeile ${rowIndex}) geschrieben.`);
    
  } catch (e) {
    logToSheet('ERROR', `[copyObservedToForecast] Fehler beim Kopieren der Werte: ${e.message}`);
  }
}

/**
 * NEU (V28): Asynchroner Trigger für den Activity Review.
 * Wird von onFormSubmit aufgerufen, um Timeouts beim Formular-Senden zu vermeiden.
 */
function triggerActivityReview() {
  // KI-frei (2026-10-26): Den Aktivitätskommentar schreibt der Mac (coach-kira) selbst.
  logToSheet('INFO', '[ActivityReview] Läuft lokal auf dem Mac – kein Apps-Script-Trigger mehr.');
  return true; // Wichtig für das Form-Skript
}

/**
 * NEU (V35.2): Übernimmt den 8-Tage-KI-Plan aus AI_REPORT_PLAN
 * und schreibt ihn in 'timeline' (coachE_ESS_day) als neue Basis.
 * SICHERT die alten Werte zuerst in 'AI_PLAN_BACKUP'.
 * (V35.1: Entfernt ui.alert-Befehle für WebApp-Kompatibilität)
 * (V35.2: KORREKTUR der Konstantennamen: _SHEET statt _SHT)
 */
function applyKiraPlanToTimeline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // (Falls du die Konstante nicht global hinzugefügt hast, 
  // wird sie hier als Fallback definiert)
  const BACKUP_SHEET = 'AI_PLAN_BACKUP'; 

  logToSheet('INFO', '[V35.2] Starte Funktion: Übernehme KI-Plan (mit Backup)...');

  try {
    // 1. Plan-Daten (Quelle) lesen
    // KORREKTUR (V35.2): OUTPUT_PLAN_SHEET
    const planSheet = ss.getSheetByName(OUTPUT_PLAN_SHEET); // "AI_REPORT_PLAN"
    if (!planSheet || planSheet.getLastRow() < 2) {
      throw new Error(`Blatt '${OUTPUT_PLAN_SHEET}' ist leer oder nicht gefunden.`);
    }

    const planHeaders = planSheet.getRange(1, 1, 1, planSheet.getLastColumn()).getValues()[0];
    const loadColIndex = planHeaders.indexOf('Empfohlener Load (ESS)');
    
    if (loadColIndex === -1) {
      throw new Error(`Spalte 'Empfohlener Load (ESS)' in '${OUTPUT_PLAN_SHEET}' nicht gefunden.`);
    }

    // Lese die 8 NEUEN Load-Werte
    const planLoads_new = planSheet.getRange(2, loadColIndex + 1, 8, 1).getValues();
    // Wandle in flaches Array für das Logging um
    const newValuesFlat = planLoads_new.map(r => r[0]); 
    logToSheet('DEBUG', `[V35.2] 8 Tage NEUEN KI-Plan gelesen (Werte: ${JSON.stringify(newValuesFlat)})`);

    // 2. Timeline-Blatt (Ziel) finden
    // KORREKTUR (V35.2): SOURCE_TIMELINE_SHEET
    const targetSheet = ss.getSheetByName(SOURCE_TIMELINE_SHEET); // "timeline"
    if (!targetSheet) {
      throw new Error(`Zielblatt '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);
    }

    const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    const targetLoadColIndex = targetHeaders.indexOf('coachE_ESS_day');
    const targetIsTodayIndex = targetHeaders.indexOf('is_today');
    const targetDateIndex = targetHeaders.indexOf('date'); // Für Backup-Info

    if (targetLoadColIndex === -1 || targetIsTodayIndex === -1 || targetDateIndex === -1) {
      throw new Error(`Spalten 'coachE_ESS_day', 'is_today' oder 'date' in '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);
    }

    // 3. Finde 'is_today=1' Zeile
    const data = targetSheet.getDataRange().getValues(); // Hole alle Daten
    let startRowIndex = -1; // 0-basierter Index
    for (let i = 1; i < data.length; i++) { // Start bei 1 (Header überspringen)
      if (data[i][targetIsTodayIndex] == 1) {
        startRowIndex = i;
        break;
      }
    }

    if (startRowIndex === -1) {
      // KORREKTUR (V35.2): SOURCE_TIMELINE_SHEET
      throw new Error(`Keine Zeile mit 'is_today=1' in '${SOURCE_TIMELINE_SHEET}' gefunden.`);
    }
    
    const targetSheetStartRow = startRowIndex + 1; // 1-basierte Zeile für .getRange()
    // Hole das Startdatum für das Log
    const startDate = data[startRowIndex][targetDateIndex]; 
    logToSheet('DEBUG', `[V35.2] 'is_today=1' gefunden in Zeile ${targetSheetStartRow} (Datum: ${startDate})`);

    // 4. LESEN der 8 ALTEN Werte (BACKUP)
    const targetRange = targetSheet.getRange(targetSheetStartRow, targetLoadColIndex + 1, 8, 1);
    const planLoads_old = targetRange.getValues();
    const oldValuesFlat = planLoads_old.map(r => r[0]); // Flaches Array für Backup
    logToSheet('DEBUG', `[V35.2] 8 Tage ALTEN Plan gelesen (Werte: ${JSON.stringify(oldValuesFlat)})`);

    // 5. SCHREIBEN des Backups
    let backupSheet = ss.getSheetByName(BACKUP_SHEET);
    if (!backupSheet) {
      backupSheet = ss.insertSheet(BACKUP_SHEET);
      // Header für das neue Blatt
      backupSheet.appendRow(["Timestamp", "Aktion", "Startdatum", "Tag 1", "Tag 2", "Tag 3", "Tag 4", "Tag 5", "Tag 6", "Tag 7", "Tag 8"]);
      logToSheet('INFO', `[V35.2] Backup-Blatt '${BACKUP_SHEET}' wurde erstellt.`);
    }
    
    const timestamp = new Date();
    // Schreibe 2 Zeilen ins Backup-Log: Alter Plan und Neuer Plan
    backupSheet.appendRow([timestamp, "ALTER PLAN (Überschrieben)", startDate, ...oldValuesFlat]);
    backupSheet.appendRow([timestamp, "NEUER PLAN (Übernommen)", startDate, ...newValuesFlat]);
    logToSheet('INFO', `[V35.2] Backup der alten und neuen Werte in '${BACKUP_SHEET}' geschrieben.`);

    // 6. SCHREIBEN der 8 NEUEN Werte in die 'timeline' (Der eigentliche Befehl)
    targetRange.setValues(planLoads_new);

    // KORREKTUR (V35.2): SOURCE_TIMELINE_SHEET
    logToSheet('INFO', `[V35.2] KI-Plan (8 Tage) erfolgreich in '${SOURCE_TIMELINE_SHEET}' (ab Zeile ${targetSheetStartRow}) geschrieben.`);
    
    // --- NEU (V35.1): Gib Erfolgs-Text an die WebApp zurück ---
    return `KI-Plan übernommen (V35.2)! 'timeline' aktualisiert, Backup in '${BACKUP_SHEET}' erstellt.`;

  } catch (e) {
    logToSheet('ERROR', `[V35.2] Fehler beim Übernehmen des KI-Plans: ${e.message} \nStack: ${e.stack}`);
    
    // --- NEU (V35.1): Wirf den Fehler, damit die WebApp ihn fängt ---
    throw new Error(`Plan konnte nicht übernommen werden: ${e.message}`);
  }
}

// KI-frei (Stufe 3): createDashboardPrompt_V39 entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

/**
 * (V108-FIX): Ersetzt die V42 Open-Meteo-Funktion, die Fehler meldet.
 * Nutzt stattdessen die OpenWeatherMap (OWM) 3.0 OneCall-API.
 * Benötigt einen 'OPENWEATHER_API_KEY' in den Skripteigenschaften.
 * Konvertiert m/s Wind in km/h, um das alte Format [cite: 848] für die KI beizubehalten.
 */
function getWetterdaten() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("KK_CONFIG");

  // 1. Hole den NEUEN Key
  const API_KEY = PropertiesService.getScriptProperties().getProperty('OPENWEATHER_API_KEY');
  if (!API_KEY) {
    logToSheet('ERROR', '[Wetter V108] OPENWEATHER_API_KEY nicht in Skripteigenschaften gefunden.');
    return "Wetterdaten nicht verfügbar (API-Key fehlt).";
  }

  // 2. Hole Koordinaten (wie V42 [cite: 839-840])
  if (!configSheet) {
    logToSheet('WARN', '[Wetter V108] KK_CONFIG nicht gefunden. Überspringe Wetter.');
    return "Wetterdaten nicht verfügbar (KK_CONFIG fehlt).";
  }
  const configData = configSheet.getDataRange().getValues();
  let lat, lon;
  configData.forEach(row => {
    if (row[0] === 'LAT') lat = String(row[1]).replace(',', '.');
    if (row[0] === 'LON') lon = String(row[1]).replace(',', '.');
  });

  if (!lat || !lon) {
    logToSheet('WARN', `[Wetter V108] LAT/LON in KK_CONFIG nicht gefunden.`);
    return "Wetterdaten nicht verfügbar (LAT/LON fehlt).";
  }

  // 3. Baue die OWM OneCall API 3.0 URL
  const url = `https://api.openweathermap.org/data/3.0/onecall?lat=${lat}&lon=${lon}&exclude=minutely,hourly,alerts&appid=${API_KEY}&units=metric&lang=de`;
  
  try {
    const response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const data = JSON.parse(responseBody);
      
      if (!data.daily || data.daily.length < 2) {
        throw new Error("OWM API-Antwort enthält keine 'daily'-Daten.");
      }

      // 4. Parse die OWM-Antwort (die anders ist als Open-Meteo)
      const heute = data.daily[0];
      const morgen = data.daily[1];
      
      // Hilfsfunktion, da OWM Unix-Timestamps (Sekunden) sendet
      const formatDate = (timestamp) => Utilities.formatDate(new Date(timestamp * 1000), "UTC", "yyyy-MM-dd");
      
      const heuteDesc = heute.weather[0] ? heute.weather[0].description : 'N/A';
      const morgenDesc = morgen.weather[0] ? morgen.weather[0].description : 'N/A';

      // 5. Baue den exakt gleichen String wie V42 [cite: 846-850]
      let wetterString = "HEUTE (" + formatDate(heute.dt) + "):\n";
      wetterString += `  - Max Temp: ${heute.temp.max.toFixed(0)}°C\n`;
      wetterString += `  - Min Temp: ${heute.temp.min.toFixed(0)}°C\n`;
      // OWM gibt m/s, wir rechnen es in km/h um [cite: 848]
      wetterString += `  - Max Wind: ${(heute.wind_speed * 3.6).toFixed(0)} km/h\n`; 
      wetterString += `  - Wetter-Code: ${heute.weather[0].id} (${heuteDesc})\n`;
      
      wetterString += "MORGEN (" + formatDate(morgen.dt) + "):\n";
      wetterString += `  - Max Temp: ${morgen.temp.max.toFixed(0)}°C\n`;
      wetterString += `  - Min Temp: ${morgen.temp.min.toFixed(0)}°C\n`;
      wetterString += `  - Max Wind: ${(morgen.wind_speed * 3.6).toFixed(0)} km/h\n`;
      wetterString += `  - Wetter-Code: ${morgen.weather[0].id} (${morgenDesc})\n`;
      
      logToSheet('INFO', `[Wetter V108] OpenWeatherMap-Daten abgerufen: ${wetterString.replace(/\n/g, ' ')}`);
      return wetterString;

    } else {
      logToSheet('ERROR', `[Wetter V108] API-Fehler (Code ${responseCode}): ${responseBody}`);
      return "Wetter-API-Fehler (OWM).";
    }
  } catch (e) {
    logToSheet('ERROR', `[Wetter V108] Schwerer Fehler beim Abruf: ${e.message}`);
    return "Wetter-Abruf fehlgeschlagen (OWM).";
  }
}

/**
 * V46: Chat-Backend mit Gedächtnis (Read & Write).
 */
function askKira(frage) {
  // KI-frei (2026-10-26): Gemini-Aufrufe entfernt. Fragen: Cockpit-Chat oder Telegram (lokal auf dem Mac).
  return 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.';
}

// KI-frei (2026-10-26): createChatPrompt_V45 entfernt (Prompt-Bau liegt nicht mehr im Apps Script).

/**
 * NEU (V128): Berechnet Recovery- und Training-Subscores.
 * UPDATE: "KEI" wurde zum Training-Score hinzugefügt!
 */
function calculateSubScores(fitnessScores) {
    let recoveryMetrics = ["RHR", "Schlafdauer", "Schlafscore", "HRV Status"];
    
    // HIER IST DAS UPDATE: "KEI" hinzufügen
    let trainingMetrics = [  
        "TE Balance (% Intensiv)", 
        "ACWR (Forecast)", 
        "Training Status",
        "KEI" // <--- NEU!
    ];

    let recoveryTotalScore = 0;
    let recoveryTotalWeight = 0;
    let trainingTotalScore = 0;
    let trainingTotalWeight = 0;

    if (!fitnessScores || fitnessScores.length === 0) {
        return { recoveryScore: 0, trainingScore: 0, recoveryAmpel: "ROT", trainingAmpel: "ROT" };
    }

    fitnessScores.forEach(s => {
        // Überspringe ungültige/offene Werte
        if (s.ampel === "OFFEN") return;

        const score = s.num_score || 0;
        
        // Recovery Score Calculation (HRV x2)
        if (recoveryMetrics.includes(s.metrik)) {
            if (s.metrik === "HRV Status") {
                recoveryTotalScore += score * 2;
                recoveryTotalWeight += 2;
            } else {
                recoveryTotalScore += score;
                recoveryTotalWeight += 1;
            }
        }
        
        // Training Score Calculation
        if (trainingMetrics.includes(s.metrik)) {
            // Optional: Willst du KEI höher gewichten? 
            // Aktuell zählt er einfach (Faktor 1) wie alle anderen.
            trainingTotalScore += score;
            trainingTotalWeight += 1;
        }
    });

    const recoveryScoreNum = (recoveryTotalWeight > 0) ? (recoveryTotalScore / recoveryTotalWeight).toFixed(0) : 0;
    const trainingScoreNum = (trainingTotalWeight > 0) ? (trainingTotalScore / trainingTotalWeight).toFixed(0) : 0;

    return { 
        recoveryScore: parseFloat(recoveryScoreNum), 
        trainingScore: parseFloat(trainingScoreNum),
        recoveryAmpel: getGesamtAmpel(recoveryScoreNum),
        trainingAmpel: getGesamtAmpel(trainingScoreNum)
    };
}

/**
 * V87.4-R1C1-FIX: Speichert Plan-Änderungen UND kopiert Formeln via R1C1.
 * Nutzt setFormulaR1C1 statt copyTo, um Abstürze zu vermeiden.
 */
function updatePlannedLoad(planData) {
  logToSheet('INFO', `[V87.4 Update] Empfange ${planData.length} Updates. Erstes Datum: ${planData[0].datum}`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = (typeof SOURCE_TIMELINE_SHEET !== 'undefined') ? SOURCE_TIMELINE_SHEET : 'timeline';
    const timelineSheet = ss.getSheetByName(sheetName);
    
    if (!timelineSheet) throw new Error(`Blatt '${sheetName}' nicht gefunden.`);

    const range = timelineSheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];

    // Indizes für Werte
    const dateIndex = headers.indexOf('date');
    const loadIndex = headers.indexOf('coachE_ESS_day');
    const sportIndex = headers.indexOf('Sport_x');
    const zoneIndex = headers.indexOf('Zone');
    const teAeIndex = headers.indexOf('Target_Aerobic_TE');
    const teAnIndex = headers.indexOf('Target_Anaerobic_TE');

    const findHeaderIndex = (names) => {
      for (const name of names) {
        const idx = headers.indexOf(name);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    // Indizes für FORMELN (Werden NICHT überschrieben)
    const formulaCols = [
      // coachE_ATL_forecast, coachE_CTL_forecast, coachE_ACWR_forecast wurden hier entfernt
      headers.indexOf('Monotony7'),
      headers.indexOf('Strain7'),
      headers.indexOf('coachE_ATL_morning'),
      headers.indexOf('coachE_CTL_morning'),
      findHeaderIndex(['fKEI', 'fkei']),
      findHeaderIndex(['Load*', 'load_star', 'load*']),
      findHeaderIndex(['CTL*', 'ctl_star', 'ctl*']),
      findHeaderIndex(['ctl_gap', 'CTL-CTL*', 'CTL−CTL*'])
    ].filter(idx => idx !== -1);

    // Helper: Datum normalisieren
    const normalizeDate = (d) => {
      if (!d) return null;
      if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (typeof d === 'string') {
        if (d.includes('-')) return d.split('T')[0];
        const parts = d.split('.');
        if (parts.length === 3) return `${parts[2]}-${parts[1]}-${parts[0]}`;
      }
      return String(d);
    };

    const updateMap = new Map();
    planData.forEach(item => updateMap.set(item.datum, item));

    let changesMade = 0;

    // Loop durch das Sheet (Start ab Zeile 2)
    for (let i = 1; i < values.length; i++) {
      const sheetDateRaw = values[i][dateIndex];
      if (!sheetDateRaw) continue;

      const sheetDateStr = normalizeDate(sheetDateRaw);

      if (updateMap.has(sheetDateStr)) {
        const update = updateMap.get(sheetDateStr);
        let rowChanged = false;
        const currentRow = i + 1; // 1-basiert

        const toSheetFmt = (val) => String(val).replace('.', ',');

        // 1. WERTE SCHREIBEN
        const setCell = (colIdx, newVal, oldVal) => {
          if (colIdx === -1) return;
          let valToWrite = newVal;
          if ([loadIndex, teAeIndex, teAnIndex].includes(colIdx)) {
            valToWrite = toSheetFmt(newVal);
          }
          if (String(valToWrite) != String(toSheetFmt(oldVal))) {
            timelineSheet.getRange(currentRow, colIdx + 1).setValue(valToWrite);
            rowChanged = true;
          }
        };

        setCell(loadIndex, update.load, values[i][loadIndex]);
        setCell(sportIndex, update.sport, values[i][sportIndex]);
        setCell(zoneIndex, update.zone, values[i][zoneIndex]);
        setCell(teAeIndex, update.te_ae, values[i][teAeIndex]);
        setCell(teAnIndex, update.te_an, values[i][teAnIndex]);

        // 2. FORMELN KOPIEREN (R1C1 METHODE - ROBUST)
        if (i > 1 && formulaCols.length > 0) { 
           const prevRow = i; // Zeile davor (Sheet Row)
           
           formulaCols.forEach(colIdx => {
             // WICHTIG: Wir holen die Formel in R1C1-Notation (relativ)
             const sourceRange = timelineSheet.getRange(prevRow, colIdx + 1);
             const formulaR1C1 = sourceRange.getFormulaR1C1();
             
             // Nur schreiben, wenn es wirklich eine Formel ist
             if (formulaR1C1) {
                const targetRange = timelineSheet.getRange(currentRow, colIdx + 1);
                // Checken, ob wir überschreiben müssen (Performance)
                if (targetRange.getFormulaR1C1() !== formulaR1C1) {
                    targetRange.setFormulaR1C1(formulaR1C1);
                }
             }
           });
        }

        if (rowChanged) changesMade++;
      }
    }

    if (changesMade > 0) {
      logToSheet('INFO', `[V87.4] Erfolgreich: ${changesMade} Tage aktualisiert & Formeln fortgeschrieben.`);
      copyTimelineData(); // Sync
      return `Erfolg: ${changesMade} Tage aktualisiert!`;
    } else {
      return "Warnung: Keine passenden Tage im Plan gefunden.";
    }

  } catch (e) {
    logToSheet('ERROR', `[V87.4] CRASH: ${e.message}`);
    // Stacktrace für Debugging
    console.error(e.stack);
    throw new Error(e.message); // Fehler an WebApp weitergeben
  }
}

/**
 * V92-SUPER-HYBRID: Fix für WebApp Karten & iOS Nerd-Stats.
 * Stellt KI-Load, KI-Zone, KI-Hinweis, Review-Stats und History-Meta wieder her.
 */
function getDashboardDataAsStringV76() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statusData = getDashboardDataV76();

    // --- Header alias helper (minimal, local to this hot path) ---
    const norm_ = (s) => String(s || '').toLowerCase().trim();
    const getCol_ = (headersLC, aliases) => {
      for (let i = 0; i < aliases.length; i++) {
        const idx = headersLC.indexOf(norm_(aliases[i]));
        if (idx !== -1) return idx;
      }
      return -1;
    };
    const requireCols_ = (headersRaw, headersLC, reqMap) => {
      const missing = [];
      Object.keys(reqMap).forEach(k => {
        const aliases = reqMap[k];
        if (getCol_(headersLC, aliases) === -1) missing.push(`${k} (${aliases.join('|')})`);
      });
      if (missing.length) {
        const msg = '[Contract] Timeline headers fehlen: ' + missing.join(', ');
        logToSheet('ERROR', msg);
        throw new Error(msg);
      }
    };

    // 1. Timeline Daten laden
    const timelineSheet = ss.getSheetByName('KK_TIMELINE') || ss.getSheetByName('timeline');
    if (!timelineSheet) throw new Error("Timeline Sheet fehlt.");

    // Performance-Fix: keine Full-Sheet Reads (timeline hat ~50k Zeilen)
    const tLastRow = timelineSheet.getLastRow();
    const tLastCol = timelineSheet.getLastColumn();
    if (tLastRow < 2 || tLastCol < 1) throw new Error('Timeline Sheet ist leer.');

    const tHeaders = timelineSheet.getRange(1, 1, 1, tLastCol).getValues()[0];
    const tHeadersLC = tHeaders.map(h => String(h || '').toLowerCase().trim());

    // Minimal contract check for the dashboard path
    requireCols_(tHeaders, tHeadersLC, {
      date: ['date','datum','day'],
      isToday: ['is_today'],
      loadSoll: ['coachE_ESS_day','coache_ess_day'],
      sport: ['sport_x','Sport_x'],
      zone: ['zone','Zone','coach_zone']
    });

    const idx = {
      date: getCol_(tHeadersLC, ['date','datum','day']),
      isToday: getCol_(tHeadersLC, ['is_today']),
      loadIst: getCol_(tHeadersLC, ['load_fb_day']),
      loadSoll: getCol_(tHeadersLC, ['coachE_ESS_day','coache_ess_day']),
      atl: getCol_(tHeadersLC, ['fbATL_obs','fbatl_obs']),
      ctl: getCol_(tHeadersLC, ['fbCTL_obs','fbctl_obs']),
      acwr: getCol_(tHeadersLC, ['fbACWR_obs','fbacwr_obs']),
      atlFc: getCol_(tHeadersLC, ['coachE_ATL_forecast','coache_atl_forecast']),
      ctlFc: getCol_(tHeadersLC, ['coachE_CTL_forecast','coache_ctl_forecast']),
      acwrFc: getCol_(tHeadersLC, ['coachE_ACWR_forecast','coache_acwr_forecast']),
      smartGainFc: getCol_(tHeadersLC, ['coachE_Smart_Gains','coache_smart_gains']),
      sport: getCol_(tHeadersLC, ['sport_x','Sport_x']),
      zone: getCol_(tHeadersLC, ['zone','Zone','coach_zone']),
      done: getCol_(tHeadersLC, ['activity_done']),
      fix: getCol_(tHeadersLC, ['fix']),
      mono7: getCol_(tHeadersLC, ['Monotony7','monotony7']),
      strain7: getCol_(tHeadersLC, ['Strain7','strain7']),
      teAe: getCol_(tHeadersLC, ['Target_Aerobic_TE','target_aerobic_te']),
      teAn: getCol_(tHeadersLC, ['Target_Anaerobic_TE','target_anaerobic_te']),
      weekPhase: getCol_(tHeadersLC, ['Week_Phase','week_phase'])
    };

    // Today-Row effizient finden (nur is_today Spalte laden)
    let todayRow = -1; // 1-based Sheet row
    if (idx.isToday !== -1) {
      const isTodayCol = timelineSheet.getRange(2, idx.isToday + 1, tLastRow - 1, 1).getValues();
      for (let i = 0; i < isTodayCol.length; i++) {
        if (parseFloat(String(isTodayCol[i][0]).replace(',', '.')) == 1) { todayRow = i + 2; break; }
      }
    }

    // Fallback: wenn is_today nicht gefunden wird, nimm letzte Datenzeile
    if (todayRow < 2) todayRow = tLastRow;

    // Slice laden: Lookback + Forecast + Puffer
    const LOOKBACK_DAYS = 60;
    const FORWARD_DAYS = 21;
    const startRow = Math.max(2, todayRow - LOOKBACK_DAYS);
    const endRow = Math.min(tLastRow, todayRow + FORWARD_DAYS);
    const sliceRows = Math.max(1, endRow - startRow + 1);
    const slice = timelineSheet.getRange(startRow, 1, sliceRows, tLastCol).getValues();

    // Rebuild tData im alten Format: [headers, ...rows]
    const tData = [tHeaders].concat(slice);
    const todayRowIndex = (todayRow - startRow) + 1; // Index in tData (0=headers)


    // 3. Activity Reviews (Fix für ATL/CTL/ACWR + Java-Object-Fix)
    let activityReviews = [];
    const revSheet = ss.getSheetByName("AI_ACTIVITY_REVIEWS");
    if (revSheet && revSheet.getLastRow() > 1) {
        const revData = revSheet.getDataRange().getValues();
        for(let i = revData.length - 1; i >= 1 && activityReviews.length < 7; i--) {
            const rRow = revData[i];
            if (rRow[0] instanceof Date) {
                const pF = (v) => { 
                  let n = parseFloat(String(v).replace(',','.'));
                  return isNaN(n) ? 0 : n;
                };

                // Java-Object-Fix: Wir erzwingen, dass der Review-Text ein String ist
                let reviewText = rRow[9];
                if (Array.isArray(reviewText)) reviewText = reviewText[0]; // Falls es ein Array ist, nimm das erste Element
                reviewText = String(reviewText || "Keine Details.");

                activityReviews.push({
                    datum: Utilities.formatDate(rRow[0], Session.getScriptTimeZone(), "dd.MM."),
                    loadIst: Math.round(pF(rRow[1])), // Spalte B
                    loadSoll: Math.round(pF(rRow[2])), // Spalte C
                    atl: Math.round(pF(rRow[3])),      // Spalte D (ATL IST)
                    ctl: Math.round(pF(rRow[5])),      // Spalte F (CTL IST)
                    acwr: pF(rRow[7]).toFixed(2),      // Spalte H (ACWR IST)
                    text: reviewText                   // Spalte J
                });
            }
        }
    }

    // 3. KI-Plan Map (Für KI-Load, KI-Zone, KI-Hinweis)
    const aiPlanMap = new Map(); 
    try {
      const planReportSheet = ss.getSheetByName('AI_REPORT_PLAN');
      if (planReportSheet && planReportSheet.getLastRow() > 1) {
        const pData = planReportSheet.getDataRange().getValues();
        const pHead = pData[0];
        const dIdx = pHead.indexOf('Datum');
        const lIdx = pHead.indexOf('Empfohlener Load (ESS)');
        const zIdx = pHead.indexOf('Empfohlene Zone (KI)');
        const iIdx = pHead.indexOf('Prognostizierte Auswirkung (KI)');

        for (let i = 1; i < pData.length; i++) {
          if (pData[i][dIdx] instanceof Date) {
            const key = Utilities.formatDate(pData[i][dIdx], Session.getScriptTimeZone(), "yyyy-MM-dd");
            aiPlanMap.set(key, {
              load: pData[i][lIdx],
              zone: pData[i][zIdx] || "-",
              impact: pData[i][iIdx] || ""
            });
          }
        }
      }
    } catch(e) {}

    // 4. Heute finden & Nerd-Stats (Commander-Update V122)
    // Performance-Fix: todayRowIndex wurde bereits beim Slice-Read ermittelt.

    // --- ACWR FIREWALL ---
    // Wir holen den ACWR-Wert direkt aus den bereits geladenen Status-Daten (AI_REPORT_STATUS)
    // Das verhindert, dass Google Apps Script fälschlicherweise ein Datum aus der Timeline liest.
    const acwrObj = statusData.scores.find(s => s.metrik.includes("ACWR"));
    let displayACWR = "0.88"; // Fallback
    if (acwrObj) {
      displayACWR = String(acwrObj.raw_wert).replace(',', '.');
      // Falls der Wert immer noch wie ein Datum aussieht oder kein Punkt-Format hat:
      if (displayACWR.includes("T") || isNaN(parseFloat(displayACWR))) {
         displayACWR = "1.10"; // Dein aktueller Zielwert als Notfall-Anker
      }
    }

    let nerdStats = { ctl: 0, atl: 0, tsb: 0, acwr: displayACWR };
    let isActivityDone = false;
    let planData = [];

    if (todayRowIndex !== -1) {
        const row = tData[todayRowIndex];
        const p = (v, fc) => {
          let n = parseFloat(String(v).replace(',','.')) || 0;
          return n === 0 ? (parseFloat(String(fc).replace(',','.')) || 0) : n;
        };

        // CTL, ATL und TSB (Form) berechnen
        let ctl = p(row[idx.ctl], row[idx.ctlFc]);
        let atl = p(row[idx.atl], row[idx.atlFc]);
        
        nerdStats.ctl = Math.round(ctl);
        nerdStats.atl = Math.round(atl);
        nerdStats.tsb = Math.round(ctl - atl);
        // acwr ist bereits oben via Firewall gesetzt!

        // Check ob Training erledigt
        isActivityDone = (row[idx.done] === "x" || row[idx.done] === true || parseFloat(String(row[idx.loadIst]).replace(',','.')) > 0);

        // 14-Tage Plan bauen (Mit KI-Feldern)
        // --- VORBEREITUNG TE BALANCE BERECHNUNG ---
        const actualIndices = { 
          load: tHeaders.indexOf('load_fb_day'), 
          aerobic: tHeaders.indexOf('Aerobic_TE'), 
          anaerobic: tHeaders.indexOf('Anaerobic_TE') 
        };
        const targetIndices = { 
          load: tHeaders.indexOf('coachE_ESS_day'), 
          aerobic: tHeaders.indexOf('Target_Aerobic_TE'), 
          anaerobic: tHeaders.indexOf('Target_Anaerobic_TE') 
        };
        const LOOKBACK_DAYS_VARIANZ = 28;

        for (let i = 0; i < 14; i++) {
            let r = todayRowIndex + i;
            if (r < tData.length) {
                let d = tData[r][idx.date];
                if (!(d instanceof Date)) d = new Date(d);
                let dStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
                
                // --- NEU: LIVE BERECHNUNG DER TE BALANCE FÜR DIESEN TAG ---
                let teRatio = 0;
                try {
                    // Hybrid-Kalkulation: Nimmt Ist-Werte für die Vergangenheit und Soll-Werte für die Zukunft
                    teRatio = calculateForecastTEVarianz(tData, r, LOOKBACK_DAYS_VARIANZ, actualIndices, targetIndices, todayRowIndex);
                } catch(e) { teRatio = 0; }

                let formattedTeBalance = (teRatio * 100).toFixed(1);

                let ki = aiPlanMap.get(dStr) || { load: "-", zone: "-", impact: "" };

                planData.push({
                    datum: dStr,
                    tag: i === 0 ? "Heute" : Utilities.formatDate(d, Session.getScriptTimeZone(), "EEE"),
                    load: tData[r][idx.loadSoll],
                    sport: tData[r][idx.sport],
                    zone: tData[r][idx.zone],
                    fix: tData[r][idx.fix],
                    te_ae: tData[r][idx.teAe],
                    te_an: tData[r][idx.teAn],
                    kiLoad: ki.load,   
                    kiZone: ki.zone,   
                    kiImpact: ki.impact,
                    mono: tData[r][idx.mono7] || 0,
                    strain: tData[r][idx.strain7] || 0,
                    // HIER IST DER FIX:
                    teBalance: formattedTeBalance,
                    weekPhase: tData[r][idx.weekPhase] || "A" // Hier muss der Wert rein! 
                });
            }
        }
    }

    // 5. History Report (WIEDERHERGESTELLT)
    let historyReport = [];
    let historyMeta = "unbekannt";
    const hSheet = ss.getSheetByName('AI_REPORT_HISTORY');
    if (hSheet) {
        const hValues = hSheet.getRange("A2:C4").getValues();
        historyReport = hValues.map(r => ({ cat: r[0], ampel: r[1], text: r[2] }));
        historyMeta = hSheet.getRange("A6").getDisplayValue() || "unbekannt";
    }

    // --- NEU: SMART GAIN UPDATE AUS TIMELINE ---
    if (todayRowIndex !== -1 && idx.smartGainFc !== -1) {
        const rawSmartGain = tData[todayRowIndex][idx.smartGainFc];
        const smartGainValue = parseFloat(String(rawSmartGain).replace(',', '.')) || 0;

        // Wir suchen im statusData (aus AI_REPORT_STATUS) nach dem Eintrag "KEI"
        let smartGainScoreObj = statusData.scores.find(s => s.metrik === "KEI") || statusData.scores.find(s => s.metrik === "Smart Gains");
        
        if (smartGainScoreObj) {
            // Wir überschreiben den Wert aus dem Report mit dem Live-Wert aus der Timeline
            smartGainScoreObj.raw_wert = smartGainValue.toFixed(2);
            
            // Optional: Hier könnten wir auch die Ampel-Farbe basierend auf deinen neuen Grenzen setzen
            // smartGainScoreObj.ampel = (smartGainValue < -15) ? "ROT" : (smartGainValue > 5) ? "GRÜN" : "BLAU";
        }
    }
    

    // --- START INTEGRATION PERFORMANCE KIRA 3.0 (REALITY MODE) ---
    // FIX: Wir optimieren NICHT neu, sondern zeigen exakt das an, was in der Timeline steht.
    const baseData = getLabBaselines(); 
    
    // Wir bereiten die Timeline-Daten für die V2-Berechnung vor
    // Wir nutzen den 'load' aus planData (das ist 'coachE_ESS_day' aus der Timeline)
    const simulationInput = planData.map(d => ({
        ...d,
        load: parseFloat(d.load) || 0 // Wir zwingen die Engine, DEINEN Plan zu nutzen
    }));
    
    // Wir berechnen nur die physiologischen Folgen (KEI V2) basierend auf DEINEM Plan
    const calculatedV2Plan = enrichPlanWithProjections(simulationInput, baseData);

    // Wir bauen das finale Objekt
    const finalPerformancePlan = planData.map((day, index) => {
        const v2Data = calculatedV2Plan[index]; // Das berechnete Ergebnis für diesen Tag
        
        return {
            ...day,
            // WICHTIG: Wir zeigen exakt die Werte aus der Timeline/PlanApp an!
            load: day.load,          // Linke Spalte: Dein geplanter Load
            kiLoad: day.load,        // Rechte Spalte: Identisch (damit keine Verwirrung entsteht)
            
            // Hier kommen die frischen V2 Metriken rein
            projectedSG: v2Data.projectedSG,
            projectedKEI: v2Data.projectedSG,
            projectedACWR: v2Data.projectedACWR
        };
    });

    // Forecast Sheet Update (Optional, nur wenn du willst, dass der Forecast auch im Sheet landet)
    // writeProjectionsToForecastSheet(calculatedV2Plan);
    // updateAIReportPlanSheet(finalPerformancePlan); 
    // ^ VORSICHT: updateAIReportPlanSheet würde AI_REPORT_PLAN überschreiben. 
    // Wenn du PlanApp nutzt, besser hier auskommentieren, damit PlanApp die Hoheit behält.
    
    // ----------------------------------------------

    // ---------------------------------------------------------
    // WICHTIG: STRATEGIE BRIEFING LADEN (MIT MARKDOWN CLEANER)
    // ---------------------------------------------------------
    let strategyBriefing = "Keine Strategie-Daten verfügbar.";
    try {
      const futureSheet = ss.getSheetByName('AI_FUTURE_STATUS');
      if (futureSheet) {
        // Wir lesen die ersten 5 Zeilen der Spalte A
        const fData = futureSheet.getRange(1, 1, 5, 1).getValues();
        
        for (let r = 0; r < fData.length; r++) {
          let rawText = String(fData[r][0]).trim();
          if (!rawText) continue;

          // >>> DER CRUCIAL FIX: MARKDOWN ENTFERNEN <<<
          if (rawText.includes("```")) {
            rawText = rawText.replace(/```json/gi, "").replace(/```/g, "").trim();
          }

          if (rawText.startsWith('{')) {
            try {
              const parsed = JSON.parse(rawText);
              if (parsed.briefing) {
                strategyBriefing = parsed.briefing; // Text gefunden!
                break; 
              }
            } catch (jsonErr) {
              console.warn("JSON Parse Error: " + jsonErr.message);
            }
          }
        }
      }
    } catch(e) {
      strategyBriefing = "Fehler: " + e.message;
    }

    // 6. PAYLOAD (Hier nutzen wir jetzt 'performancePlan' statt 'planData')
    const payload = {
      // Scriptable Widget & Global
      planStatus: statusData.scores.find(s => s.metrik === "Plan Status")?.text_info || "OK",
      gesamtScoreNum: statusData.scores.find(s => s.metrik === "Gesamtscore")?.num_score || 0,
      gesamtScoreAmpel: statusData.scores.find(s => s.metrik === "Gesamtscore")?.ampel || "GRAU",
      trainingScore: statusData.trainingScore,
      trainingAmpel: statusData.trainingAmpel,
      recoveryScore: statusData.recoveryScore,
      recoveryAmpel: statusData.recoveryAmpel,
      readinessScore: statusData.scores.find(s => s.metrik === "Training Readiness")?.num_score || 0,
      readinessAmpel: statusData.scores.find(s => s.metrik === "Training Readiness")?.ampel || "BLAU",
      ernaehrungWert: statusData.scores.find(s => s.metrik.includes("Bilanz"))?.num_score || 0,
      nerdStats: nerdStats,

      // WebApp Karten
      empfehlungText: statusData.gesamtText || "Keine Analyse verfügbar.",
      planData: finalPerformancePlan, // <--- WICHTIG: Hier performancePlan einsetzen!
      isActivityDone: isActivityDone,
      scores: statusData.scores,
      strategy_text: strategyBriefing,
      historyReport: historyReport,
      historyMeta: historyMeta, 
      activityReviews: activityReviews,
      sparklines: statusData.sparklines
    };

    return JSON.stringify(payload);

  } catch (e) {
    // Das ist das fehlende Catch-Glied in der Kette
    return JSON.stringify({ error: true, empfehlungText: "Fehler im Dashboard-Sync: " + e.message });
  }
}


/**
 * (V87-MOD): Füllt auf 14 Tage auf und reicht TE-Daten an die WebApp weiter.
 */
function getPlanDataForUIV76(datenPaket, isActivityDone, kiLoadMap) {
  const planData = [];
  const heute = datenPaket.heute;
  const timeZone = Session.getScriptTimeZone();
  const PLAN_TAGE = 14; 

  // ... in getPlanDataForUIV76 ...

  // --- 1. Tag (Heute) ---
  let heuteDatumObj = (heute['date'] instanceof Date) ? heute['date'] : new Date(heute['date']);
  if (isNaN(heuteDatumObj.getTime())) { heuteDatumObj = new Date(); }
  const heuteDatumString = Utilities.formatDate(heuteDatumObj, timeZone, "yyyy-MM-dd");

  let lastMetrics = {
    atl: heute['coachE_ATL_forecast'], ctl: heute['coachE_CTL_forecast'],
    acwr: heute['coachE_ACWR_forecast'], mono: heute['Monotony7'], strain: heute['strain7']
  };

  planData.push({
    datum: heuteDatumString,
    tag: "Heute",
    load: isActivityDone ? (parseGermanFloat(heute['load_fb_day']) || 0) : (parseGermanFloat(heute['coachE_ESS_day']) || 0),
    
    // --- KORREKTUR: Zone anzeigen statt "Erledigt" ---
    // Wir zeigen immer den Wert aus der Spalte 'Zone' (Plan oder Ist), auch wenn der Tag vorbei ist.
    zone: heute['Zone'] || 'N/A', 
    // -----------------------------------------------
    
    sport: heute['Sport_x'] || 'N/A',
    fix: heute['fix'] || '',
    kiLoad: kiLoadMap.get(heuteDatumString) || null,
    te_ae: parseGermanFloat(heute['Target_Aerobic_TE']) || 0,
    te_an: parseGermanFloat(heute['Target_Anaerobic_TE']) || 0
  });

  // --- 2. Zukunft ---
  const futurePlan = datenPaket.zukunft.plan;

  for (let i = 0; i < PLAN_TAGE; i++) {
    let datumString, tagLabel, load, zone, sport, fix, kiLoad, teAe, teAn;

    if (i < futurePlan.length) {
      // A) Tag existiert
      const tag = futurePlan[i];
      datumString = tag.datum;
      if (i === 0) tagLabel = "Morgen";
      else if (i === 1) tagLabel = "Übermorgen";
      else tagLabel = tag.tag;

      load = tag.original_load_ess;
      zone = tag.original_zone;
      sport = tag.original_sport;
      fix = tag.original_fix_marker;
      kiLoad = kiLoadMap.get(tag.datum) || null;
      teAe = tag.original_te_ae || 0; // <-- V87
      teAn = tag.original_te_an || 0; // <-- V87

      lastMetrics = { atl: tag.original_atl, ctl: tag.original_ctl, acwr: tag.original_acwr, mono: tag.original_monotony, strain: tag.original_strain };

    } else {
      // B) Tag existiert NICHT (Filler)
      const nextDate = new Date(heuteDatumObj);
      nextDate.setDate(heuteDatumObj.getDate() + (i + 1));
      datumString = Utilities.formatDate(nextDate, timeZone, "yyyy-MM-dd");
      tagLabel = Utilities.formatDate(nextDate, timeZone, "E"); 

      load = 0; zone = "N/A"; fix = ""; kiLoad = null; teAe = 0; teAn = 0;
      sport = "Noch nicht geplant";
      if (i === futurePlan.length) {
         sport += ` (End-Status: ATL ${lastMetrics.atl}, ACWR ${lastMetrics.acwr})`;
      }
    }

    planData.push({
      datum: datumString, tag: tagLabel, load: load,
      zone: zone, sport: sport, fix: fix, kiLoad: kiLoad,
      te_ae: teAe, te_an: teAn,
      mono: lastMetrics.mono,   // NEU: Wert mitgeben
      strain: lastMetrics.strain // NEU: Wert mitgeben
    });
  }
  return planData;
}

/**
 * V6.4-GAP-FIX: Erstellt Vorschläge für 3 Tage.
 * Sucht DYNAMISCH nach Lücken.
 * FEATURES: Phasen-Check (A/E) + Kontext-Check (Was liegt VOR der Lücke?).
 */
function getNext3DaySuggestions() {
  // KI-frei (2026-10-26): 3-Tage-Vorschläge (Gemini) entfernt – Wochenplan im Cockpit nutzen.
  return { success: false, error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' };
}

// KI-frei (2026-10-26): buildNext3DayPrompt entfernt (Prompt-Bau liegt nicht mehr im Apps Script).


/**
 * V3.0 (TE-Target Update): Schreibt den genehmigten Vorschlag in die Timeline.
 * - Trennt Sport/Zone falls nötig ("Firewall").
 * - Schreibt Target_Aerobic_TE und Target_Anaerobic_TE.
 */
function approveDaySuggestion(suggestion) {
  logToSheet('INFO', `[DEBUG CHECK] Empfangene Daten: ${JSON.stringify(suggestion)}`);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(SOURCE_TIMELINE_SHEET);
    if (!targetSheet) throw new Error(`Quell-Blatt '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);

    const data = targetSheet.getDataRange().getValues();
    const headers = data[0];

    // --- SPALTEN MAPPING ---
    const findHeaderIndex = (names) => {
      for (const name of names) {
        const idx = headers.indexOf(name);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const indices = {
      date: headers.indexOf('date'),
      weekday: headers.indexOf('Weekday'),
      load: headers.indexOf('coachE_ESS_day'),
      sport: headers.indexOf('Sport_x'),
      zone: headers.indexOf('Zone'),
      // NEUE TE ZIEL SPALTEN:
      teAerobic: headers.indexOf('Target_Aerobic_TE'),
      teAnaerobic: headers.indexOf('Target_Anaerobic_TE'),
      // FORMEL SPALTEN:
      atlForecast: headers.indexOf('coachE_ATL_forecast'),
      ctlForecast: headers.indexOf('coachE_CTL_forecast'),
      acwrForecast: headers.indexOf('coachE_ACWR_forecast'),
      monotony: headers.indexOf('Monotony7'),
      strain: headers.indexOf('Strain7'),
      atlMorning: headers.indexOf('coachE_ATL_morning'),
      ctlMorning: headers.indexOf('coachE_CTL_morning'),
      fkei: findHeaderIndex(['fKEI', 'fkei']),
      loadStar: findHeaderIndex(['Load*', 'load_star', 'load*']),
      ctlStar: findHeaderIndex(['CTL*', 'ctl_star', 'ctl*']),
      ctlGap: findHeaderIndex(['ctl_gap', 'CTL-CTL*', 'CTL−CTL*'])
    };

    if (indices.date === -1 || indices.sport === -1 || indices.zone === -1) {
      throw new Error(`Kritische Spalten (Date, Sport_x, Zone) fehlen in 'timeline'.`);
    }

    // --- CLEANING & SPLITTING LOGIC (Die Firewall) ---
    let cleanSport = suggestion.sport;
    let cleanZone = suggestion.zone;

    // Falls Sport so aussieht: "Run (Z2)" -> extrahieren
    if (cleanSport) {
        const zoneMatch = cleanSport.match(/\((.*?)\)/); 
        if (zoneMatch) {
          const content = zoneMatch[1]; 
          // Wenn Inhalt nach Zone aussieht (enthält "Z" oder Zahl)
          if (content.includes("Z") || content.match(/\d/)) {
            if (!cleanZone || cleanZone === "N/A" || cleanZone === "") {
                cleanZone = content; // Zone übernehmen falls leer
            }
            cleanSport = cleanSport.replace(/\s*\(.*?\)/g, "").trim(); // Klammer löschen
            logToSheet('INFO', `[Cleaning] Sport bereinigt: "${suggestion.sport}" -> "${cleanSport}" | "${cleanZone}"`);
          }
        }
    }
    // ------------------------------------------------

    // Zeile finden (erstes leeres Datum oder anhängen)
    let targetRowIndex = -1;
    let sourceRowIndex = -1; // Für Formeln (Zeile darüber)
    
    // Suche erste leere Zeile im Datumsfeld
    for (let i = 1; i < data.length; i++) {
      if (data[i][indices.date] === "" || data[i][indices.date] === null) {
        targetRowIndex = i;
        sourceRowIndex = i - 1;
        break;
      }
    }
    // Falls voll, neue Zeile anhängen
    if (targetRowIndex === -1) {
       targetRowIndex = data.length;
       sourceRowIndex = data.length - 1;
    }
    
    const targetSheetRow = targetRowIndex + 1;
    const sourceSheetRow = sourceRowIndex + 1;

    // Datum parsen
    const parts = suggestion.date.split('-').map(Number);
    const dateObj = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));

    // --- DATEN SCHREIBEN ---
    targetSheet.getRange(targetSheetRow, indices.date + 1).setValue(dateObj).setNumberFormat("yyyy-mm-dd");
    targetSheet.getRange(targetSheetRow, indices.weekday + 1).setValue(suggestion.weekday);
    targetSheet.getRange(targetSheetRow, indices.load + 1).setValue(suggestion.load);
    targetSheet.getRange(targetSheetRow, indices.sport + 1).setValue(cleanSport);
    targetSheet.getRange(targetSheetRow, indices.zone + 1).setValue(cleanZone);

    // NEU: TE Ziele schreiben (wenn Spalten vorhanden)
    if (indices.teAerobic !== -1 && suggestion.te_aerobic_target !== undefined) {
       targetSheet.getRange(targetSheetRow, indices.teAerobic + 1).setValue(suggestion.te_aerobic_target);
    }
    if (indices.teAnaerobic !== -1 && suggestion.te_anaerobic_target !== undefined) {
       targetSheet.getRange(targetSheetRow, indices.teAnaerobic + 1).setValue(suggestion.te_anaerobic_target);
    }

    // --- FORMELN KOPIEREN (Forecasts) ---
    const formulaCols = [
      indices.atlForecast, indices.ctlForecast, indices.acwrForecast,
      indices.monotony, indices.strain, indices.atlMorning, indices.ctlMorning,
      indices.fkei, indices.loadStar, indices.ctlStar, indices.ctlGap
    ];

    if (sourceRowIndex > 0) {
      for (const colIndex of formulaCols) {
        if (colIndex !== -1) {
          targetSheet.getRange(sourceSheetRow, colIndex + 1).copyTo(targetSheet.getRange(targetSheetRow, colIndex + 1));
        }
      }
    }

    // Sync anstoßen
    copyTimelineData(); 

    return `Genehmigt: ${suggestion.date} | ${cleanSport} (${cleanZone}) | TE: ${suggestion.te_aerobic_target}/${suggestion.te_anaerobic_target}`;

  } catch (e) {
    logToSheet('ERROR', `[approveDaySuggestion] FEHLER: ${e.message}`);
    return `FEHLER: ${e.message}`;
  }
}

/**
 * NEU (V109): Spezielle Garmin-Farbskala für Training Readiness.
 * 1-24: ROT (Schlecht)
 * 25-49: ORANGE (Niedrig)
 * 50-74: GRÜN (Mäßig)
 * 75-94: BLAU (Hoch)
 * 95-100: VIOLETT (Optimal)
 */
function getGarminReadinessColor(score) {
  if (score >= 95) return "VIOLETT";
  if (score >= 75) return "BLAU";
  if (score >= 50) return "GRÜN";
  if (score >= 25) return "ORANGE";
  return "ROT";
}

/**
 * NEU (V120): Berechnet den Score für die Proteinversorgung.
 * Ziel: Ausdauersportler im Training brauchen ca. 1.6g - 2.0g pro kg Körpergewicht.
 * Unter 1.2g ist bei hohem Load kritisch (Muskelabbau).
 */
function normalizeProteinScore(g_per_kg) {
  const OPTIMAL_MIN = 1.6; // Ab 1.6g/kg ist alles super (Score 100)
  const ACCEPTABLE_MIN = 1.2; // Ab 1.2g/kg ist es "Okay" (Score 50)
  
  if (g_per_kg >= OPTIMAL_MIN) {
    return 100;
  }
  
  if (g_per_kg < ACCEPTABLE_MIN) {
    // Unter 1.2g fällt der Score rapide ab (0 bei 0.5g)
    let score = (g_per_kg - 0.5) / (ACCEPTABLE_MIN - 0.5) * 50;
    return Math.max(0, Math.round(score));
  }
  
  // Dazwischen (1.2 bis 1.6): Score 50 bis 100
  // Steigung: 50 Punkte auf 0.4g Unterschied = 125
  let score = 50 + (g_per_kg - ACCEPTABLE_MIN) * 125;
  return Math.round(score);
}

/**
 * V132-FIXED: Berechnet den Gesamtscore.
 * Nimmt jetzt ein fertiges Array aller Scores entgegen.
 */
function calculateGesamtScore(inputScores) {
  let totalScore = 0;
  let totalWeight = 0;

  // Gewichtung (HRV, Readiness, KEI zählt doppelt)
  const weights = {
    "HRV Status": 2,          
    "Training Readiness": 2,  
    "KEI": 2,
    "Smart Gains": 2,         // Legacy Fallback
    
    // Basis (Faktor 1)
    "RHR": 1,
    "Schlafdauer": 1,
    "Schlafscore": 1,
    "ACWR (Forecast)": 1,
    "TE Balance (% Intensiv)": 1,
    "Training Status": 1,
    "Protein-Invest": 1,
    "7-Tage-Bilanz": 1
  };

  inputScores.forEach(s => {
    // Nur gültige Scores nutzen
    if (s && s.num_score !== undefined && s.ampel !== "OFFEN" && s.ampel !== "GRAU") {
      const weight = weights[s.metrik] || 1;
      totalScore += (s.num_score * weight);
      totalWeight += weight;
    }
  });

  if (totalWeight === 0) return { num: 0, ampel: "GRAU" };

  const finalScore = Math.round(totalScore / totalWeight);
  
  return { 
    num: finalScore, 
    ampel: getGesamtAmpel(finalScore) 
  };
}

/**
 * V119-FIX: Berechnet die TE-Balance (Rolling 28 Tage) für die ZUKUNFT (Hybrid).
 * KORREKTUR: Nutzt jetzt exakt dieselben Grenzwerte (TE_LIMIT...) wie die Historie.
 */
function calculateForecastTEVarianz(allRawData, currentRowIndex, lookbackDays, actualIndices, targetIndices, todayIndex) {
    let sumBaseLoad = 0;
    let sumIntenseLoad = 0;
    const startIndex = Math.max(1, currentRowIndex - (lookbackDays - 1));

    for (let i = startIndex; i <= currentRowIndex; i++) {
        if (!allRawData[i]) continue;
        const row = allRawData[i];
        let load, aerobic, anaerobic;

        // ENTSCHEIDUNG: Historie oder Zukunft?
        if (i <= todayIndex) {
            // IST-Daten (Vergangenheit & Heute)
            load = parseGermanFloat(row[actualIndices.load]);
            aerobic = parseGermanFloat(row[actualIndices.aerobic]);
            anaerobic = parseGermanFloat(row[actualIndices.anaerobic]);
        } else {
            // SOLL-Daten (Zukunft Targets)
            load = parseGermanFloat(row[targetIndices.load]);
            aerobic = parseGermanFloat(row[targetIndices.aerobic]);
            anaerobic = parseGermanFloat(row[targetIndices.anaerobic]);
        }

        if (!load || load === 0) continue;

        // Intensitäts-Check (Exakt synchronisiert mit Historie!)
        let isIntenseDay = false;
        
        // 1. Anaerob: Nutzt die GLOBALE Konstante (oder Fallback 2.0)
        // Stelle sicher, dass TE_LIMIT_ANAEROBIC oben definiert ist!
        const limitAnaerob = (typeof TE_LIMIT_ANAEROBIC !== 'undefined') ? TE_LIMIT_ANAEROBIC : 2.0;
        
        if (typeof anaerobic === 'number' && anaerobic >= limitAnaerob) { 
             isIntenseDay = true;
        }
        // 2. Aerob: Nutzt die GLOBALE Konstante
        else if (typeof aerobic === 'number' && aerobic >= TE_LIMIT_HIGH_AEROBIC) {
             isIntenseDay = true;
        }

        if (isIntenseDay) {
            sumIntenseLoad += load;
        } else {
            sumBaseLoad += load;
        }
    }

    const totalLoad = sumBaseLoad + sumIntenseLoad;
    if (totalLoad > 0) {
        return sumIntenseLoad / totalLoad;
    } else {
        return 0.0;
    }
}

/**
 * Speichert eine Chat-Nachricht in das Blatt 'AI_CHAT_LOG'.
 * Erstellt das Blatt automatisch, falls es fehlt.
 */
function logChatMessage(sender, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAME = 'AI_CHAT_LOG';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Blatt erstellen, falls nicht vorhanden
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Sender', 'Message']); // Header
    sheet.setColumnWidth(1, 150); // Datum
    sheet.setColumnWidth(2, 100); // Sender
    sheet.setColumnWidth(3, 600); // Nachricht
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Zeitstempel
  const timestamp = new Date();
  
  // Speichern
  sheet.appendRow([timestamp, sender, message]);
}

/**
 * Liest die letzten N Nachrichten aus dem Chat-Log für den Kontext.
 * Gibt einen formatierten String zurück.
 */
function getRecentChatHistory(limit = 6) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AI_CHAT_LOG');
  
  if (!sheet || sheet.getLastRow() < 2) return "Keine vorherigen Nachrichten.";

  // Hole die letzten N Zeilen
  const lastRow = sheet.getLastRow();
  const startRow = Math.max(2, lastRow - limit + 1); // Header überspringen
  const numRows = lastRow - startRow + 1;
  
  const data = sheet.getRange(startRow, 1, numRows, 3).getValues(); // Timestamp, Sender, Message
  
  let historyText = "";
  const timeZone = Session.getScriptTimeZone();

  data.forEach(row => {
    const timeStr = Utilities.formatDate(new Date(row[0]), timeZone, "dd.MM. HH:mm");
    const sender = row[1];
    const msg = row[2];
    historyText += `[${timeStr}] ${sender}: ${msg}\n`;
  });

  return historyText;
}

/**
 * V89-FINAL: Liest Dashboard-Daten + Sparklines.
 * FIX: Speichert Text-Felder redundant (text UND text_info), um "undefined" zu verhindern.
 */
function getDashboardDataV76() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAME = 'AI_REPORT_STATUS';
  const sheet = ss.getSheetByName(SHEET_NAME);

  // Fallback Objekt, falls Sheet fehlt
  if (!sheet) {
    return { 
      trainingScore: 0, trainingAmpel: "GRAU", 
      recoveryScore: 0, recoveryAmpel: "GRAU", 
      scores: [], 
      gesamtText: "Keine Daten verfügbar.", 
      sparklines: {} 
    };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  let idxMetrik = headers.indexOf('Metrik'); 
  let idxAmpel = headers.indexOf('Status_Ampel');
  let idxRaw = headers.indexOf('Status_Wert_Num');
  let idxScore = headers.indexOf('Status_Score_100');
  let idxText = headers.indexOf('Text_Info'); 

  // Fallback Indices
  if (idxMetrik === -1) idxMetrik = 1;
  if (idxAmpel === -1) idxAmpel = 2;
  if (idxRaw === -1) idxRaw = 3;   
  if (idxScore === -1) idxScore = 4; 
  if (idxText === -1) idxText = 5;

  let trainingScore = 0;
  let trainingAmpel = "GRAU";
  let recoveryScore = 0;
  let recoveryAmpel = "GRAU";
  let gesamtText = "";
  let scores = [];

  // 1. Daten lesen
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const metrik = row[idxMetrik];
    const valScore = parseFloat(row[idxScore]) || 0; // Sicherstellen, dass es eine Zahl ist
    const valRaw = String(row[idxRaw]);
    const ampel = (idxAmpel > -1 && row[idxAmpel]) ? row[idxAmpel] : "GRAU";
    const textInfo = (idxText > -1 && row[idxText]) ? row[idxText] : ""; 

    if (metrik === 'Training Score') { trainingScore = valScore; trainingAmpel = ampel; }
    if (metrik === 'Recovery Score') { recoveryScore = valScore; recoveryAmpel = ampel; }
    if (metrik === 'Gesamtscore') gesamtText = textInfo;

    if (metrik && metrik !== "") {
        scores.push({ 
            metrik: metrik, 
            num_score: valScore, 
            raw_wert: valRaw,    
            ampel: ampel,
            text: textInfo,      // Standard Name
            text_info: textInfo  // Backup Name (für Scriptable sicherheitshalber)
        });
    }
  }

  // --- 2. SPARKLINES ---
  const sparklines = {};
  const tlSheet = ss.getSheetByName('KK_TIMELINE');

  if (tlSheet && tlSheet.getLastRow() > 1) {
      try {
          const allValues = tlSheet.getDataRange().getValues();
          const tlHead = allValues[0];
          
          const getRobustTrend = (colName, limit = 14) => {
             const idx = tlHead.indexOf(colName);
             if (idx === -1) return [];
             const validValues = [];
             // Rückwärts suchen nach gültigen Zahlen
             for (let i = allValues.length - 1; i >= 1; i--) {
                 let val = allValues[i][idx];
                 if (typeof val === 'string') val = val.replace(/"/g, '').replace(/,/g, '.').trim();
                 let num = parseFloat(val);
                 if (!isNaN(num) && num !== null && val !== "") validValues.unshift(num);
                 if (validValues.length >= limit) break;
             }
             return validValues;
          };

          // Mapping
          sparklines["RHR"] = getRobustTrend('rhr_bpm');
          sparklines["Schlafdauer"] = getRobustTrend('sleep_hours');
          sparklines["Schlafscore"] = getRobustTrend('sleep_score_0_100');
          sparklines["Training Readiness"] = getRobustTrend('Garmin_Training_Readiness');
          sparklines["HRV Status"] = getRobustTrend('hrv_status'); 
          sparklines["ACWR (Forecast)"] = getRobustTrend('coachE_ACWR_forecast'); 
          sparklines["Training Status"] = getRobustTrend('coachE_CTL_forecast'); 
          sparklines["Strain7"] = getRobustTrend('Strain7');
          sparklines["TE Balance (% Intensiv)"] = getRobustTrend('Monotony7'); 
          sparklines["7-Tage-Bilanz"] = getRobustTrend('deficit');
          sparklines["Protein-Invest"] = getRobustTrend('protein_g');
          
          const ctlArr = getRobustTrend('coachE_CTL_forecast');
          if (ctlArr.length > 0) sparklines["KEI"] = ctlArr;

          // --- LIVE FIX: Bilanz Score ---
          // Falls Bilanz nicht im Status-Report steht, bauen wir sie hier künstlich ein
          const hasBilanz = scores.find(s => s.metrik.includes("Bilanz"));
          if (!hasBilanz) {
              const defTrend = sparklines["7-Tage-Bilanz"];
              if (defTrend.length > 0) {
                  const lastDeficit = defTrend[defTrend.length - 1]; 
                  let bAmpel = "GRÜN";
                  if (lastDeficit < -500) bAmpel = "ROT";
                  else if (lastDeficit < -300) bAmpel = "GELB";
                  
                  scores.push({
                      metrik: "7-Tage-Bilanz",
                      num_score: lastDeficit < 0 ? Math.max(0, 100 + (lastDeficit/10)) : 100,
                      raw_wert: lastDeficit + " kcal",
                      ampel: bAmpel,
                      text: "Live berechnet",
                      text_info: "Live berechnet"
                  });
              }
          }

      } catch(e) {
          console.log("Data/Sparkline Error: " + e.message);
      }
  }

  return {
    trainingScore: trainingScore,
    trainingAmpel: trainingAmpel, 
    recoveryScore: recoveryScore,
    recoveryAmpel: recoveryAmpel,
    scores: scores,
    gesamtText: gesamtText,
    sparklines: sparklines 
  };
}

/**
 * Notfall-Funktion, falls die Header-Namen komplett anders sind.
 * Nutzt die Standard-Positionen (Spalte 2, 3, 4).
 */
function readDataFallback(data) {
   let tScore="N/A", rScore="N/A", scores=[];
   for(let i=1; i<data.length; i++){
     const row = data[i];
     const met = row[1]; // Col B
     const amp = row[2]; // Col C
     const val = row[3]; // Col D
     if(met==='Training Score') tScore=val;
     if(met==='Recovery Score') rScore=val;
     scores.push({metrik: met, num_score: val, ampel: amp});
   }
   return {trainingScore: tScore, recoveryScore: rScore, scores: scores};
}

/**
 * Sendet eine Nachricht via Telegram Bot API.
 * Benötigt Bot-Token (von BotFather) und Chat-ID (von userinfobot).
 */
function sendTelegram(message) {
  const cfg = getTelegramConfig_();
  if (!cfg.token || !cfg.chatId) {
    logToSheet('ERROR', '[Telegram] TELEGRAM_BOT_TOKEN oder TELEGRAM_CHAT_ID fehlt in ScriptProperties');
    return;
  }

  // Reuse the unified sender
  try {
    sendTelegramMessage(cfg.chatId, String(message || ''));
    logToSheet('INFO', '✅ Telegram-Nachricht erfolgreich gesendet.');
  } catch (e) {
    logToSheet('ERROR', `❌ Telegram Sende-Crash: ${e.message}`);
  }
}

/**
 * V1-STRUCT: Spezial-Funktion für den Report.
 * Nutzt 'responseSchema', um valides JSON für die Tabelle zu ERZWINGEN.
 * Damit sind Fließtext-Antworten technisch unmöglich.
 */
function callGeminiStructured(promptText) {
  // KI-frei (2026-10-26): Gemini-Aufrufe entfernt.
  throw new Error('callGeminiStructured: KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.');
}

/**
 * V3-OPENAI-ROBUST: Mit Fehler-Handling für HTML-Antworten.
 */
function callOpenAI(promptText) {
  // KI-frei (2026-10-26): OpenAI-Aufrufe entfernt. Der Supervisor läuft auf dem Mac (coach-kira).
  throw new Error('callOpenAI: KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.');
}

/**
 * NEU (V8.0 - Open-Meteo): Wetter ohne API-Key!
 * Stabil, schnell und stürzt nicht ab.
 * Mappt WMO-Codes auf OpenWeather-Icons für das Frontend.
 */
function getWeatherIconMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("KK_CONFIG");
  
  // Koordinaten laden (Fallback auf Berlin, falls Config fehlt)
  let lat = "52.52";
  let lon = "13.41";

  if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      configData.forEach(row => {
          if (row[0] === 'LAT') lat = String(row[1]).replace(',', '.');
          if (row[0] === 'LON') lon = String(row[1]).replace(',', '.');
      });
  }

  // Open-Meteo API URL (Kein Key nötig!)
  const url = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&daily=weather_code,temperature_2m_max,temperature_2m_min,wind_speed_10m_max&timezone=auto`;

  try {
    const response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    if (response.getResponseCode() !== 200) {
        logToSheet('WARN', '[Wetter] Open-Meteo nicht erreichbar. Fahre ohne Wetter fort.');
        return {};
    }
    
    const json = JSON.parse(response.getContentText());
    if (!json.daily) return {};

    const map = {};
    const timeZone = Session.getScriptTimeZone();

    // Wir iterieren durch die Tage
    for (let i = 0; i < json.daily.time.length; i++) {
       const dateRaw = json.daily.time[i]; // Format YYYY-MM-DD kommt direkt so an
       
       // Sicherheitshalber parsen wir das Datum, um sicher zu sein
       // OpenMeteo liefert YYYY-MM-DD String.
       const dateStr = dateRaw; 

       const code = json.daily.weather_code[i];
       const tempMax = json.daily.temperature_2m_max[i];
       const tempMin = json.daily.temperature_2m_min[i];
       
       // Mapping WMO Code -> Icon Name (für die App)
       let iconCode = "03d"; // Default Cloud
       
       // WMO Codes: https://open-meteo.com/en/docs
       if (code === 0) iconCode = "01d"; // Clear
       else if (code === 1) iconCode = "02d"; // Mainly clear
       else if (code === 2) iconCode = "03d"; // Partly cloudy
       else if (code === 3) iconCode = "04d"; // Overcast
       else if (code >= 45 && code <= 48) iconCode = "50d"; // Fog
       else if (code >= 51 && code <= 67) iconCode = "09d"; // Drizzle / Rain
       else if (code >= 71 && code <= 77) iconCode = "13d"; // Snow
       else if (code >= 80 && code <= 82) iconCode = "09d"; // Showers
       else if (code >= 95) iconCode = "11d"; // Thunderstorm
       
       map[dateStr] = {
          icon: iconCode,
          max: Math.round(tempMax),
          min: Math.round(tempMin)
       };
    }
    return map;

  } catch (e) {
    logToSheet('WARN', '[Wetter] Fehler bei Open-Meteo: ' + e.message);
    return {}; // Leeres Objekt zurückgeben, damit die App nicht abstürzt!
  }
}

/**
 * NEU: Liest die Historien-Analyse aus AI_REPORT_HISTORY.
 */
function getHistoryReportData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AI_REPORT_HISTORY');
  
  // Default-Rückgabe, falls leer
  if (!sheet) return { timestamp: "Nie", entries: [] };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { timestamp: "Nie", entries: [] };

  // Datenbereich lesen (Zeile 2 bis Ende, Spalten A bis C)
  // Struktur laut runHistoricalAnalysis: [Kategorie, Ampel, Text]
  const data = sheet.getRange(2, 1, Math.min(3, lastRow - 1), 3).getValues(); 
  
  // Zeitstempel aus Zelle A6 holen (siehe Zeile 1578 im Originalcode)
  let timestamp = "";
  try {
    timestamp = sheet.getRange("A6").getDisplayValue().replace("Stand: ", "");
  } catch(e) { timestamp = "Unbekannt"; }

  const entries = data.map(row => ({
    category: row[0],
    ampel: row[1],
    text: row[2]
  }));

  return { timestamp: timestamp, entries: entries };
}

function getFactorsForClient() {
  return {
    zoneFactors: getLoadFactors(), // [cite: 617]
    elevFactors: getElevFactors()  // [cite: 629]
  };
}

/**
 * Löscht alle existierenden Trigger für eine bestimmte Funktion.
 * Verhindert den "Too many triggers" Fehler.
 */
function deleteTriggersForFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Lädt die Chat-Historie für das Frontend.
 * Greift auf 'AI_CHAT_LOG' zu (Spalten A=Zeit, B=Sender, C=Nachricht).
 */
function getChatHistoryForUI(limit = 50) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AI_CHAT_LOG');
  
  // Falls Sheet fehlt oder leer ist (nur Header), leere Liste zurückgeben
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const lastRow = sheet.getLastRow();
  // Berechne Startzeile: Entweder Zeile 2 oder die letzten N Zeilen
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  
  // Hole Datenbereich
  const data = sheet.getRange(startRow, 1, numRows, 3).getValues();
  
  // Mapping für das Frontend
  return data.map(row => ({
    // WICHTIG: Datum in String wandeln, sonst kommt NULL im Browser an!
    timestamp: row[0] ? row[0].toString() : "", 
    sender: row[1],    
    message: row[2]    
  }));
}

/**
 * HILFSFUNKTION: Falls dein Log leer ist, führe diese Funktion EINMAL aus,
 * um Testdaten zu erzeugen und das Blatt zu reparieren.
 */
function debugChatLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('AI_CHAT_LOG');
  
  if (!sheet) {
    sheet = ss.insertSheet('AI_CHAT_LOG');
    sheet.appendRow(['Timestamp', 'Sender', 'Message']); // Header
    Logger.log("Blatt AI_CHAT_LOG erstellt.");
  }
  
  if (sheet.getLastRow() < 2) {
    const now = new Date();
    sheet.appendRow([now, 'Kira', 'Systemtest: Historie initialisiert.']);
    sheet.appendRow([now, 'Commander', 'Hallo Kira, hörst du mich?']);
    sheet.appendRow([now, 'Kira', 'Laut und deutlich, Commander.']);
    Logger.log("Test-Daten eingefügt.");
  } else {
    Logger.log("Blatt hat bereits Daten (" + sheet.getLastRow() + " Zeilen).");
  }
}

// KI-frei (Stufe 3): getDeepKnowledgeBase entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

// --- PINNWAND / NOTIZEN FUNKTIONEN ---

/**
 * Speichert eine neue Notiz im Blatt 'KK_NOTES'.
 */
function saveNote(title, content) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAME = 'KK_NOTES';
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Titel', 'Inhalt']); // Header
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 400);
    sheet.getRange("A1:C1").setFontWeight("bold");
  }

  const timestamp = new Date();
  // Neue Notiz OBEN einfügen (nach Header), damit die neuesten zuerst kommen
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 3).setValues([[timestamp, title, content]]);
  
  return "Notiz gespeichert!";
}

/**
 * Lädt alle Notizen für die WebApp (V2 - Robust).
 * Wandelt alles in Strings um, um Übertragungsfehler zu vermeiden.
 */
function getNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('KK_NOTES');
  
  // Sicherheits-Check: Gibt es das Blatt?
  if (!sheet) return [];
  
  // Nutzung von getDataRange() ist sicherer als getLastRow()
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  // Wenn weniger als 2 Zeilen (nur Header oder leer), gib leere Liste zurück
  if (values.length < 2) return [];

  // Wir schneiden den Header (Zeile 1) ab -> slice(1)
  const notes = values.slice(1).map(row => ({
    // WICHTIG: Explizite Umwandlung in String!
    date: row[0] ? String(row[0]) : "", 
    title: row[1] ? String(row[1]) : "(Kein Titel)",
    content: row[2] ? String(row[2]) : ""
  }));
  
  return notes;
}



/**
 * V146: Vollständiger Refresh inkl. RIS und 14-Tage-Forecast.
 */
function runDataRefreshOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    logToSheet('INFO', '⚡ [Fast-Calc] Starte System-Refresh (RIS + Forecast)...');

    // 1. Daten-Synchronisation
    copyTimelineData();
    const datenPaket = getSheetData();
    if (!datenPaket) return;

    // 2. Scores berechnen (für das Dashboard/Status)
    const fitnessScores = calculateFitnessMetrics(datenPaket);
    const nutritionScore = calculateNutritionScore(datenPaket);
    const subScores = calculateSubScores(fitnessScores);
    const scoreResult = calculateGesamtScore(fitnessScores.concat([nutritionScore]));
    updateStatusSheetValues(ss, fitnessScores, nutritionScore, subScores, scoreResult, datenPaket);

    // 3. RIS Update (Spalte O)
    updateHistoryWithRIS();

    // --- NEU: FORECAST ENGINE FÜR SHEET AKTIVIEREN ---
    Logger.log("📈 Generiere Performance-Forecast für AI_DATA_FORECAST...");
    
    // 1. Hole die berechneten Plan-Daten inkl. TE-Balance aus dem Dashboard-Modul
    const jsonString = getDashboardDataAsStringV76(); 
    const fullData = JSON.parse(jsonString);
    
    if (fullData.planData) {
      const baseData = getLabBaselines();
      // 2. Führe die Simulation mit den neuen Penalty-Regeln aus
      const performancePlan = enrichPlanWithProjections(fullData.planData, baseData);
      
      // 3. Schreibe ALLES (ohne .toFixed-Text-Fehler) ins Sheet
      writeProjectionsToForecastSheet(performancePlan);
      Logger.log("✅ Master-Forecast (Kira-Flow) erfolgreich exportiert.");
    }

    exportLookerChartsData();
    logToSheet('INFO', '✅ System-Refresh erfolgreich abgeschlossen.');
    
  } catch (e) {
    logToSheet('ERROR', `[Fast-Calc] Kritischer Abbruch: ${e.message}`);
    throw e;
  }
}

/**
 * Helper: Aktualisiert Werte und loggt die Entscheidungslogik.
 */
function updateStatusSheetValues(ss, fitnessScores, nutritionScore, subScores, scoreResult, datenPaket) {
  const sheet = ss.getSheetByName('AI_REPORT_STATUS');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const metricMap = new Map();
  for (let i = 1; i < data.length; i++) { 
    metricMap.set(data[i][1], i + 1); 
  }

  const writeVal = (metrik, ampel, raw, score) => {
    if (metricMap.has(metrik)) {
      const row = metricMap.get(metrik);
      sheet.getRange(row, 3).setValue(ampel); 
      
      // --- COMMANDER-FIX: ANTI-DATUM-FIREWALL ---
      const valCell = sheet.getRange(row, 4); // Spalte D
      
      if (metrik.includes("ACWR")) {
        // 1. Zwinge die Zelle in das Zahlenformat 0.00 (überschreibt Datum!)
        valCell.setNumberFormat("0.00"); 
        // 2. Sicherstellen, dass wir eine echte Zahl schreiben
        let numVal = parseFloat(String(raw).replace(',', '.'));
        valCell.setValue(numVal);
      } else {
        // Für andere Werte: Als reinen Text schreiben (verhindert Auto-Format)
        valCell.setNumberFormat("@");
        valCell.setValue(raw);
      }
      // ------------------------------------------

      sheet.getRange(row, 5).setValue(score); 
    }
  };

  // A. Werte schreiben
  writeVal("Gesamtscore", scoreResult.ampel, scoreResult.num, scoreResult.num);
  writeVal("Recovery Score", subScores.recoveryAmpel, subScores.recoveryScore, subScores.recoveryScore);
  writeVal("Training Score", subScores.trainingAmpel, subScores.trainingScore, subScores.trainingScore);

  // B. Plan Status Logik & LOGGING
  const alleScores = fitnessScores.concat([nutritionScore]);
  
  // Kriterien prüfen
  const isRecoveryCritical = alleScores.some(s => ["HRV Status","RHR","Training Readiness"].includes(s.metrik) && s.num_score < 40);
  const isRecoveryBad = alleScores.some(s => ["HRV Status","RHR","Training Readiness"].includes(s.metrik) && s.num_score < 80);
  const rtpActive = (datenPaket.rtp_status && datenPaket.rtp_status.includes("RTP"));
  
  const acwrObj = fitnessScores.find(s => s.metrik.includes("ACWR"));
  const acwrVal = acwrObj ? parseFloat(acwrObj.raw_wert.replace(',','.')) : 0;
  const acwrHigh = acwrVal > 1.3;

  // --- NEUE BEGRIFFE (Commander Edition) ---
  let planStatus = "FREIGABE"; // Ersetzt GO und OK
  let statusAmpel = "GRÜN";
  let grund = "Alles im grünen Bereich.";

  if (isRecoveryCritical) {
      planStatus = "INTERVENTION"; // Ersetzt ANPASSUNG STOP
      statusAmpel = "ROT";
      grund = "Kritische Bio-Marker (<40)";
  } else if (rtpActive) {
      planStatus = "INTERVENTION"; 
      statusAmpel = "ROT";
      grund = `RTP Protokoll aktiv (${datenPaket.rtp_status})`;
  } else if (isRecoveryBad) {
      planStatus = "BEOBACHTUNG"; // Ersetzt ANPASSUNG VORSICHT
      statusAmpel = "GELB";
      grund = "Warnende Bio-Marker (<80)";
  } else if (acwrHigh) {
      planStatus = "BEOBACHTUNG"; 
      statusAmpel = "GELB";
      grund = `ACWR zu hoch (${acwrVal})`;
  }
  
  // LOGGING DER ENTSCHEIDUNG
  logToSheet('INFO', `[Logic-Check] Status gesetzt auf: '${planStatus}'. Grund: ${grund}`);

  if (metricMap.has("Plan Status")) {
      const row = metricMap.get("Plan Status");
      sheet.getRange(row, 3).setValue(statusAmpel); 
      sheet.getRange(row, 6).setValue(planStatus); 
  }

  // C. Restliche Scores schreiben
  fitnessScores.forEach(s => {
      writeVal(s.metrik, s.ampel, s.raw_wert, s.num_score);
  });
  writeVal("7-Tage-Bilanz", nutritionScore.ampel, nutritionScore.raw_wert, nutritionScore.num_score);
}

function getLabBaselines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timelineSheet = ss.getSheetByName(TIMELINE_SHEET_NAME); 
  const configSheet = ss.getSheetByName('config');

  let results = { atl: 0, ctl: 0, ctlHistory: [], officialAcwr: 0, config: {} };

  // 1. Config laden
  if (configSheet) {
    const confData = configSheet.getDataRange().getValues();
    confData.forEach(row => {
      if (row[0]) {
        let val = row[1];
        if (typeof val === 'string') val = parseFloat(val.replace(',', '.'));
        results.config[row[0]] = isNaN(val) ? row[1] : val;
      }
    });
  }

  // 2. Timeline: Historie & Baseline extrahieren
  if (timelineSheet) {
    const data = timelineSheet.getDataRange().getValues();
    const headers = data[0];
    const idxIsToday = headers.indexOf('is_today');
    const idxAtlForecast = headers.indexOf('coachE_ATL_forecast'); 
    const idxCtlForecast = headers.indexOf('coachE_CTL_forecast');
    const idxAcwrObs = headers.indexOf('fbACWR_obs');

    for (let i = 1; i < data.length; i++) {
      if (parseFloat(String(data[i][idxIsToday]).replace(',','.')) == 1) {
        // Startwerte (Stand gestern Abend)
        results.atl = parseFloat(String(data[i-1][idxAtlForecast]).replace(',','.')) || 0;
        results.ctl = parseFloat(String(data[i-1][idxCtlForecast]).replace(',','.')) || 1;
        results.officialAcwr = parseFloat(String(data[i][idxAcwrObs]).replace(',','.')) || 0;

        // --- NEU: ECHTES 7-TAGE-GEDÄCHTNIS LADEN ---
        const startIdx = Math.max(1, i - 7);
        results.ctlHistory = data.slice(startIdx, i).map(r => 
          parseFloat(String(r[idxCtlForecast]).replace(',','.')) || results.ctl
        );
        
        // Puffer auffüllen, falls die Timeline zu kurz ist
        while (results.ctlHistory.length < 7) {
          results.ctlHistory.unshift(results.ctlHistory[0] || results.ctl);
        }
        break;
      }
    }
  }
  return results;
}

function getDashboardDataForTelegram() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('AI_REPORT_STATUS');
    if (!sheet) throw new Error("AI_REPORT_STATUS fehlt!");
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const scores = [];
    let empfehlungText = "";
    let trainingScore = 0, recoveryScore = 0;

    const idx = {
      met: headers.indexOf('Metrik'), amp: headers.indexOf('Status_Ampel'),
      sco: headers.indexOf('Status_Score_100'), raw: headers.indexOf('Status_Wert_Num'),
      txt: headers.indexOf('Text_Info')
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const metrik = row[idx.met];
      if (!metrik) continue;

      const val = {
        metrik: metrik,
        num_score: parseFloat(row[idx.sco]) || 0,
        ampel: row[idx.amp] || "GRAU",
        raw_wert: row[idx.raw] || "",
        text_info: row[idx.txt] || ""
      };

      if (metrik === 'Training Score') trainingScore = val.num_score;
      if (metrik === 'Recovery Score') recoveryScore = val.num_score;
      if (metrik === 'Gesamtscore') empfehlungText = val.text_info;

      scores.push(val);
    }

    // --- ROBUSERER ZUGRIFF AUF TIMELINE ---
    const timelineSheet = ss.getSheetByName('KK_TIMELINE') || ss.getSheetByName('timeline');
    const lastRow = timelineSheet.getLastRow();
    // 10 Zeilen Puffer, um "Heute" sicher zu finden
    const tData = timelineSheet.getRange(Math.max(1, lastRow - 10), 1, Math.min(11, lastRow), timelineSheet.getLastColumn()).getValues();
    const tHead = timelineSheet.getRange(1, 1, 1, timelineSheet.getLastColumn()).getValues()[0];
    
    const tIdx = {
      isToday: tHead.indexOf('is_today'),
      acwr: tHead.indexOf('fbACWR_obs'), acwrFc: tHead.indexOf('coachE_ACWR_forecast'),
      ctl: tHead.indexOf('fbCTL_obs'), ctlFc: tHead.indexOf('coachE_CTL_forecast'),
      atl: tHead.indexOf('fbATL_obs'), atlFc: tHead.indexOf('coachE_ATL_forecast')
    };

    let nerdStats = { acwr: "N/A", tsb: "N/A" };
    for (let i = tData.length - 1; i >= 0; i--) {
      // Flexibler Check auf 1 (Zahl oder Text)
      const isTodayVal = parseFloat(String(tData[i][tIdx.isToday]).replace(',','.'));
      if (isTodayVal === 1) {
        const r = tData[i];
        const p = (v, fc) => {
          let n = parseFloat(String(v).replace(',','.')) || 0;
          return n === 0 ? (parseFloat(String(fc).replace(',','.')) || 0) : n;
        };
        nerdStats.acwr = p(r[tIdx.acwr], r[tIdx.acwrFc]).toFixed(2);
        nerdStats.tsb = Math.round(p(r[tIdx.ctl], r[tIdx.ctlFc]) - p(r[tIdx.atl], r[tIdx.atlFc]));
        break;
      }
    }

    return { success: true, empfehlungText, scores, trainingScore, recoveryScore, nerdStats };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getTacticalLogData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName("AI_DATA_HISTORY");
    const timelineSheet = ss.getSheetByName("timeline");

    if (!historySheet || !timelineSheet) return [];

    const historyData = historySheet.getDataRange().getValues();
    const timelineData = timelineSheet.getDataRange().getValues();
    if (!historyData || historyData.length < 2) return [];
    if (!timelineData || timelineData.length < 2) return [];

    const tz = Session.getScriptTimeZone(); // wichtig: konsistente Datum-Keys

    const timelineMap = {};
    const tHeaders = timelineData[0].map(h => String(h).trim().toLowerCase());

    // Indizes finden
    const colFlags   = tHeaders.indexOf("sg_flags");
    const colDate    = tHeaders.indexOf("date");
    const colSport   = tHeaders.indexOf("sport_x");
    const colPhase   = tHeaders.indexOf("week_phase");
    const colACWR    = tHeaders.indexOf("fbacwr_obs");
    const colAerTE   = tHeaders.indexOf("aerobic_te");
    const colAnaerTE = tHeaders.indexOf("anaerobic_te");
    const colCtlFc   = tHeaders.indexOf("coache_ctl_forecast");
    const colCtlObs  = tHeaders.indexOf("fbctl_obs");

    if (colDate < 0) throw new Error("timeline: Spalte 'date' nicht gefunden.");

    const isNonEmpty = (v) => v !== "" && v !== null && v !== undefined;
    const toKeyDate = (v) => {
      if (v instanceof Date) return Utilities.formatDate(v, tz, "dd.MM.yyyy");
      return String(v || "").trim();
    };

    // ✅ timelineMap befüllen
    timelineData.slice(1).forEach(row => {
      const dStr = toKeyDate(row[colDate]);
      if (!dStr) return;

      const ctlVal =
        (colCtlFc > -1 && isNonEmpty(row[colCtlFc])) ? row[colCtlFc]
        : (colCtlObs > -1 && isNonEmpty(row[colCtlObs])) ? row[colCtlObs]
        : "";

      timelineMap[dStr] = {
        sport:   (colSport   > -1 ? row[colSport]   : "") || "",
        phase:   (colPhase   > -1 ? row[colPhase]   : "") || "",
        acwr:    (colACWR    > -1 ? row[colACWR]    : "") || "",
        aer:     (colAerTE   > -1 ? row[colAerTE]   : "") || "",
        anaer:   (colAnaerTE > -1 ? row[colAnaerTE] : "") || "",
        flags:   (colFlags   > -1 ? row[colFlags]   : "") || "",
        ctl:     ctlVal
      };
    });

    const rows = historyData.slice(1);

    const fallbackInfo = {
      sport: "Pause",
      phase: "N/A",
      acwr: "0,00",
      aer: "0,0",
      anaer: "0,0",
      flags: "",
      ctl: ""
    };

    // ✅ falls HISTORY_CHART_TAGE nicht existiert, nicht craschen
    const limit = (typeof HISTORY_CHART_TAGE !== "undefined" && HISTORY_CHART_TAGE)
      ? HISTORY_CHART_TAGE
      : 90;

    return rows.slice(-limit).reverse().map(row => {
      const dStr = toKeyDate(row[0]);
      const tInfo = timelineMap[dStr] || fallbackInfo;

      return {
        datum: dStr,
        phase: tInfo.phase || "RTP",
        load: row[2] || "0",
        teAerobic: tInfo.aer || "0,0",
        teAnaerobic: tInfo.anaer || "0,0",
        sg: row[13] || "0,0",
        ctlTrend: row[12] || "0",
        ctl: tInfo.ctl || "0",          // ✅ neu
        acwr: tInfo.acwr || "0,88",
        rhr: row[7] || "--",
        sport: tInfo.sport || "",
        v3Status: tInfo.flags || ""
      };
    });

  } catch (e) {
    console.error("Sync-Error getTacticalLogData: " + e.toString());
    return [];
  }
}


/**
 * Offizieller Einstiegspunkt für den Web-Button "Calc"
 */
function triggerDataRefresh_WebApp() {
  try {
    runDataRefreshOnly(); // Ruft deine bestehende Logik auf
    return "✅ Daten erfolgreich synchronisiert und Formate korrigiert!";
  } catch (e) {
    return "❌ Fehler beim Refresh: " + e.message;
  }
}

/**
 * Berechnet den Realized Impact Score (RIS) in der Tabelle AI_DATA_HISTORY
 * Schreibt das Ergebnis in Spalte O (Index 14).
 */
function updateHistoryWithRIS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AI_DATA_HISTORY');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Indizes finden
  const idxTR = headers.indexOf("Training Readiness");
  const idxSG = headers.indexOf("KEI Score") > -1 ? headers.indexOf("KEI Score") : headers.indexOf("Smart Gains Score");
  const idxRIS = 14; // Spalte O

  if (idxTR === -1 || idxSG === -1) {
    Logger.log("Fehler: Header für TR oder KEI nicht gefunden.");
    return;
  }

  // Header für Spalte O setzen, falls leer
  if (headers[idxRIS] !== "Realized Impact (RIS)") {
    sheet.getRange(1, idxRIS + 1).setValue("Realized Impact (RIS)");
  }

  const outputValues = [];

  for (let i = 1; i < data.length; i++) {
    const tr = parseFloat(String(data[i][idxTR]).replace(',', '.')) || 0;
    const sg = parseFloat(String(data[i][idxSG]).replace(',', '.')) || 0;
    
    // Berechnung: RIS = KEI * (TR / 100)
    // Wir runden auf 2 Dezimalstellen
    let ris = sg * (tr / 100);
    
    outputValues.push([ris.toFixed(2)]);
  }

  // In Spalte O schreiben (ab Zeile 2)
  if (outputValues.length > 0) {
    sheet.getRange(2, idxRIS + 1, outputValues.length, 1).setValues(outputValues);
  }
}

/**
 * V151: Aktive Load-Optimierung basierend auf KEI & ACWR-Schutz.
 * Sucht für jeden Tag den mathematisch effizientesten Load.
 */
function generateOptimized14DayPlan(planData, base) {
  const cfg = base.config;
  let currentATL = base.atl;
  let currentCTL = base.ctl;
  
  // CTL-Anker stabilisieren
  let startCTL = base.ctl6DaysAgo || (base.ctl * 0.95);
  let step = (base.ctl - startCTL) / 6;
  let ctlHistory = [];
  for (let i = 0; i < 7; i++) ctlHistory.push(startCTL + (i * step));

  return planData.map((day) => {
    const originalLoad = parseFloat(day.load) || 0;
    const phase = day.Week_Phase || 'A';
    const acwrLimit = (phase === 'E') ? 1.0 : 1.2;

    // NEU: Wir limitieren den maximalen Tages-Load auf das 1.5-fache der aktuellen Fitness,
    // um 200er-Ausreißer bei niedriger Fitness zu verhindern.
    const maxDailyLoadCap = Math.max(80, currentCTL * 0.4);

    let bestLoad = 0;
    let bestScore = -9999;
    
    for (let testLoad = 0; testLoad <= 180; testLoad += 5) {
      const adjATL = (cfg.S_ATL * testLoad) + (cfg.B_ATL || 0);
      const adjCTL = (cfg.S_CTL * testLoad) + (cfg.B_CTL || 0);
      const alpha = (adjATL > currentATL) ? cfg.alpha_ATL_up : cfg.alpha_ATL_down;
      
      let tATL = (1 - alpha) * currentATL + alpha * adjATL;
      const tCTL = (1 - cfg.alpha_CTL) * currentCTL + cfg.alpha_CTL * adjCTL;
      const tACWR = tATL / (tCTL || 1);
      
      // KEI V2 Logik für den Optimizer
      const fitnessGain = (tCTL - ctlHistory[0]) * 2.0; // Trend doppelt gewichtet
      const absoluteBonus = tCTL / 10.0;                // Belohnung für hohes Niveau
      const fatigueCost = (tATL * 7) / 500.0;           // Strain Proxy (ATL*7) / 500
      
      // --- DAS NEUE SCORING-SYSTEM V2 ---
      let currentScore = fitnessGain + absoluteBonus - fatigueCost;

      // 1. Progressive ACWR Penalty (startet schon ab 1.1 sanft)
      // if (tACWR > 1.1) {
      //   const excess = tACWR - 1.1;
      //   currentScore -= (excess * 200); // Sanfter Anstieg
      // }
      
      // 2. Harte Mauer bei absolutem Limit
      // if (tACWR > acwrLimit) {
      //   currentScore -= 500; 
      // }

      // 3. Strafe für extreme Tages-Last (Verhindert die 200er Sprünge)
      if (testLoad > maxDailyLoadCap) {
        currentScore -= (testLoad - maxDailyLoadCap) * 2;
      }

      if (currentScore > bestScore) {
        bestScore = currentScore;
        bestLoad = testLoad;
      }
    }

    // Finalisierung des Tages für die Kette
    const fAdjATL = (cfg.S_ATL * bestLoad) + (cfg.B_ATL || 0);
    const fAdjCTL = (cfg.S_CTL * bestLoad) + (cfg.B_CTL || 0);
    const fAlpha = (fAdjATL > currentATL) ? cfg.alpha_ATL_up : cfg.alpha_ATL_down;
    currentATL = (1 - fAlpha) * currentATL + fAlpha * fAdjATL;
    currentCTL = (1 - cfg.alpha_CTL) * currentCTL + cfg.alpha_CTL * fAdjCTL;
    ctlHistory.shift();
    ctlHistory.push(currentCTL);

    return {
      ...day,
      originalLoad: parseFloat(day.load) || 0,
      recommendedLoad: bestLoad,
      projectedSG: bestScore,
      projectedKEI: bestScore,
      projectedACWR: currentATL / (currentCTL || 1)
    };
  });
}

/**
 * Kern-Simulation der Coach-E Formel für einen Test-Load
 */
function simulateDay(load, oldATL, oldCTL, dayMeta, cfg, ctl7dAgo) {
  // Schlaf-Defaults aus Config
  const sH = cfg.sleep_hours_default || 7.5;
  const sS = cfg.sleep_score_default || 85;

  const adjLoadATL = (cfg.S_ATL * load) + (cfg.B_ATL || 0); 
  const adjLoadCTL = (cfg.S_CTL * load) + (cfg.B_CTL || 0);

  const alphaATL = (adjLoadATL > oldATL) ? cfg.alpha_ATL_up : cfg.alpha_ATL_down;
  let newATL = (1 - alphaATL) * oldATL + alphaATL * adjLoadATL;
  // Capping
  newATL = Math.max(oldATL - cfg.cap_down_ATL, Math.min(oldATL + cfg.cap_up_ATL, newATL));

  const newCTL = (1 - cfg.alpha_CTL) * oldCTL + cfg.alpha_CTL * adjLoadCTL;
  
  const acwr = newATL / (newCTL || 1);
  const trend = newCTL - ctl7dAgo;
  const smart = trend - ((newATL * 7) / 400); // Deine KEI-Formel

  return { atl: newATL, ctl: newCTL, acwr: acwr, smartGain: smart };
}

/**
 * V2-ULTRA-STABLE: Nutzt echtes 7-Tage-Gedächtnis für KEI.
 */
function enrichPlanWithProjections(planData, base) {
  const cfg = base.config || {};
  const S_ATL = cfg.S_ATL || 18.0; 
  const S_CTL = cfg.S_CTL || 12.5;
  const alphaATL = cfg.alpha_ATL_down || 0.17; 
  const alphaCTL = cfg.alpha_CTL || 0.018;
  
  let currentATL = base.atl || 0; 
  let currentCTL = base.ctl || 0;
  
  // --- FIX: Wir nehmen die echte Historie von base.ctlHistory ---
  let ctlHistory = (base.ctlHistory && base.ctlHistory.length >= 7) 
                   ? [...base.ctlHistory] 
                   : [currentCTL - 0.5]; // Nur Not-Fallback

  return planData.map((day) => {
    const load = parseFloat(day.recommendedLoad) || parseFloat(day.kiLoad) || parseFloat(day.load) || 0;
    
    // Banister Modell
    const adjLoadATL = (S_ATL * load); 
    const adjLoadCTL = (S_CTL * load);
    
    let nextATL = (1 - alphaATL) * currentATL + alphaATL * adjLoadATL;
    let nextCTL = (1 - alphaCTL) * currentCTL + alphaCTL * adjLoadCTL;

    // --- SMART GAINS V2 BERECHNUNG (UNVERÄNDERTE FORMEL) ---
    // Der Wert von vor 7 Tagen ist IMMER an Index 0 des Puffers
    const ctl7dAgo = ctlHistory[0]; 
    
    const trendComp = (nextCTL - ctl7dAgo) * 2.0;
    const statusComp = nextCTL / 10.0;
    const strainEst = nextATL * 7;
    const costComp = strainEst / 500.0; // Deine optimierte Gewichtung
    
    const smartGainV2 = trendComp + statusComp - costComp;
    
    const acwr = nextATL / (nextCTL || 1);

    // Puffer rotieren (FIFO)
    ctlHistory.push(nextCTL);
    ctlHistory.shift(); // Hält den Puffer immer bei exakt 7 Tagen Historie
    
    currentATL = nextATL;
    currentCTL = nextCTL;

    return {
      ...day,
      projectedATL: nextATL,
      projectedCTL: nextCTL,
      projectedACWR: acwr,
      projectedSG: smartGainV2,
      projectedKEI: smartGainV2
    };
  });
}

function writeProjectionsToForecastSheet(performancePlan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const forecastSheet = ss.getSheetByName('AI_DATA_FORECAST');
  const statusData = getDashboardDataV76(); 
  if (!forecastSheet) return;

  const headers = ["Datum", "Geplanter Load (ESS)", "ATL Prognose", "CTL Prognose", "ACWR Prognose", 
                   "Monotony7 Prognose", "Strain7 Prognose", "TE Balance", "Recovery Score", "Training Score", "KEI Forecast"];

  // Helfer: Macht aus jedem Input eine saubere Zahl für die Berechnung/Speicherung
  const toNum = (val) => {
    if (val === undefined || val === null || val === "") return 0;
    let clean = String(val).replace(',', '.');
    let n = parseFloat(clean);
    return isNaN(n) ? 0 : n;
  };
  
  const outputRows = performancePlan.map(day => [
    day.datum,
    toNum(day.load),
    toNum(day.projectedATL),
    toNum(day.projectedCTL),
    toNum(day.projectedACWR),
    toNum(day.mono),
    toNum(day.strain),
    toNum(day.teBalance),
    toNum(statusData.recoveryScore),
    toNum(statusData.trainingScore),
    toNum(day.projectedSG)
  ]);

  forecastSheet.clearContents(); 
  forecastSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (outputRows.length > 0) {
    const range = forecastSheet.getRange(2, 1, outputRows.length, headers.length);
    range.setValues(outputRows);

    // --- DER FORMAL-FIX ---
    // 1. Datum Spalte (A)
    forecastSheet.getRange(2, 1, outputRows.length, 1).setNumberFormat("dd.MM.yyyy");

    // 2. Alle Zahlen-Spalten (B bis K)
    // Wir setzen das Format auf Zahl mit Komma. Da dein Sheet auf DE steht, 
    // wandelt Google die internen JS-Punkte automatisch in Kommas um.
    forecastSheet.getRange(2, 2, outputRows.length, headers.length - 1).setNumberFormat("#,##0.00");
  }
  
  Logger.log("✅ AI_DATA_FORECAST wurde mit echten Zahlenwerten befüllt.");
}

/**
 * Schreibt die optimierten KI-Vorschläge zurück in das Sheet AI_REPORT_PLAN,
 * damit WebApp und Tabelle synchron sind.
 */
function updateAIReportPlanSheet(optimizedPlan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AI_REPORT_PLAN');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dIdx = headers.indexOf('Datum');
  const rIdx = headers.indexOf('Empfohlener Load (ESS)');

  optimizedPlan.forEach(planDay => {
    for (let i = 1; i < data.length; i++) {
      let sheetDate = data[i][dIdx];
      if (sheetDate instanceof Date) {
        let sheetDateStr = Utilities.formatDate(sheetDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (sheetDateStr === planDay.datum) {
          sheet.getRange(i + 1, rIdx + 1).setValue(planDay.kiLoad);
        }
      }
    }
  });
  Logger.log("✅ AI_REPORT_PLAN wurde mit KI-Vorschlägen synchronisiert.");
}

/**
 * V181-SYNC: Liest Startwerte (Gestern) & Pläne (Ab Heute) INKL. Lock-Status.
 * Basis: User-V180 + Time-Shift für korrekte Mathe.
 */
function getSimStartValues() {
  try {
    // 1. Config & Sheet laden
    const base = getLabBaselines(); // Deine existierende Config-Funktion
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Wichtig für PlanApp: zuerst SOURCE 'timeline' lesen (enthält frisch gespeicherte Locks),
    // erst danach KK_TIMELINE als Fallback.
    let sheet = ss.getSheetByName('timeline');
    if (!sheet) sheet = ss.getSheetByName('KK_TIMELINE'); // Fallback

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    
    // 2. Spalten-Mapping (Robust durch Namenssuche)
    const idx = {
  // Prio 1: CoachE Forecast (Die "reine" Lehre)
  atlFc: headers.indexOf('coache_atl_forecast'),
  ctlFc: headers.indexOf('coache_ctl_forecast'),
  // Prio 2: Garmin/Observed (Fallback)
  atlObs: headers.indexOf('fbatl_obs'),
  ctlObs: headers.indexOf('fbctl_obs'),
  ctlStar: headers.indexOf('ctl_star'),
  acwrFc: headers.indexOf('coache_acwr_forecast'),
  monotony7: headers.indexOf('monotony7'),
  sleepH: headers.indexOf('sleep_hours'),
  sleepS: headers.indexOf('sleep_score_0_100'),
  today: headers.indexOf('is_today'),

  // ✅ HIER: nur einmal
  date: headers.indexOf('date'),
  fb: headers.indexOf('load_fb_day'),
  phase: headers.indexOf('week_phase'),

  ess: headers.indexOf('coache_ess_day'),
  teAe: headers.indexOf('target_aerobic_te'),
  teAn: headers.indexOf('target_anaerobic_te'),
  sport: headers.indexOf('sport_x'),
  zone: headers.indexOf('coach_zone'),
  zoneBackup: headers.indexOf('zone'),
  fix: headers.indexOf('fix'),
  kei: headers.indexOf('coache_smart_gains')
};
if (idx.kei === -1) idx.kei = headers.indexOf('smart gain forecast');


    // Helper: Zahlen sicher parsen (1,2 -> 1.2)
    const parseVal = (v) => {
       if (!v) return 0;
       if (typeof v === 'number') return v;
       return parseFloat(String(v).replace(',', '.')) || 0;
    };

    // Logik: Erst Observation, dann CoachE Forecast, dann Baseline
    const pickObservedOrForecast = (rowData, obsIdx, fcIdx, fallback) => {
      const obsVal = parseVal(rowData[obsIdx]);
      if (obsVal > 0) return obsVal;
      const fcVal = parseVal(rowData[fcIdx]);
      if (fcVal > 0) return fcVal;
      return fallback;
    };
    // --- PlanApp Snapshot (Stabilität) ---
let snap = _readPlanAppSnapshot_();
let snapshotActive = !!snap; // wird später noch validiert


    // 3. "Heute" finden (robust: Datum > Flag)
const tz = Session.getScriptTimeZone();
const _dateKey_ = (d) => Utilities.formatDate(new Date(d), tz, "yyyy-MM-dd");
const todayDate = new Date();
todayDate.setHours(0,0,0,0);
const todayKey = _dateKey_(todayDate);

let todayRow = -1;

// Primär: exakte Datumsübereinstimmung (robust gegen versehentlich gesetzte is_today-Flags)
for (let i=1; i<data.length; i++) {
  const d = data[i][idx.date];
  if (!d) continue;
  if (_dateKey_(d) === todayKey) todayRow = i; // letzte passende Zeile gewinnt
}

// Fallback: is_today (falls Datum nicht gefunden)
if (todayRow === -1 && idx.today > -1) {
  for (let i=1; i<data.length; i++) {
    const v = data[i][idx.today];
    if (v == 1 || v === "1" || v === true || v === "TRUE") todayRow = i; // letzte gewinnt
  }
}

// Fallback: erste Zeile >= heute
if (todayRow === -1) {
  for (let i=1; i<data.length; i++) {
    const d = data[i][idx.date];
    if (!d) continue;
    if (_dateKey_(d) >= todayKey) { todayRow = i; break; }
  }
}

if(todayRow === -1) throw new Error("Heute-Zeile in Timeline nicht gefunden.");

// --- NEU: Startdatum & vorhandenen FB-Load für den "Heute"-Tag aus Timeline ableiten ---
let todayDateMs = 0;
let loadFbToday = null;

try {
  // Datum der Heute-Zeile (damit die PlanAppFB exakt dort startet)
  if (idx.date > -1) {
    const rawDate = data[todayRow][idx.date];
    const d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (!isNaN(d.getTime())) todayDateMs = d.getTime();
  }

  // FB-Load aus Timeline übernehmen, wenn befüllt
  if (idx.fb > -1) {
    const rawFb = data[todayRow][idx.fb];
    const v = parseFloat(String(rawFb).replace(',', '.'));
    if (Number.isFinite(v)) loadFbToday = v;
  }
} catch (e) {
  Logger.log("FB Startwerte Warnung: " + e.message);
}

// --- NEU: CTL Timeline (letzte 7 Tage bis GESTERN) direkt aus KK_TIMELINE ---
// Reihenfolge: ältester -> neuester (t-7 ... t-1), damit deltaCtl7d korrekt bleibt.
let timelineCtl7d = [];
for (let offset = 7; offset >= 1; offset--) {
  const r = todayRow - offset;
  if (r >= 1 && data[r]) {
    timelineCtl7d.push(pickObservedOrForecast(data[r], idx.ctlObs, idx.ctlFc, base.ctl));
  } else {
    timelineCtl7d.push(base.ctl);
  }
}

// --- DEBUG: CTL/ACWR/Monotonie HEUTE inkl. Datum (wie Sheet-Formeln) ---
const todayRowDateRaw = (idx.date > -1) ? data[todayRow][idx.date] : null;
const todayRowDateKey = todayRowDateRaw ? _dateKey_(todayRowDateRaw) : "unknown";
const ctlTodaySheet = (idx.ctlFc > -1) ? parseVal(data[todayRow][idx.ctlFc]) : 0;
const ctl7dAgoRow = todayRow - 7;
const ctl7dAgoSheet = (idx.ctlFc > -1 && ctl7dAgoRow >= 1 && data[ctl7dAgoRow])
  ? parseVal(data[ctl7dAgoRow][idx.ctlFc])
  : base.ctl;
const acwrTodaySheet = (idx.acwrFc > -1) ? parseVal(data[todayRow][idx.acwrFc]) : 0;
const monotonyTodaySheet = (idx.monotony7 > -1) ? parseVal(data[todayRow][idx.monotony7]) : 0;

Logger.log(
  `[SIM-INIT] Timeline TodayRow date=${todayRowDateKey} (row ${todayRow}) ` +
  `CTL_today(coache_ctl_forecast)=${ctlTodaySheet} ` +
  `CTL_t-7(coache_ctl_forecast)=${ctl7dAgoSheet} ` +
  `ACWR_today(coache_acwr_forecast)=${acwrTodaySheet} ` +
  `Monotony_today(monotony7)=${monotonyTodaySheet}`
);


// Snapshot von einem anderen Tag automatisch verwerfen (sonst bleibt PlanApp auf 'gestern' hängen)
if (snap && snap.startDate) {
  try {
    const snapKey = _dateKey_(new Date(snap.startDate));
    if (snapKey !== todayKey) {
      PropertiesService.getDocumentProperties().deleteProperty(PLANAPP_SNAPSHOT_KEY);
      snap = null;
    }
  } catch(e) {
    try { PropertiesService.getDocumentProperties().deleteProperty(PLANAPP_SNAPSHOT_KEY); } catch(_e){}
    snap = null;
  }
}

    
    // Fallback falls "is_today" fehlt: Datum Vergleich (TZ-sicher)
if (todayRow === -1) {
  const TZ = "Europe/Berlin";
  const todayStr = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");

  for (let i = 1; i < data.length; i++) {
    const cell = data[i][idx.date];
    if (!cell) continue;

    let rowStr = "";
    if (cell instanceof Date) {
      rowStr = Utilities.formatDate(cell, TZ, "yyyy-MM-dd");
    } else {
      // String/Number -> best effort
      const d = new Date(cell);
      rowStr = isNaN(d) ? String(cell).slice(0, 10) : Utilities.formatDate(d, TZ, "yyyy-MM-dd");
    }

    if (rowStr === todayStr) { todayRow = i; break; }
  }
}


    if(todayRow === -1) throw new Error("Heute-Zeile in Timeline nicht gefunden.");

    // --- HEUTE geschlossen? (Training schon abgeschlossen -> OBS ist Wahrheit) ---
const todayATL_obs = (idx.atlObs > -1) ? parseVal(data[todayRow][idx.atlObs]) : 0;
const todayCTL_obs = (idx.ctlObs > -1) ? parseVal(data[todayRow][idx.ctlObs]) : 0;

// "geschlossen", wenn OBS-Werte vorhanden sind
let todayIsClosed = (todayATL_obs > 0 && todayCTL_obs > 0);


    // --- NEU: Sleep Inputs für HEUTE (für Timeline-kompatible ATL/CTL Stimuli) ---
const sleepHoursDefault = parseVal(base.config.sleep_hours_default);
const sleepScoreDefault = parseVal(base.config.sleep_score_default);

const sleepHoursToday =
  (idx.sleepH > -1 && data[todayRow][idx.sleepH] !== "" && data[todayRow][idx.sleepH] != null)
    ? parseVal(data[todayRow][idx.sleepH])
    : (sleepHoursDefault || 0);

const sleepScoreToday =
  (idx.sleepS > -1 && data[todayRow][idx.sleepS] !== "" && data[todayRow][idx.sleepS] != null)
    ? parseVal(data[todayRow][idx.sleepS])
    : (sleepScoreDefault || 0);


    // --- A) STARTWERTE (DER TRICK: WIR NEHMEN GESTERN) ---
    // Damit die Simulation für "Heute" korrekt rechnet, brauchen wir den Endstand von Gestern.
    // Seed: wenn heute geschlossen -> Seed ist HEUTE (finale Werte), sonst GESTERN
let startRowIndex = todayIsClosed ? todayRow : (todayRow - 1);
if (startRowIndex < 1) startRowIndex = todayRow;


    const startRowData = data[startRowIndex];

    let finalATL = pickObservedOrForecast(startRowData, idx.atlObs, idx.atlFc, base.atl);
    let finalCTL = pickObservedOrForecast(startRowData, idx.ctlObs, idx.ctlFc, base.ctl);
    let finalCTLStar = (idx.ctlStar > -1) ? parseVal(startRowData[idx.ctlStar]) : finalCTL;
    if (!Number.isFinite(finalCTLStar) || finalCTLStar === 0) finalCTLStar = finalCTL;

    // --- B) HISTORIE (Letzte 7 Tage BIS GESTERN) ---
// Wir wollen CTL(t-7 ... t-1), wobei t = HEUTE. Seed ist GESTERN (= startRowIndex).
let ctlHistory = [];
for (let k = 6; k >= 0; k--) {
  const r = startRowIndex - k;   // endet bei startRowIndex (= gestern)
  if (r >= 1) {
    const histVal = pickObservedOrForecast(data[r], idx.ctlObs, idx.ctlFc, base.ctl);
    ctlHistory.push(histVal);
  } else {
    ctlHistory.push(base.ctl);
  }
}

// Snapshot nur für Seed/History übernehmen, wenn er explizit als "locked truth" markiert ist
// ODER wenn Observed-Spalten leer sind (sonst Priorität der Observed-Werte bewahren).
const atlObsRaw = (idx.atlObs > -1) ? startRowData[idx.atlObs] : null;
const ctlObsRaw = (idx.ctlObs > -1) ? startRowData[idx.ctlObs] : null;
const observedColumnsEmpty =
  (atlObsRaw === '' || atlObsRaw == null) &&
  (ctlObsRaw === '' || ctlObsRaw == null);
const snapshotOverridesAllowed = !!(snapshotActive && snap && observedColumnsEmpty);

// Snapshot aktiv? -> Seed/History fixieren (mehr Stabilität der KEI-Werte)
if (snapshotActive && snap) {
  if (typeof snap.todayIsClosed === 'boolean') todayIsClosed = snap.todayIsClosed;
  if (typeof snap.startRowIndex === 'number') startRowIndex = snap.startRowIndex;

  if (snapshotOverridesAllowed) {
    if (typeof snap.atlYesterday === 'number') finalATL = snap.atlYesterday;
    if (typeof snap.ctlYesterday === 'number') finalCTL = snap.ctlYesterday;

    if (Array.isArray(snap.ctlHistoryYesterday) && snap.ctlHistoryYesterday.length) {
      ctlHistory = snap.ctlHistoryYesterday.slice(-7);
      while (ctlHistory.length < 7) ctlHistory.unshift(finalCTL);
    }
  }
}

// Startdatum (Snapshot oder Live)
const startDateMs = (snapshotActive && snap && typeof snap.startDate === 'number')
  ? snap.startDate
  : new Date(data[todayRow][idx.date]).getTime();

    // --- C) ZUKUNFT (14 Tage ab Heute) ---
    let phases=[], plannedLoads=[], plannedTeAe=[], plannedTeAn=[], plannedSports=[], plannedZones=[], lockedDays=[], keiBaseline=[];
    
    for(let i=0; i<14; i++) {
      let r = todayRow + i;
      if (r < data.length) {
        phases.push(data[r][idx.phase] || "E");
        // ESS aus Timeline (coachE_ESS_day / coache_ess_day)
let essVal = parseVal(data[r][idx.ess]);

// Optional: wenn leer/0 ist, nimm FB-Load als Fallback (nur wenn du das willst)
if ((!Number.isFinite(essVal) || essVal <= 0) && idx.fb > -1) {
  const fbVal = parseVal(data[r][idx.fb]);
  if (Number.isFinite(fbVal) && fbVal > 0) essVal = fbVal;
}

plannedLoads.push(essVal);

        plannedTeAe.push(parseVal(data[r][idx.teAe]));
        plannedTeAn.push(parseVal(data[r][idx.teAn]));
        
        plannedSports.push(data[r][idx.sport] || "");
        
        // Zone robust holen
        let z = data[r][idx.zone];
        if (!z && idx.zoneBackup > -1) z = data[r][idx.zoneBackup];
        plannedZones.push(z || "");

        let keiVal = null;
        if (idx.kei > -1) {
          const rawKei = data[r][idx.kei];
          if (rawKei !== "" && rawKei != null) {
            const parsedKei = parseFloat(String(rawKei).replace(',', '.'));
            if (Number.isFinite(parsedKei)) keiVal = parsedKei;
          }
        }
        keiBaseline.push(keiVal);
        
        // LOCK STATUS LESEN
let isLocked = false;
if (idx.fix > -1) {
    const rawFix = data[r][idx.fix];
    const fixStr = String(rawFix == null ? '' : rawFix).trim().toLowerCase();
    if (rawFix == 1 || rawFix === true || fixStr === "1" || fixStr === "x" || fixStr === "true" || fixStr === "fix") {
        isLocked = true;
    }
}

// ✅ HIER EINBAUEN:
if (todayIsClosed && i === 0) isLocked = true;

lockedDays.push(isLocked);
      } else {
        // Falls Tabelle zu Ende ist, fülle auf
        plannedLoads.push(0);
        phases.push("E");
        lockedDays.push(false);
        keiBaseline.push(null);
      }
    }

    // --- AI_FUTURE_STATUS: Rohstring (JSON) 1:1 ---
let kiraBriefing = "";
const futureSheet = ss.getSheetByName('AI_FUTURE_STATUS');
if (!futureSheet) throw new Error("Sheet 'AI_FUTURE_STATUS' nicht gefunden.");
kiraBriefing = String(futureSheet.getRange(1, 1).getValue() || "");

// Snapshot aktiv? Dann Plan-Horizont auf den gespeicherten Anker-Row setzen
// (damit KEI & Forecast nicht jeden Tag "springen")
const liveTodayRowIndex = todayRow;

if (snapshotActive && snap && typeof snap === 'object') {
  const r = snap.todayRowIndex;

  if (typeof r === 'number' && r > 1) {
    todayRow = r;
    startDate = new Date(snap.startDate);
    snapshotLabel = `📌 Snapshot aktiv (${Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'd.M.yyyy')})`;
  } else {
    // Snapshot unbrauchbar -> hart deaktivieren
    snapshotActive = false;
    snap = null;
  }
} else {
  snapshotActive = false;
  snap = null;
}


    // HEUTE: Vergleichswerte (Forecast vs Observed)
const todayATL_fc = (idx.atlFc > -1) ? parseVal(data[todayRow][idx.atlFc]) : 0;
// const todayATL_obs = (idx.atlObs > -1) ? parseVal(data[todayRow][idx.atlObs]) : 0;
const todayCTL_fc = (idx.ctlFc > -1) ? parseVal(data[todayRow][idx.ctlFc]) : 0;
// const todayCTL_obs = (idx.ctlObs > -1) ? parseVal(data[todayRow][idx.ctlObs]) : 0;

    // ---------- NEU (FB / Garmin Inputs) ----------
    // Training Readiness heute: fb_TR_obs (0..100)
    // HRV Status Gate: hrv_status minus untere Schwelle aus hrv_threshholds
    // => negativ: unter Schwelle (Bremse)

    // Robust: falls idx.* nicht existiert, versuchen wir über Header zu finden.
    const _idxOf = (key) => {
      if (idx && typeof idx[key] === 'number') return idx[key];
      // Fallback: wenn du ein header-array hast: header.findIndex(...)
      // Da ich deinen Header-Variablennamen hier nicht sehe, bleibt das optional.
      return -1;
    };

    const colTR   = _idxOf('fb_TR_obs');
    const colHRV  = _idxOf('hrv_status');
    const colTH   = _idxOf('hrv_threshholds');

    let trObsToday = null;
    let hrvDeltaLowToday = null;

    // fb_TR_obs
    if (typeof todayRow === 'number' && todayRow >= 0 && colTR >= 0) {
      const v = parseVal(data[todayRow][colTR]);
      if (Number.isFinite(v)) trObsToday = v;
    }

    // hrv_status - lower(hrv_threshholds)
    // hrv_threshholds kann Zahl ODER JSON/Text sein
    if (typeof todayRow === 'number' && todayRow >= 0 && colHRV >= 0 && colTH >= 0) {
      const hrvStatus = parseVal(data[todayRow][colHRV]);

      let hrvLow = null;
      const rawTh = data[todayRow][colTH];

      // Variante A: rawTh ist direkt eine Zahl
      const thNum = parseVal(rawTh);
      if (Number.isFinite(thNum)) {
        hrvLow = thNum;
      } else {
        // Variante B: rawTh ist JSON/Text mit "lower"
        try {
          const thObj = JSON.parse(String(rawTh || ""));
          const low = parseVal(thObj && thObj.lower);
          if (Number.isFinite(low)) hrvLow = low;
        } catch (e2) {
          // ignorieren
        }
      }

      if (Number.isFinite(hrvStatus) && Number.isFinite(hrvLow)) {
        hrvDeltaLowToday = hrvStatus - hrvLow;
      }
    }
    // ---------- /NEU ----------
    // --- NEU: KEI HEUTE aus KK_TIMELINE holen ---
let keiToday_obs = 0;
try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kk = ss.getSheetByName('KK_TIMELINE');
  if (kk) {
    const kkData = kk.getDataRange().getValues();
    const kkHeaders = kkData[0].map(h => String(h).trim().toLowerCase());

    const idxSG = kkHeaders.indexOf('coache_smart_gains');
    const idxToday = kkHeaders.indexOf('is_today');

    if (idxSG > -1 && idxToday > -1) {
      // heute finden (1 in is_today)
      let kkTodayRow = -1;
      for (let r=1; r<kkData.length; r++) {
        const v = String(kkData[r][idxToday]).toUpperCase();
        if (v === "1" || v === "TRUE") { kkTodayRow = r; break; }
      }
      if (kkTodayRow > -1) {
        keiToday_obs = Number(String(kkData[kkTodayRow][idxSG]).replace(',', '.')) || 0;
      }
    }
  }
} catch(e) {
  Logger.log("KEI read warning: " + e.message);
}



    // RETURN PAYLOAD (Das Paket für die WebApp)
    return {
  atl: finalATL, 
  ctl: finalCTL,
  ctl_star: finalCTLStar,
  todayATL_fc, todayATL_obs, todayCTL_fc, todayCTL_obs,
todayRowIndex: todayRow,
startRowIndex: startRowIndex,
todayLoad: parseVal(data[todayRow][idx.ess]),
  sleepHoursToday: sleepHoursToday,
      sleepScoreToday: sleepScoreToday,
      trObsToday: trObsToday,
hrvDeltaLowToday: hrvDeltaLowToday,
      sleepHoursDefault: sleepHoursDefault || 0,
      sleepScoreDefault: sleepScoreDefault || 0,
  atlYesterday: finalATL,
  ctlYesterday: finalCTL,
  ctlHistoryYesterday: ctlHistory,
  timelineCtl7d: timelineCtl7d,
  keiToday_obs: keiToday_obs,
  smartGainsToday_obs: keiToday_obs,
  config: base.config,
  startDate: startDateMs,
  todayDateMs: todayDateMs,
loadFbToday: loadFbToday,
  phases,
  plannedLoads,
  todayIsClosed: todayIsClosed,
  snapshotActive: snapshotActive,
snapshotCreatedAt: (snap && snap.createdAt) ? snap.createdAt : 0,
snapshotLiveTodayRowIndex: (typeof liveTodayRowIndex === 'number') ? liveTodayRowIndex : -1,
todayATL_obs: todayATL_obs,
todayCTL_obs: todayCTL_obs,
  plannedTeAe,
  plannedTeAn,
  plannedSports,
  plannedZones,
  keiBaseline,
  lockedDays,
  ctlHistory, // kannst du lassen für Backward-Compatibility
  kiraBriefing
};

    
  } catch(e) { 
    Logger.log("CRITICAL ERROR SIM-INIT: " + e.message);
    throw new Error("Sim-Init Fehler: " + e.message); 
  }
}

/**
 * V185-SYNC-FB: Speichert Plan in 'timeline' UND startet Refresh + Kalender-Sync.
 * - Backward-compatible: Legacy PlanApp kann weiterhin exakt so aufrufen:
 *     saveSimulatedPlan(loads, teAeList, teAnList, sports, zones, locks)
 * - Erweiterung optional (PlanAppFB):
 *     saveSimulatedPlan(loads, teAeList, teAnList, sports, zones, locks,
 *                       focusLow, focusHigh, focusAna, planMode, intentList)
 *
 * ERSETZEN:
 * - Ersetze deine komplette bestehende saveSimulatedPlan(...) Funktion durch diese hier.
 *
 * ANPASSEN (falls deine Timeline-Spalten anders heißen):
 * - plan_mode
 * - plan_session_intent
 * - plan_focus_low_aerobic_pct
 * - plan_focus_high_aerobic_pct
 * - plan_focus_anaerobic_pct
 */
function saveSimulatedPlan(loads, teAeList, teAnList, sports, zones, locks) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1) Zwingend 'timeline'
    const sheet = ss.getSheetByName('timeline');
    if (!sheet) throw new Error("CRITICAL: Blatt 'timeline' existiert nicht!");

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) throw new Error("timeline ist leer oder hat keine Datenzeilen.");

    const headers = data[0].map(h => String(h || "").trim().toLowerCase());

    // 2) Indizes suchen (Legacy + FB optional)
    const idx = {
      today: headers.indexOf('is_today'),
      ess: headers.indexOf('coache_ess_day'),
      teAe: headers.indexOf('target_aerobic_te'),
      teAn: headers.indexOf('target_anaerobic_te'),
      sport: headers.indexOf('sport_x'),
      zone: headers.indexOf('zone'),
      coachZone: headers.indexOf('coach_zone'),
      fix: headers.indexOf('fix')
    };

    if (idx.today === -1) throw new Error("Spalte 'is_today' fehlt in timeline.");
    if (idx.ess === -1) throw new Error("Spalte 'coache_ess_day' fehlt in timeline.");

    // 3) Heute-Zeile finden (0-basiert in data)
    let todayRow0 = -1;
    for (let i = 1; i < data.length; i++) {
      const val = String(data[i][idx.today]).replace(',', '.').trim();
      if (val === "1" || val === "1.0" || val.toUpperCase() === "TRUE") {
        todayRow0 = i;
        break;
      }
    }
    if (todayRow0 === -1) throw new Error("Kein 'Heute' (1) in is_today gefunden.");

    // 4) Optional Args (FB Mode) – backward-compatible über arguments[]
    const focusLowArr  = (arguments.length > 6 && Array.isArray(arguments[6])) ? arguments[6] : null;
    const focusHighArr = (arguments.length > 7 && Array.isArray(arguments[7])) ? arguments[7] : null;
    const focusAnaArr  = (arguments.length > 8 && Array.isArray(arguments[8])) ? arguments[8] : null;
    const planModeVal  = (arguments.length > 9 && arguments[9] !== undefined && arguments[9] !== null) ? String(arguments[9]) : null;
    const intentArr    = (arguments.length > 10 && Array.isArray(arguments[10])) ? arguments[10] : null;

    // 5) Schreibe Daten batch-weise (viel schneller als setValue in Schleife)
    const n = Array.isArray(loads) ? loads.length : 0;
    if (n === 0) throw new Error("loads ist leer oder kein Array.");

    // Zielbereich (1-basiert im Sheet)
    const startRow1 = todayRow0 + 1; // todayRow0 ist 0-basiert auf data; Sheet ist 1-basiert
    const firstWriteRow1 = startRow1; // wie in deiner Version: ab morgen/ab heute+? (Legacy Verhalten: todayRow + i + 1)
    // HINWEIS: Dein Original schreibt ab (todayRow + i + 1) => also ab der Zeile NACH heute.
    // Wenn du stattdessen ab HEUTE schreiben willst, ändere auf: const firstWriteRow1 = startRow1;

    const lastRow1 = sheet.getLastRow();
    const maxWritable = Math.max(0, Math.min(n, lastRow1 - firstWriteRow1 + 1));
    if (maxWritable === 0) throw new Error("Kein Platz zum Schreiben (zu wenig Zeilen in timeline).");

    // Wir lesen den bestehenden Bereich einmal, modifizieren im Array, schreiben einmal zurück.
    // Helper: Nummern robust parse
    const toNum = (x, def) => {
      const v = parseFloat(String(x).replace(',', '.'));
      return Number.isFinite(v) ? v : def;
    };

    const loadValues = [];
    const teAeValues = [];
    const teAnValues = [];
    const sportValues = [];
    const zoneValues = [];
    const coachZoneValues = [];
    const fixValues = [];

    for (let i = 0; i < maxWritable; i++) {
      const valLoad = toNum(loads[i], 0);
      const valAe = (Array.isArray(teAeList) && teAeList[i] !== undefined) ? toNum(teAeList[i], 0) : 0;
      const valAn = (Array.isArray(teAnList) && teAnList[i] !== undefined) ? toNum(teAnList[i], 0) : 0;
      const sportVal = (idx.sport > -1 && Array.isArray(sports)) ? (sports[i] !== undefined ? sports[i] : "") : "";
      const zoneVal = Array.isArray(zones) ? (zones[i] !== undefined ? zones[i] : "") : "";
      const lockVal = (Array.isArray(locks) && locks[i]) ? 'x' : '';

      loadValues.push([valLoad]);
      teAeValues.push([valAe]);
      teAnValues.push([valAn]);
      sportValues.push([sportVal]);
      zoneValues.push([zoneVal]);
      coachZoneValues.push([zoneVal]);
      fixValues.push([lockVal]);
    }

    // Batch write only the intended columns (preserve formulas/other data)
    sheet.getRange(firstWriteRow1, idx.ess + 1, maxWritable, 1).setValues(loadValues);
    if (idx.teAe > -1) sheet.getRange(firstWriteRow1, idx.teAe + 1, maxWritable, 1).setValues(teAeValues);
    if (idx.teAn > -1) sheet.getRange(firstWriteRow1, idx.teAn + 1, maxWritable, 1).setValues(teAnValues);
    if (idx.sport > -1) sheet.getRange(firstWriteRow1, idx.sport + 1, maxWritable, 1).setValues(sportValues);
    if (idx.zone > -1) sheet.getRange(firstWriteRow1, idx.zone + 1, maxWritable, 1).setValues(zoneValues);
    if (idx.coachZone > -1) sheet.getRange(firstWriteRow1, idx.coachZone + 1, maxWritable, 1).setValues(coachZoneValues);
    if (idx.fix > -1) sheet.getRange(firstWriteRow1, idx.fix + 1, maxWritable, 1).setValues(fixValues);

    // Flush bevor Refresh/Sync
    SpreadsheetApp.flush();

    // 6) Schweren Refresh + Kalender-Sync asynchron triggern (UI nicht blockieren)
    let syncMsg = "";
    try {
      syncMsg = schedulePostSaveRefreshAndCalendarSync_();
    } catch (e) {
      syncMsg = " (Async Trigger Fehler: " + e.message + ")";
    }

    // 7) Snapshot aktualisieren
    try {
      if (typeof refreshPlanAppSnapshot === 'function') refreshPlanAppSnapshot();
    } catch (e) {
      console.log("Snapshot Warning: " + e.message);
    }

    return "✅ Plan gespeichert" + syncMsg;

  } catch (e) {
    return "❌ FEHLER: " + e.message;
  }
}


/**
 * Plant den Post-Save-Refresh + Kalender-Sync als einmaligen Trigger ein.
 * Verhindert WebApp-Timeout in saveSimulatedPlan().
 */
function schedulePostSaveRefreshAndCalendarSync_() {
  const triggerFn = 'runPostSaveRefreshAndCalendarSyncAsync_';
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction && t.getHandlerFunction() === triggerFn) {
      ScriptApp.deleteTrigger(t);
    }
  }

  ScriptApp.newTrigger(triggerFn)
    .timeBased()
    .after(10 * 1000)
    .create();

  return " (Refresh + Kalendersync asynchron gestartet ⚡🗓️)";
}

function runPostSaveRefreshAndCalendarSyncAsync_() {
  // Wichtig: zuerst Kalender-Sync, damit er nicht vom schweren Refresh "verhungert"
  try {
    if (typeof syncToGoogleCalendar === 'function') {
      syncToGoogleCalendar();
    }
  } catch (e) {
    logToSheet('WARN', '[PlanApp] Async Kalender-Sync Fehler: ' + e.message);
  }

  try {
    if (typeof runDataRefreshOnly === 'function') {
      runDataRefreshOnly();
    }
  } catch (e) {
    logToSheet('WARN', '[PlanApp] Async Refresh Fehler: ' + e.message);
  } finally {
    const triggerFn = 'runPostSaveRefreshAndCalendarSyncAsync_';
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction && t.getHandlerFunction() === triggerFn) {
        ScriptApp.deleteTrigger(t);
      }
    }
  }
}

/**
 * Plant den Kalendersync asynchron als einmaligen Trigger ein.
 * Vorteil: saveSimulatedPlan() kann sofort an die WebApp zurückkehren.
 */
function schedulePlanCalendarSync_() {
  if (typeof syncToGoogleCalendar !== 'function') {
    return " (Kein Kalender-Modul gefunden)";
  }

  const triggerFn = 'runPlanCalendarSyncAsync_';
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction && t.getHandlerFunction() === triggerFn) {
      ScriptApp.deleteTrigger(t);
    }
  }

  ScriptApp.newTrigger(triggerFn)
    .timeBased()
    .after(10 * 1000)
    .create();

  return " (Kalendersync asynchron gestartet 🗓️)";
}

/**
 * Trigger-Entry-Point für den entkoppelten Plan-Kalendersync.
 */
function runPlanCalendarSyncAsync_() {
  try {
    syncToGoogleCalendar();
  } catch (e) {
    console.warn('[PlanApp] Async Kalender-Sync Fehler: ' + e.message);
  } finally {
    const triggerFn = 'runPlanCalendarSyncAsync_';
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction && t.getHandlerFunction() === triggerFn) {
        ScriptApp.deleteTrigger(t);
      }
    }
  }
}


/**
 * V172: Strategische Projektion mit konkreten Last-Empfehlungen
 */
function getStrategicKiraBriefing(simData) {
  // KI-frei (2026-10-26): Gemini-Briefing entfernt – Plan-Briefing im Cockpit nutzen.
  return { success: false, error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' };
}

/**
 * Hilfsfunktion: Holt kompakten Kontext der letzten Tage
 */
function getBriefHistoryContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AI_DATA_HISTORY');
  if(!sheet) return "Keine Historie verfügbar.";
  const vals = sheet.getRange("A2:N8").getValues(); // Letzte 7 Zeilen
  return vals.map(r => `Tag: ${r[0]}, Load: ${r[2]}, KEI: ${r[13]}`).join(" | ");
}

function syncCalendarFromTimeline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('timeline');
  if (!sheet) return "Abbruch: Blatt 'timeline' nicht gefunden.";
  
  const data = sheet.getDataRange().getValues();
  
  // 1. Header robust einlesen (Kleinschreibung)
  const headers = data[0].map(h => h.toString().trim().toLowerCase());
  const idx = {
  date: headers.indexOf('date'),
  isToday: headers.indexOf('is_today'),
  load: headers.indexOf('coache_ess_day'), // Kleingeschrieben!
  sport: headers.indexOf('sport_x'),       // Kleingeschrieben!
  zone: headers.indexOf('zone'),
  teAe: headers.indexOf('target_aerobic_te'),
  teAn: headers.indexOf('target_anaerobic_te')
};

  // Sicherheitscheck: Wurden die Spalten gefunden?
  if (idx.isToday === -1 || idx.teAe === -1) {
    return "Fehler: Spalten in timeline nicht gefunden.";
  }

  // 2. "Heute"-Zeile finden
  let todayRow = data.findIndex(r => r[idx.isToday] == 1);
  if (todayRow === -1) return "Abbruch: is_today=1 nicht gefunden.";

  // 3. Kalender laden
  const calId = getKiraConfig('CALENDAR_ID') || "primary";
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) return "Abbruch: Kalender nicht gefunden.";
  
  // 4. Bereinigung (nächste 14 Tage)
  const start = new Date(); start.setHours(0,0,0,0);
  const end = new Date(); end.setDate(start.getDate() + 14);
  cal.getEvents(start, end).forEach(ev => {
    if (ev.getTitle().includes("[Kira]")) ev.deleteEvent();
  });

  // 5. Integration: Neue Trainingstage schreiben
  let count = 0;
  for (let i = 0; i < 14; i++) {
    const r = todayRow + i;
    if (r >= data.length) break;
    const row = data[r]; // Hier wird 'row' für diesen spezifischen Tag definiert!

    // Werte sicher einlesen
    const load = parseFloat(String(row[idx.load]).replace(',', '.')) || 0;
    const sport = row[idx.sport] || "";
    const zone = row[idx.zone] || "N/A";

    // NUR eintragen, wenn Load geplant oder Sportart benannt ist
    if (load > 0 || sport !== "") {
      
      // >>> DIESE ZEILEN MÜSSEN HIER INNEN STEHEN <<<
      const valTeAe = parseFloat(String(row[idx.teAe]).replace(',', '.')) || 0;
      const valTeAn = parseFloat(String(row[idx.teAn]).replace(',', '.')) || 0;

      const title = `[Kira] ${sport || "Training"} (${zone}) | ${load} ESS`;
      const desc = `Strategische Planung (PlanApp):\n- Load: ${load} ESS\n- Zone: ${zone}\n- TE Ziel: ${valTeAe.toFixed(1)} / ${valTeAn.toFixed(1)}`;
      
      cal.createAllDayEvent(title, new Date(row[idx.date]), {description: desc});
      count++;
    }
  }
  return "OK: " + count + " Einheiten synchronisiert.";
}


function getKiraConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('KK_CONFIG');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

/**
 * KEI-Bewertung (AI_REPORT_STATUS)
 * Legende:
 * - High Efficiency (>= 8)
 * - Productive (3 bis < 8)
 * - Maintenance / Transition (0 bis < 3)
 * - < 0 = ineffizient / rote Flagge
 */
function getKEIAssessment(value) {
  const v = Number(value);
  if (!isFinite(v)) return { score: 0, ampel: "GRAU", text: "n/a" };

  // 4-stufige KEI-Abstufung (wie Legende):
  // 1) >= 8: High Efficiency
  // 2) 3 bis < 8: Productive
  // 3) 0 bis < 3: Maintenance / Transition
  // 4) < 0: Ineffizient
  if (v >= 8) return { score: 100, ampel: "GRÜN", text: "High Efficiency" };
  if (v >= 3) return { score: 80, ampel: "GRÜN", text: "Productive" };
  if (v >= 0) return { score: 55, ampel: "GELB", text: "Maintenance / Transition" };
  return { score: 25, ampel: "ROT", text: "Ineffizient" };
}

function getSmartGainsAssessment(value) {
  return getKEIAssessment(value);
}



function debugTimelineColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('KK_TIMELINE');
  if (!sheet) sheet = ss.getSheetByName('timeline');
  
  if (!sheet) {
    Logger.log("❌ FEHLER: Kein Blatt 'KK_TIMELINE' oder 'timeline' gefunden!");
    return;
  }
  
  Logger.log("✅ Blatt gefunden: " + sheet.getName());
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log("--- SPALTEN-CHECK ---");
  
  // Wir prüfen genau diese kritischen Spalten
  const checkList = ['is_today', 'coache_ess_day', 'fix'];
  
  checkList.forEach(such => {
    // Suche exakt und fuzzy
    const exactIdx = headers.indexOf(such);
    const fuzzyIdx = headers.findIndex(h => h.toString().toLowerCase().trim() === such.toLowerCase());
    
    if (fuzzyIdx > -1) {
      Logger.log(`✅ Gefunden: '${such}' in Spalte ${fuzzyIdx + 1} (Header: '${headers[fuzzyIdx]}')`);
    } else {
      Logger.log(`❌ FEHLT: '${such}' konnte nicht gefunden werden!`);
    }
  });
  
  Logger.log("---------------------");
  Logger.log("Alle Header roh: " + JSON.stringify(headers));
}

// --- HILFSFUNKTION FÜR DAS PLAN-WIDGET (Timeline-Version V2 + CTL) ---
function getPlanForWidget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('KK_TIMELINE');
  if (!sheet) sheet = ss.getSheetByName('timeline'); // Fallback
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({error: "Timeline Sheet not found"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; 
  const getCol = (name) => headers.indexOf(name);
  
  // SPALTEN MAPPING
  const colDate = getCol('date');
  const colDay = getCol('Weekday');
  const colLoad = getCol('coachE_ESS_day');       
  const colSport = getCol('Sport_x');             
  const colZone = getCol('coach_Zone');          
  const colZoneBackup = getCol('Zone'); 
  const colACWR = getCol('coachE_ACWR_forecast'); 
  const colSG = getCol('coachE_Smart_Gains');     
  const colCTL = getCol('coachE_CTL_forecast');   // NEU: CTL Spalte

  // HEUTE FINDEN
  let startIndex = -1;
  const today = new Date();
  today.setHours(0,0,0,0);
  
  for (let i = 1; i < data.length; i++) {
    let rowDate = new Date(data[i][colDate]);
    rowDate.setHours(0,0,0,0);
    if (rowDate.getTime() === today.getTime()) {
      startIndex = i;
      break;
    }
  }
  
  // Fallback: Nächstes Datum suchen
  if (startIndex === -1) {
    for (let i = 1; i < data.length; i++) {
        let d = new Date(data[i][colDate]);
        if (d >= today) { startIndex = i; break; }
    }
  }
  
  if (startIndex === -1) {
     return ContentService.createTextOutput(JSON.stringify({error: "Date not found"})).setMimeType(ContentService.MimeType.JSON);
  }

  // DATEN SAMMELN (7 Tage)
  let output = [];
  const endIndex = Math.min(data.length, startIndex + 7); 
  
  for (let i = startIndex; i < endIndex; i++) {
    let row = data[i];
    
    let dateRaw = new Date(row[colDate]);
    let dateStr = Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), "dd.MM.");
    
    let load = parseFloat(String(row[colLoad]).replace(',', '.')) || 0;
    let acwr = parseFloat(String(row[colACWR]).replace(',', '.')) || 0;
    let sg = parseFloat(String(row[colSG]).replace(',', '.')) || 0;
    let ctl = parseFloat(String(row[colCTL]).replace(',', '.')) || 0; // NEU
    
    let sport = row[colSport] || "-";
    let zone = row[colZone] || row[colZoneBackup] || "";
    
    // Farblogik (KEI)
let color = "gray";
if (sg > 189) color = "red";           // Danger
else if (sg >= 142) color = "purple";  // Prime
else if (sg >= 95) color = "green";    // Productive
else if (sg >= 39) color = "orange";   // Maintenance
else color = "red";                    // Detraining                

    output.push({
      day: row[colDay],        
      date: dateStr,           
      load: Math.round(load),  
      sport: sport,            
      zone: zone,              
      acwr: acwr.toFixed(2),   
      sg: Math.round(sg),  
      ctl: Math.round(ctl),    // NEU: CTL als ganze Zahl
      color: color             
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.JSON);
}

// ===========================================
// 🗓️ KALENDER SYNC MODUL (FEHLTE BISHER)
// ===========================================

/**
 * Synchronisiert die Timeline (ab Heute + 14 Tage) in den Google Kalender.
 * Logik: Löscht alte "Kira"-Einträge an den betroffenen Tagen und schreibt den aktuellen Plan neu.
 */
function syncToGoogleCalendar() {
  try {
    // 1. Konfiguration laden
    const config = getKiraConfig_Safe(); 
    const calId = config['CALENDAR_ID'];
    
    if (!calId) {
      console.warn("❌ SYNC ABBRUCH: Keine 'CALENDAR_ID' in KK_CONFIG gefunden.");
      return;
    }

    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) {
      console.warn("❌ SYNC ABBRUCH: Kalender mit ID nicht gefunden (Zugriffsrechte?).");
      return;
    }

    // 2. Timeline Daten holen
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Frisch gespeicherte Planwerte liegen zuerst in 'timeline'
    let sheet = ss.getSheetByName('timeline');
    if (!sheet) sheet = ss.getSheetByName('KK_TIMELINE'); 

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toLowerCase());

    // Indizes finden
    const idx = {
      today: headers.indexOf('is_today'),
      date: headers.indexOf('date'),
      sport: headers.indexOf('sport_x'),
      ess: headers.indexOf('coache_ess_day'),
      zone: headers.indexOf('coach_zone'),
      teAe: headers.indexOf('target_aerobic_te'),
      teAn: headers.indexOf('target_anaerobic_te'),
      fix: headers.indexOf('fix')
    };

    if (idx.today === -1) {
      throw new Error("Kritische Spalte 'is_today' fehlt.");
    }

    // 3. Heute finden
    let todayRow = -1;
    for (let i = 1; i < data.length; i++) {
      let val = String(data[i][idx.today]);
      if (val == "1" || val == "1.0" || val == "TRUE" || val == "true") { 
        todayRow = i; 
        break; 
      }
    }
    
    // Fallback via Datum (falls is_today formel kaputt)
    if (todayRow === -1) {
       const todayDate = new Date();
       todayDate.setHours(0,0,0,0);
       for(let i=1; i<data.length; i++) {
          let rDate = new Date(data[i][idx.date]);
          rDate.setHours(0,0,0,0);
          if(rDate.getTime() === todayDate.getTime()) { todayRow = i; break; }
       }
    }

    if (todayRow === -1) {
      console.warn("Heute nicht gefunden - Sync übersprungen.");
      return;
    }

    // 4. Sync Loop (Heute + 14 Tage)
    console.log(`🔄 Starte Kalender-Sync ab Zeile ${todayRow + 1}...`);

    for (let i = 0; i < 14; i++) {
      let r = todayRow + i;
      if (r >= data.length) break;

      let rowData = data[r];
      let dateVal = new Date(rowData[idx.date]);
      
      // Werte sicher lesen
      let sport = rowData[idx.sport];
      let ess = parseFloat(String(rowData[idx.ess]).replace(',', '.')) || 0;
      let zone = rowData[idx.zone] || "";
      
      // Lock Check
      let isLocked = false;
      if (idx.fix > -1) {
         const fixVal = rowData[idx.fix];
         const fixStr = String(fixVal == null ? '' : fixVal).trim().toLowerCase();
         if (fixVal == 1 || fixVal === true || fixStr === "1" || fixStr === "x" || fixStr === "true" || fixStr === "fix") isLocked = true;
      }

      // A) CLEANUP: Alte Kira-Events an diesem Tag löschen
      // Wir löschen nur Events mit dem Schild-Icon 🛡️ oder "Coach Kira" im Text
      const existingEvents = cal.getEventsForDay(dateVal);
      for (let ev of existingEvents) {
        if (ev.getTitle().includes("🛡️") || ev.getDescription().includes("Coach Kira")) {
          ev.deleteEvent();
        }
      }

      // B) NEUER EINTRAG (Nur wenn Training geplant)
      // Bedingung: Load > 0 ODER Sport eingetragen (und nicht 'Rest')
      if (ess > 0 || (sport && sport !== "" && sport.toLowerCase() !== "rest" && sport !== "Ruhetag")) {
        
        let title = `🛡️ ${sport} [${Math.round(ess)}]`;
        if (zone) title += ` (${zone})`;
        if (isLocked) title += " 🔒";

        let desc = `Planung von Coach Kira:\n`;
        desc += `-------------------------\n`;
        desc += `⚡ Load: ${Math.round(ess)}\n`;
        desc += `🎯 Zone: ${zone}\n`;
        if (idx.teAe > -1) desc += `❤️ TE Aerob: ${rowData[idx.teAe]}\n`;
        if (idx.teAn > -1) desc += `🔥 TE Anaerob: ${rowData[idx.teAn]}\n`;
        desc += `\nStatus: ${isLocked ? "Fixiert" : "Flexibel"}`;

        cal.createAllDayEvent(title, dateVal, {description: desc});
      }
    }
    console.log("✅ Kalender-Sync fertig.");

  } catch (e) {
    console.error("❌ SYNC FEHLER: " + e.message);
  }
}

/**
 * Hilfsfunktion: Liest KK_CONFIG robust ein
 */
function getKiraConfig_Safe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('KK_CONFIG');
  let config = {};
  if (!sheet) return config;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    let key = String(data[i][0]).trim();
    let val = data[i][1];
    if (key) config[key] = val;
  }
  return config;
}

function getPrimeRangeData(opts) {
  opts = opts || {};
  const daysBack = Number(opts.daysBack || 90);
  const acwrMin  = Number(opts.acwrMin  ?? 0.8);
  const acwrMax  = Number(opts.acwrMax  ?? 1.3);
  const sleepMin = Number(opts.sleepMin ?? 70);
  const redFlagsRegex = String(opts.redFlagsRegex || ""); // z.B. "H|R|SICK"
  const qLow = Number(opts.qLow ?? 75);
  const qHigh = Number(opts.qHigh ?? 95);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("KK_TIMELINE");
  if (!sh) throw new Error("KK_TIMELINE nicht gefunden.");

  const data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return { meta:{ totalDays:0, greenDays:0 }, rows:[], debug:{ msg:"Keine Daten" } };

  const headersRaw = data[0].map(h => String(h || "").trim());
  const headers = headersRaw.map(h => h.toLowerCase());

  const findCol = (cands) => {
    for (const c of cands) {
      const idx = headers.indexOf(String(c).toLowerCase());
      if (idx > -1) return idx;
    }
    return -1;
  };

  // --- EXAKTE Header-Kandidaten (bitte ggf. an deine echten Header anpassen) ---
  const idx = {
    date:  findCol(["date", "datum"]),
    sg:    findCol(["coache_smart_gains", "coache_smart_gains_v3", "coachE_smart_gains"]),
    flags: findCol(["sg_flags", "sg_flags "]),
    acwr:  findCol(["coache_acwr_forecast", "coache_acwr", "coachE_acwr_forecast"]),
    ctl:   findCol(["coache_ctl_forecast", "coachE_ctl_forecast"]),
    atl:   findCol(["coache_atl_forecast", "coachE_atl_forecast"]),
    load:  findCol(["coache_ess_day", "ess_day", "coachE_ess_day"]),
    sleep: findCol(["sleep_score_0_100", "sleep_score"]),
    rhr:   findCol(["rhr_bpm", "rhr"]),
    hrvStatus: findCol(["hrv_status"]),
    hrvThresh: findCol(["hrv_threshholds", "hrv_thresholds"])
  };

  const num = (v) => {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return isFinite(v) ? v : null;
    const s = String(v).trim().replace(/\s/g, "").replace(",", ".");
    const x = parseFloat(s);
    return isFinite(x) ? x : null;
  };

  const toISO = (d) => {
    if (d instanceof Date) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    // falls schon String "YYYY-MM-DD"
    const s = String(d || "").trim();
    return s;
  };

  const hrvDeviation = (statusVal, threshVal) => {
    // user-spec: "Abweichung von hrv_status zu unterem Wert von hrv_threshholds"
    // threshVal kann z.B. "40;70" sein → lower = 40
    const st = num(statusVal);
    if (st === null) return null;

    let lower = null;
    if (typeof threshVal === "number") {
      lower = threshVal;
    } else {
      const s = String(threshVal || "").trim();
      const m = s.match(/(-?\d+([.,]\d+)?)/); // erste Zahl als "lower"
      if (m) lower = num(m[0]);
    }
    if (lower === null) return null;
    return st - lower;
  };

  // --- Zeitraum schneiden ---
  // --- Zeitraum schneiden (FIX): Nur Historie <= heute, dann letzte N Tage ---
const rowsAll = data.slice(1);

const todayIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

// 1) erst valid dates herausfiltern + nur bis heute
let histRows = rowsAll
  .map(r => {
    const d = (idx.date > -1) ? toISO(r[idx.date]) : "";
    return { r, d };
  })
  .filter(x => x.d && x.d.length >= 10 && x.d <= todayIso); // ISO string compare klappt bei YYYY-MM-DD

// 2) sortieren (falls Sheet nicht strikt sortiert ist)
histRows.sort((a,b) => a.d > b.d ? 1 : (a.d < b.d ? -1 : 0));

// 3) letzte daysBack nehmen
histRows = histRows.slice(Math.max(0, histRows.length - daysBack));

// 4) zurück in row-array
const rows = histRows.map(x => x.r);


  const redRe = redFlagsRegex ? new RegExp(redFlagsRegex, "i") : null;

  const out = [];
  for (const r of rows) {
    const date = (idx.date > -1) ? toISO(r[idx.date]) : "";
    const sg   = (idx.sg   > -1) ? num(r[idx.sg]) : null;

    const ctl  = (idx.ctl  > -1) ? num(r[idx.ctl]) : null;
    const atl  = (idx.atl  > -1) ? num(r[idx.atl]) : null;

    // ACWR: Spalte oder fallback ATL/CTL
    let acwr = (idx.acwr > -1) ? num(r[idx.acwr]) : null;
    if (acwr === null && atl !== null && ctl !== null && ctl !== 0) {
      acwr = atl / ctl;
    }

    const load  = (idx.load  > -1) ? num(r[idx.load]) : null;
    const sleep = (idx.sleep > -1) ? num(r[idx.sleep]) : null;
    const rhr   = (idx.rhr   > -1) ? num(r[idx.rhr]) : null;

    const flags = (idx.flags > -1) ? String(r[idx.flags] || "") : "";

    const hrvDev = (idx.hrvStatus > -1 && idx.hrvThresh > -1)
      ? hrvDeviation(r[idx.hrvStatus], r[idx.hrvThresh])
      : null;

    // --- GREEN FILTER ---
    const okAcwr  = (acwr !== null && acwr >= acwrMin && acwr <= acwrMax);
    const okSleep = (sleep === null) ? true : (sleep >= sleepMin); // wenn kein Schlaf → nicht blockieren
    const okFlags = (!redRe) ? true : !redRe.test(flags);
    const okHrv   = (hrvDev === null) ? true : (hrvDev >= 0); // simple default

    const isGreen = okAcwr && okSleep && okFlags && okHrv;

    out.push({
  date,
  sg,
  acwr,
  ctl,
  atl,
  load,
  sleep,
  rhr,
  flags,
  hrvDev,
  isGreen,

  // === ALIASES (Frontend-Robustheit) ===
  // damit Frontend sowohl "sleepScore" als auch "sleep" versteht
  sleepScore: sleep,
  // damit Frontend sowohl "sgFlags" als auch "flags" versteht
  sgFlags: flags
});
  }

  const greenSg = out.filter(x => x.isGreen && x.sg !== null).map(x => x.sg).sort((a,b)=>a-b);

  const quantile = (arr, q) => {
    if (!arr.length) return null;
    const p = q / 100;
    const pos = (arr.length - 1) * p;
    const base = Math.floor(pos);
    const rest = pos - base;
    if (arr[base+1] === undefined) return arr[base];
    return arr[base] + rest * (arr[base+1] - arr[base]);
  };

  const meta = {
    totalDays: out.length,
    greenDays: out.filter(x => x.isGreen).length,
    primeBase: quantile(greenSg, qLow),
    primeTop:  quantile(greenSg, qHigh),
    sgMedian:  quantile(greenSg, 50),
    sgP90:     quantile(greenSg, 90)
  };

  const debug = {
  headersSample: headersRaw.slice(0, 50),
  idx,
  sampleRow: out[0] || null,
  todayIso,
  histCount: histRows.length,
  lastHistDate: histRows.length ? toISO(histRows[histRows.length - 1].r[idx.date]) : null
};


  return { meta, rows: out, debug };
}

function getWetterdatenCached_() {
  const props = PropertiesService.getDocumentProperties();
  const key = 'OWM_DAILY_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const hit = props.getProperty(key);
  if (hit) return hit;
  const val = getWetterdaten();
  props.setProperty(key, val);
  return val;
}

function hb_(a, b, c, d) {
  try {
    // Zwei Aufrufformen: hb_(stage, extra) und hb_(ss, runId, stage, extra) (Supervisor)
    let stage, extra, runIdArg = null;
    if (a && typeof a === 'object' && typeof a.getSheetByName === 'function') {
      runIdArg = b; stage = c; extra = d;
    } else {
      stage = a; extra = b;
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("AI_HEARTBEAT") || ss.insertSheet("AI_HEARTBEAT");

    // Header einmalig setzen
    if (sh.getLastRow() === 0) {
      sh.getRange("A1:D1").setValues([["timestamp", "stage", "runId", "extra"]]);
      sh.setFrozenRows(1);
    }

    let runId = runIdArg ? String(runIdArg) : PropertiesService.getScriptProperties().getProperty("HB_RUN_ID");
    if (!runId) {
      runId = "RUN_" + Utilities.getUuid().slice(0, 8);
      PropertiesService.getScriptProperties().setProperty("HB_RUN_ID", runId);
    }

    // In feste Zellen schreiben (übersichtlich)
    sh.getRange("A2:D2").setValues([[new Date(), String(stage || ""), runId, String(extra || "").slice(0, 500)]]);
    SpreadsheetApp.flush(); // damit du es sofort siehst
  } catch (e) {
    // Herzschlag darf nie den Job killen
    Logger.log("HB ERROR: " + (e && e.stack ? e.stack : e));
  }
}


function hbStart_(name) {
  PropertiesService.getScriptProperties().setProperty("HB_RUN_ID", "RUN_" + Utilities.getUuid().slice(0, 8));
  hb_("START " + (name || ""), "");
}

function hbEnd_(status, extra) {
  hb_("END " + (status || "OK"), extra || "");
  PropertiesService.getScriptProperties().deleteProperty("HB_RUN_ID");
}

/**
 * Universal: Liefert KK_TIMELINE als JSON (rows + meta).
 * Unterstützt Filter: days, from, to.
 *
 * Beispiel-Aufrufe:
 *   .../exec?mode=timeline
 *   .../exec?mode=timeline&days=90
 *   .../exec?mode=timeline&from=2025-01-01&to=2025-12-31
 */
function getTimelineUniversalJson(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const days = p.days ? parseInt(p.days, 10) : null;
  const from = p.from ? String(p.from) : null; // yyyy-mm-dd
  const to   = p.to   ? String(p.to)   : null; // yyyy-mm-dd

  const payload = getTimelineUniversal_({ days, from, to });
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Core Loader
 */
function getTimelineUniversal_({ days=null, from=null, to=null } = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("KK_TIMELINE");
  if (!sh) sh = ss.getSheetByName("timeline");
  if (!sh) throw new Error("Sheet KK_TIMELINE (oder timeline) nicht gefunden.");

  const tz = Session.getScriptTimeZone() || "Europe/Berlin";
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { meta: { sheet: sh.getName(), rows: 0 }, headers: [], rows: [] };
  }

  // Header normalisieren
  const rawHeaders = values[0].map(h => String(h || "").trim());
  const headers = rawHeaders.map(normalizeHeader_);

  // Date-Index finden (robust)
  const idxDate =
    headers.indexOf("date") > -1 ? headers.indexOf("date") :
    headers.indexOf("datum") > -1 ? headers.indexOf("datum") :
    headers.indexOf("day") > -1 ? headers.indexOf("day") : -1;

  // rows -> objects
  let rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const obj = {};

    for (let c = 0; c < headers.length; c++) {
      const key = headers[c] || ("col_" + c);
      let v = row[c];

      // Dates sauber rausgeben
      if (v instanceof Date) {
        v = Utilities.formatDate(v, tz, "yyyy-MM-dd");
      }

      // Numbers/Strings: wir lassen Numbers als Number
      // Strings bleiben String (Charts können später parseFloat machen)
      obj[key] = v;
    }

    // Zusätzlich: canonical dateKey
    if (idxDate > -1) {
      const dv = row[idxDate];
      if (dv instanceof Date) obj.dateKey = Utilities.formatDate(dv, tz, "yyyy-MM-dd");
      else if (dv) {
        const d2 = new Date(dv);
        obj.dateKey = isNaN(d2) ? String(dv).slice(0,10) : Utilities.formatDate(d2, tz, "yyyy-MM-dd");
      } else {
        obj.dateKey = "";
      }
    } else {
      obj.dateKey = "";
    }

    rows.push(obj);
  }

  // Filter anwenden
  rows = filterRowsByDate_(rows, { days, from, to });

  // Meta
  const meta = {
    sheet: sh.getName(),
    tz,
    rows: rows.length,
    generatedAt: new Date().toISOString()
  };

  return { meta, headers, rows };
}

function normalizeHeader_(h) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_")
    .replace(/[^\w]/g, ""); // nur [a-z0-9_]
}

/**
 * Filtert über dateKey (yyyy-mm-dd). Wenn dateKey fehlt -> keine Filterung.
 */
function filterRowsByDate_(rows, { days=null, from=null, to=null } = {}) {
  const hasDate = rows.some(r => r.dateKey && String(r.dateKey).length >= 10);
  if (!hasDate) return rows;

  // from/to -> strings yyyy-mm-dd
  let fromKey = from && from.length >= 10 ? from.slice(0,10) : null;
  let toKey   = to   && to.length   >= 10 ? to.slice(0,10)   : null;

  // days -> fromKey setzen relativ zu "heute" (über dateKey max)
  if (days && Number.isFinite(days) && days > 0) {
    // wir nehmen das max dateKey aus den Daten als "heute/letzter Stand"
    const maxKey = rows
      .map(r => r.dateKey)
      .filter(k => k && k.length >= 10)
      .sort()
      .slice(-1)[0];

    if (maxKey) {
      const dMax = new Date(maxKey + "T00:00:00");
      const dFrom = new Date(dMax);
      dFrom.setDate(dFrom.getDate() - (days - 1));

      const pad = (n)=>String(n).padStart(2,"0");
      fromKey = `${dFrom.getFullYear()}-${pad(dFrom.getMonth()+1)}-${pad(dFrom.getDate())}`;
      toKey = maxKey;
    }
  }

  return rows.filter(r => {
    const k = String(r.dateKey || "");
    if (!k) return false;
    if (fromKey && k < fromKey) return false;
    if (toKey && k > toKey) return false;
    return true;
  });
}

/**
 * Liefert KK_TIMELINE als JSON Payload (Headers + Rows).
 * Optional:
 * - days=N -> letzte N Tage bis inkl. heute (robust via date Spalte)
 * - futureDays=M -> zusätzlich M Tage ab heute (inkl.) nach vorne (für Forecast)
 *
 * Zusätzlich:
 * - requiredHeadersCheck + missingHeaders (damit neue Charts nicht "still" leer sind)
 */
function getTimelinePayload(days, futureDays) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Robust: KK_TIMELINE oder timeline
  let sh = ss.getSheetByName('KK_TIMELINE');
  if (!sh) sh = ss.getSheetByName('timeline');
  if (!sh) throw new Error("Sheet 'KK_TIMELINE' (oder 'timeline') nicht gefunden.");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const fDefault = 14;
  const f = (futureDays === undefined || futureDays === null || futureDays === "")
    ? fDefault
    : (Number(futureDays) || fDefault);

  if (lastRow < 2 || lastCol < 1) {
    return { ok: true, sheet: sh.getName(), generatedAt: Date.now(), count: 0, headers: [], rows: [], missingHeaders: [], days: null, futureDays: f };
  }

  // Header separat laden (Performance)
  const headersRaw = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const headersLC  = headersRaw.map(h => h.toLowerCase());
  const idxIsToday = headersLC.indexOf('is_today');

  // date Spalte finden (date/datum/day) – via alias helper
  const idxDate = (function(){
    const pick = (aliases) => {
      for (let i = 0; i < aliases.length; i++) {
        const j = headersLC.indexOf(String(aliases[i]).toLowerCase());
        if (j !== -1) return j;
      }
      return -1;
    };
    return pick(['date','datum','day']);
  })();

  // ---- REQUIRED HEADERS CHECK für neue Charts (B–I) ----
  const REQUIRED_GROUPS = [
    { name: "Recovery_vs_Training",
      anyOf: [
        ["coachE_ESS_day", "load_fb_day"],
        ["recovery_7d", "recovery7"],
        ["Garmin_Training_Readiness", "fb_TR_obs", "training_readiness"],
        ["sleep_score_0_100", "sleepscore"]
      ]
    },
    { name: "Load_Forecast",
      allOf: ["coachE_ESS_day", "coachE_CTL_forecast", "coachE_Smart_Gains"]
    },
    { name: "ACWR_TSB",
      anyOf: [
        ["coachE_ACWR_forecast", "fbACWR_obs", "fbacwr_obs"],
        ["coachE_CTL_forecast", "fbCTL_obs", "fbctl_obs"],
        ["coachE_ATL_forecast", "fbATL_obs", "fbatl_obs"]
      ]
    },
    { name: "Fuel_to_Perform",
      anyOf: [
        ["protein_g", "protein"],
        ["carb_g", "carbs_g", "carb"],
        ["fat_g", "fat"],
        ["coachE_ESS_day", "load_fb_day"]
      ]
    },
    { name: "Sleep_Quality_Scatter",
      anyOf: [
        ["sleep_hours", "sleep_h", "sleep_duration_h"],
        ["sleep_score_0_100", "sleepscore"],
        ["coachE_ESS_day", "load_fb_day"],
        ["date", "datum", "day"]
      ]
    },
    { name: "Efficiency_Quadrant",
      anyOf: [
        ["coachE_ESS_day", "load_fb_day"],
        ["coachE_CTL_forecast", "fbCTL_obs", "fbctl_obs"]
      ]
    },
    { name: "Efficiency_Monitor_FB",
      anyOf: [
        ["fbCTL_obs", "fbctl_obs"],
        ["rhr", "resting_hr", "rhr_bpm"],
        ["hrv", "hrv_ms", "hrvstatus", "hrv_status"]
      ]
    }
  ];

  function hasAny(arr){
    return arr.some(h => headersLC.includes(String(h).toLowerCase()));
  }
  function hasAll(arr){
    return arr.every(h => headersLC.includes(String(h).toLowerCase()));
  }

  const missingHeaders = [];
  REQUIRED_GROUPS.forEach(g => {
    if (g.allOf){
      if (!hasAll(g.allOf)) {
        missingHeaders.push({
          chart: g.name,
          missing: g.allOf.filter(h => !headersLC.includes(String(h).toLowerCase()))
        });
      }
    } else if (g.anyOf){
      const groupMissing = [];
      g.anyOf.forEach(group => { if (!hasAny(group)) groupMissing.push(group); });
      if (groupMissing.length) missingHeaders.push({ chart: g.name, missingAnyOfGroups: groupMissing });
    }
  });

  // -------- Parameter normalisieren --------
  const dHist = (days === undefined || days === null || days === "")
    ? null
    : ((Number.isFinite(Number(days)) && Number(days) > 0) ? Math.floor(Number(days)) : null);

  // DEFAULT: futureDays=14 (das ist dein "forecast=14" Wunsch)
  const dFut = (futureDays === undefined || futureDays === null || futureDays === "")
    ? fDefault
    : ((Number.isFinite(Number(futureDays)) && Number(futureDays) > 0) ? Math.floor(Number(futureDays)) : 0);

  const tz = Session.getScriptTimeZone();

  function serializeCell(raw){
    if (raw instanceof Date && !isNaN(raw.getTime())) {
      return Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
    }
    return raw;
  }

  function serializeRows(rowsIn){
    return (rowsIn || []).map(r => (r || []).map(serializeCell));
  }

  // Ohne Date-Spalte: Future/Range nicht möglich -> tail return (Performance-Schutz)
  const MAX_ROWS_WITHOUT_DATE = 1000;
  if (idxDate === -1) {
    const dataStart = Math.max(2, lastRow - MAX_ROWS_WITHOUT_DATE + 1);
    const rowCount = Math.max(0, lastRow - dataStart + 1);
    const rowsNoDate = rowCount > 0 ? sh.getRange(dataStart, 1, rowCount, lastCol).getValues() : [];
    const message = (rowCount < (lastRow - 1))
      ? `Date-Spalte fehlt: Rückgabe auf die letzten ${rowCount} Zeilen begrenzt (von ${lastRow - 1}).`
      : undefined;

    return {
      ok: true,
      sheet: sh.getName(),
      generatedAt: Date.now(),
      count: rowsNoDate.length,
      headers: headersRaw,
      rows: serializeRows(rowsNoDate),
      missingHeaders,
      days: dHist,
      futureDays: dFut,
      ...(message ? { message } : {})
    };
  }

  const fmtKey = (d) => Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd');

  const today = new Date();
  today.setHours(0,0,0,0);
  const todayKey = fmtKey(today);

  function parseSheetDate(raw){
    if (!raw) return null;

    if (raw instanceof Date && !isNaN(raw.getTime())) return raw;

    // Google Sheets serial number (days since 1899-12-30)
    if (typeof raw === "number" && isFinite(raw)) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      const d = new Date(epoch.getTime() + raw * 24 * 60 * 60 * 1000);
      return isNaN(d.getTime()) ? null : d;
    }

    const s = String(raw).trim();
    if (!s) return null;

    // ISO yyyy-mm-dd
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      const d = new Date(s + "T00:00:00");
      return isNaN(d.getTime()) ? null : d;
    }

    // German dd.mm.yyyy
    const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) {
      const dd = Number(m[1]), mm = Number(m[2]), yy = Number(m[3]);
      const d = new Date(yy, mm - 1, dd);
      return isNaN(d.getTime()) ? null : d;
    }

    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  // Nur date-Spalte lesen (Performance)
  const dateCol = sh.getRange(2, idxDate + 1, lastRow - 1, 1).getValues();
  const dateRows = [];
  for (let i = 0; i < dateCol.length; i++) {
    const dt = parseSheetDate(dateCol[i][0]);
    if (!dt) continue;
    dateRows.push({ rowNum: i + 2, key: fmtKey(dt) });
  }

  if (!dateRows.length) {
    return {
      ok: true,
      sheet: sh.getName(),
      generatedAt: Date.now(),
      days: dHist,
      futureDays: dFut,
      count: 0,
      headers: headersRaw,
      rows: [],
      missingHeaders
    };
  }

  let todayRow = -1;
  if (idxIsToday !== -1) {
    const isTodayCol = sh.getRange(2, idxIsToday + 1, lastRow - 1, 1).getValues();
    for (let i = isTodayCol.length - 1; i >= 0; i--) {
      const v = Number(String(isTodayCol[i][0]).replace(',', '.'));
      if (v === 1) {
        todayRow = i + 2;
        break;
      }
    }
  }

  // todayKey bevorzugt aus is_today-Reihe, sonst "heute"
  let effectiveTodayKey = todayKey;
  if (todayRow >= 2) {
    const match = dateRows.find(r => r.rowNum === todayRow);
    if (match && match.key) effectiveTodayKey = match.key;
  }

  const fromHistKey = (dHist && dHist > 0)
    ? (function () {
        const d = new Date(effectiveTodayKey + 'T00:00:00');
        d.setDate(d.getDate() - (dHist - 1));
        return fmtKey(d);
      })()
    : null;

  const toFutureKey = (dFut > 0)
    ? (function () {
        const d = new Date(effectiveTodayKey + 'T00:00:00');
        d.setDate(d.getDate() + (dFut - 1));
        return fmtKey(d);
      })()
    : effectiveTodayKey;

  // Historie + kleiner Vorwärtsbereich um today
  let selectedMeta = dateRows.filter(r => {
    const isInHist = (fromHistKey ? (r.key >= fromHistKey && r.key <= effectiveTodayKey) : (r.key <= effectiveTodayKey));
    const isInFuture = (dFut > 0) ? (r.key >= effectiveTodayKey && r.key <= toFutureKey) : false;
    return isInHist || isInFuture;
  });

  // Merge by rowNum, anschließend stabil sortieren via date key
  const byRow = new Map();
  selectedMeta.forEach(r => byRow.set(r.rowNum, r));
  selectedMeta = Array.from(byRow.values()).sort((a, b) => {
    const k = a.key.localeCompare(b.key);
    return k !== 0 ? k : (a.rowNum - b.rowNum);
  });

  // 4) Optionale harte Obergrenze
  const MAX_EXPORT_ROWS = 800;
  let truncationMessage = '';
  if (selectedMeta.length > MAX_EXPORT_ROWS) {
    const dropped = selectedMeta.length - MAX_EXPORT_ROWS;
    selectedMeta = selectedMeta.slice(-MAX_EXPORT_ROWS);
    truncationMessage = `Timeline export wurde auf ${MAX_EXPORT_ROWS} Zeilen begrenzt (${dropped} ältere Zeilen ausgelassen).`;
  }

  if (!selectedMeta.length) {
    return {
      ok: true,
      sheet: sh.getName(),
      generatedAt: Date.now(),
      days: dHist,
      futureDays: dFut,
      count: 0,
      headers: headersRaw,
      rows: [],
      missingHeaders,
      ...(truncationMessage ? { message: truncationMessage } : {})
    };
  }

  // 3) Nur den minimal nötigen Zeilenbereich in voller Breite laden
  const minRow = selectedMeta.reduce((m, r) => Math.min(m, r.rowNum), Number.POSITIVE_INFINITY);
  const maxRow = selectedMeta.reduce((m, r) => Math.max(m, r.rowNum), 0);
  const spanRows = Math.max(1, maxRow - minRow + 1);
  const fullSlice = sh.getRange(minRow, 1, spanRows, lastCol).getValues();
  const byRowData = new Map();
  for (let i = 0; i < fullSlice.length; i++) {
    byRowData.set(minRow + i, fullSlice[i]);
  }

  const rowsOut = selectedMeta
    .map(r => byRowData.get(r.rowNum))
    .filter(Boolean);
  const outRowsSerialized = serializeRows(rowsOut);
  const message = truncationMessage || undefined;

  return {
    ok: true,
    sheet: sh.getName(),
    generatedAt: Date.now(),
    days: dHist,
    futureDays: dFut,
    count: outRowsSerialized.length,
    headers: headersRaw,
    rows: outRowsSerialized,
    missingHeaders,
    ...(message ? { message } : {})
  };
}
function getTimelinePayloadForCharts(days, futureDays) {
  try {
    const payload = getTimelinePayload(days, futureDays);
    if (payload && Array.isArray(payload.headers) && Array.isArray(payload.rows)) {
      return (payload.ok === false) ? payload : { ...payload, ok: true };
    }
    return {
      ok: false,
      error: true,
      message: 'timeline payload leer oder ungültig (headers/rows fehlen)'
    };
  } catch (err) {
    return {
      ok: false,
      error: true,
      message: String((err && err.stack) || (err && err.message) || err || 'getTimelinePayloadForCharts failed')
    };
  }
}


/**
 * Backward-compatible Apps Script runner endpoint used by charts_V2.
 * Keeps old frontend calls working and proxies to the hardened payload builder.
 */
function getTimelineJsonUniversal(days, futureDays) {
  return getTimelinePayloadForCharts(days, futureDays);
}

function _median_(arr) {
  const a = (arr || []).filter(x => Number.isFinite(x)).slice().sort((x,y)=>x-y);
  if (!a.length) return null;
  const m = Math.floor(a.length/2);
  return (a.length % 2) ? a[m] : (a[m-1] + a[m]) / 2;
}

function _isBadRhrFlag_(rhrSeries, i, lookbackDays, deltaBpm) {
  const lb = lookbackDays || 7;
  const d  = deltaBpm || 5;

  const today = Number(rhrSeries?.[i]);
  if (!Number.isFinite(today)) return false;

  const start = Math.max(0, i - lb);
  const window = [];
  for (let k = start; k < i; k++) {
    const v = Number(rhrSeries?.[k]);
    if (Number.isFinite(v)) window.push(v);
  }
  if (window.length < 3) return false; // zu wenig Historie

  const base = _median_(window);
  if (!Number.isFinite(base)) return false;

  return today >= (base + d);
}

/**
 * Coach-Kira Analyse: STRICT JSON (für Frontend JSON.parse)
 * Erwartet simLast = das Objekt, das du in charts.html persistierst.
 */
function getCoachKiraAnalysis(simLast) {
  // KI-frei (2026-10-26): Gemini-Analyse entfernt – Cockpit nutzen.
  return { success: false, error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' };
}


/** Kürzt simLast, damit Prompt stabil bleibt */
function slimSimLast_(simLast) {
  const o = simLast || {};
  const take = (arr, n) => Array.isArray(arr) ? arr.slice(0, n) : [];

  return {
    generatedAt: o.generatedAt,
    startDate: o.startDate,
    todayIsClosed: !!o.todayIsClosed,
    snapshotActive: !!o.snapshotActive,
    snapshotCreatedAt: o.snapshotCreatedAt,

    labels: take(o.labels, 14),
    loads:  take(o.loads, 14),
    ctl:    take(o.ctl, 14),
    acwr:   take(o.acwr, 14),
    tsb:    take(o.tsb, 14),
    sg:     take(o.sg, 14),

    plannedLoads: take(o.plannedLoads, 14),
    phases:       take(o.phases, 14),
    lockedDays:   take(o.lockedDays, 14),
    plannedSports: take(o.plannedSports, 14),
    plannedZones:  take(o.plannedZones, 14),

    hrvToday: o.hrvToday,
    sleepToday: o.sleepToday,
    monoToday: o.monoToday,
    sleepHoursToday: o.sleepHoursToday,
    sleepScoreToday: o.sleepScoreToday
  };
}


/** Erzwingt: ein parsebares JSON-Objekt */
function coerceStrictJson_(rawText) {
  const s = String(rawText ?? "").trim();
  if (!s) return null;

  // 1) Direkt parsebar?
  try { return JSON.parse(s); } catch (e) {}

  // 2) Häufig: Modell liefert Text + JSON → extrahiere erstes {...}
  const extracted = extractFirstJsonObject_(s);
  if (extracted) {
    try { return JSON.parse(extracted); } catch (e) {}
  }

  // 3) Notfalls: null
  return null;
}

function extractFirstJsonObject_(s) {
  // naive aber robust genug: finde ersten '{' und match bis balanciert
  const start = s.indexOf("{");
  if (start < 0) return null;

  let depth = 0;
  let inStr = false;
  let esc = false;

  for (let i = start; i < s.length; i++) {
    const ch = s[i];

    if (inStr) {
      if (esc) { esc = false; continue; }
      if (ch === "\\") { esc = true; continue; }
      if (ch === '"') inStr = false;
      continue;
    } else {
      if (ch === '"') { inStr = true; continue; }
      if (ch === "{") depth++;
      if (ch === "}") depth--;
      if (depth === 0) return s.slice(start, i + 1);
    }
  }
  return null;
}

// ============================================================
// Coach Kira Dashboard API – Minimalendpunkte für pptx.app
// Stand: MVP
// ============================================================

function ckDash_json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function ckDash_now_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'Europe/Berlin',
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
}

/**
 * (2026-09-28) Token-Schutz auf Wunsch deaktiviert.
 * Wieder einschalten: Prüfung gegen ScriptProperty DASHBOARD_ACTION_TOKEN zurückholen (siehe unten).
 */
function ckDash_requireToken_(p) {
  return { ok: true };
}


function ckDash_getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);

  if (!sh) {
    throw new Error('Sheet nicht gefunden: ' + name);
  }

  return sh;
}

function ckDash_readKeyValueRows_(sheetName, keyCol, valueCol, maxRows) {
  const sh = ckDash_getSheet_(sheetName);
  const lastRow = Math.min(sh.getLastRow(), maxRows || sh.getLastRow());
  const values = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();

  return values.map(function(row, i) {
    return {
      row: i + 1,
      key: row[keyCol - 1],
      value: row[valueCol - 1],
      raw: row
    };
  });
}

function ckDash_status_() {
  try {
    const heartbeat = ckDash_getHeartbeat_();
    const report = ckDash_getReportStatusSummary_();

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      heartbeat: heartbeat,
      summary: report
    });
  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_getHeartbeat_() {
  const sh = ckDash_getSheet_('AI_HEARTBEAT');
  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return null;
  }

  const header = values[0];
  const row = values[values.length - 1];

  return {
    timestamp: row[0],
    stage: row[1],
    runId: row[2],
    extra: row[3],
    header: header
  };
}

function ckDash_getReportStatusSummary_() {
  const sh = ckDash_getSheet_('AI_REPORT_STATUS');
  const values = sh.getDataRange().getValues();

  const out = {
    gesamtAmpel: null,
    gesamtScore: null,
    recoveryScore: null,
    trainingScore: null,
    planStatus: null,
    items: []
  };

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const kategorie = row[0];
    const metrik = row[1];
    const ampel = row[2];
    const wert = row[3];
    const score = row[4];
    const text = row[5];

    if (!metrik) continue;

    out.items.push({
      kategorie: kategorie,
      metrik: metrik,
      ampel: ampel,
      wert: wert,
      score: score,
      text: text
    });

    if (metrik === 'Gesamtscore') {
      out.gesamtAmpel = ampel;
      out.gesamtScore = score || wert;
    }

    if (metrik === 'Recovery Score') {
      out.recoveryScore = score || wert;
    }

    if (metrik === 'Training Score') {
      out.trainingScore = score || wert;
    }

    if (metrik === 'Plan Status') {
      out.planStatus = text || wert || ampel;
    }
  }

  return out;
}

function ckDash_planJson_() {
  try {
    const sh = ckDash_getSheet_('AI_REPORT_PLAN');
    const values = sh.getDataRange().getValues();

    if (values.length < 2) {
      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        plan: []
      });
    }

    const headers = values[0];

    const plan = values.slice(1)
      .filter(function(row) {
        return row[0];
      })
      .map(function(row, idx) {
        const obj = {
          index: idx + 1
        };

        headers.forEach(function(h, colIdx) {
          if (h) obj[h] = row[colIdx];
        });

        return obj;
      });

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      rows: plan.length,
      plan: plan
    });
  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_chartData_(p) {
  try {
    const rangeParam = p.range || '28d';
    const metricsParam = p.metrics || 'load_fb_day,fbATL_obs,fbCTL_obs,fbACWR_obs,sleep_hours,sleep_score_0_100,rhr_bpm,Garmin_Training_Readiness';

    const days = ckDash_parseDays_(rangeParam);
    const requestedMetrics = metricsParam.split(',')
      .map(function(s) { return String(s).trim(); })
      .filter(Boolean);

    const sh = ckDash_getSheet_('KK_TIMELINE');
    const values = sh.getDataRange().getValues();

    if (values.length < 2) {
      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        rows: 0,
        data: []
      });
    }

    const headers = values[0].map(function(h) {
      return String(h || '').trim();
    });

    const dateIdx = headers.indexOf('date');
    if (dateIdx < 0) {
      throw new Error('Spalte date nicht gefunden in KK_TIMELINE');
    }

    const metricIndexes = requestedMetrics.map(function(m) {
      return {
        name: m,
        index: headers.indexOf(m)
      };
    }).filter(function(x) {
      return x.index >= 0;
    });

    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';

    const todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    const todayDate = new Date(todayKey + 'T23:59:59');
    const minDate = new Date(todayDate.getTime() - (days - 1) * 24 * 60 * 60 * 1000);

    const data = [];

    for (let i = 1; i < values.length; i++) {
  const row = values[i];
  const d = row[dateIdx];

  if (!d || !(d instanceof Date)) continue;
  if (d < minDate) continue;
  if (d > todayDate) continue;

  const obj = {
    row: i + 1,
    date: Utilities.formatDate(d, tz, 'yyyy-MM-dd')
  };

  metricIndexes.forEach(function(m) {
    obj[m.name] = row[m.index];
  });

  data.push(obj);
}

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      range: rangeParam,
      requestedMetrics: requestedMetrics,
      returnedMetrics: metricIndexes.map(function(m) { return m.name; }),
      rows: data.length,
      data: data
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_parseDays_(rangeParam) {
  const s = String(rangeParam || '28d').toLowerCase();
  const m = s.match(/^(\d+)d$/);

  if (m) {
    return Math.max(1, Math.min(365, Number(m[1])));
  }

  return 28;
}

function ckDash_queueAction_(p, functionName) {
  try {
    const auth = ckDash_requireToken_(p);

    if (!auth.ok) {
      return ckDash_json_({
        ok: false,
        timestamp: ckDash_now_(),
        error: auth.error
      });
    }

    const allowed = {
      runKiraGeminiSupervisor: true,
      generateOptimized14DayPlan: true
    };

    if (!allowed[functionName]) {
      return ckDash_json_({
        ok: false,
        timestamp: ckDash_now_(),
        error: 'action_not_allowed',
        functionName: functionName
      });
    }

    // KI-frei (Stufe 3): gewünschte KI (lokal/Cloud) für den nächsten Supervisor-Auftrag merken
    if (functionName === 'runKiraGeminiSupervisor') {
      PropertiesService.getScriptProperties().setProperty('CK_SUP_ENGINE_NEXT', String(p.engine || '') === 'cloud' ? 'cloud' : 'local');
    }

    const lock = LockService.getScriptLock();

    if (!lock.tryLock(5000)) {
      return ckDash_json_({
        ok: false,
        timestamp: ckDash_now_(),
        error: 'lock_busy'
      });
    }

    try {
      const runId = 'DASH_' + Utilities.getUuid().slice(0, 8);

      PropertiesService
        .getScriptProperties()
        .setProperty('DASH_LAST_QUEUED_RUN_ID', runId);

      PropertiesService
        .getScriptProperties()
        .setProperty('DASH_LAST_QUEUED_FUNCTION', functionName);

      PropertiesService
        .getScriptProperties()
        .setProperty('DASH_LAST_QUEUED_AT', ckDash_now_());

      ckDash_deleteTriggersForFunction_(functionName);

      ScriptApp
        .newTrigger(functionName)
        .timeBased()
        .after(1000)
        .create();

      try {
        logToSheet('INFO', '[Dashboard] Queued ' + functionName + ' runId=' + runId);
      } catch (logErr) {
        // Logging darf nie den Start verhindern.
      }

      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        queued: true,
        functionName: functionName,
        runId: runId
      });

    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_deleteTriggersForFunction_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction && trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function ckDash_runStatus_(p) {
  try {
    const props = PropertiesService.getScriptProperties();

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      requestedRunId: p.runId || null,
      lastQueued: {
        runId: props.getProperty('DASH_LAST_QUEUED_RUN_ID'),
        functionName: props.getProperty('DASH_LAST_QUEUED_FUNCTION'),
        queuedAt: props.getProperty('DASH_LAST_QUEUED_AT')
      },
      heartbeat: ckDash_getHeartbeat_()
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_logs_(p) {
  try {
    const limit = Math.max(1, Math.min(100, Number(p.limit || 25)));
    const sh = ckDash_getSheet_('AI_LOG');

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 1) {
      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        logs: []
      });
    }

    const startRow = Math.max(1, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;

    const values = sh.getRange(startRow, 1, numRows, lastCol).getValues();

    const logs = values.map(function(row, idx) {
      return {
        row: startRow + idx,
        timestamp: row[0],
        level: row[1],
        message: row[2],
        raw: row
      };
    });

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      rows: logs.length,
      logs: logs
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_planJsonV2_() {
  try {
    const sh = ckDash_getSheet_('AI_REPORT_PLAN');
    const values = sh.getDataRange().getValues();

    if (values.length < 2) {
      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        rows: 0,
        plan: []
      });
    }

    const plan = values.slice(1)
      .filter(function(row) {
        return row[0];
      })
      .map(function(row, idx) {
        const d = row[0];

        return {
          index: idx + 1,
          date: d instanceof Date
            ? Utilities.formatDate(d, Session.getScriptTimeZone() || 'Europe/Berlin', 'yyyy-MM-dd')
            : String(d || ''),
          day: row[1] || '',
          original_load_ess: Number(row[2] || 0),
          recommended_load_ess: Number(row[3] || 0),
          recommended_zone: row[4] || '',
          training_1: row[5] || '',
          training_2: row[6] || '',
          projected_effect: row[7] || '',
          weather_recommendation: row[8] || ''
        };
      });

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      rows: plan.length,
      plan: plan
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}


function ckDash_planSimulationJson_() {
  var CACHE_KEY = 'ckDash_planSimulationJson_v1';
  var CACHE_TTL = 60;
  var DEPLOY_MARKER = 'v2-cache-2026-09-18';

  try {
    var cache = CacheService.getScriptCache();
    var hit = cache.get(CACHE_KEY);
    if (hit) {
      return ContentService
        .createTextOutput(hit)
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (typeof getSimStartValues !== 'function') {
      throw new Error('getSimStartValues() nicht gefunden.');
    }

    var base = getSimStartValues();
    var tz = Session.getScriptTimeZone() || 'Europe/Berlin';

    var startDate = base.startDate
      ? new Date(base.startDate)
      : new Date();

    var dayNames = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

    var fmtDate = function(d) {
      return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    };

    var toNum = function(v, def) {
      var n = parseFloat(String(v).replace(',', '.'));
      return isFinite(n) ? n : (def || 0);
    };

    var days = [];
    for (var i = 0; i < 14; i++) {
      var d = new Date(startDate.getTime());
      d.setDate(startDate.getDate() + i);

      days.push({
        index: i + 1,
        date: fmtDate(d),
        day: dayNames[d.getDay()],
        phase: (base.phases && base.phases[i]) || 'E',
        load: toNum(base.plannedLoads && base.plannedLoads[i], 0),
        sport: (base.plannedSports && base.plannedSports[i]) || '',
        zone: (base.plannedZones && base.plannedZones[i]) || '',
        te_ae: toNum(base.plannedTeAe && base.plannedTeAe[i], 0),
        te_an: toNum(base.plannedTeAn && base.plannedTeAn[i], 0),
        locked: !!(base.lockedDays && base.lockedDays[i])
      });
    }

    var payload = {
      ok: true,
      timestamp: ckDash_now_(),
      source: 'getSimStartValues',
      cached: false,
      deployMarker: DEPLOY_MARKER,
      base: {
        atl: toNum(base.atl, 0),
        ctl: toNum(base.ctl, 0),
        ctlHistory: Array.isArray(base.ctlHistoryYesterday)
          ? base.ctlHistoryYesterday
          : (Array.isArray(base.ctlHistory) ? base.ctlHistory : []),
        config: base.config || {},
        todayIsClosed: !!base.todayIsClosed,
        startDate: fmtDate(startDate),
        todayRowIndex: base.todayRowIndex,
        startRowIndex: base.startRowIndex,
        snapshotActive: !!base.snapshotActive
      },
      days: days
    };

    var json = JSON.stringify(payload, null, 2);

    if (json.length < 95 * 1024) {
      try { cache.put(CACHE_KEY, json, CACHE_TTL); } catch (_) { /* ignore */ }
    }

    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_statusAiAnalysis_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('AI_STATUS_ANALYSIS');
    if (!sh || sh.getLastRow() < 2) {
      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        analysis: null
      });
    }

    const row = sh.getRange(sh.getLastRow(), 1, 1, 3).getValues()[0];
    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      analysisTimestamp: row[0],
      analysis: row[1]
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_writeStatusAiAnalysis_(analysis, raw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('AI_STATUS_ANALYSIS');
  if (!sh) {
    sh = ss.insertSheet('AI_STATUS_ANALYSIS');
    sh.getRange(1, 1, 1, 3).setValues([['timestamp', 'analysis', 'raw']]);
  }
  sh.appendRow([new Date(), analysis, raw || '']);
}

// ============================================================
// Coach Kira Dashboard API – Iteration 3 Zusatzfunktionen
// Plan Save + Perplexity Status-AI
// Am Ende von KiraGeminiSupervisor.gs einfügen
// ============================================================

function ckDash_parseJsonArray_(raw, fallback) {
  if (!raw) return fallback || [];
  try {
    const v = JSON.parse(raw);
    return Array.isArray(v) ? v : (fallback || []);
  } catch (e) {
    return fallback || [];
  }
}

function ckDash_saveSimulatedPlan_(p) {
  try {
    const auth = ckDash_requireToken_(p);
    if (!auth.ok) {
      return ckDash_json_({
        ok: false,
        timestamp: ckDash_now_(),
        error: auth.error
      });
    }

    if (typeof saveSimulatedPlan !== 'function') {
      throw new Error('saveSimulatedPlan(...) nicht gefunden.');
    }

    const loads = ckDash_parseJsonArray_(p.loads, []);
    const teAe = ckDash_parseJsonArray_(p.teAe, []);
    const teAn = ckDash_parseJsonArray_(p.teAn, []);
    const sports = ckDash_parseJsonArray_(p.sports, []);
    const zones = ckDash_parseJsonArray_(p.zones, []);
    const locks = ckDash_parseJsonArray_(p.locks, []);

    if (!loads.length) throw new Error('loads ist leer.');

    const raw = saveSimulatedPlan(loads, teAe, teAn, sports, zones, locks);
    const ok = String(raw || '').indexOf('✅') === 0;

    try {
      logToSheet('INFO', '[Dashboard] saveSimulatedPlan via cockpit: ' + String(raw));
    } catch (e) {}

    return ckDash_json_({
      ok: ok,
      timestamp: ckDash_now_(),
      message: String(raw || ''),
      raw: String(raw || '')
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_runStatusAiAnalysis_(p) {
  // KI-frei (2026-10-26): Statusanalyse wird im Cockpit erstellt (lokal oder Cloud) und per saveBriefing gespeichert.
  return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' });
}

function ckDash_statusAiAnalysis_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('AI_STATUS_ANALYSIS');

    if (!sh || sh.getLastRow() < 2) {
      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        analysis: null
      });
    }

    const row = sh.getRange(sh.getLastRow(), 1, 1, 3).getValues()[0];

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      analysisTimestamp: row[0],
      analysis: row[1]
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_writeStatusAiAnalysis_(analysis, raw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('AI_STATUS_ANALYSIS');

  if (!sh) {
    sh = ss.insertSheet('AI_STATUS_ANALYSIS');
    sh.getRange(1, 1, 1, 3).setValues([
      ['timestamp', 'analysis', 'raw']
    ]);
  }

  sh.appendRow([
    new Date(),
    analysis,
    raw || ''
  ]);
}

function ckDash_planBriefing_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('AI_PLAN_BRIEFING');

    if (!sh || sh.getLastRow() < 2) {
      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        briefing: null,
        recommendations: null
      });
    }

    const row = sh.getRange(sh.getLastRow(), 1, 1, 4).getValues()[0];
    let recommendations = null;

    try {
      recommendations = row[2] ? JSON.parse(row[2]) : null;
    } catch (e) {
      recommendations = null;
    }

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      briefingTimestamp: row[0],
      briefing: row[1],
      recommendations: recommendations,
      raw: row[3] || ''
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_generatePlanBriefing_(p) {
  // KI-frei (2026-10-26): Plan-Briefing wird im Cockpit erstellt (lokal oder Cloud) und per saveBriefing gespeichert.
  return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' });
}

function ckDash_collectPlanBriefingContext_() {
  const sim = typeof getSimStartValues === 'function' ? getSimStartValues() : {};
  const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
  const startDate = sim.startDate ? new Date(sim.startDate) : new Date();
  const dayNames = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  const toNum = function(v, def) {
    const n = parseFloat(String(v).replace(',', '.'));
    return Number.isFinite(n) ? n : (def || 0);
  };
  const fmtDate = function(d) {
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  };

  const days = [];
  for (let i = 0; i < 14; i++) {
    const d = new Date(startDate.getTime());
    d.setDate(startDate.getDate() + i);
    days.push({
      index: i + 1,
      date: fmtDate(d),
      day: dayNames[d.getDay()],
      phase: (sim.phases && sim.phases[i]) || '',
      current_load_ess: toNum(sim.plannedLoads && sim.plannedLoads[i], 0),
      sport: (sim.plannedSports && sim.plannedSports[i]) || '',
      zone: (sim.plannedZones && sim.plannedZones[i]) || '',
      te_ae: toNum(sim.plannedTeAe && sim.plannedTeAe[i], 0),
      te_an: toNum(sim.plannedTeAn && sim.plannedTeAn[i], 0),
      locked: !!(sim.lockedDays && sim.lockedDays[i])
    });
  }

  return {
    base: {
      atl: toNum(sim.atl, 0),
      ctl: toNum(sim.ctl, 0),
      todayIsClosed: !!sim.todayIsClosed,
      startDate: fmtDate(startDate),
      config: sim.config || {}
    },
    days: days
  };
}

function ckDash_parsePlanBriefingContent_(content) {
  let txt = String(content || '').trim();
  txt = txt.replace(/^```json\\s*/i, '').replace(/^```\\s*/i, '').replace(/```$/i, '').trim();

  try {
    const parsed = JSON.parse(txt);
    return {
      briefing: String(parsed.briefing || ''),
      recommendations: Array.isArray(parsed.recommendations) ? parsed.recommendations : []
    };
  } catch (e) {
    return {
      briefing: txt,
      recommendations: []
    };
  }
}

function ckDash_writePlanBriefing_(briefing, recommendations, raw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('AI_PLAN_BRIEFING');

  if (!sh) {
    sh = ss.insertSheet('AI_PLAN_BRIEFING');
    sh.getRange(1, 1, 1, 4).setValues([
      ['timestamp', 'briefing', 'recommendations_json', 'raw']
    ]);
  }

  sh.appendRow([
    new Date(),
    briefing || '',
    JSON.stringify(recommendations || []),
    raw || ''
  ]);
}

function ckDash_consolidatedAnalysis_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('AI_CONSOLIDATED_ANALYSIS');
    if (!sh || sh.getLastRow() < 2) {
      return ckDash_json_({
        ok: true,
        timestamp: ckDash_now_(),
        analysis: null
      });
    }

    const row = sh.getRange(sh.getLastRow(), 1, 1, 3).getValues()[0];
    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      analysisTimestamp: row[0],
      analysis: row[1],
      raw: row[2] || ''
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

function ckDash_runConsolidatedAnalysis_(p) {
  // KI-frei (2026-10-26): Konsolidierte Analyse wird im Cockpit erstellt (lokal oder Cloud) und per saveBriefing gespeichert.
  return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' });
}

function ckDash_collectConsolidatedContext_() {
  const status = ckDash_getReportStatusSummary_();
  const planBriefing = ckDash_getLatestSheetJsonLike_('AI_PLAN_BRIEFING', 4);
  const sim = typeof getSimStartValues === 'function' ? getSimStartValues() : {};
  const history = ckDash_readTailObjects_('AI_DATA_HISTORY', 21);
  const forecast = ckDash_readTailObjects_('AI_DATA_FORECAST', 21);
  const activityReviews = ckDash_readTailObjects_('AI_ACTIVITY_REVIEWS', 10);

  return {
    generatedAt: ckDash_now_(),
    status: status,
    planBriefing: planBriefing,
    planSimulation: ckDash_slimSimStartValues_(sim),
    historyTail: history,
    forecastTail: forecast,
    activityReviewsTail: activityReviews
  };
}

function ckDash_slimSimStartValues_(sim) {
  const toNum = function(v, def) {
    const n = parseFloat(String(v).replace(',', '.'));
    return Number.isFinite(n) ? n : (def || 0);
  };
  const days = [];
  const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
  const startDate = sim.startDate ? new Date(sim.startDate) : new Date();
  const dayNames = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  const fmtDate = function(d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
  for (let i = 0; i < 14; i++) {
    const d = new Date(startDate.getTime());
    d.setDate(startDate.getDate() + i);
    days.push({
      date: fmtDate(d),
      day: dayNames[d.getDay()],
      load: toNum(sim.plannedLoads && sim.plannedLoads[i], 0),
      sport: (sim.plannedSports && sim.plannedSports[i]) || '',
      zone: (sim.plannedZones && sim.plannedZones[i]) || '',
      te_ae: toNum(sim.plannedTeAe && sim.plannedTeAe[i], 0),
      te_an: toNum(sim.plannedTeAn && sim.plannedTeAn[i], 0),
      locked: !!(sim.lockedDays && sim.lockedDays[i])
    });
  }
  return {
    base: {
      atl: toNum(sim.atl, 0),
      ctl: toNum(sim.ctl, 0),
      todayIsClosed: !!sim.todayIsClosed
    },
    days: days
  };
}

function ckDash_getLatestSheetJsonLike_(sheetName, columns) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return null;
    const row = sh.getRange(sh.getLastRow(), 1, 1, columns || sh.getLastColumn()).getValues()[0];
    return row;
  } catch (e) {
    return null;
  }
}

function ckDash_readTailObjects_(sheetName, limit) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return [];
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const start = Math.max(2, lastRow - (limit || 10) + 1);
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const rows = sh.getRange(start, 1, lastRow - start + 1, lastCol).getValues();
    return rows.map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) {
        if (h) obj[h] = row[i];
      });
      return obj;
    });
  } catch (e) {
    return [];
  }
}

function ckDash_writeConsolidatedAnalysis_(analysis, raw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('AI_CONSOLIDATED_ANALYSIS');
  if (!sh) {
    sh = ss.insertSheet('AI_CONSOLIDATED_ANALYSIS');
    sh.getRange(1, 1, 1, 3).setValues([
      ['timestamp', 'analysis', 'raw']
    ]);
  }
  sh.appendRow([new Date(), analysis || '', raw || '']);
}

/**
 * saveWellbeing — speichert das morgendliche Befinden (1–5) für heute oder gestern.
 * Schreibt in 'timeline' (Quelle) UND 'KK_TIMELINE' (Spiegel), Spalte 'befinden_1_5'.
 * Parameter: token, date (yyyy-MM-dd), value (1–5)
 */
function ckDash_saveWellbeing_(p) {
  var COL_NAME = 'befinden_1_5';
  var lock = LockService.getScriptLock();
  try {
    var auth = ckDash_requireToken_(p);
    if (!auth.ok) {
      return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: auth.error });
    }

    var value = parseInt(String(p.value || '').trim(), 10);
    if (!(value >= 1 && value <= 5)) throw new Error('value muss 1–5 sein.');

    var tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var y = new Date();
    y.setDate(y.getDate() - 1);
    var yesterdayKey = Utilities.formatDate(y, tz, 'yyyy-MM-dd');

    var dateKey = String(p.date || todayKey).trim().slice(0, 10);
    if (dateKey !== todayKey && dateKey !== yesterdayKey) {
      throw new Error('Nur heute oder gestern speicherbar (' + dateKey + ').');
    }

    lock.waitLock(10000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var written = [];
    ['timeline', 'KK_TIMELINE'].forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) throw new Error('Blatt ' + name + ' nicht gefunden.');

      var lastCol = sh.getLastColumn();
      var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
        return String(h || '').trim().toLowerCase();
      });
      var colIdx = headers.indexOf(COL_NAME);
      var dateIdx = headers.indexOf('date');
      if (colIdx < 0) throw new Error('Spalte ' + COL_NAME + ' fehlt in ' + name + '.');
      if (dateIdx < 0) throw new Error('Spalte date fehlt in ' + name + '.');

      var lastRow = sh.getLastRow();
      var dates = sh.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues();
      var row = -1;
      for (var i = 0; i < dates.length; i++) {
        var d = dates[i][0];
        var key = (d instanceof Date)
          ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
          : String(d || '').trim().slice(0, 10);
        if (key === dateKey) { row = i + 2; break; }
      }
      if (row < 0) throw new Error('Datum ' + dateKey + ' nicht in ' + name + ' gefunden.');

      sh.getRange(row, colIdx + 1).setValue(value);
      written.push(name + '!R' + row);
    });

    SpreadsheetApp.flush();
    try { logToSheet('INFO', '[Dashboard] saveWellbeing ' + dateKey + ' = ' + value + ' → ' + written.join(', ')); } catch (_) {}

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      date: dateKey,
      value: value,
      written: written
    });

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * submitCockpitEntry — Cockpit-Eingabe anstelle des Google-Formulars.
 * Baut ein Formular-Event (namedValues) und ruft die UNVERÄNDERTE onFormSubmit(e) auf.
 * Damit laufen Schreiben, Tageswechsel, IST→SOLL, Activity Review und KI-Report wie beim Formular.
 * Parameter: token, type (Morgens | Nach Aktivität | Abends), date (yyyy-MM-dd), values (JSON-Objekt)
 */
function ckDash_submitCockpitEntry_(p) {
  const ALLOWED = {
    'Morgens': ['garminATL', 'garminCTL', 'garminACWR', 'sleep_hours', 'sleep_score_0_100', 'rhr_bpm',
                'hrv_status', 'hrv_threshholds', 'Garmin_Training_Readiness', 'Trainingszustand'],
    'Nach Aktivität': ['load_fb_day', 'fbATL_obs', 'fbCTL_obs', 'fbACWR_obs', 'garminEnduranceScore',
                       'Sport_x', 'Zone', 'fb_TR_obs', 'Aerobic_TE', 'Anaerobic_TE'],
    'Abends': ['kcal_in', 'kcal_out', 'deficit', 'carb_g', 'protein_g', 'fat_g', 'fiber_g']
  };
  const TEXT_COLS = ['Sport_x', 'Zone', 'hrv_threshholds', 'Trainingszustand'];
  const lock = LockService.getDocumentLock();
  const started = Date.now();

  try {
    const auth = ckDash_requireToken_(p);
    if (!auth.ok) return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: auth.error });

    const type = String(p.type || '').trim();
    if (!ALLOWED[type]) throw new Error('Unbekannter Typ: ' + type);
    if (typeof onFormSubmit !== 'function') throw new Error('onFormSubmit() nicht gefunden.');

    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const now = new Date();
    const todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const y = new Date(now.getTime()); y.setDate(y.getDate() - 1);
    const yesterdayKey = Utilities.formatDate(y, tz, 'yyyy-MM-dd');
    const dateKey = String(p.date || todayKey).trim().slice(0, 10);
    if (type === 'Morgens' && dateKey !== todayKey) throw new Error('Morgenwerte nur für heute (' + todayKey + ').');
    if (dateKey !== todayKey && dateKey !== yesterdayKey) throw new Error('Nur heute oder gestern erlaubt (' + dateKey + ').');

    let values = {};
    try { values = JSON.parse(String(p.values || '{}')); } catch (_) { throw new Error('values ist kein gültiges JSON.'); }

    const [yy, mm, dd] = dateKey.split('-');
    const namedValues = {
      'Zeitstempel': [Utilities.formatDate(now, tz, 'dd.MM.yyyy HH:mm:ss')],
      'Was willst Du eingeben?': [type],
      'Datum': [dd + '.' + mm + '.' + yy]
    };
    const used = [];
    ALLOWED[type].forEach(function(col) {
      if (!Object.prototype.hasOwnProperty.call(values, col)) return;
      let v = values[col];
      if (v === null || v === undefined || String(v).trim() === '') return;
      v = TEXT_COLS.indexOf(col) >= 0 ? String(v).trim() : String(v).trim().replace('.', ','); // wie Formular: Dezimalkomma
      namedValues[col] = [v];
      used.push(col);
    });
    if (!used.length) throw new Error('Keine Werte übergeben.');

    lock.waitLock(30000);
    onFormSubmit({ namedValues: namedValues, values: [], source: 'cockpit' });
    SpreadsheetApp.flush();

    // Rücklesen zur Bestätigung
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_TIMELINE_SHEET);
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
    const dateIdx = headers.indexOf('date');
    const dates = sh.getRange(2, dateIdx + 1, sh.getLastRow() - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < dates.length; i++) {
      const d = dates[i][0];
      const k = (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || '').slice(0, 10);
      if (k === dateKey) { row = i + 2; break; }
    }
    const readback = {};
    if (row > 0) {
      const rowVals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
      used.forEach(function(col) {
        const idx = headers.indexOf(col);
        if (idx >= 0) readback[col] = rowVals[idx] instanceof Date ? String(rowVals[idx]) : rowVals[idx];
      });
    }

    // Protokoll
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let log = ss.getSheetByName('COCKPIT_ENTRIES');
    if (!log) {
      log = ss.insertSheet('COCKPIT_ENTRIES');
      log.getRange(1, 1, 1, 5).setValues([['timestamp', 'type', 'date', 'row', 'values_json']]).setFontWeight('bold');
    }
    log.appendRow([namedValues['Zeitstempel'][0], type, dateKey, row, JSON.stringify(values)]);
    try { logToSheet('INFO', '[Dashboard] submitCockpitEntry ' + type + ' ' + dateKey + ' → Zeile ' + row + ' (' + used.join(', ') + ')'); } catch (_) {}

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      type: type,
      date: dateKey,
      row: row,
      written: used,
      readback: readback,
      triggersStarted: dateKey === todayKey,
      durationMs: Date.now() - started
    });
  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ============================================================================
// Kira-Chat (Dashboard, Heute-Seite) – Perplexity sonar-pro
// ============================================================================
const CK_CHAT_MODEL = 'sonar-pro';
const CK_CHAT_DAILY_LIMIT = 50;

function ckDash_chat_(e) {
  // KI-frei (2026-10-26): Der Chat läuft im Cockpit (lokal oder Cloud über den Cockpit-Server). Ablage weiter per chatLog.
  return ckDash_json_({ ok: false, error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' });
}

// Protokoll lokaler Kira-Antworten (Ollama auf dem Mac) in AI_CHAT_LOG
function ckDash_chatLog_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const q = String(b.question || '').slice(0, 4000).trim();
    const a = String(b.answer || '').slice(0, 8000).trim();
    const m = String(b.model || 'lokal').replace(/[^\w:.\-]/g, '').slice(0, 40);
    const eng = String(b.engine || 'local') === 'cloud' ? 'Cloud' : 'lokal';
    if (!q || !a) return ckDash_json_({ ok: false, error: 'Frage oder Antwort fehlt.' });
    logChatMessage('Marc (Cockpit)', q);
    logChatMessage('Kira (' + eng + ', ' + m + ')', a);
    return ckDash_json_({ ok: true });
  } catch (err) {
    return ckDash_json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// KI-frei (Stufe 3): ckDash_buildBriefingMessages_ entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

function ckDash_briefingPrompt_(p) {
  // KI-frei (Stufe 3): Prompt-Bau liegt in kira_prompts (Cockpit/Mac). Daten: mode=briefingData.
  return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'briefingPrompt entfällt – bitte mode=briefingData nutzen.' });
}

/** KI-frei (Stufe 3): Daten für Status-, Plan- und konsolidiertes Briefing. */
function ckDash_briefingData_(p) {
  try {
    const type = String((p && p.type) || '');
    let daten;
    if (type === 'status') daten = { status: ckDash_getReportStatusSummary_() };
    else if (type === 'plan') daten = { status: ckDash_getReportStatusSummary_(), sim: ckDash_collectPlanBriefingContext_() };
    else if (type === 'consolidated') daten = { context: ckDash_collectConsolidatedContext_() };
    else throw new Error('Unbekannter Briefing-Typ: ' + type);
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), type: type, daten: JSON.parse(JSON.stringify(daten)) });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

function ckDash_parseBriefingJsonLocal_(content) {
  let txt = String(content || '').trim();
  txt = txt.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
  try {
    const parsed = JSON.parse(txt);
    return {
      briefing: String(parsed.briefing || ''),
      recommendations: Array.isArray(parsed.recommendations) ? parsed.recommendations : []
    };
  } catch (e) {
    return { briefing: txt, recommendations: [] };
  }
}

function ckDash_saveBriefing_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const type = String(b.type || '');
    const text = String(b.text || '').slice(0, 30000).trim();
    const model = String(b.model || 'lokal').replace(/[^\w:.\-]/g, '').slice(0, 40);
    if (!text) throw new Error('Kein Briefing-Text übermittelt.');
    const engine = String(b.engine || 'local') === 'cloud' ? 'cloud' : 'local';
    const raw = JSON.stringify({ engine: engine, model: model, savedAt: ckDash_now_() });

    if (type === 'status') {
      ckDash_writeStatusAiAnalysis_(text, raw);
      return ckDash_json_({ ok: true, timestamp: ckDash_now_(), analysis: text });
    }
    if (type === 'consolidated') {
      ckDash_writeConsolidatedAnalysis_(text, raw);
      return ckDash_json_({ ok: true, timestamp: ckDash_now_(), analysis: text });
    }
    if (type === 'plan') {
      const parsed = ckDash_parseBriefingJsonLocal_(text);
      ckDash_writePlanBriefing_(parsed.briefing, parsed.recommendations, raw);
      return ckDash_json_({ ok: true, timestamp: ckDash_now_(), briefing: parsed.briefing, recommendations: parsed.recommendations });
    }
    throw new Error('Unbekannter Briefing-Typ: ' + type);
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

function ckDash_briefingSources_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const defs = [
      ['status', 'AI_STATUS_ANALYSIS', 3],
      ['consolidated', 'AI_CONSOLIDATED_ANALYSIS', 3],
      ['plan', 'AI_PLAN_BRIEFING', 4]
    ];
    const out = {};
    defs.forEach(function (d) {
      const sh = ss.getSheetByName(d[1]);
      if (!sh || sh.getLastRow() < 2) { out[d[0]] = null; return; }
      const row = sh.getRange(sh.getLastRow(), 1, 1, d[2]).getValues()[0];
      let engine = 'cloud';
      let model = null;
      try {
        const r = JSON.parse(String(row[d[2] - 1] || '{}'));
        if (r && r.engine === 'local') engine = 'local';
        model = (r && r.model) ? String(r.model) : null;
      } catch (e) {}
      out[d[0]] = { engine: engine, model: model, timestamp: row[0] };
    });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), sources: out });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

// ============================================================================
// Aktivitäts-Kommentare lokal (Mac-Dienst) – speichern und lesen
// ============================================================================
function ckDash_activitySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('AI_ACTIVITY_REVIEWS');
  if (!sh) {
    sh = ss.insertSheet('AI_ACTIVITY_REVIEWS');
    sh.getRange(1, 1, 1, 10).setValues([[
      'Datum', 'Load (IST)', 'Load (SOLL)', 'ATL (IST)', 'ATL (SOLL)',
      'CTL (IST)', 'CTL (SOLL)', 'ACWR (IST)', 'ACWR (SOLL)', 'KI Review Text'
    ]]);
  }
  if (String(sh.getRange(1, 11).getValue() || '') === '') {
    sh.getRange(1, 11).setValue('Quelle');
  }
  return sh;
}

function ckDash_saveActivityReview_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const date = String(b.date || '');
    const m = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) throw new Error('date fehlt oder ist nicht YYYY-MM-DD.');
    const text = String(b.text || '').slice(0, 4000).trim();
    if (!text) throw new Error('Kein Kommentar-Text übermittelt.');
    const model = String(b.model || 'lokal').replace(/[^\w:.\-]/g, '').slice(0, 40);
    const ist = b.ist || {};
    const soll = b.soll || {};
    const val = function (v) { return (v === null || v === undefined || v === '') ? 'N/A' : v; };
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);

    const sh = ckDash_activitySheet_();
    sh.appendRow([
      d,
      val(ist.load), val(soll.load),
      val(ist.atl), val(soll.atl),
      val(ist.ctl), val(soll.ctl),
      val(ist.acwr), val(soll.acwr),
      text,
      JSON.stringify({ engine: 'local', model: model, savedAt: ckDash_now_() })
    ]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd.MM.yyyy');
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), date: date });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

function ckDash_activityReviews_(p) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AI_ACTIVITY_REVIEWS');
    if (!sh || sh.getLastRow() < 2) return ckDash_json_({ ok: true, timestamp: ckDash_now_(), reviews: [] });
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const lastCol = Math.max(10, Math.min(11, sh.getLastColumn()));
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    const byDate = {};
    rows.forEach(function (r) {
      if (!(r[0] instanceof Date)) return;
      const date = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
      let text = String(r[9] || '').trim();
      if (!text) return;
      if (text.charAt(0) === '{') {
        try { const j = JSON.parse(text); text = String(j.coach_feedback || j.review || j.text || text); } catch (x) {}
      }
      let engine = 'cloud';
      let model = 'gemini';
      try {
        const q = JSON.parse(String(r[10] || '{}'));
        if (q && q.engine === 'local') engine = 'local';
        if (q && q.model) model = String(q.model);
      } catch (x) {}
      byDate[date] = {
        date: date, text: text, engine: engine, model: model,
        loadIst: r[1], loadSoll: r[2], acwrIst: r[7]
      };
    });
    const limit = Math.max(1, Math.min(2000, parseInt((p && p.limit) || '500', 10) || 500));
    const reviews = Object.keys(byDate).sort().slice(-limit).map(function (k) { return byDate[k]; });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), reviews: reviews });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

// ============================================================================
// Supervisor lokal (Mac-Dienst, coach-kira) – Auftrag ablegen, abholen, speichern
// ============================================================================
function ckSup_engine_() {
  return 'local'; // KI-frei (2026-10-26): Supervisor-KI läuft immer auf dem Mac (lokal oder Cloud über den Cockpit-Server).
}

// KI-frei (Stufe 3): ckSup_schema_ entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

function ckSup_promptData_(data, fitnessScores, nutritionScore, gesamtScoreNum, gesamtScoreAmpel, zoneFactors, elevFactors, weekConfigCSV, isActivityDone, wetterDaten) {
  const heute = (data && data.heute) || {};
  const tz = Session.getScriptTimeZone();

  // TE-Balance heute aus AI_DATA_FORECAST (null = nicht gefunden → Prompt nutzt Monotonie-Varianz)
  let teBalanceHeute = null;
  try {
    const fcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AI_DATA_FORECAST');
    if (fcSheet && fcSheet.getLastRow() > 1) {
      const fcData = fcSheet.getDataRange().getValues();
      const fcHeaders = fcData[0].map(h => h.toString().trim());
      const fcHeadersLc = fcHeaders.map(h => h.toLowerCase());

      const teIdx = fcHeadersLc.indexOf('te balance') > -1
        ? fcHeadersLc.indexOf('te balance')
        : fcHeadersLc.indexOf('te balance (% intensiv)');

      if (teIdx !== -1) {
        const tz = Session.getScriptTimeZone();
        const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
        let todayRowIdx = -1;

        // 1) Bevorzugt: is_today == 1
        const isTodayIdx = fcHeadersLc.indexOf('is_today');
        if (isTodayIdx !== -1) {
          for (let r = 1; r < fcData.length; r++) {
            if (parseGermanFloat(fcData[r][isTodayIdx]) === 1) {
              todayRowIdx = r;
              break;
            }
          }
        }

        // 2) Fallback: Datum == heute
        if (todayRowIdx === -1) {
          const dateIdx = fcHeadersLc.indexOf('datum') > -1
            ? fcHeadersLc.indexOf('datum')
            : fcHeadersLc.indexOf('date');

          if (dateIdx !== -1) {
            for (let r = 1; r < fcData.length; r++) {
              const rawDate = fcData[r][dateIdx];
              const dObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
              if (dObj instanceof Date && !isNaN(dObj.getTime())) {
                const rowKey = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
                if (rowKey === todayKey) {
                  todayRowIdx = r;
                  break;
                }
              }
            }
          }
        }

        if (todayRowIdx !== -1) {
          let teRaw = fcData[todayRowIdx][teIdx];
          let teNum = Number.NaN;

          if (typeof teRaw === 'number') {
            teNum = teRaw;
          } else {
            teNum = parseGermanFloat(String(teRaw).replace('%', ''));
          }

          if (Number.isFinite(teNum)) {
            if (Math.abs(teNum) <= 1.0 && teNum !== 0) teNum = teNum * 100;
            teBalanceHeute = teNum.toFixed(1).replace('.', ',') + "%";
          }
        }
      }
    }
  } catch (e) {
    teBalanceHeute = null;
  }

  // Datum und Wochentag von heute
  let heuteDateObj = heute['date'];
  if (!(heuteDateObj instanceof Date)) heuteDateObj = new Date(heuteDateObj);
  const heuteDatum = Utilities.formatDate(heuteDateObj, tz, "yyyy-MM-dd");
  let heuteTag = heute['date.2'];
  if (!heuteTag || heuteTag === "") heuteTag = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"][heuteDateObj.getDay()];

  // Wissensbasis KK_WISSEN: aktive Inhalte (Spalte B = 'x', Spalte C = Text); null = Blatt fehlt
  let wissen = null;
  const wSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KK_WISSEN');
  if (wSheet) {
    wissen = [];
    const w = wSheet.getDataRange().getValues();
    for (let i = 1; i < w.length; i++) {
      if (String(w[i][1]).toLowerCase() === 'x' && w[i][2]) wissen.push(String(w[i][2]));
    }
  }

  // Treffsicherheit (letzte Wochen-Auswertung) und Planung (Ziele, Saisonrahmen) aus den KI-Extras
  let treffer = null;
  let planung = null;
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CKX_SHEET);
    if (sh && sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        if (rows[i][1] === 'treffer_woche' && rows[i][2]) { treffer = { datum: rows[i][0], text: String(rows[i][2]) }; break; }
      }
      const heuteKey = Utilities.formatDate(new Date(), tz || 'Europe/Berlin', 'yyyy-MM-dd');
      const ziele = [];
      let saison = null;
      rows.forEach(function (r) {
        if (r[1] === 'ziel' && r[0] >= heuteKey) ziele.push({ datum: r[0], text: String(r[2]) });
        if (r[1] === 'saison' && (!saison || r[0] > saison[0])) saison = r;
      });
      let wochen = [];
      if (saison && saison[3]) { try { wochen = JSON.parse(saison[3]).wochen || []; } catch (e) {} }
      planung = { heute: heuteKey, ziele: ziele, wochen: wochen };
    }
  } catch (e) {}

  // JSON-sicher (Date → ISO-Text), wie es der Mac per supervisorJob bekommt
  return JSON.parse(JSON.stringify({
    v: 1,
    heute: heute,
    heute_datum: heuteDatum,
    heute_tag: heuteTag,
    zukunft_plan: (data && data.zukunft && data.zukunft.plan) || [],
    baseline: (data && data.baseline) || {},
    history: (data && data.history) || '',
    monotony_varianz: data ? data.monotony_varianz : 0,
    rtp_status: data ? data.rtp_status : '',
    fitnessScores: fitnessScores,
    nutritionScore: nutritionScore,
    gesamtScoreNum: gesamtScoreNum,
    gesamtScoreAmpel: gesamtScoreAmpel,
    zoneFactors: zoneFactors,
    elevFactors: elevFactors,
    weekConfigCSV: weekConfigCSV,
    isActivityDone: !!isActivityDone,
    wetterDaten: wetterDaten,
    teBalanceHeute: teBalanceHeute,
    wissen: wissen,
    treffer: treffer,
    planung: planung
  }));
}

function ckSup_enqueue_(promptData, ctx, runId) {
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();
  const engine = String(props.getProperty('CK_SUP_ENGINE_NEXT') || 'local') === 'cloud' ? 'cloud' : 'local';
  props.deleteProperty('CK_SUP_ENGINE_NEXT');
  const job = {
    id: String(runId) + '-' + Date.now(),
    runId: String(runId),
    createdAt: ckDash_now_(),
    engine: engine,
    data: promptData,
    ctx: ctx
  };
  const s = JSON.stringify(job);
  const size = 90000;
  const parts = {};
  const n = Math.ceil(s.length / size);
  for (let i = 0; i < n; i++) parts['ckSup_job_' + i] = s.substr(i * size, size);
  parts['ckSup_job_n'] = String(n);
  parts['ckSup_job_id'] = job.id;
  cache.putAll(parts, 21600); // 6 Stunden
}

function ckSup_readJob_() {
  const cache = CacheService.getScriptCache();
  const n = parseInt(cache.get('ckSup_job_n') || '0', 10);
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push('ckSup_job_' + i);
  const got = cache.getAll(keys);
  let s = '';
  for (let i = 0; i < n; i++) {
    if (got['ckSup_job_' + i] == null) return null;
    s += got['ckSup_job_' + i];
  }
  try { return JSON.parse(s); } catch (e) { return null; }
}

function ckSup_clearJob_() {
  const cache = CacheService.getScriptCache();
  const n = parseInt(cache.get('ckSup_job_n') || '0', 10);
  const keys = ['ckSup_job_n', 'ckSup_job_id'];
  for (let i = 0; i < n; i++) keys.push('ckSup_job_' + i);
  cache.removeAll(keys);
}

function ckSup_getJob_() {
  try {
    const job = ckSup_readJob_();
    if (!job) return ckDash_json_({ ok: true, timestamp: ckDash_now_(), engine: ckSup_engine_(), job: null });
    return ckDash_json_({
      ok: true, timestamp: ckDash_now_(), engine: ckSup_engine_(),
      job: { id: job.id, createdAt: job.createdAt, engine: job.engine || 'local', data: job.data || null }
    });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

function ckSup_save_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let runId = 'local';
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const job = ckSup_readJob_();
    if (!job) throw new Error('Kein offener Supervisor-Auftrag (abgelaufen oder schon gespeichert).');
    if (String(b.id || '') !== job.id) throw new Error('Auftrags-ID passt nicht (neuerer Auftrag vorhanden).');
    const content = String(b.content || '').trim();
    if (!content) throw new Error('Leeres Ergebnis.');
    runId = job.runId || runId;
    ckSup_clearJob_();
    logToSheet('INFO', '[Supervisor] Lokales Ergebnis empfangen (' + String(b.model || 'coach-kira').slice(0, 40) + ').');
    ckSup_finish_(ss, content, job.ctx, runId);
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), id: job.id });
  } catch (err) {
    try { hb_(ss, runId, "END_ERROR", String(err && err.message ? err.message : err)); } catch (x) {}
    try { logToSheet('ERROR', '[Supervisor] Lokal speichern fehlgeschlagen: ' + (err && err.message ? err.message : err)); } catch (x) {}
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** Speicherteil des Supervisors – unverändert aus runKiraGeminiSupervisor übernommen. */
function ckSup_finish_(ss, apiResponse, ctx, HB_RUN_ID) {
  const fitnessScores = ctx.fitnessScores;
  const nutritionScore = ctx.nutritionScore;
  const scoreResult = ctx.scoreResult;
  const subScores = ctx.subScores;
  const teResult = ctx.teResult;

  // Original-Load nicht der KI überlassen: fest aus der timeline übernehmen
  try { apiResponse = ckSup_fixOriginalLoads_(ss, apiResponse); } catch (e) { logToSheet('WARN', '[Supervisor] Original-Load nicht korrigiert: ' + e.message); }

  // AI_REPORT_PLAN wieder aktuell halten (wird von Kalender, „KI-Plan anwenden“ und Plan-Daten gelesen)
  try { ckSup_writeReportPlan_(ss, apiResponse); } catch (e) { logToSheet('WARN', '[Supervisor] AI_REPORT_PLAN nicht geschrieben: ' + e.message); }



    hb_(ss, HB_RUN_ID, "AI_OK", "OpenAI Antwort da");
    hb_(ss, HB_RUN_ID, "WRITE", "writeOutputToSheets startet");

    // ------------------------------------------------------------
    // SCHRITT 4: SPEICHERN (Hier schreibt die KI den Text in F11!)
    // ------------------------------------------------------------
    writeOutputToSheets(
        ss, 
        apiResponse, 
        fitnessScores, 
        nutritionScore, 
        scoreResult.num, 
        scoreResult.ampel, 
        subScores.recoveryScore, 
        subScores.recoveryAmpel, 
        subScores.trainingScore, 
        subScores.trainingAmpel
    );

    // +++ ANZEIGE-FIX (NUR FORMAT & AMPEL) +++
    // Wir lassen den Text in F11 (den die KI gerade geschrieben hat) in Ruhe!
    try {
       const statusSheet = ss.getSheetByName('AI_REPORT_STATUS');
       if (statusSheet && teResult.val > 0) {
           // Wir fixieren nur die Zahl als Text-Format (gegen "0,106...")
           statusSheet.getRange("D11").setNumberFormat("@").setValue(teResult.text);
           
           // Wir fixieren Score & Ampel (falls writeOutput geschlampt hat)
           let ampel = (teResult.val >= 0.2) ? "LILA" : (teResult.val >= 0.1) ? "GRÜN" : "ROT";
           let score = (teResult.val >= 0.2) ? 100 : (teResult.val >= 0.1) ? 80 : 25;
           
           statusSheet.getRange("C11").setValue(ampel);
           statusSheet.getRange("E11").setValue(score);
           
           // WICHTIG: KEIN setValue für F11 hier!
       }
    } catch(e) { console.log("Cell Fix Fehler: " + e.message); }

    logToSheet('INFO', `Berichte & Scores gespeichert.`);

    // ------------------------------------------------------------
    // SCHRITT 5: NACHARBEIT
    // ------------------------------------------------------------
    try {
       const uhrzeit = new Date().toLocaleTimeString('de-DE', {hour: '2-digit', minute:'2-digit'});
       const appUrl = "https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec"; 
       const configSheet = ss.getSheetByName("KK_CONFIG");
       if (configSheet) configSheet.getRange("B9").setValue(appUrl);
       if (typeof sendTelegram === 'function') {
         try {
            sendTelegram(`🫡 Commander, Analyse fertig (${uhrzeit})!\n📱 Dashboard: ${appUrl}`);
         } catch(e) { Utilities.sleep(1000); sendTelegram("🫡 Analyse fertig (Retry)"); }
       }
    } catch(e) {}

    // ------------------------------------------------------------
    // SCHRITT 6: NACHARBEIT (ASYNC)
    // ------------------------------------------------------------
    // Diese Schritte können (je nach Datenmenge) viel Zeit kosten. Damit der Supervisor
    // schneller fertig ist, laufen sie asynchron in einem Trigger.
    schedulePostWork_();

    Logger.log("✅ Supervisor-Prozess komplett abgeschlossen.");
    logToSheet('INFO', `✅ KI-Bewertung V137 erfolgreich.`);
    hb_(ss, HB_RUN_ID, "END_OK", "Supervisor fertig");

}
// ============================================================================
// Wartung 19.10.2026: Doppelstarts verhindern, AI_REPORT_PLAN schreiben, Diagnose
// ============================================================================
/** true = dieser Lauf darf starten. Läuft schon einer (< 6 min), wird ein Nachlauf vorgemerkt. */
function ckSup_begin_(runId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return true; // im Zweifel laufen lassen
  try {
    const props = PropertiesService.getScriptProperties();
    const running = Number(props.getProperty('CK_SUP_RUNNING') || 0);
    if (running && Date.now() - running < 6 * 60 * 1000) {
      props.setProperty('CK_SUP_RERUN', '1');
      try { logToSheet('INFO', '[Supervisor] Läuft bereits – Nachlauf vorgemerkt (' + runId + ').'); } catch (x) {}
      return false;
    }
    props.setProperty('CK_SUP_RUNNING', String(Date.now()));
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Lauf beendet: Sperre lösen; war ein Nachlauf vorgemerkt, einmal neu starten (frische Daten). */
function ckSup_end_() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('CK_SUP_RUNNING');
    if (props.getProperty('CK_SUP_RERUN')) {
      props.deleteProperty('CK_SUP_RERUN');
      ScriptApp.getProjectTriggers().forEach(function (t) {
        if (t.getHandlerFunction() === 'runKiraGeminiSupervisor' && t.getEventType() === ScriptApp.EventType.CLOCK) {
          try { ScriptApp.deleteTrigger(t); } catch (x) {}
        }
      });
      ScriptApp.newTrigger('runKiraGeminiSupervisor').timeBased().after(5000).create();
      try { logToSheet('INFO', '[Supervisor] Nachlauf gestartet.'); } catch (x) {}
    }
  } catch (e) {
    try { logToSheet('WARN', '[Supervisor] ckSup_end_: ' + e.message); } catch (x) {}
  }
}

/** Schreibt empfohlener_plan aus der KI-Antwort nach AI_REPORT_PLAN (gleiche Spalten wie früher). */
function ckSup_writeReportPlan_(ss, apiResponse) {
  let data = apiResponse;
  if (typeof data === 'string') {
    const s = data.replace(/`{3}(?:json)?/g, '').trim();
    try { data = JSON.parse(s); } catch (e) { data = JSON.parse(s.substring(s.indexOf('{'), s.lastIndexOf('}') + 1)); }
  }
  const plan = (data && Array.isArray(data.empfohlener_plan)) ? data.empfohlener_plan : [];
  if (!plan.length) {
    // Kein KI-Plan mehr (Plan-Briefing übernimmt): altes Blatt leeren, damit nichts Veraltetes gelesen wird
    const old = ss.getSheetByName(OUTPUT_PLAN_SHEET);
    if (old && old.getLastRow() > 1) old.getRange(2, 1, old.getLastRow() - 1, Math.max(9, old.getLastColumn())).clearContent();
    return;
  }
  const toDate = function (v) {
    const t = String(v || '').trim();
    let m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
    m = t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0);
    return t;
  };
  const header = ['Datum', 'Tag', 'Original Load (ESS)', 'Empfohlener Load (ESS)', 'Empfohlene Zone (KI)',
    'Beispiel-Training 1 (KI)', 'Beispiel-Training 2 (KI)', 'Prognostizierte Auswirkung (KI)', 'Wetterempfehlung'];
  const rows = plan.map(function (d) {
    return [toDate(d.datum), d.tag || '', Number(d.original_load_ess) || 0, Number(d.empfohlener_load_ess) || 0,
      d.empfohlene_zone || '', d.beispiel_training_1 || '', d.beispiel_training_2 || '',
      d.prognostizierte_auswirkung || '', d.wetterempfehlung || ''];
  });
  let sh = ss.getSheetByName(OUTPUT_PLAN_SHEET);
  if (!sh) sh = ss.insertSheet(OUTPUT_PLAN_SHEET);
  const lr = sh.getLastRow();
  if (lr > 1) sh.getRange(2, 1, lr - 1, Math.max(9, sh.getLastColumn())).clearContent();
  sh.getRange(1, 1, 1, 9).setValues([header]);
  sh.getRange(2, 1, rows.length, 9).setValues(rows);
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-MM-dd');
  logToSheet('INFO', '[Supervisor] AI_REPORT_PLAN aktualisiert (' + rows.length + ' Tage ab ' + String(plan[0].datum) + ').');
}

/** Diagnose für Wartung: welche Funktionen sind aktiv, Log-Zustand, Trigger. Schreibt einen Testeintrag. */
function ckDash_diag_() {
  const out = { ok: true, timestamp: ckDash_now_() };
  try {
    const props = PropertiesService.getScriptProperties();
    out.logToSheetHead = String(logToSheet).slice(0, 220);
    out.hbHead = String(hb_).slice(0, 120);
    out.logSheetName = LOG_SHEET_NAME;
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    out.logSheet = sh ? { lastRow: sh.getLastRow(), maxRows: sh.getMaxRows(), lastCol: sh.getLastColumn() } : null;
    const before = sh ? sh.getLastRow() : 0;
    logToSheet('DEBUG', '[Diag] Testeintrag');
    const sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    out.testLogWritten = sh2 ? (sh2.getLastRow() !== before || String(sh2.getRange(sh2.getLastRow(), 3).getValue()).indexOf('[Diag]') >= 0) : false;
    out.lastLogError = props.getProperty('CK_LOG_LAST_ERROR');
    out.supRunning = props.getProperty('CK_SUP_RUNNING');
    out.supRerun = props.getProperty('CK_SUP_RERUN');
    out.supervisorEngine = props.getProperty('SUPERVISOR_ENGINE');
    out.triggers = ScriptApp.getProjectTriggers().map(function (t) {
      return { fn: t.getHandlerFunction(), type: String(t.getEventType()), source: String(t.getTriggerSource()) };
    });
  } catch (e) {
    out.ok = false;
    out.error = String(e && e.message ? e.message : e);
  }
  return ckDash_json_(out);
}
/** Setzt original_load_ess im KI-Plan auf den geplanten Wert aus der timeline (coachE_ESS_day). */
function ckSup_fixOriginalLoads_(ss, apiResponse) {
  let data = apiResponse;
  const wasString = (typeof data === 'string');
  if (wasString) {
    const s = data.replace(/`{3}(?:json)?/g, '').trim();
    try { data = JSON.parse(s); } catch (e) { data = JSON.parse(s.substring(s.indexOf('{'), s.lastIndexOf('}') + 1)); }
  }
  const plan = (data && Array.isArray(data.empfohlener_plan)) ? data.empfohlener_plan : [];
  if (!plan.length) return apiResponse;

  const sh = ss.getSheetByName('timeline');
  if (!sh) return apiResponse;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const dIdx = head.indexOf('date');
  const eIdx = head.indexOf('coachE_ESS_day');
  const fIdx = head.indexOf('fix');
  const zIdx = head.indexOf('Zone');
  const sIdx = head.indexOf('Sport_x');
  if (dIdx < 0 || eIdx < 0) return apiResponse;

  const n = sh.getLastRow() - 1;
  const dates = sh.getRange(2, dIdx + 1, n, 1).getValues();
  const ess = sh.getRange(2, eIdx + 1, n, 1).getValues();
  const fix = fIdx >= 0 ? sh.getRange(2, fIdx + 1, n, 1).getValues() : null;
  const zon = zIdx >= 0 ? sh.getRange(2, zIdx + 1, n, 1).getDisplayValues() : null;
  const spo = sIdx >= 0 ? sh.getRange(2, sIdx + 1, n, 1).getDisplayValues() : null;
  const locked = {};

  const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
  const map = {};
  for (let i = 0; i < n; i++) {
    const d = dates[i][0];
    if (!(d instanceof Date)) continue;
    const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    map[k] = Number(String(ess[i][0]).replace(',', '.')) || 0;
    if (fix && String(fix[i][0]).trim().toLowerCase() === 'x') {
      locked[k] = { zone: zon ? String(zon[i][0] || '') : '', sport: spo ? String(spo[i][0] || '') : '' };
    }
  }
  let changed = 0;
  plan.forEach(function (day) {
    const key = String(day.datum || '').slice(0, 10);
    if (key in map && Number(day.original_load_ess) !== map[key]) { day.original_load_ess = map[key]; changed++; }
    // Gesperrte Tage (fix = x) bleiben wie geplant – die KI darf sie nicht verändern
    if (locked[key]) {
      const lk = locked[key];
      if (Number(day.empfohlener_load_ess) !== map[key]) { day.empfohlener_load_ess = map[key]; changed++; }
      day.empfohlene_zone = map[key] > 0 ? (lk.zone || day.empfohlene_zone || '') : 'Ruhetag (Fixed)';
      if (map[key] <= 0) { day.beispiel_training_1 = 'Geplanter Ruhetag'; day.beispiel_training_2 = 'Mobility / Spaziergang'; }
      else if (lk.sport) day.beispiel_training_1 = lk.sport + ' (fest geplant)';
    }
  });

  if (changed) logToSheet('INFO', '[Supervisor] Original-Load an ' + changed + ' Tag(en) aus der timeline übernommen.');
  return wasString ? JSON.stringify(data) : data;
}

// ============================================================
// Coach Kira – KI-Extras (Treffsicherheit, Selbsttest, Plausibilität,
// Ernährung, Alltag-Check) · Speicher: Blatt AI_KIRA_EXTRAS
// ============================================================
const CKX_SHEET = 'AI_KIRA_EXTRAS';
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis', 'garmin_api', 'garmin_tag'];


function ckX_sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CKX_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CKX_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['Datum', 'Typ', 'Text', 'Daten (JSON)', 'Quelle', 'Gespeichert']]);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

/** POST kind=saveExtra · {typ, datum:'YYYY-MM-DD', text, data, quelle} · ersetzt Eintrag mit gleichem Typ+Datum */
function ckX_save_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const typ = String(b.typ || '');
    if (CKX_TYPES.indexOf(typ) < 0) throw new Error('Unbekannter Typ: ' + typ);
    const datum = String(b.datum || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(datum)) throw new Error('datum fehlt oder ist nicht YYYY-MM-DD.');
    const text = String(b.text || '').slice(0, 6000);
    const data = b.data === undefined ? '' : JSON.stringify(b.data).slice(0, 40000);
    const quelle = String(b.quelle || 'lokal').replace(/[^\w:.\- ]/g, '').slice(0, 60);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = ckX_sheet_();
      const row = [datum, typ, text, data, quelle, ckDash_now_()];
      const n = sh.getLastRow() - 1;
      // Eintrag löschen (z. B. ein Ziel entfernen): {typ, datum, loeschen:true}
      if (b.loeschen === true) {
        if (n > 0) {
          const keys0 = sh.getRange(2, 1, n, 2).getDisplayValues();
          for (let i = n - 1; i >= 0; i--) {
            if (keys0[i][0] === datum && keys0[i][1] === typ) sh.deleteRow(i + 2);
          }
        }
        return ckDash_json_({ ok: true, timestamp: ckDash_now_(), deleted: true });
      }

      if (n > 0) {
        const keys = sh.getRange(2, 1, n, 2).getDisplayValues();
        for (let i = n - 1; i >= 0; i--) {
          if (keys[i][0] === datum && keys[i][1] === typ) {
            sh.getRange(i + 2, 1, 1, 6).setValues([row]);
            return ckDash_json_({ ok: true, timestamp: ckDash_now_(), updated: true });
          }
        }
      }
      sh.appendRow(row);
    } finally {
      lock.releaseLock();
    }
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), updated: false });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** GET mode=extras&typ=a,b&n=30 · neueste zuerst */
function ckX_get_(p) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CKX_SHEET);
    if (!sh || sh.getLastRow() < 2) return ckDash_json_({ ok: true, timestamp: ckDash_now_(), items: [] });
    const types = String(p.typ || '').split(',').map(function (s) { return s.trim(); }).filter(String);
    const n = Math.max(1, Math.min(1000, Number(p.n) || 30));
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues();
    const items = [];
    for (let i = rows.length - 1; i >= 0 && items.length < n; i--) {
      const r = rows[i];
      if (types.length && types.indexOf(r[1]) < 0) continue;
      let data = null;
      if (r[3]) { try { data = JSON.parse(r[3]); } catch (x) { data = null; } }
      items.push({ datum: r[0], typ: r[1], text: r[2], data: data, quelle: r[4], gespeichert: r[5] });
    }
    items.sort(function (a, b) { return a.datum < b.datum ? 1 : a.datum > b.datum ? -1 : 0; });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), items: items });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/**
 * GET mode=calendarDays&days=14 · nur Belegung, KEINE Termintitel
 * (die Web-App ist öffentlich erreichbar).
 */
function ckX_calendar_(p) {
  try {
    const days = Math.max(1, Math.min(21, Number(p.days) || 14));
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const skip = /coach\s*kira|feiertag|holiday|geburtstag|birthday|wochennummer|week numbers/i;
    const reise = /reise|flug|flight|hotel|dienstreise|messe|zug nach|bahn/i;
    const frei = /urlaub|ferien|frei\b|vacation/i;
    const krank = /krank|arzt|zahnarzt|physio|op\b/i;
    // optional: nur bestimmte Kalender (kal=Name|Name|…, exakter Name ohne Groß/Klein) und ganztag=0 = ganztägige Termine ignorieren
    const nurKal = String(p.kal || '').split('|').map(function (x) { return x.trim().toLowerCase(); }).filter(String);
    const ohneGanztag = String(p.ganztag) === '0';
    const cals = CalendarApp.getAllCalendars().filter(function (c) {
      const n = String(c.getName() || '').trim().toLowerCase();
      return nurKal.length ? nurKal.indexOf(n) >= 0 : !skip.test(c.getName());
    });
    const start = new Date(); start.setHours(0, 0, 0, 0);
    const out = [];
    for (let d = 0; d < days; d++) {
      // Kalendertage statt fester 24 h (sonst doppelter Tag bei der Zeitumstellung)
      const a = new Date(start.getFullYear(), start.getMonth(), start.getDate() + d, 0, 0, 0);
      const b = new Date(start.getFullYear(), start.getMonth(), start.getDate() + d + 1, 0, 0, 0);
      const dayStart = new Date(a.getFullYear(), a.getMonth(), a.getDate(), 6, 0, 0);
      const dayEnd = new Date(a.getFullYear(), a.getMonth(), a.getDate(), 22, 0, 0);

      let allDay = 0, evening = false, morning = false, isReise = false, isFrei = false, isKrank = false, count = 0;
      const iv = [];
      const seen = {};
      let first = null, last = null;
      cals.forEach(function (c) {
        c.getEvents(a, b).forEach(function (ev) {
          if (/^\\[coach kira\\]/i.test(ev.getTitle())) return;
          if (ohneGanztag && ev.isAllDayEvent()) return;
          const t = ev.getTitle() || '';
          // gleicher Termin in mehreren Kalendern nur einmal zählen
          const key = t + '|' + ev.getStartTime().getTime() + '|' + ev.getEndTime().getTime();
          if (seen[key]) return;
          seen[key] = true;
          count++;
          if (reise.test(t)) isReise = true;
          if (frei.test(t)) isFrei = true;
          if (krank.test(t)) isKrank = true;
          if (ev.isAllDayEvent()) { allDay++; return; }
          const s = new Date(Math.max(ev.getStartTime().getTime(), dayStart.getTime()));
          const e2 = new Date(Math.min(ev.getEndTime().getTime(), dayEnd.getTime()));
          if (e2 > s) iv.push([s.getTime(), e2.getTime()]);
          if (ev.getEndTime().getHours() >= 18 || ev.getStartTime().getHours() >= 17) evening = true;
          if (ev.getStartTime().getHours() < 8) morning = true;
          if (!first || ev.getStartTime() < first) first = ev.getStartTime();
          if (!last || ev.getEndTime() > last) last = ev.getEndTime();
        });
      });
      // Überschneidungen zusammenfassen
      iv.sort(function (x, y) { return x[0] - y[0]; });
      let busyMin = 0, curS = null, curE = null;
      iv.forEach(function (x) {
        if (curE === null || x[0] > curE) { if (curE !== null) busyMin += (curE - curS) / 60000; curS = x[0]; curE = x[1]; }
        else if (x[1] > curE) curE = x[1];
      });
      if (curE !== null) busyMin += (curE - curS) / 60000;
      out.push({

        date: Utilities.formatDate(a, tz, 'yyyy-MM-dd'),
        termine: count,
        ganztaegig: allDay,
        belegt_h: Math.round(Math.min(busyMin, 16 * 60) / 6) / 10,
        frueh: morning,
        abends: evening,
        von: first ? Utilities.formatDate(first, tz, 'HH:mm') : '',
        bis: last ? Utilities.formatDate(last, tz, 'HH:mm') : '',
        reise: isReise,
        frei: isFrei,
        gesundheit: isKrank
      });
    }
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), calendars: cals.length, days: out,
      filter: { kalender: nurKal.length ? cals.map(function (c) { return c.getName(); }) : null, ganztaegig: !ohneGanztag } });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** GET mode=physBaseline · Gewicht, Max-HF, Laktatschwelle, FTP … aus phys_baseline */
function ckX_phys_() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('phys_baseline');
    if (!sh) throw new Error('Blatt phys_baseline fehlt.');
    const v = sh.getRange(1, 1, 2, sh.getLastColumn()).getValues();
    const o = {};
    v[0].forEach(function (h, i) { if (String(h).trim()) o[String(h).trim()] = v[1][i]; });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), baseline: o });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

// KI-frei (Stufe 3): ckX_trefferPromptBlock_ entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

// KI-frei (Stufe 3): ckRegel_monotonie_ entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

// KI-frei (Stufe 3): ckRegel_monotonieBlock_ entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

/** GET mode=loadFactors · KK_LOAD_FACTORS (ESS pro Stunde je Sport/Zone) und KK_ELEV_FACTORS (nur lesen) */
function ckX_loadFactors_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const read = function (name) {
      const sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return [];
      const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
      const h = v[0].map(function (x) { return String(x).trim(); });
      return v.slice(1).filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
        const o = {}; h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o;
      });
    };
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), load: read('KK_LOAD_FACTORS'), elev: read('KK_ELEV_FACTORS') });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

// KI-frei (Stufe 3): ckX_planungsBlock_ entfernt – Prompt-Bau liegt in kira_prompts (Cockpit/Mac).

// ============================================================
// Coach Kira – Garmin km/hm (Mac-Dienst kira_garmin.py, Garmin Personal API)
// POST kind=saveDayMetrics · {rows:[{datum:'YYYY-MM-DD', km, hm}], modus:'leer'|'alle', probe:true|false}
// Schreibt distance_km_day / elev_m_day in 'timeline' (Quelle) und 'KK_TIMELINE' (Spiegel).
// modus 'leer' (Standard): nur leere Zellen füllen · 'alle': vorhandene Werte überschreiben
// probe: nichts schreiben, nur zählen
// ============================================================
function ckGA_saveDays_(e) {
  const lock = LockService.getScriptLock();
  let gesperrt = false;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const rows = Array.isArray(b.rows) ? b.rows : [];
    if (!rows.length) return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'rows fehlt' });
    if (rows.length > 400) return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'max 400 Tage je Aufruf' });
    const modus = b.modus === 'alle' ? 'alle' : 'leer';
    const probe = b.probe === true;
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const heute = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    const tage = {};
    let verworfen = 0;
    rows.forEach(function (r) {
      const datum = String((r && r.datum) || '');
      const km = Number(r && r.km);
      const hm = Number(r && r.hm);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(datum) || datum > heute ||
          !isFinite(km) || km < 0 || km > 1000 || !isFinite(hm) || hm < 0 || hm > 20000) {
        verworfen++;
        return;
      }
      tage[datum] = { km: Math.round(km * 100) / 100, hm: Math.round(hm) };
    });

    lock.waitLock(20000);
    gesperrt = true;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leer = function (x) { return x === '' || x === null; };
    const blaetter = {};

    ['timeline', 'KK_TIMELINE'].forEach(function (name) {
      const sh = ss.getSheetByName(name);
      if (!sh) throw new Error("Blatt '" + name + "' nicht gefunden.");
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
      const cD = head.indexOf('date'), cK = head.indexOf('distance_km_day'), cH = head.indexOf('elev_m_day');
      [['date', cD], ['distance_km_day', cK], ['elev_m_day', cH]].forEach(function (p) {
        if (p[1] < 0) throw new Error("Spalte '" + p[0] + "' fehlt in Blatt '" + name + "'.");
      });
      const n = sh.getLastRow() - 1;
      const st = { geschrieben: 0, unveraendert: 0, behalten: 0, ohne_zeile: 0 };
      if (n < 1) { st.ohne_zeile = Object.keys(tage).length; blaetter[name] = st; return; }

      const rK = sh.getRange(2, cK + 1, n, 1), rH = sh.getRange(2, cH + 1, n, 1);
      const fK = rK.getFormulas(), fH = rH.getFormulas();
      for (let i = 0; i < n; i++) {
        if (fK[i][0]) throw new Error("Formel in Spalte 'distance_km_day' von Blatt '" + name + "' (Zeile " + (i + 2) + ').');
        if (fH[i][0]) throw new Error("Formel in Spalte 'elev_m_day' von Blatt '" + name + "' (Zeile " + (i + 2) + ').');
      }
      const dates = sh.getRange(2, cD + 1, n, 1).getValues();
      const vK = rK.getValues(), vH = rH.getValues();
      const zeilen = {};
      for (let i = 0; i < n; i++) {
        const d = dates[i][0];
        const key = d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || '').slice(0, 10);
        if (key) (zeilen[key] = zeilen[key] || []).push(i);
      }

      Object.keys(tage).forEach(function (datum) {
        const idx = zeilen[datum];
        if (!idx) { st.ohne_zeile++; return; }
        const w = tage[datum];
        idx.forEach(function (i) {
          const aK = vK[i][0], aH = vH[i][0];
          if (aK === w.km && aH === w.hm) { st.unveraendert++; return; }
          if (modus !== 'alle' && (!leer(aK) || !leer(aH))) { st.behalten++; return; }
          vK[i][0] = w.km;
          vH[i][0] = w.hm;
          st.geschrieben++;
        });
      });

      if (!probe && st.geschrieben > 0) {
        rK.setValues(vK);
        rH.setValues(vH);
      }
      blaetter[name] = st;
    });

    return ckDash_json_({
      ok: true, timestamp: ckDash_now_(), modus: modus, probe: probe,
      tage: Object.keys(tage).length, verworfen: verworfen, blaetter: blaetter,
    });
  } catch (err) {
    try { logToSheet('ERROR', '[ckGA_saveDays_] ' + err); } catch (e2) {}
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  } finally {
    if (gesperrt) lock.releaseLock();
  }
}

// ============================================================================
// KI-frei (Stufe 3, 2026-10-26): Telegram-Fragen – Warteschlange für den Mac
//   Bot → ckTg_enqueue_ · Mac holt mode=telegramJobs · antwortet POST kind=telegramAnswer
// ============================================================================
const CK_TG_QUEUE_KEY = 'CK_TG_QUEUE';
const CK_TG_VERLAUF_KEY = 'CK_TG_VERLAUF';

function ckTg_read_(key) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(key) || '[]'); } catch (e) { return []; }
}

function ckTg_send_(chatId, text) {
  // ohne Markdown, damit Sonderzeichen in Kiras Antwort nichts kaputt machen
  const cfg = getTelegramConfig_();
  if (!cfg.token) { logToSheet('ERROR', '[Telegram] TELEGRAM_BOT_TOKEN fehlt in ScriptProperties'); return; }
  UrlFetchApp.fetch('https://api.telegram.org/bot' + cfg.token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chatId, text: String(text).slice(0, 4000) })
  });
}

function ckTg_enqueue_(chatId, text, user) {
  const cfg = getTelegramConfig_();
  if (cfg.chatId && String(chatId) !== String(cfg.chatId)) return; // nur Marcs Chat
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const q = ckTg_read_(CK_TG_QUEUE_KEY);
    q.push({ id: 'tg-' + Date.now(), chatId: String(chatId), frage: String(text).slice(0, 2000), user: String(user || ''), ts: ckDash_now_() });
    PropertiesService.getScriptProperties().setProperty(CK_TG_QUEUE_KEY, JSON.stringify(q.slice(-10)));
  } finally {
    lock.releaseLock();
  }
  ckTg_send_(chatId, 'Kira denkt nach – die Antwort kommt vom Mac (lokal), meist in 1–2 Minuten.');
}

function ckTg_jobs_() {
  return ckDash_json_({
    ok: true, timestamp: ckDash_now_(),
    jobs: ckTg_read_(CK_TG_QUEUE_KEY).map(function (j) { return { id: j.id, frage: j.frage, ts: j.ts }; }),
    verlauf: ckTg_read_(CK_TG_VERLAUF_KEY)
  });
}

function ckTg_answer_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const id = String(b.id || '');
    const text = String(b.text || '').trim();
    if (!id || !text) return ckDash_json_({ ok: false, error: 'id oder text fehlt.' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    let job = null;
    try {
      const q = ckTg_read_(CK_TG_QUEUE_KEY);
      job = q.filter(function (j) { return j.id === id; })[0] || null;
      if (!job) return ckDash_json_({ ok: false, error: 'Frage nicht (mehr) in der Warteschlange.' });
      PropertiesService.getScriptProperties().setProperty(CK_TG_QUEUE_KEY, JSON.stringify(q.filter(function (j) { return j.id !== id; })));
      const v = ckTg_read_(CK_TG_VERLAUF_KEY);
      v.push({ frage: job.frage.slice(0, 500), antwort: text.slice(0, 1500) });
      PropertiesService.getScriptProperties().setProperty(CK_TG_VERLAUF_KEY, JSON.stringify(v.slice(-3)));
    } finally {
      lock.releaseLock();
    }
    ckTg_send_(job.chatId, text);
    const m = String(b.model || 'coach-kira').replace(/[^\w:.\-]/g, '').slice(0, 40);
    const eng = String(b.engine || 'local') === 'cloud' ? 'Cloud' : 'lokal';
    logChatMessage('Marc (Telegram)', job.frage);
    logChatMessage('Kira (' + eng + ', ' + m + ', Telegram)', text);
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), id: id });
  } catch (err) {
    return ckDash_json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}


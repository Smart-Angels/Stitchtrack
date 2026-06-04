// ════════════════════════════════════════════════════════════════
//  StitchTrack v7 — Google Apps Script Backend
//  Paste this entire file into your Apps Script project.
//  Deploy → New deployment → Web app → Execute as: Me → Anyone
// ════════════════════════════════════════════════════════════════

var SS_NAME  = 'StitchTrack Data';   // Name of the Google Sheet
var RAW_SHEET = 'RAW_JSON';          // Sheet that stores the full JSON blob
var LOG_SHEET = 'ChangeLog';         // Audit log of every save

// ── ENTRY POINTS ────────────────────────────────────────────────

/**
 * Handles GET requests.
 * Supports JSONP: ?action=getAll&callback=cb_xxx
 * Also supports plain JSON: ?action=getAll  (returns JSON directly)
 */
function doGet(e) {
  var params  = e.parameter || {};
  var action  = params.action   || 'getAll';
  var callback = params.callback || '';

  var result;
  try {
    if (action === 'ping') {
      result = { ok: true, ts: new Date().toISOString() };
    } else if (action === 'getAll') {
      result = getData();
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  var json = JSON.stringify(result);

  if (callback) {
    // JSONP response — used by the app when loaded from file://
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Plain JSON — used when hosted (CORS handled by Apps Script)
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles POST requests.
 * Body: JSON string { action: 'saveAll', data: { ...DB } }
 * Also accepts GET-style: ?action=saveAll&payload=encodedJSON
 */
function doPost(e) {
  var result;
  try {
    var body = {};

    // Try reading from POST body first
    if (e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (_) {}
    }

    // Fallback: payload in GET params (no-cors fetch sends it this way sometimes)
    if (!body.data && e.parameter && e.parameter.payload) {
      try { body = { action: 'saveAll', data: JSON.parse(e.parameter.payload) }; } catch (_) {}
    }

    var action = body.action || (e.parameter && e.parameter.action) || 'saveAll';

    if (action === 'saveAll' && body.data) {
      result = saveData(body.data);
    } else {
      result = { error: 'No data received or unknown action.' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CORE DATA FUNCTIONS ─────────────────────────────────────────

/**
 * Returns the full DB object from RAW_JSON sheet.
 * If sheet is empty / never saved, returns { savedAt: null }.
 */
function getData() {
  var sheet = getOrCreateSheet(RAW_SHEET);
  var cell  = sheet.getRange('A1').getValue();
  if (!cell) return { savedAt: null };
  try {
    var db = JSON.parse(cell);
    return db;
  } catch (_) {
    return { savedAt: null, parseError: true };
  }
}

/**
 * Saves the full DB object into RAW_JSON sheet as a single JSON string.
 * Also writes human-readable sheets for key collections.
 * Returns { ok: true, savedAt: '...' }
 */
function saveData(db) {
  db.savedAt = new Date().toISOString();

  // ── 1. Write raw JSON blob ────────────────────────────────────
  var rawSheet = getOrCreateSheet(RAW_SHEET);
  rawSheet.getRange('A1').setValue(JSON.stringify(db));

  // ── 2. Write human-readable sheets ───────────────────────────
  try { writeWorkersSheet(db); }   catch(_) {}
  try { writeOrdersSheet(db); }    catch(_) {}
  try { writeProdLogSheet(db); }   catch(_) {}
  try { writeDailyOutputSheet(db); } catch(_) {}
  try { writeLeaveSheet(db); }     catch(_) {}

  // ── 3. Audit log ─────────────────────────────────────────────
  try {
    var logSheet = getOrCreateSheet(LOG_SHEET);
    logSheet.appendRow([
      new Date(),
      'saveAll',
      (db.workers  || []).length + ' workers',
      (db.orders   || []).length + ' orders',
      Object.keys(db.prodLog  || {}).length + ' log days',
      Object.keys(db.dailyOutput || {}).length + ' output days'
    ]);
  } catch(_) {}

  return { ok: true, savedAt: db.savedAt };
}

// ── HUMAN-READABLE SHEET WRITERS ────────────────────────────────

function writeWorkersSheet(db) {
  var sh = getOrCreateSheet('Workers');
  sh.clearContents();
  sh.appendRow(['ID', 'Name', 'Department', 'Salary (Rs.)']);
  (db.workers || []).forEach(function(w) {
    sh.appendRow([w.id, w.name, w.dept, w.salary || '']);
  });
  styleHeader(sh);
}

function writeOrdersSheet(db) {
  var sh = getOrCreateSheet('Orders');
  sh.clearContents();
  sh.appendRow(['Order ID', 'Product Name', 'Client', 'Total Qty', 'Status', 'Incentive/Piece']);
  (db.orders || []).forEach(function(o) {
    sh.appendRow([o.id, o.name, o.client, o.totalQty, o.status, o.incentivePerPiece || 0]);
  });
  styleHeader(sh);
}

function writeProdLogSheet(db) {
  var sh = getOrCreateSheet('Production Log');
  sh.clearContents();
  sh.appendRow(['Date', 'Worker ID', 'Worker Name', 'Slot ID', 'Entry #', 'Order ID', 'Part', 'Qty']);

  var workerMap = {};
  (db.workers || []).forEach(function(w) { workerMap[w.id] = w.name; });

  var log = db.prodLog || {};
  Object.keys(log).sort().forEach(function(date) {
    var dayLog = log[date];
    Object.keys(dayLog).forEach(function(wid) {
      var workerLog = dayLog[wid];
      Object.keys(workerLog).forEach(function(slotId) {
        var entries = workerLog[slotId] || [];
        entries.forEach(function(ent, idx) {
          if (ent.qty || ent.part) {
            sh.appendRow([date, wid, workerMap[wid] || wid, slotId, idx + 1,
                          ent.orderId || '', ent.part || '', parseInt(ent.qty) || 0]);
          }
        });
      });
    });
  });
  styleHeader(sh);
}

function writeDailyOutputSheet(db) {
  var sh = getOrCreateSheet('Daily Output');
  sh.clearContents();
  sh.appendRow(['Date', 'Order ID', 'Actual Completed', 'Notes']);

  var output = db.dailyOutput || {};
  Object.keys(output).sort().forEach(function(date) {
    var dayOut = output[date];
    Object.keys(dayOut).forEach(function(oid) {
      var rec = dayOut[oid];
      sh.appendRow([date, oid, rec.supervisorActual || 0, rec.notes || '']);
    });
  });
  styleHeader(sh);
}

function writeLeaveSheet(db) {
  var sh = getOrCreateSheet('Leave Log');
  sh.clearContents();
  sh.appendRow(['Date', 'Worker ID', 'Worker Name', 'Slot ID', 'On Leave']);

  var workerMap = {};
  (db.workers || []).forEach(function(w) { workerMap[w.id] = w.name; });

  var leave = db.leaveLog || {};
  Object.keys(leave).sort().forEach(function(date) {
    var dayLeave = leave[date];
    Object.keys(dayLeave).forEach(function(wid) {
      var wLeave = dayLeave[wid];
      Object.keys(wLeave).forEach(function(slotId) {
        if (wLeave[slotId]) {
          sh.appendRow([date, wid, workerMap[wid] || wid, slotId, 'Yes']);
        }
      });
    });
  });
  styleHeader(sh);
}

// ── UTILITIES ───────────────────────────────────────────────────

/**
 * Gets or creates a sheet by name in the active spreadsheet.
 * Creates the spreadsheet named SS_NAME if it doesn't exist.
 */
function getOrCreateSheet(name) {
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch(_) {
    // Script not bound to a spreadsheet — find or create by name
    var files = DriveApp.getFilesByName(SS_NAME);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(SS_NAME);
    }
  }

  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Bolds and freezes the first row of a sheet.
 */
function styleHeader(sh) {
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn());
  hdr.setFontWeight('bold');
  hdr.setBackground('#2a2a2a');
  hdr.setFontColor('#f5a623');
  sh.setFrozenRows(1);
}

// ── MANUAL TEST FUNCTION ─────────────────────────────────────────
// Run this in the Apps Script editor to verify the setup works.
function testSetup() {
  var result = getData();
  Logger.log('getData result: ' + JSON.stringify(result).slice(0, 200));

  var testDB = {
    settings: { capacity: 200, otRate: 60 },
    workers:  [{ id: 'W001', name: 'Test Worker', dept: 'stitching', salary: 12000 }],
    orders:   [{ id: 'ORD-001', name: 'Test Shirt', client: 'TestCo', totalQty: 100, status: 'active', incentivePerPiece: 5 }],
    departments: [], parts: [], holidays: [], workingSundays: [], timeSlots: [],
    schedule: {}, prodLog: {}, leaveLog: {}, dailyOutput: {}, materials: {}
  };
  var saveResult = saveData(testDB);
  Logger.log('saveData result: ' + JSON.stringify(saveResult));
  Logger.log('✅ Setup OK — check your Google Sheet for data.');
}

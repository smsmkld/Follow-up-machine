// ============================================================
// FOLLOW UP MACHINE - WEB APP BACKEND
// ============================================================
//
// Add this as a SECOND script file in the same Apps Script project
// (Files + -> Script -> name it "WebApp"). Do not paste it into Code.gs.
// It reuses CONFIG, getSetting, getSheet, todayStr, chainIsRunning etc.
//
// DEPLOY:
//   Deploy -> New deployment -> type "Web app"
//   Execute as:      Me
//   Who has access:  Only myself      <- keep this unless you host remotely
//   Copy the /exec URL, open it on your phone, Add to Home Screen.
//
// ============================================================

/**
 * Serves the single-page app. Called when you open the /exec URL.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Follow Up Machine')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * JSON endpoint, used ONLY when the page is hosted somewhere else
 * (Cloudflare Pages, etc). When Apps Script serves the page itself the
 * front-end calls the functions below directly and never touches this.
 *
 * To use it you must:
 *   1. add a row to Settings:  webAppToken | <a long random string>
 *   2. redeploy with "Who has access: Anyone"
 *   3. put the same string in API_TOKEN at the top of Index.html
 *
 * Read the security note in Index.html before you turn this on.
 */
function doPost(e) {
  const reply = function (obj) {
    return ContentService
      .createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
  };

  // Explicit allow-list. Anything not named here cannot be called remotely,
  // so a leaked token still cannot reach arbitrary script functions.
  const HANDLERS = {
    uiGetStats:      uiGetStats,
    uiGetLeads:      uiGetLeads,
    uiSetLeadStatus: uiSetLeadStatus,
    uiGetSettings:   uiGetSettings,
    uiSaveSetting:   uiSaveSetting,
    uiSearchThreads: uiSearchThreads,
    uiEnrollLead:    uiEnrollLead,
    uiGetLog:        uiGetLog,
    uiGetSequences:  uiGetSequences,
    uiRunChainNow:   uiRunChainNow,
    uiAddSequence:   uiAddSequence,
    uiSaveStep:      uiSaveStep,
    uiDeleteStep:    uiDeleteStep,
    uiDeleteSequence:uiDeleteSequence,
    uiGetLeadRow:    uiGetLeadRow,
    uiSaveLeadRow:   uiSaveLeadRow,
    uiGetVariables:  uiGetVariables,
    uiPreviewStep:   uiPreviewStep,
    uiTestSendStep:  uiTestSendStep,
    uiAddLeadColumn: uiAddLeadColumn,
    uiDeleteLeadColumn: uiDeleteLeadColumn,
    uiRunAction:     uiRunAction
  };

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return reply({ ok: false, error: 'Empty request' });
    }
    const req = JSON.parse(e.postData.contents);

    const expected = getSetting('webAppToken');
    if (!expected) {
      return reply({ ok: false, error: 'Remote access is off. Add a webAppToken row to Settings to enable it.' });
    }

    // Refuse while locked out, before even looking at the token.
    const locked = _authLockMessage();
    if (locked) return reply({ ok: false, error: locked });

    if (String(req.token || '') !== String(expected)) {
      _authNoteFailure();
      Logger.log('doPost: rejected request with a bad token');
      return reply({ ok: false, error: 'Unauthorized' });
    }
    _authClearFailures();

    const fn = HANDLERS[req.fn];
    if (!fn) return reply({ ok: false, error: 'Unknown function: ' + req.fn });

    const data = fn.apply(null, req.args || []);
    return reply({ ok: true, data: data });

  } catch (err) {
    Logger.log('doPost ERROR: ' + err.message);
    return reply({ ok: false, error: err.message });
  }
}

// ============================================================
// REMOTE AUTH - failure lockout
// ============================================================
//
// With the deployment set to "Anyone" the /exec URL is public and the token is
// the only lock. Apps Script does nothing to slow down guessing, so this does:
// 10 bad tokens inside 15 minutes and everything is refused until the window
// passes. Irrelevant when Apps Script serves the page itself - the front end
// calls the functions directly and never reaches doPost.

var _AUTH_KEY      = 'authFails';
var _AUTH_MAX      = 10;
var _AUTH_WINDOWMS = 15 * 60 * 1000;

function _authState() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(_AUTH_KEY) || '{}'); }
  catch (e) { return {}; }
}

/** Returns a message while locked out, or null when requests may proceed. */
function _authLockMessage() {
  const st  = _authState();
  const now = Date.now();
  if (!st.first || !st.n) return null;
  if (now - st.first > _AUTH_WINDOWMS) { _authClearFailures(); return null; }
  if (st.n < _AUTH_MAX) return null;
  const mins = Math.ceil((_AUTH_WINDOWMS - (now - st.first)) / 60000);
  return 'Too many failed attempts. Locked for about ' + mins + ' more minute' +
         (mins === 1 ? '' : 's') + '.';
}

function _authNoteFailure() {
  const now = Date.now();
  let st = _authState();
  if (!st.first || now - st.first > _AUTH_WINDOWMS) st = { first: now, n: 0 };
  st.n = (st.n || 0) + 1;
  try {
    PropertiesService.getScriptProperties().setProperty(_AUTH_KEY, JSON.stringify(st));
  } catch (e) { Logger.log('_authNoteFailure: ' + e.message); }
  Logger.log('doPost: failed auth ' + st.n + '/' + _AUTH_MAX + ' in this window');
}

function _authClearFailures() {
  try { PropertiesService.getScriptProperties().deleteProperty(_AUTH_KEY); }
  catch (e) {}
}

// ============================================================
// SHARED HELPERS
// ============================================================

/**
 * Normalises a cell that may hold a Date or a string into YYYY-MM-DD.
 * Delegates to sheetDateToStr in Code.gs, which reads the date back in the
 * SPREADSHEET's timezone - the same one Sheets used to store it.
 */
function _uiDate(v) {
  return sheetDateToStr(v);
}

/**
 * Finds a lead's CURRENT row by threadId. Never trust a stored row number -
 * enrolling a lead inserts at row 2 and shifts everything below it down.
 */
function _uiFindRowByThreadId(sheet, threadId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, CONFIG.cols.threadId, lastRow - 1, 1).getValues();
  const want = String(threadId).trim();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === want) return i + 2;
  }
  return -1;
}

// ============================================================
// API - DASHBOARD
// ============================================================

function uiGetStats() {
  const c      = CONFIG.cols;
  const counts = { Active: 0, Replied: 0, OOO: 0, Bounced: 0, Done: 0, Error: 0, Paused: 0 };

  const sheet   = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, c.status, lastRow - 1, 1).getValues().forEach(function (r) {
      const s = String(r[0]).trim();
      if (counts[s] !== undefined) counts[s]++;
    });
  }

  let running = false;
  try { running = chainIsRunning(); } catch (e) {}

  return {
    counts:       counts,
    sendsToday:   getSendsToday(),
    maxSends:     parseInt(getSetting('maxSendsPerDay') || '50'),
    paused:       getSetting('pauseAllFollowUps').toUpperCase() === 'TRUE',
    chainRunning: running,
    today:        todayStr(),
    windowFrom:   getSetting('followUpStartTime') || '',
    windowTo:     getSetting('followUpEndTime') || '',
    timezone:     getSetting('timezone') || 'UTC'
  };
}

/** Kicks the chain off by hand. Same function the 8am trigger calls. */
function uiRunChainNow() {
  startDailyRun();
  return 'Chain started. Check back in a few minutes.';
}

// ============================================================
// API - LEADS
// ============================================================

function uiGetLeads() {
  const sheet   = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const c    = CONFIG.cols;
  const data = sheet.getRange(2, 1, lastRow - 1, c.notes).getValues();

  return data.map(function (r, i) {
    return {
      row:          i + 2,
      leadName:     String(r[c.leadName - 1] || ''),
      leadEmail:    String(r[c.leadEmail - 1] || ''),
      threadId:     String(r[c.threadId - 1] || ''),
      sequenceName: String(r[c.sequenceName - 1] || ''),
      step:         Number(r[c.sequenceStep - 1]) || 0,
      totalSteps:   Number(r[c.totalSteps - 1]) || 0,
      nextSendDate: _uiDate(r[c.nextSendDate - 1]),
      status:       String(r[c.status - 1] || ''),
      lastSentDate: _uiDate(r[c.lastSentDate - 1]),
      snippet:      String(r[c.lastReplySnippet - 1] || ''),
      notes:        String(r[c.notes - 1] || '')
    };
  }).filter(function (l) { return l.leadEmail || l.threadId; });
}

/**
 * Changes one lead's status. Finds the row by threadId, never by a row number
 * sent up from the phone - rows move.
 */
function uiSetLeadStatus(threadId, status) {
  const ALLOWED = ['Active', 'Replied', 'Paused', 'Done', 'Error', 'OOO', 'Bounced'];
  if (ALLOWED.indexOf(status) === -1) throw new Error('Unknown status: ' + status);

  const sheet = getSheet(CONFIG.sheets.activeFollowUps);
  const c     = CONFIG.cols;
  const row   = _uiFindRowByThreadId(sheet, threadId);
  if (row === -1) throw new Error('That lead is no longer in ActiveFollowUps.');

  sheet.getRange(row, c.status).setValue(status);

  // Reactivating: make it due today so the next run picks it up.
  if (status === 'Active') {
    sheet.getRange(row, c.nextSendDate).setValue(todayStr());
  }

  SpreadsheetApp.flush();
  Logger.log('uiSetLeadStatus: row ' + row + ' -> ' + status);
  return uiGetLeads();
}

// ============================================================
// API - ADD LEAD
// ============================================================

/**
 * Searches Gmail for threads involving an address and returns them as plain
 * objects. Writes nothing - the InboxScanner sheet is not involved at all,
 * so adding a lead from your phone is a single round trip.
 */
function uiSearchThreads(email) {
  email = String(email || '').trim();
  if (!email || email.indexOf('@') === -1) throw new Error('Enter a valid email address.');

  // Capped at 25: every thread costs a Gmail round trip below.
  const threads  = GmailApp.search('from:' + email + ' OR to:' + email, 0, 25);
  const enrolled = {};
  try {
    uiGetLeads().forEach(function (l) { enrolled[l.threadId] = l.status; });
  } catch (e) {}

  return threads.map(function (t) {
    const msgs = t.getMessages();
    let body = '';
    try {
      body = (msgs[0].getPlainBody() || stripHtml(msgs[0].getBody()))
               .trim().replace(/\s+/g, ' ').substring(0, 300);
    } catch (e) { body = ''; }

    const d = t.getLastMessageDate();
    const id = t.getId();
    return {
      threadId:  id,
      subject:   t.getFirstMessageSubject(),
      date:      d.getFullYear() + '-' +
                 String(d.getMonth() + 1).padStart(2, '0') + '-' +
                 String(d.getDate()).padStart(2, '0'),
      msgCount:  msgs.length,
      body:      body,
      alreadyIn: enrolled[id] || ''
    };
  });
}

/**
 * Enrols a thread straight into ActiveFollowUps.
 * Takes the same script lock enrollLead uses, so a phone enrolment and a
 * checkbox enrolment can never interleave and corrupt the row order.
 */
function uiEnrollLead(threadId, leadEmail, leadName, sequenceName, totalSteps) {
  threadId  = String(threadId || '').trim();
  leadEmail = String(leadEmail || '').trim();
  if (!threadId)  throw new Error('No thread selected.');
  if (!leadEmail) throw new Error('Lead email is required.');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('Another enrolment is running. Try again in a moment.');
  }

  try {
    const sheet = getSheet(CONFIG.sheets.activeFollowUps);

    if (_uiFindRowByThreadId(sheet, threadId) !== -1) {
      throw new Error('That thread is already enrolled.');
    }

    // Today if today's sending window is still open, otherwise tomorrow.
    // Uses the same clock as the engine - see firstSendDate in Code.gs.
    const firstSend = firstSendDate();

    const newRow = [
      String(leadName || ''),
      leadEmail,
      threadId,
      getSendingAccounts()[0] || '',
      String(sequenceName || getSetting('enrollDefaultSequence') || ''),
      0,
      Number(totalSteps) || parseInt(getSetting('enrollDefaultTotalSteps') || '0'),
      firstSend,
      'Active',
      getSetting('enrollDefaultResumeOnReply').toUpperCase() === 'TRUE',
      '', '', ''
    ];

    sheet.insertRowBefore(2);
    sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
    sheet.getRange(2, CONFIG.cols.nextSendDate).setNumberFormat('yyyy-MM-dd');
    SpreadsheetApp.flush();

    Logger.log('uiEnrollLead: ' + leadEmail + ' | thread ' + threadId +
               ' | seq ' + newRow[4] + ' | first send ' + firstSend);
    return uiGetLeads();

  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// API - SETTINGS
// ============================================================

function uiGetSettings() {
  const sheet   = getSheet(CONFIG.sheets.settings);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      const name = String(r[0]).trim();
      return {
        name:  name,
        // Never ship the remote-access token down to the page.
        value: name === 'webAppToken' ? '••••••••' : String(r[1]).trim()
      };
    });
}

function uiSaveSetting(name, value) {
  name = String(name || '').trim();
  if (!name) throw new Error('No setting name given.');
  if (name === 'webAppToken') throw new Error('Change webAppToken in the Sheet, not here.');

  const sheet   = getSheet(CONFIG.sheets.settings);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Settings sheet is empty.');

  const rows = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === name) {
      sheet.getRange(i + 2, 2).setValue(value);
      SpreadsheetApp.flush();
      // The settings cache is per-execution, but clear it anyway so anything
      // called later in THIS request sees the new value.
      for (const k in _settingsCache) delete _settingsCache[k];
      Logger.log('uiSaveSetting: ' + name + ' = ' + value);
      return uiGetSettings();
    }
  }
  throw new Error('Unknown setting: ' + name);
}

// ============================================================
// API - SEND LOG
// ============================================================

function uiGetLog(limit) {
  limit = Number(limit) || 50;
  const sheet   = getSheet(CONFIG.sheets.sendLog);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const take  = Math.min(limit, lastRow - 1);
  const start = lastRow - take + 1;
  const rows  = sheet.getRange(start, 1, take, 9).getValues();

  return rows.map(function (r) {
    let ts = '';
    if (r[0] instanceof Date) {
      const d = r[0];
      ts = d.getFullYear() + '-' +
           String(d.getMonth() + 1).padStart(2, '0') + '-' +
           String(d.getDate()).padStart(2, '0') + ' ' +
           String(d.getHours()).padStart(2, '0') + ':' +
           String(d.getMinutes()).padStart(2, '0');
    } else {
      ts = String(r[0] || '');
    }
    return {
      time:     ts,
      email:    String(r[1] || ''),
      name:     String(r[2] || ''),
      sequence: String(r[3] || ''),
      step:     String(r[4] || ''),
      preview:  String(r[6] || ''),
      result:   String(r[8] || '')
    };
  }).reverse(); // newest first
}

// ============================================================
// API - SEQUENCES (read only)
// ============================================================

function uiGetSequences() {
  const sheet   = getSheet(CONFIG.sheets.sequences);
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 1) return [];

  const all     = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = all[0];
  const out     = [];

  for (let ci = 0; ci < headers.length; ci++) {
    const name = String(headers[ci]).trim();
    if (!name || name === 'awaitDays') continue;
    if (ci === 0) continue; // a sequence needs an awaitDays column to its left

    const steps = [];
    for (let r = 1; r < all.length; r++) {
      const msg = String(all[r][ci] || '').trim();
      // The engine stops at the first empty row, so this view does too.
      if (!msg) break;
      steps.push({
        n:         steps.length + 1,
        awaitDays: Number(all[r][ci - 1]) || 1,
        message:   msg
      });
    }
    out.push({ name: name, steps: steps });
  }
  return out;
}

// ============================================================
// API - SEQUENCE EDITING
// ============================================================
//
// The Sequences sheet is column PAIRS sharing rows: awaitDays, then a message
// column whose header is the sequence name. Because the pairs share rows, we
// never insert or delete whole rows here - that would shift every other
// sequence. We only ever touch the two columns belonging to one sequence.
//
// The engine (getSequenceStep) stops at the first empty message row, so steps
// must stay contiguous from row 2 down. Every function below preserves that.

/** Returns the 1-based message column for a sequence, or throws. */
function _uiSeqCol(sheet, seqName) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error('Sequences sheet is empty.');
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const want = String(seqName).trim();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === want) {
      if (i < 1) throw new Error('"' + want + '" has no awaitDays column to its left.');
      return i + 1;
    }
  }
  throw new Error('Sequence not found: ' + want);
}

/** Reads one sequence's steps as [{awaitDays, message}], stopping at the first blank. */
function _uiReadSteps(sheet, msgCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const vals = sheet.getRange(2, msgCol - 1, lastRow - 1, 2).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const msg = String(vals[i][1] || '').trim();
    if (!msg) break;
    out.push({ awaitDays: Number(vals[i][0]) || 1, message: msg });
  }
  return out;
}

/**
 * Creates a new sequence: an awaitDays column plus a named message column,
 * placed after a spacer so the layout matches what Setup builds.
 */
function uiAddSequence(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('Enter a sequence name.');
  if (name.toLowerCase() === 'awaitdays') throw new Error('That name is reserved.');

  const sheet   = getSheet(CONFIG.sheets.sequences);
  const lastCol = sheet.getLastColumn();

  if (lastCol >= 1) {
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === name) {
        throw new Error('A sequence called "' + name + '" already exists.');
      }
    }
  }

  const start = lastCol > 0 ? lastCol + 2 : 1;   // +2 leaves a spacer column
  const need  = start + 1;
  if (need > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());

  sheet.getRange(1, start).setValue('awaitDays').setFontWeight('bold');
  sheet.getRange(1, start + 1).setValue(name).setFontWeight('bold');
  sheet.setColumnWidth(start, 100);
  sheet.setColumnWidth(start + 1, 350);
  SpreadsheetApp.flush();

  Logger.log('uiAddSequence: created "' + name + '" in columns ' + start + '/' + (start + 1));
  return uiGetSequences();
}

/**
 * Writes one step. stepNumber may be an existing step (edit) or exactly one
 * past the end (append) - never further, since a gap would silently end the
 * sequence early for every lead running it.
 */
function uiSaveStep(seqName, stepNumber, awaitDays, message) {
  const sheet  = getSheet(CONFIG.sheets.sequences);
  const msgCol = _uiSeqCol(sheet, seqName);
  const steps  = _uiReadSteps(sheet, msgCol);

  const n = Number(stepNumber);
  if (!n || n < 1) throw new Error('Bad step number.');
  if (n > steps.length + 1) {
    throw new Error('Step ' + n + ' would leave a gap. The next step here is ' + (steps.length + 1) + '.');
  }

  const msg = String(message || '').trim();
  if (!msg) throw new Error('Message cannot be empty. Use delete to remove a step.');

  const row = n + 1;
  if (row > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows());
  }

  sheet.getRange(row, msgCol - 1).setValue(Number(awaitDays) || 1);
  sheet.getRange(row, msgCol).setValue(msg);
  SpreadsheetApp.flush();

  Logger.log('uiSaveStep: ' + seqName + ' step ' + n + ' (' + msg.length + ' chars)');
  return uiGetSequences();
}

/**
 * Removes a step and closes the gap, rewriting only this sequence's two
 * columns so other sequences sharing those rows are untouched.
 */
function uiDeleteStep(seqName, stepNumber) {
  const sheet  = getSheet(CONFIG.sheets.sequences);
  const msgCol = _uiSeqCol(sheet, seqName);
  const steps  = _uiReadSteps(sheet, msgCol);

  const n = Number(stepNumber);
  if (!n || n < 1 || n > steps.length) throw new Error('Step ' + stepNumber + ' does not exist.');

  steps.splice(n - 1, 1);

  // Rewrite the whole block, one blank row longer, to clear the old tail.
  const block = steps.map(function (s) { return [s.awaitDays, s.message]; });
  block.push(['', '']);
  sheet.getRange(2, msgCol - 1, block.length, 2).setValues(block);
  SpreadsheetApp.flush();

  Logger.log('uiDeleteStep: ' + seqName + ' step ' + n + ' removed, ' + steps.length + ' left');
  return uiGetSequences();
}

// ============================================================
// API - FULL LEAD ROW
// ============================================================

/** Field kinds so the phone can pick the right input for each column. */
function _uiKinds() {
  const c = CONFIG.cols, k = {};
  k[c.status]           = 'status';
  k[c.nextSendDate]     = 'date';
  k[c.lastSentDate]     = 'date';
  k[c.sequenceStep]     = 'number';
  k[c.totalSteps]       = 'number';
  k[c.resumeOnReply]    = 'bool';
  k[c.sequenceName]     = 'sequence';
  k[c.notes]            = 'long';
  k[c.lastReplySnippet] = 'long';
  return k;
}

/**
 * Returns every column of one lead's row, custom columns included, so the
 * app can show and edit the whole row rather than a fixed subset.
 */
function uiGetLeadRow(threadId) {
  const sheet = getSheet(CONFIG.sheets.activeFollowUps);
  const row   = _uiFindRowByThreadId(sheet, threadId);
  if (row === -1) throw new Error('That lead is no longer in ActiveFollowUps.');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values  = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const c = CONFIG.cols, kinds = _uiKinds();

  const cells = [];
  for (let i = 0; i < lastCol; i++) {
    const name = String(headers[i] || '').trim();
    if (!name) continue;
    const col = i + 1;
    const isDate = (col === c.nextSendDate || col === c.lastSentDate);
    cells.push({
      col:      col,
      name:     name,
      value:    isDate ? _uiDate(values[i])
                       : (values[i] === null || values[i] === undefined ? '' : String(values[i])),
      kind:     kinds[col] || 'text',
      // threadId is the key we look the row up by - editing it would orphan the lead
      readonly: col === c.threadId
    });
  }
  return { row: row, threadId: threadId, cells: cells };
}

/** Writes back any subset of a lead's columns. Keyed by header name. */
function uiSaveLeadRow(threadId, updates) {
  const sheet = getSheet(CONFIG.sheets.activeFollowUps);
  const row   = _uiFindRowByThreadId(sheet, threadId);
  if (row === -1) throw new Error('That lead is no longer in ActiveFollowUps.');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const c = CONFIG.cols;
  const ALLOWED_STATUS = ['Active', 'Replied', 'Paused', 'Done', 'Error', 'OOO', 'Bounced'];

  let changed = 0;
  for (let i = 0; i < lastCol; i++) {
    const name = String(headers[i] || '').trim();
    if (!name || !updates.hasOwnProperty(name)) continue;

    const col = i + 1;
    if (col === c.threadId) continue;   // never editable

    let v = updates[name];
    if (col === c.status) {
      v = String(v).trim();
      if (ALLOWED_STATUS.indexOf(v) === -1) throw new Error('Unknown status: ' + v);
    } else if (col === c.sequenceStep || col === c.totalSteps) {
      v = Number(v) || 0;
    } else if (col === c.resumeOnReply) {
      v = (v === true || String(v).toUpperCase() === 'TRUE');
    }

    sheet.getRange(row, col).setValue(v);
    changed++;
  }

  SpreadsheetApp.flush();
  Logger.log('uiSaveLeadRow: row ' + row + ', ' + changed + ' field(s) written');
  return uiGetLeads();
}

// ============================================================
// API - VARIABLES
// ============================================================

/**
 * Every {{token}} usable in a message: the three built-ins plus any extra
 * column header added to ActiveFollowUps past the notes column.
 */
function uiGetVariables() {
  const out = [
    { token: '{{firstName}}', name: 'firstName', kind: 'built-in' },
    { token: '{{leadName}}',  name: 'leadName',  kind: 'built-in' },
    { token: '{{leadEmail}}', name: 'leadEmail', kind: 'built-in' }
  ];
  try {
    const sheet   = getSheet(CONFIG.sheets.activeFollowUps);
    const lastCol = sheet.getLastColumn();
    if (lastCol > CONFIG.cols.notes) {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (let i = CONFIG.cols.notes; i < lastCol; i++) {
        const h = String(headers[i] || '').trim();
        if (h) out.push({ token: '{{' + h + '}}', name: h, kind: 'custom' });
      }
    }
  } catch (e) {
    Logger.log('uiGetVariables: ' + e.message);
  }
  return out;
}

// ============================================================
// API - STEP PREVIEW & TEST SEND
// ============================================================

/** Pulls a lead's variable values. Returns blanks when threadId is empty. */
function _uiResolveLead(threadId) {
  const out = { leadName: '', leadEmail: '', customVars: {} };
  if (!threadId) return out;

  const sheet = getSheet(CONFIG.sheets.activeFollowUps);
  const row   = _uiFindRowByThreadId(sheet, threadId);
  if (row === -1) return out;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const vals    = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const c = CONFIG.cols;

  out.leadName  = String(vals[c.leadName - 1] || '');
  out.leadEmail = String(vals[c.leadEmail - 1] || '');
  for (let i = c.notes; i < lastCol; i++) {
    const h = String(headers[i] || '').trim();
    if (h) out.customVars[h] = String(vals[i] || '').trim();
  }
  return out;
}

/**
 * Renders a step exactly as a lead would receive it, image included.
 * The HTML comes back already escaped, so the page inserts it as-is.
 */
function uiPreviewStep(seqName, stepNumber, threadId) {
  const stepData = getSequenceStep(seqName, Number(stepNumber));
  if (!stepData || !stepData.message) {
    throw new Error('Step ' + stepNumber + ' does not exist in "' + seqName + '".');
  }

  const L      = _uiResolveLead(threadId);
  const parsed = parseMessage(stepData.message);
  const text   = replaceVariables(parsed.text, L.leadName, L.leadEmail, L.customVars);

  // Escape the message before it becomes HTML. {{IMG_PLACEHOLDER}} has no
  // HTML characters so it survives this untouched.
  const safe = String(text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  let body = safe.replace(/\n/g, '<br>');

  let imgTag = '';
  let note   = '';
  if (parsed.imgUrl) {
    try {
      const blob = getImageBlob(parsed.imgUrl);
      if (!blob) {
        note = 'Image URL could not be read.';
      } else {
        const bytes = blob.getBytes();
        if (bytes.length > 3 * 1024 * 1024) {
          note = 'Image is ' + Math.round(bytes.length / 1048576) +
                 ' MB - too big to preview here, but it will still send.';
        } else {
          imgTag = '<img src="data:' + (blob.getContentType() || 'image/jpeg') +
                   ';base64,' + Utilities.base64Encode(bytes) +
                   '" style="width:' + (parsed.imgSize || '100%') +
                   ';max-width:100%;border-radius:8px;display:block;margin:10px 0">';
        }
      }
    } catch (e) {
      note = 'Image failed to load: ' + e.message;
    }
  }

  body = body.indexOf('{{IMG_PLACEHOLDER}}') !== -1
    ? body.replace('{{IMG_PLACEHOLDER}}', imgTag)
    : body + imgTag;

  return {
    html:      body,
    note:      note,
    awaitDays: stepData.awaitDays,
    usedLead:  L.leadEmail,
    hasVars:   /\{\{[^}]+\}\}/.test(text)   // an unresolved token is left visible
  };
}

/**
 * Sends one step to yourself as a real email. No banner - the body is exactly
 * what a lead would get. Only the [TEST] subject prefix marks it.
 * Images go inline via Content-ID, same as the live sender, because Gmail
 * strips data: URIs out of email bodies.
 */
function uiTestSendStep(seqName, stepNumber, threadId, toEmail) {
  const to = String(toEmail || '').trim() || getSetting('notificationEmail');
  if (!to) throw new Error('No recipient. Set notificationEmail in Settings, or type an address.');
  if (to.indexOf('@') === -1) throw new Error('"' + to + '" is not an email address.');

  const stepData = getSequenceStep(seqName, Number(stepNumber));
  if (!stepData || !stepData.message) {
    throw new Error('Step ' + stepNumber + ' does not exist in "' + seqName + '".');
  }

  const L        = _uiResolveLead(threadId);
  const parsed   = parseMessage(stepData.message);
  const textPart = replaceVariables(parsed.text, L.leadName, L.leadEmail, L.customVars);
  const subject  = '[TEST] ' + seqName + ' step ' + stepNumber +
                   (L.leadEmail ? ' - ' + L.leadEmail : '');

  let imgBlob = null;
  if (parsed.imgUrl) {
    try { imgBlob = getImageBlob(parsed.imgUrl); }
    catch (e) { Logger.log('uiTestSendStep: image - ' + e.message); }
  }

  const htmlBody = textPart.replace(/\n/g, '<br>');

  if (imgBlob) {
    const cid      = 'test-' + Date.now();
    const boundary = 'b_' + Utilities.getUuid().replace(/-/g, '');
    const imgTag   = '<br><img src="cid:' + cid + '" style="width:' +
                     (parsed.imgSize || '100%') + ';max-width:100%"><br>';
    const withImg  = htmlBody.indexOf('{{IMG_PLACEHOLDER}}') !== -1
      ? htmlBody.replace('{{IMG_PLACEHOLDER}}', imgTag)
      : htmlBody + imgTag;

    const raw = [
      'MIME-Version: 1.0',
      'From: ' + Session.getActiveUser().getEmail(),
      'To: ' + to,
      'Subject: ' + subject,
      'Content-Type: multipart/related; boundary="' + boundary + '"',
      '',
      '--' + boundary,
      'Content-Type: text/html; charset=UTF-8',
      '',
      withImg,
      '',
      '--' + boundary,
      'Content-Type: ' + (imgBlob.getContentType() || 'image/jpeg') + '; name="image"',
      'Content-Transfer-Encoding: base64',
      'Content-ID: <' + cid + '>',
      'Content-Disposition: inline; filename="image"',
      '',
      Utilities.base64Encode(imgBlob.getBytes()),
      '',
      '--' + boundary + '--'
    ].join('\r\n');

    Gmail.Users.Messages.send({ raw: Utilities.base64EncodeWebSafe(raw) }, 'me');
  } else {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: htmlBody.replace('{{IMG_PLACEHOLDER}}', '')
    });
  }

  Logger.log('uiTestSendStep: ' + seqName + ' step ' + stepNumber + ' -> ' + to);
  return 'Test sent to ' + to;
}

// ============================================================
// API - DELETE A SEQUENCE
// ============================================================

/**
 * Removes a sequence's two columns. Deleting columns is safe here (unlike
 * deleting rows) because every lookup is by header name, not position.
 *
 * Refuses while leads are still running it. Finished leads - Done, Replied,
 * Bounced, Error - never read the sequence again, so they do not block.
 */
function uiDeleteSequence(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('No sequence named.');

  const sheet  = getSheet(CONFIG.sheets.sequences);
  const msgCol = _uiSeqCol(sheet, name);

  const LIVE    = ['Active', 'OOO', 'Paused'];
  const active  = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow = active.getLastRow();
  if (lastRow >= 2) {
    const c    = CONFIG.cols;
    const rows = active.getRange(2, 1, lastRow - 1, c.status).getValues();
    let live = 0;
    rows.forEach(function (r) {
      if (String(r[c.sequenceName - 1]).trim() === name &&
          LIVE.indexOf(String(r[c.status - 1]).trim()) !== -1) live++;
    });
    if (live) {
      throw new Error(live + ' lead' + (live > 1 ? 's are' : ' is') + ' still running "' + name +
        '". Move them to another sequence, or mark them Done, then try again.');
    }
  }

  sheet.deleteColumns(msgCol - 1, 2);
  SpreadsheetApp.flush();
  Logger.log('uiDeleteSequence: removed "' + name + '"');
  return uiGetSequences();
}

// ============================================================
// API - CUSTOM LEAD COLUMNS
// ============================================================

/**
 * Adds a new column to ActiveFollowUps. Any header past the notes column
 * automatically becomes a {{variable}} usable in messages, so this is how you
 * add a merge field without opening the Sheet.
 */
function uiAddLeadColumn(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('Enter a column name.');
  if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(name)) {
    throw new Error('Use letters, numbers and underscores only, starting with a letter. ' +
                    'It becomes {{' + name.replace(/[^A-Za-z0-9_]/g, '') + '}} in your messages.');
  }

  const sheet   = getSheet(CONFIG.sheets.activeFollowUps);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === name.toLowerCase()) {
      throw new Error('A column called "' + name + '" already exists.');
    }
  }

  const col = lastCol + 1;
  if (col > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
  sheet.getRange(1, col).setValue(name).setFontWeight('bold');
  sheet.setColumnWidth(col, 180);
  SpreadsheetApp.flush();

  Logger.log('uiAddLeadColumn: added "' + name + '" at column ' + col);
  return { name: name, col: col, token: '{{' + name + '}}' };
}

/**
 * Runs one of the maintenance actions by name. An explicit allow-list, so the
 * remote endpoint can never be talked into calling something arbitrary.
 */
function uiRunAction(action) {
  switch (String(action)) {
    case 'startDailyRun':
      startDailyRun();
      return 'Chain started. Watch the Leads tab in a few minutes.';

    case 'manualProcessOneLead':
      processOneLead(null);
      return 'Processed one lead. Check the Log tab.';

    case 'createTriggers':
      createTriggers();
      return 'Triggers recreated: 8am daily, Monday notification, enroll watcher.';

    case 'cleanUpStuckTriggers':
      cleanUpStuckTriggers();
      return 'Stuck triggers cleared and the chain lock released.';

    case 'sendNotificationEmail':
      sendNotificationEmail();
      return 'Summary email sent to your notificationEmail.';

    case 'clearParked':
      PropertiesService.getScriptProperties().deleteProperty('parkedLeads');
      return 'Parked list cleared - today\'s already-sent guard is reset.';

    case 'clearAuthLock':
      _authClearFailures();
      return 'Sign-in lockout cleared.';

    case 'clearQuotaStop':
      PropertiesService.getScriptProperties().deleteProperty('quotaStopDate');
      return 'Quota stop cleared. The chain can run again today.';

    default:
      throw new Error('Unknown action: ' + action);
  }
}

/**
 * Removes a column you previously added to ActiveFollowUps.
 *
 * Refuses for the built-in columns the engine depends on, and refuses while
 * any sequence message still uses the token - otherwise the next email a real
 * prospect receives would contain the literal {{token}}.
 */
function uiDeleteLeadColumn(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('No column named.');

  const c       = CONFIG.cols;
  const sheet   = getSheet(CONFIG.sheets.activeFollowUps);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let col = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === name) { col = i + 1; break; }
  }
  if (col === -1) throw new Error('No column called "' + name + '".');
  if (col <= c.notes) {
    throw new Error('"' + name + '" is a built-in column the system needs. ' +
                    'Only columns you added yourself can be removed.');
  }

  const token = '{{' + name + '}}';
  const used  = [];
  try {
    uiGetSequences().forEach(function (s) {
      s.steps.forEach(function (st) {
        if (st.message.indexOf(token) !== -1) used.push(s.name + ' step ' + st.n);
      });
    });
  } catch (e) {
    Logger.log('uiDeleteLeadColumn: could not scan sequences - ' + e.message);
  }
  if (used.length) {
    throw new Error(token + ' is still used in ' + used.join(', ') +
      '. Take it out of those messages first, or a lead would be emailed the raw token.');
  }

  sheet.deleteColumns(col, 1);
  SpreadsheetApp.flush();
  Logger.log('uiDeleteLeadColumn: removed "' + name + '" (was column ' + col + ')');
  return { name: name };
}

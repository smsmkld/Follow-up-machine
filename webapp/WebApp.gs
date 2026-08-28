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
    uiRunChainNow:   uiRunChainNow
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
    if (String(req.token || '') !== String(expected)) {
      Logger.log('doPost: rejected request with a bad token');
      return reply({ ok: false, error: 'Unauthorized' });
    }

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
// SHARED HELPERS
// ============================================================

/** Normalises a cell that may hold a Date or a string into YYYY-MM-DD. */
function _uiDate(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' +
           String(v.getMonth() + 1).padStart(2, '0') + '-' +
           String(v.getDate()).padStart(2, '0');
  }
  return String(v || '').trim();
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

    const now = new Date();
    const tom = new Date(now);
    tom.setDate(tom.getDate() + 1);
    const tomorrow = tom.getFullYear() + '-' +
                     String(tom.getMonth() + 1).padStart(2, '0') + '-' +
                     String(tom.getDate()).padStart(2, '0');

    const newRow = [
      String(leadName || ''),
      leadEmail,
      threadId,
      getSendingAccounts()[0] || '',
      String(sequenceName || getSetting('enrollDefaultSequence') || ''),
      0,
      Number(totalSteps) || parseInt(getSetting('enrollDefaultTotalSteps') || '0'),
      tomorrow,
      'Active',
      getSetting('enrollDefaultResumeOnReply').toUpperCase() === 'TRUE',
      '', '', ''
    ];

    sheet.insertRowBefore(2);
    sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
    sheet.getRange(2, CONFIG.cols.nextSendDate).setNumberFormat('yyyy-MM-dd');
    SpreadsheetApp.flush();

    Logger.log('uiEnrollLead: ' + leadEmail + ' | thread ' + threadId + ' | seq ' + newRow[4]);
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

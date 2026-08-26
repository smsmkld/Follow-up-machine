// ============================================================
// SECTION: CONFIGURATION
// ============================================================

const CONFIG = {
  // Auto-detected from the Google account running this script.
  // All replies go here and sending always comes from this account.
  sendingAccounts: [Session.getActiveUser().getEmail()],

  // OOO/bounce keywords (case-insensitive match against reply body)
  oooKeywords: [
    'out of office', 'on leave', 'away from', 'vacation',
    'undeliverable', 'delivery failed', 'bounce',
    'auto-reply', 'automatic reply', 'not available', 'on holiday'
  ],

  // Sheet names
  sheets: {
    settings:        'Settings',
    inboxScanner:    'InboxScanner',
    sequences:       'Sequences',
    activeFollowUps: 'ActiveFollowUps',
    sendLog:         'SendLog',
  },

  // Column indices for ActiveFollowUps (1-based)
  cols: {
    leadName:         1,
    leadEmail:        2,
    threadId:         3,
    fromAccount:      4,
    sequenceName:     5,
    sequenceStep:     6,
    totalSteps:       7,
    nextSendDate:     8,
    status:           9,
    resumeOnReply:    10,
    lastSentDate:     11,
    lastReplySnippet: 12,
    notes:            13,
  },

  // Column indices for InboxScanner (1-based)
  scannerCols: {
    leadEmail:     1,
    threadId:      2,
    subject:       3,
    fromAccount:   4,
    lastReplyDate: 5,
    body:          6,  // ← NEW: first message plain body (truncated)
    enroll:        7,  // ← shifted from 6 to 7
  },

  // PropertiesService keys for daily send counter
  props: {
    sendsToday:     'sendsToday',
    sendsTodayDate: 'sendsTodayDate',
  }
};

// ============================================================
// SECTION: TIME API HELPER
// ============================================================

/**
 * Cache for timeapi.io responses within a single script run.
 * Key: timezone string  →  Value: { dateTime: '2024-01-15T09:30:00', ... }
 */
const _timeApiCache = {};

/**
 * Fetches the current datetime for the given timezone from timeapi.io.
 * Returns the parsed response object, or null on failure.
 * Results are cached per script run to avoid repeated API calls.
 *
 * API endpoint: https://timeapi.io/api/v1/timezone/zone?timeZone=<tz>
 * Response shape: { timeZone, currentLocalTime, currentUtcOffset, ... }
 * currentLocalTime format: "2024-01-15 09:30:00.123456"
 */
function fetchTimeApiData(tz) {
  if (!tz) tz = 'UTC';
  if (_timeApiCache[tz] !== undefined) return _timeApiCache[tz];

  try {
    const url      = 'https://timeapi.io/api/v1/timezone/zone?timeZone=' + encodeURIComponent(tz);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code     = response.getResponseCode();
    if (code !== 200) {
      Logger.log('fetchTimeApiData: HTTP ' + code + ' for tz=' + tz + ' - falling back to Utilities.formatDate');
      _timeApiCache[tz] = null;
      return null;
    }
    const data = JSON.parse(response.getContentText());
    _timeApiCache[tz] = data;
    return data;
  } catch (e) {
    Logger.log('fetchTimeApiData ERROR for tz=' + tz + ': ' + e.message + ' - falling back to Utilities.formatDate');
    _timeApiCache[tz] = null;
    return null;
  }
}

/**
 * Returns the current time as HH:mm string using timeapi.io.
 * Falls back to Utilities.formatDate if the API is unreachable.
 */
function currentTimeHHMM() {
  const tz   = getSetting('timezone') || 'UTC';
  const data = fetchTimeApiData(tz);
  if (data && data.currentLocalTime) {
    // currentLocalTime: "2024-01-15 09:30:00.123456"
    // Extract HH:mm from position 11..15
    return data.currentLocalTime.substring(11, 16);
  }
  // Fallback
  return Utilities.formatDate(new Date(), tz, 'HH:mm');
}

/**
 * Returns today's date as YYYY-MM-DD string using timeapi.io.
 * Falls back to Utilities.formatDate if the API is unreachable.
 */
function todayStr() {
  const tz   = getSetting('timezone') || 'UTC';
  const data = fetchTimeApiData(tz);
  if (data && data.currentLocalTime) {
    // currentLocalTime: "2024-01-15 09:30:00.123456"
    // Extract YYYY-MM-DD from the first 10 characters
    return data.currentLocalTime.substring(0, 10);
  }
  // Fallback
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

/**
 * Returns the current day abbreviation (Mon, Tue, Wed, Thu, Fri, Sat, Sun)
 * using timeapi.io. Falls back to Utilities.formatDate if the API is unreachable.
 *
 * timeapi.io returns dayOfWeek as a full English name e.g. "Monday".
 * We truncate to the first 3 characters to match the Settings format.
 */
function currentDayAbbr() {
  const tz   = getSetting('timezone') || 'UTC';
  const data = fetchTimeApiData(tz);
  if (data && data.dayOfWeek) {
    // data.dayOfWeek: "Monday", "Tuesday", etc.
    return data.dayOfWeek.substring(0, 3); // "Mon", "Tue", etc.
  }
  // Fallback
  return Utilities.formatDate(new Date(), tz, 'EEE');
}

/**
 * Returns the current full day name (e.g. "Monday") using timeapi.io.
 * Falls back to Utilities.formatDate if the API is unreachable.
 * Used by weeklyNotificationTrigger() to match notificationDay setting.
 */
function currentDayFull() {
  const tz   = getSetting('timezone') || 'UTC';
  const data = fetchTimeApiData(tz);
  if (data && data.dayOfWeek) {
    return data.dayOfWeek; // "Monday", "Tuesday", etc.
  }
  // Fallback
  return Utilities.formatDate(new Date(), tz, 'EEEE');
}

// ============================================================
// SECTION: MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Follow-Up System')
    .addItem('Scan inbox for email (Sheet 1)', 'scanInboxForEmail')
    .addSeparator()
    .addItem('Start follow-up chain now (manual)', 'startDailyRun')
    .addItem('Run one lead now (manual single step)', 'manualProcessOneLead')
    .addSeparator()
    .addItem('Setup spreadsheet (first time)', 'setupSpreadsheet')
    .addItem('Create daily trigger', 'createTriggers')
    .addItem('View all sequence names', 'listSequenceNames')
    .addSeparator()
    .addItem('Send notification email now', 'sendNotificationEmail')
    .addSeparator()
    .addItem('Send test email (first lead, step 1)', 'sendTestEmail')
    .addToUi();
}

// ============================================================
// SECTION: SETUP
// ============================================================

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function recreateSheet(name) {
    const existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
    return ss.insertSheet(name);
  }

  // ── Settings sheet (created FIRST so it ends up leftmost) ─────
  const settingsSheet = recreateSheet(CONFIG.sheets.settings);
  const settingsHeaders = ['setting', 'value'];
  settingsSheet.getRange(1, 1, 1, 2).setValues([settingsHeaders]).setFontWeight('bold');
  settingsSheet.setFrozenRows(1);

  const settingsRows = [
    ['followUpStartTime',          '08:00'],
    ['followUpEndTime',            '18:00'],
    ['followUpSendingDays',        'Mon,Tue,Wed,Thu,Fri'],
    ['notificationEmail',          ''],
    ['notificationFrequency',      'daily'],
    ['notificationDay',            'Monday'],
    ['minDelayMinutes',            '4'],
    ['maxDelayMinutes',            '8'],
    ['maxSendsPerDay',             '50'],
    ['timezone',                   'UTC'],
    ['pauseAllFollowUps',          'FALSE'],
    ['replyCheckBeforeSend',       'TRUE'],
    ['oooAutoResume',              'FALSE'],
    ['oooResumeDays',              '7'],
    ['enrollDefaultSequence',      ''],
    ['enrollDefaultTotalSteps',    '0'],
    ['enrollDefaultResumeOnReply', 'FALSE'],
  ];
  settingsSheet.getRange(2, 1, settingsRows.length, 2).setValues(settingsRows);

  // Lock column A (setting names) visually with a light background
  settingsSheet.getRange('A2:A' + (settingsRows.length + 1)).setBackground('#f3f3f3').setFontStyle('italic');
  settingsSheet.setColumnWidth(1, 220);
  settingsSheet.setColumnWidth(2, 260);

  // ── InboxScanner ──────────────────────────────────────────
  // Columns: leadEmail | threadId | subject | fromAccount | lastReplyDate | body | enroll
  const scanner = recreateSheet(CONFIG.sheets.inboxScanner);
  const scannerHeaders = ['leadEmail', 'threadId', 'subject', 'fromAccount', 'lastReplyDate', 'body', 'enroll'];
  scanner.getRange(1, 1, 1, scannerHeaders.length).setValues([scannerHeaders]).setFontWeight('bold');
  scanner.setFrozenRows(1);

  // Enroll checkbox validation is now on column G (7)
  const enrollValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  scanner.getRange('G2:G1000').setDataValidation(enrollValidation);

  scanner.setColumnWidth(CONFIG.scannerCols.leadEmail,     220); // A
  scanner.setColumnWidth(CONFIG.scannerCols.threadId,      180); // B
  scanner.setColumnWidth(CONFIG.scannerCols.subject,       280); // C
  scanner.setColumnWidth(CONFIG.scannerCols.fromAccount,   220); // D
  scanner.setColumnWidth(CONFIG.scannerCols.lastReplyDate, 160); // E
  scanner.setColumnWidth(CONFIG.scannerCols.body,          400); // F ← NEW - wider for body preview
  scanner.setColumnWidth(CONFIG.scannerCols.enroll,         80); // G

  // Wrap text on body column so it's readable
  scanner.getRange('F2:F1000').setWrap(true);

  // ── Sequences ──────────────────────────────────────────────
  const sequences = recreateSheet(CONFIG.sheets.sequences);
  sequences.getRange('A1').setValue('awaitDays').setFontWeight('bold');
  sequences.getRange('B1').setValue('loomFollowUpMessage').setFontWeight('bold');
  sequences.getRange('D1').setValue('awaitDays').setFontWeight('bold');
  sequences.getRange('E1').setValue('callFollowUpMessage').setFontWeight('bold');
  sequences.setFrozenRows(1);
  sequences.setColumnWidth(1, 100);
  sequences.setColumnWidth(2, 350);
  sequences.setColumnWidth(3, 30);
  sequences.setColumnWidth(4, 100);
  sequences.setColumnWidth(5, 350);

  // ── ActiveFollowUps ────────────────────────────────────────
  // Note: the fromAccount column (D) is informational / read-only.
  // It records which account the original cold email came from.
  // Sending always uses the authenticated Google account - fromAccount does not control sending.
  const active = recreateSheet(CONFIG.sheets.activeFollowUps);
  const activeHeaders = [
    'leadName', 'leadEmail', 'threadId', 'fromAccount', 'sequenceName',
    'sequenceStep', 'totalSteps', 'nextSendDate', 'status', 'resumeOnReply',
    'lastSentDate', 'lastReplySnippet', 'notes'
  ];
  active.getRange(1, 1, 1, activeHeaders.length).setValues([activeHeaders]).setFontWeight('bold');
  active.setFrozenRows(1);
  active.getRange('F2:F1000').setNumberFormat('0');
  active.getRange('G2:G1000').setNumberFormat('0');
  active.getRange('H2:H1000').setNumberFormat('yyyy-MM-dd');
  active.getRange('K2:K1000').setNumberFormat('yyyy-MM-dd');
  const resumeValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  active.getRange('J2:J1000').setDataValidation(resumeValidation);
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Replied', 'Paused', 'Done', 'Error', 'OOO'], true)
    .build();
  active.getRange('I2:I1000').setDataValidation(statusValidation);
  active.setColumnWidth(1, 130);
  active.setColumnWidth(2, 220);
  active.setColumnWidth(3, 180);
  active.setColumnWidth(4, 220);
  active.setColumnWidth(5, 200);
  active.setColumnWidth(6, 110);
  active.setColumnWidth(7, 100);
  active.setColumnWidth(8, 130);
  active.setColumnWidth(9, 100);
  active.setColumnWidth(10, 130);
  active.setColumnWidth(11, 130);
  active.setColumnWidth(12, 250);
  active.setColumnWidth(13, 300);

  // ── SendLog ────────────────────────────────────────────────
  const logSheet = recreateSheet(CONFIG.sheets.sendLog);
  const logHeaders = [
    'timestamp', 'leadEmail', 'leadName', 'sequenceName',
    'stepNumber', 'fromAccount', 'messagePreview', 'threadId', 'result'
  ];
  logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]).setFontWeight('bold');
  logSheet.setFrozenRows(1);
  logSheet.setColumnWidth(1, 170);
  logSheet.setColumnWidth(2, 220);
  logSheet.setColumnWidth(3, 130);
  logSheet.setColumnWidth(4, 200);
  logSheet.setColumnWidth(5, 90);
  logSheet.setColumnWidth(6, 220);
  logSheet.setColumnWidth(7, 300);
  logSheet.setColumnWidth(8, 180);
  logSheet.setColumnWidth(9, 80);

  // ── Reorder sheets: Settings first, then the rest ─────────
  const sheetOrder = [
    CONFIG.sheets.settings,
    CONFIG.sheets.inboxScanner,
    CONFIG.sheets.sequences,
    CONFIG.sheets.activeFollowUps,
    CONFIG.sheets.sendLog,
  ];
  sheetOrder.forEach((name, idx) => {
    const s = ss.getSheetByName(name);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(idx + 1);
    }
  });

  Logger.log(
    'Setup complete! All 5 sheets created.\n\n' +
    'Next steps:\n' +
    '1. Review the Settings sheet and adjust values\n' +
    '2. Run "Create daily trigger" from the menu\n' +
    '3. Add your sequences to the Sequences sheet\n\n' +
    'Emails will be sent from: ' + Session.getActiveUser().getEmail() + '\n\n' +
    'InboxScanner columns: leadEmail | threadId | subject | fromAccount | lastReplyDate | body | enroll\n' +
    'When multiple threads exist for the same email they each get a row - tick Enroll on the right one.'
  );
}

function createTriggers() {
  // Remove existing managed triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'startDailyRun' || fn === 'watchEnrollCheckbox' || fn === 'weeklyNotificationTrigger') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 8am daily trigger → startDailyRun
    const dailyTrigger = ScriptApp.newTrigger('startDailyRun')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  PropertiesService.getScriptProperties().setProperty('recurringDailyTriggerUid', dailyTrigger.getUniqueId());
  Logger.log('createTriggers: recurring daily trigger UID saved: ' + dailyTrigger.getUniqueId());

  // Weekly notification trigger: fires every Monday at 9am
  // (day is also controlled by notificationDay setting, but the trigger fires Monday;
  //  the function checks notificationDay and only sends if it matches)
  ScriptApp.newTrigger('weeklyNotificationTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // Installable onEdit trigger for InboxScanner checkbox (needs Gmail access)
  ScriptApp.newTrigger('watchEnrollCheckbox')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  Logger.log(
    'Triggers created!\n\n' +
    '• startDailyRun fires every day at 8am\n' +
    '• weeklyNotificationTrigger fires every Monday at 9am\n' +
    '• watchEnrollCheckbox fires on every spreadsheet edit\n\n' +
    'Verify in Extensions → Apps Script → Triggers.'
  );
}

// ============================================================
// SECTION: SETTINGS HELPER
// ============================================================

/**
 * Reads a value from the Settings sheet by setting name.
 * Returns the value as a trimmed string, or '' if not found.
 * Caches within a single script run using a closure cache object.
 */
const _settingsCache = {};

function getSetting(name) {
  if (_settingsCache[name] !== undefined) return _settingsCache[name];

  try {
    const sheet = getSheet(CONFIG.sheets.settings);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return '';
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (const row of data) {
      const key = String(row[0]).trim();
      const val = String(row[1]).trim();
      _settingsCache[key] = val; // cache everything at once
    }
    return _settingsCache[name] !== undefined ? _settingsCache[name] : '';
  } catch (e) {
    Logger.log('getSetting ERROR for "' + name + '": ' + e.message);
    return '';
  }
}

// ============================================================
// SECTION: TIME WINDOW & DATE HELPERS
// ============================================================

/**
 * Returns true if current time is within the configured sending window.
 * Uses timeapi.io via currentTimeHHMM().
 */
function isWithinSendingWindow() {
  const startTime = getSetting('followUpStartTime') || '08:00';
  const endTime   = getSetting('followUpEndTime')   || '18:00';
  const now       = currentTimeHHMM();
  // Normal window e.g. 08:00 - 22:00: start <= now <= end
  // Midnight-crossing window e.g. 17:00 - 02:00: now >= start OR now <= end
  if (startTime <= endTime) {
    return now >= startTime && now <= endTime;
  } else {
    return now >= startTime || now <= endTime;
  }
}

/**
 * Returns true if the sending window has not opened yet today.
 * Used by startDailyRun to schedule a delayed trigger for when the window opens.
 * Only relevant for non-midnight-crossing windows.
 */
function windowOpensLaterToday() {
  const startTime = getSetting('followUpStartTime') || '08:00';
  const endTime   = getSetting('followUpEndTime')   || '18:00';
  const now       = currentTimeHHMM(); // from timeapi.io - correct timezone
  // Normal window (e.g. 09:00-22:00): opens later if now < startTime
  // Midnight-crossing (e.g. 17:00-02:00): already open if now >= start OR now <= end
  // so "opens later" only applies to normal windows
  if (startTime <= endTime) {
    return now < startTime;
  }
  return false; // midnight-crossing: if we're outside it means we're between end and start
                // i.e. the window already closed, not that it opens later
}

/**
 * Returns true if today is a configured sending day.
 * followUpSendingDays is a comma-separated list of 3-letter day abbreviations, e.g. "Mon,Tue,Wed,Thu,Fri"
 * Uses timeapi.io via currentDayAbbr().
 */
function isSendingDay() {
  const sendingDaysRaw = getSetting('followUpSendingDays') || 'Mon,Tue,Wed,Thu,Fri';
  const todayDayAbbr   = currentDayAbbr(); // from timeapi.io
  const allowedDays    = sendingDaysRaw.split(',').map(d => d.trim());
  return allowedDays.some(d => d.toLowerCase() === todayDayAbbr.toLowerCase());
}

/**
 * Parses a YYYY-MM-DD string into a Date object.
 */
function parseDate(str) {
  const parts = String(str).split('-');
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
}

/**
 * Adds N days to a YYYY-MM-DD date string and returns a YYYY-MM-DD string.
 * Uses plain JS Date arithmetic - no timezone-sensitive formatting needed here
 * because we are only doing arithmetic on a date string, not reading the wall clock.
 */
function addDays(dateStr, n) {
  const d = parseDate(dateStr);
  d.setDate(d.getDate() + parseInt(n));
  // Format manually to avoid any timezone shift in Utilities.formatDate
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const dd   = String(d.getDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd;
}

/**
 * Returns true if date1 (YYYY-MM-DD) is on or before date2 (YYYY-MM-DD).
 */
function dateIsOnOrBefore(date1, date2) {
  return date1 <= date2;
}

// ============================================================
// SECTION: DAILY SEND COUNTER (PropertiesService)
// ============================================================

/**
 * Reads the today's send count from PropertiesService.
 * Resets to 0 if the stored date doesn't match today.
 * Returns current count as a number.
 */
function getSendsToday() {
  const props = PropertiesService.getScriptProperties();
  const storedDate  = props.getProperty(CONFIG.props.sendsTodayDate) || '';
  const today       = todayStr();
  if (storedDate !== today) {
    props.setProperty(CONFIG.props.sendsTodayDate, today);
    props.setProperty(CONFIG.props.sendsToday, '0');
    return 0;
  }
  return parseInt(props.getProperty(CONFIG.props.sendsToday) || '0');
}

/**
 * Increments the daily send counter by 1.
 */
function incrementSendsToday() {
  const props   = PropertiesService.getScriptProperties();
  const current = getSendsToday(); // also resets date if needed
  props.setProperty(CONFIG.props.sendsToday, String(current + 1));
}

/**
 * Returns true if we have hit the maxSendsPerDay cap.
 */
function hitMaxSendsPerDay() {
  const max     = parseInt(getSetting('maxSendsPerDay') || '50');
  const current = getSendsToday();
  return current >= max;
}

// ============================================================
// SECTION: CHAINED TRIGGER ENGINE
// ============================================================

/**
 * Entry point called by the 8am time-based trigger.
 * Performs all pre-flight checks, then starts the chain.
 */
function startDailyRun() {
  Logger.log('=== startDailyRun: ' + new Date().toISOString() + ' ===');

  // Reset caches for a fresh run
  for (const key in _settingsCache) delete _settingsCache[key];
  for (const key in _timeApiCache)  delete _timeApiCache[key];

  // Clean up any one-time retry triggers for startDailyRun (created by scheduleWindowOpenTrigger).
  // The recurring daily trigger was created with everyDays(1) — we keep exactly one
  // startDailyRun trigger total. Any extras beyond the first are retry triggers to delete.
  const _allT = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'startDailyRun');
  Logger.log('startDailyRun: found ' + _allT.length + ' startDailyRun trigger(s)');
  _allT.forEach(t => {
    // TriggerSource.CLOCK + everyDays = the recurring daily trigger - never delete it
    // after()-based triggers have no repeat interval - those are the retry ones to clean up
    try {
      const isRecurring = t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK &&
                          t.getEventType() === ScriptApp.EventType.CLOCK;
      // We can't directly distinguish everyDays vs after() from the trigger object,
      // so we use a stored property to mark the recurring trigger's UID at creation time
      const recurringUid = PropertiesService.getScriptProperties().getProperty('recurringDailyTriggerUid');
      if (recurringUid && t.getUniqueId() !== recurringUid) {
        ScriptApp.deleteTrigger(t);
        Logger.log('startDailyRun: deleted retry trigger ' + t.getUniqueId());
      } else if (!recurringUid) {
        Logger.log('startDailyRun: WARNING - recurringDailyTriggerUid not set, not deleting any triggers to be safe');
      }
    } catch (e) {
      Logger.log('startDailyRun: error checking trigger ' + t.getUniqueId() + ': ' + e.message);
    }
  });

  Logger.log('startDailyRun: pauseAllFollowUps=' + getSetting('pauseAllFollowUps'));

  // Check pauseAllFollowUps
  if (getSetting('pauseAllFollowUps').toUpperCase() === 'TRUE') {
    Logger.log('pauseAllFollowUps is TRUE - exiting without sending.');
    return;
  }

  Logger.log('startDailyRun: isSendingDay=' + isSendingDay() + ' | today=' + currentDayAbbr() + ' | allowedDays=' + getSetting('followUpSendingDays'));

  // Check sending day
  if (!isSendingDay()) {
    Logger.log('Today is not a configured sending day - exiting.');
    return;
  }

  Logger.log('startDailyRun: isWithinSendingWindow=' + isWithinSendingWindow() + ' | now=' + currentTimeHHMM() + ' | window=' + getSetting('followUpStartTime') + '-' + getSetting('followUpEndTime'));

  // Check time window
  if (!isWithinSendingWindow()) {
    if (windowOpensLaterToday() && isSendingDay()) {
      // Window hasn't opened yet today - schedule a one-time trigger to fire at window open time
      scheduleWindowOpenTrigger();
      Logger.log('Outside sending window - window opens later today, trigger scheduled.');
    } else {
      Logger.log('Outside sending window - exiting.');
    }
    return;
  }

  // Reset daily send counter if date rolled over
  getSendsToday(); // side-effect: resets counter if new day

  Logger.log('startDailyRun: sendsToday=' + getSendsToday() + ' | maxSendsPerDay=' + getSetting('maxSendsPerDay') + ' | hitMax=' + hitMaxSendsPerDay());

  // Check max sends
  if (hitMaxSendsPerDay()) {
    Logger.log('maxSendsPerDay already reached - exiting.');
    return;
  }

  // Check OOO auto-resume
  checkOooAutoResume();

  // Record sends count at chain start so processOneLead can detect if THIS run sent anything
  PropertiesService.getScriptProperties().setProperty('sendsAtChainStart', String(getSendsToday()));

  const _startT = String(getSetting('followUpStartTime')).substring(0, 5);
  const _endT   = String(getSetting('followUpEndTime')).substring(0, 5);
  Logger.log('startDailyRun: pre-flight passed | sending window: ' + _startT + '-' + _endT + ' | sendsToday: ' + getSendsToday() + '/' + getSetting('maxSendsPerDay'));

  // Count how many leads are due today (for logging)
  Logger.log('startDailyRun: scanning for due leads to schedule...');
  const activeSheet = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow = activeSheet.getLastRow();
  let dueCount = 0;
  let firstLeadEmail = '';
  let firstLeadRow = -1;

  if (lastRow >= 2) {
    const today = todayStr();
    const data = activeSheet.getRange(2, 1, lastRow - 1, CONFIG.cols.lastSentDate).getValues();
    const c = CONFIG.cols;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const status = String(row[c.status - 1]).trim();
      if (status !== 'Active') continue;

      let nextSendDate = '';
      const nextRaw = row[c.nextSendDate - 1];
      if (nextRaw instanceof Date) {
        nextSendDate = nextRaw.getFullYear() + '-' +
          String(nextRaw.getMonth() + 1).padStart(2, '0') + '-' +
          String(nextRaw.getDate()).padStart(2, '0');
      } else {
        nextSendDate = String(nextRaw || '').trim();
      }
      if (!nextSendDate || !dateIsOnOrBefore(nextSendDate, today)) continue;

      let lastSentDate = '';
      const lastRaw = row[c.lastSentDate - 1];
      if (lastRaw instanceof Date) {
        lastSentDate = lastRaw.getFullYear() + '-' +
          String(lastRaw.getMonth() + 1).padStart(2, '0') + '-' +
          String(lastRaw.getDate()).padStart(2, '0');
      } else {
        lastSentDate = String(lastRaw || '').trim();
      }
      if (lastSentDate === today) continue;

      dueCount++;
      if (dueCount === 1) {
        firstLeadEmail = String(row[c.leadEmail - 1]);
        firstLeadRow = i + 2;
        Logger.log('startDailyRun: first due lead = ' + firstLeadEmail +
          ' | sequence: ' + String(row[c.sequenceName - 1]) +
          ' | step: ' + String(row[c.sequenceStep - 1]) +
          ' | nextSendDate: ' + nextSendDate +
          ' | row: ' + firstLeadRow);
      }
    }
  }

  Logger.log('startDailyRun: found ' + dueCount + ' lead(s) due today - NOT sending directly, scheduling processOneLead chain instead');

  if (dueCount === 0) {
    Logger.log('startDailyRun: no leads due today - chain not started.');
    return;
  }

  // Record chain start send count for notification logic
  PropertiesService.getScriptProperties().setProperty('sendsAtChainStart', String(getSendsToday()));
  Logger.log('startDailyRun: sendsAtChainStart set to ' + getSendsToday());

  // Schedule the first processOneLead - it will chain itself from there
  scheduleNextProcessOneLead();
  Logger.log('startDailyRun: first processOneLead trigger scheduled - chain will begin shortly');
}

/**
 * Called by chained time-based triggers. e contains triggerUid.
 * Deletes its own trigger first, then processes one lead, then reschedules.
 */
function processOneLead(e) {
  Logger.log('processOneLead: triggerUid=' + (e ? e.triggerUid : 'manual') + ' | sendsToday=' + getSendsToday() + ' | maxSends=' + getSetting('maxSendsPerDay'));

  // Step 1: Delete this trigger immediately to prevent orphan buildup
  if (e && e.triggerUid) {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getUniqueId() === e.triggerUid) {
        ScriptApp.deleteTrigger(t);
        Logger.log('Deleted self-trigger: ' + e.triggerUid);
      }
    });
  }

  // Reset caches for fresh read
  for (const key in _settingsCache) delete _settingsCache[key];
  for (const key in _timeApiCache)  delete _timeApiCache[key];

  // Check time window - stop chain if past end time
  if (!isWithinSendingWindow()) {
    Logger.log('Outside sending window - stopping chain.');
    return;
  }

  // Check pause flag
  if (getSetting('pauseAllFollowUps').toUpperCase() === 'TRUE') {
    Logger.log('pauseAllFollowUps is TRUE - stopping chain.');
    return;
  }

  // Check max sends cap
  if (hitMaxSendsPerDay()) {
    Logger.log('maxSendsPerDay reached - stopping chain.');
    return;
  }

  // Step 2: Find next due lead
  const lead = findNextDueLead();

  // Step 3: No lead found → end of chain
  if (!lead) {
    Logger.log('All leads processed for today. Chain complete.');
    Logger.log('processOneLead: chain ended | sendsAtChainStart=' + PropertiesService.getScriptProperties().getProperty('sendsAtChainStart') + ' | sendsNow=' + getSendsToday() + ' | sentThisRun=' + (getSendsToday() - parseInt(PropertiesService.getScriptProperties().getProperty('sendsAtChainStart') || '0')));
    // Only send daily notification if THIS chain run actually sent at least one email.
    // We compare sendsToday now vs when the chain started (stored in script properties).
    const props = PropertiesService.getScriptProperties();
    const sendsAtChainStart = parseInt(props.getProperty('sendsAtChainStart') || '0');
    const sendsNow = getSendsToday();
    const sentThisRun = sendsNow - sendsAtChainStart;
    if (sentThisRun > 0 && (getSetting('notificationFrequency') || 'daily').toLowerCase() === 'daily') {
      Logger.log('Chain complete: ' + sentThisRun + ' email(s) sent this run - sending notification.');
      sendNotificationEmail();
    } else {
      Logger.log('Chain complete: 0 emails sent this run (total today: ' + sendsNow + ') - skipping notification.');
    }
    return;
  }

  // Step 4: Process this lead
  Logger.log('Processing lead: ' + lead.leadEmail + ' (row ' + lead.sheetRow + ')');
  try {
    processSingleLead(lead);
  } catch (e) {
    // Uncaught exception inside processSingleLead - log it and notify but keep chain alive
    Logger.log('processSingleLead UNCAUGHT ERROR for ' + lead.leadEmail + ': ' + e.message);
    const notifEmail = getSetting('notificationEmail');
    if (notifEmail) {
      try {
        MailApp.sendEmail({
          to: notifEmail,
          subject: 'Follow-Up System - unexpected error, chain continuing',
          body: 'An unexpected error occurred processing lead: ' + lead.leadEmail +
            '\nRow: ' + lead.sheetRow + '\nError: ' + e.message +
            '\n\nThe chain is continuing to the next lead. Check this lead manually.'
        });
      } catch (ne) { Logger.log('Could not send error notification: ' + ne.message); }
    }
  }

  // Step 5: Schedule the next trigger in the chain
  scheduleNextProcessOneLead();
}

/**
 * Manual wrapper so the menu item works without an event object.
 */
function manualProcessOneLead() {
  processOneLead(null);
}

/**
 * Schedules a one-time trigger to call startDailyRun() at the configured
 * followUpStartTime so the chain fires automatically when the window opens,
 * even though the 8am daily trigger ran too early.
 *
 * Deletes any existing window-open trigger first to avoid duplicates.
 * The trigger calls startDailyRun() which is idempotent - safe to call multiple times.
 */
function scheduleWindowOpenTrigger() {
  // Uses timeapi.io to get the EXACT current time in the configured timezone,
  // calculates how many milliseconds until followUpStartTime,
  // and schedules startDailyRun() to fire at that precise moment.
  //
  // This is the correct approach — no new Date() timezone math which uses the
  // Apps Script server timezone instead of the user's configured timezone.
  //
  // Deduplication: if a retry trigger already exists (>1 startDailyRun triggers)
  // we skip to avoid stacking.
  const existingDailyRun = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'startDailyRun');
  if (existingDailyRun.length > 1) {
    Logger.log('scheduleWindowOpenTrigger: retry already scheduled - skipping duplicate');
    return;
  }

  const startTime = getSetting('followUpStartTime') || '08:00';
  const tz        = getSetting('timezone') || 'UTC';
  const parts     = startTime.split(':');
  const targetHour   = parseInt(parts[0], 10);
  const targetMinute = parseInt(parts[1] || '0', 10);

  // Ask timeapi.io what the current time is in the configured timezone
  // Use the /time/current/zone endpoint which returns hour, minute, seconds as integers
  let diffMs;
  const MAX_RETRIES = 3;
  let fetched = false;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const url      = 'https://timeapi.io/api/time/current/zone?timeZone=' + encodeURIComponent(tz);
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json     = JSON.parse(response.getContentText());

      const nowHour   = parseInt(json.hour,    10);
      const nowMinute = parseInt(json.minute,  10);
      const nowSecond = parseInt(json.seconds, 10);

      const nowTotalMinutes    = nowHour * 60 + nowMinute;
      const targetTotalMinutes = targetHour * 60 + targetMinute;

      let diffMinutes = targetTotalMinutes - nowTotalMinutes;
      // If target is "earlier" in the day than now, it means it's tomorrow
      // (shouldn't happen here since windowOpensLaterToday() confirmed window is later today,
      // but guard anyway)
      if (diffMinutes < 0) diffMinutes += 24 * 60;

      // Subtract seconds already elapsed in current minute for precision
      diffMs = (diffMinutes * 60 - nowSecond) * 1000;

      Logger.log('scheduleWindowOpenTrigger: tz=' + tz +
        ' now=' + String(nowHour).padStart(2,'0') + ':' + String(nowMinute).padStart(2,'0') +
        ' target=' + startTime +
        ' diffMinutes=' + diffMinutes +
        ' diffMs=' + diffMs +
        ' (attempt ' + attempt + ')');

      fetched = true;
      break;
    } catch (e) {
      Logger.log('scheduleWindowOpenTrigger: timeapi attempt ' + attempt + ' failed - ' + e.message);
      if (attempt < MAX_RETRIES) Utilities.sleep(2000);
    }
  }

  if (!fetched || isNaN(diffMs) || diffMs <= 0) {
    // Fallback: if API failed or time already passed, retry in 15 minutes
    diffMs = 15 * 60 * 1000;
    Logger.log('scheduleWindowOpenTrigger: could not calculate exact delay - falling back to 15 min retry');
  }

  ScriptApp.newTrigger('startDailyRun')
    .timeBased()
    .after(diffMs)
    .create();

  const minsUntil = Math.round(diffMs / 60000);
  Logger.log('scheduleWindowOpenTrigger: startDailyRun scheduled in ' + minsUntil + ' min (window opens at ' + startTime + ' ' + tz + ')');
}

/**
 * Schedules the next processOneLead() trigger with a random delay from Settings.
 */
function scheduleNextProcessOneLead() {
  const minMin = parseInt(getSetting('minDelayMinutes') || '4');
  const maxMin = parseInt(getSetting('maxDelayMinutes') || '8');
  const delayMs = Math.floor(Math.random() * (maxMin - minMin + 1) + minMin) * 60 * 1000;
  const delayMinutes = Math.round(delayMs / 60000);

  ScriptApp.newTrigger('processOneLead')
    .timeBased()
    .after(delayMs)
    .create();

  Logger.log('scheduleNextProcessOneLead: delayMs=' + delayMs + ' | delayMinutes=' + delayMinutes + ' | minMin=' + minMin + ' | maxMin=' + maxMin);
  Logger.log('scheduleNextProcessOneLead: trigger will fire at ~' + new Date(Date.now() + delayMs).toISOString());
}

/**
 * Scans ActiveFollowUps and returns the first row where:
 *   status = Active AND nextSendDate <= today AND lastSentDate != today
 * Returns a lead object or null.
 */
function findNextDueLead() {
  let activeSheet;
  try {
    activeSheet = getSheet(CONFIG.sheets.activeFollowUps);
  } catch (e) {
    Logger.log('findNextDueLead: could not open ActiveFollowUps sheet - ' + e.message);
    const notifEmail = getSetting('notificationEmail');
    if (notifEmail) {
      try {
        MailApp.sendEmail({
          to: notifEmail,
          subject: 'Follow-Up System - ActiveFollowUps sheet not found, chain stopped',
          body: 'The ActiveFollowUps sheet could not be found. The chain has stopped.\n\n' +
            'Error: ' + e.message + '\n\nCheck the sheet name has not been renamed.'
        });
      } catch (ne) {}
    }
    return null;
  }
  const lastRow = activeSheet.getLastRow();
  if (lastRow < 2) return null;

  const numRows  = lastRow - 1;
  const lastCol  = activeSheet.getLastColumn();
  const data     = activeSheet.getRange(2, 1, numRows, lastCol).getValues();
  // Read headers row once so we can build customVars for columns beyond the fixed 13
  const headers  = activeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const today    = todayStr();
  const c        = CONFIG.cols;

  Logger.log('findNextDueLead: scanning ' + data.length + ' rows in ActiveFollowUps, today=' + today);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const status = String(row[c.status - 1]).trim();
    const rowEmail = String(row[c.leadEmail - 1]).trim();
    const rowStatus = String(row[c.status - 1]).trim();
    if (rowEmail || rowStatus) {
      Logger.log('findNextDueLead: row ' + (i+2) + ' | email=' + rowEmail + ' | status=' + rowStatus);
    }
    if (status !== 'Active') continue;

    let nextSendDate = '';
    const nextRaw = row[c.nextSendDate - 1];
    if (nextRaw instanceof Date) {
      // Convert the Date object to YYYY-MM-DD using plain JS (no tz shift needed - already a date value)
      const d    = nextRaw;
      const yyyy = d.getFullYear();
      const mm   = String(d.getMonth() + 1).padStart(2, '0');
      const dd   = String(d.getDate()).padStart(2, '0');
      nextSendDate = yyyy + '-' + mm + '-' + dd;
    } else {
      nextSendDate = String(nextRaw || '').trim();
    }
    if (!nextSendDate || !dateIsOnOrBefore(nextSendDate, today)) {
      Logger.log('findNextDueLead: row ' + (i+2) + ' skipped - nextSendDate=' + nextSendDate + ' is in the future');
      continue;
    }

    let lastSentDate = '';
    const lastRaw = row[c.lastSentDate - 1];
    if (lastRaw instanceof Date) {
      const d    = lastRaw;
      const yyyy = d.getFullYear();
      const mm   = String(d.getMonth() + 1).padStart(2, '0');
      const dd   = String(d.getDate()).padStart(2, '0');
      lastSentDate = yyyy + '-' + mm + '-' + dd;
    } else {
      lastSentDate = String(lastRaw || '').trim();
    }
    if (lastSentDate === today) {
      Logger.log('findNextDueLead: row ' + (i+2) + ' skipped - already sent today (' + lastSentDate + ')');
      continue; // already processed today
    }
    Logger.log('findNextDueLead: FOUND due lead at row ' + (i+2) + ' | email=' + String(row[c.leadEmail-1]) + ' | nextSendDate=' + nextSendDate + ' | lastSentDate=' + lastSentDate + ' | step=' + String(row[c.sequenceStep-1]) + '/' + String(row[c.totalSteps-1]));

    // Build customVars from any columns beyond the fixed 13 (notes is col 13)
    // Key = header name (e.g. 'companyName'), value = cell value as string
    const customVars = {};
    for (let col = c.notes; col < headers.length; col++) {
      const headerName = String(headers[col]).trim();
      if (headerName) customVars[headerName] = String(row[col] || '').trim();
    }

    // Found a due lead
    return {
      sheetRow:      i + 2, // 1-based sheet row (row 1 = header)
      dataIndex:     i,
      leadName:      String(row[c.leadName - 1]),
      leadEmail:     String(row[c.leadEmail - 1]),
      threadId:      String(row[c.threadId - 1]),
      fromAccount:   String(row[c.fromAccount - 1]),
      sequenceName:  String(row[c.sequenceName - 1]),
      sequenceStep:  Number(row[c.sequenceStep - 1]) || 0,
      totalSteps:    Number(row[c.totalSteps - 1]) || 0,
      nextSendDate:  nextSendDate,
      resumeOnReply: row[c.resumeOnReply - 1],
      notes:         String(row[c.notes - 1] || ''),
      customVars:    customVars,
    };
  }

  return null;
}

/**
 * Processes a single lead object (found by findNextDueLead).
 * Runs reply detection, OOO check, sends the message, updates the sheet row.
 */
function processSingleLead(lead) {
  const activeSheet = getSheet(CONFIG.sheets.activeFollowUps);
  const today = todayStr();
  const c = CONFIG.cols;

  const { sheetRow, leadName, leadEmail, threadId, fromAccount,
          sequenceName, sequenceStep, totalSteps, customVars } = lead;

  Logger.log('processSingleLead START: email=' + leadEmail + ' | seq=' + sequenceName + ' | step=' + sequenceStep + '/' + totalSteps + ' | row=' + sheetRow);

  // Helper: write a single cell — throws on failure so callers can catch it
  function writeCell(col, value) {
    try {
      activeSheet.getRange(sheetRow, col).setValue(value);
    } catch (e) {
      throw new Error('writeCell col=' + col + ' value=' + value + ': ' + e.message);
    }
  }

  // Helper: append to notes (never overwrite)
  function appendNote(text) {
    const existing = activeSheet.getRange(sheetRow, c.notes).getValue();
    const updated  = existing ? existing + '\n' + text : text;
    writeCell(c.notes, updated);
  }

  // ── Step 3: Fetch the Gmail thread ────────────────────────
  let thread;
  try {
    thread = GmailApp.getThreadById(threadId);
  } catch (e) {
    Logger.log('Thread fetch error for ' + leadEmail + ': ' + e.message);
    writeCell(c.status, 'Error');
    appendNote('[' + today + '] Thread fetch error: ' + e.message);
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Failed');
    return;
  }

  if (!thread) {
    Logger.log('Thread not found for ' + leadEmail);
    writeCell(c.status, 'Error');
    appendNote('[' + today + '] Thread not found');
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Failed');
    return;
  }

  // ── Step 4: Get messages ──────────────────────────────────
  let messages;
  try {
    messages = thread.getMessages();
  } catch (e) {
    Logger.log('Error reading messages for ' + leadEmail + ': ' + e.message);
    writeCell(c.status, 'Error');
    appendNote('[' + today + '] Error reading messages: ' + e.message);
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Failed');
    return;
  }

  const latestMessage = messages[messages.length - 1];
  const latestSender  = latestMessage.getFrom();

  // ── Step 5: Reply detection (if enabled in Settings) ──────
  const replyCheckEnabled = getSetting('replyCheckBeforeSend').toUpperCase() !== 'FALSE';
  Logger.log('processSingleLead: replyCheckEnabled=' + replyCheckEnabled + ' | latestSender=' + latestSender + ' | isSendingAccount=' + isSendingAccount(latestSender));
  if (replyCheckEnabled && !isSendingAccount(latestSender)) {
    let replyBody = '';
    try { replyBody = stripHtml(latestMessage.getBody()); }
    catch (e) { replyBody = latestMessage.getPlainBody() || ''; }
    const snippet = replyBody.substring(0, 120);
    Logger.log(leadEmail + ': lead replied - setting Replied');
    writeCell(c.status, 'Replied');
    writeCell(c.lastReplySnippet, snippet);
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Skipped - Lead replied');
    SpreadsheetApp.flush();
    return;
  }

  Logger.log('processSingleLead: reply check passed for ' + leadEmail + ' - latest sender: ' + latestSender);

  // ── Step 6: OOO / bounce detection ────────────────────────
  let latestBody = '';
  try { latestBody = stripHtml(latestMessage.getBody()).toLowerCase(); }
  catch (e) { latestBody = (latestMessage.getPlainBody() || '').toLowerCase(); }

  const oooDetected = CONFIG.oooKeywords.some(kw => latestBody.includes(kw));
  Logger.log('processSingleLead: oooDetected=' + oooDetected + ' | bodyPreview=' + latestBody.substring(0, 80));
  if (oooDetected) {
    Logger.log(leadEmail + ': OOO/bounce detected');
    writeCell(c.status, 'OOO');
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Skipped - OOO/bounce detected');
    SpreadsheetApp.flush();
    return;
  }

  // ── Step 7: Determine step to send ────────────────────────
  const currentStep = sequenceStep; // 0-indexed

  // ── Step 8: Check totalSteps cap ──────────────────────────
  Logger.log('processSingleLead: currentStep=' + currentStep + ' | totalSteps=' + totalSteps + ' | hitsCap=' + (totalSteps > 0 && currentStep >= totalSteps));
  if (totalSteps > 0 && currentStep >= totalSteps) {
    Logger.log(leadEmail + ': Done - max steps reached');
    writeCell(c.status, 'Done');
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Skipped - Max steps reached');
    SpreadsheetApp.flush();
    return;
  }

  // ── Step 9: Check sequence has this step ──────────────────
  let stepData;
  try {
    stepData = getSequenceStep(sequenceName, currentStep + 1);
  } catch (e) {
    Logger.log(leadEmail + ': Sequence read error - ' + e.message);
    writeCell(c.status, 'Error');
    appendNote('[' + today + '] Sequence read error: ' + e.message);
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Failed');
    SpreadsheetApp.flush();
    return;
  }
  Logger.log('processSingleLead: stepData found=' + !!stepData + ' | awaitDays=' + (stepData ? stepData.awaitDays : 'N/A') + ' | msgLength=' + (stepData ? stepData.message.length : 0));
  if (!stepData || !stepData.message) {
    Logger.log(leadEmail + ': Done - sequence exhausted at step ' + currentStep);
    writeCell(c.status, 'Done');
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Skipped - Sequence exhausted');
    SpreadsheetApp.flush();
    return;
  }

  Logger.log('processSingleLead: step ' + currentStep + ' found in sequence - awaitDays=' + stepData.awaitDays + ' | messageLength=' + stepData.message.length);

  // ── Step 10: Parse message content & replace variables ───────────────────
  const parsed   = parseMessage(stepData.message);
  const imgUrl   = parsed.imgUrl;
  const imgSize  = parsed.imgSize; // CSS width string e.g. '50%', '300px', '100%'
  // Replace {{variableName}} tokens with lead data.
  // Built-in: {{leadName}}, {{leadEmail}}, {{firstName}} (first word of leadName).
  // Custom: any column header added to ActiveFollowUps beyond col 13.
  const textPart = replaceVariables(parsed.text, leadName, leadEmail, customVars);
  Logger.log(leadEmail + ': variables replaced - preview: ' + textPart.substring(0, 80) + (imgUrl ? ' | imgSize: ' + imgSize : ''));

  // ── Step 11: Send the reply ────────────────────────────────
  // imgBlob is passed to sendReplyFromAlias which embeds it inline in the HTML body.
  // If no image, imgBlob is null and a plain HTML email is sent.
  let imgBlob = null;
  if (imgUrl) {
    try {
      imgBlob = getImageBlob(imgUrl);
      Logger.log(leadEmail + ': image blob fetched - ' + (imgBlob ? imgBlob.getName() : 'null'));
    } catch (e) {
      Logger.log(leadEmail + ': WARNING - image fetch failed: ' + e.message);
      appendNote('[' + today + '] WARNING: image fetch failed at step ' + currentStep + ': ' + e.message);
    }
  }
  Logger.log('processSingleLead: ABOUT TO SEND | email=' + leadEmail + ' | step=' + currentStep + ' | fromAccount=' + fromAccount);
  try {
    sendReplyRaw(thread, leadEmail, textPart, imgBlob, imgSize);
  } catch (e) {
    Logger.log(leadEmail + ': SEND FAILED - ' + e.message);
    const isQuota = e.message && (
      e.message.toLowerCase().includes('quota') ||
      e.message.toLowerCase().includes('rate') ||
      e.message.toLowerCase().includes('limit')
    );
    writeCell(c.status, 'Error');
    appendNote('[' + today + '] Send failed at step ' + currentStep + ': ' + e.message);
    logToSendLog(leadEmail, leadName, sequenceName, currentStep, fromAccount,
      textPart.substring(0, 100), threadId, 'Failed');
    SpreadsheetApp.flush();
    // If quota hit, notify immediately and stop the chain
    if (isQuota) {
      Logger.log('QUOTA HIT - stopping chain and sending notification');
      const notifEmail = getSetting('notificationEmail');
      if (notifEmail) {
        try {
          MailApp.sendEmail({
            to: notifEmail,
            subject: 'Follow-Up System - Gmail quota hit, chain stopped',
            body: 'The follow-up chain stopped because Gmail quota was hit.\n' +
              'Last lead: ' + leadEmail + ' (step ' + currentStep + ')\n' +
              'Error: ' + e.message + '\n' +
              'All unsent leads remain Active and will be picked up tomorrow.'
          });
        } catch (ne) { Logger.log('Could not send quota notification: ' + ne.message); }
      }
      return; // stop chain — don't schedule next trigger
    }
    return;
  }

  // ── Step 12: On successful send ────────────────────────────
  // Write order matters for duplicate-send safety:
  //   lastSentDate FIRST  - findNextDueLead checks this; if set, lead is skipped today
  //   nextSendDate SECOND - schedules tomorrow correctly
  //   sequenceStep LAST   - if we crash between send and this write, lastSentDate
  //                         already guards against a re-send today; sequenceStep being
  //                         stale is recoverable manually; a double-send is not.
  const newStep     = currentStep + 1;
  const newSendDate = addDays(today, stepData.awaitDays);

  try {
    writeCell(c.lastSentDate, today);
    writeCell(c.nextSendDate, newSendDate);
    writeCell(c.sequenceStep, newStep);
  } catch (e) {
    // Sheet write failed after email was already sent.
    // Log clearly so the user can manually fix sequenceStep in the sheet.
    Logger.log(leadEmail + ': CRITICAL - email sent but sheet write failed: ' + e.message +
      ' | manually set sequenceStep=' + newStep + ' lastSentDate=' + today + ' nextSendDate=' + newSendDate);
    appendNote('[' + today + '] CRITICAL: email sent (step ' + currentStep + ') but sheet write failed. ' +
      'Manually set sequenceStep=' + newStep + ' lastSentDate=' + today + ' nextSendDate=' + newSendDate);
    logToSendLog(leadEmail, leadName, sequenceName, currentStep, fromAccount,
      textPart.substring(0, 100), threadId, 'Sent-WriteError');
    SpreadsheetApp.flush();
    return;
  }

  logToSendLog(leadEmail, leadName, sequenceName, currentStep, fromAccount,
    textPart.substring(0, 100), threadId, 'Sent');

  incrementSendsToday();

  Logger.log('processSingleLead: SEND SUCCESS | email=' + leadEmail + ' | step=' + currentStep + ' → ' + newStep + ' | nextSendDate=' + newSendDate + ' | sendsToday=' + getSendsToday());

  // ── Step 13: Peek at next step - mark Done immediately if sequence is over ──
  // This avoids the lead sitting as Active overnight just to be marked Done the next morning.
  // Check totalSteps cap first, then check if the sequence sheet has a next row.
  let isDoneNow = false;
  if (totalSteps > 0 && newStep >= totalSteps) {
    isDoneNow = true;
    Logger.log(leadEmail + ': sequence complete - hit totalSteps cap (' + totalSteps + ') - marking Done now');
  } else {
    try {
      const nextStepData = getSequenceStep(sequenceName, newStep + 1);
      if (!nextStepData || !nextStepData.message) {
        isDoneNow = true;
        Logger.log(leadEmail + ': sequence complete - no step ' + (newStep + 1) + ' in sheet - marking Done now');
      }
    } catch (e) {
      Logger.log(leadEmail + ': could not peek at next step - leaving Active: ' + e.message);
    }
  }

  if (isDoneNow) {
    writeCell(c.status, 'Done');
    Logger.log(leadEmail + ': status set to Done');
  }

  SpreadsheetApp.flush();
}

// ============================================================
// SECTION: OOO AUTO-RESUME
// ============================================================

/**
 * Called at the start of startDailyRun().
 * If oooAutoResume is TRUE, finds all OOO leads where the last update was
 * >= oooResumeDays ago and flips them back to Active.
 */
function checkOooAutoResume() {
  if (getSetting('oooAutoResume').toUpperCase() !== 'TRUE') return;

  const resumeDays  = parseInt(getSetting('oooResumeDays') || '7');
  const activeSheet = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow     = activeSheet.getLastRow();
  if (lastRow < 2) return;

  const data  = activeSheet.getRange(2, 1, lastRow - 1, CONFIG.cols.notes).getValues();
  const today = todayStr();
  const c     = CONFIG.cols;

  data.forEach((row, i) => {
    if (String(row[c.status - 1]).trim() !== 'OOO') return;

    // Use the date the OOO status was SET, not lastSentDate.
    // Parse the most recent [YYYY-MM-DD] OOO note entry from the notes column.
    // Falls back to lastSentDate if no note found.
    let oooSetDate = '';
    const notes = String(row[c.notes - 1] || '');
    const oooNoteMatch = notes.match(/\[(\d{4}-\d{2}-\d{2})\][^\n]*OOO/i);
    if (oooNoteMatch) {
      oooSetDate = oooNoteMatch[1];
    } else {
      // Fallback: use lastSentDate
      const lastRaw = row[c.lastSentDate - 1];
      if (lastRaw instanceof Date) {
        const d  = lastRaw;
        oooSetDate = d.getFullYear() + '-' +
          String(d.getMonth() + 1).padStart(2, '0') + '-' +
          String(d.getDate()).padStart(2, '0');
      } else {
        oooSetDate = String(lastRaw || '').trim();
      }
    }

    if (!oooSetDate) return; // can't determine when OOO was set
    const lastSent = oooSetDate;

    const resumeDate = addDays(lastSent, resumeDays);
    if (today >= resumeDate) {
      const sheetRow = i + 2;
      activeSheet.getRange(sheetRow, c.status).setValue('Active');
      // Set nextSendDate to today so it gets picked up in this run
      activeSheet.getRange(sheetRow, c.nextSendDate).setValue(today);
      const existing = activeSheet.getRange(sheetRow, c.notes).getValue();
      const note = '[' + today + '] Auto-resumed from OOO after ' + resumeDays + ' days.';
      activeSheet.getRange(sheetRow, c.notes).setValue(existing ? existing + '\n' + note : note);
      Logger.log('OOO auto-resumed: ' + String(row[c.leadEmail - 1]));
    }
  });

  SpreadsheetApp.flush();
}

// ============================================================
// SECTION: TEST SEND
// ============================================================

/**
 * Sends a preview of step 1 of the first Active lead's sequence to notificationEmail.
 * Nothing is written to the sheet. No real email goes to the lead.
 * Use this to verify message text, variable replacement, image size, and layout
 * before running the real chain.
 *
 * HOW TO USE:
 *   1. Make sure the first row in ActiveFollowUps has a lead with status=Active
 *      and a sequenceName filled in.
 *   2. Make sure notificationEmail is set in Settings.
 *   3. Run Follow-Up System → Send test email (first lead, step 1)
 *   4. Check your notificationEmail inbox - the email shows exactly what the lead
 *      would receive, with all variables replaced and the image rendered inline.
 */
function sendTestEmail() {
  const recipient = getSetting('notificationEmail');
  if (!recipient) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Set notificationEmail in Settings first.', 'Test Send', 6);
    return;
  }

  // Ask which step number to preview
  const ui       = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Test Email - Choose Step',
    'Enter the step number to preview (1 = first message, 2 = second, etc):',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const stepToPreview = parseInt(response.getResponseText().trim());
  if (isNaN(stepToPreview) || stepToPreview < 1) {
    ui.alert('Invalid step number. Enter a number like 1, 2, or 3.');
    return;
  }

  // Grab the first row in ActiveFollowUps that has a sequenceName - no status check,
  // no sequenceStep check - just use it as the variable source for the preview.
  const activeSheet = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow     = activeSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No leads in ActiveFollowUps.', 'Test Send', 6);
    return;
  }

  const lastCol = activeSheet.getLastColumn();
  const headers = activeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data    = activeSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const c       = CONFIG.cols;

  let lead = null;
  for (let i = 0; i < data.length; i++) {
    const seqName = String(data[i][c.sequenceName - 1]).trim();
    if (!seqName) continue; // skip rows with no sequence set
    const customVars = {};
    for (let col = c.notes; col < headers.length; col++) {
      const h = String(headers[col]).trim();
      if (h) customVars[h] = String(data[i][col] || '').trim();
    }
    lead = {
      leadName:     String(data[i][c.leadName     - 1]),
      leadEmail:    String(data[i][c.leadEmail    - 1]),
      sequenceName: seqName,
      customVars:   customVars,
    };
    break;
  }

  if (!lead) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'No lead with a sequenceName found in ActiveFollowUps.', 'Test Send', 6
    );
    return;
  }

  // Pull the requested step directly - no sequenceStep offset, user asked for exact step number
  let stepData;
  try {
    stepData = getSequenceStep(lead.sequenceName, stepToPreview);
  } catch (e) {
    ui.alert('Sequence error: ' + e.message);
    return;
  }

  if (!stepData || !stepData.message) {
    ui.alert('Step ' + stepToPreview + ' does not exist in sequence "' + lead.sequenceName + '". Check your Sequences sheet has that many rows.');
    return;
  }

  // Parse, replace variables, fetch image - identical to real send path
  const parsed   = parseMessage(stepData.message);
  const imgUrl   = parsed.imgUrl;
  const imgSize  = parsed.imgSize;
  const textPart = replaceVariables(parsed.text, lead.leadName, lead.leadEmail, lead.customVars);

  let imgBlob = null;
  if (imgUrl) {
    try {
      imgBlob = getImageBlob(imgUrl);
    } catch (e) {
      Logger.log('sendTestEmail: image fetch failed - ' + e.message);
    }
  }

  // Yellow banner so it is obvious this is a test
  const banner =
    '<div style="background:#fff3cd;border:1px solid #ffc107;padding:10px 14px;' +
    'margin-bottom:16px;font-family:sans-serif;font-size:13px;border-radius:4px">' +
    '<b>TEST PREVIEW - Step ' + stepToPreview + '</b><br>' +
    'Sequence: <b>' + lead.sequenceName + '</b><br>' +
    'Variables pulled from lead: <b>' + lead.leadEmail + '</b> (' + lead.leadName + ')<br>' +
    'awaitDays: ' + stepData.awaitDays + ' | image: ' + (imgUrl ? (imgSize + ' - ' + imgUrl.substring(0, 60) + '...') : 'none') +
    '</div>';

  const htmlBody  = banner + textPart.replace(/\n/g, '<br>');
  let   finalHtml = htmlBody;
  const subject   = '[TEST] Step ' + stepToPreview + ' | ' + lead.sequenceName + ' | ' + lead.leadEmail;

  if (imgBlob) {
    const cid       = 'test-inline-' + stepToPreview;
    const boundary  = 'testbnd_' + Utilities.getUuid().replace(/-/g, '');
    const imgBase64 = Utilities.base64Encode(imgBlob.getBytes());
    const mimeType  = imgBlob.getContentType() || 'image/jpeg';
    const sizeStyle = imgSize ? 'width:' + imgSize + ';max-width:100%' : 'max-width:100%';
    const imgTag    = '<br><img src="cid:' + cid + '" style="' + sizeStyle + '"><br>';

    finalHtml = finalHtml.includes('{{IMG_PLACEHOLDER}}')
      ? finalHtml.replace('{{IMG_PLACEHOLDER}}', imgTag)
      : finalHtml + imgTag;

    const rawEmail = [
      'MIME-Version: 1.0',
      'From: ' + Session.getActiveUser().getEmail(),
      'To: ' + recipient,
      'Subject: ' + subject,
      'Content-Type: multipart/related; boundary="' + boundary + '"',
      '',
      '--' + boundary,
      'Content-Type: text/html; charset=UTF-8',
      '',
      finalHtml,
      '',
      '--' + boundary,
      'Content-Type: ' + mimeType + '; name="image"',
      'Content-Transfer-Encoding: base64',
      'Content-ID: <' + cid + '>',
      'Content-Disposition: inline; filename="image"',
      '',
      imgBase64,
      '',
      '--' + boundary + '--'
    ].join('\r\n');

    try {
      Gmail.Users.Messages.send({ raw: Utilities.base64EncodeWebSafe(rawEmail) }, 'me');
      Logger.log('sendTestEmail: step ' + stepToPreview + ' sent with image to ' + recipient);
    } catch (e) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Send failed: ' + e.message, 'Test Send', 8);
      return;
    }
  } else {
    try {
      MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: finalHtml });
      Logger.log('sendTestEmail: step ' + stepToPreview + ' sent (no image) to ' + recipient);
    } catch (e) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Send failed: ' + e.message, 'Test Send', 8);
      return;
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Test for step ' + stepToPreview + ' sent to ' + recipient + ' - check your inbox.',
    'Test Send ✓', 8
  );
}

// ============================================================
// SECTION: NOTIFICATION EMAILS
// ============================================================

/**
 * Sends a summary notification email based on notificationFrequency setting.
 * Called automatically at chain end (daily) or by weeklyNotificationTrigger.
 * If notificationEmail is blank, does nothing.
 */
function sendNotificationEmail() {
  const recipient = getSetting('notificationEmail');
  if (!recipient) {
    Logger.log('sendNotificationEmail: no notificationEmail set - skipping.');
    return;
  }

  const frequency = (getSetting('notificationFrequency') || 'daily').toLowerCase();
  const today     = todayStr();

  // Determine date range
  let rangeStart, rangeLabel;
  if (frequency === 'weekly') {
    rangeStart = addDays(today, -7);
    rangeLabel = 'Weekly';
  } else {
    rangeStart = today;
    rangeLabel = 'Daily';
  }

  // ── Pull data from SendLog ─────────────────────────────────
  let totalSent    = 0;
  let totalFailed  = 0;
  const logSheet   = getSheet(CONFIG.sheets.sendLog);
  const logLastRow = logSheet.getLastRow();
  if (logLastRow >= 2) {
    const logData = logSheet.getRange(2, 1, logLastRow - 1, 9).getValues();
    logData.forEach(row => {
      let ts = '';
      if (row[0] instanceof Date) {
        const d    = row[0];
        const yyyy = d.getFullYear();
        const mm   = String(d.getMonth() + 1).padStart(2, '0');
        const dd   = String(d.getDate()).padStart(2, '0');
        ts = yyyy + '-' + mm + '-' + dd;
      } else {
        ts = String(row[0]).substring(0, 10);
      }
      if (ts < rangeStart) return;
      const result = String(row[8]).trim();
      if (result === 'Sent')   totalSent++;
      if (result === 'Failed') totalFailed++;
    });
  }

  // ── Pull data from ActiveFollowUps ─────────────────────────
  let totalReplied = 0;
  let totalDone    = 0;
  let totalActive  = 0;
  let totalErrors  = 0;
  const errorLeads = []; // { email, notes }

  const activeSheet   = getSheet(CONFIG.sheets.activeFollowUps);
  const activeLastRow = activeSheet.getLastRow();
  if (activeLastRow >= 2) {
    const activeData = activeSheet.getRange(2, 1, activeLastRow - 1, CONFIG.cols.notes).getValues();
    const c = CONFIG.cols;
    activeData.forEach(row => {
      const status = String(row[c.status - 1]).trim();
      if (status === 'Replied') totalReplied++;
      if (status === 'Done')    totalDone++;
      if (status === 'Active')  totalActive++;
      if (status === 'Error') {
        totalErrors++;
        errorLeads.push({
          email: String(row[c.leadEmail - 1]),
          notes: String(row[c.notes - 1] || '').split('\n').slice(-3).join(' | '),
        });
      }
    });
  }

  // ── Build email ────────────────────────────────────────────
  const subjectDate = frequency === 'weekly'
    ? rangeStart + ' to ' + today
    : today;
  const subject = 'Follow-Up System - ' + rangeLabel + ' Summary - ' + subjectDate;

  let body = '<h2>Follow-Up System ' + rangeLabel + ' Summary</h2>';
  body += '<p><strong>Period:</strong> ' + subjectDate + '</p>';
  body += '<hr>';
  body += '<table style="border-collapse:collapse;font-family:sans-serif;font-size:14px">';

  const row = (label, value) =>
    '<tr><td style="padding:6px 16px 6px 0;color:#555">' + label + '</td>' +
    '<td style="padding:6px 0;font-weight:bold">' + value + '</td></tr>';

  body += row('Emails sent', totalSent);
  body += row('Send failures', totalFailed);
  body += row('Leads that replied', totalReplied);
  body += row('Sequences completed (Done)', totalDone);
  body += row('Leads with errors', totalErrors);
  body += row('Active leads still in progress', totalActive);
  body += '</table>';

  if (errorLeads.length > 0) {
    body += '<h3 style="color:#c0392b">Leads with Errors</h3>';
    body += '<table style="border-collapse:collapse;font-family:sans-serif;font-size:13px;width:100%">';
    body += '<tr style="background:#f5f5f5"><th style="padding:6px 12px;text-align:left">Email</th>' +
            '<th style="padding:6px 12px;text-align:left">Last Note</th></tr>';
    errorLeads.forEach((el, idx) => {
      const bg = idx % 2 === 0 ? '#fff' : '#fafafa';
      body += '<tr style="background:' + bg + '">' +
              '<td style="padding:6px 12px">' + el.email + '</td>' +
              '<td style="padding:6px 12px;color:#c0392b">' + el.notes + '</td></tr>';
    });
    body += '</table>';
  }

  body += '<p style="color:#999;font-size:12px;margin-top:24px">Sent by Follow-Up System · AK</p>';

  try {
    MailApp.sendEmail({
      to:       recipient,
      subject:  subject,
      htmlBody: body,
    });
    Logger.log('Notification email sent to ' + recipient);
  } catch (e) {
    Logger.log('sendNotificationEmail ERROR: ' + e.message);
  }
}

/**
 * Called by the weekly Monday 9am trigger.
 * Only sends if notificationFrequency = weekly AND today matches notificationDay.
 * Uses timeapi.io via currentDayFull().
 */
function weeklyNotificationTrigger() {
  const freq = (getSetting('notificationFrequency') || 'daily').toLowerCase();
  if (freq !== 'weekly') {
    Logger.log('weeklyNotificationTrigger: frequency is not weekly - skipping.');
    return;
  }

  // Check the notificationDay matches today (from timeapi.io)
  const notifDay  = getSetting('notificationDay') || 'Monday';
  const todayFull = currentDayFull(); // e.g. "Monday" from timeapi.io
  if (todayFull.toLowerCase() !== notifDay.toLowerCase()) {
    Logger.log('weeklyNotificationTrigger: today (' + todayFull + ') != notificationDay (' + notifDay + ') - skipping.');
    return;
  }

  sendNotificationEmail();
}

// ============================================================
// SECTION: SHEET 1 - INBOX SCANNER
// ============================================================

/**
 * Scans Gmail for ALL threads involving the email address in the active row's column A.
 * Each thread gets its own row in InboxScanner so the user can compare them
 * and tick Enroll (col G) on the correct thread to follow up on.
 *
 * Columns written per row:
 *   A leadEmail | B threadId | C subject | D fromAccount | E lastReplyDate | F body | G enroll
 *
 * The body column (F) shows the plain-text content of the FIRST message in the thread,
 * truncated to 500 characters - enough to identify which thread is the right one.
 */
function scanInboxForEmail() {
  const sheet   = getSheet(CONFIG.sheets.inboxScanner);
  const lastRow = sheet.getLastRow();

  // Find the first row in column A that contains a valid email and has no threadId yet
  // (threadId in col B being empty means it hasn't been scanned yet)
  let email = '';
  let row   = -1;

  if (lastRow >= 2) {
    const colAValues = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // cols A and B
    for (let i = 0; i < colAValues.length; i++) {
      const candidate  = String(colAValues[i][0]).trim();
      const hasThread  = String(colAValues[i][1]).trim(); // col B = threadId
      if (candidate.includes('@') && !hasThread) {
        email = candidate;
        row   = i + 2; // 1-based sheet row
        break;
      }
    }
  }

  if (!email || row < 2) {
    Logger.log('No unscanned email address found in column A. Add an email to column A of a new row (leave column B empty) then run again.');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'No unscanned email found in column A. Add the email to column A of a new row (leave column B blank) and run again.',
      'Inbox Scanner', 8
    );
    return;
  }

  Logger.log('Scanning Gmail for ALL threads involving: ' + email);
  const query = 'from:' + email + ' OR to:' + email;
  let threads = [];
  try {
    threads = GmailApp.search(query, 0, 500);
  } catch (e) {
    Logger.log('Gmail search failed: ' + e.message);
    return;
  }

  if (threads.length === 0) {
    Logger.log('No Gmail threads found involving ' + email);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'No threads found for ' + email,
      'Inbox Scan', 5
    );
    return;
  }

  const checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const sendingAccounts    = getSendingAccounts();
  const sc                 = CONFIG.scannerCols;

  // Build all rows as a batch - one array per thread
  // We write each row individually because we also need to set checkbox validation on col G
  let writeRow = row;

  threads.forEach((thread, threadIndex) => {
    const msgs = thread.getMessages();

    // Determine fromAccount: first message whose from/to matches a sending account
    let fromAccount = '';
    for (const msg of msgs) {
      const msgFrom = msg.getFrom();
      const matched = sendingAccounts.find(a => msgFrom.toLowerCase().includes(a.toLowerCase()));
      if (matched) { fromAccount = matched; break; }
      const msgTo     = msg.getTo() + ',' + msg.getCc();
      const matchedTo = sendingAccounts.find(a => msgTo.toLowerCase().includes(a.toLowerCase()));
      if (matchedTo) { fromAccount = matchedTo; break; }
    }

    // Last reply date as YYYY-MM-DD - use plain JS formatting to avoid tz dependency
    const lastDate = thread.getLastMessageDate();
    const yyyy     = lastDate.getFullYear();
    const mm       = String(lastDate.getMonth() + 1).padStart(2, '0');
    const dd       = String(lastDate.getDate()).padStart(2, '0');
    const lastDateStr = yyyy + '-' + mm + '-' + dd;

    // Body: plain text of the FIRST message in the thread, truncated to 500 chars.
    // The first message is the original cold email - most useful for identifying the thread.
    let bodyPreview = '';
    try {
      const firstMsg = msgs[0];
      const rawBody  = firstMsg.getPlainBody() || stripHtml(firstMsg.getBody());
      bodyPreview    = rawBody.trim().replace(/\s+/g, ' ').substring(0, 500);
    } catch (e) {
      bodyPreview = '[Could not read body: ' + e.message + ']';
    }

    // Write all scalar columns in one setValues call for this row
    sheet.getRange(writeRow, sc.leadEmail,     1, 1).setValue(email);
    sheet.getRange(writeRow, sc.threadId,      1, 1).setValue(thread.getId());
    sheet.getRange(writeRow, sc.subject,       1, 1).setValue(thread.getFirstMessageSubject());
    sheet.getRange(writeRow, sc.fromAccount,   1, 1).setValue(fromAccount);
    sheet.getRange(writeRow, sc.lastReplyDate, 1, 1).setValue(lastDateStr);
    sheet.getRange(writeRow, sc.body,          1, 1).setValue(bodyPreview);

    // Enroll column: checkbox (col G = 7)
    sheet.getRange(writeRow, sc.enroll).insertCheckboxes().setDataValidation(checkboxValidation);

    writeRow++;
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Found ' + threads.length + ' thread(s) for ' + email + '. ' +
    'Review the subject and body columns to identify the right thread, ' +
    'then tick Enroll (col G) on that row.',
    'Inbox Scan Complete', 10
  );
}

/**
 * Installable onEdit trigger - watches column G (enroll) of InboxScanner for checkbox = TRUE.
 * Column shifted from 6 to 7 because body column was inserted at position 6.
 */
function watchEnrollCheckbox(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.sheets.inboxScanner) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  // Enroll is now column 7 (G)
  if (col !== CONFIG.scannerCols.enroll || row < 2) return;
  if (e.range.getValue() !== true) return;
  enrollLead(sheet, row);
}

function enrollLead(sheet, row) {
  // Acquire a script-wide lock to prevent concurrent enrollments from two rapid
  // checkbox clicks corrupting the row insertion order in ActiveFollowUps.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000); // wait up to 8 seconds; throws if cannot acquire
  } catch (e) {
    Logger.log('enrollLead: could not acquire lock - ' + e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Another enrollment is in progress. Please try again in a moment.', 'Enroll', 5
    );
    return;
  }

  try {
    _enrollLeadInner(sheet, row);
  } finally {
    lock.releaseLock();
  }
}

function _enrollLeadInner(sheet, row) {
  const sc = CONFIG.scannerCols;

  const leadEmail   = sheet.getRange(row, sc.leadEmail).getValue().toString().trim();
  const threadId    = sheet.getRange(row, sc.threadId).getValue().toString().trim();
  const fromAccount = sheet.getRange(row, sc.fromAccount).getValue().toString().trim();

  Logger.log('enrollLead called: email=' + leadEmail + ' | threadId=' + threadId + ' | row=' + row);

  if (!leadEmail || !threadId) {
    Logger.log('enrollLead: missing leadEmail or threadId on row ' + row);
    return;
  }

  // Use plain JS Date here - UrlFetchApp (used by todayStr via timeapi.io) is NOT
  // available inside onEdit installable triggers. Plain JS Date is always safe.
  const now      = new Date();
  const todayVal = now.getFullYear() + '-' +
                   String(now.getMonth() + 1).padStart(2, '0') + '-' +
                   String(now.getDate()).padStart(2, '0');
  const tom      = new Date(now);
  tom.setDate(tom.getDate() + 1);
  const tomorrow = tom.getFullYear() + '-' +
                   String(tom.getMonth() + 1).padStart(2, '0') + '-' +
                   String(tom.getDate()).padStart(2, '0');

  // Read enrollment defaults from Settings
  const defaultSequence      = getSetting('enrollDefaultSequence')      || '';
  const defaultTotalSteps    = parseInt(getSetting('enrollDefaultTotalSteps') || '0');
  const defaultResumeOnReply = getSetting('enrollDefaultResumeOnReply').toUpperCase() === 'TRUE';

  const activeSheet = getSheet(CONFIG.sheets.activeFollowUps);
  Logger.log('enrollLead: defaultSequence=' + defaultSequence + ' | defaultTotalSteps=' + defaultTotalSteps + ' | defaultResumeOnReply=' + defaultResumeOnReply);
  Logger.log('enrollLead: nextSendDate will be set to ' + tomorrow + ' | activeSheet lastRow before insert=' + activeSheet.getLastRow());
  const newRow = [
    '',                    // leadName
    leadEmail,             // leadEmail
    threadId,              // threadId
    fromAccount,           // fromAccount
    defaultSequence,       // sequenceName
    0,                     // sequenceStep
    defaultTotalSteps,     // totalSteps
    tomorrow,              // nextSendDate
    'Active',              // status
    defaultResumeOnReply,  // resumeOnReply
    '',                    // lastSentDate
    '',                    // lastReplySnippet
    '',                    // notes
  ];

  // Insert at row 2 (just below the header) so the newest lead appears at the TOP.
  // insertRowBefore(2) pushes all existing data down by one row first.
  Logger.log('enrollLead: inserting new lead at row 2 (top of ActiveFollowUps)');
  Logger.log('enrollLead: row data = ' + JSON.stringify(newRow));
  try {
    activeSheet.insertRowBefore(2);
    activeSheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
    Logger.log('enrollLead: insert succeeded - lead is now at row 2, total rows = ' + activeSheet.getLastRow());
  } catch (e) {
    Logger.log('enrollLead: INSERT FAILED - ' + e.message);
    return;
  }

  // Apply formatting and validation to the newly inserted row 2
  const resumeVal = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  activeSheet.getRange(2, CONFIG.cols.resumeOnReply).setDataValidation(resumeVal);
  activeSheet.getRange(2, CONFIG.cols.nextSendDate).setNumberFormat('yyyy-MM-dd');
  Logger.log('enrollLead: formatting applied to row 2');

  // Mark InboxScanner row as Enrolled to prevent re-firing
  sheet.getRange(row, sc.enroll).clearDataValidations().setValue('Enrolled');
  Logger.log('enrollLead: InboxScanner row ' + row + ' marked as Enrolled');

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    defaultSequence
      ? 'Lead enrolled at the TOP of ActiveFollowUps with sequence: ' + defaultSequence
      : 'Lead enrolled at the TOP of ActiveFollowUps. Set sequenceName before tomorrow.',
    'Lead Enrolled ✓', 8
  );
  Logger.log('=== enrollLead DONE: ' + leadEmail + ' | threadId: ' + threadId + ' | nextSendDate: ' + tomorrow + ' ===');
}

// ============================================================
// SECTION: SHEET 2 - SEQUENCES
// ============================================================

/**
 * Returns { awaitDays, message } for the given sequence name and 1-based step number.
 * Returns null if the step row is empty or doesn't exist.
 */
function getSequenceStep(sequenceName, stepNumber) {
  const sheet   = getSheet(CONFIG.sheets.sequences);
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 1) throw new Error('Sequences sheet is empty');

  const headers     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let msgColIndex   = -1;
  for (let c = 0; c < headers.length; c++) {
    if (headers[c] === sequenceName) { msgColIndex = c + 1; break; }
  }
  if (msgColIndex === -1) throw new Error('Sequence "' + sequenceName + '" not found in Sequences sheet');

  const awaitColIndex = msgColIndex - 1;
  if (awaitColIndex < 1) throw new Error('No awaitDays column to the left of sequence "' + sequenceName + '"');

  const dataRow = stepNumber + 1; // row 1 = header; step 1 → row 2
  if (dataRow > lastRow) return null;

  const awaitDays = sheet.getRange(dataRow, awaitColIndex).getValue();
  const message   = sheet.getRange(dataRow, msgColIndex).getValue();
  if (message === '' || message === null || message === undefined) return null;

  return { awaitDays: Number(awaitDays) || 1, message: message.toString() };
}

/**
 * Counts non-empty rows in the message column of the named sequence.
 */
function getSequenceTotalSteps(sequenceName) {
  const sheet   = getSheet(CONFIG.sheets.sequences);
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 2) return 0;

  const headers   = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let msgColIndex = -1;
  for (let c = 0; c < headers.length; c++) {
    if (headers[c] === sequenceName) { msgColIndex = c + 1; break; }
  }
  if (msgColIndex === -1) return 0;

  const colValues = sheet.getRange(2, msgColIndex, lastRow - 1, 1).getValues();
  return colValues.filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined).length;
}

/**
 * Returns an array of all sequence names from row 1 of the Sequences sheet.
 */
function getAllSequenceNames() {
  const sheet   = getSheet(CONFIG.sheets.sequences);
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return headers.filter(h => h !== '' && h !== 'awaitDays');
}

/**
 * Shows a dialog listing all available sequence names.
 */
function listSequenceNames() {
  const names = getAllSequenceNames();
  if (names.length === 0) {
    Logger.log('No sequences found in the Sequences sheet.\n\nAdd column headers next to awaitDays columns to get started.');
    return;
  }
  Logger.log(
    'Available sequences:\n\n• ' + names.join('\n• ') +
    '\n\nUse these exact names in the sequenceName column of ActiveFollowUps.'
  );
}

// ============================================================
// SECTION: SEND FUNCTIONS
// ============================================================

/**
 * Sends a reply into the given Gmail thread from the authenticated Google account.
 * fromAccount is accepted in the signature for call-site compatibility but is not used —
 * sending always comes from the account that owns this Apps Script.
 */
/**
 * Sends a threaded reply to leadEmail using a raw RFC 2822 message via Gmail API.
 *
 * Why raw send instead of thread.reply():
 *   thread.reply() always replies to the sender of the LAST message in the thread.
 *   After step 1 the last sender is US, so thread.reply() on step 2 would email
 *   ourselves. Building a raw message with an explicit To: header and the correct
 *   In-Reply-To / References chain ensures every step goes to the lead and stays
 *   threaded correctly in both our inbox and theirs.
 *
 * If imgBlob is provided the email is multipart/related with the image embedded
 * inline via a Content-ID reference (<img src="cid:inline-image">) so it appears
 * in the email body, not as an attachment.
 *
 * Requires the Gmail advanced service to be enabled:
 *   Extensions → Apps Script → Services → Gmail API (Gmail)
 */
function sendReplyRaw(thread, toEmail, bodyText, imgBlob, imgSize) {
  const messages       = thread.getMessages();
  const lastMessage    = messages[messages.length - 1];
  const lastGmailId    = lastMessage.getId();
  const originalSubject = thread.getFirstMessageSubject();
  const replySubject    = originalSubject.startsWith('Re: ')
    ? originalSubject
    : 'Re: ' + originalSubject;

  // Fetch RFC headers from the last message so we can set In-Reply-To / References
  let messageIdHeader  = '';
  let referencesHeader = '';
  try {
    const rawMsg = Gmail.Users.Messages.get('me', lastGmailId, { format: 'full' });
    if (rawMsg.payload && rawMsg.payload.headers) {
      for (const h of rawMsg.payload.headers) {
        const name = h.name.toLowerCase();
        if (name === 'message-id') messageIdHeader  = h.value;
        if (name === 'references') referencesHeader = h.value;
      }
    }
  } catch (e) {
    Logger.log('sendReplyRaw: could not fetch RFC headers - ' + e.message + ' - using fallback');
  }

  if (!messageIdHeader) {
    messageIdHeader = '<' + lastGmailId + '@mail.gmail.com>';
    Logger.log('sendReplyRaw: using fallback Message-ID ' + messageIdHeader);
  }

  const references = referencesHeader
    ? referencesHeader + ' ' + messageIdHeader
    : messageIdHeader;

  const fromEmail = Session.getActiveUser().getEmail();
  const htmlBody  = bodyText.replace(/\n/g, '<br>');

  let rawEmail;

  if (imgBlob) {
    // ── Multipart/related: inline image embedded in body ──────────────────
    const cid       = 'inline-image-' + Date.now();
    const boundary  = 'boundary_' + Utilities.getUuid().replace(/-/g, '');
    const imgBase64 = Utilities.base64Encode(imgBlob.getBytes());
    const mimeType  = imgBlob.getContentType() || 'image/jpeg';

    const sizeStyle = imgSize ? 'width:' + imgSize + ';max-width:100%' : 'max-width:100%';
    const imgTag    = '<br><img src="cid:' + cid + '" style="' + sizeStyle + '"><br>';
    // If the text contains {{IMG_PLACEHOLDER}} (set by parseMessage when there is text
    // both above and below the [IMG] line), replace it with the img tag in position.
    // Otherwise fall back to appending the image at the end.
    const htmlWithImg = htmlBody.includes('{{IMG_PLACEHOLDER}}')
      ? htmlBody.replace('{{IMG_PLACEHOLDER}}', imgTag)
      : htmlBody + imgTag;

    rawEmail = [
      'MIME-Version: 1.0',
      'From: ' + fromEmail,
      'To: ' + toEmail,
      'Subject: ' + replySubject,
      'In-Reply-To: ' + messageIdHeader,
      'References: ' + references,
      'Content-Type: multipart/related; boundary="' + boundary + '"',
      '',
      '--' + boundary,
      'Content-Type: text/html; charset=UTF-8',
      '',
      htmlWithImg,
      '',
      '--' + boundary,
      'Content-Type: ' + mimeType + '; name="image"',
      'Content-Transfer-Encoding: base64',
      'Content-ID: <' + cid + '>',
      'Content-Disposition: inline; filename="image"',
      '',
      imgBase64,
      '',
      '--' + boundary + '--'
    ].join('\r\n');

    Logger.log('sendReplyRaw: sending multipart/related with inline image to ' + toEmail);
  } else {
    // ── Plain HTML email, no image ────────────────────────────────────────
    rawEmail = [
      'MIME-Version: 1.0',
      'From: ' + fromEmail,
      'To: ' + toEmail,
      'Subject: ' + replySubject,
      'In-Reply-To: ' + messageIdHeader,
      'References: ' + references,
      'Content-Type: text/html; charset=UTF-8',
      '',
      htmlBody
    ].join('\r\n');

    Logger.log('sendReplyRaw: sending HTML email to ' + toEmail);
  }

  Gmail.Users.Messages.send(
    { raw: Utilities.base64EncodeWebSafe(rawEmail), threadId: thread.getId() },
    'me'
  );
}

/**
 * Replaces {{variableName}} tokens in a message template with lead data.
 *
 * Built-in variables (always available):
 *   {{leadName}}   - full name from col A of ActiveFollowUps
 *   {{leadEmail}}  - email from col B
 *   {{firstName}}  - first word of leadName (e.g. "John" from "John Smith")
 *
 * Custom variables:
 *   Any column header you add to ActiveFollowUps beyond col 13 becomes a variable.
 *   Example: add column header "companyName" → use {{companyName}} in sequences.
 *   The value in that cell for each lead is substituted automatically.
 *
 * Unknown tokens (no matching column) are left as-is so you notice them.
 */
function replaceVariables(text, leadName, leadEmail, customVars) {
  if (!text) return text;

  // Built-in variables
  const firstName = leadName ? leadName.trim().split(/\s+/)[0] : '';
  let result = text
    .replace(/\{\{leadName\}\}/g,  leadName  || '')
    .replace(/\{\{leadEmail\}\}/g, leadEmail || '')
    .replace(/\{\{firstName\}\}/g, firstName || '');

  // Custom variables from extra columns
  if (customVars) {
    for (const key in customVars) {
      const regex = new RegExp('\\{\\{' + key + '\\}\\}', 'g');
      result = result.replace(regex, customVars[key]);
    }
  }

  return result;
}

/**
 * Splits a raw message cell value into text, optional image Drive URL, and image size.
 *
 * Handles all line break styles Google Sheets may store (\n, \r\n, \r).
 *
 * [IMG] syntax examples in the sequence cell:
 *   [IMG] https://drive.google.com/...          ← full width (default)
 *   [IMG 50%] https://drive.google.com/...      ← half width
 *   [IMG 300px] https://drive.google.com/...    ← fixed 300px width
 *   [IMG 75%] https://drive.google.com/...      ← 75% width
 *
 * Returns { text, imgUrl, imgSize }
 *   imgSize is a CSS width string e.g. '50%', '300px', or '100%' (default)
 */
function parseMessage(rawMessage) {
  // Normalise all line break styles to plain \n so split works reliably
  const normalised   = rawMessage.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines        = normalised.split('\n');
  const imgLineIndex = lines.findIndex(l => l.trim().match(/^\[IMG(\s[^\]]*)?\]/i));
  if (imgLineIndex === -1) return { text: normalised.trim(), imgUrl: null, imgSize: null };

  const textBefore = lines.slice(0, imgLineIndex).join('\n').trim();
  const textAfter  = lines.slice(imgLineIndex + 1).join('\n').trim();
  // Combine text before and after the [IMG] line with a placeholder so the image
  // sits between the two text blocks when rendered in the email body.
  // The placeholder {{IMG_PLACEHOLDER}} is replaced in sendReplyRaw with the actual <img> tag.
  // Only prepend a newline before the placeholder if there is text above the image
  const text = (textBefore ? textBefore + '\n' : '') +
               '{{IMG_PLACEHOLDER}}' +
               (textAfter ? '\n' + textAfter : '');
  const imgLine   = lines[imgLineIndex].trim();

  // Parse optional size from [IMG 50%] or [IMG 300px]
  const sizeMatch = imgLine.match(/^\[IMG\s+([^\]]+)\]/i);
  const imgSize   = sizeMatch ? sizeMatch[1].trim() : '100%';

  // Extract the URL - everything after the closing ]
  const imgUrl = imgLine.replace(/^\[IMG[^\]]*\]\s*/i, '').trim();

  return { text, imgUrl, imgSize };
}

/**
 * Fetches a file from Google Drive by Drive share URL and returns it as a Blob.
 * Used for inline image embedding in sendReplyRaw().
 */
function getImageBlob(driveUrl) {
  const match = driveUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) {
    Logger.log('getImageBlob: could not extract file ID from URL: ' + driveUrl);
    return null;
  }
  const file = DriveApp.getFileById(match[1]);
  const blob  = file.getBlob();
  blob.setName(file.getName());
  return blob;
}

// ============================================================
// SECTION: HELPER FUNCTIONS
// ============================================================

/**
 * Returns a spreadsheet sheet by name. Throws if not found.
 */
function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: "' + name + '". Run "Setup spreadsheet" from the menu.');
  return sheet;
}

/**
 * Strips HTML tags from a string.
 */
function stripHtml(html) {
  if (!html) return '';
  let text = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<\/tr>/gi, '\n');
  text = text.replace(/<[^>]+>/g, '');
  text = text
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ');
  return text.replace(/\n{3,}/g, '\n\n').trim();
}

/**
 * Returns the array of configured sending account email addresses.
 */
function getSendingAccounts() {
  return CONFIG.sendingAccounts;
}

/**
 * Returns true if the given email string is one of the sending accounts.
 */
function isSendingAccount(emailStr) {
  if (!emailStr) return false;
  const lower = emailStr.toLowerCase();
  return CONFIG.sendingAccounts.some(a => lower.includes(a.toLowerCase()));
}

/**
 * Appends one row to the SendLog sheet.
 */
function logToSendLog(leadEmail, leadName, sequenceName, stepNumber, fromAccount, messagePreview, threadId, result) {
  try {
    getSheet(CONFIG.sheets.sendLog).appendRow([
      new Date(), leadEmail, leadName, sequenceName,
      stepNumber, fromAccount, messagePreview, threadId, result,
    ]);
  } catch (e) {
    Logger.log('logToSendLog ERROR: ' + e.message);
  }
}

/**
 * Returns all data rows from ActiveFollowUps as a 2D array (header excluded).
 */
function getActiveLeads() {
  const sheet   = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, CONFIG.cols.notes).getValues();
}

/**
 * Writes a value to a specific cell by row and column number (both 1-based).
 */
function setCell(sheet, row, col, value) {
  sheet.getRange(row, col).setValue(value);
}

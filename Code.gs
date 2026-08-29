// ============================================================
// SECTION: CONFIGURATION
// ============================================================

const CONFIG = {
  // Auto-detected from the Google account running this script.
  // All replies go here and sending always comes from this account.
  //
  // getActiveUser() can come back EMPTY under a time-based trigger on a
  // consumer Gmail account. An empty string here is poison: JavaScript says
  // 'anything'.includes('') === true, so isSendingAccount would match every
  // sender alive and silently switch off reply, bounce AND OOO detection.
  // getEffectiveUser() is the account the script actually runs as and is
  // reliable, so it goes first. Blanks are dropped, and a throw is caught so
  // this can never break the whole script at load time.
  sendingAccounts: (function () {
    const out = [];
    try { const e = Session.getEffectiveUser().getEmail(); if (e) out.push(e); } catch (err) {}
    try { const a = Session.getActiveUser().getEmail();    if (a && out.indexOf(a) === -1) out.push(a); } catch (err) {}
    return out;
  })(),

  // TEMPORARY absence. A real person who is simply away - safe to resume later.
  oooKeywords: [
    'out of office', 'on leave', 'away from the office', 'on vacation',
    'auto-reply', 'automatic reply', 'autoreply', 'on holiday',
    'annual leave', 'parental leave', 'maternity leave', 'currently away',
    'back in the office', 'limited access to email'
  ],

  // PERMANENT failure. The address is dead - never resume, never retry.
  // Checked BEFORE oooKeywords, because a bounce is the more specific case.
  bounceKeywords: [
    'undeliverable', 'delivery failed', 'delivery has failed',
    'delivery status notification', 'address not found',
    'recipient not found', 'no such user', 'user unknown',
    'mailbox unavailable', 'mailbox full', 'account has been disabled',
    'permanent error', 'message blocked', 'does not exist'
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
    .addItem('Clean up stuck triggers', 'cleanUpStuckTriggers')
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
    ['replyResumeDays',            '7'],
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
    .requireValueInList(['Active', 'Replied', 'Paused', 'Done', 'Error', 'OOO', 'Bounced'], true)
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

  // Weekly notification trigger, on whatever day notificationDay names.
  //
  // This used to be hardcoded to MONDAY while weeklyNotificationTrigger only
  // sends when today matches notificationDay - so any other day meant the
  // trigger fired Monday, saw a mismatch, and skipped. You simply never got a
  // weekly summary, with nothing to show why.
  //
  // Re-run this function after changing notificationDay.
  const WEEKDAYS = {
    sunday:    ScriptApp.WeekDay.SUNDAY,
    monday:    ScriptApp.WeekDay.MONDAY,
    tuesday:   ScriptApp.WeekDay.TUESDAY,
    wednesday: ScriptApp.WeekDay.WEDNESDAY,
    thursday:  ScriptApp.WeekDay.THURSDAY,
    friday:    ScriptApp.WeekDay.FRIDAY,
    saturday:  ScriptApp.WeekDay.SATURDAY
  };
  const _notifDay = String(getSetting('notificationDay') || 'Monday').trim().toLowerCase();
  const _weekDay  = WEEKDAYS[_notifDay] || ScriptApp.WeekDay.MONDAY;
  if (!WEEKDAYS[_notifDay]) {
    Logger.log('createTriggers: notificationDay "' + _notifDay + '" not recognised - using Monday');
  }
  ScriptApp.newTrigger('weeklyNotificationTrigger')
    .timeBased()
    .onWeekDay(_weekDay)
    .atHour(9)
    .create();
  Logger.log('createTriggers: weekly summary trigger set for ' + _notifDay);

  // Installable onEdit trigger for InboxScanner checkbox (needs Gmail access)
  ScriptApp.newTrigger('watchEnrollCheckbox')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  Logger.log(
    'Triggers created!\n\n' +
    '• startDailyRun fires every day at 8am\n' +
    '• weeklyNotificationTrigger fires every ' + _notifDay + ' at 9am\n' +
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

  // Only one chain at a time. Without this, hitting "Start follow-up chain now"
  // while the 8am chain is running gives two independent chains that can both
  // grab the same lead and email it twice. Returning here also protects the
  // running chain from the trigger reap below, which would otherwise delete
  // its pending next step.
  if (chainIsRunning()) {
    Logger.log('startDailyRun: a chain is already running - not starting a second one.');
    return;
  }

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

  // Reap orphaned processOneLead triggers. Google allows only 20 triggers per
  // script. A brand new chain is about to start, so every existing
  // processOneLead trigger is stale by definition. Delete them all.
  let _reaped = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processOneLead') {
      try { ScriptApp.deleteTrigger(t); _reaped++; }
      catch (err) { Logger.log('startDailyRun: could not delete stale processOneLead trigger - ' + err.message); }
    }
  });
  Logger.log('startDailyRun: reaped ' + _reaped + ' stale processOneLead trigger(s) | total triggers now: ' + ScriptApp.getProjectTriggers().length);

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

  // Put "said yes then went quiet" leads back into their sequence.
  checkReplyAutoResume();

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

      const nextSendDate = sheetDateToStr(row[c.nextSendDate - 1]);
      if (!nextSendDate || !dateIsOnOrBefore(nextSendDate, today)) continue;

      const lastSentDate = sheetDateToStr(row[c.lastSentDate - 1]);
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

  // Claim the lock before scheduling. Without this the lock is not set until
  // the first processOneLead fires 4-8 min later, leaving a window where a
  // second startDailyRun sees no lock and starts a parallel chain.
  markChainRunning();

  // Schedule the first processOneLead - it will chain itself from there
  scheduleNextProcessOneLead();
  Logger.log('startDailyRun: first processOneLead trigger scheduled - chain will begin shortly');
}

/**
 * Called by chained time-based triggers. e contains triggerUid.
 * Deletes its own trigger first, then processes one lead, then reschedules.
 */
function processOneLead(e) {
  const runStart   = Date.now();
  const MAX_RUN_MS = 4 * 60 * 1000; // bail out before the 6-minute Apps Script limit
  const seen       = {};            // threadIds already handled in this execution
  let   skipped    = 0;

  Logger.log('processOneLead: triggerUid=' + (e ? e.triggerUid : 'manual') + ' | sendsToday=' + getSendsToday() + ' | maxSends=' + getSetting('maxSendsPerDay'));

  // Delete this trigger immediately to prevent orphan buildup
  if (e && e.triggerUid) {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getUniqueId() === e.triggerUid) {
        ScriptApp.deleteTrigger(t);
        Logger.log('Deleted self-trigger: ' + e.triggerUid);
      }
    });
  }

  for (const key in _settingsCache) delete _settingsCache[key];
  for (const key in _timeApiCache)  delete _timeApiCache[key];

  // Refresh the chain lock so startDailyRun knows a chain is alive.
  markChainRunning();

  // Work through leads. ONLY A REAL SEND ends this execution.
  while (true) {

    if (!isWithinSendingWindow()) {
      Logger.log('Outside sending window - stopping chain. (' + skipped + ' skipped this execution)');
      clearChainLock();
      return;
    }
    if (getSetting('pauseAllFollowUps').toUpperCase() === 'TRUE') {
      Logger.log('pauseAllFollowUps is TRUE - stopping chain.');
      clearChainLock();
      return;
    }
    if (hitMaxSendsPerDay()) {
      Logger.log('maxSendsPerDay reached - stopping chain.');
      clearChainLock();
      return;
    }
    if (PropertiesService.getScriptProperties().getProperty('quotaStopDate') === todayStr()) {
      Logger.log('Gmail quota was hit earlier today - chain stays stopped.');
      clearChainLock();
      return;
    }

    const lead = findNextDueLead();

    if (!lead) {
      const props             = PropertiesService.getScriptProperties();
      const sendsAtChainStart = parseInt(props.getProperty('sendsAtChainStart') || '0');
      const sendsNow          = getSendsToday();
      const sentThisRun       = sendsNow - sendsAtChainStart;
      Logger.log('All leads processed for today. Chain complete. (' + skipped + ' skipped this execution)');
      Logger.log('processOneLead: chain ended | sendsAtChainStart=' + sendsAtChainStart + ' | sendsNow=' + sendsNow + ' | sentThisRun=' + sentThisRun);
      if (sentThisRun > 0 && (getSetting('notificationFrequency') || 'daily').toLowerCase() === 'daily') {
        Logger.log('Chain complete: ' + sentThisRun + ' email(s) sent this run - sending notification.');
        sendNotificationEmail();
      } else {
        Logger.log('Chain complete: 0 emails sent this run (total today: ' + sendsNow + ') - skipping notification.');
      }
      clearChainLock();
      return;
    }
    // If the same lead comes back, its status write failed. Park it so the
    // loop below cannot spin on one lead forever.
    if (lead.threadId && seen[lead.threadId]) {
      Logger.log('processOneLead: ' + lead.leadEmail + ' came back a second time - status write must have failed. Parking it.');
      parkLead(lead.threadId, 'status write failed - parked to break loop');
      continue;
    }
    if (lead.threadId) seen[lead.threadId] = true;

    Logger.log('Processing lead: ' + lead.leadEmail + ' (row ' + lead.sheetRow + ')');
    let sent = false;
    try {
      sent = processSingleLead(lead);
    } catch (err) {
      Logger.log('processSingleLead UNCAUGHT ERROR for ' + lead.leadEmail + ': ' + err.message);
      parkLead(lead.threadId, 'uncaught error: ' + err.message);
      const notifEmail = getSetting('notificationEmail');
      if (notifEmail) {
        try {
          MailApp.sendEmail({
            to: notifEmail,
            subject: 'Follow-Up System - unexpected error, chain continuing',
            body: 'An unexpected error occurred processing lead: ' + lead.leadEmail +
              '\nRow: ' + lead.sheetRow + '\nError: ' + err.message +
              '\n\nThat lead has been skipped for today. The chain is continuing.'
          });
        } catch (ne) { Logger.log('Could not send error notification: ' + ne.message); }
      }
      sent = false;
    }

    // An email actually went out -> apply the normal 4-8 min spam pacing.
    if (sent) {
      Logger.log('processOneLead: sent to ' + lead.leadEmail + ' | ' + skipped + ' lead(s) skipped for free in this execution');
      scheduleNextProcessOneLead();
      return;
    }

    // Nothing was emailed -> move straight to the next lead, no delay.
    skipped++;

    if (Date.now() - runStart > MAX_RUN_MS) {
      Logger.log('processOneLead: ' + skipped + ' skips took ' + Math.round((Date.now() - runStart) / 1000) + 's - continuing in a fresh execution.');
      scheduleContinueSoon();
      return;
    }
  }
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

  try {
    ScriptApp.newTrigger('processOneLead')
      .timeBased()
      .after(delayMs)
      .create();
  } catch (e) {
    // Nearly always the 20-triggers-per-script limit. Without this the chain
    // would just stop, with nothing anywhere explaining why.
    Logger.log('scheduleNextProcessOneLead: FAILED to create trigger - ' + e.message);
    notifyTriggerFailure(e);
    return;
  }

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
  const parked   = getParkedLeads(); // leads already handled today - never pick again

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

    // A lead is parked the instant its email is sent, BEFORE any sheet write,
    // so this still holds when the sheet write fails. This is what makes a
    // duplicate send impossible instead of merely unlikely.
    const rowThreadId = String(row[c.threadId - 1]).trim();
    if (rowThreadId && parked[rowThreadId]) {
      Logger.log('findNextDueLead: row ' + (i+2) + ' skipped - parked today (' + parked[rowThreadId] + ')');
      continue;
    }

    const nextSendDate = sheetDateToStr(row[c.nextSendDate - 1]);
    if (!nextSendDate || !dateIsOnOrBefore(nextSendDate, today)) {
      Logger.log('findNextDueLead: row ' + (i+2) + ' skipped - nextSendDate=' + nextSendDate + ' is in the future');
      continue;
    }

    const lastSentDate = sheetDateToStr(row[c.lastSentDate - 1]);
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

  // Row numbers go stale the moment someone ticks an Enroll box - enrollLead
  // does insertRowBefore(2), which slides every row down by one. The row number
  // findNextDueLead handed us can therefore point at a DIFFERENT lead by the
  // time we write. So we never trust it: we re-confirm by threadId first.
  let cachedRow = sheetRow;
  function currentRow() {
    // Fast path: the cached row still holds this lead's threadId (one cell read).
    try {
      if (String(activeSheet.getRange(cachedRow, c.threadId).getValue()).trim() === String(threadId).trim()) {
        return cachedRow;
      }
    } catch (e) {}
    // Rows shifted underneath us - find this lead again by threadId.
    const lastRow = activeSheet.getLastRow();
    if (lastRow >= 2) {
      const ids = activeSheet.getRange(2, c.threadId, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(threadId).trim()) {
          cachedRow = i + 2;
          Logger.log(leadEmail + ': rows shifted - lead is now on row ' + cachedRow);
          return cachedRow;
        }
      }
    }
    throw new Error('lead row for threadId ' + threadId + ' no longer exists in ActiveFollowUps');
  }

  // Helper: write a single cell — throws on failure so callers can catch it
  function writeCell(col, value) {
    try {
      activeSheet.getRange(currentRow(), col).setValue(value);
    } catch (e) {
      throw new Error('writeCell col=' + col + ' value=' + value + ': ' + e.message);
    }
  }

  // Helper: write the status column, and if that write fails, park the lead.
  // The status column is what stops a lead being picked again. If we cannot
  // write it, the lead stays Active and due and would be handed back forever.
  function setStatusSafe(newStatus) {
    try {
      writeCell(c.status, newStatus);
    } catch (err) {
      Logger.log(leadEmail + ': could not write status "' + newStatus + '" - parking lead. ' + err.message);
      parkLead(threadId, 'status write failed (' + newStatus + ')');
    }
  }

  // Helper: append to notes (never overwrite). Never throws.
  function appendNote(text) {
    try {
      const existing = activeSheet.getRange(currentRow(), c.notes).getValue();
      const updated  = existing ? existing + '\n' + text : text;
      writeCell(c.notes, updated);
    } catch (err) {
      Logger.log(leadEmail + ': could not append note - ' + err.message);
    }
  }
  // ── Step 3: Fetch the Gmail thread ────────────────────────
  let thread;
  try {
    thread = GmailApp.getThreadById(threadId);
  } catch (e) {
    Logger.log('Thread fetch error for ' + leadEmail + ': ' + e.message);
    setStatusSafe('Error');
    appendNote('[' + today + '] Thread fetch error: ' + e.message);
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Failed');
    return;
  }

  if (!thread) {
    Logger.log('Thread not found for ' + leadEmail);
    setStatusSafe('Error');
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
    setStatusSafe('Error');
    appendNote('[' + today + '] Error reading messages: ' + e.message);
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Failed');
    return;
  }

  const latestMessage = messages[messages.length - 1];
  const latestSender  = latestMessage.getFrom();

  // ── Step 5: Read the latest message ONCE, then classify it ──────
  let latestBody = '';
  try { latestBody = stripHtml(latestMessage.getBody()); }
  catch (e) { latestBody = latestMessage.getPlainBody() || ''; }
  const bodyLower = latestBody.toLowerCase();
  const fromUs    = isSendingAccount(latestSender);
  const snippet   = latestBody.substring(0, 120);

  // Bounce is checked first - it is the more specific case, and some bounce
  // notices also contain wording that would match an OOO keyword.
  const isBounce = !fromUs && CONFIG.bounceKeywords.some(kw => bodyLower.includes(kw));
  const oooDetected = !fromUs && !isBounce &&
                      CONFIG.oooKeywords.some(kw => bodyLower.includes(kw));
  Logger.log('processSingleLead: fromUs=' + fromUs + ' | isBounce=' + isBounce +
             ' | oooDetected=' + oooDetected +
             ' | latestSender=' + latestSender + ' | bodyPreview=' + bodyLower.substring(0, 80));

  // ── Hard bounce: the address is dead. Stop permanently. ───
  // Deliberately NOT marked OOO, so oooAutoResume can never bring it back and
  // start mailing a dead address on a loop.
  if (isBounce) {
    Logger.log(leadEmail + ': BOUNCE detected from ' + latestSender);
    setStatusSafe('Bounced');
    appendNote('[' + today + '] Bounced - address appears dead. From: ' +
               String(latestSender).replace(/[\r\n]+/g, ' '));
    try { writeCell(c.lastReplySnippet, snippet); }
    catch (e) { Logger.log(leadEmail + ': could not write bounce snippet - ' + e.message); }
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Skipped - Bounced');
    SpreadsheetApp.flush();
    return false;
  }

  // ── Step 5a: OOO / bounce is checked FIRST ────────────────
  // An auto-reply or a mailer-daemon bounce is not from our account, so the
  // reply check below would file it as a genuine "Replied" and kill the
  // sequence permanently. Checking here catches it instead. This runs even
  // when replyCheckBeforeSend is FALSE - there is no point mailing a dead
  // address either way.
  if (oooDetected) {
    Logger.log(leadEmail + ': OOO/bounce detected from ' + latestSender);
    setStatusSafe('OOO');
    // This note carries the date checkOooAutoResume parses to decide when to
    // put the lead back to Active. Without it, auto-resume has nothing to read.
    appendNote('[' + today + '] OOO/bounce detected - paused. From: ' +
               String(latestSender).replace(/[\r\n]+/g, ' '));
    try { writeCell(c.lastReplySnippet, snippet); }
    catch (e) { Logger.log(leadEmail + ': could not write OOO snippet - ' + e.message); }
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Skipped - OOO/bounce detected');
    SpreadsheetApp.flush();
    return false;
  }

  // ── Step 5b: Genuine reply detection (if enabled in Settings) ──
  const replyCheckEnabled = getSetting('replyCheckBeforeSend').toUpperCase() !== 'FALSE';
  if (!replyCheckEnabled) {
    Logger.log('WARNING: replyCheckBeforeSend is FALSE - ' + leadEmail +
               ' will be mailed even if they already replied. Set it TRUE in Settings.');
  }

  // If this lead was resumed from Replied, every message they sent BEFORE that
  // resume has already been dealt with. Without this cutoff the search below
  // finds that old reply, sets Replied again, and the lead can never actually
  // resume - it just bounces between Active and Replied forever.
  // Stored to the MINUTE, not the day. A date-only cutoff can never let a
  // resume send on the same day the lead replied, which makes the whole
  // feature impossible to see working until tomorrow.
  let resumeCutoff = '';
  try {
    const rNotes = String(activeSheet.getRange(currentRow(), c.notes).getValue() || '');
    const rMarks = rNotes.match(/\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2})\][^\n]*Resumed from Replied/gi);
    if (rMarks) resumeCutoff = rMarks[rMarks.length - 1].match(/\[([^\]]+)\]/)[1];
  } catch (e) { Logger.log(leadEmail + ': could not read resume cutoff - ' + e.message); }

  // Search EVERY message in the thread, newest first - not just the latest one.
  // Checking only the latest misses the most common real case: the lead
  // replies, you answer them by hand from Gmail, and now the thread ends with
  // YOUR message - so the sequence marches on at someone you are already
  // talking to. Once a human exchange has started, automation stops.
  // Matching the lead's own address keeps mailer-daemon bounces and any other
  // thread participants out of this, so bounces still reach the bounce branch.
  let leadReply = null;
  const leadAddr = String(leadEmail).toLowerCase().trim();
  if (leadAddr) {
    for (let m = messages.length - 1; m >= 0; m--) {
      if (String(messages[m].getFrom()).toLowerCase().indexOf(leadAddr) === -1) continue;
      // Strictly before the resume moment = already handled. Anything older is
      // older still, so stop looking. Formatted in the configured timezone so
      // it compares against the note on the same clock. 'yyyy-MM-dd HH:mm' is
      // fixed width, so a plain string compare orders it correctly.
      if (resumeCutoff) {
        const mStamp = Utilities.formatDate(messages[m].getDate(),
                         getSetting('timezone') || 'UTC', 'yyyy-MM-dd HH:mm');
        if (mStamp < resumeCutoff) break;
      }
      leadReply = messages[m];
      break;
    }
  }

  if (replyCheckEnabled && leadReply) {
    let replyText = '';
    try { replyText = stripHtml(leadReply.getBody()); }
    catch (e) { replyText = leadReply.getPlainBody() || ''; }
    Logger.log(leadEmail + ': lead replied at ' + leadReply.getDate() + ' - setting Replied');
    setStatusSafe('Replied');
    // Dated marker. checkReplyAutoResume measures replyResumeDays from this,
    // so without it a lead with resumeOnReply ticked can never come back.
    appendNote('[' + today + '] Lead replied - paused.');
    try { writeCell(c.lastReplySnippet, replyText.substring(0, 120)); }
    catch (e) { Logger.log(leadEmail + ': could not write reply snippet - ' + e.message); }
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Skipped - Lead replied');
    SpreadsheetApp.flush();
    return false;
  }

  Logger.log('processSingleLead: reply/OOO checks passed for ' + leadEmail);

  // ── Step 7: Determine step to send ────────────────────────
  const currentStep = sequenceStep; // 0-indexed

  // ── Step 8: Check totalSteps cap ──────────────────────────
  Logger.log('processSingleLead: currentStep=' + currentStep + ' | totalSteps=' + totalSteps + ' | hitsCap=' + (totalSteps > 0 && currentStep >= totalSteps));
  if (totalSteps > 0 && currentStep >= totalSteps) {
    Logger.log(leadEmail + ': Done - max steps reached');
    setStatusSafe('Done');
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
    setStatusSafe('Error');
    appendNote('[' + today + '] Sequence read error: ' + e.message);
    logToSendLog(leadEmail, leadName, sequenceName, sequenceStep, fromAccount, '', threadId, 'Failed');
    SpreadsheetApp.flush();
    return;
  }
  Logger.log('processSingleLead: stepData found=' + !!stepData + ' | awaitDays=' + (stepData ? stepData.awaitDays : 'N/A') + ' | msgLength=' + (stepData ? stepData.message.length : 0));
  if (!stepData || !stepData.message) {
    Logger.log(leadEmail + ': Done - sequence exhausted at step ' + currentStep);
    setStatusSafe('Done');
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
    setStatusSafe('Error');
    appendNote('[' + today + '] Send failed at step ' + currentStep + ': ' + e.message);
    logToSendLog(leadEmail, leadName, sequenceName, currentStep, fromAccount,
      textPart.substring(0, 100), threadId, 'Failed');
    SpreadsheetApp.flush();
    // If quota hit, notify immediately and stop the chain
    if (isQuota) {
      Logger.log('QUOTA HIT - stopping chain and sending notification');
      // Tell the chain to stay stopped for the rest of today. Without this the
      // loop races through every remaining lead, fails each on quota, marks
      // them all Error and emails a warning for every one.
      PropertiesService.getScriptProperties().setProperty('quotaStopDate', today);
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
  // ══════════════════════════════════════════════════════════════════
  // THE EMAIL IS OUT THE DOOR. Everything below is bookkeeping.
  // Park the lead FIRST, before touching the sheet. parkLead writes to Script
  // Properties, which is completely independent of the spreadsheet. Once this
  // line runs, findNextDueLead will refuse to hand this lead back today no
  // matter what happens to any sheet write below.
  // ══════════════════════════════════════════════════════════════════
  parkLead(threadId, 'sent step ' + currentStep + ' on ' + today);

  // Count the send straight away - the email really did go out, so it must
  // count against the daily cap even if the sheet writes below fail.
  try {
    incrementSendsToday();
  } catch (ce) {
    // Never let a counter write throw - the email is already sent, and an
    // escape here would make the chain treat this as a skip and drop the delay.
    Logger.log(leadEmail + ': could not increment send counter - ' + ce.message);
  }

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
    setStatusSafe('Error');
    SpreadsheetApp.flush();
    return true; // the email DID go out - the chain must still pace the next one
  }

  logToSendLog(leadEmail, leadName, sequenceName, currentStep, fromAccount,
    textPart.substring(0, 100), threadId, 'Sent');

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
    setStatusSafe('Done');
    Logger.log(leadEmail + ': status set to Done');
  }

  SpreadsheetApp.flush();
  return true;
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
    // Take the LAST OOO note, not the first. .match without /g returns the
    // oldest one, so a lead who went OOO in June, resumed, then went OOO
    // again yesterday gets measured from June - already past resumeDays -
    // and gets emailed while they are still away.
    // Matching "OOO/bounce detected" specifically also stops the
    // "Auto-resumed from OOO" line being read as an OOO date.
    const oooNotes = notes.match(/\[(\d{4}-\d{2}-\d{2})\][^\n]*OOO\/bounce detected/gi);
    if (oooNotes) {
      oooSetDate = oooNotes[oooNotes.length - 1].match(/\[(\d{4}-\d{2}-\d{2})\]/)[1];
    } else {
      // Fallback: use lastSentDate
      oooSetDate = sheetDateToStr(row[c.lastSentDate - 1]);
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

  const ui = SpreadsheetApp.getUi();

  // Ask WHICH LEAD to pull variables from
  const leadResponse = ui.prompt(
    'Test Email - Choose Lead',
    'Enter the lead\'s EMAIL ADDRESS, or their ROW NUMBER from ActiveFollowUps.\n\n' +
    'Leave blank to use the first lead that has a sequenceName.',
    ui.ButtonSet.OK_CANCEL
  );
  if (leadResponse.getSelectedButton() !== ui.Button.OK) return;
  const leadQuery = leadResponse.getResponseText().trim();

  // Ask which step number to preview
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
    const seqName  = String(data[i][c.sequenceName - 1]).trim();
    const rowEmail = String(data[i][c.leadEmail - 1]).trim();
    const rowNum   = i + 2;

    // Blank query -> first lead with a sequence. Contains @ -> match on email.
    // Otherwise -> match on row number.
    let matches;
    if (!leadQuery) {
      matches = !!seqName;
    } else if (leadQuery.indexOf('@') !== -1) {
      matches = rowEmail.toLowerCase() === leadQuery.toLowerCase();
    } else {
      matches = String(rowNum) === leadQuery;
    }
    if (!matches) continue;

    if (!seqName) {
      ui.alert('That lead (row ' + rowNum + ', ' + rowEmail + ') has no sequenceName set.\n\n' +
               'Fill in column E first, then try again.');
      return;
    }

    const customVars = {};
    for (let col = c.notes; col < headers.length; col++) {
      const h = String(headers[col]).trim();
      if (h) customVars[h] = String(data[i][col] || '').trim();
    }
    lead = {
      leadName:     String(data[i][c.leadName  - 1]),
      leadEmail:    String(data[i][c.leadEmail - 1]),
      sequenceName: seqName,
      sheetRow:     rowNum,
      customVars:   customVars,
    };
    break;
  }

  if (!lead) {
    ui.alert(leadQuery
      ? 'No lead found matching "' + leadQuery + '" in ActiveFollowUps.'
      : 'No lead with a sequenceName found in ActiveFollowUps.');
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

  // No banner - the body is now byte-for-byte what the lead would receive.
  // The [TEST] subject prefix is the only marker, so read the subject line.
  Logger.log('sendTestEmail: step ' + stepToPreview + ' | sequence ' + lead.sequenceName +
             ' | lead ' + lead.leadEmail + ' (row ' + lead.sheetRow + ')' +
             ' | awaitDays ' + stepData.awaitDays +
             ' | image ' + (imgUrl ? imgSize : 'none'));

  const safeText  = imgBlob ? textPart : textPart.replace(/\{\{IMG_PLACEHOLDER\}\}\n?/g, '');
  const htmlBody  = safeText.replace(/\n/g, '<br>');
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
  let totalBounced = 0;
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
      if (status === 'Bounced') totalBounced++;
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
  body += row('Dead addresses (bounced)', totalBounced);
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
    // Every thread costs a Gmail round trip below, so 500 is what makes this
    // time out. 50 is far more than you need to pick the right thread.
    threads = GmailApp.search(query, 0, 50);
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

  // Build every row in memory first, then write the whole block in ONE call.
  // The old version made 7 separate sheet calls per thread - 350 round trips
  // for 50 threads, which blew the 6-minute execution limit.
  const rows = threads.map(thread => {
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

    const lastDate    = thread.getLastMessageDate();
    const lastDateStr = lastDate.getFullYear() + '-' +
      String(lastDate.getMonth() + 1).padStart(2, '0') + '-' +
      String(lastDate.getDate()).padStart(2, '0');

    // Body: plain text of the FIRST message (the original cold email), 500 chars.
    let bodyPreview = '';
    try {
      const rawBody = msgs[0].getPlainBody() || stripHtml(msgs[0].getBody());
      bodyPreview   = rawBody.trim().replace(/\s+/g, ' ').substring(0, 500);
    } catch (e) {
      bodyPreview = '[Could not read body: ' + e.message + ']';
    }

    return [email, thread.getId(), thread.getFirstMessageSubject(),
            fromAccount, lastDateStr, bodyPreview, false];
  });

  // Make room. The old version wrote straight down from the found row, painting
  // over any rows already sitting below it. Inserting pushes them down instead,
  // so nothing already in the sheet is ever destroyed.
  if (rows.length > 1) sheet.insertRowsAfter(row, rows.length - 1);

  sheet.getRange(row, sc.leadEmail, rows.length, 7).setValues(rows);
  sheet.getRange(row, sc.enroll, rows.length, 1)
       .insertCheckboxes()
       .setDataValidation(checkboxValidation);
  sheet.getRange(row, sc.body, rows.length, 1).setWrap(true);

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

  // Today if today's sending window is still open, otherwise tomorrow.
  // Uses the same clock as the sending engine - see firstSendDate below.
  const tomorrow = firstSendDate();

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

  // ── Clean up InboxScanner ──────────────────────────────────
  // Every thread we found for this lead is now noise: the one that matters is
  // enrolled and tracked in ActiveFollowUps. Left alone this sheet fills up
  // fast, since it holds one row per thread per lead. Delete every row for
  // this email, including the one just ticked.
  //
  // Runs only AFTER the ActiveFollowUps insert succeeded above, so a failed
  // enrollment never destroys the scan results.
  try {
    const scanLastRow = sheet.getLastRow();
    if (scanLastRow >= 2) {
      const scanEmails = sheet.getRange(2, sc.leadEmail, scanLastRow - 1, 1).getValues();
      const target = leadEmail.toLowerCase();
      const hits = [];
      for (let i = 0; i < scanEmails.length; i++) {
        if (String(scanEmails[i][0]).trim().toLowerCase() === target) hits.push(i + 2);
      }
      // Delete in contiguous runs, bottom-up, so row numbers above stay valid
      // and 20 threads cost one delete call instead of twenty.
      let removed = 0;
      let end = hits.length - 1;
      while (end >= 0) {
        let start = end;
        while (start > 0 && hits[start - 1] === hits[start] - 1) start--;
        sheet.deleteRows(hits[start], end - start + 1);
        removed += end - start + 1;
        end = start - 1;
      }
      Logger.log('enrollLead: removed ' + removed + ' InboxScanner row(s) for ' + leadEmail);
    }
  } catch (e) {
    // The lead IS enrolled - cleanup failing is cosmetic, never block on it.
    Logger.log('enrollLead: InboxScanner cleanup failed (lead is still enrolled) - ' + e.message);
  }

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
  // A blank name would match an empty header cell - the spacer column between
  // two sequences - read an empty message, return null, and processSingleLead
  // would take that as "sequence finished" and mark the lead Done. The lead is
  // never emailed and nothing anywhere says why. Fail loudly instead: this
  // routes to the Error branch, which writes a note you can actually see.
  if (!String(sequenceName || '').trim()) {
    throw new Error('sequenceName is blank - set it in ActiveFollowUps');
  }
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
  // If the Drive image could not be fetched there is no <img> tag to swap in,
  // and without this the lead reads the literal text {{IMG_PLACEHOLDER}}
  // sitting in the middle of the email.
  const safeText  = imgBlob ? bodyText : bodyText.replace(/\{\{IMG_PLACEHOLDER\}\}\n?/g, '');
  const htmlBody  = safeText.replace(/\n/g, '<br>');

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
  // Drop blank entries before comparing. A single '' in this list would make
  // .includes() true for every sender, which reads every reply as our own.
  // If nothing resolves we return false, so a reply is DETECTED and the
  // follow-up is held - the safe direction to fail in.
  return CONFIG.sendingAccounts
    .filter(a => a && String(a).trim())
    .some(a => lower.includes(String(a).toLowerCase().trim()));
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

// ============================================================
// SECTION: PARKED LEADS (duplicate-send guard)
// ============================================================

function getParkedLeads() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('parkedLeads');
    if (!raw) return {};
    const obj = JSON.parse(raw);
    if (!obj || obj.date !== todayStr()) return {}; // stale - new day, start clean
    return obj.ids || {};
  } catch (e) {
    Logger.log('getParkedLeads ERROR: ' + e.message);
    return {};
  }
}

function parkLead(threadId, reason) {
  if (!threadId) return;
  try {
    const ids = getParkedLeads();
    // Cap the reason - a Script Property value maxes out near 9 KB, and long
    // error messages across many leads would silently overflow it, which would
    // take the duplicate-send guard down with it.
    ids[String(threadId)] = String(reason || 'parked').substring(0, 30);
    PropertiesService.getScriptProperties()
      .setProperty('parkedLeads', JSON.stringify({ date: todayStr(), ids: ids }));
    Logger.log('parkLead: ' + threadId + ' parked for today - ' + (reason || 'parked'));
  } catch (e) {
    Logger.log('parkLead ERROR for ' + threadId + ': ' + e.message);
  }
}

function cleanUpStuckTriggers() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processOneLead') {
      try { ScriptApp.deleteTrigger(t); removed++; }
      catch (e) { Logger.log('cleanUpStuckTriggers: ' + e.message); }
    }
  });
  // A chain that died from the trigger limit never got to clear its lock, and
  // that lock would block a restart for 30 min. This button is the "unstick
  // everything" button, so drop the lock too.
  clearChainLock();

  const left = ScriptApp.getProjectTriggers().length;
  const msg  = 'Removed ' + removed + ' stuck processOneLead trigger(s). ' +
               left + ' trigger(s) remain (Google allows 20).';
  Logger.log('cleanUpStuckTriggers: ' + msg);
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Trigger Cleanup', 8); } catch (e) {}
}

function notifyTriggerFailure(err) {
  let count = '?';
  try { count = ScriptApp.getProjectTriggers().length; } catch (e) {}
  const notifEmail = getSetting('notificationEmail');
  if (!notifEmail) return;
  try {
    MailApp.sendEmail({
      to: notifEmail,
      subject: 'Follow-Up System - CHAIN STOPPED, could not create trigger',
      body: 'The follow-up chain could not schedule its next step and has stopped.\n\n' +
        'Error: ' + err.message + '\n' +
        'Triggers currently in this project: ' + count + ' (Google allows 20)\n\n' +
        'If that number is at or near 20, run "Clean up stuck triggers" from the ' +
        'Follow-Up System menu, then start the chain again.\n\n' +
        'No leads were lost - unsent leads stay Active and are picked up on the next run.'
    });
  } catch (e) { Logger.log('notifyTriggerFailure: could not send email - ' + e.message); }
}

function scheduleContinueSoon() {
  try {
    ScriptApp.newTrigger('processOneLead')
      .timeBased()
      .after(60 * 1000)
      .create();
    Logger.log('scheduleContinueSoon: processOneLead continues in ~1 min (no email was sent)');
  } catch (e) {
    Logger.log('scheduleContinueSoon: FAILED to create trigger - ' + e.message);
    notifyTriggerFailure(e);
  }
}

// ============================================================
// SECTION: CHAIN LOCK (one chain at a time)
// ============================================================

// The lock stores the timestamp of the last chain activity. If a chain dies
// without clearing it, the lock goes stale on its own so it can never block
// you forever. Normal gap between chain steps is 4-8 min, so 30 is safe.
const CHAIN_LOCK_KEY      = 'chainRunningSince';
const CHAIN_LOCK_STALE_MS = 30 * 60 * 1000;

function chainIsRunning() {
  const raw = PropertiesService.getScriptProperties().getProperty(CHAIN_LOCK_KEY);
  if (!raw) return false;
  const age = Date.now() - parseInt(raw, 10);
  if (isNaN(age) || age > CHAIN_LOCK_STALE_MS) {
    Logger.log('chainIsRunning: stale lock (' + Math.round(age / 60000) + ' min old) - clearing it');
    clearChainLock();
    return false;
  }
  return true;
}

function markChainRunning() {
  try { PropertiesService.getScriptProperties().setProperty(CHAIN_LOCK_KEY, String(Date.now())); }
  catch (e) { Logger.log('markChainRunning ERROR: ' + e.message); }
}

function clearChainLock() {
  try { PropertiesService.getScriptProperties().deleteProperty(CHAIN_LOCK_KEY); }
  catch (e) { Logger.log('clearChainLock ERROR: ' + e.message); }
}

// ============================================================
// SECTION: DATE HELPERS (one clock, always the configured timezone)
// ============================================================

/**
 * Today in YOUR configured timezone, as YYYY-MM-DD.
 *
 * Same answer as todayStr() but with no network call, so it is safe inside
 * onEdit triggers and instant everywhere else. Never use plain new Date()
 * for a calendar date - that silently uses the Apps Script PROJECT timezone,
 * which is a different clock from the one the sending engine runs on.
 */
function todayInTz() {
  const tz = getSetting('timezone') || 'UTC';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

/**
 * Converts a value read out of a sheet cell into YYYY-MM-DD.
 *
 * Sheets stores a date you wrote as midnight in the SPREADSHEET's timezone,
 * so it must be read back in that same timezone. Formatting it with
 * getFullYear()/getMonth()/getDate() uses the script timezone instead and
 * shifts the date by a day whenever the two differ.
 */
function sheetDateToStr(v) {
  if (!(v instanceof Date)) return String(v || '').trim();
  try {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  } catch (e) {
    return Utilities.formatDate(v, getSetting('timezone') || 'UTC', 'yyyy-MM-dd');
  }
}

/**
 * Current time as HH:mm in YOUR configured timezone.
 * Network-free twin of currentTimeHHMM(), so it is safe inside onEdit
 * triggers and instant everywhere else. Same answer, no timeapi.io call.
 */
function nowInTz() {
  const tz = getSetting('timezone') || 'UTC';
  return Utilities.formatDate(new Date(), tz, 'HH:mm');
}

/**
 * The date a newly enrolled lead should first be contacted.
 *
 * Today if today's sending window has not closed yet, otherwise tomorrow.
 *
 * Enrolment used to always say "tomorrow", which made the wait depend on the
 * hour you happened to click: enrol at 00:01 and the first follow-up sat until
 * the NEXT day's window, roughly 33 hours; enrol at 23:00 and it went out 10
 * hours later. Same rule for both now - if there is still time today, use today.
 */
function firstSendDate() {
  const today  = todayInTz();
  const startT = String(getSetting('followUpStartTime') || '08:00').substring(0, 5);
  const endT   = String(getSetting('followUpEndTime')   || '18:00').substring(0, 5);

  // Midnight-crossing window (e.g. 17:00-02:00) is open on both sides of
  // midnight, so there is always time left today.
  if (startT > endT) return today;

  const now = nowInTz();
  return now < endT ? today : addDays(today, 1);
}

// ============================================================
// SECTION: REPLY AUTO-RESUME
// ============================================================

/**
 * Timestamp in the configured timezone, to the minute: 'yyyy-MM-dd HH:mm'.
 * Network free, so it is safe anywhere.
 */
function nowStampInTz() {
  const tz = getSetting('timezone') || 'UTC';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
}

/**
 * Called at the start of startDailyRun(), next to checkOooAutoResume().
 *
 * A lead who says "yeah, interesting" and then goes quiet is the most valuable
 * lead there is, and without this they sit at Replied forever. When
 * resumeOnReply is ticked for that lead, this puts them back into the sequence
 * after replyResumeDays of silence - continuing at the step they stopped on,
 * not starting over. Reply again and they pause again, and the cycle repeats
 * from the new date.
 *
 * replyResumeDays = 0 means resume on the next run. Useful for testing.
 */
function checkReplyAutoResume() {
  const resumeDays  = parseInt(getSetting('replyResumeDays') || '7');
  const activeSheet = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow     = activeSheet.getLastRow();
  if (lastRow < 2) return;

  const data  = activeSheet.getRange(2, 1, lastRow - 1, CONFIG.cols.notes).getValues();
  const today = todayStr();
  const c     = CONFIG.cols;

  data.forEach((row, i) => {
    if (String(row[c.status - 1]).trim() !== 'Replied') return;

    // Per-lead opt in. The column is a checkbox so it arrives as a real
    // boolean, but a hand-typed TRUE arrives as text - accept both.
    const flag    = row[c.resumeOnReply - 1];
    const optedIn = flag === true || String(flag).trim().toUpperCase() === 'TRUE';
    if (!optedIn) return;

    // When did they reply? Newest "Lead replied" note wins, so a lead who has
    // been round this loop before is measured from their LATEST reply.
    const notes = String(row[c.notes - 1] || '');
    const marks = notes.match(/\[(\d{4}-\d{2}-\d{2})\][^\n]*Lead replied/gi);
    const repliedOn = marks
      ? marks[marks.length - 1].match(/\[(\d{4}-\d{2}-\d{2})\]/)[1]
      : sheetDateToStr(row[c.lastSentDate - 1]);   // fallback for older rows
    if (!repliedOn) return;

    if (today < addDays(repliedOn, resumeDays)) return;   // still inside the quiet period

    const sheetRow = i + 2;
    const step     = Number(row[c.sequenceStep - 1]) || 0;
    activeSheet.getRange(sheetRow, c.status).setValue('Active');
    activeSheet.getRange(sheetRow, c.nextSendDate).setValue(today);

    // This note is the cutoff processSingleLead reads. It MUST be written, and
    // it MUST carry the time, or the lead is flipped straight back to Replied.
    const existing = activeSheet.getRange(sheetRow, c.notes).getValue();
    const note = '[' + nowStampInTz() + '] Resumed from Replied after ' + resumeDays +
                 ' days of silence - continuing at step ' + step + '.';
    activeSheet.getRange(sheetRow, c.notes).setValue(existing ? existing + '\n' + note : note);
    Logger.log('Reply auto-resumed: ' + String(row[c.leadEmail - 1]) +
               ' (replied ' + repliedOn + ', silent ' + resumeDays + ' days, step ' + step + ')');
  });

  SpreadsheetApp.flush();
}

// ============================================================
// SECTION: TESTING HELPERS
// ============================================================

/**
 * Forget that a lead was already handled today so the chain will pick it up
 * again on the next run.
 *
 * The duplicate-send guard lives in Script Properties, NOT in the sheet, so
 * editing the row by hand does not clear it. That is deliberate - it is what
 * makes a double send impossible even when a sheet write fails - but it means
 * you need this to re-test a lead on the same day.
 */
function unparkLead() {
  const THREAD_ID = '';   // blank = clear them all

  const props = PropertiesService.getScriptProperties();
  if (!THREAD_ID) {
    props.deleteProperty('parkedLeads');
    Logger.log('Cleared ALL parked leads for today.');
    return;
  }
  const raw = props.getProperty('parkedLeads');
  if (!raw) { Logger.log('Nothing is parked right now.'); return; }

  const obj = JSON.parse(raw);
  if (!obj.ids || !obj.ids[THREAD_ID]) {
    Logger.log(THREAD_ID + ' is not parked. Currently parked: ' + JSON.stringify(obj.ids || {}));
    return;
  }
  const was = obj.ids[THREAD_ID];
  delete obj.ids[THREAD_ID];
  props.setProperty('parkedLeads', JSON.stringify(obj));
  Logger.log('Unparked ' + THREAD_ID + ' (was: "' + was + '"). Still parked: ' + JSON.stringify(obj.ids));
}

/**
 * Prints exactly what the reply check sees for one lead.
 * Set LEAD_EMAIL, run, then read the log in Executions.
 */
function debugReplyCheck() {
  const LEAD_EMAIL = 'put-the-lead-email-here@example.com';

  let active, effective;
  try { active    = Session.getActiveUser().getEmail(); }    catch (e) { active    = 'THREW: ' + e.message; }
  try { effective = Session.getEffectiveUser().getEmail(); } catch (e) { effective = 'THREW: ' + e.message; }

  Logger.log('getActiveUser().getEmail()    = "' + active + '"');
  Logger.log('getEffectiveUser().getEmail() = "' + effective + '"');
  Logger.log('CONFIG.sendingAccounts        = ' + JSON.stringify(CONFIG.sendingAccounts));
  Logger.log('replyCheckBeforeSend setting  = "' + getSetting('replyCheckBeforeSend') + '"');
  Logger.log('replyResumeDays setting       = "' + getSetting('replyResumeDays') + '"');
  Logger.log('now in configured tz          = ' + nowStampInTz());

  const sheet   = getSheet(CONFIG.sheets.activeFollowUps);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No leads in ActiveFollowUps.'); return; }

  const c    = CONFIG.cols;
  const data = sheet.getRange(2, 1, lastRow - 1, c.notes).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][c.leadEmail - 1]).trim().toLowerCase() !==
        LEAD_EMAIL.trim().toLowerCase()) continue;

    const threadId = String(data[i][c.threadId - 1]).trim();
    Logger.log('--- row ' + (i + 2) + ' | ' + LEAD_EMAIL + ' | threadId ' + threadId);
    Logger.log('status=' + data[i][c.status - 1] + ' | resumeOnReply=' + data[i][c.resumeOnReply - 1] +
               ' | step=' + data[i][c.sequenceStep - 1]);
    Logger.log('notes:\n' + String(data[i][c.notes - 1] || '(empty)'));

    const thread = GmailApp.getThreadById(threadId);
    if (!thread) { Logger.log('THREAD NOT FOUND - the reply landed in a different thread.'); return; }

    const msgs = thread.getMessages();
    Logger.log('thread has ' + msgs.length + ' message(s):');
    msgs.forEach((m, n) => Logger.log('  [' + n + '] ' +
      Utilities.formatDate(m.getDate(), getSetting('timezone') || 'UTC', 'yyyy-MM-dd HH:mm') +
      '   from: ' + m.getFrom()));

    const from = msgs[msgs.length - 1].getFrom();
    Logger.log('latest sender            = ' + from);
    Logger.log('isSendingAccount(latest) = ' + isSendingAccount(from));
    return;
  }
  Logger.log('Lead ' + LEAD_EMAIL + ' not found in ActiveFollowUps.');
}

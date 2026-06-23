const SHEETS = {
  REGISTRATIONS: 'Registrations',
  SETTINGS: 'Settings',
  EMAIL_TEMPLATE: 'Email Template',
  REMINDER_TEMPLATE: 'Reminder Template',
  EMAIL_LOG: 'Email Log',
  REMINDER_LOG: 'Reminder Log',
  SYSTEM_LOG: 'System Log'
};

const COL = {
  TIMESTAMP: 1,
  FULL_NAME: 2,
  EMAIL: 3,
  EMAIL_NORMALIZED: 4,
  LEVEL: 5,
  BOOKING_STATUS: 6,
  EMAIL_STATUS: 7,
  EMAIL_SENT_AT: 8,
  SOURCE: 9,
  USER_AGENT: 10,
  ERROR: 11,
  NOTES: 12,
  SEND_REMINDER: 13,
  REMINDER_STATUS: 14,
  REMINDER_SENT_AT: 15,
  REMINDER_BATCH_ID: 16,
  REMINDER_NOTES: 17
};

const DATA_COLUMNS = 17;
const DEFAULT_EMAILS_PER_RUN = 30;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Session Tools')
    .addItem('Send Pending Confirmations', 'sendPendingEmails')
    .addSeparator()
    .addItem('Select All Confirmed For Reminder', 'selectAllConfirmedForReminder')
    .addItem('Clear Reminder Selection', 'clearReminderSelection')
    .addItem('Send Reminder To Selected', 'sendReminderToSelected')
    .addItem('Send Reminder To All Confirmed', 'sendReminderNow')
    .addItem('Reset Reminder Status For Selected', 'resetReminderStatusForSelected')
    .addSeparator()
    .addItem('Delete Selected Registration Rows', 'deleteSelectedRegistrationRows')
    .addItem('Clean Empty Rows', 'cleanEmptyRows')
    .addItem('Refresh Booking Cache', 'clearBookingCache')
    .addToUi();
}

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback || '';
  let result;

  try {
    const action = String(params.action || 'status').toLowerCase();
    if (action === 'register') result = register_(params);
    else if (action === 'status') result = getStatus_();
    else result = { status: 'error', message: 'Unknown action.' };
  } catch (err) {
    logSystem_('doGet', 'ERROR', String(err && err.stack ? err.stack : err), JSON.stringify(params), 'webapp');
    result = { status: 'error', message: 'حدث خطأ مؤقت. جرّب مرة أخرى.' };
  }

  return output_(result, callback);
}

function register_(params) {
  const name = clean_(params.name);
  const email = clean_(params.email);
  const emailNormalized = normalizeEmail_(email);
  const level = clean_(params.level);
  const source = clean_(params.source || 'landing-page');
  const userAgent = clean_(params.userAgent);

  if (name.length < 2) return { status: 'error', message: 'اكتب اسمك الكامل.' };
  if (!isValidEmail_(emailNormalized)) return { status: 'error', message: 'اكتب بريد إلكتروني صحيح.' };
  if (!['Beginner', 'Junior', 'Middle Level', 'Senior'].includes(level)) return { status: 'error', message: 'اختار مستواك الحالي.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { status: 'error', message: 'في ضغط على التسجيل. جرّب مرة أخرى.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
    const settings = getSettings_();
    const totalSeats = Number(settings.totalSeats) || 50;
    const siteMode = String(settings.siteMode || 'registration');

    if (siteMode === 'closed') return withStats_({ status: 'closed', message: 'التسجيل مغلق الآن.' }, settings);
    if (siteMode === 'live') return withStats_({ status: 'closed', message: settings.siteLiveMessage || 'السيشن بدأت. راجع إيميلك للدخول.' }, settings);
    if (emailExists_(sheet, emailNormalized)) return withStats_({ status: 'duplicate', message: 'الإيميل ده مسجل بالفعل.' }, settings);

    const confirmed = countConfirmed_(sheet);
    const remaining = Math.max(totalSeats - confirmed, 0);
    if (remaining <= 0) return withStats_({ status: 'full', message: 'اكتمل عدد المقاعد.' }, settings);

    const newRow = getNextEmptyRegistrationRow_(sheet);
    sheet.getRange(newRow, 1, 1, DATA_COLUMNS).setValues([[
      new Date(),
      name,
      email,
      emailNormalized,
      level,
      'CONFIRMED',
      'PENDING',
      '',
      source,
      userAgent,
      '',
      '',
      false,
      '',
      '',
      '',
      ''
    ]]);
    sheet.getRange(newRow, COL.SEND_REMINDER).insertCheckboxes();

    clearBookingCache();

    return withStats_({
      status: 'success',
      message: 'تم تأكيد الحجز.',
      popupTitle: settings.successPopupTitle || 'تم تأكيد الحجز',
      popupText: settings.successPopupText || 'سيتم إرسال تفاصيل الحضور على بريدك الإلكتروني خلال دقائق.'
    }, settings, confirmed + 1);
  } finally {
    lock.releaseLock();
  }
}

function getNextEmptyRegistrationRow_(sheet) {
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < 2) sheet.insertRowsAfter(1, 50);

  const maxRows = sheet.getMaxRows();
  const emailValues = sheet.getRange(2, COL.EMAIL_NORMALIZED, maxRows - 1, 1).getValues();

  for (let i = 0; i < emailValues.length; i++) {
    const email = String(emailValues[i][0] || '').trim();
    if (!email) return i + 2;
  }

  sheet.insertRowsAfter(maxRows, 50);
  return maxRows + 1;
}

function getStatus_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('booking_status');
  if (cached) return JSON.parse(cached);

  const settings = getSettings_();
  const result = withStats_({ status: 'ok' }, settings);
  cache.put('booking_status', JSON.stringify(result), 20);
  return result;
}

function withStats_(base, settings, confirmedOverride) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
  const totalSeats = Number(settings.totalSeats) || 50;
  const confirmed = typeof confirmedOverride === 'number' ? confirmedOverride : countConfirmed_(sheet);

  return Object.assign({}, base, {
    totalSeats,
    confirmed,
    remaining: Math.max(totalSeats - confirmed, 0),
    siteMode: String(settings.siteMode || 'registration'),
    sessionTitle: settings.sessionTitle || '',
    sessionDate: settings.sessionDate || '',
    successPopupTitle: settings.successPopupTitle || 'تم تأكيد الحجز',
    successPopupText: settings.successPopupText || 'سيتم إرسال تفاصيل الحضور على بريدك الإلكتروني خلال دقائق.'
  });
}

function sendPendingEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
    const settings = getSettings_();
    const template = getTemplate_(SHEETS.EMAIL_TEMPLATE);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const maxPerRun = Number(settings.confirmationBatchSize) || DEFAULT_EMAILS_PER_RUN;
    const values = sheet.getRange(2, 1, lastRow - 1, DATA_COLUMNS).getValues();
    let sentCount = 0;

    for (let i = 0; i < values.length; i++) {
      if (sentCount >= maxPerRun) break;

      const row = values[i];
      const rowNumber = i + 2;
      const name = row[COL.FULL_NAME - 1];
      const email = row[COL.EMAIL - 1];
      const bookingStatus = row[COL.BOOKING_STATUS - 1];
      const emailStatus = row[COL.EMAIL_STATUS - 1];

      if (bookingStatus !== 'CONFIRMED' || emailStatus !== 'PENDING') continue;

      try {
        const subject = sendTemplatedEmail_(email, name, settings, template, 'confirmation');
        sheet.getRange(rowNumber, COL.EMAIL_STATUS).setValue('SENT');
        sheet.getRange(rowNumber, COL.EMAIL_SENT_AT).setValue(new Date());
        logEmail_(email, subject, 'SENT', '', rowNumber);
        sentCount++;
      } catch (err) {
        const error = String(err && err.message ? err.message : err);
        sheet.getRange(rowNumber, COL.EMAIL_STATUS).setValue('FAILED');
        sheet.getRange(rowNumber, COL.ERROR).setValue(error);
        logEmail_(email, template.emailSubject || 'تأكيد حجزك في السيشن', 'FAILED', error, rowNumber);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function selectAllConfirmedForReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
  const settings = getSettings_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const skipAlreadySent = String(settings.reminderSkipAlreadySent || 'TRUE').toUpperCase() === 'TRUE';
  const values = sheet.getRange(2, 1, lastRow - 1, DATA_COLUMNS).getValues();
  const selection = [];
  const statuses = [];

  values.forEach(row => {
    const bookingStatus = row[COL.BOOKING_STATUS - 1];
    const reminderStatus = row[COL.REMINDER_STATUS - 1];
    const eligible = bookingStatus === 'CONFIRMED' && !(skipAlreadySent && reminderStatus === 'SENT');
    selection.push([eligible]);
    statuses.push([eligible && !reminderStatus ? 'READY' : reminderStatus]);
  });

  sheet.getRange(2, COL.SEND_REMINDER, selection.length, 1).setValues(selection);
  sheet.getRange(2, COL.REMINDER_STATUS, statuses.length, 1).setValues(statuses);
  SpreadsheetApp.getUi().alert('تم تحديد المسجلين المؤهلين للـ reminder.');
}

function clearReminderSelection() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.REGISTRATIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = Array.from({ length: lastRow - 1 }, () => [false]);
  sheet.getRange(2, COL.SEND_REMINDER, lastRow - 1, 1).setValues(values);
  SpreadsheetApp.getUi().alert('تم مسح تحديد الـ reminder.');
}

function sendReminderToSelected() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Send Reminder To Selected',
    'هيتم إرسال Reminder فقط للصفوف المحددة TRUE في عمود Send Reminder?. متأكد؟',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const result = sendReminderBatch_({ selectedOnly: true });
  ui.alert('Reminder Result', `تم الإرسال: ${result.sent}\nتم تخطيهم: ${result.skipped}\nفشل: ${result.failed}`, ui.ButtonSet.OK);
}

function sendReminderNow() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Send Reminder To All Confirmed',
    'هيتم إرسال Reminder لكل المسجلين CONFIRMED الذين لم يُرسل لهم Reminder قبل كده. متأكد؟',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const result = sendReminderBatch_({ selectedOnly: false });
  ui.alert('Reminder Result', `تم الإرسال: ${result.sent}\nتم تخطيهم: ${result.skipped}\nفشل: ${result.failed}`, ui.ButtonSet.OK);
}

function sendReminderBatch_(options) {
  options = options || {};
  const selectedOnly = options.selectedOnly === true;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { sent: 0, skipped: 0, failed: 0, message: 'locked' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
    const settings = getSettings_();
    const template = getTemplate_(SHEETS.REMINDER_TEMPLATE);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { sent: 0, skipped: 0, failed: 0 };

    const maxPerRun = Number(settings.reminderBatchSize) || 80;
    const skipAlreadySent = String(settings.reminderSkipAlreadySent || 'TRUE').toUpperCase() === 'TRUE';
    const onlyConfirmed = String(settings.reminderOnlyConfirmed || 'TRUE').toUpperCase() === 'TRUE';
    const values = sheet.getRange(2, 1, lastRow - 1, DATA_COLUMNS).getValues();
    const batchId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');

    let sent = 0;
    let skipped = 0;
    let failed = 0;

    for (let i = 0; i < values.length; i++) {
      if (sent >= maxPerRun) break;

      const row = values[i];
      const rowNumber = i + 2;
      const name = row[COL.FULL_NAME - 1];
      const email = row[COL.EMAIL - 1];
      const bookingStatus = row[COL.BOOKING_STATUS - 1];
      const selected = row[COL.SEND_REMINDER - 1] === true;
      const reminderStatus = row[COL.REMINDER_STATUS - 1];

      if (!email) { skipped++; continue; }
      if (onlyConfirmed && bookingStatus !== 'CONFIRMED') { skipped++; continue; }
      if (selectedOnly && !selected) { skipped++; continue; }
      if (skipAlreadySent && reminderStatus === 'SENT') { skipped++; continue; }

      try {
        const subject = sendTemplatedEmail_(email, name, settings, template, 'reminder');
        sheet.getRange(rowNumber, COL.SEND_REMINDER).setValue(false);
        sheet.getRange(rowNumber, COL.REMINDER_STATUS).setValue('SENT');
        sheet.getRange(rowNumber, COL.REMINDER_SENT_AT).setValue(new Date());
        sheet.getRange(rowNumber, COL.REMINDER_BATCH_ID).setValue(batchId);
        sheet.getRange(rowNumber, COL.REMINDER_NOTES).setValue('');
        logReminder_(email, subject, 'SENT', '', rowNumber, batchId);
        sent++;
      } catch (err) {
        const error = String(err && err.message ? err.message : err);
        sheet.getRange(rowNumber, COL.REMINDER_STATUS).setValue('FAILED');
        sheet.getRange(rowNumber, COL.REMINDER_NOTES).setValue(error);
        logReminder_(email, template.reminderSubject || 'تذكير بالسيشن', 'FAILED', error, rowNumber, batchId);
        failed++;
      }
    }

    return { sent, skipped, failed };
  } finally {
    lock.releaseLock();
  }
}

function resetReminderStatusForSelected() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Reminder Status',
    'هيتم مسح Reminder Status للصفوف المحددة TRUE فقط. استخدمها للاختبار أو إعادة الإرسال. متأكد؟',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.REGISTRATIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, DATA_COLUMNS).getValues();
  let reset = 0;

  values.forEach((row, index) => {
    if (row[COL.SEND_REMINDER - 1] === true) {
      const rowNumber = index + 2;
      sheet.getRange(rowNumber, COL.REMINDER_STATUS, 1, 4).clearContent();
      reset++;
    }
  });

  ui.alert(`تم مسح حالة الـ reminder لعدد ${reset} صف.`);
}

function deleteSelectedRegistrationRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEETS.REGISTRATIONS) {
    ui.alert('افتح Tab اسمها Registrations الأول، وحدد الصفوف اللي عايز تمسحها.');
    return;
  }

  const range = sheet.getActiveRange();
  if (!range) {
    ui.alert('حدد الصفوف اللي عايز تمسحها الأول.');
    return;
  }

  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const endRow = startRow + numRows - 1;

  if (startRow === 1) {
    ui.alert('مينفعش تمسح صف العناوين Header.');
    return;
  }

  const response = ui.alert(
    'Delete Selected Registration Rows',
    `هيتم حذف الصفوف من ${startRow} إلى ${endRow}. متأكد؟`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  for (let row = endRow; row >= startRow; row--) {
    if (row > 1) sheet.deleteRow(row);
  }

  clearBookingCache();
  ui.alert('تم حذف الصفوف المحددة وتحديث عداد المقاعد.');
}

function cleanEmptyRows() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(2, 1, lastRow - 1, DATA_COLUMNS);
    const values = range.getValues();
    const kept = values.filter(row => rowHasMeaningfulData_(row));

    range.clearContent();
    if (kept.length) {
      sheet.getRange(2, 1, kept.length, DATA_COLUMNS).setValues(kept);
      for (let i = 0; i < kept.length; i++) {
        sheet.getRange(i + 2, COL.SEND_REMINDER).insertCheckboxes();
      }
    }

    clearBookingCache();
    SpreadsheetApp.getUi().alert(`تم ترتيب الصفوف. عدد الصفوف المحفوظة: ${kept.length}`);
  } finally {
    lock.releaseLock();
  }
}

function rowHasMeaningfulData_(row) {
  return row.some((value, index) => {
    if (index === COL.SEND_REMINDER - 1 && value === false) return false;
    return value !== '' && value !== null && value !== false;
  });
}

function sendTemplatedEmail_(email, name, settings, template, type) {
  const data = {
    name: name || '',
    sessionTitle: settings.sessionTitle || 'أفضل Roadmap لـ UI/UX في عصر الـ AI',
    sessionDate: settings.sessionDate || 'سيتم تأكيد الموعد قريبًا',
    meetingLink: settings.meetingLink || '',
    senderName: settings.senderName || 'Ahmed Hussein Community'
  };

  const isReminder = type === 'reminder';
  const subject = render_(
    template[isReminder ? 'reminderSubject' : 'emailSubject'] || (isReminder ? 'تذكير بالسيشن' : 'تأكيد حجزك في السيشن'),
    data
  );
  const title = render_(
    template[isReminder ? 'reminderTitle' : 'emailTitle'] || (isReminder ? 'تذكير بسيشن اليوم' : 'تم تأكيد حجزك بنجاح'),
    data
  );
  const intro = render_(
    template[isReminder ? 'reminderIntro' : 'emailIntro'] || `أهلًا {{name}}، ${isReminder ? 'بنذكرك بسيشن {{sessionTitle}}.' : 'تم تأكيد حجزك في سيشن {{sessionTitle}}.'}`,
    data
  );
  const details = render_(template[isReminder ? 'reminderDetails' : 'emailDetails'] || 'الموعد: {{sessionDate}}', data);
  const buttonText = render_(template.buttonText || 'الدخول إلى السيشن', data);
  const footerText = render_(template.footerText || '{{senderName}}', data);

  const plainBody = `${intro}\n\n${details}\n\nرابط الدخول:\n${data.meetingLink}\n\n${footerText}`;
  const htmlBody = buildEmailHtml_({ title, intro, details, buttonText, footerText, meetingLink: data.meetingLink });

  GmailApp.sendEmail(email, subject, plainBody, { name: data.senderName, htmlBody });
  return subject;
}

function buildEmailHtml_(data) {
  const detailsHtml = escapeHtml_(data.details).replace(/\n/g, '<br>');
  const footerHtml = escapeHtml_(data.footerText).replace(/\n/g, '<br>');
  const link = escapeHtml_(data.meetingLink || '#');

  return `
  <div dir="rtl" style="margin:0;padding:0;background:#f6f2e8;font-family:Arial,Tahoma,sans-serif;color:#111827;line-height:1.9">
    <div style="max-width:620px;margin:0 auto;padding:28px 16px">
      <div style="background:#07111F;border-radius:24px;padding:2px;background-image:linear-gradient(135deg,#F4DFA5,#D6B76A,#B98D39)">
        <div style="background:#07111F;border-radius:22px;padding:28px 22px;text-align:right">
          <div style="color:#D6B76A;font-size:13px;font-weight:700;margin-bottom:12px">Ahmed Hussein Community</div>
          <h1 style="margin:0 0 14px;color:#fff;font-size:26px;line-height:1.4">${escapeHtml_(data.title)}</h1>
          <p style="margin:0 0 18px;color:#D8DEE8;font-size:16px">${escapeHtml_(data.intro)}</p>
          <div style="background:rgba(255,255,255,.06);border:1px solid rgba(244,223,165,.25);border-radius:16px;padding:16px;margin:18px 0;color:#F8F4EA;font-size:15px">${detailsHtml}</div>
          <div style="text-align:center;margin:24px 0">
            <a href="${link}" style="display:inline-block;background:linear-gradient(135deg,#F4DFA5,#D6B76A,#B98D39);color:#111319;text-decoration:none;font-weight:800;padding:13px 24px;border-radius:14px;font-size:16px">${escapeHtml_(data.buttonText)}</a>
          </div>
          <p style="margin:12px 0 0;color:#A9B4C3;font-size:13px;text-align:center">لو الزر مش شغال، استخدم الرابط ده:</p>
          <p style="margin:6px 0 20px;text-align:center;direction:ltr"><a href="${link}" style="color:#F4DFA5;text-decoration:underline">${link}</a></p>
          <div style="border-top:1px solid rgba(244,223,165,.18);padding-top:16px;color:#A9B4C3;font-size:14px">${footerHtml}</div>
        </div>
      </div>
    </div>
  </div>`;
}

function getSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 2).getValues();
  const settings = {};
  values.forEach(([key, value]) => {
    if (key) settings[String(key).trim()] = value;
  });
  return settings;
}

function getTemplate_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 2).getValues();
  const template = {};
  values.forEach(([key, value]) => {
    if (key) template[String(key).trim()] = String(value || '');
  });
  return template;
}

function emailExists_(sheet, emailNormalized) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < 2) return false;
  const emails = sheet.getRange(2, COL.EMAIL_NORMALIZED, maxRows - 1, 1).getValues();
  return emails.some(row => String(row[0] || '').trim().toLowerCase() === emailNormalized);
}

function countConfirmed_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < 2) return 0;
  const statuses = sheet.getRange(2, COL.BOOKING_STATUS, maxRows - 1, 1).getValues();
  return statuses.filter(row => row[0] === 'CONFIRMED').length;
}

function output_(payload, callback) {
  const json = JSON.stringify(payload);
  if (callback) return ContentService.createTextOutput(`${callback}(${json});`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function logEmail_(email, subject, status, error, rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.EMAIL_LOG);
  sheet.appendRow([new Date(), email, subject, status, error, rowNumber]);
}

function logReminder_(email, subject, status, error, rowNumber, batchId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.REMINDER_LOG);
  sheet.appendRow([new Date(), email, subject, status, error, rowNumber, batchId, Session.getActiveUser().getEmail()]);
}

function logSystem_(action, status, message, payload, source) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.SYSTEM_LOG);
    sheet.appendRow([new Date(), action, status, message, payload, source]);
  } catch (_) {}
}

function clearBookingCache() {
  CacheService.getScriptCache().remove('booking_status');
}

function render_(template, data) {
  return String(template || '').replace(/{{\s*(\w+)\s*}}/g, (_, key) => data[key] != null ? String(data[key]) : '');
}

function clean_(value) {
  return String(value || '').trim();
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * ============================================================
 *  أتمتة جدول مراجعة الجودة (QA) - قراءة تلقائية من جدول CSM
 * ============================================================
 *
 * الفكرة:
 * هذا السكريبت يُثبَّت على ملف "جدول الجودة" (مو ملف الـ CSM).
 * كل ما يشتغل (يدويًا أو بجدولة يومية)، يفتح ملف الـ CSM بصلاحية
 * قراءة فقط (SpreadsheetApp.openById) — ما يعدل ولا يحذف ولا
 * يضيف أي شي بملف الـ CSM إطلاقًا.
 *
 * يكتشف تلقائيًا آخر يوم مُضاف بجدول الـ CSM، ولو كان يوم جديد
 * لم يُعالَج من قبل، يبني له جدول مراجعة جودة كامل ويضيفه
 * كأعمدة جديدة بجانب آخر يوم بملف الجودة.
 *
 * ============================================================
 *  الإعداد المطلوب (مرة وحدة):
 * ============================================================
 *
 * 1) افتح ملف "جدول الجودة" (14qfbh60-...).
 * 2) من القائمة: Extensions > Apps Script
 * 3) الصق هذا الكود كامل.
 * 4) عدّل CSM_SHEET_ID تحت لو تغيّر رابط ملف الـ CSM.
 * 5) من أيقونة الساعة (Triggers) أضف Trigger:
 *      - Function: runDailyQAAutomation
 *      - Event source: Time-driven
 *      - Type: Day timer، الوقت المفضل (مثلاً 9-10 صباحًا)
 * 6) جرّب أول مرة بتشغيل runDailyQAAutomation يدويًا من المحرر.
 *
 * ============================================================
 *  قاعدة تصنيف الجلسات (عدّلها هنا لو ظهر اسم حملة جديد):
 * ============================================================
 */

// ============ الإعدادات ============

const CSM_SHEET_ID = 'PASTE_CSM_SHEET_ID_HERE'; // من رابط ملف الـ CSM
const QA_SHEET_NAME = 'ورقة1'; // اسم التبويب داخل ملف الجودة (عدّله لاسم تبويبك الفعلي)

// فرق الأيام بين تاريخ الأودت وتاريخ مكالمة الـ CSM المُراجَعة
const AUDIT_DAY_OFFSET = 2;

// أنواع الجلسات الداخلية (اجتماعات/استراحات) - لا تُراجع أبدًا
const INTERNAL_KEYWORDS = [
  'CSM Daily Standup', 'CSM Handover', 'Task CSM Automation',
  'Monshaat Review', 'Pray', 'Break', 'Intercom', 'New Client',
  'Webinar Training', 'Side Meeting', 'NPS', 'Thaki Digital PMO',
  'Drop ', 'Not Activated', 'No Contact Clients', 'Zatca',
  'Grubtech', 'Salla & Zid', 'Main Stage', 'Sentiment',
];

// قاعدة التصنيف: كل نمط (يبحث كـ "يحتوي على") يربط بنوع معيار ومدة مراجعة بالدقائق
// الترتيب مهم: يفحص من الأعلى للأسفل، وأول تطابق يفوز
const CLASSIFICATION_RULES = [
  { pattern: /open\.?cx/i, type: 'Chat — Open CX', reviewMinutes: 10 },
  { pattern: /^cases/i, type: 'Case Handling', reviewMinutes: 20 },
  { pattern: /reactivation case/i, type: 'Case Handling', reviewMinutes: 20 },
  // أنماط Reactivation (كوهورت/ألوان/أشهر معينة) - عدّل هنا لو ظهر اسم جديد
  { pattern: /cohort|orange|yellow|gray|black|jan\b|feb\b|red\b|q1|q2/i, type: 'Reactivation', reviewMinutes: 12 },
  // أي شي فيه "Activation" وما طابق النمط أعلاه (شهر عادي زي Aug/July/June Activation)
  { pattern: /activation/i, type: 'Activation', reviewMinutes: 15 },
];

const TYPE_COLORS = {
  'Chat — Open CX': { bg: '#E1F5EE', text: '#04342C' },
  'Case Handling': { bg: '#FAECE7', text: '#4A1B0C' },
  'Activation': { bg: '#E6F1FB', text: '#042C53' },
  'Reactivation': { bg: '#FAEEDA', text: '#412402' },
  'admin': { bg: '#F1EFE8', text: '#2C2C2A' },
  'UNKNOWN': { bg: '#FCEBEB', text: '#791F1F' }, // أحمر تحذيري - نوع غير معروف يحتاج مراجعة يدوية
};

// ============ نهاية الإعدادات ============


/**
 * الدالة الرئيسية - شغّلها يدويًا أول مرة، وبعدها اربطها بـ Trigger يومي
 */
function runDailyQAAutomation() {
  const props = PropertiesService.getScriptProperties();
  const lastProcessed = props.getProperty('LAST_PROCESSED_CSM_DATE');

  const csmData = findLatestCsmDay(lastProcessed);
  if (!csmData) {
    Logger.log('لا يوجد يوم جديد بجدول الـ CSM لم تتم معالجته بعد.');
    return;
  }

  Logger.log('يوم جديد وُجد: ' + csmData.dateLabel);

  const sessions = extractDaySessions(csmData.sheet, csmData.headerRow, csmData.col);
  const classified = classifySessions(sessions);
  const schedule = buildQaSchedule(classified);

  const auditDate = addDays(csmData.dateObj, AUDIT_DAY_OFFSET);
  writeScheduleToQaSheet(schedule, auditDate, csmData.dateLabel);

  props.setProperty('LAST_PROCESSED_CSM_DATE', csmData.dateLabel);
  Logger.log('تمت الإضافة بنجاح لتاريخ الأودت: ' + Utilities.formatDate(auditDate, Session.getScriptTimeZone(), 'dd/MMM/yyyy'));
}

/**
 * يفتح ملف الـ CSM بصلاحية قراءة فقط ويلقى آخر يوم مضاف
 * (أحدث تاريخ موجود بكل الشيت) الذي لم تتم معالجته من قبل
 */
function findLatestCsmDay(lastProcessedLabel) {
  const ss = SpreadsheetApp.openById(CSM_SHEET_ID); // قراءة فقط - لا كتابة إطلاقًا على هذا الملف
  const sheets = ss.getSheets();

  let best = null; // { dateObj, dateLabel, sheet, headerRow, col }
  const dateRegex = /^(\d{1,2})\/([A-Za-z]{3})\/(\d{4})$/;
  const monthMap = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };

  sheets.forEach(function (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const val = String(data[r][c] || '').trim();
        const m = val.match(dateRegex);
        if (m) {
          const dateObj = new Date(parseInt(m[3]), monthMap[m[2]], parseInt(m[1]));
          if (!best || dateObj > best.dateObj) {
            best = { dateObj: dateObj, dateLabel: val, sheet: sheet, headerRow: r, col: c, allData: data };
          }
        }
      }
    }
  });

  if (!best) return null;
  if (lastProcessedLabel && best.dateLabel === lastProcessedLabel) return null; // مُعالَج من قبل
  return best;
}

/**
 * يستخرج قائمة الجلسات (الاسم + المدة بالدقائق) لعمود يوم معين،
 * بدمج الصفوف المتتالية لنفس الجلسة في مدة واحدة تراكمية
 */
function extractDaySessions(sheet, headerRow, dateCol) {
  const lastRow = sheet.getLastRow();
  const nameCol = dateCol; // اسم الجلسة عادة بنفس عمود التاريخ بالصف اللي تحته
  const durCol = dateCol + 1; // عمود المدة يليه غالبًا (Time)

  const sessions = [];
  let current = null;

  for (let r = headerRow + 1; r <= lastRow; r++) {
    const name = String(sheet.getRange(r + 1, nameCol + 1).getValue() || '').trim();
    const durText = String(sheet.getRange(r + 1, durCol + 1).getValue() || '').trim();

    if (!name) continue;
    // توقف لو وصلنا لصف تاريخ جديد (بداية يوم/أسبوع ثاني)
    if (/^\d{1,2}\/[A-Za-z]{3}\/\d{4}$/.test(name)) break;

    const durMin = parseDurationToMinutes(durText);

    if (current && current.name === name) {
      current.minutes += durMin || 15; // لو ما فيه مدة صريحة بالصف، افترض ربع ساعة (خانة واحدة)
    } else {
      if (current) sessions.push(current);
      current = { name: name, minutes: durMin || 15 };
    }
  }
  if (current) sessions.push(current);
  return sessions;
}

function parseDurationToMinutes(text) {
  if (!text) return null;
  let total = 0;
  const hourMatch = text.match(/(\d+)\s*hour/i);
  const minMatch = text.match(/(\d+)\s*min/i);
  if (hourMatch) total += parseInt(hourMatch[1]) * 60;
  if (minMatch) total += parseInt(minMatch[1]);
  return total > 0 ? total : null;
}

/**
 * يصنف كل جلسة: يستبعد الداخلية، ويحدد نوع المعيار للباقي
 */
function classifySessions(sessions) {
  const result = [];
  sessions.forEach(function (s) {
    const isInternal = INTERNAL_KEYWORDS.some(function (kw) {
      return s.name.toLowerCase().indexOf(kw.toLowerCase()) !== -1;
    });
    if (isInternal) return;

    let matchedType = 'UNKNOWN';
    let reviewMinutes = 15;
    for (let i = 0; i < CLASSIFICATION_RULES.length; i++) {
      if (CLASSIFICATION_RULES[i].pattern.test(s.name)) {
        matchedType = CLASSIFICATION_RULES[i].type;
        reviewMinutes = CLASSIFICATION_RULES[i].reviewMinutes;
        break;
      }
    }
    result.push({ name: s.name.trim(), type: matchedType, reviewMinutes: reviewMinutes });
  });
  return result;
}

/**
 * يبني جدول اليوم الكامل: 3 جولات مراجعة + بريك + إداري،
 * موزع من 9 صباحًا إلى 6 مساءً
 */
function buildQaSchedule(classifiedSessions) {
  const rows = [];
  let cursorMinutes = 9 * 60; // 9:00 صباحًا بالدقائق من منتصف الليل

  function addRow(label, type, minutes) {
    const start = cursorMinutes;
    const end = cursorMinutes + minutes;
    rows.push({
      hourLabel: minutesToHourRange(start),
      label: label,
      type: type,
      minutes: minutes,
    });
    cursorMinutes = end;
  }

  for (let round = 1; round <= 3; round++) {
    classifiedSessions.forEach(function (s) {
      addRow('Audit ' + round + ' - ' + s.name, s.type, s.reviewMinutes);
    });
    if (round === 1) addRow('غداء وصلاة الظهر (بريك)', 'admin', 30);
    if (round === 2) addRow('صلاة العصر', 'admin', 15);
  }

  addRow('تحديث لوحة الأداء وتسجيل السكورات', 'admin', 45);
  addRow('مراجعة الأخطاء الحرجة وتنبيهات سلاك', 'admin', 30);

  return rows;
}

function minutesToHourRange(startMinutes) {
  const h1 = Math.floor(startMinutes / 60);
  const h2 = h1 + 1;
  const fmt = function (h) {
    const hh = h > 12 ? h - 12 : h;
    return hh;
  };
  return fmt(h1) + ' - ' + fmt(h2);
}

/**
 * يكتب الجدول كأعمدة جديدة في ملف الجودة، بجانب آخر يوم موجود.
 * لو التبويب الحالي فيه 7 أيام (أسبوع كامل)، ينشئ صف مجموعة جديد تحت.
 */
function writeScheduleToQaSheet(scheduleRows, auditDate, csmDateLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(QA_SHEET_NAME);
  if (!sheet) {
    throw new Error('ما لقيت تبويب باسم "' + QA_SHEET_NAME + '" - تأكد من QA_SHEET_NAME بالإعدادات');
  }

  const lastCol = sheet.getLastColumn();
  const startCol = lastCol > 0 ? lastCol + 1 : 1;

  const auditDateStr = Utilities.formatDate(auditDate, Session.getScriptTimeZone(), 'dd/MMM/yyyy');
  const dayName = Utilities.formatDate(auditDate, Session.getScriptTimeZone(), 'EEE').toUpperCase();

  // عنوان
  sheet.getRange(1, startCol, 1, 4).merge().setValue('QA CSM')
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(2, startCol, 1, 4).merge().setValue('دوام فريق الجودة: 9:00 ص - 6:00 م')
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');

  const headers = [dayName, auditDateStr, 'Tim', 'Total Quality Reviews'];
  headers.forEach(function (h, i) {
    sheet.getRange(3, startCol + i).setValue(h)
      .setBackground('#943126').setFontColor('#FFFFFF').setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  let r = 4;
  let groupStartRow = r;
  let lastHourLabel = null;

  scheduleRows.forEach(function (row) {
    const style = TYPE_COLORS[row.type] || TYPE_COLORS['UNKNOWN'];

    if (row.hourLabel !== lastHourLabel) {
      if (lastHourLabel !== null && r - 1 >= groupStartRow) {
        sheet.getRange(groupStartRow, startCol, r - groupStartRow, 1).merge();
      }
      groupStartRow = r;
      lastHourLabel = row.hourLabel;
    }
    sheet.getRange(r, startCol).setValue(row.hourLabel)
      .setHorizontalAlignment('center').setFontWeight('bold');

    sheet.getRange(r, startCol + 1).setValue(row.label)
      .setBackground(style.bg).setFontColor(style.text).setFontWeight('bold')
      .setHorizontalAlignment('center');

    sheet.getRange(r, startCol + 2).setValue(row.minutes)
      .setBackground(style.bg).setFontColor(style.text)
      .setHorizontalAlignment('center');

    r++;
  });
  if (r - 1 >= groupStartRow) {
    sheet.getRange(groupStartRow, startCol, r - groupStartRow, 1).merge();
  }

  // ملاحظة مرجعية: أي يوم CSM تمت مراجعته
  sheet.getRange(r + 1, startCol, 1, 2).merge()
    .setValue('يراجع مكالمات يوم: ' + csmDateLabel)
    .setFontStyle('italic').setFontColor('#666666');

  Logger.log('تمت كتابة الجدول بالأعمدة من ' + startCol + ' إلى ' + (startCol + 3));
}

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * دالة اختبار - تعرض بالـ Log وش راح يستخرجه السكريبت بدون ما يكتب فعليًا
 * شغّلها للتأكد قبل ما تفعّل التريقر التلقائي
 */
function testDryRun() {
  const csmData = findLatestCsmDay(null);
  if (!csmData) {
    Logger.log('ما لقيت أي تاريخ بالملف');
    return;
  }
  Logger.log('آخر يوم موجود: ' + csmData.dateLabel);

  const sessions = extractDaySessions(csmData.sheet, csmData.headerRow, csmData.col);
  Logger.log('الجلسات الخام:');
  sessions.forEach(function (s) { Logger.log('  ' + s.name + ' - ' + s.minutes + ' دقيقة'); });

  const classified = classifySessions(sessions);
  Logger.log('بعد التصنيف (استبعاد الداخلية):');
  classified.forEach(function (s) { Logger.log('  [' + s.type + '] ' + s.name + ' (مراجعة ' + s.reviewMinutes + ' د)'); });

  const unknowns = classified.filter(function (s) { return s.type === 'UNKNOWN'; });
  if (unknowns.length) {
    Logger.log('⚠️ تحذير: فيه أسماء جلسات ما قدر يصنفها - راجع CLASSIFICATION_RULES:');
    unknowns.forEach(function (s) { Logger.log('  - ' + s.name); });
  }
}

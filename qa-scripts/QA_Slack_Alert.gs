/**
 * ============================================================
 *  تنبيه سلاك التلقائي - سكور الجودة (QA Score Alert)
 * ============================================================
 *
 * الفكرة:
 * كل ما يوصل تقييم جديد (صف جديد) للشيت، السكريبت يشيك على
 * عمود "علامة القسم %" (Section Score %). لو كان أقل من الحد
 * المحدد (70% افتراضيًا)، يرسل تنبيه لقناة سلاك محددة يمنشن
 * فيها المسؤولين عن الكوالتي.
 *
 * ============================================================
 *  الإعداد المطلوب منك (مرة وحدة بس):
 * ============================================================
 *
 * 1) افتح الشيت اللي فيه بيانات السكور.
 * 2) من القائمة: Extensions > Apps Script
 * 3) امسح أي كود موجود، والصق هذا الكود كامل.
 * 4) عدّل القيم في قسم "الإعدادات" تحت مباشرة:
 *      - SLACK_WEBHOOK_URL: رابط الـ Incoming Webhook (شرح تحت)
 *      - SCORE_THRESHOLD: الحد الأدنى المقبول (70 حاليًا)
 *      - SCORE_COLUMN_NAME: اسم عمود السكور بالضبط كما في الهيدر
 *      - AGENT_COLUMN_NAME: اسم عمود الموظف بالضبط كما في الهيدر
 *      - MENTIONS: قائمة اليوزرات المطلوب منشنتهم
 *
 * 5) لازم "تربط" السكريبت بحدث تلقائي (Trigger):
 *      من القائمة اليسار في Apps Script اضغط على أيقونة الساعة
 *      (Triggers) > Add Trigger، واختر:
 *        - Function to run: onFormSubmit
 *        - Event source: From spreadsheet
 *        - Event type: On form submit
 *      (لأن هذا شيت مربوط بفورم - كل تقييم جديد = form submission)
 *
 *      لو تحتاج كشف فوري حتى على التعديل اليدوي (مش بس عبر
 *      الفورم)، ضيف Trigger ثاني بنفس الطريقة لكن اختر:
 *        - Function to run: onEditTrigger
 *        - Event type: On edit
 *
 * 6) احصل على رابط Slack Incoming Webhook:
 *      a. روح https://api.slack.com/apps > Create New App > From scratch
 *      b. اختر اسم التطبيق (مثلاً "QA Score Alert") والـ Workspace
 *      c. من القائمة اليسار: Incoming Webhooks > فعّلها (Activate)
 *      d. اضغط "Add New Webhook to Workspace"
 *      e. اختر قناة #quality-and-training-team
 *      f. انسخ الرابط اللي يبدأ بـ https://hooks.slack.com/services/...
 *      g. حطه مكان SLACK_WEBHOOK_URL تحت
 *
 * 7) عشان يتم منشن سلمان وفدا فعليًا (مو بس كتابة اسمهم كنص)،
 *    محتاج توكن بوت إضافي (بصلاحية بسيطة تحول الإيميل لـ ID تلقائيًا):
 *      a. في نفس صفحة التطبيق اللي سويته (خطوة 6)، روح
 *         OAuth & Permissions من القائمة اليسار
 *      b. تحت "Scopes" > "Bot Token Scopes" اضغط "Add an OAuth Scope"
 *         وأضف: users:read.email
 *      c. اطلع لأعلى الصفحة واضغط "Install to Workspace" (أو
 *         "Reinstall to Workspace" لو مثبت من قبل) ووافق
 *      d. انسخ "Bot User OAuth Token" (يبدأ بـ xoxb-...)
 *      e. حطه مكان SLACK_BOT_TOKEN تحت
 *
 *    بعد كذا، السكريبت بيحول الإيميلات في MENTIONS تلقائيًا
 *    لمعرّفات سلاك الحقيقية عند كل إرسال - ما تحتاج تدور على
 *    الـ Member ID يدويًا.
 *
 * ============================================================
 */

// ============ الإعدادات - عدّل هذي القيم ============

const SLACK_WEBHOOK_URL = 'PASTE_YOUR_SLACK_WEBHOOK_URL_HERE'; // رابط الـ Incoming Webhook
const SLACK_BOT_TOKEN = 'PASTE_YOUR_BOT_TOKEN_HERE'; // يبدأ بـ xoxb- (خطوة 7 أعلاه)

const SCORE_THRESHOLD = 70; // الحد الأدنى المقبول بالنسبة المئوية

// اسم العمود بالضبط كما هو مكتوب في الهيدر (الصف الأول) في الشيت
const SCORE_COLUMN_NAME = 'علامة القسم % | Section Score %';
const AGENT_COLUMN_NAME = 'Agent - الموظف';
const AUDITOR_COLUMN_NAME = 'الموظف المتابع - Auditor';

// منشن الأشخاص المطلوب تنبيههم - بالإيميل، ويتحول تلقائيًا لمنشن حقيقي
const MENTIONS = [
  { name: 'Salman Alalshikh', email: 's.alalshikh@rewaa.com' },
  { name: 'Fidaa Yousef', email: 'f.Yousef@rewaa.com' },
];

// ============ نهاية الإعدادات ============


/**
 * يشتغل تلقائيًا عند استلام رد فورم جديد (تقييم جديد)
 */
function onFormSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    processRow(sheet, row);
  } catch (err) {
    Logger.log('onFormSubmit error: ' + err);
  }
}

/**
 * يشتغل عند أي تعديل يدوي في الشيت (اختياري - لو تبي كشف فوري
 * حتى للتعديل اليدوي مو بس الفورم)
 */
function onEditTrigger(e) {
  try {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    if (row === 1) return; // تجاهل تعديل الهيدر
    processRow(sheet, row);
  } catch (err) {
    Logger.log('onEditTrigger error: ' + err);
  }
}

/**
 * المنطق الأساسي: يقرأ الصف، يتحقق من السكور، ويرسل تنبيه لو لازم
 */
function processRow(sheet, rowNumber) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const scoreColIndex = headers.indexOf(SCORE_COLUMN_NAME);
  const agentColIndex = headers.indexOf(AGENT_COLUMN_NAME);
  const auditorColIndex = headers.indexOf(AUDITOR_COLUMN_NAME);

  if (scoreColIndex === -1 || agentColIndex === -1) {
    Logger.log('لم يتم العثور على أعمدة السكور أو الموظف - تحقق من أسماء الأعمدة في الإعدادات');
    return;
  }

  const rowValues = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  const rawScore = rowValues[scoreColIndex];
  const agentName = rowValues[agentColIndex];
  const auditorName = auditorColIndex !== -1 ? rowValues[auditorColIndex] : '';

  const scorePercent = parsePercent(rawScore);

  if (scorePercent === null) {
    Logger.log('تعذر قراءة قيمة السكور: ' + rawScore);
    return;
  }

  if (scorePercent < SCORE_THRESHOLD) {
    sendSlackAlert({
      agentName: agentName,
      auditorName: auditorName,
      scorePercent: scorePercent,
      rowNumber: rowNumber,
      sheetUrl: sheet.getParent().getUrl() + '#gid=' + sheet.getSheetId(),
    });
  }
}

/**
 * يحول القيمة (سواء رقم أو نص فيه %) إلى رقم نسبة مئوية
 */
function parsePercent(value) {
  if (value === null || value === undefined || value === '') return null;

  if (typeof value === 'number') {
    // لو كانت مكتوبة كنسبة عشرية (0.85) بدل 85
    return value <= 1 ? value * 100 : value;
  }

  const str = String(value).replace('%', '').trim();
  const num = parseFloat(str);
  return isNaN(num) ? null : num;
}

/**
 * يحول إيميل إلى Slack Member ID عن طريق Slack Web API
 * (users.lookupByEmail) - يحتاج SLACK_BOT_TOKEN بصلاحية
 * users:read.email. يخزن النتيجة مؤقتًا (Cache) لتقليل الطلبات.
 */
function lookupSlackIdByEmail(email) {
  if (!SLACK_BOT_TOKEN || SLACK_BOT_TOKEN === 'PASTE_YOUR_BOT_TOKEN_HERE') {
    return null;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'slackid_' + email;
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const url = 'https://slack.com/api/users.lookupByEmail?email=' + encodeURIComponent(email);
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN },
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.ok && json.user && json.user.id) {
      cache.put(cacheKey, json.user.id, 21600); // يخزن 6 ساعات
      return json.user.id;
    }
    Logger.log('لم يتم العثور على مستخدم سلاك بالإيميل: ' + email + ' | رد سلاك: ' + response.getContentText());
    return null;
  } catch (err) {
    Logger.log('خطأ أثناء البحث عن الإيميل ' + email + ': ' + err);
    return null;
  }
}

/**
 * يبني ويرسل رسالة سلاك بتنسيق منظم (شبيه بتنبيهات JIRA)
 */
function sendSlackAlert(data) {
  const mentionsText = MENTIONS.map(function (m) {
    const slackId = lookupSlackIdByEmail(m.email);
    return slackId ? '<@' + slackId + '>' : m.name;
  }).join(' ');

  const message = {
    blocks: [
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text:
            ':rotating_light: *تنبيه - سكور جودة منخفض* :rotating_light:',
        },
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: '*الموظف:*\n' + data.agentName },
          { type: 'mrkdwn', text: '*السكور:*\n' + data.scorePercent + '%' },
          { type: 'mrkdwn', text: '*المدقق:*\n' + (data.auditorName || '—') },
          { type: 'mrkdwn', text: '*الحد الأدنى المطلوب:*\n' + SCORE_THRESHOLD + '%' },
        ],
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: '<' + data.sheetUrl + '|فتح الشيت> | ' + mentionsText,
        },
      },
    ],
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(message),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
  Logger.log('Slack response: ' + response.getContentText());
}

/**
 * دالة اختبار يدوية - شغّلها من Apps Script لتجربة الإرسال
 * بدون الحاجة تنتظر تقييم فعلي أو تعدل شي في الشيت
 */
function testSlackAlert() {
  sendSlackAlert({
    agentName: 'اسم تجريبي',
    auditorName: 'Fidaa Yousef',
    scorePercent: 55,
    rowNumber: 999,
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  });
}

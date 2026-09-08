/**
 * ============================================================
 *  إضافة: تنبيه سلاك عند سكور منخفض
 * ============================================================
 *
 *  طريقة التركيب:
 *  1) الصق كل الكود اللي تحت في آخر السكريبت الموجود عندك
 *     (بدون ما تحذف أو تعدل أي شي من كودك الحالي).
 *
 *  2) روح لدالة onFormSubmitHandler الموجودة عندك حاليًا،
 *     وعدّلها من:
 *
 *        function onFormSubmitHandler(e) {
 *          refreshAll();
 *        }
 *
 *     إلى:
 *
 *        function onFormSubmitHandler(e) {
 *          refreshAll();
 *          checkScoreAndAlert(e.range.getRow());
 *        }
 *
 *     (سطر واحد بس تضيفه - كل شي ثاني في كودك يفضل زي ما هو)
 *
 *  3) عدّل الإعدادات تحت مباشرة (الويبهوك، التوكن، الإيميلات).
 *
 *  4) ما تحتاج تسوي Trigger جديد - الـ Trigger الموجود عندك
 *     أصلاً (onFormSubmitHandler) بيشغل كل شي تلقائيًا.
 * ============================================================
 */

// ============ إعدادات تنبيه سلاك ============

const SLACK_WEBHOOK_URL = 'PASTE_YOUR_SLACK_WEBHOOK_URL_HERE';
const SLACK_BOT_TOKEN = 'PASTE_YOUR_BOT_TOKEN_HERE'; // xoxb-... (خطوات الحصول عليه بالأسفل)

// اختر طريقة الإرسال:
// 'app'  -> الرسالة ترسل من التطبيق (فيها علامة APP صغيرة، إعداد أبسط)
// 'user' -> الرسالة ترسل فعليًا من حسابك بدون أي علامة (يحتاج تفويض إضافي)
const SEND_MODE = 'app';

// لو SEND_MODE = 'user'، لازم توكن مستخدم (xoxp-) بدل توكن البوت أعلاه
const SLACK_USER_TOKEN = 'PASTE_YOUR_USER_TOKEN_HERE'; // xoxp-...
const SLACK_CHANNEL_ID = 'PASTE_CHANNEL_ID_HERE'; // مطلوب فقط مع SEND_MODE = 'user'، مثال: C0123ABCD

const ALERT_SCORE_THRESHOLD = 0.70; // نفس صيغة sectionScore عندك (كسر عشري 0-1)

const ALERT_MENTIONS = [
  { name: 'Salman Alalshikh', email: 's.alalshikh@rewaa.com' },
  { name: 'Fidaa Yousef', email: 'f.Yousef@rewaa.com' },
];

// ============ نهاية الإعدادات ============


/**
 * يشيك على السكور اللي تم حسابه لصف معين (بعد ما refreshAll
 * يكون حسبه وكتبه في العمود 3)، ويرسل تنبيه لو تحت الحد
 */
function checkScoreAndAlert(rowNumber) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    if (rowNumber < 2) return;

    var rowValues = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sectionScore = rowValues[2]; // العمود C - علامة القسم % (كسر عشري)
    var auditorName = rowValues[4];  // العمود E - الموظف المتابع
    var agentName = rowValues[AGENT_COL]; // العمود F - Agent
    var type = String(rowValues[TYPE_COL] || '').trim();

    if (sectionScore === '' || sectionScore === null || typeof sectionScore !== 'number') return;

    if (sectionScore < ALERT_SCORE_THRESHOLD) {
      sendSlackAlert({
        agentName: agentName,
        auditorName: auditorName,
        scorePercent: Math.round(sectionScore * 100),
        weakPoints: findWeakPoints(sheet, rowValues, type),
        sheetUrl: ss.getUrl() + '#gid=' + sheet.getSheetId(),
      });
    }
  } catch (err) {
    Logger.log('checkScoreAndAlert error: ' + err);
  }
}

/**
 * يرجع أسماء المعايير اللي جاوبها المدقق "No" في هذا التقييم
 * بالذات (نقاط الضعف)، بالإضافة للأخطاء الحرجة إن وجدت
 */
function findWeakPoints(sheet, rowValues, type) {
  var weak = [];
  var range = SECTION_RANGES[type];

  if (range) {
    var labels = sheet.getRange(1, FIRST_CRIT_COL + range.start, 1, range.len).getValues()[0];
    for (var i = 0; i < range.len; i++) {
      var answer = rowValues[FIRST_CRIT_COL - 1 + range.start + i];
      if (answer === 'No') weak.push(labels[i]);
    }
  }

  var criticalLabels = sheet.getRange(1, FIRST_CRIT_COL + CRITICAL_OFFSET, 1, CRITICAL_LEN).getValues()[0];
  for (var c = 0; c < CRITICAL_LEN; c++) {
    var critAnswer = rowValues[FIRST_CRIT_COL - 1 + CRITICAL_OFFSET + c];
    if (critAnswer === 'Yes') weak.push('⚠️ خطأ حرج: ' + criticalLabels[c]);
  }

  return weak;
}

/**
 * يحول إيميل إلى Slack Member ID (يحتاج SLACK_BOT_TOKEN أو
 * SLACK_USER_TOKEN بصلاحية users:read.email)
 */
function lookupSlackIdByEmail(email) {
  var token = SEND_MODE === 'user' ? SLACK_USER_TOKEN : SLACK_BOT_TOKEN;
  if (!token || token.indexOf('PASTE_') === 0) return null;

  var cache = CacheService.getScriptCache();
  var cacheKey = 'slackid_' + email;
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var url = 'https://slack.com/api/users.lookupByEmail?email=' + encodeURIComponent(email);
  var options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    if (json.ok && json.user && json.user.id) {
      cache.put(cacheKey, json.user.id, 21600);
      return json.user.id;
    }
    Logger.log('Slack lookup failed for ' + email + ': ' + response.getContentText());
    return null;
  } catch (err) {
    Logger.log('Slack lookup error for ' + email + ': ' + err);
    return null;
  }
}

/**
 * يبني رسالة سلاك ويرسلها إما عن طريق Webhook (SEND_MODE='app')
 * أو عن طريق chat.postMessage بتوكن المستخدم (SEND_MODE='user')
 */
function sendSlackAlert(data) {
  var mentionsText = ALERT_MENTIONS.map(function (m) {
    var slackId = lookupSlackIdByEmail(m.email);
    return slackId ? '<@' + slackId + '>' : m.name;
  }).join(' ');

  var blocks = [
    {
      type: 'section',
      text: { type: 'mrkdwn', text: ':rotating_light: *تنبيه - سكور جودة منخفض* :rotating_light:' },
    },
    {
      type: 'section',
      fields: [
        { type: 'mrkdwn', text: '*الموظف:*\n' + data.agentName },
        { type: 'mrkdwn', text: '*السكور:*\n' + data.scorePercent + '%' },
        { type: 'mrkdwn', text: '*المدقق:*\n' + (data.auditorName || '—') },
        { type: 'mrkdwn', text: '*الحد الأدنى المطلوب:*\n' + Math.round(ALERT_SCORE_THRESHOLD * 100) + '%' },
      ],
    },
  ];

  if (data.weakPoints && data.weakPoints.length) {
    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: '*نقاط الضعف في هذا التقييم:*\n' + data.weakPoints.map(function (w) { return '• ' + w; }).join('\n'),
      },
    });
  }

  blocks.push({
    type: 'section',
    text: { type: 'mrkdwn', text: '<' + data.sheetUrl + '|فتح الشيت> | ' + mentionsText },
  });

  if (SEND_MODE === 'user') {
    sendViaUserToken(blocks);
  } else {
    sendViaWebhook(blocks);
  }
}

function sendViaWebhook(blocks) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ blocks: blocks }),
    muteHttpExceptions: true,
  };
  var response = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
  Logger.log('Slack webhook response: ' + response.getContentText());
}

function sendViaUserToken(blocks) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + SLACK_USER_TOKEN },
    payload: JSON.stringify({
      channel: SLACK_CHANNEL_ID,
      blocks: blocks,
      as_user: true,
    }),
    muteHttpExceptions: true,
  };
  var response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', options);
  Logger.log('Slack chat.postMessage response: ' + response.getContentText());
}

/**
 * دالة اختبار يدوية - شغّلها من Apps Script لتجربة الإرسال
 */
function testSlackAlert() {
  sendSlackAlert({
    agentName: 'اسم تجريبي',
    auditorName: 'Fidaa Yousef',
    scorePercent: 55,
    weakPoints: ['الإنصات الفعال وفهم العميل', 'دفع العميل نحو التفعيل', '⚠️ خطأ حرج: سلوك غير احترافي'],
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  });
}

/**
 * ============================================================
 *  خطوات الحصول على الـ Webhook والتوكنات (مرة وحدة بس):
 * ============================================================
 *
 * أ) Slack Incoming Webhook (لـ SEND_MODE = 'app'):
 *    1. https://api.slack.com/apps > Create New App > From scratch
 *    2. اختر اسم واختر الـ Workspace
 *    3. Incoming Webhooks > فعّلها > Add New Webhook to Workspace
 *    4. اختر قناة quality-and-training-team
 *    5. انسخ الرابط وحطه في SLACK_WEBHOOK_URL
 *
 * ب) Bot Token (لتحويل الإيميل لمنشن حقيقي مع SEND_MODE = 'app'):
 *    1. نفس التطبيق > OAuth & Permissions > Bot Token Scopes
 *    2. أضف: users:read.email
 *    3. Install to Workspace (أو Reinstall لو مثبت من قبل)
 *    4. انسخ Bot User OAuth Token (xoxb-...) وحطه في SLACK_BOT_TOKEN
 *
 * ج) User Token (لـ SEND_MODE = 'user' - الرسالة تطلع من حسابك):
 *    1. نفس التطبيق > OAuth & Permissions > User Token Scopes
 *    2. أضف: chat:write و users:read.email
 *    3. Install to Workspace > بيطلب منك (أنت شخصيًا) الموافقة
 *    4. انسخ User OAuth Token (xoxp-...) وحطه في SLACK_USER_TOKEN
 *    5. احصل على Channel ID لقناة quality-and-training-team:
 *       افتح القناة في سلاك من المتصفح > آخر الرابط بعد /archives/
 *       (شكله مثل C0123ABCD) وحطه في SLACK_CHANNEL_ID
 * ============================================================
 */

/**
 * ============================================================
 *  مطابقة أرقام TRSM بين صفحتين وتلوين النتيجة
 * ============================================================
 *
 * الفكرة:
 * - الصفحة الأولى ("Geidea Training 2"): فيها عمود TRSM (آخر عمود)
 * - الصفحة الثانية ("Nearpay TRSM" أو أي اسم آخر): فيها قائمة
 *   أكواد NearPay (سواء بعمود واحد أو صف واحد ممتد)
 *
 * السكريبت يقرأ كل قيم الصفحة الثانية (بغض النظر عن شكلها:
 * عمود، صف، أو موزعة)، يبنيها كمجموعة مرجعية، ثم يمر على كل
 * صف بالصفحة الأولى:
 *   - لو TRSM موجود بالمجموعة المرجعية → يلوّن الصف بالكامل أخضر
 *   - لو مو موجود → يلوّن الصف بالكامل أحمر
 *
 * ============================================================
 *  طريقة التركيب والتشغيل:
 * ============================================================
 *
 * 1) افتح ملفك (Geidea Training) في المتصفح.
 * 2) من القائمة: Extensions > Apps Script
 * 3) الصق هذا الكود كامل.
 * 4) عدّل الإعدادات تحت لو أسماء التبويبات مختلفة عندك.
 * 5) شغّل دالة matchAndColorTRSM من محرر Apps Script.
 * 6) أول مرة بيطلب صلاحيات (Authorize) - وافق.
 * 7) ارجع لملفك بالمتصفح - الصفوف بتكون ملوّنة تلقائيًا.
 *
 * ============================================================
 */

// ============ الإعدادات - عدّلها لو أسماء التبويبات مختلفة ============

const SOURCE_SHEET_NAME = 'Geidea Training 2';   // الصفحة اللي فيها بياناتك (تتلون)
// ملاحظة: مو محتاج تحدد اسم تبويب NearPay - السكريبت يختاره تلقائيًا
// (أول تبويب غير SOURCE_SHEET_NAME بالملف). لو عندك أكثر من تبويبين
// وتبي تحدده يدويًا، فعّل السطر التالي واكتب اسمه الصحيح:
// const REFERENCE_SHEET_NAME_OVERRIDE = 'اسم_التبويب_هنا';

const TRSM_COLUMN_IN_SOURCE = 6; // رقم العمود اللي فيه TRSM بالصفحة الأولى (F = 6)
const HEADER_ROW = 1;            // رقم صف العناوين بالصفحة الأولى (يُستثنى من التلوين)

const COLOR_MATCH = '#d9ead3';    // أخضر - موجود بـ NearPay
const COLOR_NO_MATCH = '#f4cccc'; // أحمر - غير موجود بـ NearPay

// ============ نهاية الإعدادات ============


/**
 * يختار صفحة المرجع تلقائيًا: أول تبويب غير SOURCE_SHEET_NAME.
 * لو عرّفت REFERENCE_SHEET_NAME_OVERRIDE فوق، يستخدمه بدل التخمين.
 */
function getReferenceSheet(ss, sourceSheet) {
  if (typeof REFERENCE_SHEET_NAME_OVERRIDE !== 'undefined') {
    const named = ss.getSheetByName(REFERENCE_SHEET_NAME_OVERRIDE);
    if (named) return named;
  }
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    if (all[i].getSheetId() !== sourceSheet.getSheetId()) return all[i];
  }
  return null;
}

function matchAndColorTRSM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    throw new Error('ما لقيت صفحة باسم "' + SOURCE_SHEET_NAME + '" - تأكد من SOURCE_SHEET_NAME بالإعدادات');
  }
  const referenceSheet = getReferenceSheet(ss, sourceSheet);
  if (!referenceSheet) {
    throw new Error('ما لقيت صفحة مرجعية ثانية بالملف. حدد REFERENCE_SHEET_NAME_OVERRIDE يدويًا بالإعدادات.');
  }

  // 1) نبني مجموعة مرجعية من كل قيم الصفحة الثانية (بغض النظر عن شكلها)
  const referenceValues = referenceSheet.getDataRange().getValues();
  const referenceSet = new Set();

  referenceValues.forEach(function (row) {
    row.forEach(function (cell) {
      const val = String(cell || '').trim().toUpperCase();
      if (val && val !== 'NEARPAY' && val !== 'TRSM') {
        // لو الخلية فيها أكثر من كود مفصول بمسافات، نقسمها
        val.split(/\s+/).forEach(function (code) {
          if (code) referenceSet.add(code);
        });
      }
    });
  });

  Logger.log('عدد أكواد NearPay المرجعية: ' + referenceSet.size);

  // 2) نمر على كل صف بالصفحة الأولى ونلوّن
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();

  let matchCount = 0;
  let noMatchCount = 0;
  let emptyCount = 0;

  for (let r = HEADER_ROW + 1; r <= lastRow; r++) {
    const trsmValue = String(sourceSheet.getRange(r, TRSM_COLUMN_IN_SOURCE).getValue() || '').trim().toUpperCase();
    const rowRange = sourceSheet.getRange(r, 1, 1, lastCol);

    if (!trsmValue) {
      emptyCount++;
      continue; // نتجاهل الصفوف اللي ما فيها TRSM أصلاً (نتركها بدون تلوين)
    }

    if (referenceSet.has(trsmValue)) {
      rowRange.setBackground(COLOR_MATCH);
      matchCount++;
    } else {
      rowRange.setBackground(COLOR_NO_MATCH);
      noMatchCount++;
    }
  }

  Logger.log('تم التلوين - مطابق (أخضر): ' + matchCount + ' | غير مطابق (أحمر): ' + noMatchCount + ' | بدون TRSM: ' + emptyCount);

  SpreadsheetApp.getUi().alert(
    'تمت المطابقة بنجاح ✅\n\n' +
    'أخضر (موجود بـ NearPay): ' + matchCount + '\n' +
    'أحمر (غير موجود): ' + noMatchCount + '\n' +
    'بدون رقم TRSM: ' + emptyCount
  );
}

/**
 * دالة اختبار - تعرض بالـ Log تفاصيل المطابقة بدون ما تلوّن فعليًا
 * شغّلها أول مرة للتأكد إن كل شي مضبوط
 */
function testDryRun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    Logger.log('تأكد من SOURCE_SHEET_NAME بالإعدادات. الصفحات الموجودة فعليًا:');
    ss.getSheets().forEach(function (s) { Logger.log(' - ' + s.getName()); });
    return;
  }
  const referenceSheet = getReferenceSheet(ss, sourceSheet);
  if (!referenceSheet) {
    Logger.log('ما لقيت صفحة مرجعية ثانية. الصفحات الموجودة فعليًا:');
    ss.getSheets().forEach(function (s) { Logger.log(' - ' + s.getName()); });
    return;
  }
  Logger.log('صفحة المرجع المكتشفة تلقائيًا: ' + referenceSheet.getName());

  const referenceValues = referenceSheet.getDataRange().getValues();
  const referenceSet = new Set();
  referenceValues.forEach(function (row) {
    row.forEach(function (cell) {
      const val = String(cell || '').trim().toUpperCase();
      if (val && val !== 'NEARPAY' && val !== 'TRSM') {
        val.split(/\s+/).forEach(function (code) { if (code) referenceSet.add(code); });
      }
    });
  });

  Logger.log('عدد أكواد NearPay: ' + referenceSet.size);
  Logger.log('أول 5 أكواد كعينة: ' + Array.from(referenceSet).slice(0, 5).join(', '));

  const lastRow = Math.min(sourceSheet.getLastRow(), HEADER_ROW + 10); // أول 10 صفوف فقط للاختبار
  for (let r = HEADER_ROW + 1; r <= lastRow; r++) {
    const trsmValue = String(sourceSheet.getRange(r, TRSM_COLUMN_IN_SOURCE).getValue() || '').trim().toUpperCase();
    const found = referenceSet.has(trsmValue);
    Logger.log('صف ' + r + ' | TRSM: ' + trsmValue + ' | ' + (found ? '✅ موجود' : '❌ غير موجود'));
  }
}

/**
 * webapp.js
 * ------------------------------------------------------------
 * Web アプリ（HtmlService）の入口。doGet でフォーム画面を返し、
 * google.script.run 経由で saveDay / loadDay を呼び出せるようにする。
 * ここでは SpreadsheetApp / HtmlService の呼び出しと、
 * logic.js・sheets.js の橋渡しのみを行う。
 * ------------------------------------------------------------
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('フォーム');
  template.defsJson = JSON.stringify(buildClientDefs_());
  template.todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  return template.evaluate()
    .setTitle('キャリア実績 日次入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// フォーム画面(クライアント側JS)に渡す項目定義。
// 項目を追加したときはlogic.jsの定義配列を直せば、ここを経由して自動的にフォームへ反映される。
function buildClientDefs_() {
  return {
    CARRIERS: Logic.CARRIERS,
    CARRIER_GROUP: Logic.CARRIER_GROUP,
    GROUPS: Logic.GROUPS,
    GROUP_LABEL: Logic.GROUP_LABEL,
    GROUP_COLOR: Logic.GROUP_COLOR,
    MIGRATIONS: Logic.MIGRATIONS,
    SOURCES: Logic.SOURCES,
    SUB_DEF: Logic.SUB_DEF,
    YAGOU_PRODUCTS: Logic.YAGOU_PRODUCTS,
    YAGOU_GROUP_ORDER: Logic.YAGOU_GROUP_ORDER,
    YAGOU_GROUP_LABEL: Logic.YAGOU_GROUP_LABEL
  };
}

// ============================================================
// クライアントから呼ばれる関数（google.script.run）
// ============================================================

// 保存。日付をキーに上書き／追記し、その日を含む週のシェア率を返す。
function saveDay(raw) {
  if (!raw || !Logic.toDateStr(raw.date)) {
    throw new Error('日付が指定されていません。');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var savedDate = saveDayToLog_(ss, raw);
  regenerateWeeklyAndMonthly_(ss);

  var monday = Logic.getMonday(savedDate);
  var sunday = Logic.getWeekEnd(monday);
  var weekModels = loadDayModelsInRange_(ss, monday, sunday);
  var weekAgg = Logic.aggregateDayModels(weekModels);

  return {
    ok: true,
    date: savedDate,
    weekStart: monday,
    weekEnd: sunday,
    weekShare: {
      mb: weekAgg.computed.mbShare,
      rikuKomi: weekAgg.computed.rikuKomiShare,
      rikuNuki: weekAgg.computed.rikuNukiShare
    }
  };
}

// 指定日の登録済みデータを読み込む（未登録なら null）
function loadDay(dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return loadDayFromLog_(ss, dateStr);
}

// 「集計シートを開く」用のリンク集
function getSheetLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    summary: getSheetUrl_(ss, SHEET_NAMES.SUMMARY),
    weekly: getSheetUrl_(ss, SHEET_NAMES.WEEKLY),
    monthly: getSheetUrl_(ss, SHEET_NAMES.MONTHLY),
    settings: getSheetUrl_(ss, SHEET_NAMES.SETTINGS)
  };
}

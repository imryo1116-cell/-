/**
 * report.js
 * ------------------------------------------------------------
 * 報告テキストの取得（google.script.runの入口）、CSV書き出し、メール送信。
 * 集計そのものは logic.js / sheets.js に任せ、ここでは
 * 「どの期間を取り出すか」「どう出力するか」だけを担当する。
 * ------------------------------------------------------------
 */

// 指定日を含む週の報告テキストを返す（「今週の報告テキストをコピー」用）
function getWeeklyReportText(dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var monday = Logic.getMonday(Logic.toDateStr(dateStr));
  var sunday = Logic.getWeekEnd(monday);
  var models = loadDayModelsInRange_(ss, monday, sunday);
  var agg = Logic.aggregateDayModels(models);

  var d = Logic.parseDateStr(monday);
  var weekIndex = Logic.getWeekIndexInMonth(monday);
  var title = Logic.buildTitle(d.getFullYear(), d.getMonth() + 1, weekIndex, monday, sunday, models.length);
  var staffing = getStaffingText_(ss, monday.slice(0, 7));

  return Logic.generateReportText(agg, { title: title, staffing: staffing });
}

// 指定月(yyyy-MM)の報告テキストを返す（「今月の報告テキストをコピー」用）
function getMonthlyReportText(yyyyMM) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = Number(yyyyMM.slice(0, 4));
  var month = Number(yyyyMM.slice(5, 7));
  var lastDay = new Date(year, month, 0).getDate();
  var start = Logic.formatDateStr(year, month, 1);
  var end = Logic.formatDateStr(year, month, lastDay);

  var models = loadDayModelsInRange_(ss, start, end);
  var agg = Logic.aggregateDayModels(models);
  var title = '【' + year + '年' + month + '月 月まとめ（' + models.length + '日）】';
  var staffing = getStaffingText_(ss, yyyyMM);

  return Logic.generateReportText(agg, { title: title, staffing: staffing });
}

// 指定月(yyyy-MM)の日次ログをCSV文字列として返す（BOM付きUTF-8）
function exportMonthCsv(yyyyMM) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = Number(yyyyMM.slice(0, 4));
  var month = Number(yyyyMM.slice(5, 7));
  var lastDay = new Date(year, month, 0).getDate();
  var start = Logic.formatDateStr(year, month, 1);
  var end = Logic.formatDateStr(year, month, lastDay);

  var sh = ss.getSheetByName(SHEET_NAMES.LOG);
  var defs = Logic.getLogHeaderDefs();
  var outHeader = defs.map(function (d) { return d.label; });
  var rows = [];

  if (sh) {
    var lastRow = sh.getLastRow();
    if (lastRow > LOG_HEADER_ROWS) {
      var sheetHeader = sh.getRange(LOG_HEADER_ROWS, 1, 1, sh.getLastColumn()).getValues()[0];
      var dateIdx = sheetHeader.indexOf('日付');
      var values = sh.getRange(LOG_HEADER_ROWS + 1, 1, lastRow - LOG_HEADER_ROWS, sh.getLastColumn()).getValues();
      values.forEach(function (row) {
        var dateStr = Logic.toDateStr(row[dateIdx]);
        if (dateStr >= start && dateStr <= end) {
          var dict = {};
          sheetHeader.forEach(function (label, i) { dict[label] = row[i]; });
          rows.push(outHeader.map(function (label) { return dict.hasOwnProperty(label) ? dict[label] : ''; }));
        }
      });
    }
  }

  var csv = Logic.buildCsv(outHeader, rows);
  var BOM = '\uFEFF';
  return { filename: yyyyMM + '_日次ログ.csv', content: BOM + csv };
}

// 「報告をメール送信」：設定シートの宛先へ今週の報告テキストを送る
function sendReportMail(dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recipients = getMailRecipients_(ss);
  if (recipients.length === 0) {
    return { ok: false, message: '宛先が未設定です。「設定」シートに送信先メールアドレスを入力してください。' };
  }
  var text = getWeeklyReportText(dateStr);
  var monday = Logic.getMonday(Logic.toDateStr(dateStr));
  MailApp.sendEmail({
    to: recipients.join(','),
    subject: '【実績報告】' + monday + ' の週',
    body: text
  });
  return { ok: true, message: recipients.join('、') + ' に送信しました。' };
}

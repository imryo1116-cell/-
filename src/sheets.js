/**
 * sheets.js
 * ------------------------------------------------------------
 * SpreadsheetApp を使う処理（シート生成・読み書き）のみを置くファイル。
 * 集計そのものは行わず、集計は必ず logic.js の Logic.* を呼び出す。
 * ------------------------------------------------------------
 */

var SHEET_NAMES = {
  LOG: '日次ログ',
  SUMMARY: '日次サマリー',
  WEEKLY: '週次集計',
  MONTHLY: '月次集計',
  SETTINGS: '設定'
};

var LOG_HEADER_ROWS = 2; // 1行目=大分類、2行目=見出し名
var LOG_FROZEN_COLS = 1; // 日付列(A列)のみ固定

var COLOR_NAVY = '#1F3A68';
var COLOR_TEAL = '#0E7C7B';
var COLOR_GRAY_BG = '#F2F2F2';
var COLOR_WHITE = '#FFFFFF';

// ============================================================
// 初期セットアップ（この関数をスクリプトエディタから実行する）
// ============================================================
function 初期セットアップ() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  setupLogSheet_(ss);
  setupSummarySheet_(ss);
  setupSettingsSheet_(ss);
  regenerateWeeklyAndMonthly_(ss);

  ss.toast('初期セットアップが完了しました。「日次ログ」以降のシートを確認してください。', '完了', 5);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日次入力ツール')
    .addItem('初期セットアップ（再実行）', '初期セットアップ')
    .addItem('週次・月次集計を再生成', 'menuRegenerate_')
    .addToUi();
}

function menuRegenerate_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  regenerateWeeklyAndMonthly_(ss);
  ss.toast('週次・月次集計を再生成しました。');
}

// ============================================================
// 共通ユーティリティ
// ============================================================
function getOrCreateSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// 新規シートは26列しかない（不具合(3)対策）。書き込み前に必ず呼ぶこと。
function ensureGridSize(sheet, rows, cols) {
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if (maxRows < rows) sheet.insertRowsAfter(maxRows, rows - maxRows);
  if (maxCols < cols) sheet.insertColumnsAfter(maxCols, cols - maxCols);
}

function columnToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// シートの2行目（見出し名）から Logic.getColMap を作る
function readColMap_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var header = sheet.getRange(LOG_HEADER_ROWS, 1, 1, lastCol).getValues()[0];
  return Logic.getColMap(header);
}

// ============================================================
// シート①「日次ログ」
// 不具合(1)(2)(3)(4)(5) すべての対策がここに集約されている。
// ============================================================
function setupLogSheet_(ss) {
  var sh = getOrCreateSheet_(ss, SHEET_NAMES.LOG);
  var defs = Logic.getLogHeaderDefs();
  var numCols = defs.length;

  // --- (5) 再実行時のデータ退避：見出し名(ラベル)をキーに既存データを退避する ---
  var savedRows = [];
  var lastRow = sh.getLastRow();
  if (lastRow > LOG_HEADER_ROWS && sh.getLastColumn() > 0) {
    var oldHeader = sh.getRange(LOG_HEADER_ROWS, 1, 1, sh.getLastColumn()).getValues()[0];
    var oldValues = sh.getRange(LOG_HEADER_ROWS + 1, 1, lastRow - LOG_HEADER_ROWS, sh.getLastColumn()).getValues();
    oldValues.forEach(function (row) {
      var dict = {};
      oldHeader.forEach(function (label, i) {
        if (label) dict[label] = row[i];
      });
      savedRows.push(dict);
    });
  }

  // --- (1) 再実行に備え、結合とフリーズを一旦解除してから作り直す ---
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  if (sh.getMaxRows() > 0 && sh.getMaxColumns() > 0) {
    try {
      sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    } catch (e) {
      // 結合がない場合は何もしない
    }
  }
  sh.clear();

  // --- (3) 必要な行列数を先に確保する ---
  var reservedDataRows = Math.max(savedRows.length, 1) + 50; // 余裕を持たせる
  ensureGridSize(sh, LOG_HEADER_ROWS + reservedDataRows, numCols);

  // --- ヘッダー2行を書き込む ---
  var headerRows = Logic.buildHeaderRows(defs);
  sh.getRange(1, 1, 1, numCols).setValues([headerRows.row1]);
  sh.getRange(2, 1, 1, numCols).setValues([headerRows.row2]);

  // --- (2) 日付列は書式なしテキストにしてから書き込む ---
  var dateColIndex = defs.indexOf(defs.filter(function (d) { return d.key === 'date'; })[0]) + 1;
  sh.getRange(1, dateColIndex, sh.getMaxRows(), 1).setNumberFormat('@');

  // --- 退避しておいたデータをラベルで引き直して書き戻す ---
  if (savedRows.length > 0) {
    var newLabels = headerRows.row2;
    var dataToWrite = savedRows.map(function (dict) {
      return newLabels.map(function (label) {
        return dict.hasOwnProperty(label) ? dict[label] : '';
      });
    });
    // 日付は必ず正規化して文字列で書き戻す（不具合(2)対策）
    dataToWrite.forEach(function (row) {
      row[dateColIndex - 1] = Logic.toDateStr(row[dateColIndex - 1]);
    });
    sh.getRange(LOG_HEADER_ROWS + 1, 1, dataToWrite.length, numCols).setValues(dataToWrite);
  }

  // --- 大分類セルの結合（固定列の境界をまたぐ結合がないかガードする） ---
  var ranges = Logic.buildMergeRanges(defs);
  Logic.validateMergeRanges(ranges, LOG_FROZEN_COLS);
  ranges.forEach(function (r) {
    sh.getRange(1, r.startCol, 1, r.span).merge();
  });

  // --- 見た目 ---
  sh.getRange(1, 1, 1, numCols).setFontWeight('bold').setFontColor(COLOR_WHITE).setBackground(COLOR_NAVY).setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, numCols).setFontWeight('bold').setBackground('#DCE3EC');
  defs.forEach(function (d, i) {
    if (d.major === '番号移行' || d.major === '移転元') {
      sh.getRange(1, i + 1, 1, 1).setBackground(COLOR_TEAL);
    } else if (Logic.GROUPS.indexOf(d.group) !== -1) {
      sh.getRange(1, i + 1, 1, 1).setBackground(Logic.GROUP_COLOR[d.group]);
    }
  });
  sh.setRowHeights(1, 1, 22);
  sh.setRowHeights(2, 1, 22);

  // --- (1) 結合をすべて終えてから最後にフリーズをかける ---
  sh.setFrozenRows(LOG_HEADER_ROWS);
  sh.setFrozenColumns(LOG_FROZEN_COLS);
}

// ============================================================
// 日次ログへの保存（日付をキーに上書き更新。なければ追記して日付順ソート）
// ============================================================
function saveDayToLog_(ss, raw) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = ss.getSheetByName(SHEET_NAMES.LOG);
    if (!sh) throw new Error('「日次ログ」シートが見つかりません。先に初期セットアップを実行してください。');

    var lastCol = sh.getLastColumn();
    var header = sh.getRange(LOG_HEADER_ROWS, 1, 1, lastCol).getValues()[0];
    var dateColIndex = header.indexOf('日付') + 1;
    if (dateColIndex === 0) throw new Error('「日付」列が見つかりません。');

    raw.updatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var targetDate = Logic.toDateStr(raw.date);

    var defs = Logic.getLogHeaderDefs();
    var defRow = Logic.rowFromRaw(raw, defs);
    var labelToValue = {};
    defs.forEach(function (d, i) { labelToValue[d.label] = defRow[i]; });
    var newRow = header.map(function (label) {
      return labelToValue.hasOwnProperty(label) ? labelToValue[label] : '';
    });

    var lastRow = sh.getLastRow();
    var foundRowNum = -1;
    if (lastRow > LOG_HEADER_ROWS) {
      var dateValues = sh.getRange(LOG_HEADER_ROWS + 1, dateColIndex, lastRow - LOG_HEADER_ROWS, 1).getValues();
      for (var i = 0; i < dateValues.length; i++) {
        if (Logic.toDateStr(dateValues[i][0]) === targetDate) {
          foundRowNum = LOG_HEADER_ROWS + 1 + i;
          break;
        }
      }
    }

    if (foundRowNum > 0) {
      sh.getRange(foundRowNum, 1, 1, header.length).setValues([newRow]);
    } else {
      ensureGridSize(sh, sh.getLastRow() + 1, header.length);
      sh.getRange(sh.getLastRow() + 1, 1, 1, header.length).setValues([newRow]);
    }

    // 日付昇順に並び替え（日付は 'yyyy-MM-dd' のテキストなので文字列比較で正しく並ぶ）
    lastRow = sh.getLastRow();
    if (lastRow > LOG_HEADER_ROWS) {
      sh.getRange(LOG_HEADER_ROWS + 1, 1, lastRow - LOG_HEADER_ROWS, header.length)
        .sort({ column: dateColIndex, ascending: true });
    }

    return targetDate;
  } finally {
    lock.releaseLock();
  }
}

// 指定日のデータを読み込む（未登録なら null）
function loadDayFromLog_(ss, dateStr) {
  var sh = ss.getSheetByName(SHEET_NAMES.LOG);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow <= LOG_HEADER_ROWS) return null;

  var header = sh.getRange(LOG_HEADER_ROWS, 1, 1, lastCol).getValues()[0];
  var dateColIndex = header.indexOf('日付') + 1;
  if (dateColIndex === 0) return null;

  var target = Logic.toDateStr(dateStr);
  var dateValues = sh.getRange(LOG_HEADER_ROWS + 1, dateColIndex, lastRow - LOG_HEADER_ROWS, 1).getValues();
  var foundIndex = -1;
  for (var i = 0; i < dateValues.length; i++) {
    if (Logic.toDateStr(dateValues[i][0]) === target) { foundIndex = i; break; }
  }
  if (foundIndex === -1) return null;

  var rowValues = sh.getRange(LOG_HEADER_ROWS + 1 + foundIndex, 1, 1, lastCol).getValues()[0];
  var defs = Logic.getLogHeaderDefs();
  var raw = {};
  header.forEach(function (label, i) {
    var def = defs.filter(function (d) { return d.label === label; })[0];
    if (def) raw[def.key] = def.type === 'date' ? Logic.toDateStr(rowValues[i]) : rowValues[i];
  });
  return raw;
}

// 指定期間([startDate, endDate] 両端含む 'yyyy-MM-dd') の日次モデル配列を取得する
function loadDayModelsInRange_(ss, startDate, endDate) {
  var sh = ss.getSheetByName(SHEET_NAMES.LOG);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow <= LOG_HEADER_ROWS) return [];

  var header = sh.getRange(LOG_HEADER_ROWS, 1, 1, lastCol).getValues()[0];
  var defs = Logic.getLogHeaderDefs();
  var values = sh.getRange(LOG_HEADER_ROWS + 1, 1, lastRow - LOG_HEADER_ROWS, lastCol).getValues();

  var models = [];
  values.forEach(function (row) {
    var raw = {};
    header.forEach(function (label, i) {
      var def = defs.filter(function (d) { return d.label === label; })[0];
      if (def) raw[def.key] = def.type === 'date' ? Logic.toDateStr(row[i]) : row[i];
    });
    var dateStr = Logic.toDateStr(raw.date);
    if (dateStr && dateStr >= startDate && dateStr <= endDate) {
      models.push(Logic.buildDayModel(raw));
    }
  });
  return models;
}

function loadAllDayModels_(ss) {
  return loadDayModelsInRange_(ss, '0000-01-01', '9999-12-31');
}

// ============================================================
// シート②「日次サマリー」（日次ログと1:1の行対応。数式で自動計算する）
// ============================================================
var SUMMARY_HEADER_ROW = 1;
var SUMMARY_RESERVED_ROWS = 1000; // 約2.7年分。不足したら初期セットアップを再実行する

function summaryColumnDefs_() {
  var cols = [
    { label: '日付' }, { label: '曜日' },
    { label: 'MB総販_DCM' }, { label: 'MB総販_SBYM' }, { label: 'MB総販_AUUQ' }, { label: 'MB総販_計' },
    { label: 'MB総販_DCM構成比' }, { label: 'MB総販_SBYM構成比' }, { label: 'MB総販_AUUQ構成比' },
    { label: 'MNPリク込み_DCM' }, { label: 'MNPリク込み_SBYM' }, { label: 'MNPリク込み_AUUQ' }, { label: 'MNPリク込み_計' },
    { label: 'MNPリク除き_DCM' }, { label: 'MNPリク除き_SBYM' }, { label: 'MNPリク除き_AUUQ' }, { label: 'MNPリク除き_計' },
    { label: '内SIMのみ_計' },
    { label: '機変_DCM' }, { label: '機変_SBYM' }, { label: '機変_AUUQ' },
    { label: '純新規_DCM' }, { label: '純新規_SBYM' }, { label: '純新規_AUUQ' },
    { label: '番号移行_計' }, { label: '副商材_計' }, { label: '屋号商材_計' }
  ];
  return cols;
}

function setupSummarySheet_(ss) {
  var logSh = ss.getSheetByName(SHEET_NAMES.LOG);
  if (!logSh) throw new Error('先に日次ログを作成してください。');
  var logColMap = readColMap_(logSh); // ラベル→列番号(1始まり)
  var L = function (label) {
    var col = logColMap[label];
    if (!col) throw new Error('日次ログに見出し「' + label + '」が見つかりません。');
    return columnToLetter_(col);
  };
  var carrierCol = function (carrier, field) { return L(carrier + '_' + field); };

  var sh = getOrCreateSheet_(ss, SHEET_NAMES.SUMMARY);
  sh.clear();
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);

  var cols = summaryColumnDefs_();
  ensureGridSize(sh, SUMMARY_RESERVED_ROWS + SUMMARY_HEADER_ROW, cols.length);
  sh.getRange(1, 1, 1, cols.length).setValues([cols.map(function (c) { return c.label; })])
    .setFontWeight('bold').setBackground(COLOR_NAVY).setFontColor(COLOR_WHITE);
  sh.getRange(1, 1, 1, cols.length).setNumberFormat('@');

  var groupCarriers = { DCM: ['DCM'], SBYM: ['SB', 'YM'], AUUQ: ['au', 'UQ'] };
  var logRowOffset = LOG_HEADER_ROWS + 1 - (SUMMARY_HEADER_ROW + 1); // summary行R → log行(R+offset)

  var fieldSum = function (group, field, logRow) {
    return groupCarriers[group].map(function (c) {
      return "日次ログ!" + carrierCol(c, field) + logRow;
    }).join('+');
  };
  var mbSum = function (group, logRow) {
    return groupCarriers[group].map(function (c) {
      return "日次ログ!" + carrierCol(c, '機変') + logRow + '+' + "日次ログ!" + carrierCol(c, '純新規') + logRow + '+' + "日次ログ!" + carrierCol(c, 'MNP') + logRow;
    }).join('+');
  };

  var rows = [];
  for (var r = SUMMARY_HEADER_ROW + 1; r <= SUMMARY_HEADER_ROW + SUMMARY_RESERVED_ROWS; r++) {
    var logRow = r + logRowOffset;
    var dateRef = "日次ログ!" + L('日付') + logRow;
    var weekdayRef = "日次ログ!" + L('曜日') + logRow;

    var mbDcm = mbSum('DCM', logRow), mbSbym = mbSum('SBYM', logRow), mbAuuq = mbSum('AUUQ', logRow);
    var mbTotalFormula = '(' + mbDcm + ')+(' + mbSbym + ')+(' + mbAuuq + ')';

    var mnpDcm = "日次ログ!" + carrierCol('DCM', 'MNP') + logRow;
    var mnpSbym = fieldSum('SBYM', 'MNP', logRow);
    var mnpAuuq = fieldSum('AUUQ', 'MNP', logRow);
    var mnpTotalFormula = '(' + mnpDcm + ')+(' + mnpSbym + ')+(' + mnpAuuq + ')';

    var migSbYm = "日次ログ!" + L('移行_SB→YM') + logRow + '+' + "日次ログ!" + L('移行_YM→SB') + logRow;
    var migAuUq = "日次ログ!" + L('移行_au→UQ') + logRow + '+' + "日次ログ!" + L('移行_UQ→au') + logRow;
    var nukiDcm = mnpDcm;
    var nukiSbym = '(' + mnpSbym + ')-(' + migSbYm + ')';
    var nukiAuuq = '(' + mnpAuuq + ')-(' + migAuUq + ')';
    var nukiTotalFormula = '(' + nukiDcm + ')+(' + nukiSbym + ')+(' + nukiAuuq + ')';

    var uchiSimTotal = Logic.CARRIERS.map(function (c) { return "日次ログ!" + carrierCol(c, '内SIMのみ') + logRow; }).join('+');

    var subTotalRefs = Logic.SUB_DEF.filter(function (d) { return d.type === 'input'; })
      .map(function (d) { return "日次ログ!" + L(d.label) + logRow; }).join('+');

    var yagouTotalRefs = [];
    Logic.YAGOU_PRODUCTS.forEach(function (p) {
      Logic.YAGOU_GROUP_ORDER.forEach(function (g) {
        yagouTotalRefs.push("日次ログ!" + L('屋号_' + p.label + '_' + Logic.YAGOU_GROUP_LABEL[g]) + logRow);
      });
    });

    var migTotalFormula = migSbYm + '+' + migAuUq;

    var ifBlank = function (formula) {
      return '=IF($A' + r + '="","",' + formula + ')';
    };

    rows.push([
      '=IF(' + dateRef + '="","",' + dateRef + ')',
      '=IF(' + weekdayRef + '="","",' + weekdayRef + ')',
      ifBlank(mbDcm), ifBlank(mbSbym), ifBlank(mbAuuq), ifBlank(mbTotalFormula),
      '=IF($F' + r + '="","",IF($F' + r + '=0,0,$C' + r + '/$F' + r + '))',
      '=IF($F' + r + '="","",IF($F' + r + '=0,0,$D' + r + '/$F' + r + '))',
      '=IF($F' + r + '="","",IF($F' + r + '=0,0,$E' + r + '/$F' + r + '))',
      ifBlank(mnpDcm), ifBlank(mnpSbym), ifBlank(mnpAuuq), ifBlank(mnpTotalFormula),
      ifBlank(nukiDcm), ifBlank(nukiSbym), ifBlank(nukiAuuq), ifBlank(nukiTotalFormula),
      ifBlank(uchiSimTotal),
      ifBlank(fieldSum('DCM', '機変', logRow)), ifBlank(fieldSum('SBYM', '機変', logRow)), ifBlank(fieldSum('AUUQ', '機変', logRow)),
      ifBlank(fieldSum('DCM', '純新規', logRow)), ifBlank(fieldSum('SBYM', '純新規', logRow)), ifBlank(fieldSum('AUUQ', '純新規', logRow)),
      ifBlank(migTotalFormula), ifBlank(subTotalRefs), ifBlank(yagouTotalRefs.join('+'))
    ]);
  }
  sh.getRange(SUMMARY_HEADER_ROW + 1, 1, rows.length, cols.length).setValues(rows);
  sh.getRange(SUMMARY_HEADER_ROW + 1, 7, rows.length, 3).setNumberFormat('0.0%');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
}

// ============================================================
// シート⑤「設定」
// ============================================================
var SETTINGS_STAFF_HEADER = ['適用開始年月', 'DCM稼働人数', 'SB/YM稼働人数', 'AU/UQ稼働人数'];
var SETTINGS_MAIL_HEADER = ['設定項目', '値'];
var SETTINGS_MAIL_ROW_LABEL = '報告メール宛先(カンマ区切り)';

function setupSettingsSheet_(ss) {
  var sh = getOrCreateSheet_(ss, SHEET_NAMES.SETTINGS);
  ensureGridSize(sh, 20, 8);

  // 稼働人数の履歴テーブル（既存データがあれば触らない＝データを消さない）
  if (sh.getRange(1, 1).getValue() !== SETTINGS_STAFF_HEADER[0]) {
    sh.getRange(1, 1, 1, 4).setValues([SETTINGS_STAFF_HEADER]).setFontWeight('bold').setBackground(COLOR_NAVY).setFontColor(COLOR_WHITE);
    var thisMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    sh.getRange(2, 1, 1, 4).setValues([[thisMonth, '', '', '']]);
  }

  // 報告メール宛先などの単発設定テーブル（F列以降）
  if (sh.getRange(1, 6).getValue() !== SETTINGS_MAIL_HEADER[0]) {
    sh.getRange(1, 6, 1, 2).setValues([SETTINGS_MAIL_HEADER]).setFontWeight('bold').setBackground(COLOR_NAVY).setFontColor(COLOR_WHITE);
    sh.getRange(2, 6, 1, 2).setValues([[SETTINGS_MAIL_ROW_LABEL, '']]);
  }
  sh.setColumnWidths(1, 4, 130);
  sh.setColumnWidths(6, 2, 220);
}

// 指定した年月(yyyy-MM)時点で有効な稼働人数の表示テキストを取得する
function getStaffingText_(ss, yyyymm) {
  var sh = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sh) return '';
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return '';
  var values = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  var best = null;
  values.forEach(function (row) {
    var applyFrom = row[0] ? String(row[0]) : '';
    if (applyFrom && applyFrom <= yyyymm) {
      if (!best || applyFrom > best[0]) best = row;
    }
  });
  if (!best) return '';
  var parts = [];
  if (best[1]) parts.push('DCM ' + best[1]);
  if (best[2]) parts.push('SB/YM ' + best[2]);
  if (best[3]) parts.push('AU/UQ ' + best[3]);
  return parts.join('／');
}

function getMailRecipients_(ss) {
  var sh = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var label = sh.getRange(r, 6).getValue();
    if (label === SETTINGS_MAIL_ROW_LABEL) {
      var raw = sh.getRange(r, 7).getValue();
      if (!raw) return [];
      return String(raw).split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
    }
  }
  return [];
}

function getSheetUrl_(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return ss.getUrl();
  return ss.getUrl() + '#gid=' + sh.getSheetId();
}

// ============================================================
// シート③④「週次集計」「月次集計」
// 数式ではなく、日次ログを読み込んで Logic で計算した結果を
// まとめて setValues する（保存の都度、自動で再生成される）。
// ============================================================
var BLOCK_COLS = ['項目', 'DCM', 'SB/YM', 'AU/UQ', '計', '構成比', '備考'];

// 1件分のブロック（見出し行＋データ行）を書き込み、次に書き込むべき行番号を返す
function writeAggBlock_(sh, startRow, agg, title) {
  ensureGridSize(sh, startRow + 60, BLOCK_COLS.length);

  sh.getRange(startRow, 1, 1, BLOCK_COLS.length).merge()
    .setValue(title).setBackground(COLOR_NAVY).setFontColor(COLOR_WHITE)
    .setFontWeight('bold').setHorizontalAlignment('left');

  var headerRow = startRow + 1;
  sh.getRange(headerRow, 1, 1, BLOCK_COLS.length).setValues([BLOCK_COLS]).setFontWeight('bold').setBackground('#DCE3EC');
  sh.getRange(headerRow, 2).setBackground(Logic.GROUP_COLOR.DCM).setFontColor(COLOR_WHITE);
  sh.getRange(headerRow, 3).setBackground(Logic.GROUP_COLOR.SBYM).setFontColor(COLOR_WHITE);
  sh.getRange(headerRow, 4).setBackground(Logic.GROUP_COLOR.AUUQ).setFontColor(COLOR_WHITE);

  var rows = Logic.buildAggBlockRows(agg);
  var dataStart = headerRow + 1;
  if (rows.length > 0) {
    var values = rows.map(function (r) { return [r.label, r.dcm, r.sbym, r.auuq, r.total, r.share, r.note]; });
    sh.getRange(dataStart, 1, values.length, BLOCK_COLS.length).setValues(values);
    sh.getRange(dataStart, 2, values.length, 3).setNumberFormat('0');
    rows.forEach(function (r, i) {
      if (r.header) {
        sh.getRange(dataStart + i, 1, 1, BLOCK_COLS.length).setFontWeight('bold').setBackground(COLOR_GRAY_BG);
      }
    });
  }
  return dataStart + rows.length + 1; // +1 = 空白行
}

function regenerateWeeklyAndMonthly_(ss) {
  var allModels = loadAllDayModels_(ss);
  writeWeeklySheet_(ss, allModels);
  writeMonthlySheet_(ss, allModels);
}

function clearSheetAndCharts_(sh) {
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.clear();
}

function groupModelsByWeek_(models) {
  var map = {};
  models.forEach(function (m) {
    var monday = Logic.getMonday(m.date);
    if (!map[monday]) map[monday] = [];
    map[monday].push(m);
  });
  return Object.keys(map).sort().map(function (monday) {
    return { weekStart: monday, models: map[monday] };
  });
}

function groupModelsByMonth_(models) {
  var map = {};
  models.forEach(function (m) {
    var ym = m.date.slice(0, 7); // 'yyyy-MM'
    if (!map[ym]) map[ym] = [];
    map[ym].push(m);
  });
  return Object.keys(map).sort().map(function (ym) {
    return { yearMonth: ym, models: map[ym] };
  });
}

function writeWeeklySheet_(ss, allModels) {
  var sh = getOrCreateSheet_(ss, SHEET_NAMES.WEEKLY);
  clearSheetAndCharts_(sh);
  sh.setColumnWidths(1, 1, 160);
  sh.setColumnWidths(2, 4, 90);
  sh.setColumnWidths(6, 1, 220);
  sh.setColumnWidths(7, 1, 300);

  var weeks = groupModelsByWeek_(allModels);
  var row = 1;
  weeks.forEach(function (w) {
    var agg = Logic.aggregateDayModels(w.models);
    var d = Logic.parseDateStr(w.weekStart);
    var weekIndex = Logic.getWeekIndexInMonth(w.weekStart);
    var title = Logic.buildTitle(d.getFullYear(), d.getMonth() + 1, weekIndex, w.weekStart, Logic.getWeekEnd(w.weekStart), w.models.length);
    row = writeAggBlock_(sh, row, agg, title);
  });
  if (weeks.length === 0) {
    sh.getRange(1, 1).setValue('日次ログにデータが入力されるとここに週次集計が表示されます。');
  }
}

function writeMonthlySheet_(ss, allModels) {
  var sh = getOrCreateSheet_(ss, SHEET_NAMES.MONTHLY);
  clearSheetAndCharts_(sh);
  sh.setColumnWidths(1, 1, 160);
  sh.setColumnWidths(2, 4, 90);
  sh.setColumnWidths(6, 1, 220);
  sh.setColumnWidths(7, 1, 300);

  var months = groupModelsByMonth_(allModels);
  var row = 1;
  months.forEach(function (mo) {
    var year = Number(mo.yearMonth.slice(0, 4));
    var month = Number(mo.yearMonth.slice(5, 7));
    var weeks = groupModelsByWeek_(mo.models);

    weeks.forEach(function (w) {
      var agg = Logic.aggregateDayModels(w.models);
      var weekIndex = Logic.getWeekIndexInMonth(w.weekStart);
      var title = Logic.buildTitle(year, month, weekIndex, w.weekStart, Logic.getWeekEnd(w.weekStart), w.models.length);
      row = writeAggBlock_(sh, row, agg, title);
    });

    var monthAgg = Logic.aggregateDayModels(mo.models);
    var monthTitle = '【' + year + '年' + month + '月 月合計（' + mo.models.length + '日）】';
    var monthTotalStartRow = row;
    row = writeAggBlock_(sh, row, monthAgg, monthTitle);

    insertMonthCharts_(sh, monthTotalStartRow, weeks, monthAgg, year, month);
    row += 20; // グラフ用の余白
  });

  if (months.length === 0) {
    sh.getRange(1, 1).setValue('日次ログにデータが入力されるとここに月次集計が表示されます。');
  }
}

// 3区分シェア率の週次推移グラフ（折れ線・MB総販とMNPで2枚）＋ 月合計構成比円グラフ（3枚）
function insertMonthCharts_(sh, anchorRow, weeks, monthAgg, year, month) {
  var dataCol = 10; // J列以降にグラフ用データテーブルを作る
  var startRow = anchorRow;

  // --- 週次推移（折れ線）用データ ---
  var trendHeader = ['週', 'DCM%', 'SB/YM%', 'AU/UQ%'];
  var mbTrendRows = [trendHeader];
  var mnpTrendRows = [trendHeader];
  weeks.forEach(function (w) {
    var agg = Logic.aggregateDayModels(w.models);
    var idx = Logic.getWeekIndexInMonth(w.weekStart);
    var label = idx + '週目';
    var mbShare = agg.computed.mbShare;
    var mnpShare = agg.computed.rikuKomiShare;
    mbTrendRows.push([label, mbShare.byGroup.DCM.pct / 100, mbShare.byGroup.SBYM.pct / 100, mbShare.byGroup.AUUQ.pct / 100]);
    mnpTrendRows.push([label, mnpShare.byGroup.DCM.pct / 100, mnpShare.byGroup.SBYM.pct / 100, mnpShare.byGroup.AUUQ.pct / 100]);
  });

  ensureGridSize(sh, startRow + mbTrendRows.length + mnpTrendRows.length + 20, dataCol + 3);

  sh.getRange(startRow, dataCol, mbTrendRows.length, 4).setValues(mbTrendRows);
  sh.getRange(startRow, dataCol, 1, 4).setFontWeight('bold');
  sh.getRange(startRow + 1, dataCol + 1, mbTrendRows.length - 1, 3).setNumberFormat('0.0%');

  var mnpTrendStartRow = startRow + mbTrendRows.length + 2;
  sh.getRange(mnpTrendStartRow, dataCol, mnpTrendRows.length, 4).setValues(mnpTrendRows);
  sh.getRange(mnpTrendStartRow, dataCol, 1, 4).setFontWeight('bold');
  sh.getRange(mnpTrendStartRow + 1, dataCol + 1, mnpTrendRows.length - 1, 3).setNumberFormat('0.0%');

  var groupColors = [Logic.GROUP_COLOR.DCM, Logic.GROUP_COLOR.SBYM, Logic.GROUP_COLOR.AUUQ];

  if (weeks.length > 0) {
    var mbLineChart = sh.newChart().asLineChart()
      .addRange(sh.getRange(startRow, dataCol, mbTrendRows.length, 4))
      .setOption('title', year + '年' + month + '月 MB総販 3区分シェア率推移')
      .setOption('colors', groupColors)
      .setPosition(anchorRow, dataCol + 6, 0, 0)
      .build();
    sh.insertChart(mbLineChart);

    var mnpLineChart = sh.newChart().asLineChart()
      .addRange(sh.getRange(mnpTrendStartRow, dataCol, mnpTrendRows.length, 4))
      .setOption('title', year + '年' + month + '月 MNP(リク込み) 3区分シェア率推移')
      .setOption('colors', groupColors)
      .setPosition(anchorRow + 18, dataCol + 6, 0, 0)
      .build();
    sh.insertChart(mnpLineChart);
  }

  // --- 月合計 構成比 円グラフ（MB総販／リク込み／リク除き） ---
  var pieStartRow = mnpTrendStartRow + mnpTrendRows.length + 2;
  var pieDefs = [
    { title: 'MB総販 構成比', byGroup: monthAgg.computed.mbShare.byGroup },
    { title: 'MNPリク込み 構成比', byGroup: monthAgg.computed.rikuKomiShare.byGroup },
    { title: 'MNPリク除き 構成比', byGroup: monthAgg.computed.rikuNukiShare.byGroup }
  ];
  pieDefs.forEach(function (pd, i) {
    var r = pieStartRow + i * 5;
    var tableRows = [['区分', '件数']];
    Logic.GROUPS.forEach(function (g) { tableRows.push([Logic.GROUP_LABEL[g], pd.byGroup[g].count]); });
    ensureGridSize(sh, r + tableRows.length, dataCol + 1);
    sh.getRange(r, dataCol, tableRows.length, 2).setValues(tableRows);
    sh.getRange(r, dataCol, 1, 2).setFontWeight('bold');

    var pieChart = sh.newChart().asPieChart()
      .addRange(sh.getRange(r, dataCol, tableRows.length, 2))
      .setOption('title', year + '年' + month + '月 ' + pd.title)
      .setOption('colors', groupColors)
      .setPosition(anchorRow + 36 + i * 16, dataCol + 6, 0, 0)
      .build();
    sh.insertChart(pieChart);
  });
}

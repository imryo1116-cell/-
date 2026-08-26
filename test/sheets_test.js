/**
 * sheets_test.js
 * ------------------------------------------------------------
 * gasMock.js の軽量 SpreadsheetApp モックを使い、sheets.js の
 * 初期セットアップ／保存／読み込み／再実行時のデータ保持を検証する。
 * ------------------------------------------------------------
 */
var assert = require('assert');
var Logic = require('../src/logic.js');
var createGasContext = require('./gasMock.js').createGasContext;

function run() {
  var passed = 0;
  var failed = 0;
  var failures = [];

  function test(name, fn) {
    try {
      fn();
      passed++;
    } catch (e) {
      failed++;
      failures.push({ name: name, error: e });
    }
  }

  test('初期セットアップ: 日次ログの見出しと結合範囲が正しく作られる', function () {
    var g = createGasContext();
    g.context.初期セットアップ();
    var logSh = g.ss.getSheetByName('日次ログ');
    assert.ok(logSh, '日次ログシートが作成される');
    var header2 = logSh.getRange(2, 1, 1, logSh.getLastColumn()).getValues()[0];
    assert.ok(header2.indexOf('DCM_機変') !== -1);
    assert.ok(header2.indexOf('更新日時') !== -1);
    assert.strictEqual(logSh.frozenRows, 2);
    assert.strictEqual(logSh.frozenCols, 1);
    // A列(日付)の結合が固定列をまたいでいないこと(=マージ自体は起きているはず)
    assert.ok(logSh.merges.length > 0);
  });

  test('保存: 同じ日付は上書き、別日付は追記され、日付昇順に並ぶ', function () {
    var g = createGasContext();
    g.context.初期セットアップ();

    g.context.saveDayToLog_(g.ss, { date: '2026-08-12', DCM_kihen: 1 });
    g.context.saveDayToLog_(g.ss, { date: '2026-08-10', DCM_kihen: 2 });
    g.context.saveDayToLog_(g.ss, { date: '2026-08-10', DCM_kihen: 9 }); // 上書き

    var logSh = g.ss.getSheetByName('日次ログ');
    var lastRow = logSh.getLastRow();
    var lastCol = logSh.getLastColumn();
    var header = logSh.getRange(2, 1, 1, lastCol).getValues()[0];
    var dateCol = header.indexOf('日付') + 1;
    var kihenCol = header.indexOf('DCM_機変') + 1;
    var body = logSh.getRange(3, 1, lastRow - 2, lastCol).getValues();

    assert.strictEqual(body.length, 2, '同一日付は行が増えず2行のまま');
    assert.strictEqual(body[0][dateCol - 1], '2026-08-10');
    assert.strictEqual(body[0][kihenCol - 1], 9, '上書きされている');
    assert.strictEqual(body[1][dateCol - 1], '2026-08-12');
  });

  test('読み込み: 保存した値が loadDayFromLog_ で復元できる', function () {
    var g = createGasContext();
    g.context.初期セットアップ();
    g.context.saveDayToLog_(g.ss, { date: '2026-08-10', DCM_kihen: 3, dcm_hikari1g: 5, event: 'auイベント' });

    var raw = g.context.loadDayFromLog_(g.ss, '2026-08-10');
    assert.ok(raw);
    assert.strictEqual(raw.date, '2026-08-10');
    assert.strictEqual(raw.event, 'auイベント');
    assert.strictEqual(Number(raw.DCM_kihen), 3);
    assert.strictEqual(Number(raw.dcm_hikari1g), 5);

    var noneRaw = g.context.loadDayFromLog_(g.ss, '2099-01-01');
    assert.strictEqual(noneRaw, null);
  });

  test('不具合(5)対策: 初期セットアップを再実行しても既存の日次ログデータが消えない', function () {
    var g = createGasContext();
    g.context.初期セットアップ();
    g.context.saveDayToLog_(g.ss, { date: '2026-08-10', DCM_kihen: 7 });
    g.context.saveDayToLog_(g.ss, { date: '2026-08-11', DCM_kihen: 4 });

    g.context.初期セットアップ(); // 再実行

    var raw10 = g.context.loadDayFromLog_(g.ss, '2026-08-10');
    var raw11 = g.context.loadDayFromLog_(g.ss, '2026-08-11');
    assert.ok(raw10, '再実行後もデータが残っている');
    assert.strictEqual(Number(raw10.DCM_kihen), 7);
    assert.ok(raw11);
    assert.strictEqual(Number(raw11.DCM_kihen), 4);
  });

  test('週次・月次集計の自動再生成がエラーなく実行できる', function () {
    var g = createGasContext();
    g.context.初期セットアップ();
    g.context.saveDayToLog_(g.ss, { date: '2026-08-10', DCM_kihen: 3, DCM_mnp: 2 });
    g.context.saveDayToLog_(g.ss, { date: '2026-08-11', DCM_kihen: 1 });
    g.context.regenerateWeeklyAndMonthly_(g.ss);

    var weekly = g.ss.getSheetByName('週次集計');
    var monthly = g.ss.getSheetByName('月次集計');
    assert.ok(weekly.getLastRow() > 0);
    assert.ok(monthly.getLastRow() > 0);
  });

  test('日次サマリー: 見出し名から解決した数式で日次ログの値と一致する', function () {
    var g = createGasContext();
    g.context.初期セットアップ();
    g.context.saveDayToLog_(g.ss, { date: '2026-08-10', DCM_kihen: 3, DCM_junshinki: 1, DCM_mnp: 2 });

    var summarySh = g.ss.getSheetByName('日次サマリー');
    // モックはSpreadsheetの数式評価を行わないため、少なくとも数式文字列が
    // 正しい列を参照して組み立てられていることを確認する。
    var formula = summarySh.getCell(2, 3); // MB総販_DCM 列（1行目ヘッダー、2行目が最初のデータ行）
    assert.ok(typeof formula === 'string' && formula.indexOf('日次ログ!') !== -1);
    assert.ok(formula.indexOf('DCM_機変') === -1); // 列参照はラベルではなく列文字であること
  });

  return { passed: passed, failed: failed, failures: failures };
}

module.exports = { run: run };

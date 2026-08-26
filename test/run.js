/**
 * test/run.js
 * ------------------------------------------------------------
 * logic.js のローカルテスト。SpreadsheetApp 等は使わないため
 * `node test/run.js` だけで実行できる。
 * ------------------------------------------------------------
 */
var assert = require('assert');
var Logic = require('../src/logic.js');

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

// ============================================================
// toDateStr: Date型／シリアル値／'2026/8/5'／'2026-08-05' が
// すべて '2026-08-05' になる
// ============================================================
test('toDateStr: Date object', function () {
  var d = new Date(2026, 7, 5); // month is 0-indexed -> August
  assert.strictEqual(Logic.toDateStr(d), '2026-08-05');
});

test('toDateStr: slash format single digit', function () {
  assert.strictEqual(Logic.toDateStr('2026/8/5'), '2026-08-05');
});

test('toDateStr: iso format', function () {
  assert.strictEqual(Logic.toDateStr('2026-08-05'), '2026-08-05');
});

test('toDateStr: spreadsheet serial value', function () {
  // 1899-12-30 が 0。2026-08-05 のシリアル値を計算して検証する。
  var epoch = Date.UTC(1899, 11, 30);
  var target = Date.UTC(2026, 7, 5);
  var serial = Math.round((target - epoch) / 86400000);
  assert.strictEqual(Logic.toDateStr(serial), '2026-08-05');
});

// ============================================================
// MB総販：機変＋純新規＋MNP と一致する
// ============================================================
test('MB総販 = 機変+純新規+MNP（キャリア別・グループ別・合計）', function () {
  var model = Logic.buildDayModel({
    date: '2026-08-10',
    DCM_kihen: 3, DCM_junshinki: 2, DCM_mnp: 5,
    SB_kihen: 1, SB_junshinki: 1, SB_mnp: 1,
    YM_kihen: 0, YM_junshinki: 0, YM_mnp: 0,
    au_kihen: 2, au_junshinki: 0, au_mnp: 3,
    UQ_kihen: 0, UQ_junshinki: 1, UQ_mnp: 0
  });
  var mb = model.computed.mb;
  assert.strictEqual(mb.byCarrier.DCM, 10);
  assert.strictEqual(mb.byCarrier.SB, 3);
  assert.strictEqual(mb.byGroup.SBYM, 3); // SB(3) + YM(0)
  assert.strictEqual(mb.byGroup.AUUQ, 6); // au(5) + UQ(1)
  assert.strictEqual(mb.total, 10 + 3 + 6);
});

// ============================================================
// リク除き：SB/YM = リク込み-(SB→YM+YM→SB)、AU/UQも同様、
// DCM はリク込みと必ず一致
// ============================================================
test('MNPリク除き：グループ単位で番号移行を控除する', function () {
  var model = Logic.buildDayModel({
    date: '2026-08-10',
    DCM_mnp: 10,
    SB_mnp: 6, YM_mnp: 4,
    au_mnp: 5, UQ_mnp: 3,
    migSbToYm: 1, migYmToSb: 2,
    migAuToUq: 1, migUqToAu: 0
  });
  var mnp = model.computed.mnp;
  assert.strictEqual(mnp.rikuKomiByGroup.DCM, 10);
  assert.strictEqual(mnp.rikuKomiByGroup.SBYM, 10);
  assert.strictEqual(mnp.rikuKomiByGroup.AUUQ, 8);

  assert.strictEqual(mnp.rikuNukiByGroup.DCM, mnp.rikuKomiByGroup.DCM); // DCMは常に一致
  assert.strictEqual(mnp.rikuNukiByGroup.SBYM, 10 - (1 + 2));
  assert.strictEqual(mnp.rikuNukiByGroup.AUUQ, 8 - (1 + 0));
});

test('MNPリク除き：DCMは番号移行入力欄自体が存在しない前提でも常にリク込み=リク除き', function () {
  var model = Logic.buildDayModel({ date: '2026-08-10', DCM_mnp: 7 });
  assert.strictEqual(model.computed.mnp.rikuKomiByGroup.DCM, 7);
  assert.strictEqual(model.computed.mnp.rikuNukiByGroup.DCM, 7);
});

// ============================================================
// 内SIMのみ：MNPリク込みを超える値を入れてもロジックが壊れない
// ============================================================
test('内SIMのみはMNPリク込みを超えても計算が破綻しない（内数チェックは行わない）', function () {
  var model = Logic.buildDayModel({
    date: '2026-08-10',
    DCM_mnp: 2, DCM_uchiSim: 999
  });
  assert.strictEqual(model.computed.mnp.uchiSimByGroup.DCM, 999);
  assert.strictEqual(model.computed.mnp.uchiSimTotal, 999);
  // 他の計算に影響しないこと
  assert.strictEqual(model.computed.mnp.rikuKomiByGroup.DCM, 2);
});

// ============================================================
// 週区切り：月曜起算。日曜は前週に入ることを確認する境界テストを含む
// ============================================================
test('週区切り：月曜起算で同じ週内は同一の月曜日になる', function () {
  // 2026-08-10 は月曜日
  assert.strictEqual(Logic.getWeekdayJa('2026-08-10'), '月');
  assert.strictEqual(Logic.getMonday('2026-08-10'), '2026-08-10');
  assert.strictEqual(Logic.getMonday('2026-08-13'), '2026-08-10'); // 木曜
  assert.strictEqual(Logic.getMonday('2026-08-16'), '2026-08-10'); // 日曜（週の最終日）
});

test('週区切り：日曜日は前の週（月曜起算）に属する境界テスト', function () {
  // 2026-08-16 は日曜日。この週の月曜は 2026-08-10。
  assert.strictEqual(Logic.getWeekdayJa('2026-08-16'), '日');
  assert.strictEqual(Logic.getMonday('2026-08-16'), '2026-08-10');
  // 翌日の月曜（2026-08-17）は次の週として扱われる
  assert.strictEqual(Logic.getMonday('2026-08-17'), '2026-08-17');
  assert.notStrictEqual(Logic.getMonday('2026-08-16'), Logic.getMonday('2026-08-17'));
});

test('週区切り：月をまたぐ週番号が連番で振られる', function () {
  var weeks = Logic.getWeeksOfMonth(2026, 8);
  for (var i = 0; i < weeks.length; i++) {
    assert.strictEqual(weeks[i].index, i + 1);
  }
  var idx = Logic.getWeekIndexInMonth('2026-08-16');
  assert.ok(idx >= 1);
});

// ============================================================
// 報告テキスト：0件の項目・カテゴリが行ごと省略される
// ============================================================
test('報告テキスト：値が0の副商材カテゴリ行は出力されない', function () {
  var model = Logic.buildDayModel({
    date: '2026-08-10',
    DCM_kihen: 1, DCM_junshinki: 0, DCM_mnp: 0,
    dcm_hikari1g: 5
    // dcm_denki系・dcm_card系はすべて未入力(=0)
  });
  var text = Logic.generateReportText(model, { title: 'テスト' });
  assert.ok(text.indexOf('d光 1Gbps') !== -1, '入力ありの項目は出力される');
  assert.ok(text.indexOf('エネルギー事業') === -1, '0件のカテゴリは省略される');
  assert.ok(text.indexOf('クレジット事業') === -1, '0件のカテゴリは省略される');
});

test('報告テキスト：屋号商材が全0のときは商材行が出ない', function () {
  var model = Logic.buildDayModel({ date: '2026-08-10' });
  var text = Logic.generateReportText(model, { title: 'テスト' });
  assert.ok(text.indexOf('■屋号商材（計0件）') !== -1);
  assert.ok(text.indexOf('つながる機器保険：') === -1);
});

test('報告テキスト：番号移行・移転元が0件なら行自体が出ない', function () {
  var model = Logic.buildDayModel({ date: '2026-08-10', DCM_mnp: 1 });
  var text = Logic.generateReportText(model, { title: 'テスト' });
  assert.ok(text.indexOf('番号移行：') === -1);
  assert.ok(text.indexOf('移転元：') === -1);
});

// ============================================================
// 全ゼロのとき：0除算にならず 0.0% になる
// ============================================================
test('全ゼロのとき、シェア率は0.0%になり例外も出ない', function () {
  var model = Logic.buildDayModel({ date: '2026-08-10' });
  var share = model.computed.mbShare;
  assert.strictEqual(share.total, 0);
  Logic.GROUPS.forEach(function (g) {
    assert.strictEqual(share.byGroup[g].pct, 0);
    assert.strictEqual(Logic.fmtPct(share.byGroup[g].pct), '0.0%');
  });
  var text = Logic.generateReportText(model, { title: 'テスト' });
  assert.ok(text.indexOf('NaN') === -1);
});

// ============================================================
// 列マップ：SUB_DEF に項目を1つ足しても既存の列参照が壊れない
// ============================================================
test('列マップ：見出し名から列位置を取得でき、順序を変えても既存ラベルの取得は壊れない', function () {
  var header1 = ['日付', '曜日', '副_d光1G', '副_HOME5G'];
  var map1 = Logic.getColMap(header1);
  assert.strictEqual(map1['副_HOME5G'], 4);

  // 途中に新しい項目を1つ挿入しても、既存ラベルは自分の名前で引けるので壊れない
  var header2 = ['日付', '曜日', '副_d光1G', '副_新項目', '副_HOME5G'];
  var map2 = Logic.getColMap(header2);
  assert.strictEqual(map2['副_d光1G'], 3);
  assert.strictEqual(map2['副_HOME5G'], 5); // 位置は変わるが名前で引けば正しく取得できる
});

test('SUB_DEF に項目を1つ追加しても既存キーの解決が壊れない', function () {
  var extendedDef = Logic.SUB_DEF.concat([
    { key: 'test_new_item', label: '副_テスト新項目', group: 'DCM', category: '(Wi-Fi)', name: 'テスト新項目', type: 'input' }
  ]);
  assert.strictEqual(extendedDef.length, Logic.SUB_DEF.length + 1);
  // 既存の SUB_DEF 自体は変更されない（immutableな追加であることの確認）
  assert.strictEqual(Logic.SUB_DEF.filter(function (d) { return d.key === 'test_new_item'; }).length, 0);
});

// ============================================================
// 大分類の結合範囲が固定列の境界をまたいでいないこと（不具合(1)再発防止）
// ============================================================
test('大分類の結合範囲は固定列(A列)の境界をまたがない', function () {
  var headerDefs = [
    { label: '日付', major: '日付' },
    { label: '曜日', major: '基本情報' },
    { label: 'イベント', major: '基本情報' },
    { label: 'メモ', major: '基本情報' },
    { label: 'DCM_機変', major: 'DCM' },
    { label: 'DCM_純新規', major: 'DCM' },
    { label: 'DCM_MNP', major: 'DCM' },
    { label: 'DCM_内SIMのみ', major: 'DCM' }
  ];
  var ranges = Logic.buildMergeRanges(headerDefs);
  // frozenCols = 1（A列のみ固定）
  assert.doesNotThrow(function () {
    Logic.validateMergeRanges(ranges, 1);
  });
});

test('大分類の結合範囲が固定列をまたぐ場合は例外を投げる（設計ミスの検出）', function () {
  var badRanges = [{ startCol: 1, span: 2, major: '不正な結合' }]; // A列とB列にまたがる
  assert.throws(function () {
    Logic.validateMergeRanges(badRanges, 1);
  });
});

// ============================================================
// CSV / HTMLエスケープ
// ============================================================
test('CSV: カンマや改行を含む値はダブルクォートでエスケープされる', function () {
  var csv = Logic.buildCsv(['a', 'b'], [['1,2', 'line1\nline2'], ['x"y', 'z']]);
  assert.ok(csv.indexOf('"1,2"') !== -1);
  assert.ok(csv.indexOf('"line1\nline2"') !== -1);
  assert.ok(csv.indexOf('"x""y"') !== -1);
});

test('HTMLエスケープ', function () {
  assert.strictEqual(Logic.escapeHtml('<script>&"\''), '&lt;script&gt;&amp;&quot;&#39;');
});

// ============================================================
// 週次・月次の集計（複数日合算）
// ============================================================
test('複数日の集計：機変等が単純合算され、MB総販等も再計算される', function () {
  var day1 = Logic.buildDayModel({ date: '2026-08-10', DCM_kihen: 2, DCM_mnp: 1 });
  var day2 = Logic.buildDayModel({ date: '2026-08-11', DCM_kihen: 3, DCM_mnp: 2 });
  var agg = Logic.aggregateDayModels([day1, day2]);
  assert.strictEqual(agg.carriers.DCM.kihen, 5);
  assert.strictEqual(agg.computed.mb.byCarrier.DCM, 5 + 3);
  assert.strictEqual(agg.dayCount, 2);
});

test('入力値のクランプ：空欄は0、マイナスは0、小数は切り捨て', function () {
  var model = Logic.buildDayModel({ date: '2026-08-10', DCM_kihen: '', DCM_junshinki: -5, DCM_mnp: 3.9 });
  assert.strictEqual(model.carriers.DCM.kihen, 0);
  assert.strictEqual(model.carriers.DCM.junshinki, 0);
  assert.strictEqual(model.carriers.DCM.mnp, 3);
});

// ============================================================
// 結果表示
// ============================================================
console.log('');
console.log('=== logic.js local test results ===');
console.log('passed: ' + passed + ' / failed: ' + failed);
if (failed > 0) {
  failures.forEach(function (f) {
    console.log('');
    console.log('FAIL: ' + f.name);
    console.log(f.error && f.error.stack ? f.error.stack : f.error);
  });
  process.exit(1);
} else {
  console.log('ALL TESTS PASSED');
  process.exit(0);
}

/**
 * コード.gs（build.js による自動結合ファイル）
 * ------------------------------------------------------------
 * このファイルは src/logic.js → src/sheets.js → src/webapp.js → src/report.js
 * を結合したものです。直接このファイルを編集しないでください。
 * 修正するときは src/ 配下の該当ファイルを直してから `node build.js` を
 * 実行し、生成しなおしたこのファイルをスクリプトエディタに貼り直してください。
 *
 * 【項目を追加・変更したいときにどこを直すか】
 *   - 副商材の項目を増やす   → 下の Logic.SUB_DEF（src/logic.js）に1行追加
 *   - 屋号商材を増やす       → Logic.YAGOU_PRODUCTS（src/logic.js）に1行追加
 *   - MNP移転元の選択肢を増やす → Logic.SOURCES（src/logic.js）に1行追加
 *   - キャリア構成を変える（通常は不要） → Logic.CARRIERS / CARRIER_GROUP
 *   これらはすべて「見出し名（label）」から列位置を解決する設計なので、
 *   配列の途中に追加しても既存のシート列や集計式は壊れません。
 *   追加後は「初期セットアップ」を再実行してください（既存データは消えません）。
 * ------------------------------------------------------------
 */

// ============================================================
// logic.js
// ============================================================
/**
 * logic.js
 * ------------------------------------------------------------
 * 集計ロジックのみを持つ「純関数」ファイルです。
 * SpreadsheetApp / HtmlService など GAS 固有の API は一切呼び出しません。
 * そのため node コマンドだけでローカルテストできます（test/run.js 参照）。
 *
 * 【項目を追加・変更したいときはここを直す】
 *   - 副商材を増やす   → SUB_DEF 配列に1行追加する（type:'input'）。
 *                        小計（◯◯計）を増やす場合は type:'sum' の行を追加し、
 *                        of に集計したい項目の key を並べる。
 *   - 屋号商材を増やす → YAGOU_PRODUCTS 配列に1行追加する。
 *   - MNP移転元を増やす → SOURCES 配列に1行追加する。
 *   - キャリア構成を変える（通常は不要） → CARRIERS / CARRIER_GROUP を編集する。
 *   列の位置は「見出し名（label）」から動的に解決するので、
 *   配列の途中に追加しても既存の列参照は壊れません。
 * ------------------------------------------------------------
 */
var Logic = (function () {
  'use strict';

  // ============================================================
  // 基本マスタ
  // ============================================================
  var CARRIERS = ['DCM', 'SB', 'YM', 'au', 'UQ'];
  var CARRIER_GROUP = { DCM: 'DCM', SB: 'SBYM', YM: 'SBYM', au: 'AUUQ', UQ: 'AUUQ' };
  var GROUPS = ['DCM', 'SBYM', 'AUUQ'];
  var GROUP_LABEL = { DCM: 'DCM', SBYM: 'SB/YM', AUUQ: 'AU/UQ' };
  var GROUP_COLOR = { DCM: '#D8272C', SBYM: '#6E6E73', AUUQ: '#F39A1F' };

  // 同一グループ内の番号移行（＝「リク」）
  var MIGRATIONS = [
    { key: 'migSbToYm', label: '移行_SB→YM', group: 'SBYM', text: 'SB→YM' },
    { key: 'migYmToSb', label: '移行_YM→SB', group: 'SBYM', text: 'YM→SB' },
    { key: 'migAuToUq', label: '移行_au→UQ', group: 'AUUQ', text: 'au→UQ' },
    { key: 'migUqToAu', label: '移行_UQ→au', group: 'AUUQ', text: 'UQ→au' }
  ];

  // MNP 移転元キャリア（任意入力）
  var SOURCES = [
    { key: 'srcDocomo', label: '元_docomo', text: 'docomo' },
    { key: 'srcAu', label: '元_au', text: 'au' },
    { key: 'srcUq', label: '元_UQ', text: 'UQ' },
    { key: 'srcSb', label: '元_SB', text: 'SB' },
    { key: 'srcYm', label: '元_YM', text: 'YM' },
    { key: 'srcRakuten', label: '元_楽天モバイル', text: '楽天モバイル' },
    { key: 'srcMvno', label: '元_MVNO格安', text: 'MVNO・格安' },
    { key: 'srcOther', label: '元_その他', text: 'その他' }
  ];

  // 副商材定義（キャリア別・カテゴリ別）
  // type: 'input'（入力欄） / 'sum'（自動計算・of で指定した key の合計）
  var SUB_DEF = [
    // --- DCM ---
    { key: 'dcm_hikari1g', label: '副_d光1G', group: 'DCM', category: '(Wi-Fi)', name: 'd光 1Gbps', type: 'input' },
    { key: 'dcm_hikari10g', label: '副_d光10G', group: 'DCM', category: '(Wi-Fi)', name: 'd光 10Gbps', type: 'input' },
    { key: 'dcm_hikariTotal', label: null, group: 'DCM', category: '(Wi-Fi)', name: 'd光計', type: 'sum', of: ['dcm_hikari1g', 'dcm_hikari10g'] },
    { key: 'dcm_home5g', label: '副_HOME5G', group: 'DCM', category: '(Wi-Fi)', name: 'HOME 5G', type: 'input' },

    { key: 'dcm_denkiBasic', label: '副_ドコモでんきBasic', group: 'DCM', category: '(エネルギー事業)', name: 'ドコモでんき Basic', type: 'input' },
    { key: 'dcm_denkiGreen', label: '副_ドコモでんきGreen', group: 'DCM', category: '(エネルギー事業)', name: 'ドコモでんき Green', type: 'input' },
    { key: 'dcm_denkiTotal', label: null, group: 'DCM', category: '(エネルギー事業)', name: 'ドコモでんき計', type: 'sum', of: ['dcm_denkiBasic', 'dcm_denkiGreen'] },
    { key: 'dcm_gas', label: '副_ドコモガス', group: 'DCM', category: '(エネルギー事業)', name: 'ドコモガス', type: 'input' },
    { key: 'dcm_energyTotal', label: null, group: 'DCM', category: '(エネルギー事業)', name: 'エネルギー計', type: 'sum', of: ['dcm_denkiBasic', 'dcm_denkiGreen', 'dcm_gas'] },

    { key: 'dcm_cardR', label: '副_dカードR', group: 'DCM', category: '(クレジット事業)', name: 'dカード(R)', type: 'input' },
    { key: 'dcm_cardG', label: '副_dカードG', group: 'DCM', category: '(クレジット事業)', name: 'dカード(G)', type: 'input' },
    { key: 'dcm_cardGu', label: '副_dカードGU', group: 'DCM', category: '(クレジット事業)', name: 'dカード(GU)', type: 'input' },
    { key: 'dcm_cardP', label: '副_dカードP', group: 'DCM', category: '(クレジット事業)', name: 'dカード(P)', type: 'input' },
    { key: 'dcm_cardTotal', label: null, group: 'DCM', category: '(クレジット事業)', name: 'dカード計', type: 'sum', of: ['dcm_cardR', 'dcm_cardG', 'dcm_cardGu', 'dcm_cardP'] },

    // --- AU/UQ ---
    { key: 'auuq_hikariHome', label: '副_au光ホーム', group: 'AUUQ', category: '(Wi-Fi)', name: 'au光 ホームタイプ', type: 'input' },
    { key: 'auuq_hikariMansion', label: '副_au光マンション', group: 'AUUQ', category: '(Wi-Fi)', name: 'au光 マンションタイプ', type: 'input' },
    { key: 'auuq_hikariTotal', label: null, group: 'AUUQ', category: '(Wi-Fi)', name: 'au光計', type: 'sum', of: ['auuq_hikariHome', 'auuq_hikariMansion'] },
    { key: 'auuq_biglobe1g', label: '副_Biglobe光1G', group: 'AUUQ', category: '(Wi-Fi)', name: 'Biglobe光 1Gbps', type: 'input' },
    { key: 'auuq_biglobe10g', label: '副_Biglobe光10G', group: 'AUUQ', category: '(Wi-Fi)', name: 'Biglobe光 10Gbps', type: 'input' },
    { key: 'auuq_biglobeTotal', label: null, group: 'AUUQ', category: '(Wi-Fi)', name: 'Biglobe光計', type: 'sum', of: ['auuq_biglobe1g', 'auuq_biglobe10g'] },
    { key: 'auuq_wimax', label: '副_WiMax', group: 'AUUQ', category: '(Wi-Fi)', name: 'WiMax', type: 'input' },

    { key: 'auuq_denki', label: '副_auでんき', group: 'AUUQ', category: '(エネルギー事業)', name: 'auでんき', type: 'input' },

    // --- SB/YM ---
    { key: 'sbym_hikari1g', label: '副_SB光1G', group: 'SBYM', category: '(Wi-Fi)', name: 'SB光 1Gbps', type: 'input' },
    { key: 'sbym_hikari10g', label: '副_SB光10G', group: 'SBYM', category: '(Wi-Fi)', name: 'SB光 10Gbps', type: 'input' },
    { key: 'sbym_hikariTotal', label: null, group: 'SBYM', category: '(Wi-Fi)', name: 'SB光計', type: 'sum', of: ['sbym_hikari1g', 'sbym_hikari10g'] },
    { key: 'sbym_air', label: '副_SBAir', group: 'SBYM', category: '(Wi-Fi)', name: 'SB Air', type: 'input' },

    { key: 'sbym_denki', label: '副_おうちでんき', group: 'SBYM', category: '(エネルギー事業)', name: 'おうちでんき', type: 'input' },
    { key: 'sbym_gas', label: '副_おうちガス', group: 'SBYM', category: '(エネルギー事業)', name: 'おうちガス', type: 'input' },
    { key: 'sbym_energyTotal', label: null, group: 'SBYM', category: '(エネルギー事業)', name: 'エネルギー計', type: 'sum', of: ['sbym_denki', 'sbym_gas'] }
  ];

  // 屋号商材（3商材 × 3区分）
  var YAGOU_PRODUCTS = [
    { key: 'kikiHoken', label: 'つながる機器保険' },
    { key: 'kakeiSodan', label: '家計相談' },
    { key: 'marugotoHoken', label: 'まるごと保険' }
  ];

  function yagouKey(productKey, group) {
    return 'yagou_' + productKey + '_' + group;
  }
  function yagouLabel(productKey, group) {
    return '屋号_' + (YAGOU_PRODUCTS.filter(function (p) { return p.key === productKey; })[0] || {}).label + '_' + group;
  }

  // ============================================================
  // 数値・日付の正規化（不具合(2)(4)対策の要）
  // ============================================================

  // 空欄→0、マイナス→0、小数→切り捨ての非負整数に丸める
  function clampNonNegInt(v) {
    var n = Number(v);
    if (v === '' || v === null || v === undefined || isNaN(n)) return 0;
    n = Math.floor(n);
    if (n < 0) return 0;
    return n;
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function formatDateStr(y, m, d) {
    return y + '-' + pad2(m) + '-' + pad2(d);
  }

  // GAS の Date / スプレッドシートのシリアル値 / '2026/8/5' / '2026-08-05' を
  // すべて 'yyyy-MM-dd' の文字列に正規化する。
  // 日付が Date 型に化ける不具合(2)対策の要となる関数。呼び出し側は必ずこれを通すこと。
  function toDateStr(v) {
    if (v === null || v === undefined || v === '') return '';

    // Date オブジェクト（GASのDateも含む）
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
      return formatDateStr(v.getFullYear(), v.getMonth() + 1, v.getDate());
    }

    // スプレッドシートのシリアル値（数値）: 1899-12-30 を 0 とする
    if (typeof v === 'number') {
      var epoch = Date.UTC(1899, 11, 30);
      var ms = epoch + Math.round(v) * 86400000;
      var d2 = new Date(ms);
      return formatDateStr(d2.getUTCFullYear(), d2.getUTCMonth() + 1, d2.getUTCDate());
    }

    var s = String(v).trim();

    // 'yyyy-MM-dd' or 'yyyy/M/d' 系
    var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (m) {
      return formatDateStr(Number(m[1]), Number(m[2]), Number(m[3]));
    }

    // 上記以外はDateパースを試みる
    var parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      return formatDateStr(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
    }
    return s;
  }

  // 'yyyy-MM-dd' 文字列 → ローカルの Date オブジェクト（時刻部分は 00:00:00）
  function parseDateStr(dateStr) {
    var parts = String(dateStr).split('-');
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }

  var WEEKDAY_JA = ['日', '月', '火', '水', '木', '金', '土'];
  function getWeekdayJa(dateStr) {
    var d = parseDateStr(dateStr);
    return WEEKDAY_JA[d.getDay()];
  }

  // 月曜起算の週の開始日（月曜）を 'yyyy-MM-dd' で返す。日曜は前の週に属する。
  function getMonday(dateStr) {
    var d = parseDateStr(dateStr);
    var dow = d.getDay(); // 0=日 .. 6=土
    var diff = (dow === 0 ? -6 : 1 - dow);
    d.setDate(d.getDate() + diff);
    return formatDateStr(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  function addDays(dateStr, n) {
    var d = parseDateStr(dateStr);
    d.setDate(d.getDate() + n);
    return formatDateStr(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  function getWeekEnd(mondayStr) {
    return addDays(mondayStr, 6);
  }

  // 指定の年月に含まれる日を、月曜起算の週でグルーピングし、
  // その月内で最初に出てくる順に「n週目」の番号を振る。
  function getWeeksOfMonth(year, month) {
    var lastDay = new Date(year, month, 0).getDate();
    var order = [];
    var seen = {};
    for (var day = 1; day <= lastDay; day++) {
      var dateStr = formatDateStr(year, month, day);
      var monday = getMonday(dateStr);
      if (!seen[monday]) {
        seen[monday] = true;
        order.push(monday);
      }
    }
    return order.map(function (monday, idx) {
      return { index: idx + 1, weekStart: monday, weekEnd: getWeekEnd(monday) };
    });
  }

  function getWeekIndexInMonth(dateStr) {
    var d = parseDateStr(dateStr);
    var weeks = getWeeksOfMonth(d.getFullYear(), d.getMonth() + 1);
    var monday = getMonday(dateStr);
    for (var i = 0; i < weeks.length; i++) {
      if (weeks[i].weekStart === monday) return weeks[i].index;
    }
    return null;
  }

  // ============================================================
  // 1日分の入力値を正規化・集計する
  // ============================================================

  // raw: { date, event, memo, DCM_kihen, DCM_junshinki, DCM_mnp, DCM_uchiSim, ... (CARRIERS分)
  //        migSbToYm, migYmToSb, migAuToUq, migUqToAu,
  //        srcDocomo ... (SOURCES分),
  //        SUB_DEF の type:'input' の key たち,
  //        yagou_{product}_{group} } を受け取り、
  // 正規化済みの値と computed（自動計算結果）を持つモデルを返す。
  function buildDayModel(raw) {
    raw = raw || {};
    var model = {
      date: toDateStr(raw.date),
      event: raw.event ? String(raw.event) : '',
      memo: raw.memo ? String(raw.memo) : '',
      carriers: {},
      migrations: {},
      sources: {},
      sub: {},
      yagou: {}
    };
    model.weekday = model.date ? getWeekdayJa(model.date) : '';

    CARRIERS.forEach(function (c) {
      model.carriers[c] = {
        kihen: clampNonNegInt(raw[c + '_kihen']),
        junshinki: clampNonNegInt(raw[c + '_junshinki']),
        mnp: clampNonNegInt(raw[c + '_mnp']),
        uchiSim: clampNonNegInt(raw[c + '_uchiSim'])
      };
    });

    MIGRATIONS.forEach(function (m) {
      model.migrations[m.key] = clampNonNegInt(raw[m.key]);
    });

    SOURCES.forEach(function (s) {
      model.sources[s.key] = clampNonNegInt(raw[s.key]);
    });

    SUB_DEF.forEach(function (def) {
      if (def.type === 'input') {
        model.sub[def.key] = clampNonNegInt(raw[def.key]);
      }
    });

    YAGOU_PRODUCTS.forEach(function (p) {
      GROUPS.forEach(function (g) {
        var k = yagouKey(p.key, g);
        model.yagou[k] = clampNonNegInt(raw[k]);
      });
    });

    model.computed = computeAll(model);
    return model;
  }

  // ------------------------------------------------------------
  // MB総販（機変＋純新規＋MNP）
  // ------------------------------------------------------------
  function computeMbSales(model) {
    var byCarrier = {};
    var byGroup = { DCM: 0, SBYM: 0, AUUQ: 0 };
    var total = 0;
    CARRIERS.forEach(function (c) {
      var v = model.carriers[c];
      var mb = v.kihen + v.junshinki + v.mnp;
      byCarrier[c] = mb;
      byGroup[CARRIER_GROUP[c]] += mb;
      total += mb;
    });
    return { byCarrier: byCarrier, byGroup: byGroup, total: total };
  }

  // ------------------------------------------------------------
  // MNP（リク込み／リク除き／内SIMのみ）
  // ------------------------------------------------------------
  function computeMnp(model) {
    var rikuKomiByGroup = { DCM: 0, SBYM: 0, AUUQ: 0 };
    var uchiSimByGroup = { DCM: 0, SBYM: 0, AUUQ: 0 };
    CARRIERS.forEach(function (c) {
      var g = CARRIER_GROUP[c];
      rikuKomiByGroup[g] += model.carriers[c].mnp;
      uchiSimByGroup[g] += model.carriers[c].uchiSim;
    });

    var migTotalByGroup = { DCM: 0, SBYM: 0, AUUQ: 0 };
    MIGRATIONS.forEach(function (m) {
      migTotalByGroup[m.group] += model.migrations[m.key];
    });

    var rikuNukiByGroup = {};
    GROUPS.forEach(function (g) {
      // DCM はグループ内移行の概念がないため常にリク込み＝リク除き
      rikuNukiByGroup[g] = g === 'DCM' ? rikuKomiByGroup.DCM : rikuKomiByGroup[g] - migTotalByGroup[g];
    });

    var sourceTotal = 0;
    SOURCES.forEach(function (s) { sourceTotal += model.sources[s.key]; });

    return {
      rikuKomiByGroup: rikuKomiByGroup,
      rikuKomiTotal: sum(rikuKomiByGroup),
      rikuNukiByGroup: rikuNukiByGroup,
      rikuNukiTotal: sum(rikuNukiByGroup),
      uchiSimByGroup: uchiSimByGroup,
      uchiSimTotal: sum(uchiSimByGroup),
      migTotalByGroup: migTotalByGroup,
      migTotal: sum(migTotalByGroup),
      sourceTotal: sourceTotal
    };
  }

  function sum(obj) {
    var t = 0;
    Object.keys(obj).forEach(function (k) { t += obj[k]; });
    return t;
  }

  // ------------------------------------------------------------
  // 副商材（sum 型は of の合計を再帰的に解決する。参照順に依存しない）
  // ------------------------------------------------------------
  function computeSubProducts(model) {
    var defByKey = {};
    SUB_DEF.forEach(function (d) { defByKey[d.key] = d; });
    var resolved = {};

    function resolve(key) {
      if (resolved.hasOwnProperty(key)) return resolved[key];
      var def = defByKey[key];
      var value;
      if (!def) {
        value = 0;
      } else if (def.type === 'input') {
        value = model.sub[key] || 0;
      } else {
        value = def.of.reduce(function (acc, k) { return acc + resolve(k); }, 0);
      }
      resolved[key] = value;
      return value;
    }

    SUB_DEF.forEach(function (d) { resolve(d.key); });

    // 副商材の総合計は type:'input' のみを足す（sum 列を含めると二重計上になる）
    var grandTotal = 0;
    var byGroup = { DCM: 0, SBYM: 0, AUUQ: 0 };
    SUB_DEF.forEach(function (d) {
      if (d.type === 'input') {
        grandTotal += resolved[d.key];
        byGroup[d.group] += resolved[d.key];
      }
    });

    return { values: resolved, grandTotal: grandTotal, byGroup: byGroup };
  }

  // ------------------------------------------------------------
  // 屋号商材
  // ------------------------------------------------------------
  function computeYagouTotals(model) {
    var byProduct = {};
    var byGroup = { DCM: 0, SBYM: 0, AUUQ: 0 };
    var grandTotal = 0;
    YAGOU_PRODUCTS.forEach(function (p) {
      var perGroup = {};
      var productTotal = 0;
      GROUPS.forEach(function (g) {
        var v = model.yagou[yagouKey(p.key, g)] || 0;
        perGroup[g] = v;
        productTotal += v;
        byGroup[g] += v;
      });
      byProduct[p.key] = { perGroup: perGroup, total: productTotal };
      grandTotal += productTotal;
    });
    return { byProduct: byProduct, byGroup: byGroup, grandTotal: grandTotal };
  }

  // ------------------------------------------------------------
  // 3区分の構成比（0件のときも 0除算にせず 0.0% を返す）
  // ------------------------------------------------------------
  function computeShare(byGroup) {
    var total = sum(byGroup);
    var result = { total: total, byGroup: {} };
    GROUPS.forEach(function (g) {
      var count = byGroup[g] || 0;
      var pct = total > 0 ? (count / total) * 100 : 0;
      result.byGroup[g] = { count: count, pct: pct };
    });
    return result;
  }

  function computeAll(model) {
    var mb = computeMbSales(model);
    var mnp = computeMnp(model);
    var sub = computeSubProducts(model);
    var yagou = computeYagouTotals(model);
    return {
      mb: mb,
      mnp: mnp,
      sub: sub,
      yagou: yagou,
      mbShare: computeShare(mb.byGroup),
      rikuKomiShare: computeShare(mnp.rikuKomiByGroup),
      rikuNukiShare: computeShare(mnp.rikuNukiByGroup)
    };
  }

  // ============================================================
  // 複数日の集計（週次・月次はこれで日次を合算してから computeAll と同じ計算式を再利用する）
  // ============================================================
  function aggregateDayModels(models) {
    var agg = {
      date: '',
      event: '',
      memo: '',
      carriers: {},
      migrations: {},
      sources: {},
      sub: {},
      yagou: {}
    };
    CARRIERS.forEach(function (c) {
      agg.carriers[c] = { kihen: 0, junshinki: 0, mnp: 0, uchiSim: 0 };
    });
    MIGRATIONS.forEach(function (m) { agg.migrations[m.key] = 0; });
    SOURCES.forEach(function (s) { agg.sources[s.key] = 0; });
    SUB_DEF.forEach(function (d) { if (d.type === 'input') agg.sub[d.key] = 0; });
    YAGOU_PRODUCTS.forEach(function (p) {
      GROUPS.forEach(function (g) { agg.yagou[yagouKey(p.key, g)] = 0; });
    });

    var events = [];
    models.forEach(function (m) {
      CARRIERS.forEach(function (c) {
        agg.carriers[c].kihen += m.carriers[c].kihen;
        agg.carriers[c].junshinki += m.carriers[c].junshinki;
        agg.carriers[c].mnp += m.carriers[c].mnp;
        agg.carriers[c].uchiSim += m.carriers[c].uchiSim;
      });
      MIGRATIONS.forEach(function (mg) { agg.migrations[mg.key] += m.migrations[mg.key]; });
      SOURCES.forEach(function (s) { agg.sources[s.key] += m.sources[s.key]; });
      SUB_DEF.forEach(function (d) { if (d.type === 'input') agg.sub[d.key] += m.sub[d.key]; });
      YAGOU_PRODUCTS.forEach(function (p) {
        GROUPS.forEach(function (g) {
          var k = yagouKey(p.key, g);
          agg.yagou[k] += m.yagou[k];
        });
      });
      if (m.event && events.indexOf(m.event) === -1) events.push(m.event);
    });

    agg.event = events.join('／');
    agg.computed = computeAll(agg);
    agg.dayCount = models.length;
    return agg;
  }

  // ============================================================
  // 「日次ログ」シートの列定義（列の並び順の唯一の正）
  // 項目を追加したいときは、該当する定義配列（CARRIERS/MIGRATIONS/SOURCES/
  // SUB_DEF/YAGOU_PRODUCTS）に追加すれば、この関数が自動的に列を追加する。
  // ============================================================
  var YAGOU_GROUP_LABEL = { DCM: 'DCM', AUUQ: 'auUQ', SBYM: 'SBYM' };
  var YAGOU_GROUP_ORDER = ['DCM', 'AUUQ', 'SBYM'];

  function getLogHeaderDefs() {
    var defs = [];
    defs.push({ key: 'date', label: '日付', major: '日付', type: 'date' });
    defs.push({ key: 'weekday', label: '曜日', major: '基本情報', type: 'text' });
    defs.push({ key: 'event', label: 'イベント', major: '基本情報', type: 'text' });
    defs.push({ key: 'memo', label: 'メモ', major: '基本情報', type: 'text' });

    CARRIERS.forEach(function (c) {
      defs.push({ key: c + '_kihen', label: c + '_機変', major: c, type: 'int' });
      defs.push({ key: c + '_junshinki', label: c + '_純新規', major: c, type: 'int' });
      defs.push({ key: c + '_mnp', label: c + '_MNP', major: c, type: 'int' });
      defs.push({ key: c + '_uchiSim', label: c + '_内SIMのみ', major: c, type: 'int' });
    });

    MIGRATIONS.forEach(function (m) {
      defs.push({ key: m.key, label: m.label, major: '番号移行', type: 'int' });
    });

    SOURCES.forEach(function (s) {
      defs.push({ key: s.key, label: s.label, major: '移転元', type: 'int' });
    });

    SUB_DEF.forEach(function (d) {
      if (d.type !== 'input') return; // 計(sum)列はシートに持たず、日次サマリー側で数式計算する
      defs.push({ key: d.key, label: d.label, major: GROUP_LABEL[d.group] + d.category, type: 'int' });
    });

    YAGOU_PRODUCTS.forEach(function (p) {
      YAGOU_GROUP_ORDER.forEach(function (g) {
        defs.push({ key: yagouKey(p.key, g), label: '屋号_' + p.label + '_' + YAGOU_GROUP_LABEL[g], major: '屋号:' + p.label, type: 'int' });
      });
    });

    defs.push({ key: 'updatedAt', label: '更新日時', major: '更新日時', type: 'text' });
    return defs;
  }

  // headerDefs から2行分のヘッダー配列を組み立てる（row1=大分類・結合用、row2=見出し名）
  function buildHeaderRows(headerDefs) {
    var row2 = headerDefs.map(function (d) { return d.label; });
    var row1 = headerDefs.map(function () { return ''; });
    var i = 0;
    while (i < headerDefs.length) {
      var major = headerDefs[i].major;
      row1[i] = major || '';
      i++;
      while (i < headerDefs.length && headerDefs[i].major === major) i++;
    }
    return { row1: row1, row2: row2 };
  }

  // raw（フォーム入力オブジェクト）→ シート1行分の配列
  function rowFromRaw(raw, headerDefs) {
    raw = raw || {};
    return headerDefs.map(function (d) {
      if (d.key === 'weekday') {
        var ds = toDateStr(raw.date);
        return ds ? getWeekdayJa(ds) : '';
      }
      if (d.type === 'date') return toDateStr(raw[d.key]);
      if (d.type === 'int') return clampNonNegInt(raw[d.key]);
      return raw[d.key] === undefined || raw[d.key] === null ? '' : String(raw[d.key]);
    });
  }

  // シート1行分の配列 → raw（buildDayModel にそのまま渡せるオブジェクト）
  function rawFromRow(rowValues, headerDefs) {
    var raw = {};
    headerDefs.forEach(function (d, i) {
      var v = rowValues[i];
      raw[d.key] = d.type === 'date' ? toDateStr(v) : v;
    });
    return raw;
  }

  // ============================================================
  // シート列マップ（不具合(4)対策：見出し名から列位置を解決する）
  // ============================================================
  // headerRow: 2行目の見出し配列（1始まりではなく0始まりの配列。A列=index0）
  // 戻り値: { '見出し名': 列番号(1始まり) }
  function getColMap(headerRow) {
    var map = {};
    headerRow.forEach(function (label, idx) {
      if (label !== '' && label !== null && label !== undefined) {
        map[label] = idx + 1;
      }
    });
    return map;
  }

  // ============================================================
  // 大分類セルの結合範囲（不具合(1)対策）
  // ranges: [{ startCol, span, major }] （1始まり列番号）
  // frozenCols: 固定されている列数（例: 1 = A列のみ固定）
  // 固定列と非固定列をまたぐ結合があれば例外を投げる。
  // ============================================================
  function validateMergeRanges(ranges, frozenCols) {
    ranges.forEach(function (r) {
      var endCol = r.startCol + r.span - 1;
      var crossesBoundary = r.startCol <= frozenCols && endCol > frozenCols;
      if (crossesBoundary) {
        throw new Error(
          '大分類「' + r.major + '」の結合範囲(列' + r.startCol + '〜' + endCol + ')が' +
          '固定列の境界(列' + frozenCols + '/' + (frozenCols + 1) + ')をまたいでいます。'
        );
      }
    });
    return true;
  }

  // ヘッダー定義（major, label の並び）から結合範囲を組み立てる。
  // 同じ major が連続する区間を1つの結合range にする。major が '' の場合は結合しない。
  function buildMergeRanges(headerDefs) {
    var ranges = [];
    var i = 0;
    while (i < headerDefs.length) {
      var major = headerDefs[i].major;
      if (!major) { i++; continue; }
      var start = i;
      while (i < headerDefs.length && headerDefs[i].major === major) i++;
      var span = i - start;
      if (span > 1) {
        ranges.push({ startCol: start + 1, span: span, major: major });
      }
    }
    return ranges;
  }

  // ============================================================
  // CSV
  // ============================================================
  function csvEscapeCell(v) {
    var s = v === null || v === undefined ? '' : String(v);
    if (/[",\n]/.test(s)) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function buildCsv(headerRow, dataRows) {
    var lines = [headerRow.map(csvEscapeCell).join(',')];
    dataRows.forEach(function (row) {
      lines.push(row.map(csvEscapeCell).join(','));
    });
    return lines.join('\r\n');
  }

  // ============================================================
  // HTMLエスケープ（フォーム表示文字列に必ず通す）
  // ============================================================
  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // ============================================================
  // 報告テキスト生成（0件の項目・カテゴリは行ごと省略）
  // ============================================================
  function fmtPct(pct) {
    return pct.toFixed(1) + '%';
  }

  function shareLine(label, share) {
    if (share.total === 0) return label + '（計0件）：実績なし';
    var parts = GROUPS.map(function (g) {
      var s = share.byGroup[g];
      return GROUP_LABEL[g] + ' ' + s.count + '件/' + fmtPct(s.pct);
    });
    return label + '（計' + share.total + '件）：' + parts.join('、');
  }

  function buildTitle(year, month, weekIndex, weekStart, weekEnd, dayCount) {
    var range = shortDate(weekStart) + '〜' + shortDate(weekEnd);
    return '【' + year + '年' + month + '月 ' + weekIndex + '週目 ' + range + '（' + dayCount + '日）】';
  }

  function shortDate(dateStr) {
    var p = dateStr.split('-');
    return Number(p[1]) + '/' + Number(p[2]);
  }

  function generateReportText(agg, opts) {
    opts = opts || {};
    var c = agg.computed;
    var lines = [];
    lines.push(opts.title || '【集計】');

    // ---- MB総販 ----
    lines.push('■MB総販');
    lines.push(shareLine('MB総販', c.mbShare));
    if (c.mb.total > 0) {
      var kihenTotal = 0, junTotal = 0;
      CARRIERS.forEach(function (cc) { kihenTotal += agg.carriers[cc].kihen; junTotal += agg.carriers[cc].junshinki; });
      lines.push('内訳：機変 ' + kihenTotal + '件＋純新規 ' + junTotal + '件＋MNP ' + c.mnp.rikuKomiTotal + '件＝' + c.mb.total + '件');
      GROUPS.forEach(function (g) {
        var kihen = 0, jun = 0;
        CARRIERS.forEach(function (cc) { if (CARRIER_GROUP[cc] === g) { kihen += agg.carriers[cc].kihen; jun += agg.carriers[cc].junshinki; } });
        var mnp = c.mnp.rikuKomiByGroup[g];
        var mb = c.mb.byGroup[g];
        if (mb > 0) {
          lines.push('　' + GROUP_LABEL[g] + '：機変' + kihen + '／純新規' + jun + '／MNP' + mnp + '＝' + mb + '件');
        }
      });
    }

    // ---- MNP ----
    lines.push('■MNP');
    lines.push(shareLine('リク込み', c.rikuKomiShare));
    lines.push(shareLine('リク除き', c.rikuNukiShare));
    if (c.mnp.uchiSimTotal > 0) {
      lines.push('内SIMのみ：計' + c.mnp.uchiSimTotal + '件（' +
        GROUPS.map(function (g) { return GROUP_LABEL[g] + ' ' + c.mnp.uchiSimByGroup[g]; }).join('／') + '）');
    }
    if (c.mnp.migTotal > 0) {
      var migParts = MIGRATIONS.filter(function (m) { return agg.migrations[m.key] > 0; })
        .map(function (m) { return m.text + ' ' + agg.migrations[m.key] + '件'; });
      lines.push('番号移行：' + migParts.join('／') + '（計' + c.mnp.migTotal + '件）');
    }
    if (c.mnp.sourceTotal > 0) {
      var srcParts = SOURCES.filter(function (s) { return agg.sources[s.key] > 0; })
        .map(function (s) { return s.text + ' ' + agg.sources[s.key] + '件'; });
      lines.push('移転元：' + srcParts.join('／') + '（計' + c.mnp.sourceTotal + '件）');
    }

    // ---- 副商材 ----
    lines.push('■副商材（計' + c.sub.grandTotal + '件）');
    if (c.sub.grandTotal > 0) {
      GROUPS.forEach(function (g) {
        if (c.sub.byGroup[g] === 0) return;
        lines.push(GROUP_LABEL[g]);
        var categories = [];
        SUB_DEF.forEach(function (d) {
          if (d.group === g && categories.indexOf(d.category) === -1) categories.push(d.category);
        });
        categories.forEach(function (cat) {
          var defsInCat = SUB_DEF.filter(function (d) { return d.group === g && d.category === cat; });
          var hasAny = defsInCat.some(function (d) { return c.sub.values[d.key] > 0; });
          if (!hasAny) return;
          var parts = defsInCat
            .filter(function (d) { return c.sub.values[d.key] > 0; })
            .map(function (d) {
              var suffix = d.type === 'sum' ? '　' + d.name + ' ' + c.sub.values[d.key] + '件' : d.name + ' ' + c.sub.values[d.key] + '件';
              return suffix;
            })
            .join('／')
            .replace(/／　/g, '　');
          lines.push('　' + cat + ' ' + parts);
        });
      });
    }

    // ---- 屋号商材 ----
    lines.push('■屋号商材（計' + c.yagou.grandTotal + '件）');
    if (c.yagou.grandTotal > 0) {
      YAGOU_PRODUCTS.forEach(function (p) {
        var pd = c.yagou.byProduct[p.key];
        if (pd.total === 0) return;
        var parts = GROUPS.filter(function (g) { return pd.perGroup[g] > 0; })
          .map(function (g) { return GROUP_LABEL[g] + ' ' + pd.perGroup[g] + '件'; });
        lines.push(p.label + '：' + parts.join('／') + '（計' + pd.total + '件）');
      });
    }

    if (opts.staffing) {
      lines.push('稼働人数：' + opts.staffing);
    }
    if (agg.event) {
      lines.push('イベント：' + agg.event);
    }
    lines.push('※リク除き＝リク込みから同一グループ内の番号移行を控除。DCMは番号移行の概念なし。');

    return lines.join('\n');
  }

  // ============================================================
  // 週次・月次集計シート用の表示行データ（0件の項目は行ごと省略する）
  // 表示用の並びだけを組み立てる純関数。実際のセル書き込みは sheets.js が行う。
  // ============================================================
  function buildAggBlockRows(agg) {
    var c = agg.computed;
    var rows = [];

    function shareRow(label, share) {
      var shareText = share.total > 0
        ? GROUPS.map(function (g) { return GROUP_LABEL[g] + ' ' + fmtPct(share.byGroup[g].pct); }).join('／')
        : '';
      rows.push({
        label: label,
        dcm: share.byGroup.DCM.count, sbym: share.byGroup.SBYM.count, auuq: share.byGroup.AUUQ.count,
        total: share.total, share: shareText, note: ''
      });
    }

    shareRow('MB総販', c.mbShare);
    if (c.mb.total > 0) {
      GROUPS.forEach(function (g) {
        var kihen = 0, jun = 0;
        CARRIERS.forEach(function (cc) {
          if (CARRIER_GROUP[cc] === g) { kihen += agg.carriers[cc].kihen; jun += agg.carriers[cc].junshinki; }
        });
        var mnp = c.mnp.rikuKomiByGroup[g];
        var mb = c.mb.byGroup[g];
        if (mb > 0) {
          rows.push({ label: '　内訳:' + GROUP_LABEL[g], dcm: '', sbym: '', auuq: '', total: mb, share: '', note: '機変' + kihen + '／純新規' + jun + '／MNP' + mnp });
        }
      });
    }

    shareRow('MNPリク込み', c.rikuKomiShare);
    shareRow('MNPリク除き', c.rikuNukiShare);

    if (c.mnp.uchiSimTotal > 0) {
      rows.push({ label: '内SIMのみ', dcm: c.mnp.uchiSimByGroup.DCM, sbym: c.mnp.uchiSimByGroup.SBYM, auuq: c.mnp.uchiSimByGroup.AUUQ, total: c.mnp.uchiSimTotal, share: '', note: '' });
    }
    if (c.mnp.migTotal > 0) {
      var migParts = MIGRATIONS.filter(function (m) { return agg.migrations[m.key] > 0; }).map(function (m) { return m.text + ' ' + agg.migrations[m.key] + '件'; });
      rows.push({ label: '番号移行', dcm: '', sbym: '', auuq: '', total: c.mnp.migTotal, share: '', note: migParts.join('／') });
    }
    if (c.mnp.sourceTotal > 0) {
      var srcParts = SOURCES.filter(function (s) { return agg.sources[s.key] > 0; }).map(function (s) { return s.text + ' ' + agg.sources[s.key] + '件'; });
      rows.push({ label: '移転元', dcm: '', sbym: '', auuq: '', total: c.mnp.sourceTotal, share: '', note: srcParts.join('／') });
    }

    rows.push({ label: '副商材', dcm: c.sub.byGroup.DCM, sbym: c.sub.byGroup.SBYM, auuq: c.sub.byGroup.AUUQ, total: c.sub.grandTotal, share: '', note: '', header: true });
    if (c.sub.grandTotal > 0) {
      GROUPS.forEach(function (g) {
        if (c.sub.byGroup[g] === 0) return;
        var categories = [];
        SUB_DEF.forEach(function (d) { if (d.group === g && categories.indexOf(d.category) === -1) categories.push(d.category); });
        categories.forEach(function (cat) {
          var defsInCat = SUB_DEF.filter(function (d) { return d.group === g && d.category === cat; });
          var hasAny = defsInCat.some(function (d) { return c.sub.values[d.key] > 0; });
          if (!hasAny) return;
          var parts = defsInCat.filter(function (d) { return c.sub.values[d.key] > 0; }).map(function (d) { return d.name + ' ' + c.sub.values[d.key] + '件'; });
          rows.push({ label: '　' + GROUP_LABEL[g] + cat, dcm: '', sbym: '', auuq: '', total: '', share: '', note: parts.join('／') });
        });
      });
    }

    rows.push({ label: '屋号商材', dcm: c.yagou.byGroup.DCM, sbym: c.yagou.byGroup.SBYM, auuq: c.yagou.byGroup.AUUQ, total: c.yagou.grandTotal, share: '', note: '', header: true });
    if (c.yagou.grandTotal > 0) {
      YAGOU_PRODUCTS.forEach(function (p) {
        var pd = c.yagou.byProduct[p.key];
        if (pd.total === 0) return;
        rows.push({ label: '　' + p.label, dcm: pd.perGroup.DCM, sbym: pd.perGroup.SBYM, auuq: pd.perGroup.AUUQ, total: pd.total, share: '', note: '' });
      });
    }

    if (agg.event) rows.push({ label: 'イベント', dcm: '', sbym: '', auuq: '', total: '', share: '', note: agg.event });

    return rows;
  }

  // ============================================================
  // 公開インターフェース
  // ============================================================
  return {
    CARRIERS: CARRIERS,
    CARRIER_GROUP: CARRIER_GROUP,
    GROUPS: GROUPS,
    GROUP_LABEL: GROUP_LABEL,
    GROUP_COLOR: GROUP_COLOR,
    MIGRATIONS: MIGRATIONS,
    SOURCES: SOURCES,
    SUB_DEF: SUB_DEF,
    YAGOU_PRODUCTS: YAGOU_PRODUCTS,
    yagouKey: yagouKey,
    yagouLabel: yagouLabel,
    YAGOU_GROUP_ORDER: YAGOU_GROUP_ORDER,
    YAGOU_GROUP_LABEL: YAGOU_GROUP_LABEL,

    getLogHeaderDefs: getLogHeaderDefs,
    buildHeaderRows: buildHeaderRows,
    rowFromRaw: rowFromRaw,
    rawFromRow: rawFromRow,

    clampNonNegInt: clampNonNegInt,
    toDateStr: toDateStr,
    parseDateStr: parseDateStr,
    formatDateStr: formatDateStr,
    getWeekdayJa: getWeekdayJa,
    getMonday: getMonday,
    getWeekEnd: getWeekEnd,
    addDays: addDays,
    getWeeksOfMonth: getWeeksOfMonth,
    getWeekIndexInMonth: getWeekIndexInMonth,

    buildDayModel: buildDayModel,
    computeMbSales: computeMbSales,
    computeMnp: computeMnp,
    computeSubProducts: computeSubProducts,
    computeYagouTotals: computeYagouTotals,
    computeShare: computeShare,
    computeAll: computeAll,
    aggregateDayModels: aggregateDayModels,

    getColMap: getColMap,
    validateMergeRanges: validateMergeRanges,
    buildMergeRanges: buildMergeRanges,

    buildCsv: buildCsv,
    csvEscapeCell: csvEscapeCell,
    escapeHtml: escapeHtml,

    fmtPct: fmtPct,
    shareLine: shareLine,
    buildAggBlockRows: buildAggBlockRows,
    buildTitle: buildTitle,
    shortDate: shortDate,
    generateReportText: generateReportText
  };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = Logic;
}

// ============================================================
// sheets.js
// ============================================================
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

// ============================================================
// webapp.js
// ============================================================
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

// ============================================================
// report.js
// ============================================================
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


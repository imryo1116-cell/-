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

/**
 * build.js
 * ------------------------------------------------------------
 * src/*.js を1つの dist/コード.gs に結合し、src/フォーム.html と
 * appsscript.json を dist/ にコピーする。
 * GAS は import/export が使えないため、単純結合だけを行う（ビルドツール不要）。
 *
 * 使い方: node build.js
 * ------------------------------------------------------------
 */
var fs = require('fs');
var path = require('path');

var SRC_DIR = path.join(__dirname, 'src');
var DIST_DIR = path.join(__dirname, 'dist');

// 結合する順番（logic.js を必ず先頭にする＝Logicオブジェクトの定義が最初に来るように）
var JS_FILES = ['logic.js', 'sheets.js', 'webapp.js', 'report.js'];

var HEADER_COMMENT =
  '/**\n' +
  ' * コード.gs（build.js による自動結合ファイル）\n' +
  ' * ------------------------------------------------------------\n' +
  ' * このファイルは src/logic.js → src/sheets.js → src/webapp.js → src/report.js\n' +
  ' * を結合したものです。直接このファイルを編集しないでください。\n' +
  ' * 修正するときは src/ 配下の該当ファイルを直してから `node build.js` を\n' +
  ' * 実行し、生成しなおしたこのファイルをスクリプトエディタに貼り直してください。\n' +
  ' *\n' +
  ' * 【項目を追加・変更したいときにどこを直すか】\n' +
  ' *   - 副商材の項目を増やす   → 下の Logic.SUB_DEF（src/logic.js）に1行追加\n' +
  ' *   - 屋号商材を増やす       → Logic.YAGOU_PRODUCTS（src/logic.js）に1行追加\n' +
  ' *   - MNP移転元の選択肢を増やす → Logic.SOURCES（src/logic.js）に1行追加\n' +
  ' *   - キャリア構成を変える（通常は不要） → Logic.CARRIERS / CARRIER_GROUP\n' +
  ' *   これらはすべて「見出し名（label）」から列位置を解決する設計なので、\n' +
  ' *   配列の途中に追加しても既存のシート列や集計式は壊れません。\n' +
  ' *   追加後は「初期セットアップ」を再実行してください（既存データは消えません）。\n' +
  ' * ------------------------------------------------------------\n' +
  ' */\n\n';

function build() {
  if (!fs.existsSync(DIST_DIR)) fs.mkdirSync(DIST_DIR, { recursive: true });

  var parts = [HEADER_COMMENT];
  JS_FILES.forEach(function (file) {
    var content = fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
    parts.push('// ============================================================\n');
    parts.push('// ' + file + '\n');
    parts.push('// ============================================================\n');
    parts.push(content.trim());
    parts.push('\n\n');
  });

  fs.writeFileSync(path.join(DIST_DIR, 'コード.gs'), parts.join(''));
  fs.copyFileSync(path.join(SRC_DIR, 'フォーム.html'), path.join(DIST_DIR, 'フォーム.html'));
  fs.copyFileSync(path.join(__dirname, 'appsscript.json'), path.join(DIST_DIR, 'appsscript.json'));

  console.log('build.js: dist/コード.gs / dist/フォーム.html / dist/appsscript.json を生成しました。');
}

build();

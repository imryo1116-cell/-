/**
 * gasMock.js
 * ------------------------------------------------------------
 * SpreadsheetApp / LockService / Utilities の軽量モック。
 * sheets.js を実機(GAS)に上げなくても node で動かすためのテスト専用コード。
 * 本番コード(src/*.js)には一切影響しない。
 * ------------------------------------------------------------
 */
var vm = require('vm');
var fs = require('fs');
var path = require('path');

function MockRange(sheet, row, col, numRows, numCols) {
  this.sheet = sheet;
  this.row = row;
  this.col = col;
  this.numRows = numRows;
  this.numCols = numCols;
}

MockRange.prototype.getValues = function () {
  var out = [];
  for (var r = 0; r < this.numRows; r++) {
    var rowArr = [];
    for (var c = 0; c < this.numCols; c++) {
      rowArr.push(this.sheet.getCell(this.row + r, this.col + c));
    }
    out.push(rowArr);
  }
  return out;
};

MockRange.prototype.getValue = function () {
  return this.sheet.getCell(this.row, this.col);
};

MockRange.prototype.setValues = function (values) {
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      this.sheet.setCell(this.row + r, this.col + c, values[r][c]);
    }
  }
  return this;
};

MockRange.prototype.setValue = function (v) {
  this.sheet.setCell(this.row, this.col, v);
  return this;
};

['setNumberFormat', 'setFontWeight', 'setFontColor', 'setBackground',
  'setHorizontalAlignment', 'setWrap'].forEach(function (name) {
  MockRange.prototype[name] = function () { return this; };
});

MockRange.prototype.merge = function () {
  this.sheet.merges.push({ row: this.row, col: this.col, numRows: this.numRows, numCols: this.numCols });
  return this;
};

MockRange.prototype.breakApart = function () {
  var r = this;
  this.sheet.merges = this.sheet.merges.filter(function (m) {
    return !(m.row >= r.row && m.row < r.row + r.numRows && m.col >= r.col && m.col < r.col + r.numCols);
  });
  return this;
};

MockRange.prototype.sort = function (opt) {
  var colIndex = opt.column - this.col; // 0-based, レンジ内の相対列
  var rows = this.getValues();
  rows.sort(function (a, b) {
    if (a[colIndex] < b[colIndex]) return opt.ascending ? -1 : 1;
    if (a[colIndex] > b[colIndex]) return opt.ascending ? 1 : -1;
    return 0;
  });
  this.setValues(rows);
  return this;
};

function MockSheet(name) {
  this.name = name;
  this.data = []; // data[row-1][col-1]
  this.maxRows = 1000;
  this.maxCols = 26;
  this.frozenRows = 0;
  this.frozenCols = 0;
  this.merges = [];
  this.charts = [];
}

MockSheet.prototype.getCell = function (row, col) {
  var r = this.data[row - 1];
  if (!r) return '';
  var v = r[col - 1];
  return v === undefined || v === null ? '' : v;
};

MockSheet.prototype.setCell = function (row, col, value) {
  if (!this.data[row - 1]) this.data[row - 1] = [];
  this.data[row - 1][col - 1] = value;
};

MockSheet.prototype.getRange = function (row, col, numRows, numCols) {
  return new MockRange(this, row, col, numRows || 1, numCols || 1);
};

MockSheet.prototype.getMaxRows = function () { return this.maxRows; };
MockSheet.prototype.getMaxColumns = function () { return this.maxCols; };
MockSheet.prototype.insertRowsAfter = function (after, count) { this.maxRows += count; };
MockSheet.prototype.insertColumnsAfter = function (after, count) { this.maxCols += count; };

MockSheet.prototype.getLastRow = function () {
  for (var r = this.data.length - 1; r >= 0; r--) {
    if (this.data[r] && this.data[r].some(function (v) { return v !== undefined && v !== null && v !== ''; })) return r + 1;
  }
  return 0;
};

MockSheet.prototype.getLastColumn = function () {
  var last = 0;
  this.data.forEach(function (row) {
    if (!row) return;
    for (var c = row.length - 1; c >= 0; c--) {
      if (row[c] !== undefined && row[c] !== null && row[c] !== '') { last = Math.max(last, c + 1); break; }
    }
  });
  return last;
};

MockSheet.prototype.clear = function () { this.data = []; this.merges = []; return this; };
MockSheet.prototype.setFrozenRows = function (n) { this.frozenRows = n; return this; };
MockSheet.prototype.setFrozenColumns = function (n) { this.frozenCols = n; return this; };
MockSheet.prototype.setColumnWidths = function () { return this; };
MockSheet.prototype.setRowHeights = function () { return this; };
MockSheet.prototype.getCharts = function () { return this.charts.slice(); };
MockSheet.prototype.removeChart = function (c) { this.charts = this.charts.filter(function (x) { return x !== c; }); };
MockSheet.prototype.getSheetId = function () { return 1; };

function MockChartBuilder() {
  this._opts = {};
}
['asLineChart', 'asPieChart', 'addRange'].forEach(function (name) {
  MockChartBuilder.prototype[name] = function () { return this; };
});
MockChartBuilder.prototype.setOption = function (k, v) { this._opts[k] = v; return this; };
MockChartBuilder.prototype.setPosition = function () { return this; };
MockChartBuilder.prototype.build = function () { return { opts: this._opts }; };

MockSheet.prototype.newChart = function () { return new MockChartBuilder(); };
MockSheet.prototype.insertChart = function (chart) { this.charts.push(chart); };

function MockSpreadsheet() {
  this.sheets = {};
}
MockSpreadsheet.prototype.getSheetByName = function (name) { return this.sheets[name] || null; };
MockSpreadsheet.prototype.insertSheet = function (name) {
  var sh = new MockSheet(name);
  this.sheets[name] = sh;
  return sh;
};
MockSpreadsheet.prototype.setSpreadsheetTimeZone = function () {};
MockSpreadsheet.prototype.getUrl = function () { return 'https://example.invalid/spreadsheet'; };
MockSpreadsheet.prototype.toast = function () {};

function pad2(n) { return (n < 10 ? '0' : '') + n; }

function formatDateMock(date, tz, pattern) {
  var y = date.getFullYear(), mo = date.getMonth() + 1, d = date.getDate();
  var h = date.getHours(), mi = date.getMinutes(), s = date.getSeconds();
  if (pattern === 'yyyy-MM-dd HH:mm:ss') {
    return y + '-' + pad2(mo) + '-' + pad2(d) + ' ' + pad2(h) + ':' + pad2(mi) + ':' + pad2(s);
  }
  if (pattern === 'yyyy-MM') {
    return y + '-' + pad2(mo);
  }
  return date.toISOString();
}

// src/logic.js + src/sheets.js を1つの vm コンテキストで実行し、
// GAS実行環境に近い形（トップレベル関数がグローバルに生える）で呼び出せるようにする。
function createGasContext() {
  var ss = new MockSpreadsheet();
  var sandbox = {
    console: console,
    SpreadsheetApp: {
      getActiveSpreadsheet: function () { return ss; },
      getUi: function () {
        return {
          createMenu: function () { return this; },
          addItem: function () { return this; },
          addToUi: function () { return this; }
        };
      }
    },
    LockService: {
      getScriptLock: function () {
        return { waitLock: function () {}, releaseLock: function () {} };
      }
    },
    Utilities: {
      formatDate: formatDateMock
    }
  };
  vm.createContext(sandbox);

  var logicSrc = fs.readFileSync(path.join(__dirname, '../src/logic.js'), 'utf8');
  var sheetsSrc = fs.readFileSync(path.join(__dirname, '../src/sheets.js'), 'utf8');
  vm.runInContext(logicSrc, sandbox, { filename: 'logic.js' });
  vm.runInContext(sheetsSrc, sandbox, { filename: 'sheets.js' });

  return { context: sandbox, ss: ss };
}

module.exports = { createGasContext: createGasContext };

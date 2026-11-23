/**
 * GitHub 上の history.json を読み、アプリごとの列に古い履歴から順に並べる。
 *
 * セットアップ:
 * 1. 対象スプレッドシートで 拡張機能 → Apps Script を開き、このファイルの内容を貼る
 * 2. プロジェクトの設定 → タイムゾーン を「東京」にする（朝6時トリガーのため）
 * 3. 初回は syncFromGitHub を実行し、外部接続（raw.githubusercontent.com）を許可する
 * 4. createMorningTrigger を一度実行する（毎日 06:00 前後に syncFromGitHub）
 *
 * 書き込みタイミング:
 *   GitHub Actions が 00:00 JST 前後に history.json を更新し、このスクリプトは 06:00 JST 前後にそれを読む。
 *   アプリごとに既存データの空行を詰め、最終行のすぐ下に新しい履歴を古い順に足す。
 *   データが空のアプリは、GitHub に保存されている最初の履歴から取り込む。
 *   iOS / デスクトップは Ver の変化だけで記入。Android は Ver またはアプデ文面の変化で記入。
 *   既存データを詰めるときは反映済み・記事URLも一緒に移動する。新規行の入力は人が行う。
 *
 * 既存の週次表（シート1 など）は触らない。無いときは「自動取得」シートを作る。
 */

var SHEET_NAME = "自動取得";
var GITHUB_OWNER = "sava110";
var GITHUB_REPO = "GetInfo-YYApplication";
var GITHUB_REF = "my-project";
var GITHUB_PATH = "data/history.json";

var COLS_PER_APP = 5;
var HEADER_ROWS = 2;
var DATA_ROW = 3;

var APPS = [
  { label: "YY文字起こし iOS", id: "com.yysystem.YYSimpleTranscript" },
  { label: "YY文字起こし Android", id: "yytranscript.yysystem.android", android: true },
  { label: "YYProbe iOS", id: "com.yysystem.YYProbe-Lite" },
  { label: "YYProbe Android", id: "yyprobe.yysystem.android", android: true },
  { label: "YYレセプション", id: "com.yysystem.YYReceptionWindow" },
  { label: "YYデスクトップ字幕", id: "yysystem.yydesktopcaption" },
  { label: "YYサイレンアラート", id: "YYSystem.YYEmergencySound" },
  { label: "YY雰囲気カメラ", id: "YYSystem.YYAtmosphereCamera" },
  { label: "YY発音分析", id: "com.yysystem.YYSpeechAssessment" },
  { label: "YYCommCards", id: "com.yysystem.YYCommCards" },
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("更新履歴")
    .addItem("GitHub から取り込む", "syncFromGitHub")
    .addItem("毎日6時のトリガーを登録", "createMorningTrigger")
    .addToUi();
}

function createMorningTrigger() {
  var handlers = ["syncFromGitHub"];
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (handlers.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("syncFromGitHub").timeBased().atHour(6).everyDays(1).create();
}

function createNoonTrigger() {
  createMorningTrigger();
}

function syncFromGitHub() {
  var document = fetchHistoryJson_();
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sheet = getOrCreateSheet_();
    ensureHeaders_(sheet);
    syncAppHistories_(sheet, document);
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

function fetchHistoryJson_() {
  // api.github.com は未認証だと IP あたり 60 回/時。Apps Script は Google の共有 IP なので
  // すぐ 403 (rate limit) になる。公開リポジトリは raw を使う。
  var url =
    "https://raw.githubusercontent.com/" +
    GITHUB_OWNER +
    "/" +
    GITHUB_REPO +
    "/" +
    GITHUB_REF +
    "/" +
    GITHUB_PATH;
  var response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "GetInfo-YYApplication-sheets",
    },
  });
  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("GitHub から history.json を読めません (" + code + "): " + response.getContentText());
  }
  return JSON.parse(response.getContentText());
}

function getOrCreateSheet_() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = book.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function ensureHeaders_(sheet) {
  var width = APPS.length * COLS_PER_APP;
  var columns = sheet.getMaxColumns();
  if (columns < width) {
    sheet.insertColumnsAfter(columns, width - columns);
  }
  var names = [];
  var fields = [];
  APPS.forEach(function (app) {
    names.push(app.label, "", "", "", "");
    fields.push("情報取得日時", "アプリVer", "アプデ内容", "反映済み", "記事URL");
  });
  sheet.getRange(1, 1, 1, width).breakApart();
  sheet.getRange(1, 1, 1, width).setValues([names]);
  sheet.getRange(2, 1, 1, width).setValues([fields]);
  sheet.getRange(1, 1, HEADER_ROWS, width).setFontWeight("bold");
  for (var i = 0; i < APPS.length; i++) {
    sheet.getRange(1, i * COLS_PER_APP + 1, 1, COLS_PER_APP).merge();
  }
  sheet.setFrozenRows(HEADER_ROWS);
  var contentCols = [];
  for (var j = 0; j < APPS.length; j++) {
    contentCols.push(j * COLS_PER_APP + 3);
  }
  contentCols.forEach(function (column) {
    sheet.setColumnWidth(column, 280);
  });
}

function syncAppHistories_(sheet, document) {
  var apps = document && document.apps && typeof document.apps === "object" ? document.apps : {};
  var width = APPS.length * COLS_PER_APP;
  var height = sheet.getLastRow() - DATA_ROW + 1;
  var data = height > 0 ? sheet.getRange(DATA_ROW, 1, height, width) : null;
  var values = data ? data.getValues() : [];
  var formulas = data ? data.getFormulas() : [];
  var notes = data ? data.getNotes() : [];
  APPS.forEach(function (app, index) {
    var offset = index * COLS_PER_APP;
    var column = offset + 1;
    var nextRow = DATA_ROW;
    var lastEntry = null;
    values.forEach(function (row, rowIndex) {
      var cells = row.slice(offset, offset + COLS_PER_APP);
      var occupied = cells.some(function (value, cellIndex) {
        var i = offset + cellIndex;
        return value !== "" || formulas[rowIndex][i] !== "" || notes[rowIndex][i] !== "";
      });
      if (!occupied) {
        return;
      }
      var sourceRow = DATA_ROW + rowIndex;
      if (sourceRow !== nextRow) {
        sheet.getRange(sourceRow, column, 1, COLS_PER_APP)
          .moveTo(sheet.getRange(nextRow, column, 1, COLS_PER_APP));
      }
      nextRow++;
      if (cells[1] !== "" || cells[2] !== "") {
        lastEntry = [cells[0], String(cells[1]), String(cells[2])];
      }
    });

    var pending = pendingHistory_(app, apps[app.id] || {}, lastEntry);
    if (!pending.length) {
      return;
    }
    var rows = sheet.getMaxRows();
    var endRow = nextRow + pending.length - 1;
    if (endRow > rows) {
      sheet.insertRowsAfter(rows, endRow - rows);
    }
    sheet.getRange(nextRow, column + 1, pending.length, 1).setNumberFormat("@");
    sheet.getRange(nextRow, column + 2, pending.length, 1).setWrap(true);
    sheet.getRange(nextRow, column, pending.length, 3).setValues(pending);
  });
}

function pendingHistory_(app, record, lastEntry) {
  var history = Array.isArray(record.history) && record.history.length
    ? record.history
    : record.latest ? [record.latest] : [];
  var entries = history.map(function (entry) {
    return [
      formatFetchedAt_(entry.fetchedAt),
      entry.version == null ? "" : String(entry.version),
      entry.recentChanges == null ? "" : String(entry.recentChanges),
    ];
  });
  if (lastEntry) {
    // 日時も照合して、同じバージョンが後日再登場した場合を区別する。
    var stamp = sheetTimestamp_(lastEntry[0]);
    var anchor = entries.length - 1;
    for (; anchor >= 0; anchor--) {
      var entry = entries[anchor];
      if ((!stamp || entry[0].slice(0, stamp.length) === stamp)
          && entry[1] === lastEntry[1] && (!app.android || entry[2] === lastEntry[2])) {
        break;
      }
    }
    if (anchor >= 0) {
      entries = entries.slice(anchor + 1);
    } else if (stamp) {
      entries = entries.filter(function (entry) {
        return entry[0].slice(0, stamp.length) > stamp;
      });
    } else {
      entries = entries.slice(-1);
    }
  }

  var previous = lastEntry || ["", "", ""];
  var pending = [];
  entries.forEach(function (entry) {
    if (appUpdated_(app, previous, entry[1], entry[2])) {
      pending.push(entry);
      previous = entry;
    }
  });
  return pending;
}

function appUpdated_(app, previous, version, changes) {
  if (app.android) {
    return (
      (version !== "" || changes !== "") &&
      (version !== previous[1] || changes !== previous[2])
    );
  }
  return version !== "" && version !== previous[1];
}

function sheetTimestamp_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  }
  var match = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::\d{2})?)?$/.exec(String(value).trim());
  if (!match) {
    return "";
  }
  var date = [match[1], match[2].padStart(2, "0"), match[3].padStart(2, "0")].join("/");
  return match[4] ? date + " " + match[4].padStart(2, "0") + ":" + match[5] : date;
}

function formatFetchedAt_(value) {
  if (!value || typeof value !== "string") {
    return "";
  }
  var parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return value;
  }
  return Utilities.formatDate(parsed, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
}

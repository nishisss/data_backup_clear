/**
 * 保存データ側（このスプレッドシート）に設置するスクリプト
 * - 定義シート（A:SS ID, B:Sheet名, C:処理対象チェック）を読み込み
 * - 対象の作業シートを保存データへスナップショット保存（値のみ / 書式はコピー）
 * - 作業シート側のF:N（5行目以降）をクリア
 */

const CONFIG = {
  DEF_SHEET_NAME: "定義",
  LOG_SHEET_NAME: "ログ",

  DEF_START_ROW: 2,
  DEF_ROW_COUNT: 5,
  DEF_COL_COUNT: 3,

  CLEAR_START_ROW: 5,
  CLEAR_START_COL: 6,
  CLEAR_COLS: 9,

  DATE_TZ: "Asia/Tokyo",
  COPY_MODE: "sheet", // "sheet" or "range"
  COPY_FALLBACK_TO_RANGE: true,
  COPY_COLUMN_WIDTHS: true,
  COPY_ROW_HEIGHTS: false,
  COPY_MERGES: true,
  COPY_BATCH_CELLS: 200000,
  COPY_BATCH_ROWS: 500,
  RETRY_MAX: 2,
  RETRY_BASE_MS: 800,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("保存データ処理")
    .addItem("定義シート初期化", "initDefinitionSheet")
    .addSeparator()
    .addItem("保存実行（確認あり）", "runSaveSnapshotConfirm")
    .addItem("保存実行（確認なし）", "runSaveSnapshot")
    .addToUi();
}

/**
 * 定義シートを作成/整形（ヘッダ、チェックボックス）
 */
function initDefinitionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const def = getOrCreateSheet_(ss, CONFIG.DEF_SHEET_NAME);

  def.getRange(1, 1, 1, 3).setValues([["作業シートID", "コピー対象シート名", "処理対象"]]);

  const range = def.getRange(CONFIG.DEF_START_ROW, 1, CONFIG.DEF_ROW_COUNT, CONFIG.DEF_COL_COUNT);
  range.clearContent();

  const checkboxRange = def.getRange(CONFIG.DEF_START_ROW, 3, CONFIG.DEF_ROW_COUNT, 1);
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  checkboxRange.setDataValidation(rule);

  def.setFrozenRows(1);
  def.autoResizeColumns(1, 3);

  ss.toast("定義シートを初期化しました。A:ID / B:シート名 / C:チェック を設定してください。", "完了", 5);
}

/**
 * メイン処理：スナップショット保存 + 作業シートのF:Nクリア
 */
function runSaveSnapshot() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const defSheet = ss.getSheetByName(CONFIG.DEF_SHEET_NAME);
    if (!defSheet) {
      throw new Error(`定義シート「${CONFIG.DEF_SHEET_NAME}」が見つかりません。先に「定義シート初期化」を実行してください。`);
    }

    const logSheet = getOrCreateSheet_(ss, CONFIG.LOG_SHEET_NAME);
    ensureLogHeader_(logSheet);

    const defValues = defSheet
      .getRange(CONFIG.DEF_START_ROW, 1, CONFIG.DEF_ROW_COUNT, CONFIG.DEF_COL_COUNT)
      .getValues();

    const today = Utilities.formatDate(new Date(), CONFIG.DATE_TZ, "yyyyMMdd");
    const maxSerialByIndex = buildMaxSerialMap_(ss, today); // {1..5}

    let processed = 0;
    let skipped = 0;
    const errors = [];

    // ★ログは配列に貯めて最後に一括書き込み
    const logRows = [];
    const srcSpreadsheetCache = {};

    ss.toast("保存処理を開始します…", "実行中", 5);

    defValues.forEach((row, i) => {
      const defNo = i + 1; // ※1 = 定義の何番目（1〜5）
      const srcSpreadsheetId = (row[0] || "").toString().trim();
      const srcSheetName = (row[1] || "").toString().trim();
      const isTarget = row[2] === true;

      if (!isTarget) return;

      if (!srcSpreadsheetId || !srcSheetName) {
        skipped++;
        logRows.push(makeLogRow_(defNo, srcSpreadsheetId, srcSheetName, "", "SKIP", "作業シートIDまたはシート名が未入力"));
        return;
      }

      try {
        const nextSerial = (maxSerialByIndex[defNo] || 0) + 1;
        if (nextSerial > 99) throw new Error("同日同番号の連番が99を超えました。シート名ルール上これ以上作れません。");
        maxSerialByIndex[defNo] = nextSerial;

        const serial2 = Utilities.formatString("%02d", nextSerial);
        const destSheetName = `${today}_${defNo}_${serial2}`;

        // 作業シートを開く（同一SSは再利用）
        const srcSS = getSpreadsheetByIdCached_(srcSpreadsheetId, srcSpreadsheetCache);
        const srcSheet = srcSS.getSheetByName(srcSheetName);
        if (!srcSheet) throw new Error(`コピー対象シートが見つかりません: ${srcSheetName}`);

        // ★スナップショット作成（タイムアウト時は軽量コピーにフォールバック）
        copySnapshot_(srcSheet, ss, destSheetName);

        // 3) 作業シート側のF:N、5行目以降をクリア
        const lastRow = srcSheet.getLastRow();
        if (lastRow >= CONFIG.CLEAR_START_ROW) {
          const numRows = lastRow - CONFIG.CLEAR_START_ROW + 1;
          srcSheet
            .getRange(CONFIG.CLEAR_START_ROW, CONFIG.CLEAR_START_COL, numRows, CONFIG.CLEAR_COLS)
            .clearContent();
        }

        processed++;
        logRows.push(makeLogRow_(defNo, srcSpreadsheetId, srcSheetName, destSheetName, "OK", "完了"));
      } catch (e) {
        const message = getErrorMessage_(e);
        errors.push(`定義${defNo}: ${message}`);
        logRows.push(makeLogRow_(defNo, srcSpreadsheetId, srcSheetName, "", "ERROR", message));
      }
    });

    // ★ログ一括書き込み
    if (logRows.length) {
      const start = logSheet.getLastRow() + 1;
      logSheet.getRange(start, 1, logRows.length, 7).setValues(logRows);
    }

    const summary = `完了：${processed}件 / スキップ：${skipped}件 / エラー：${errors.length}件`;
    ss.toast(summary, "完了", 8);

    if (errors.length) {
      SpreadsheetApp.getUi().alert(`一部エラーが発生しました。\n\n${summary}\n\n詳細は「${CONFIG.LOG_SHEET_NAME}」シートを確認してください。`);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * コピー済みシートを「値のみ」にする（高速版）
 * - getDataRange() は使わない（無駄に巨大になりがち）
 * - lastRow/lastCol で使用範囲を絞って、サーバー側コピーで値化
 */
function valueOnlyByOverwrite_(sheet) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr <= 0 || lc <= 0) return;

  const range = sheet.getRange(1, 1, lr, lc);
  range.copyTo(range, { contentsOnly: true });
}

/**
 * 当日分（yyyyMMdd）の既存シート名から、※1ごとの最大連番を作る
 * シート名形式：yyyyMMdd_1_00
 */
function buildMaxSerialMap_(ss, today) {
  const map = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  const re = new RegExp(`^${today}_(\\d)_(\\d{2})$`);

  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    const m = name.match(re);
    if (!m) return;
    const idx = parseInt(m[1], 10);
    const serial = parseInt(m[2], 10);
    if (idx >= 1 && idx <= 5) map[idx] = Math.max(map[idx], serial);
  });

  return map;
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureLogHeader_(logSheet) {
  if (logSheet.getLastRow() === 0) {
    logSheet.getRange(1, 1, 1, 7).setValues([[
      "日時",
      "定義No(※1)",
      "作業シートID",
      "作業シート名",
      "作成シート名",
      "結果",
      "メッセージ"
    ]]);
    logSheet.setFrozenRows(1);
    logSheet.autoResizeColumns(1, 7);
  }
}

function makeLogRow_(defNo, srcId, srcSheetName, destName, status, message) {
  const ts = Utilities.formatDate(new Date(), CONFIG.DATE_TZ, "yyyy-MM-dd HH:mm:ss");
  return [ts, defNo, srcId, srcSheetName, destName, status, message];
}

function getSpreadsheetByIdCached_(spreadsheetId, cache) {
  if (cache[spreadsheetId]) return cache[spreadsheetId];
  cache[spreadsheetId] = retryWithBackoff_(
    () => SpreadsheetApp.openById(spreadsheetId),
    CONFIG.RETRY_MAX,
    CONFIG.RETRY_BASE_MS
  );
  return cache[spreadsheetId];
}

function copySnapshot_(srcSheet, destSS, destSheetName) {
  if (CONFIG.COPY_MODE === "range") {
    return copyByRange_(srcSheet, destSS, destSheetName);
  }

  try {
    return copyBySheet_(srcSheet, destSS, destSheetName);
  } catch (e) {
    const message = getErrorMessage_(e);
    if (CONFIG.COPY_FALLBACK_TO_RANGE && isTimeoutError_(message)) {
      return copyByRange_(srcSheet, destSS, destSheetName);
    }
    throw e;
  }
}

function copyBySheet_(srcSheet, destSS, destSheetName) {
  // 1) 書式・列幅・結合なども含めてシート丸ごとコピー
  // 2) “使用範囲だけ”をサーバー側で値化（数式を結果値に）
  const copied = retryWithBackoff_(
    () => srcSheet.copyTo(destSS),
    CONFIG.RETRY_MAX,
    CONFIG.RETRY_BASE_MS
  );
  copied.setName(destSheetName);
  valueOnlyByOverwrite_(copied);
  return copied;
}

function copyByRange_(srcSheet, destSS, destSheetName) {
  const lr = srcSheet.getLastRow();
  const lc = srcSheet.getLastColumn();
  const dest = destSS.insertSheet(destSheetName);
  if (lr <= 0 || lc <= 0) return dest;

  const srcRange = srcSheet.getRange(1, 1, lr, lc);
  const destRange = dest.getRange(1, 1, lr, lc);

  // 書式を先にコピー（軽量）
  srcRange.copyTo(destRange, { formatOnly: true });

  // 値のみをコピー（大量セルは分割して転送）
  copyValuesInBatches_(srcSheet, dest, lr, lc);

  if (CONFIG.COPY_COLUMN_WIDTHS) {
    for (let c = 1; c <= lc; c++) {
      dest.setColumnWidth(c, srcSheet.getColumnWidth(c));
    }
  }
  if (CONFIG.COPY_ROW_HEIGHTS) {
    for (let r = 1; r <= lr; r++) {
      dest.setRowHeight(r, srcSheet.getRowHeight(r));
    }
  }
  if (CONFIG.COPY_MERGES) {
    const merges = srcRange.getMergedRanges();
    merges.forEach(range => {
      dest.getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns()).merge();
    });
  }

  return dest;
}

function copyValuesInBatches_(srcSheet, destSheet, lr, lc) {
  const totalCells = lr * lc;
  if (totalCells <= CONFIG.COPY_BATCH_CELLS) {
    const values = srcSheet.getRange(1, 1, lr, lc).getValues();
    destSheet.getRange(1, 1, lr, lc).setValues(values);
    return;
  }

  const batchRows = CONFIG.COPY_BATCH_ROWS;
  for (let r = 1; r <= lr; r += batchRows) {
    const numRows = Math.min(batchRows, lr - r + 1);
    const values = srcSheet.getRange(r, 1, numRows, lc).getValues();
    destSheet.getRange(r, 1, numRows, lc).setValues(values);
  }
}

function retryWithBackoff_(fn, maxRetries, baseMs) {
  let attempt = 0;
  while (true) {
    try {
      return fn();
    } catch (e) {
      const message = getErrorMessage_(e);
      if (attempt >= maxRetries || !isRetryableError_(message)) throw e;
      Utilities.sleep(baseMs * Math.pow(2, attempt));
      attempt++;
    }
  }
}

function isRetryableError_(message) {
  return /timed out|タイムアウト|Service invoked too many times|Rate Limit|internal error/i.test(message);
}

function isTimeoutError_(message) {
  return /timed out|タイムアウト/i.test(message);
}

function getErrorMessage_(e) {
  return e && e.message ? e.message : String(e);
}

function runSaveSnapshotConfirm() {
  const ui = SpreadsheetApp.getUi();

  const msg =
    "保存処理を実行します。\n\n" +
    "【重要】処理対象の作業シート側で、F列〜N列の5行目以降の値をクリアします。\n" +
    "この操作は元に戻せません。\n\n" +
    "実行してよろしいですか？";

  const res = ui.alert("最終確認", msg, ui.ButtonSet.OK_CANCEL);

  if (res !== ui.Button.OK) {
    SpreadsheetApp.getActiveSpreadsheet().toast("キャンセルしました。", "中止", 3);
    return;
  }

  runSaveSnapshot();
}

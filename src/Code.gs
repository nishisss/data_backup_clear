/**
 * 保存データ側（このスプレッドシート）に設置するスクリプト
 * - 定義シート（A:SS ID, B:クリア対象シート名, C:処理対象チェック, D:保存先フォルダID）を読み込み
 * - 対象の作業スプレッドシートをGoogle Driveへファイルごとバックアップ
 * - 作業シート側のF:N（5行目以降）をクリア
 */

const CONFIG = {
  DEF_SHEET_NAME: "定義",
  LOG_SHEET_NAME: "ログ",

  DEF_START_ROW: 2,
  DEF_ROW_COUNT: 5,
  DEF_COL_COUNT: 4,

  CLEAR_START_ROW: 5,
  CLEAR_START_COL: 6,
  CLEAR_COLS: 9,

  DATE_TZ: "Asia/Tokyo",
  FILE_NAME_DATETIME_FORMAT: "yyyyMMdd_HHmmss",
  SHOW_SUCCESS_ALERT: true,
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

  def.getRange(1, 1, 1, 4).setValues([[
    "作業シートID",
    "クリア対象シート名",
    "処理対象",
    "保存先フォルダID",
  ]]);

  const range = def.getRange(CONFIG.DEF_START_ROW, 1, CONFIG.DEF_ROW_COUNT, CONFIG.DEF_COL_COUNT);
  range.clearContent();

  const checkboxRange = def.getRange(CONFIG.DEF_START_ROW, 3, CONFIG.DEF_ROW_COUNT, 1);
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  checkboxRange.setDataValidation(rule);

  def.setFrozenRows(1);
  def.autoResizeColumns(1, 4);

  ss.toast(
    "定義シートを初期化しました。A:ID / B:シート名 / C:チェック / D:保存先フォルダID を設定してください。",
    "完了",
    5
  );
}

/**
 * メイン処理：ファイルバックアップ + 作業シートのF:Nクリア
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

    let processed = 0;
    let skipped = 0;
    const errors = [];

    const logRows = [];
    const srcSpreadsheetCache = {};
    const folderCache = {};

    ss.toast("保存処理を開始します…", "実行中", 5);

    defValues.forEach((row, i) => {
      const defNo = i + 1;
      const srcSpreadsheetId = (row[0] || "").toString().trim();
      const srcSheetName = (row[1] || "").toString().trim();
      const isTarget = row[2] === true;
      const backupFolderId = (row[3] || "").toString().trim();

      if (!isTarget) return;

      if (!srcSpreadsheetId || !srcSheetName || !backupFolderId) {
        skipped++;
        logRows.push(
          makeLogRow_(
            defNo,
            srcSpreadsheetId,
            srcSheetName,
            "",
            "SKIP",
            "作業シートID・シート名・保存先フォルダIDのいずれかが未入力"
          )
        );
        return;
      }

      try {
        const srcSS = getSpreadsheetByIdCached_(srcSpreadsheetId, srcSpreadsheetCache);
        const srcSheet = srcSS.getSheetByName(srcSheetName);
        if (!srcSheet) throw new Error(`クリア対象シートが見つかりません: ${srcSheetName}`);

        const backupFolder = getFolderByIdCached_(backupFolderId, folderCache);
        const backupFileName = backupSpreadsheetFile_(srcSpreadsheetId, backupFolder);

        const clearedRows = clearWorkRange_(srcSheet);
        const resultMessage = clearedRows > 0 ? `完了（F:Nを${clearedRows}行クリア）` : "完了（クリア対象なし）";

        processed++;
        logRows.push(makeLogRow_(defNo, srcSpreadsheetId, srcSheetName, backupFileName, "OK", resultMessage));
      } catch (e) {
        const message = getErrorMessage_(e);
        errors.push(`定義${defNo}: ${message}`);
        logRows.push(makeLogRow_(defNo, srcSpreadsheetId, srcSheetName, "", "ERROR", message));
      }
    });

    if (logRows.length) {
      const start = logSheet.getLastRow() + 1;
      logSheet.getRange(start, 1, logRows.length, 7).setValues(logRows);
    }

    const summary = `完了：${processed}件 / スキップ：${skipped}件 / エラー：${errors.length}件`;
    ss.toast(summary, "完了", 8);

    if (errors.length) {
      SpreadsheetApp.getUi().alert(
        `一部エラーが発生しました。\n\n${summary}\n\n詳細は「${CONFIG.LOG_SHEET_NAME}」シートを確認してください。`
      );
    } else if (CONFIG.SHOW_SUCCESS_ALERT) {
      SpreadsheetApp.getUi().alert(`保存処理が完了しました。\n\n${summary}`);
    }
  } finally {
    lock.releaseLock();
  }
}

function clearWorkRange_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.CLEAR_START_ROW) return 0;

  const numRows = lastRow - CONFIG.CLEAR_START_ROW + 1;
  sheet
    .getRange(CONFIG.CLEAR_START_ROW, CONFIG.CLEAR_START_COL, numRows, CONFIG.CLEAR_COLS)
    .clearContent();
  SpreadsheetApp.flush();
  return numRows;
}

function backupSpreadsheetFile_(spreadsheetId, folder) {
  const file = retryWithBackoff_(
    () => DriveApp.getFileById(spreadsheetId),
    CONFIG.RETRY_MAX,
    CONFIG.RETRY_BASE_MS
  );
  const backupFileName = makeBackupFileName_(file.getName());

  retryWithBackoff_(
    () => file.makeCopy(backupFileName, folder),
    CONFIG.RETRY_MAX,
    CONFIG.RETRY_BASE_MS
  );

  return backupFileName;
}

function makeBackupFileName_(baseName) {
  const timestamp = Utilities.formatDate(new Date(), CONFIG.DATE_TZ, CONFIG.FILE_NAME_DATETIME_FORMAT);
  return `${baseName}_${timestamp}`;
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
      "作成ファイル名",
      "結果",
      "メッセージ",
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

function getFolderByIdCached_(folderId, cache) {
  if (cache[folderId]) return cache[folderId];
  cache[folderId] = retryWithBackoff_(
    () => DriveApp.getFolderById(folderId),
    CONFIG.RETRY_MAX,
    CONFIG.RETRY_BASE_MS
  );
  return cache[folderId];
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

function getErrorMessage_(e) {
  return e && e.message ? e.message : String(e);
}

function runSaveSnapshotConfirm() {
  const ui = SpreadsheetApp.getUi();

  const msg =
    "保存処理を実行します。\n\n" +
    "【重要】指定フォルダへスプレッドシート全体のバックアップを作成した後、作業シート側でF列〜N列の5行目以降の値をクリアします。\n" +
    "この操作は元に戻せません。\n\n" +
    "実行してよろしいですか？";

  const res = ui.alert("最終確認", msg, ui.ButtonSet.OK_CANCEL);

  if (res !== ui.Button.OK) {
    SpreadsheetApp.getActiveSpreadsheet().toast("キャンセルしました。", "中止", 3);
    return;
  }

  runSaveSnapshot();
}

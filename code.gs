/**
 * 共有フォルダ同期くん - Server Side Logic (コンテナバインド版)
 */

// --- 初期設定・定数（日本語化） ---
const SHEET_NAME_LOGS = '転送ログ';
const SHEET_NAME_CONFIG = '設定';
const APP_TITLE = '共有フォルダ同期くん';

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=148q8_lJ6rxjyzMTenibABY2zM-nILoAH&.png');
}

// --- データベース(スプレッドシート)連携 ---

/**
 * データベースとシートを初期化して取得
 * コンテナバインド版なので、getActiveSpreadsheet() を直接使用します。
 */
function getOrCreateDb() {
  // コンテナバインドされているスプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss) {
    throw new Error('スプレッドシートが見つかりません。このスクリプトはスプレッドシートの「拡張機能」から実行してください。');
  }

  // --- 設定シート (設定) ---
  let configSheet = ss.getSheetByName(SHEET_NAME_CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_NAME_CONFIG);
    // 初期作成時はヘッダー等が空なので作成する
  }
  
  // ヘッダー等の初期化チェック（行数が1未満なら作成）
  if (configSheet.getLastRow() < 1) {
    configSheet.clear(); 
    // 日本語ヘッダー
    configSheet.appendRow(['項目キー', '設定値']); 
    
    // 初期値設定
    configSheet.appendRow(['folderPairs', '[]']); 
    configSheet.appendRow(['syncFrequency', 'hourly']);
    
    // 見栄えの調整
    configSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e5e7eb');
    configSheet.setColumnWidth(1, 150);
    configSheet.setColumnWidth(2, 400);
  }

  // --- ログシート (転送ログ) ---
  let logSheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAME_LOGS);
  }

  // ヘッダー等の初期化チェック
  if (logSheet.getLastRow() < 1) {
    logSheet.clear();
    // 日本語ヘッダー
    logSheet.appendRow(['日時', '種類', 'メッセージ', 'ファイル名']);
    
    // 見栄えの調整
    logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e5e7eb');
    logSheet.setColumnWidth(1, 160); // 日時
    logSheet.setColumnWidth(3, 350); // メッセージ
  }

  return { ss, configSheet, logSheet };
}

// --- API (クライアントとの通信) ---

function getAppConfig() {
  const { configSheet } = getOrCreateDb();
  const data = configSheet.getDataRange().getValues();
  const config = {};
  
  // 1行目はヘッダーなのでスキップ
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) config[data[i][0]] = data[i][1];
  }

  // 互換性チェック: folderPairsがない場合
  if (!config.folderPairs) config.folderPairs = '[]';
  
  // JSONパース
  let folderPairs = [];
  try {
    folderPairs = JSON.parse(config.folderPairs);
  } catch (e) {
    folderPairs = [];
  }
  // 空の場合は初期フォーム用に1つ空オブジェクトを入れるなどの処理はフロントに任せるかここでやるか
  // ここではそのまま返す

  // デフォルト値
  if (!config.syncFrequency) config.syncFrequency = 'hourly';

  const triggers = ScriptApp.getProjectTriggers();
  const isAutoSyncEnabled = triggers.some(t => t.getHandlerFunction() === 'runSyncProcess');

  return { 
    folderPairs: folderPairs, 
    syncFrequency: config.syncFrequency,
    isAutoSyncEnabled: isAutoSyncEnabled 
  };
}

function saveAppConfig(newConfig) {
  const { configSheet } = getOrCreateDb();
  
  // 現在の設定をマップ化
  const currentData = configSheet.getDataRange().getValues();
  const currentMap = {};
  for(let i=1; i<currentData.length; i++) {
    currentMap[currentData[i][0]] = currentData[i][1];
  }

  // 更新内容を反映
  if (newConfig.folderPairs) {
    currentMap['folderPairs'] = JSON.stringify(newConfig.folderPairs);
  }
  if (newConfig.syncFrequency) {
    currentMap['syncFrequency'] = newConfig.syncFrequency;
  }
  
  // シートをクリアして再書き込み
  configSheet.clearContents();
  configSheet.appendRow(['項目キー', '設定値']); // 日本語ヘッダー
  configSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e5e7eb');
  
  // 保存
  Object.keys(currentMap).forEach(key => {
    configSheet.appendRow([key, currentMap[key]]);
  });

  return { success: true };
}

function toggleAutoSync(enable, frequency = 'hourly') {
  const functionName = 'runSyncProcess';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if (t.getHandlerFunction() === functionName) ScriptApp.deleteTrigger(t); });

  // 頻度設定も保存
  const { configSheet } = getOrCreateDb();
  const data = configSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'syncFrequency') {
      configSheet.getRange(i + 1, 2).setValue(frequency);
      found = true;
      break;
    }
  }
  if (!found) configSheet.appendRow(['syncFrequency', frequency]);

  if (enable) {
    let builder = ScriptApp.newTrigger(functionName).timeBased();
    if (frequency === 'daily') builder.everyDays(1).atHour(0);
    else if (frequency === 'weekly') builder.everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8);
    else builder.everyHours(1);
    
    builder.create();
    return { enabled: true, frequency: frequency, message: '自動同期をONにしました' };
  } else {
    return { enabled: false, frequency: frequency, message: '自動同期をOFFにしました' };
  }
}

function getSyncLogs() {
  const { logSheet } = getOrCreateDb();
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return [];
  const numRows = Math.min(lastRow - 1, 50);
  const startRow = Math.max(2, lastRow - numRows + 1);
  const data = logSheet.getRange(startRow, 1, numRows, 4).getValues();
  return data.reverse().map(row => ({
    date: new Date(row[0]).toISOString(),
    type: row[1],
    message: row[2],
    fileName: row[3]
  }));
}

/**
 * 【Core】一括同期処理
 */
function runSyncProcess() {
  const { configSheet, logSheet } = getOrCreateDb();
  
  // 設定読み込み
  const configData = configSheet.getDataRange().getValues();
  let folderPairs = [];
  
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'folderPairs') {
      try {
        folderPairs = JSON.parse(configData[i][1]);
      } catch(e) { folderPairs = []; }
    }
  }

  if (!folderPairs || folderPairs.length === 0) {
    return { success: false, message: '転送設定が見つかりません。' };
  }

  const logs = [];
  const timestamp = new Date();
  let totalProcessed = 0;
  let totalErrors = 0;

  // --- ペアごとの処理ループ ---
  folderPairs.forEach((pair, index) => {
    const sourceId = pair.source;
    const targetId = pair.target;
    
    if (!sourceId || !targetId) return;

    try {
      const sourceFolder = DriveApp.getFolderById(sourceId);
      const targetFolder = DriveApp.getFolderById(targetId);
      const folderName = sourceFolder.getName();
      
      const targetFilesMap = {};
      const targetFiles = targetFolder.getFiles();
      while (targetFiles.hasNext()) {
        const file = targetFiles.next();
        targetFilesMap[file.getName()] = file;
      }

      const sourceFiles = sourceFolder.getFiles();
      
      while (sourceFiles.hasNext()) {
        const sFile = sourceFiles.next();
        const sName = sFile.getName();

        try {
          const tFile = targetFilesMap[sName];

          if (tFile) {
            // 更新判定
            if (sFile.getLastUpdated().getTime() > tFile.getLastUpdated().getTime()) {
              tFile.setTrashed(true);
              const newFile = sFile.makeCopy(sName, targetFolder);
              newFile.moveTo(targetFolder); // マイドライブ対策

              logs.push([timestamp, 'UPDATE', `[${folderName}] 更新完了`, sName]);
              totalProcessed++;
            }
          } else {
            // 新規作成
            const newFile = sFile.makeCopy(sName, targetFolder);
            newFile.moveTo(targetFolder); // マイドライブ対策

            logs.push([timestamp, 'CREATE', `[${folderName}] 新規作成`, sName]);
            totalProcessed++;
          }
        } catch (fileError) {
          Logger.log(`File Error (${sName}): ${fileError.toString()}`);
          logs.push([timestamp, 'ERROR', `[${folderName}] 失敗: ${fileError.message}`, sName]);
          totalErrors++;
        }
      }
    } catch (folderError) {
      Logger.log(`Folder Pair ${index + 1} Error: ${folderError.toString()}`);
      logs.push([timestamp, 'ERROR', `設定${index + 1}のフォルダエラー（IDを確認してください）`, '-']);
      totalErrors++;
    }
  });
  
  if (logs.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logs.length, 4).setValues(logs);
  }

  const msg = `${totalProcessed}件のファイルを処理しました。(エラー: ${totalErrors}件)`;
  return { success: true, message: msg };
}

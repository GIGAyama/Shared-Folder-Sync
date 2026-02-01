/**
 * 共有フォルダ同期くん - Server Side Logic (Ver 2.0)
 * * テストモード、進捗確認、名前付け機能を追加した強化版です。
 */

// --- 1. 定数・設定 ---
const APP_TITLE = '共有フォルダ同期くん';
const SHEET_NAME_CONFIG = '設定';
const SHEET_NAME_LOGS = '転送ログ';
const KEY_DB_SS_ID = 'DB_SS_ID';
const CACHE_KEY_PROGRESS = 'SYNC_PROGRESS'; // 進捗状況を保存するキー

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=148q8_lJ6rxjyzMTenibABY2zM-nILoAH&.png');
}

// --- 2. データベース管理 (変更なし) ---
function getOrCreateDb() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty(KEY_DB_SS_ID);
  let ss;

  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ssId = null; }
  }

  if (!ssId) {
    ss = SpreadsheetApp.create(APP_TITLE + '_データベース');
    ssId = ss.getId();
    props.setProperty(KEY_DB_SS_ID, ssId);
  }

  let configSheet = ss.getSheetByName(SHEET_NAME_CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_NAME_CONFIG);
    const defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);
  }
  
  if (configSheet.getLastRow() < 1) {
    configSheet.clear(); 
    configSheet.appendRow(['項目キー', '設定値']); 
    configSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e9ecef');
    configSheet.appendRow(['folderPairs', '[]']); 
    configSheet.appendRow(['syncFrequency', 'hourly']);
  }

  let logSheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAME_LOGS);
  }

  if (logSheet.getLastRow() < 1) {
    logSheet.clear();
    logSheet.appendRow(['日時', '種類', 'メッセージ', 'ファイル名']);
    logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e9ecef');
    logSheet.setColumnWidth(1, 160);
    logSheet.setColumnWidth(3, 300);
  }

  return { ss, configSheet, logSheet };
}

// --- 3. API (機能追加) ---

function getAppConfig() {
  const { configSheet } = getOrCreateDb();
  const data = configSheet.getDataRange().getValues();
  const config = {};
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) config[data[i][0]] = data[i][1];
  }

  let folderPairs = [];
  try {
    folderPairs = config.folderPairs ? JSON.parse(config.folderPairs) : [];
  } catch (e) {
    folderPairs = [];
  }

  const triggers = ScriptApp.getProjectTriggers();
  const isAutoSyncEnabled = triggers.some(t => t.getHandlerFunction() === 'runSyncProcess');

  return { 
    folderPairs: folderPairs, 
    syncFrequency: config.syncFrequency || 'hourly',
    isAutoSyncEnabled: isAutoSyncEnabled 
  };
}

function saveAppConfig(newConfig) {
  const { configSheet } = getOrCreateDb();
  const currentData = configSheet.getDataRange().getValues();
  const currentMap = {};
  for(let i=1; i<currentData.length; i++) {
    currentMap[currentData[i][0]] = currentData[i][1];
  }

  if (newConfig.folderPairs) {
    // ラベル情報も含めてJSON化して保存
    currentMap['folderPairs'] = JSON.stringify(newConfig.folderPairs);
  }
  if (newConfig.syncFrequency) {
    currentMap['syncFrequency'] = newConfig.syncFrequency;
  }
  
  configSheet.clearContents();
  configSheet.appendRow(['項目キー', '設定値']);
  configSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e9ecef');
  
  Object.keys(currentMap).forEach(key => {
    configSheet.appendRow([key, currentMap[key]]);
  });

  return { success: true };
}

function getSyncLogs() {
  const { logSheet } = getOrCreateDb();
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return [];

  const numRows = Math.min(lastRow - 1, 50);
  const startRow = Math.max(2, lastRow - numRows + 1);
  const data = logSheet.getRange(startRow, 1, numRows, 4).getValues();
  
  return data.reverse().map(row => ({
    date: new Date(row[0]).toLocaleString('ja-JP'),
    type: row[1],
    message: row[2],
    fileName: row[3]
  }));
}

function toggleAutoSync(enable, frequency = 'hourly') {
  const functionName = 'runSyncProcess';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if (t.getHandlerFunction() === functionName) ScriptApp.deleteTrigger(t); });

  saveAppConfig({ syncFrequency: frequency });

  if (enable) {
    let builder = ScriptApp.newTrigger(functionName).timeBased();
    if (frequency === 'daily') builder.everyDays(1).atHour(0);
    else if (frequency === 'weekly') builder.everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8);
    else builder.everyHours(1);
    
    builder.create();
    return { enabled: true, frequency: frequency, message: '自動転送をONにしました' };
  } else {
    return { enabled: false, frequency: frequency, message: '自動転送をOFFにしました' };
  }
}

/**
 * 【新機能】現在の進捗状況を取得する
 * クライアント側から定期的に呼ばれます
 */
function getProgress() {
  const cache = CacheService.getScriptCache();
  const progressJson = cache.get(CACHE_KEY_PROGRESS);
  if (progressJson) {
    return JSON.parse(progressJson);
  }
  return { percent: 0, status: '待機中...' };
}

// --- 4. コア機能（同期処理）大幅改修 ---

/**
 * ファイル同期を実行するメイン関数
 * @param {boolean} isDryRun - trueの場合、実際の書き込みを行わない（テストモード）
 */
function runSyncProcess(isDryRun = false) {
  // 自動実行(トリガー)からの呼び出しでは引数がないため、falseになる
  if (typeof isDryRun !== 'boolean') isDryRun = false;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (e) {
    return { success: false, message: '他の処理が実行中です。' };
  }

  // 進捗初期化
  const cache = CacheService.getScriptCache();
  cache.put(CACHE_KEY_PROGRESS, JSON.stringify({ percent: 0, status: '準備中...' }), 1800);

  const { configSheet, logSheet } = getOrCreateDb();
  
  const configData = configSheet.getDataRange().getValues();
  let folderPairs = [];
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'folderPairs') {
      try { folderPairs = JSON.parse(configData[i][1]); } catch(e) { folderPairs = []; }
    }
  }

  if (!folderPairs || folderPairs.length === 0) {
    lock.releaseLock();
    return { success: false, message: '転送設定が見つかりません。' };
  }

  const logs = [];
  const timestamp = new Date();
  let totalProcessed = 0;
  let totalErrors = 0;
  
  // 処理対象の総数（簡易的にペア数とする。ファイル数までは事前に分からないため）
  const totalSteps = folderPairs.length;

  folderPairs.forEach((pair, index) => {
    // 進捗更新
    const percent = Math.floor((index / totalSteps) * 100);
    const label = pair.label || `設定${index + 1}`;
    cache.put(CACHE_KEY_PROGRESS, JSON.stringify({ 
      percent: percent, 
      status: `[${label}] を確認中...` 
    }), 1800);

    const sourceId = pair.source;
    const targetId = pair.target;
    
    if (!sourceId || !targetId) return;

    try {
      const sourceFolder = DriveApp.getFolderById(sourceId);
      const targetFolder = DriveApp.getFolderById(targetId);
      const folderName = pair.label || sourceFolder.getName();
      
      const targetFilesMap = {};
      const targetFiles = targetFolder.getFiles();
      while (targetFiles.hasNext()) {
        const file = targetFiles.next();
        if (!file.isTrashed()) {
          targetFilesMap[file.getName()] = file;
        }
      }

      const sourceFiles = sourceFolder.getFiles();
      while (sourceFiles.hasNext()) {
        const sFile = sourceFiles.next();
        const sName = sFile.getName();

        try {
          const tFile = targetFilesMap[sName];
          const logPrefix = isDryRun ? '【テスト】' : '';
          const logTypeCreate = isDryRun ? 'テスト新規' : '新規';
          const logTypeUpdate = isDryRun ? 'テスト更新' : '更新';

          if (tFile) {
            if (sFile.getLastUpdated().getTime() > tFile.getLastUpdated().getTime()) {
              if (!isDryRun) {
                tFile.setTrashed(true);
                sFile.makeCopy(sName, targetFolder);
              }
              logs.push([timestamp, logTypeUpdate, `${logPrefix}[${folderName}] 上書き対象`, sName]);
              totalProcessed++;
            }
          } else {
            if (!isDryRun) {
              sFile.makeCopy(sName, targetFolder);
            }
            logs.push([timestamp, logTypeCreate, `${logPrefix}[${folderName}] 新規配布対象`, sName]);
            totalProcessed++;
          }
        } catch (fileError) {
          console.error(`File Error (${sName}): ${fileError.toString()}`);
          logs.push([timestamp, 'エラー', `失敗: ${fileError.message}`, sName]);
          totalErrors++;
        }
      }
    } catch (folderError) {
      console.error(`Folder Pair ${index + 1} Error: ${folderError.toString()}`);
      logs.push([timestamp, 'エラー', `${pair.label || 'フォルダ'}が見つかりません`, '-']);
      totalErrors++;
    }
  });
  
  // ログ保存（テストモードでもログは残す）
  if (logs.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logs.length, 4).setValues(logs);
  }

  // 完了通知
  cache.put(CACHE_KEY_PROGRESS, JSON.stringify({ percent: 100, status: '完了！' }), 1800);
  lock.releaseLock();

  const msg = isDryRun 
    ? `【テスト完了】${totalProcessed}件のファイルが対象です。(エラー: ${totalErrors}件)`
    : `${totalProcessed}件のファイルを配りました。(エラー: ${totalErrors}件)`;
    
  return { success: true, message: msg };
}

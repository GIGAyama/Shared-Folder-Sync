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
  // ⚠️ ここは以前 XFrameOptionsMode.ALLOWALL にしていた。
  //    ALLOWALL は「どのサイトからでも iframe で埋め込んでよい」という意味になる。
  //    このアプリは配布先のファイルをゴミ箱に送れる。よそのサイトに透明な iframe で
  //    重ねられ、先生が気づかないうちに「ファイルを転送する」を押させられると、
  //    児童の書き込んだファイルが消える（クリックジャッキング）。
  //    既定（DEFAULT）に戻すと、Google の画面の中でしか開けなくなる。
  //    URL を直接開いて使う分には、これで何も変わらない。
  //    Google サイトなどに貼り付けて使いたくなったときだけ、ここを見直すこと。
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
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
    // 自動実行では既存ファイルに触らない、を既定にする。
    // 名前が同じというだけで、児童が書き込んだファイルが無人で消えるのを防ぐため。
    configSheet.appendRow(['autoSyncNoOverwrite', 'true']);
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
    // 設定が無いときは「上書きしない」。安全なほうを既定にする。
    autoSyncNoOverwrite: String(config.autoSyncNoOverwrite) !== 'false',
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
  // false も保存したいので、!== undefined で見る（if (値) だと false が素通りする）
  if (newConfig.autoSyncNoOverwrite !== undefined) {
    currentMap['autoSyncNoOverwrite'] = newConfig.autoSyncNoOverwrite ? 'true' : 'false';
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
 * 配布先のファイルを、最後に更新した人の名前で言い表す。
 *
 * DriveApp には「最後に更新した人」を取る方法がない。
 * 拡張サービス「Drive」を有効にしてあれば聞けるので、あれば使い、
 * 無ければ所有者で代える。どちらも取れなければ「不明」と書く。
 * ここで例外を出して同期そのものを止めてしまわないよう、全部 try で包む。
 */
function describeLastEditor_(file) {
  try {
    if (typeof Drive !== 'undefined' && Drive.Files && Drive.Files.get) {
      let meta = null;
      // 拡張サービスの版によって返るものが違うので、両方見る
      try { meta = Drive.Files.get(file.getId(), { fields: 'lastModifyingUser(displayName,emailAddress)' }); }
      catch (e3) { meta = Drive.Files.get(file.getId()); }
      const u = meta && meta.lastModifyingUser;
      if (u && (u.displayName || u.emailAddress)) return u.displayName || u.emailAddress;
      if (meta && meta.lastModifyingUserName) return meta.lastModifyingUserName;
    }
  } catch (e) { /* 拡張サービスが無効／権限が足りないときはここに来る */ }

  try {
    const owner = file.getOwner();
    if (owner) return '所有者:' + owner.getEmail();
  } catch (e2) { /* 共有ドライブのファイルには所有者がいない */ }

  return '不明';
}

/** ログに書く日時。ファイル名にも使うので、記号を含めない形にする。 */
function stampForName_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
}

function stampForLog_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
}

/**
 * ファイル同期を実行するメイン関数
 *
 * ⚠️ 上書きは「名前が同じ」だけを見て配布先のファイルを消す処理である。
 *    配布先は児童が書き込むフォルダなので、ワークシートに答えを書いた子の
 *    ファイルが、毎時のトリガーで**誰にも聞かれずに**消えていた。
 *    そこで、
 *      ・自動実行では既定で既存ファイルに触らない（新規配布だけ行う）
 *      ・自動実行で上書きする設定にしたときは、消さずに退避（改名）する
 *      ・手動実行（確認ダイアログを通った操作）の動きは変えない。
 *        ただし上書き前の更新日時と、最後に更新した人を必ずログに残す
 *    とした。
 *
 * @param {boolean} isDryRun - trueの場合、実際の書き込みを行わない（テストモード）
 * @param {boolean} isManual - 画面のボタンから呼ばれたときだけ true。
 *                             時間主導トリガーはこの引数を渡さないので false になる。
 */
function runSyncProcess(isDryRun = false, isManual = false) {
  // 自動実行(トリガー)からの呼び出しでは引数がイベントオブジェクトになるため、falseになる
  if (typeof isDryRun !== 'boolean') isDryRun = false;
  // 画面からの呼び出しだけが true を渡す。トリガーからは絶対に true にならない。
  const isManualRun = (isManual === true);

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
  let noOverwriteSetting = true;   // 設定が無いときは安全なほう（上書きしない）
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'folderPairs') {
      try { folderPairs = JSON.parse(configData[i][1]); } catch(e) { folderPairs = []; }
    }
    if (configData[i][0] === 'autoSyncNoOverwrite') {
      noOverwriteSetting = String(configData[i][1]) !== 'false';
    }
  }

  // 自動実行のときだけ、この設定で上書きを止める。手動は今までどおり上書きする。
  const skipOverwrite = !isManualRun && noOverwriteSetting;

  if (!folderPairs || folderPairs.length === 0) {
    lock.releaseLock();
    return { success: false, message: '転送設定が見つかりません。' };
  }

  const logs = [];
  const timestamp = new Date();
  let totalProcessed = 0;
  let totalErrors = 0;
  let totalSkipped = 0;
  
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
              if (skipOverwrite) {
                // 自動実行では既存ファイルに触らない。
                // 児童が書き込んだあとかもしれないので、消さずに「見送った」と記録する。
                logs.push([timestamp, 'スキップ',
                  `[${folderName}] 自動実行のため上書きしませんでした（配布先の更新: ${stampForLog_(tFile.getLastUpdated())} / 最後に更新した人: ${describeLastEditor_(tFile)}）`,
                  sName]);
                totalSkipped++;
              } else {
                // 上書きする前に、配布先が「いつ・誰に」更新されたものだったかを必ず残す。
                // あとから「消えた」と言われたときに、これが無いと何も分からない。
                const beforeUpdated = stampForLog_(tFile.getLastUpdated());
                const beforeEditor = describeLastEditor_(tFile);
                let howRetired = '';

                if (!isDryRun) {
                  if (isManualRun) {
                    // 手動（確認ダイアログを通った操作）は今までどおりゴミ箱へ送る
                    tFile.setTrashed(true);
                    howRetired = 'ゴミ箱へ';
                  } else {
                    // 誰も見ていない自動実行では消さない。名前を変えて残しておく。
                    const backupName = `${sName}_backup_${stampForName_(timestamp)}`;
                    tFile.setName(backupName);
                    howRetired = `退避: ${backupName}`;
                  }
                  sFile.makeCopy(sName, targetFolder);
                } else {
                  howRetired = isManualRun ? 'ゴミ箱へ（予定）' : '退避（予定）';
                }

                logs.push([timestamp, logTypeUpdate,
                  `${logPrefix}[${folderName}] 上書き対象（${howRetired} / 上書き前の更新: ${beforeUpdated} / 最後に更新した人: ${beforeEditor}）`,
                  sName]);
                totalProcessed++;
              }
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

  const skipNote = totalSkipped > 0
    ? ` ／ 自動実行のため上書きしなかったもの: ${totalSkipped}件`
    : '';
  const msg = isDryRun 
    ? `【テスト完了】${totalProcessed}件のファイルが対象です。(エラー: ${totalErrors}件)${skipNote}`
    : `${totalProcessed}件のファイルを配りました。(エラー: ${totalErrors}件)${skipNote}`;
    
  return { success: true, message: msg, skipped: totalSkipped };
}

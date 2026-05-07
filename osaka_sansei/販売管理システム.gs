/*************************************************
 * 在庫管理システム 月次履歴実装完全版
 * 
 * 新機能:
 * - 08_月次履歴シートへの月次データ自動/手動保存
 * - 月替わり自動検知による前月データアーカイブ
 * - 履歴データの重複チェック・削除機能
 * - 月別サマリー分析機能
 * 
 * 設計思想:
 * - 将来のBI連携を見据えたデータベース的構造
 * - 企業運用に耐える堅牢性とエラーハンドリング
 * - DXコンサルティングでの横展開を前提とした汎用性
 *************************************************/

var MAX_DATA_ROWS = 2000;
var NO_ALERT_MESSAGE = '🎉 現在、対応が必要なアラートはありません';

/**
 * メニュー - 履歴機能を統合
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏭 在庫管理システム')
    .addItem('🚀 完全再計算・再書き込み', 'repairAll')
    .addSeparator()
    .addItem('🔄 03/04/05 データ再構築', 'rebuildAllData')
    .addItem('🔄 05 表示だけ再構築', 'rebuild05From04')
    .addSeparator()
    .addItem('📦 月次履歴を手動保存', 'saveCurrentMonthToHistory')
    .addItem('📊 履歴サマリー出力', 'outputHistorySummary')
    .addItem('🗑️ 指定月履歴削除', 'deleteSpecificMonthHistory')
    .addSeparator()
    .addItem('🎨 条件付き書式を再設定', 'setAllConditionalFormats')
    .addItem('📋 条件付き書式一覧を出力', 'listAllConditionalFormats')
    .addItem('🧹 条件付き書式をリセット', 'resetAllConditionalFormats')
    .addToUi();
}

/**
 * 完全修復 - 履歴機能対応
 */
function repairAll() {
  ensureSystemHeaders_();
  ensureMonthlyHistorySheet_(); // 履歴シート確保
  rebuildAllData();
  setAllConditionalFormats();

  SpreadsheetApp.getUi().alert(
    '✅ 完全再計算完了',
    '03/04/05の完全再計算と履歴機能の初期化が完了しました。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 全データ再構築 - 月替わり自動検知機能付き
 */
function rebuildAllData() {
  var sheets = getRequiredSheets_();
  var config = getConfigValues_();

  // 月替わり自動検知 → 前月データを履歴に保存
  autoArchiveIfMonthChanged_(sheets, config);

  var ctx = buildSystemContext_();

  var rows03 = build03Rows_(ctx);
  var rows04 = build04Rows_(ctx);
  var rows05 = build05RowsFrom04Rows_(rows04, ctx.config);

  write03Rows_(ctx.sheets.s03, rows03);
  write04Rows_(ctx.sheets.s04, rows04);
  ensure04Checkboxes_(ctx.sheets.s04, rows04.length);

  write05Rows_(ctx.sheets.s05, rows05);
  ensure05Checkboxes_(ctx.sheets.s05, countAlertRows_(rows05));

  SpreadsheetApp.flush();
}

/**
 * 04を正本として05を再構築
 */
function rebuild05From04() {
  var sheets = getRequiredSheets_();
  var config = getConfigValues_();
  var rows05 = build05RowsFromCurrent04_(sheets.s04, config);

  write05Rows_(sheets.s05, rows05);
  ensure05Checkboxes_(sheets.s05, countAlertRows_(rows05));

  SpreadsheetApp.flush();
}

/*************************************************
 * 月次履歴機能 (08_月次履歴)
 *************************************************/

/**
 * 08_月次履歴シートの初期化・取得
 * 
 * 列構成（将来のBI連携を考慮した正規化設計）:
 * A: 記録年月 (YYYY/MM)
 * B: 商品コード
 * C: 商品名
 * D: 月販売数
 * E: 月売上金額
 * F: 月末在庫数
 * G: 在庫日数
 * H: EC表示在庫
 * I: 発注残数
 * J: 当月仕入累計
 * K: 直近入荷日
 * L: 記録日時
 */
function ensureMonthlyHistorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s08 = ss.getSheetByName('08_月次履歴');

  if (!s08) {
    s08 = ss.insertSheet('08_月次履歴');

    var headers = [
      '記録年月',      // A: 2026/04 形式
      '商品コード',    // B: 
      '商品名',        // C: 
      '月販売数',      // D: 
      '月売上金額',    // E: 
      '月末在庫数',    // F: 
      '在庫日数',      // G: 
      'EC表示在庫',    // H: 
      '発注残数',      // I: 
      '当月仕入累計',  // J: 
      '直近入荷日',    // K: 
      '記録日時'       // L: システム記録時刻
    ];

    s08.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ヘッダー装飾（企業らしい品格のある色調）
    s08.getRange(1, 1, 1, headers.length)
      .setBackground('#1A237E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');

    // 使いやすさの設定
    s08.setFrozenColumns(2);  // 商品コード・商品名を固定
    s08.setFrozenRows(1);     // ヘッダー固定
    s08.autoResizeColumns(1, headers.length);
  }

  return s08;
}

/**
 * 現在の月データを手動で履歴保存
 * メニュー「📦 月次履歴を手動保存」から呼び出し
 */
function saveCurrentMonthToHistory() {
  var ui = SpreadsheetApp.getUi();
  
  var ctx = buildSystemContext_();
  var targetLabel = ctx.config.targetMonthLabel;
  
  var s08 = ensureMonthlyHistorySheet_();
  
  // 重複チェック（冪等性の保証）
  var duplicateInfo = checkDuplicateMonth_(s08, targetLabel);
  
  if (duplicateInfo.exists) {
    var response = ui.alert(
      '⚠️ 重複データ確認',
      targetLabel + ' のデータが既に ' + duplicateInfo.count + ' 件存在します。\n\n' +
      '上書きしますか？（既存データを削除して新規保存）',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      ui.alert('ℹ️ キャンセル', '履歴保存をキャンセルしました。', ui.ButtonSet.OK);
      return;
    }
    
    deleteMonthData_(s08, targetLabel);
  }
  
  // 履歴データ構築
  var historyRows = buildHistoryDataRows_(ctx, targetLabel);
  
  if (historyRows.length === 0) {
    ui.alert('⚠️ データなし', '保存対象のデータが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  
  // 履歴シートへ書き込み
  var lastRow = Math.max(s08.getLastRow(), 1);
  s08.getRange(lastRow + 1, 1, historyRows.length, 12).setValues(historyRows);
  
  // 条件付き書式適用
  setHistoryConditionalFormats_(s08);
  
  SpreadsheetApp.flush();
  
  ui.alert(
    '✅ 履歴保存完了',
    targetLabel + ' のデータを ' + historyRows.length + ' 件保存しました。\n\n' +
    '08_月次履歴シートで確認できます。',
    ui.ButtonSet.OK
  );
}

/**
 * 月替わり自動検知と前月データの自動アーカイブ
 * rebuildAllData() から自動的に呼び出される
 * 
 * 仕組み:
 * - 04シートK2の「更新年月」と現在の設定年月を比較
 * - 異なる場合は月が変わったと判定し、旧データを自動保存
 */
function autoArchiveIfMonthChanged_(sheets, config) {
  var s04 = sheets.s04;
  
  // 04シートの現在の更新年月を取得
  var currentMonthInSheet = trimString_(s04.getRange('K2').getDisplayValue());
  var newTargetMonth = config.targetMonthLabel;
  
  // 初回実行 or 月が変わっていない場合はスキップ
  if (!currentMonthInSheet || currentMonthInSheet === newTargetMonth) {
    return;
  }
  
  var s08 = ensureMonthlyHistorySheet_();
  
  // 既存の同月データがあれば削除（上書き保存）
  var duplicateInfo = checkDuplicateMonth_(s08, currentMonthInSheet);
  if (duplicateInfo.exists) {
    deleteMonthData_(s08, currentMonthInSheet);
  }
  
  // 現在のシートデータから履歴を構築
  var historyRows = buildHistoryFromCurrentSheets_(sheets, currentMonthInSheet);
  
  if (historyRows.length > 0) {
    var lastRow = Math.max(s08.getLastRow(), 1);
    s08.getRange(lastRow + 1, 1, historyRows.length, 12).setValues(historyRows);
    
    setHistoryConditionalFormats_(s08);
    
    // ユーザーへの通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      currentMonthInSheet + ' のデータを自動で履歴保存しました（' + historyRows.length + ' 件）',
      '📦 自動アーカイブ完了',
      5
    );
  }
}

/**
 * システムコンテキストから履歴データ行を構築
 * 新規計算結果から履歴を作成する場合に使用
 */
function buildHistoryDataRows_(ctx, targetLabel) {
  var rows = [];
  var now = new Date();
  
  for (var i = 0; i < ctx.masterItems.length; i++) {
    var item = ctx.masterItems[i];
    
    // 03シートデータ（売上）
    var sales = getMapNumber_(ctx.salesAgg.qtyMap, item.code);
    var salesAmount = getMapNumber_(ctx.salesAgg.amountMap, item.code);
    
    // 04シートデータ（在庫）
    var stock = getMapNumber_(ctx.ledgerAgg.stockMap, item.code);
    var orderBalance = getMapNumber_(ctx.ledgerAgg.orderBalanceMap, item.code);
    var monthReceipt = getMapNumber_(ctx.ledgerAgg.monthlyReceiptMap, item.code);
    var latestReceipt = getMapValue_(ctx.ledgerAgg.latestReceiptMap, item.code, '');
    
    // 計算値
    var stockDays = (sales === 0) ? 0 : round1_(stock / sales * ctx.config.stockConvertDays);
    var ecStock = floor1_(stock * ctx.config.ecRate);
    
    rows.push([
      targetLabel,          // A: 記録年月
      item.code,            // B: 商品コード
      item.name,            // C: 商品名
      sales,                // D: 月販売数
      salesAmount,          // E: 月売上金額
      stock,                // F: 月末在庫数
      stockDays,            // G: 在庫日数
      ecStock,              // H: EC表示在庫
      orderBalance,         // I: 発注残数
      monthReceipt,         // J: 当月仕入累計
      latestReceipt || '',  // K: 直近入荷日
      now                   // L: 記録日時
    ]);
  }
  
  return rows;
}

/**
 * 現在の03/04シートデータから履歴行を構築
 * 月替わり自動アーカイブで既存データを保存する場合に使用
 */
function buildHistoryFromCurrentSheets_(sheets, monthLabel) {
  var s03 = sheets.s03;
  var s04 = sheets.s04;
  var now = new Date();
  
  // 03シートから売上データを取得
  var lastRow03 = Math.max(s03.getLastRow(), 2);
  var values03 = s03.getRange(2, 1, Math.min(MAX_DATA_ROWS - 1, lastRow03 - 1), 4).getValues();
  
  // 04シートから在庫データを取得
  var lastRow04 = Math.max(s04.getLastRow(), 2);
  var values04 = s04.getRange(2, 1, Math.min(MAX_DATA_ROWS - 1, lastRow04 - 1), 13).getValues();
  
  // 04データを商品コードでMap化（効率的な検索のため）
  var stockDataMap = {};
  for (var i = 0; i < values04.length; i++) {
    var code = trimString_(values04[i][0]);
    if (code !== '') {
      stockDataMap[code] = {
        stock: values04[i][2],        // C: 現在在庫数
        stockDays: values04[i][4],    // E: 在庫日数
        ecStock: values04[i][5],      // F: EC表示在庫
        orderBalance: values04[i][6], // G: 発注残数
        monthReceipt: values04[i][7], // H: 当月仕入累計
        latestReceipt: values04[i][8] // I: 直近入荷日
      };
    }
  }
  
  var historyRows = [];
  
  for (var j = 0; j < values03.length; j++) {
    var itemCode = trimString_(values03[j][0]);
    if (itemCode === '') continue;
    
    var stockData = stockDataMap[itemCode] || {
      stock: 0, stockDays: 0, ecStock: 0, 
      orderBalance: 0, monthReceipt: 0, latestReceipt: ''
    };
    
    historyRows.push([
      monthLabel,               // A: 記録年月
      itemCode,                 // B: 商品コード
      values03[j][1],           // C: 商品名
      values03[j][2],           // D: 月販売数
      values03[j][3],           // E: 月売上金額
      stockData.stock,          // F: 月末在庫数
      stockData.stockDays,      // G: 在庫日数
      stockData.ecStock,        // H: EC表示在庫
      stockData.orderBalance,   // I: 発注残数
      stockData.monthReceipt,   // J: 当月仕入累計
      stockData.latestReceipt,  // K: 直近入荷日
      now                       // L: 記録日時
    ]);
  }
  
  return historyRows;
}

/**
 * 同一年月データの存在チェック
 * 重複保存を防ぐための冪等性保証
 */
function checkDuplicateMonth_(sheet, targetLabel) {
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    return { exists: false, count: 0 };
  }
  
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var count = 0;
  
  for (var i = 0; i < values.length; i++) {
    if (trimString_(values[i][0]) === targetLabel) {
      count++;
    }
  }
  
  return { exists: count > 0, count: count };
}

/**
 * 指定年月のデータを削除
 * 後ろから削除して行番号ズレを防ぐ安全な実装
 */
function deleteMonthData_(sheet, targetLabel) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  // 後ろから削除（行番号ズレ防止の定石）
  for (var i = values.length - 1; i >= 0; i--) {
    if (trimString_(values[i][0]) === targetLabel) {
      sheet.deleteRow(i + 2); // +2はヘッダー行分
    }
  }
}

/**
 * 指定月履歴削除（メニューから呼び出し）
 */
function deleteSpecificMonthHistory() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    '🗑️ 指定月履歴削除',
    '削除したい記録年月を "YYYY/MM" 形式で入力してください（例: 2026/03）',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var targetLabel = trimString_(response.getResponseText());
  
  // 入力形式チェック
  if (!targetLabel.match(/^\d{4}\/\d{2}$/)) {
    ui.alert('⚠️ フォーマット不正', 'YYYY/MM 形式で入力してください。', ui.ButtonSet.OK);
    return;
  }
  
  var s08 = ensureMonthlyHistorySheet_();
  var duplicateInfo = checkDuplicateMonth_(s08, targetLabel);
  
  if (!duplicateInfo.exists) {
    ui.alert('ℹ️ 対象なし', targetLabel + ' の履歴データは存在しません。', ui.ButtonSet.OK);
    return;
  }
  
  var confirm = ui.alert(
    '⚠️ 最終確認',
    targetLabel + ' の履歴データ ' + duplicateInfo.count + ' 行を削除します。\n' +
    'この操作は元に戻せません。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) {
    ui.alert('ℹ️ キャンセル', '削除をキャンセルしました。', ui.ButtonSet.OK);
    return;
  }
  
  deleteMonthData_(s08, targetLabel);
  
  ui.alert(
    '✅ 削除完了',
    targetLabel + ' の履歴データを削除しました。',
    ui.ButtonSet.OK
  );
}

/*************************************************
 * システムコンテキスト構築
 *************************************************/

function buildSystemContext_() {
  var sheets = getRequiredSheets_();
  var config = getConfigValues_();
  var masterItems = collectMasterItems_(sheets.s01);
  var salesAgg = collectSalesAggregatesFrom02_(sheets.s02, config);
  var ledgerAgg = collectLedgerAggregates_(sheets.s06, config);
  var existing04State = readExisting04State_(sheets.s04);

  return {
    sheets: sheets,
    config: config,
    masterItems: masterItems,
    salesAgg: salesAgg,
    ledgerAgg: ledgerAgg,
    existing04State: existing04State
  };
}

function ensureSystemHeaders_() {
  var sheets = getRequiredSheets_();
  var s03 = sheets.s03;
  var s04 = sheets.s04;
  var s05 = sheets.s05;

  // 03_月販集計
  s03.getRange('A1').setValue('商品コード');
  s03.getRange('B1').setValue('商品名');
  s03.getRange('C1').setValue('月販売数');
  s03.getRange('D1').setValue('月売上金額');

  // 04_在庫日数_運用
  s04.getRange('A1').setValue('商品コード');
  s04.getRange('B1').setValue('商品名');
  s04.getRange('C1').setValue('現在在庫数');
  s04.getRange('D1').setValue('月販売数');
  s04.getRange('E1').setValue('在庫日数');
  s04.getRange('F1').setValue('EC表示在庫');
  s04.getRange('G1').setValue('発注残数');
  s04.getRange('H1').setValue('当月仕入累計');
  s04.getRange('I1').setValue('直近入荷日');
  s04.getRange('J1').setValue('確認済');
  s04.getRange('K1').setValue('更新年月');
  s04.getRange('L1').setValue('メモ');
  s04.getRange('M1').setValue('管理指示保存');

  // 05_差分確認
  s05.getRange('A1').setValue('商品コード');
  s05.getRange('B1').setValue('商品名');
  s05.getRange('C1').setValue('現在在庫数');
  s05.getRange('D1').setValue('月販売数');
  s05.getRange('E1').setValue('在庫日数');
  s05.getRange('F1').setValue('EC表示在庫');
  s05.getRange('G1').setValue('発注残数');
  s05.getRange('H1').setValue('当月仕入累計');
  s05.getRange('I1').setValue('直近入荷日');
  s05.getRange('J1').setValue('確認済');
  s05.getRange('K1').setValue('更新年月');
  s05.getRange('L1').setValue('メモ');
  s05.getRange('M1').setValue('アラート理由');
  s05.getRange('N1').setValue('管理指示');
}

function getRequiredSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheets = {
    ss: ss,
    s00: ss.getSheetByName('00_設定'),
    s01: ss.getSheetByName('01_SKUマスタ'),
    s02: ss.getSheetByName('02_売上CSV取込'),
    s03: ss.getSheetByName('03_月販集計'),
    s04: ss.getSheetByName('04_在庫日数_運用'),
    s05: ss.getSheetByName('05_差分確認'),
    s06: ss.getSheetByName('06_入出庫台帳')
  };

  if (!sheets.s00) throw new Error('00_設定シートが見つかりません。');
  if (!sheets.s01) throw new Error('01_SKUマスタシートが見つかりません。');
  if (!sheets.s02) throw new Error('02_売上CSV取込シートが見つかりません。');
  if (!sheets.s03) throw new Error('03_月販集計シートが見つかりません。');
  if (!sheets.s04) throw new Error('04_在庫日数_運用シートが見つかりません。');
  if (!sheets.s05) throw new Error('05_差分確認シートが見つかりません。');
  if (!sheets.s06) throw new Error('06_入出庫台帳シートが見つかりません。');

  return sheets;
}

function getConfigValues_() {
  var s00 = getRequiredSheets_().s00;
  var targetMonthRaw = s00.getRange('B2').getValue();
  var monthInfo = parseTargetMonth_(targetMonthRaw);
  var stockConvertDays = toNumber_(s00.getRange('B3').getValue(), 0);
  var ecRate = toNumber_(s00.getRange('B4').getValue(), 0);
  var juuten = toNumber_(s00.getRange('B5').getValue(), 0);
  var minaoshi = toNumber_(s00.getRange('B6').getValue(), 0);
  var hacchu = toNumber_(s00.getRange('B7').getValue(), juuten);

  return {
    targetMonthRaw: targetMonthRaw,
    targetMonthLabel: monthInfo.label,
    monthStart: monthInfo.start,
    monthEnd: monthInfo.end,
    stockConvertDays: stockConvertDays,
    ecRate: ecRate,
    juuten: juuten,
    minaoshi: minaoshi,
    hacchu: hacchu
  };
}
/*************************************************
 * データ収集・集計処理
 *************************************************/

function collectMasterItems_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values = sheet.getRange(2, 2, lastRow - 1, 3).getValues();
  var items = [];
  var i;
  var code;
  var name;
  var flag;

  for (i = 0; i < values.length; i++) {
    code = trimString_(values[i][0]);
    name = trimString_(values[i][1]);
    flag = trimString_(values[i][2]);

    if (code === '') continue;
    if (!isYesFlag_(flag)) continue;

    items.push({
      code: code,
      name: name
    });
  }

  return items;
}

function collectSalesAggregatesFrom02_(sheet, config) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  var qtyMap = {};
  var amountMap = {};

  var i;
  var saleDate;
  var dateObj;
  var code;
  var qty;
  var amount;

  for (i = 0; i < values.length; i++) {
    saleDate = values[i][2];
    code = trimString_(values[i][3]);
    qty = toNumber_(values[i][7], 0);
    amount = toNumber_(values[i][8], 0);

    if (code === '') continue;

    dateObj = parseDateValue_(saleDate);
    if (!isDateInRange_(dateObj, config.monthStart, config.monthEnd)) continue;

    if (!qtyMap.hasOwnProperty(code)) qtyMap[code] = 0;
    if (!amountMap.hasOwnProperty(code)) amountMap[code] = 0;

    qtyMap[code] += qty;
    amountMap[code] += amount;
  }

  return {
    qtyMap: qtyMap,
    amountMap: amountMap
  };
}

function collectLedgerAggregates_(sheet, config) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  var stockMap = {};
  var orderBalanceMap = {};
  var monthlyReceiptMap = {};
  var latestReceiptMap = {};

  var i;
  var dateValue;
  var dateObj;
  var code;
  var type;
  var qty;
  var inTargetMonth;

  for (i = 0; i < values.length; i++) {
    dateValue = values[i][0];
    code = trimString_(values[i][1]);
    type = trimString_(values[i][3]);
    qty = toNumber_(values[i][4], 0);

    if (code === '') continue;
    if (type === '') continue;
    if (qty === 0) continue;

    if (!stockMap.hasOwnProperty(code)) stockMap[code] = 0;
    if (!orderBalanceMap.hasOwnProperty(code)) orderBalanceMap[code] = 0;
    if (!monthlyReceiptMap.hasOwnProperty(code)) monthlyReceiptMap[code] = 0;
    if (!latestReceiptMap.hasOwnProperty(code)) latestReceiptMap[code] = '';

    dateObj = parseDateValue_(dateValue);
    inTargetMonth = isDateInRange_(dateObj, config.monthStart, config.monthEnd);

    if (type === '入荷') {
      stockMap[code] += qty;
      orderBalanceMap[code] -= qty;

      if (inTargetMonth) {
        monthlyReceiptMap[code] += qty;
      }

      if (dateObj) {
        if (!latestReceiptMap[code] || dateObj.getTime() > latestReceiptMap[code].getTime()) {
          latestReceiptMap[code] = dateObj;
        }
      }
    } else if (type === '出荷') {
      stockMap[code] -= qty;
    } else if (type === '棚卸') {
      stockMap[code] += qty;
    } else if (type === '発注') {
      orderBalanceMap[code] += qty;
    }
  }

  return {
    stockMap: stockMap,
    orderBalanceMap: orderBalanceMap,
    monthlyReceiptMap: monthlyReceiptMap,
    latestReceiptMap: latestReceiptMap
  };
}

function readExisting04State_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  var map = {};
  var i;
  var code;
  var confirmed;
  var memo;
  var instruction;

  for (i = 0; i < values.length; i++) {
    code = trimString_(values[i][0]);
    if (code === '') continue;

    confirmed = values[i][9] === true;
    memo = values[i][11];
    instruction = values[i][12];

    map[code] = {
      confirmed: confirmed,
      memo: memo === null ? '' : memo,
      instruction: instruction === null ? '' : instruction
    };
  }

  return map;
}

/*************************************************
 * 行データ構築・書き込み処理
 *************************************************/

function build03Rows_(ctx) {
  var rows = [];
  var i;
  var item;
  var salesQty;
  var salesAmount;

  for (i = 0; i < ctx.masterItems.length; i++) {
    item = ctx.masterItems[i];
    salesQty = getMapNumber_(ctx.salesAgg.qtyMap, item.code);
    salesAmount = getMapNumber_(ctx.salesAgg.amountMap, item.code);

    rows.push([
      item.code,
      item.name,
      salesQty,
      salesAmount
    ]);
  }

  return rows;
}

function build04Rows_(ctx) {
  var rows = [];
  var i;
  var item;
  var stock;
  var sales;
  var stockDays;
  var ecStock;
  var orderBalance;
  var monthReceipt;
  var latestReceipt;
  var existing;
  var confirmed;
  var memo;
  var instruction;

  for (i = 0; i < ctx.masterItems.length; i++) {
    item = ctx.masterItems[i];

    stock = getMapNumber_(ctx.ledgerAgg.stockMap, item.code);
    sales = getMapNumber_(ctx.salesAgg.qtyMap, item.code);
    orderBalance = getMapNumber_(ctx.ledgerAgg.orderBalanceMap, item.code);
    monthReceipt = getMapNumber_(ctx.ledgerAgg.monthlyReceiptMap, item.code);
    latestReceipt = getMapValue_(ctx.ledgerAgg.latestReceiptMap, item.code, '');

    if (sales === 0) {
      stockDays = 0;
    } else {
      stockDays = round1_(stock / sales * ctx.config.stockConvertDays);
    }

    ecStock = floor1_(stock * ctx.config.ecRate);

    existing = ctx.existing04State[item.code] || {};
    confirmed = existing.confirmed === true;
    memo = existing.memo || '';
    instruction = existing.instruction || '';

    rows.push([
      item.code,
      item.name,
      stock,
      sales,
      stockDays,
      ecStock,
      orderBalance,
      monthReceipt,
      latestReceipt || '',
      confirmed,
      ctx.config.targetMonthLabel,
      memo,
      instruction
    ]);
  }

  return rows;
}

function build05RowsFrom04Rows_(rows04, config) {
  var rows05 = [];
  var i;
  var row04;
  var confirmed;
  var stockDays;
  var orderBalance;
  var alertReason;

  for (i = 0; i < rows04.length; i++) {
    row04 = rows04[i];
    confirmed = row04[9] === true;
    stockDays = toNumber_(row04[4], 0);
    orderBalance = toNumber_(row04[6], 0);

    if (confirmed) continue;
    if (!isAlertTarget_(stockDays, orderBalance, config)) continue;

    alertReason = buildAlertReason_(stockDays, orderBalance, config);

    rows05.push([
      row04[0],   // A 商品コード
      row04[1],   // B 商品名
      row04[2],   // C 現在在庫数
      row04[3],   // D 月販売数
      row04[4],   // E 在庫日数
      row04[5],   // F EC表示在庫
      row04[6],   // G 発注残数
      row04[7],   // H 当月仕入累計
      row04[8],   // I 直近入荷日
      false,      // J 05側チェック
      row04[10],  // K 更新年月
      row04[11],  // L メモ
      alertReason,// M アラート理由
      row04[12]   // N 管理指示
    ]);
  }

  if (rows05.length === 0) {
    rows05.push([
      NO_ALERT_MESSAGE, '', '', '', '', '', '', '', '', '', '', '', '', ''
    ]);
  }

  return rows05;
}

function build05RowsFromCurrent04_(sheet04, config) {
  var lastRow = Math.max(sheet04.getLastRow(), 2);
  var rowCount = Math.min(MAX_DATA_ROWS - 1, lastRow - 1);
  var values = sheet04.getRange(2, 1, rowCount, 13).getValues();
  var rows04 = [];
  var i;
  var code;

  for (i = 0; i < values.length; i++) {
    code = trimString_(values[i][0]);
    if (code === '') continue;
    rows04.push(values[i]);
  }

  return build05RowsFrom04Rows_(rows04, config);
}

function write03Rows_(sheet, rows) {
  writeRowsBlock_(sheet, 2, 1, 4, rows);
}

function write04Rows_(sheet, rows) {
  writeRowsBlock_(sheet, 2, 1, 13, rows);
}

function write05Rows_(sheet, rows) {
  writeRowsBlock_(sheet, 2, 1, 14, rows);
}

function writeRowsBlock_(sheet, startRow, startCol, totalCols, rows) {
  var totalRows = MAX_DATA_ROWS - 1;
  var matrix = [];
  var i;
  var row;

  for (i = 0; i < totalRows; i++) {
    if (i < rows.length) {
      row = padRow_(rows[i], totalCols);
      matrix.push(row);
    } else {
      matrix.push(buildBlankRow_(totalCols));
    }
  }

  sheet.getRange(startRow, startCol, totalRows, totalCols).setValues(matrix);
}

function ensure04Checkboxes_(sheet, dataCount) {
  var rangeAll = sheet.getRange('J2:J' + MAX_DATA_ROWS);

  try {
    rangeAll.clearDataValidations();

    if (dataCount > 0) {
      sheet.getRange(2, 10, dataCount, 1).insertCheckboxes();
    }
  } catch (err) {
    // テーブルや型付き列と衝突する場合は boolean 値のみ維持
  }
}

function ensure05Checkboxes_(sheet, dataCount) {
  var rangeAll = sheet.getRange('J2:J' + MAX_DATA_ROWS);

  try {
    rangeAll.clearDataValidations();

    if (dataCount > 0) {
      sheet.getRange(2, 10, dataCount, 1).insertCheckboxes();
    }
  } catch (err) {
    // 念のため保護
  }
}

function countAlertRows_(rows05) {
  if (!rows05 || rows05.length === 0) return 0;
  if (rows05.length === 1 && trimString_(rows05[0][0]) === NO_ALERT_MESSAGE) return 0;
  return rows05.length;
}

/*************************************************
 * イベント処理
 *************************************************/

function onEdit(e) {
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var row = range.getRow();
  var col = range.getColumn();
  var ss = e.source;

  if (row < 2) return;

  // 05_差分確認 J: 確認済みチェック
  if (sheetName === '05_差分確認' && col === 10) {
    if (range.getValue() !== true) return;

    handle05ConfirmCheck_(row);
    ss.toast('05シートの確認済みを04シートへ反映しました', '✅ 処理完了', 3);
    return;
  }

  // 05_差分確認 L: メモ編集 -> 04 L
  if (sheetName === '05_差分確認' && col === 12) {
    writeBack05FieldTo04_(row, col, range.getValue());
    rebuild05From04();
    return;
  }

  // 05_差分確認 N: 管理指示編集 -> 04 M
  if (sheetName === '05_差分確認' && col === 14) {
    writeBack05FieldTo04_(row, col, range.getValue());
    rebuild05From04();
    return;
  }

  // 04_在庫日数_運用 J/L/M編集 -> 05再構築
  if (sheetName === '04_在庫日数_運用' && (col === 10 || col === 12 || col === 13)) {
    rebuild05From04();

    if (col === 10 && range.getValue() === true) {
      ss.toast('04シートの確認済み変更を05へ反映しました', 'ℹ️ 更新', 2);
    }
    return;
  }

  // 00/01/02/06 の変更は全再計算
  if (
    sheetName === '00_設定' ||
    sheetName === '01_SKUマスタ' ||
    sheetName === '02_売上CSV取込' ||
    sheetName === '06_入出庫台帳'
  ) {
    rebuildAllData();
    return;
  }
}

function handle05ConfirmCheck_(row) {
  var sheets = getRequiredSheets_();
  var s04 = sheets.s04;
  var s05 = sheets.s05;

  var itemCode = trimString_(s05.getRange(row, 1).getDisplayValue());

  // 自分のチェックは戻す
  s05.getRange(row, 10).setValue(false);

  if (itemCode === '' || itemCode === NO_ALERT_MESSAGE) {
    rebuild05From04();
    return;
  }

  var finder = s04.getRange('A2:A' + MAX_DATA_ROWS).createTextFinder(itemCode).matchEntireCell(true).findNext();
  if (!finder) {
    rebuild05From04();
    return;
  }

  s04.getRange(finder.getRow(), 10).setValue(true);

  SpreadsheetApp.flush();
  Utilities.sleep(200);

  rebuild05From04();
}

function writeBack05FieldTo04_(row, col, value) {
  var sheets = getRequiredSheets_();
  var s04 = sheets.s04;
  var s05 = sheets.s05;

  var itemCode = trimString_(s05.getRange(row, 1).getDisplayValue());
  if (itemCode === '' || itemCode === NO_ALERT_MESSAGE) return;

  var finder = s04.getRange('A2:A' + MAX_DATA_ROWS).createTextFinder(itemCode).matchEntireCell(true).findNext();
  if (!finder) return;

  var targetRow = finder.getRow();

  // 05 L -> 04 L
  if (col === 12) {
    s04.getRange(targetRow, 12).setValue(value);
    return;
  }

  // 05 N -> 04 M
  if (col === 14) {
    s04.getRange(targetRow, 13).setValue(value);
    return;
  }
}

/*************************************************
 * アラート判定・理由構築
 *************************************************/

function isAlertTarget_(stockDays, orderBalance, config) {
  if (stockDays <= config.juuten) return true;
  if (orderBalance !== 0) return true;
  if (stockDays >= config.minaoshi) return true;
  return false;
}

function buildAlertReason_(stockDays, orderBalance, config) {
  var parts = [];

  if (stockDays <= config.juuten) {
    parts.push('🔴 在庫不足(' + stockDays + '日)');
  }

  if (orderBalance > 0) {
    parts.push('🟡 発注残あり(' + orderBalance + '個)');
  }

  if (orderBalance < 0) {
    parts.push('⚠️ 過剰入荷(' + Math.abs(orderBalance) + '個)');
  }

  if (stockDays >= config.minaoshi) {
    parts.push('🔵 在庫過多(' + stockDays + '日)');
  }

  return parts.join(' / ');
}

/*************************************************
 * 履歴サマリー出力機能
 *************************************************/

function outputHistorySummary() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var s08 = ss.getSheetByName('08_月次履歴');
  if (!s08 || s08.getLastRow() < 2) {
    ui.alert('⚠️ データなし', '08_月次履歴にデータがありません。', ui.ButtonSet.OK);
    return;
  }

  // 既存サマリーシートを削除して再作成
  var summaryName = '09_履歴サマリー';
  var s09 = ss.getSheetByName(summaryName);
  if (s09) {
    ss.deleteSheet(s09);
  }
  s09 = ss.insertSheet(summaryName);

  var lastRow = s08.getLastRow();
  var values = s08.getRange(2, 1, lastRow - 1, 12).getValues();

  // 月別集計データの構築
  var monthSummary = {};
  var i;
  var monthLabel;
  var salesQty;
  var salesAmount;
  var stock;

  for (i = 0; i < values.length; i++) {
    monthLabel = trimString_(values[i][0]);
    if (monthLabel === '') continue;

    salesQty = toNumber_(values[i][3], 0);
    salesAmount = toNumber_(values[i][4], 0);
    stock = toNumber_(values[i][5], 0);

    if (!monthSummary[monthLabel]) {
      monthSummary[monthLabel] = {
        totalQty: 0,
        totalAmount: 0,
        totalStock: 0,
        itemCount: 0
      };
    }

    monthSummary[monthLabel].totalQty += salesQty;
    monthSummary[monthLabel].totalAmount += salesAmount;
    monthSummary[monthLabel].totalStock += stock;
    monthSummary[monthLabel].itemCount += 1;
  }

  // 月リストを時系列順でソート
  var months = Object.keys(monthSummary).sort();
  
  if (months.length === 0) {
    ui.alert('⚠️ データなし', '集計対象データが見つかりません。', ui.ButtonSet.OK);
    ss.deleteSheet(s09);
    return;
  }

  // サマリーデータの構築
  var summaryData = [];
  var j;
  var monthData;
  var avgStock;

  for (j = 0; j < months.length; j++) {
    monthData = monthSummary[months[j]];
    avgStock = monthData.itemCount > 0 ? Math.round(monthData.totalStock / monthData.itemCount) : 0;

    summaryData.push([
      months[j],
      monthData.totalQty,
      monthData.totalAmount,
      avgStock,
      monthData.itemCount
    ]);
  }

  // ヘッダーとデータの書き込み
  var headers = [
    '記録年月',
    '合計販売数',
    '合計売上金額',
    '平均在庫数',
    '商品数'
  ];

  s09.getRange(1, 1, 1, headers.length).setValues([headers]);
  s09.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);

  // 装飾
  s09.getRange(1, 1, 1, headers.length)
    .setBackground('#004D40')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  // 売上金額列の数値書式
  s09.getRange(2, 3, summaryData.length, 1).setNumberFormat('#,##0');

  s09.autoResizeColumns(1, headers.length);
  s09.setFrozenRows(1);

  ui.alert(
    '📊 サマリー出力完了',
    '09_履歴サマリーシートに出力しました。\n\n' +
    '対象期間: ' + months[0] + ' 〜 ' + months[months.length - 1] + '\n' +
    '対象月数: ' + months.length + ' ヶ月',
    ui.ButtonSet.OK
  );
}

/*************************************************
 * 条件付き書式設定
 *************************************************/

function setAllConditionalFormats() {
  var sheets = getRequiredSheets_();
  var s03 = sheets.s03;
  var s04 = sheets.s04;
  var s05 = sheets.s05;
  var s06 = sheets.s06;
  var cfg = getConfigValues_();

  var juuten = cfg.juuten;
  var minaoshi = cfg.minaoshi;
  var hacchu = cfg.hacchu;

  // 04_在庫日数_運用
  var rules04 = [];
  var range04 = s04.getRange('A2:M' + MAX_DATA_ROWS);

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J2=TRUE')
      .setBackground('#F8F9FA')
      .setFontColor('#6C757D')
      .setRanges([range04])
      .build()
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($E2),$E2<=' + juuten + ')')
      .setBackground('#FFEBEE')
      .setFontColor('#C62828')
      .setBold(true)
      .setRanges([range04])
      .build()
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G2),$G2<0)')
      .setBackground('#FFEBEE')
      .setFontColor('#C62828')
      .setBold(true)
      .setRanges([range04])
      .build()
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G2),$G2>0)')
      .setBackground('#FFFDE7')
      .setFontColor('#F57F17')
      .setRanges([range04])
      .build()
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($E2),$E2<=' + hacchu + ',$E2>' + juuten + ')')
      .setBackground('#FFFDE7')
      .setFontColor('#F57F17')
      .setRanges([range04])
      .build()
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($E2),$E2>=' + minaoshi + ')')
      .setBackground('#FFF3E0')
      .setFontColor('#E65100')
      .setRanges([range04])
      .build()
  );

  s04.setConditionalFormatRules(rules04);

  // 05_差分確認
  var rules05 = [];
  var range05 = s05.getRange('A2:N' + MAX_DATA_ROWS);

  rules05.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($E2),$E2<=' + juuten + ')')
      .setBackground('#FFEBEE')
      .setFontColor('#C62828')
      .setBold(true)
      .setRanges([range05])
      .build()
  );

  rules05.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G2),$G2<0)')
      .setBackground('#FFEBEE')
      .setFontColor('#C62828')
      .setBold(true)
      .setRanges([range05])
      .build()
  );

  rules05.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G2),$G2>0)')
      .setBackground('#FFFDE7')
      .setFontColor('#F57F17')
      .setRanges([range05])
      .build()
  );

  rules05.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($E2),$E2>=' + minaoshi + ')')
      .setBackground('#E3F2FD')
      .setFontColor('#1565C0')
      .setRanges([range05])
      .build()
  );

  s05.setConditionalFormatRules(rules05);

  // 06_入出庫台帳
  var rules06 = [];
  var range06 = s06.getRange('A2:G' + MAX_DATA_ROWS);

  rules06.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="発注"')
      .setBackground('#E3F2FD')
      .setFontColor('#1565C0')
      .setRanges([range06])
      .build()
  );

  rules06.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="入荷"')
      .setBackground('#E8F5E9')
      .setFontColor('#2E7D32')
      .setRanges([range06])
      .build()
  );

  rules06.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="出荷"')
      .setBackground('#FFF3E0')
      .setFontColor('#E65100')
      .setRanges([range06])
      .build()
  );

  rules06.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="棚卸"')
      .setBackground('#F3E5F5')
      .setFontColor('#6A1B9A')
      .setRanges([range06])
      .build()
  );

  s06.setConditionalFormatRules(rules06);

  // 03_月販集計
  var rules03 = [];
  var range03 = s03.getRange('A2:D' + MAX_DATA_ROWS);

  rules03.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(NOT(ISBLANK($A2)),ISNUMBER($C2),$C2=0)')
      .setBackground('#FFF3E0')
      .setFontColor('#E65100')
      .setRanges([range03])
      .build()
  );

  s03.setConditionalFormatRules(rules03);

  // 08_月次履歴の条件付き書式
  var s08 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('08_月次履歴');
  if (s08 && s08.getLastRow() >= 2) {
    setHistoryConditionalFormats_(s08);
  }

  SpreadsheetApp.getUi().alert(
    '🎨 条件付き書式設定完了',
    '全シートの条件付き書式を再設定しました。\n\n' +
    '04_在庫日数_運用: 6件\n' +
    '05_差分確認: 4件\n' +
    '06_入出庫台帳: 4件\n' +
    '03_月販集計: 1件\n' +
    '08_月次履歴: 2件',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setHistoryConditionalFormats_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var range = sheet.getRange('A2:L' + lastRow);
  var rules = [];

  // 売上ゼロ商品の強調
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(NOT(ISBLANK($B2)),ISNUMBER($D2),$D2=0)')
      .setBackground('#FFF3E0')
      .setFontColor('#E65100')
      .setRanges([range])
      .build()
  );

  // 月別の交互色（視認性向上）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISODD(MONTH(DATEVALUE($A2&"/01")))')
      .setBackground('#F8F9FA')
      .setRanges([range])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

function listAllConditionalFormats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listName = '★条件付き書式一覧';
  var listSheet = ss.getSheetByName(listName);

  if (listSheet) {
    ss.deleteSheet(listSheet);
  }

  listSheet = ss.insertSheet(listName);

  var headers = ['シート名', '優先順位', '適用範囲', '条件種別', '条件値'];
  listSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  listSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#37474F')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  var allRows = [];
  var sheets = ss.getSheets();
  var i;
  var j;

  for (i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getName() === listName) continue;

    var rules = sheet.getConditionalFormatRules();

    if (!rules || rules.length === 0) {
      allRows.push([sheet.getName(), '-', '-', '条件付き書式なし', '-']);
      continue;
    }

    for (j = 0; j < rules.length; j++) {
      var rule = rules[j];
      var ranges = rule.getRanges();
      var rangeText = [];
      var k;

      for (k = 0; k < ranges.length; k++) {
        rangeText.push(ranges[k].getA1Notation());
      }

      var conditionType = '-';
      var conditionText = '-';

      var booleanCondition = rule.getBooleanCondition();
      var gradientCondition = rule.getGradientCondition();

      if (booleanCondition) {
        conditionType = String(booleanCondition.getCriteriaType());

        var criteriaValues = booleanCondition.getCriteriaValues();
        var valueText = [];
        var m;

        if (criteriaValues && criteriaValues.length > 0) {
          for (m = 0; m < criteriaValues.length; m++) {
            valueText.push(String(criteriaValues[m]));
          }
        }

        conditionText = valueText.join(', ');
      } else if (gradientCondition) {
        conditionType = 'GRADIENT';
        conditionText = 'グラデーション';
      }

      allRows.push([
        sheet.getName(),
        (j + 1),
        rangeText.join(', '),
        conditionType,
        conditionText
      ]);
    }
  }

  if (allRows.length > 0) {
    listSheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
  }

  listSheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert(
    '📋 出力完了',
    '「★条件付き書式一覧」シートに一覧を出力しました。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function resetAllConditionalFormats() {
  var sheets = getRequiredSheets_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  sheets.s03.clearConditionalFormatRules();
  sheets.s04.clearConditionalFormatRules();
  sheets.s05.clearConditionalFormatRules();
  sheets.s06.clearConditionalFormatRules();

  var s08 = ss.getSheetByName('08_月次履歴');
  if (s08) {
    s08.clearConditionalFormatRules();
  }

  SpreadsheetApp.getUi().alert(
    '🧹 リセット完了',
    '全シートの条件付き書式をクリアしました。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/*************************************************
 * ユーティリティ関数群
 *************************************************/

function parseTargetMonth_(value) {
  var dateObj = parseDateValue_(value);
  var year;
  var month;

  if (dateObj) {
    year = dateObj.getFullYear();
    month = dateObj.getMonth() + 1;
  } else {
    var str = trimString_(value);
    var match = str.match(/^(\d{4})[\/\-](\d{1,2})$/);

    if (match) {
      year = parseInt(match[1], 10);
      month = parseInt(match[2], 10);
    } else {
      var now = new Date();
      year = now.getFullYear();
      month = now.getMonth() + 1;
    }
  }

  var start = new Date(year, month - 1, 1);
  var end = new Date(year, month, 0);
  var label = year + '/' + ('0' + month).slice(-2);

  return {
    year: year,
    month: month,
    start: start,
    end: end,
    label: label
  };
}

function parseDateValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === 'string') {
    var str = trimString_(value);
    if (str === '') return null;

    var m1 = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m1) {
      return new Date(parseInt(m1[1], 10), parseInt(m1[2], 10) - 1, parseInt(m1[3], 10));
    }

    var m2 = str.match(/^(\d{4})[\/\-](\d{1,2})$/);
    if (m2) {
      return new Date(parseInt(m2[1], 10), parseInt(m2[2], 10) - 1, 1);
    }
  }

  return null;
}

function isDateInRange_(dateObj, startDate, endDate) {
  if (!dateObj || !startDate || !endDate) return false;

  var t = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()).getTime();
  var s = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()).getTime();
  var e = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()).getTime();

  return t >= s && t <= e;
}

function isYesFlag_(value) {
  var s = trimString_(value).toLowerCase();
  return s === 'yes' || s === 'y' || s === 'true' || s === '1';
}

function getMapNumber_(map, key) {
  if (!map || !map.hasOwnProperty(key)) return 0;
  return toNumber_(map[key], 0);
}

function getMapValue_(map, key, defaultValue) {
  if (!map || !map.hasOwnProperty(key)) return defaultValue;
  return map[key];
}

function toNumber_(value, defaultValue) {
  if (value === null || value === '' || typeof value === 'undefined') {
    return typeof defaultValue === 'undefined' ? 0 : defaultValue;
  }

  var n = Number(value);
  if (isNaN(n)) {
    return typeof defaultValue === 'undefined' ? 0 : defaultValue;
  }

  return n;
}

function trimString_(value) {
  if (value === null || typeof value === 'undefined') return '';
  return String(value).replace(/^\s+|\s+$/g, '');
}

function round1_(value) {
  return Math.round(toNumber_(value, 0) * 10) / 10;
}

function floor1_(value) {
  return Math.floor(toNumber_(value, 0));
}

function padRow_(row, totalCols) {
  var out = [];
  var i;

  for (i = 0; i < totalCols; i++) {
    if (i < row.length) {
      out.push(row[i]);
    } else {
      out.push('');
    }
  }

  return out;
}

function buildBlankRow_(totalCols) {
  var row = [];
  var i;

  for (i = 0; i < totalCols; i++) {
    row.push('');
  }

  return row;
}

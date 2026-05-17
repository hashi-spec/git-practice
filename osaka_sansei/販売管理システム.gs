/*************************************************
 * 在庫管理システム 完全版
 *
 * 修正・統合内容:
 * 1. 月販売数0時のゼロ除算バグ完全修正
 * 2. 棚卸リセット方式による店舗・EC併売対応
 * 3. 店舗販売推計機能（07_店舗販売推計シート）
 * 4. 月次履歴自動保存（08_月次履歴シート）
 * 5. EC+店舗全チャネル販売数による在庫日数計算
 * 6. ワンクリック棚卸機能（04シートN列）
 * 7. 滞留在庫の可視化
 *************************************************/

var MAX_DATA_ROWS = 2000;
var NO_ALERT_MESSAGE = "🎉 現在、対応が必要なアラートはありません";

/*************************************************
 * メニュー
 *************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏭 在庫管理システム")
    .addItem("🚀 完全再計算・再書き込み", "repairAll")
    .addSeparator()
    .addItem("🔄 03/04/05 データ再構築", "rebuildAllData")
    .addItem("🔄 05 表示だけ再構築", "rebuild05From04")
    .addSeparator()
    .addItem("📦 N列実在庫を一括棚卸記録", "processBatchInventory") // ★追加
    .addSeparator()
    .addItem("📦 月次履歴を手動保存", "saveCurrentMonthToHistory")
    .addItem("📊 履歴サマリー出力", "outputHistorySummary")
    .addItem("🗑️ 指定月履歴削除", "deleteSpecificMonthHistory")
    .addSeparator()
    .addItem("🎨 条件付き書式を再設定", "setAllConditionalFormats")
    .addItem("📋 条件付き書式一覧を出力", "listAllConditionalFormats")
    .addItem("🧹 条件付き書式をリセット", "resetAllConditionalFormats")
    .addToUi();
}

/*************************************************
 * 完全修復
 *************************************************/
function repairAll() {
  ensureSystemHeaders_();
  ensureStoreSalesSheet_();
  ensureMonthlyHistorySheet_();
  rebuildAllData();
  setAllConditionalFormats();

  SpreadsheetApp.getUi().alert(
    "✅ 完全再計算完了",
    "全シートの再計算・再書き込みと書式設定が完了しました。",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/*************************************************
 * 全データ再構築（月替わり自動検知付き）
 *************************************************/
function rebuildAllData() {
  var sheets = getRequiredSheets_();
  var config = getConfigValues_();

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

/*************************************************
 * 04を正本として05を再構築
 *************************************************/
function rebuild05From04() {
  var sheets = getRequiredSheets_();
  var config = getConfigValues_();
  var rows05 = build05RowsFromCurrent04_(sheets.s04, config);

  write05Rows_(sheets.s05, rows05);
  ensure05Checkboxes_(sheets.s05, countAlertRows_(rows05));

  SpreadsheetApp.flush();
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
  var storeSalesMap = collectStoreSalesFromHistory_(config);
  var existing04State = readExisting04State_(sheets.s04);

  return {
    sheets: sheets,
    config: config,
    masterItems: masterItems,
    salesAgg: salesAgg,
    ledgerAgg: ledgerAgg,
    storeSalesMap: storeSalesMap,
    existing04State: existing04State,
  };
}

/*************************************************
 * 必須ヘッダー整備
 *************************************************/
function ensureSystemHeaders_() {
  var sheets = getRequiredSheets_();
  var s03 = sheets.s03;
  var s04 = sheets.s04;
  var s05 = sheets.s05;

  // 03_月販集計（EC+店舗対応）
  s03.getRange("A1").setValue("商品コード");
  s03.getRange("B1").setValue("商品名");
  s03.getRange("C1").setValue("EC月販売数");
  s03.getRange("D1").setValue("EC月売上金額");
  s03.getRange("E1").setValue("店舗販売推計数");
  s03.getRange("F1").setValue("全チャネル販売数");

  // 04_在庫日数_運用
  s04.getRange("A1").setValue("商品コード");
  s04.getRange("B1").setValue("商品名");
  s04.getRange("C1").setValue("現在在庫数");
  s04.getRange("D1").setValue("全チャネル販売数");
  s04.getRange("E1").setValue("在庫日数");
  s04.getRange("F1").setValue("EC表示在庫");
  s04.getRange("G1").setValue("発注残数");
  s04.getRange("H1").setValue("当月仕入累計");
  s04.getRange("I1").setValue("直近入荷日");
  s04.getRange("J1").setValue("確認済");
  s04.getRange("K1").setValue("更新年月");
  s04.getRange("L1").setValue("メモ");
  s04.getRange("M1").setValue("管理指示保存");
  s04.getRange("N1").setValue("実在庫入力（棚卸）");
  s04.getRange("N1").setBackground("#FFE082");
  s04.getRange("N1").setFontWeight("bold");

  // 05_差分確認
  s05.getRange("A1").setValue("商品コード");
  s05.getRange("B1").setValue("商品名");
  s05.getRange("C1").setValue("現在在庫数");
  s05.getRange("D1").setValue("全チャネル販売数");
  s05.getRange("E1").setValue("在庫日数");
  s05.getRange("F1").setValue("EC表示在庫");
  s05.getRange("G1").setValue("発注残数");
  s05.getRange("H1").setValue("当月仕入累計");
  s05.getRange("I1").setValue("直近入荷日");
  s05.getRange("J1").setValue("確認済");
  s05.getRange("K1").setValue("更新年月");
  s05.getRange("L1").setValue("メモ");
  s05.getRange("M1").setValue("アラート理由");
  s05.getRange("N1").setValue("管理指示");
}

/*************************************************
 * 必須シート取得
 *************************************************/
function getRequiredSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheets = {
    ss: ss,
    s00: ss.getSheetByName("00_設定"),
    s01: ss.getSheetByName("01_SKUマスタ"),
    s02: ss.getSheetByName("02_売上CSV取込"),
    s03: ss.getSheetByName("03_月販集計"),
    s04: ss.getSheetByName("04_在庫日数_運用"),
    s05: ss.getSheetByName("05_差分確認"),
    s06: ss.getSheetByName("06_入出庫台帳"),
  };

  if (!sheets.s00) throw new Error("00_設定シートが見つかりません。");
  if (!sheets.s01) throw new Error("01_SKUマスタシートが見つかりません。");
  if (!sheets.s02) throw new Error("02_売上CSV取込シートが見つかりません。");
  if (!sheets.s03) throw new Error("03_月販集計シートが見つかりません。");
  if (!sheets.s04) throw new Error("04_在庫日数_運用シートが見つかりません。");
  if (!sheets.s05) throw new Error("05_差分確認シートが見つかりません。");
  if (!sheets.s06) throw new Error("06_入出庫台帳シートが見つかりません。");

  return sheets;
}

/*************************************************
 * 設定値取得
 *************************************************/
function getConfigValues_() {
  var s00 = getRequiredSheets_().s00;
  var targetMonthRaw = s00.getRange("B2").getValue();
  var monthInfo = parseTargetMonth_(targetMonthRaw);
  var stockConvertDays = toNumber_(s00.getRange("B3").getValue(), 0);
  var ecRate = toNumber_(s00.getRange("B4").getValue(), 0);
  var juuten = toNumber_(s00.getRange("B5").getValue(), 0);
  var minaoshi = toNumber_(s00.getRange("B6").getValue(), 0);
  var hacchu = toNumber_(s00.getRange("B7").getValue(), juuten);

  return {
    targetMonthRaw: targetMonthRaw,
    targetMonthLabel: monthInfo.label,
    monthStart: monthInfo.start,
    monthEnd: monthInfo.end,
    stockConvertDays: stockConvertDays,
    ecRate: ecRate,
    juuten: juuten,
    minaoshi: minaoshi,
    hacchu: hacchu,
  };
}

/*************************************************
 * 01_SKUマスタから対象商品収集
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

    if (code === "") continue;
    if (!isYesFlag_(flag)) continue;

    items.push({ code: code, name: name });
  }

  return items;
}

/*************************************************
 * 02_売上CSV取込の集計（EC売上のみ）
 *************************************************/
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

    if (code === "") continue;

    dateObj = parseDateValue_(saleDate);
    if (!isDateInRange_(dateObj, config.monthStart, config.monthEnd)) continue;

    if (!qtyMap.hasOwnProperty(code)) qtyMap[code] = 0;
    if (!amountMap.hasOwnProperty(code)) amountMap[code] = 0;

    qtyMap[code] += qty;
    amountMap[code] += amount;
  }

  return { qtyMap: qtyMap, amountMap: amountMap };
}

/*************************************************
 * 06_入出庫台帳の集計（棚卸リセット方式）
 *
 * 設計思想:
 * - 棚卸を「在庫確定スナップショット」として扱う
 * - 現在在庫 = 最新棚卸数 + 棚卸以降入荷 - 棚卸以降EC出荷
 * - 店舗販売は棚卸誤差として自動吸収
 *************************************************/
function collectLedgerAggregates_(sheet, config) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  var records = [];
  var i;

  for (i = 0; i < values.length; i++) {
    var dateValue = values[i][0];
    var code = trimString_(values[i][1]);
    var type = trimString_(values[i][3]);
    var qty = toNumber_(values[i][4], 0);

    if (code === "" || type === "" || qty === 0) continue;

    var dateObj = parseDateValue_(dateValue);
    if (!dateObj) continue;

    records.push({ date: dateObj, code: code, type: type, qty: qty });
  }

  // 日付昇順ソート（棚卸タイミング管理のため必須）
  records.sort(function (a, b) {
    return a.date.getTime() - b.date.getTime();
  });

  // 商品別管理Map
  var latestInventoryMap = {};
  var postInventoryReceiptMap = {};
  var postInventoryShipMap = {};
  var orderBalanceMap = {};
  var monthlyReceiptMap = {};
  var latestReceiptMap = {};

  for (i = 0; i < records.length; i++) {
    var rec = records[i];
    var code = rec.code;
    var type = rec.type;
    var qty = rec.qty;
    var dateObj = rec.date;
    var inTargetMonth = isDateInRange_(
      dateObj,
      config.monthStart,
      config.monthEnd,
    );

    if (!orderBalanceMap.hasOwnProperty(code)) orderBalanceMap[code] = 0;
    if (!monthlyReceiptMap.hasOwnProperty(code)) monthlyReceiptMap[code] = 0;
    if (!postInventoryReceiptMap.hasOwnProperty(code))
      postInventoryReceiptMap[code] = 0;
    if (!postInventoryShipMap.hasOwnProperty(code))
      postInventoryShipMap[code] = 0;

    if (type === "棚卸") {
      // 棚卸：最新スナップショットとして更新・以降の入出荷リセット
      if (
        !latestInventoryMap[code] ||
        dateObj.getTime() >= latestInventoryMap[code].date.getTime()
      ) {
        latestInventoryMap[code] = { date: dateObj, qty: qty };
        postInventoryReceiptMap[code] = 0;
        postInventoryShipMap[code] = 0;
      }
    } else if (type === "入荷") {
      var invDateForReceipt = latestInventoryMap[code]
        ? latestInventoryMap[code].date.getTime()
        : 0;

      if (dateObj.getTime() >= invDateForReceipt) {
        postInventoryReceiptMap[code] += qty;
      }

      orderBalanceMap[code] -= qty;

      if (inTargetMonth) monthlyReceiptMap[code] += qty;

      if (
        !latestReceiptMap[code] ||
        dateObj.getTime() > latestReceiptMap[code].getTime()
      ) {
        latestReceiptMap[code] = dateObj;
      }
    } else if (type === "出荷") {
      var invDateForShip = latestInventoryMap[code]
        ? latestInventoryMap[code].date.getTime()
        : 0;

      if (dateObj.getTime() >= invDateForShip) {
        postInventoryShipMap[code] += qty;
      }
    } else if (type === "発注") {
      orderBalanceMap[code] += qty;
    }
  }

  // 現在在庫の最終計算（棚卸リセット方式）
  var stockMap = {};
  var allCodes = [];
  var codeSet = {};

  var mapKeys = [
    Object.keys(latestInventoryMap),
    Object.keys(postInventoryReceiptMap),
    Object.keys(postInventoryShipMap),
  ];

  for (var m = 0; m < mapKeys.length; m++) {
    for (var n = 0; n < mapKeys[m].length; n++) {
      if (!codeSet[mapKeys[m][n]]) {
        codeSet[mapKeys[m][n]] = true;
        allCodes.push(mapKeys[m][n]);
      }
    }
  }

  for (i = 0; i < allCodes.length; i++) {
    var c = allCodes[i];
    var baseStock = latestInventoryMap[c] ? latestInventoryMap[c].qty : 0;
    var postReceipt = postInventoryReceiptMap[c] || 0;
    var postShip = postInventoryShipMap[c] || 0;
    stockMap[c] = baseStock + postReceipt - postShip;
  }

  return {
    stockMap: stockMap,
    orderBalanceMap: orderBalanceMap,
    monthlyReceiptMap: monthlyReceiptMap,
    latestReceiptMap: latestReceiptMap,
    latestInventoryMap: latestInventoryMap,
  };
}

/*************************************************
 * 07_店舗販売推計から対象月の推計データ収集
 *************************************************/
function collectStoreSalesFromHistory_(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s07 = ss.getSheetByName("07_店舗販売推計");
  var storeSalesMap = {};

  if (!s07 || s07.getLastRow() < 2) return storeSalesMap;

  var values = s07.getRange(2, 1, s07.getLastRow() - 1, 11).getValues();

  for (var i = 0; i < values.length; i++) {
    var inventoryDate = parseDateValue_(values[i][0]);
    var code = trimString_(values[i][1]);
    var storeSales = toNumber_(values[i][8], 0);

    if (!inventoryDate || code === "") continue;
    if (!isDateInRange_(inventoryDate, config.monthStart, config.monthEnd))
      continue;

    if (!storeSalesMap.hasOwnProperty(code)) storeSalesMap[code] = 0;
    storeSalesMap[code] += storeSales;
  }

  return storeSalesMap;
}

/*************************************************
 * 既存04の保持データ読み取り
 *************************************************/
function readExisting04State_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  var map = {};
  var i;
  var code;

  for (i = 0; i < values.length; i++) {
    code = trimString_(values[i][0]);
    if (code === "") continue;

    map[code] = {
      confirmed: values[i][9] === true,
      memo: values[i][11] === null ? "" : values[i][11],
      instruction: values[i][12] === null ? "" : values[i][12],
    };
  }

  return map;
}

/*************************************************
 * 03シート行データ生成（EC+店舗全チャネル対応）
 * A: 商品コード
 * B: 商品名
 * C: EC月販売数
 * D: EC月売上金額
 * E: 店舗販売推計数
 * F: 全チャネル販売数
 *************************************************/
function build03Rows_(ctx) {
  var rows = [];
  var i;
  var item;

  for (i = 0; i < ctx.masterItems.length; i++) {
    item = ctx.masterItems[i];

    var ecSales = getMapNumber_(ctx.salesAgg.qtyMap, item.code);
    var ecAmount = getMapNumber_(ctx.salesAgg.amountMap, item.code);
    var storeSales = getMapNumber_(ctx.storeSalesMap, item.code);
    var totalSales = ecSales + storeSales;

    rows.push([
      item.code,
      item.name,
      ecSales,
      ecAmount,
      storeSales,
      totalSales,
    ]);
  }

  return rows;
}

/*************************************************
 * 04シート行データ生成
 *
 * ★修正点:
 * - 月販売数0時の在庫日数をゼロではなく空欄に変更（ゼロ除算バグ修正）
 * - 全チャネル販売数（EC+店舗推計）で在庫日数計算
 *************************************************/
function build04Rows_(ctx) {
  var rows = [];
  var i;
  var item;
  var stock;
  var ecSales;
  var storeSales;
  var totalSales;
  var stockDays;
  var ecStock;
  var orderBalance;
  var monthReceipt;
  var latestReceipt;
  var existing;

  for (i = 0; i < ctx.masterItems.length; i++) {
    item = ctx.masterItems[i];

    stock = getMapNumber_(ctx.ledgerAgg.stockMap, item.code);
    ecSales = getMapNumber_(ctx.salesAgg.qtyMap, item.code);
    storeSales = getMapNumber_(ctx.storeSalesMap, item.code);
    totalSales = ecSales + storeSales;
    orderBalance = getMapNumber_(ctx.ledgerAgg.orderBalanceMap, item.code);
    monthReceipt = getMapNumber_(ctx.ledgerAgg.monthlyReceiptMap, item.code);
    latestReceipt = getMapValue_(ctx.ledgerAgg.latestReceiptMap, item.code, "");

    // ★修正箇所：全チャネル販売数0の場合は空欄（ゼロ除算・誤アラート防止）
    if (totalSales === 0) {
      stockDays = "";
    } else {
      stockDays = round1_((stock / totalSales) * ctx.config.stockConvertDays);
    }

    ecStock = floor1_(stock * ctx.config.ecRate);

    existing = ctx.existing04State[item.code] || {};

    rows.push([
      item.code,
      item.name,
      stock,
      totalSales,
      stockDays,
      ecStock,
      orderBalance,
      monthReceipt,
      latestReceipt || "",
      existing.confirmed === true,
      ctx.config.targetMonthLabel,
      existing.memo || "",
      existing.instruction || "",
    ]);
  }

  return rows;
}

/*************************************************
 * 04行配列から05行配列を生成
 *************************************************/
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
    stockDays = row04[4];
    orderBalance = toNumber_(row04[6], 0);

    if (confirmed) continue;
    if (!isAlertTarget_(stockDays, orderBalance, config)) continue;

    alertReason = buildAlertReason_(stockDays, orderBalance, config);

    rows05.push([
      row04[0],
      row04[1],
      row04[2],
      row04[3],
      row04[4],
      row04[5],
      row04[6],
      row04[7],
      row04[8],
      false,
      row04[10],
      row04[11],
      alertReason,
      row04[12],
    ]);
  }

  if (rows05.length === 0) {
    rows05.push([
      NO_ALERT_MESSAGE,
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
    ]);
  }

  return rows05;
}

/*************************************************
 * 現在の04シートから05を再構築
 *************************************************/
function build05RowsFromCurrent04_(sheet04, config) {
  var lastRow = Math.max(sheet04.getLastRow(), 2);
  var rowCount = Math.min(MAX_DATA_ROWS - 1, lastRow - 1);
  var values = sheet04.getRange(2, 1, rowCount, 13).getValues();
  var rows04 = [];
  var i;

  for (i = 0; i < values.length; i++) {
    if (trimString_(values[i][0]) === "") continue;
    rows04.push(values[i]);
  }

  return build05RowsFrom04Rows_(rows04, config);
}

/*************************************************
 * アラート対象判定（修正版）
 * ★在庫日数が空欄の場合は在庫日数判定をスキップ
 *************************************************/
function isAlertTarget_(stockDays, orderBalance, config) {
  var isValidStockDays = typeof stockDays === "number" && !isNaN(stockDays);

  if (isValidStockDays && stockDays <= config.juuten) return true;
  if (isValidStockDays && stockDays >= config.minaoshi) return true;
  if (orderBalance !== 0) return true;

  return false;
}

/*************************************************
 * アラート理由構築（修正版）
 *************************************************/
function buildAlertReason_(stockDays, orderBalance, config) {
  var parts = [];
  var isValidStockDays = typeof stockDays === "number" && !isNaN(stockDays);

  if (isValidStockDays && stockDays <= config.juuten) {
    parts.push("🔴 在庫不足(" + stockDays + "日)");
  }

  if (orderBalance > 0) {
    parts.push("🟡 発注残あり(" + orderBalance + "個)");
  }

  if (orderBalance < 0) {
    parts.push("⚠️ 過剰入荷(" + Math.abs(orderBalance) + "個)");
  }

  if (isValidStockDays && stockDays >= config.minaoshi) {
    parts.push("🔵 在庫過多(" + stockDays + "日)");
  }

  return parts.join(" / ");
}
/*************************************************
 * 書き込み処理
 *************************************************/
function write03Rows_(sheet, rows) {
  writeRowsBlock_(sheet, 2, 1, 6, rows);
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

  for (i = 0; i < totalRows; i++) {
    if (i < rows.length) {
      matrix.push(padRow_(rows[i], totalCols));
    } else {
      matrix.push(buildBlankRow_(totalCols));
    }
  }

  sheet.getRange(startRow, startCol, totalRows, totalCols).setValues(matrix);
}

function ensure04Checkboxes_(sheet, dataCount) {
  var rangeAll = sheet.getRange("J2:J" + MAX_DATA_ROWS);

  try {
    rangeAll.clearDataValidations();
    if (dataCount > 0) {
      sheet.getRange(2, 10, dataCount, 1).insertCheckboxes();
    }
  } catch (err) {}
}

function ensure05Checkboxes_(sheet, dataCount) {
  var rangeAll = sheet.getRange("J2:J" + MAX_DATA_ROWS);

  try {
    rangeAll.clearDataValidations();
    if (dataCount > 0) {
      sheet.getRange(2, 10, dataCount, 1).insertCheckboxes();
    }
  } catch (err) {}
}

function countAlertRows_(rows05) {
  if (!rows05 || rows05.length === 0) return 0;
  if (rows05.length === 1 && trimString_(rows05[0][0]) === NO_ALERT_MESSAGE)
    return 0;
  return rows05.length;
}

/*************************************************
 * onEdit イベント処理
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
  if (sheetName === "05_差分確認" && col === 10) {
    if (range.getValue() !== true) return;
    handle05ConfirmCheck_(row);
    ss.toast("確認済みを04シートへ反映しました", "✅ 処理完了", 3);
    return;
  }

  // 05_差分確認 L: メモ編集
  if (sheetName === "05_差分確認" && col === 12) {
    writeBack05FieldTo04_(row, col, range.getValue());
    rebuild05From04();
    return;
  }

  // 05_差分確認 N: 管理指示編集
  if (sheetName === "05_差分確認" && col === 14) {
    writeBack05FieldTo04_(row, col, range.getValue());
    rebuild05From04();
    return;
  }

  // 04_在庫日数_運用 J/L/M編集
  if (
    sheetName === "04_在庫日数_運用" &&
    (col === 10 || col === 12 || col === 13)
  ) {
    rebuild05From04();
    if (col === 10 && range.getValue() === true) {
      ss.toast("確認済み変更を05へ反映しました", "ℹ️ 更新", 2);
    }
    return;
  }

  // 00/01/02/06 の変更は全再計算
  if (
    sheetName === "00_設定" ||
    sheetName === "01_SKUマスタ" ||
    sheetName === "02_売上CSV取込" ||
    sheetName === "06_入出庫台帳"
  ) {
    rebuildAllData();
    return;
  }
}

/*************************************************
 * 04シートN列の実在庫を一括で棚卸として記録
 * 週次棚卸運用に最適化された現場重視設計
 *************************************************/
function processBatchInventory() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s04 = ss.getSheetByName("04_在庫日数_運用");
  var s06 = ss.getSheetByName("06_入出庫台帳");

  var lastRow = Math.max(s04.getLastRow(), 2);
  var values = s04.getRange(2, 1, lastRow - 1, 14).getValues();

  var inventoryRecords = [];
  var today = new Date();
  var updateCount = 0;
  var errorItems = [];

  // N列（インデックス13）の入力内容を検証・抽出
  for (var i = 0; i < values.length; i++) {
    var code = trimString_(values[i][0]);
    var name = trimString_(values[i][1]);
    var systemStock = toNumber_(values[i][2], 0);
    var actualStock = values[i][13];

    if (code === "") continue;

    if (actualStock !== "" && actualStock !== null) {
      var numValue = Number(actualStock);

      if (isNaN(numValue) || numValue < 0) {
        errorItems.push(name + "（" + actualStock + "）");
        continue;
      }

      // 異常値チェック（システム在庫の5倍を超える場合は確認）
      if (numValue > systemStock * 5 && systemStock > 0) {
        var confirm = ui.alert(
          "⚠️ 異常値確認",
          name +
            "の実在庫が " +
            numValue +
            " 個です。\n" +
            "システム在庫（" +
            systemStock +
            "個）と大きく異なりますが、正しいですか？",
          ui.ButtonSet.YES_NO,
        );
        if (confirm !== ui.Button.YES) continue;
      }

      inventoryRecords.push([today, code, name, "棚卸", numValue]);
      updateCount++;
    }
  }

  // エラー項目の報告
  if (errorItems.length > 0) {
    ui.alert(
      "⚠️ 入力エラー",
      "以下の商品で入力エラーがあります：\n" +
        errorItems.join("\n") +
        "\n\n正の数値を入力してください。",
      ui.ButtonSet.OK,
    );
    if (updateCount === 0) return;
  }

  if (updateCount === 0) {
    ui.alert(
      "ℹ️ 入力なし",
      "N列に実在庫が入力されていません。",
      ui.ButtonSet.OK,
    );
    return;
  }

  // 最終確認
  var execConfirm = ui.alert(
    "📦 一括棚卸の実行確認",
    updateCount +
      " 件の商品を棚卸として記録します。\n\n" +
      "・06_入出庫台帳への自動記録\n" +
      "・07_店舗販売推計の自動更新\n" +
      "・全システムデータの再計算\n\n" +
      "実行しますか？",
    ui.ButtonSet.YES_NO,
  );

  if (execConfirm !== ui.Button.YES) return;

  try {
    // 06台帳への一括書き込み
    var lastRow06 = Math.max(s06.getLastRow(), 1);
    s06
      .getRange(lastRow06 + 1, 1, inventoryRecords.length, 5)
      .setValues(inventoryRecords);

    // 04シートのN列をクリア
    s04.getRange(2, 14, lastRow - 1, 1).clearContent();

    // システム全体の再計算
    rebuildStoreSalesEstimate_();
    rebuildAllData();

    ui.alert(
      "✅ 棚卸処理完了",
      updateCount +
        " 件の棚卸を記録し、全データを更新しました。\n\n" +
        "・店舗販売推計が最新化されました\n" +
        "・在庫日数が全チャネル対応で再計算されました\n" +
        "・05_差分確認で最新のアラートを確認してください",
      ui.ButtonSet.OK,
    );
  } catch (error) {
    ui.alert(
      "❌ 処理エラー",
      "エラーが発生しました：" + error.toString(),
      ui.ButtonSet.OK,
    );
  }
}

/*************************************************
 * 06台帳から07_店舗販売推計を完全再構築
 * 週次棚卸運用に対応した冪等性保証処理
 *************************************************/
function rebuildStoreSalesEstimate_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s06 = ss.getSheetByName("06_入出庫台帳");
  var s07 = ensureStoreSalesSheet_();

  var lastRow06 = Math.max(s06.getLastRow(), 2);
  var values06 = s06.getRange(2, 1, lastRow06 - 1, 5).getValues();

  var recordsByItem = {};

  // 商品別に履歴を分類・整理
  for (var i = 0; i < values06.length; i++) {
    var dateObj = parseDateValue_(values06[i][0]);
    var code = trimString_(values06[i][1]);
    var type = trimString_(values06[i][3]);
    var qty = toNumber_(values06[i][4], 0);

    if (!dateObj || code === "" || type === "" || qty === 0) continue;

    if (!recordsByItem[code]) recordsByItem[code] = [];
    recordsByItem[code].push({ date: dateObj, type: type, qty: qty });
  }

  var rows07 = [];
  var codes = Object.keys(recordsByItem);

  // 各商品の棚卸間隔で店舗販売推計を計算
  for (var c = 0; c < codes.length; c++) {
    var code = codes[c];
    var records = recordsByItem[code];

    // 時系列順（古い順）にソート
    records.sort(function (a, b) {
      return a.date.getTime() - b.date.getTime();
    });

    var prevInventory = null;
    var prevDate = null;
    var periodReceipt = 0;
    var periodEcShip = 0;

    for (var r = 0; r < records.length; r++) {
      var rec = records[r];

      if (rec.type === "入荷") {
        if (prevDate) periodReceipt += rec.qty;
      } else if (rec.type === "出荷") {
        if (prevDate) periodEcShip += rec.qty;
      } else if (rec.type === "棚卸") {
        // 2回目以降の棚卸で推計計算
        if (prevDate !== null) {
          var theoreticalStock = prevInventory + periodReceipt - periodEcShip;
          var storeSales = Math.max(0, theoreticalStock - rec.qty);
          var periodDays = Math.round(
            (rec.date.getTime() - prevDate.getTime()) / (1000 * 60 * 60 * 24),
          );
          var dailyStoreSales =
            periodDays > 0 ? round1_(storeSales / periodDays) : 0;

          var itemName = getItemNameFromMaster_(code);

          rows07.push([
            rec.date,
            code,
            itemName,
            prevInventory,
            prevDate,
            periodReceipt,
            periodEcShip,
            rec.qty,
            storeSales,
            periodDays,
            dailyStoreSales,
          ]);
        }

        // 次の期間の開始点として設定
        prevInventory = rec.qty;
        prevDate = rec.date;
        periodReceipt = 0;
        periodEcShip = 0;
      }
    }
  }

  // 07シートを完全クリア後に再構築
  var lastRow07 = Math.max(s07.getLastRow(), 1);
  if (lastRow07 > 1) {
    s07.getRange(2, 1, lastRow07 - 1, 11).clearContent();
  }

  if (rows07.length > 0) {
    // 最新順にソートして見やすく
    rows07.sort(function (a, b) {
      return b[0].getTime() - a[0].getTime();
    });
    s07.getRange(2, 1, rows07.length, 11).setValues(rows07);
  }
}

/*************************************************
 * 01_SKUマスタから商品名を取得
 *************************************************/
function getItemNameFromMaster_(itemCode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s01 = ss.getSheetByName("01_SKUマスタ");
  if (!s01) return "";

  var finder = s01
    .getRange("B2:B" + s01.getLastRow())
    .createTextFinder(itemCode)
    .matchEntireCell(true)
    .findNext();

  if (!finder) return "";
  return trimString_(s01.getRange(finder.getRow(), 3).getValue());
}

// ↑ ここまでが追加コード

/*************************************************
 * 04シート N列：ワンクリック棚卸処理
 *************************************************/
function processQuickInventory_(range, row, actualStock) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s04 = ss.getSheetByName("04_在庫日数_運用");
  var s06 = ss.getSheetByName("06_入出庫台帳");
  var s07 = ensureStoreSalesSheet_();

  var itemCode = trimString_(s04.getRange(row, 1).getDisplayValue());
  var itemName = trimString_(s04.getRange(row, 2).getDisplayValue());
  var systemStock = toNumber_(s04.getRange(row, 3).getValue(), 0);

  if (itemCode === "") {
    range.clearContent();
    return;
  }

  var today = new Date();

  // 店舗販売数推計計算
  var estimate = calculateStoreSalesEstimate_(itemCode, actualStock, today);

  // 06台帳に棚卸記録
  var lastRow06 = Math.max(s06.getLastRow(), 1);
  s06
    .getRange(lastRow06 + 1, 1, 1, 5)
    .setValues([[today, itemCode, itemName, "棚卸", actualStock]]);

  // 07推計シートに記録（前回棚卸がある場合のみ）
  if (estimate && estimate.isValid) {
    var lastRow07 = Math.max(s07.getLastRow(), 1);
    s07
      .getRange(lastRow07 + 1, 1, 1, 11)
      .setValues([
        [
          today,
          itemCode,
          itemName,
          estimate.prevInventory,
          estimate.prevDate,
          estimate.periodReceipt,
          estimate.periodEcShip,
          actualStock,
          estimate.storeSales,
          estimate.periodDays,
          estimate.dailyStoreSales,
        ],
      ]);
  }

  // 入力セルをクリア
  range.clearContent();

  // ユーザー通知
  var diff = systemStock - actualStock;
  var message =
    itemName +
    " の棚卸を記録しました\n" +
    "システム在庫: " +
    systemStock +
    " → 実在庫: " +
    actualStock;

  if (diff !== 0) {
    message +=
      "\n誤差: " +
      Math.abs(diff) +
      "個（" +
      (diff > 0 ? "店舗販売・ロス等" : "入荷漏れ・返品等") +
      "）";
  }

  if (estimate && estimate.isValid) {
    message +=
      "\n店舗販売推計: " +
      estimate.storeSales +
      "個 / " +
      estimate.periodDays +
      "日間";
  }

  ss.toast(message, "📦 棚卸完了", 6);

  SpreadsheetApp.flush();
  Utilities.sleep(300);
  rebuildAllData();
}

/*************************************************
 * 店舗販売数推計計算
 *************************************************/
function calculateStoreSalesEstimate_(code, newInventory, newDate) {
  var s06 =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("06_入出庫台帳");
  var lastRow = Math.max(s06.getLastRow(), 2);
  var values = s06.getRange(2, 1, lastRow - 1, 5).getValues();

  var records = [];

  for (var i = 0; i < values.length; i++) {
    var recCode = trimString_(values[i][1]);
    if (recCode !== code) continue;

    var recDate = parseDateValue_(values[i][0]);
    var recType = trimString_(values[i][3]);
    var recQty = toNumber_(values[i][4], 0);

    if (!recDate || recDate.getTime() >= newDate.getTime()) continue;

    records.push({ date: recDate, type: recType, qty: recQty });
  }

  // 日付降順ソート（最新棚卸を先頭に）
  records.sort(function (a, b) {
    return b.date.getTime() - a.date.getTime();
  });

  // 直前棚卸を特定
  var prevInventory = 0;
  var prevDate = null;
  var foundPrev = false;

  for (var j = 0; j < records.length; j++) {
    if (records[j].type === "棚卸") {
      prevInventory = records[j].qty;
      prevDate = records[j].date;
      foundPrev = true;
      break;
    }
  }

  if (!foundPrev) {
    return { isValid: false, reason: "前回棚卸データなし" };
  }

  // 期間中の入荷・EC出荷を集計
  var periodReceipt = 0;
  var periodEcShip = 0;

  for (var k = 0; k < records.length; k++) {
    var rec = records[k];
    if (rec.date.getTime() <= prevDate.getTime()) continue;
    if (rec.type === "入荷") periodReceipt += rec.qty;
    if (rec.type === "出荷") periodEcShip += rec.qty;
  }

  // 店舗販売数推計
  var theoreticalStock = prevInventory + periodReceipt - periodEcShip;
  var storeSales = Math.max(0, theoreticalStock - newInventory);

  // 期間日数・日平均
  var periodDays = Math.round(
    (newDate.getTime() - prevDate.getTime()) / (1000 * 60 * 60 * 24),
  );
  var dailyStoreSales = periodDays > 0 ? round1_(storeSales / periodDays) : 0;

  return {
    isValid: true,
    prevInventory: prevInventory,
    prevDate: prevDate,
    periodReceipt: periodReceipt,
    periodEcShip: periodEcShip,
    storeSales: storeSales,
    periodDays: periodDays,
    dailyStoreSales: dailyStoreSales,
  };
}

/*************************************************
 * 05確認済みチェック処理
 *************************************************/
function handle05ConfirmCheck_(row) {
  var sheets = getRequiredSheets_();
  var s04 = sheets.s04;
  var s05 = sheets.s05;

  var itemCode = trimString_(s05.getRange(row, 1).getDisplayValue());

  s05.getRange(row, 10).setValue(false);

  if (itemCode === "" || itemCode === NO_ALERT_MESSAGE) {
    rebuild05From04();
    return;
  }

  var finder = s04
    .getRange("A2:A" + MAX_DATA_ROWS)
    .createTextFinder(itemCode)
    .matchEntireCell(true)
    .findNext();

  if (!finder) {
    rebuild05From04();
    return;
  }

  s04.getRange(finder.getRow(), 10).setValue(true);

  SpreadsheetApp.flush();
  Utilities.sleep(200);

  rebuild05From04();
}

/*************************************************
 * 05 L/N を 04固定列へ書き戻す
 *************************************************/
function writeBack05FieldTo04_(row, col, value) {
  var sheets = getRequiredSheets_();
  var s04 = sheets.s04;
  var s05 = sheets.s05;

  var itemCode = trimString_(s05.getRange(row, 1).getDisplayValue());
  if (itemCode === "" || itemCode === NO_ALERT_MESSAGE) return;

  var finder = s04
    .getRange("A2:A" + MAX_DATA_ROWS)
    .createTextFinder(itemCode)
    .matchEntireCell(true)
    .findNext();
  if (!finder) return;

  var targetRow = finder.getRow();

  if (col === 12) {
    s04.getRange(targetRow, 12).setValue(value);
    return;
  }

  if (col === 14) {
    s04.getRange(targetRow, 13).setValue(value);
    return;
  }
}

/*************************************************
 * 07_店舗販売推計シートの初期化
 *************************************************/
function ensureStoreSalesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s07 = ss.getSheetByName("07_店舗販売推計");

  if (!s07) {
    s07 = ss.insertSheet("07_店舗販売推計");

    var headers = [
      "棚卸日", // A
      "商品コード", // B
      "商品名", // C
      "前回棚卸数", // D
      "前回棚卸日", // E
      "期間中入荷数", // F
      "期間中EC出荷数", // G
      "今回棚卸数", // H
      "店舗販売推計数", // I
      "期間日数", // J
      "日平均店舗販売数", // K
    ];

    s07.getRange(1, 1, 1, headers.length).setValues([headers]);
    s07
      .getRange(1, 1, 1, headers.length)
      .setBackground("#1B5E20")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");

    s07.setFrozenRows(1);
    s07.setFrozenColumns(3);
    s07.autoResizeColumns(1, headers.length);
  }

  return s07;
}

/*************************************************
 * 08_月次履歴シートの初期化
 *************************************************/
function ensureMonthlyHistorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s08 = ss.getSheetByName("08_月次履歴");

  if (!s08) {
    s08 = ss.insertSheet("08_月次履歴");

    var headers = [
      "記録年月", // A
      "商品コード", // B
      "商品名", // C
      "EC月販売数", // D
      "EC月売上金額", // E
      "店舗販売推計数", // F
      "全チャネル販売数", // G
      "月末在庫数", // H
      "在庫日数", // I
      "EC表示在庫", // J
      "発注残数", // K
      "当月仕入累計", // L
      "直近入荷日", // M
      "記録日時", // N
    ];

    s08.getRange(1, 1, 1, headers.length).setValues([headers]);
    s08
      .getRange(1, 1, 1, headers.length)
      .setBackground("#1A237E")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");

    s08.setFrozenColumns(2);
    s08.setFrozenRows(1);
    s08.autoResizeColumns(1, headers.length);
  }

  return s08;
}

/*************************************************
 * 月次履歴：手動保存
 *************************************************/
function saveCurrentMonthToHistory() {
  var ui = SpreadsheetApp.getUi();
  var ctx = buildSystemContext_();
  var targetLabel = ctx.config.targetMonthLabel;
  var s08 = ensureMonthlyHistorySheet_();

  var duplicateInfo = checkDuplicateMonth_(s08, targetLabel);

  if (duplicateInfo.exists) {
    var response = ui.alert(
      "⚠️ 重複データ確認",
      targetLabel +
        " のデータが既に " +
        duplicateInfo.count +
        " 件存在します。\n上書きしますか？",
      ui.ButtonSet.YES_NO,
    );

    if (response !== ui.Button.YES) {
      ui.alert(
        "ℹ️ キャンセル",
        "履歴保存をキャンセルしました。",
        ui.ButtonSet.OK,
      );
      return;
    }

    deleteMonthData_(s08, targetLabel);
  }

  var historyRows = buildHistoryDataRows_(ctx, targetLabel);

  if (historyRows.length === 0) {
    ui.alert(
      "⚠️ データなし",
      "保存対象データが見つかりません。",
      ui.ButtonSet.OK,
    );
    return;
  }

  var lastRow = Math.max(s08.getLastRow(), 1);
  s08.getRange(lastRow + 1, 1, historyRows.length, 14).setValues(historyRows);

  SpreadsheetApp.flush();

  ui.alert(
    "✅ 履歴保存完了",
    targetLabel + " のデータを " + historyRows.length + " 件保存しました。",
    ui.ButtonSet.OK,
  );
}

/*************************************************
 * 月替わり自動アーカイブ
 *************************************************/
function autoArchiveIfMonthChanged_(sheets, config) {
  var s04 = sheets.s04;
  var currentMonthInSheet = trimString_(s04.getRange("K2").getDisplayValue());
  var newTargetMonth = config.targetMonthLabel;

  if (!currentMonthInSheet || currentMonthInSheet === newTargetMonth) return;

  var s08 = ensureMonthlyHistorySheet_();

  var duplicateInfo = checkDuplicateMonth_(s08, currentMonthInSheet);
  if (duplicateInfo.exists) {
    deleteMonthData_(s08, currentMonthInSheet);
  }

  var historyRows = buildHistoryFromCurrentSheets_(sheets, currentMonthInSheet);

  if (historyRows.length > 0) {
    var lastRow = Math.max(s08.getLastRow(), 1);
    s08.getRange(lastRow + 1, 1, historyRows.length, 14).setValues(historyRows);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      currentMonthInSheet +
        " のデータを自動で履歴保存しました（" +
        historyRows.length +
        " 件）",
      "📦 自動アーカイブ完了",
      5,
    );
  }
}

/*************************************************
 * 履歴データ行の構築（EC+店舗全チャネル対応）
 *************************************************/
function buildHistoryDataRows_(ctx, targetLabel) {
  var rows = [];
  var now = new Date();

  for (var i = 0; i < ctx.masterItems.length; i++) {
    var item = ctx.masterItems[i];

    var ecSales = getMapNumber_(ctx.salesAgg.qtyMap, item.code);
    var ecAmount = getMapNumber_(ctx.salesAgg.amountMap, item.code);
    var storeSales = getMapNumber_(ctx.storeSalesMap, item.code);
    var totalSales = ecSales + storeSales;
    var stock = getMapNumber_(ctx.ledgerAgg.stockMap, item.code);
    var orderBalance = getMapNumber_(ctx.ledgerAgg.orderBalanceMap, item.code);
    var monthReceipt = getMapNumber_(
      ctx.ledgerAgg.monthlyReceiptMap,
      item.code,
    );
    var latestReceipt = getMapValue_(
      ctx.ledgerAgg.latestReceiptMap,
      item.code,
      "",
    );

    var stockDays =
      totalSales === 0
        ? ""
        : round1_((stock / totalSales) * ctx.config.stockConvertDays);
    var ecStock = floor1_(stock * ctx.config.ecRate);

    rows.push([
      targetLabel,
      item.code,
      item.name,
      ecSales,
      ecAmount,
      storeSales,
      totalSales,
      stock,
      stockDays,
      ecStock,
      orderBalance,
      monthReceipt,
      latestReceipt || "",
      now,
    ]);
  }

  return rows;
}

/*************************************************
 * 現在の03/04シートから履歴行を構築
 *************************************************/
function buildHistoryFromCurrentSheets_(sheets, monthLabel) {
  var s03 = sheets.s03;
  var s04 = sheets.s04;
  var now = new Date();

  var lastRow03 = Math.max(s03.getLastRow(), 2);
  var values03 = s03
    .getRange(2, 1, Math.min(MAX_DATA_ROWS - 1, lastRow03 - 1), 6)
    .getValues();

  var lastRow04 = Math.max(s04.getLastRow(), 2);
  var values04 = s04
    .getRange(2, 1, Math.min(MAX_DATA_ROWS - 1, lastRow04 - 1), 13)
    .getValues();

  var stockDataMap = {};
  for (var i = 0; i < values04.length; i++) {
    var code = trimString_(values04[i][0]);
    if (code !== "") {
      stockDataMap[code] = {
        stock: values04[i][2],
        stockDays: values04[i][4],
        ecStock: values04[i][5],
        orderBalance: values04[i][6],
        monthReceipt: values04[i][7],
        latestReceipt: values04[i][8],
      };
    }
  }

  var historyRows = [];

  for (var j = 0; j < values03.length; j++) {
    var itemCode = trimString_(values03[j][0]);
    if (itemCode === "") continue;

    var sd = stockDataMap[itemCode] || {
      stock: 0,
      stockDays: "",
      ecStock: 0,
      orderBalance: 0,
      monthReceipt: 0,
      latestReceipt: "",
    };

    historyRows.push([
      monthLabel,
      itemCode,
      values03[j][1],
      values03[j][2], // EC月販売数
      values03[j][3], // EC月売上金額
      values03[j][4], // 店舗販売推計数
      values03[j][5], // 全チャネル販売数
      sd.stock,
      sd.stockDays,
      sd.ecStock,
      sd.orderBalance,
      sd.monthReceipt,
      sd.latestReceipt,
      now,
    ]);
  }

  return historyRows;
}

/*************************************************
 * 重複チェック・データ削除
 *************************************************/
function checkDuplicateMonth_(sheet, targetLabel) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { exists: false, count: 0 };

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var count = 0;

  for (var i = 0; i < values.length; i++) {
    if (trimString_(values[i][0]) === targetLabel) count++;
  }

  return { exists: count > 0, count: count };
}

function deleteMonthData_(sheet, targetLabel) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    if (trimString_(values[i][0]) === targetLabel) {
      sheet.deleteRow(i + 2);
    }
  }
}

function deleteSpecificMonthHistory() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    "🗑️ 指定月履歴削除",
    '削除したい記録年月を "YYYY/MM" 形式で入力してください（例: 2026/03）',
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var targetLabel = trimString_(response.getResponseText());

  if (!targetLabel.match(/^\d{4}\/\d{2}$/)) {
    ui.alert(
      "⚠️ フォーマット不正",
      "YYYY/MM 形式で入力してください。",
      ui.ButtonSet.OK,
    );
    return;
  }

  var s08 = ensureMonthlyHistorySheet_();
  var duplicateInfo = checkDuplicateMonth_(s08, targetLabel);

  if (!duplicateInfo.exists) {
    ui.alert(
      "ℹ️ 対象なし",
      targetLabel + " の履歴データは存在しません。",
      ui.ButtonSet.OK,
    );
    return;
  }

  var confirm = ui.alert(
    "⚠️ 最終確認",
    targetLabel +
      " の履歴データ " +
      duplicateInfo.count +
      " 行を削除します。\nよろしいですか？",
    ui.ButtonSet.YES_NO,
  );

  if (confirm !== ui.Button.YES) return;

  deleteMonthData_(s08, targetLabel);
  ui.alert(
    "✅ 削除完了",
    targetLabel + " の履歴データを削除しました。",
    ui.ButtonSet.OK,
  );
}

/*************************************************
 * 履歴サマリー出力
 *************************************************/
function outputHistorySummary() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var s08 = ss.getSheetByName("08_月次履歴");
  if (!s08 || s08.getLastRow() < 2) {
    ui.alert(
      "⚠️ データなし",
      "08_月次履歴にデータがありません。",
      ui.ButtonSet.OK,
    );
    return;
  }

  var summaryName = "09_履歴サマリー";
  var s09 = ss.getSheetByName(summaryName);
  if (s09) ss.deleteSheet(s09);
  s09 = ss.insertSheet(summaryName);

  var lastRow = s08.getLastRow();
  var values = s08.getRange(2, 1, lastRow - 1, 14).getValues();

  var monthSummary = {};

  for (var i = 0; i < values.length; i++) {
    var monthLabel = trimString_(values[i][0]);
    if (monthLabel === "") continue;

    var ecSales = toNumber_(values[i][3], 0);
    var ecAmount = toNumber_(values[i][4], 0);
    var storeSales = toNumber_(values[i][5], 0);
    var totalSales = toNumber_(values[i][6], 0);
    var stock = toNumber_(values[i][7], 0);

    if (!monthSummary[monthLabel]) {
      monthSummary[monthLabel] = {
        ecQty: 0,
        ecAmount: 0,
        storeSales: 0,
        totalSales: 0,
        totalStock: 0,
        itemCount: 0,
      };
    }

    monthSummary[monthLabel].ecQty += ecSales;
    monthSummary[monthLabel].ecAmount += ecAmount;
    monthSummary[monthLabel].storeSales += storeSales;
    monthSummary[monthLabel].totalSales += totalSales;
    monthSummary[monthLabel].totalStock += stock;
    monthSummary[monthLabel].itemCount += 1;
  }

  var months = Object.keys(monthSummary).sort();

  if (months.length === 0) {
    ui.alert(
      "⚠️ データなし",
      "集計対象データが見つかりません。",
      ui.ButtonSet.OK,
    );
    ss.deleteSheet(s09);
    return;
  }

  var headers = [
    "記録年月",
    "EC販売数",
    "EC売上金額",
    "店舗販売推計数",
    "全チャネル販売数",
    "平均在庫数",
    "商品数",
  ];

  s09.getRange(1, 1, 1, headers.length).setValues([headers]);

  var summaryData = [];

  for (var j = 0; j < months.length; j++) {
    var d = monthSummary[months[j]];
    var avgStock = d.itemCount > 0 ? Math.round(d.totalStock / d.itemCount) : 0;

    summaryData.push([
      months[j],
      d.ecQty,
      d.ecAmount,
      d.storeSales,
      d.totalSales,
      avgStock,
      d.itemCount,
    ]);
  }

  s09.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);

  s09
    .getRange(1, 1, 1, headers.length)
    .setBackground("#004D40")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  s09.getRange(2, 3, summaryData.length, 1).setNumberFormat("#,##0");
  s09.autoResizeColumns(1, headers.length);
  s09.setFrozenRows(1);

  ui.alert(
    "📊 サマリー出力完了",
    "09_履歴サマリーシートに出力しました。\n\n" +
      "対象期間: " +
      months[0] +
      " 〜 " +
      months[months.length - 1] +
      "\n" +
      "対象月数: " +
      months.length +
      " ヶ月",
    ui.ButtonSet.OK,
  );
}

/*************************************************
 * 条件付き書式 全再設定
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
  var range04 = s04.getRange("A2:N" + MAX_DATA_ROWS);

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$J2=TRUE")
      .setBackground("#F8F9FA")
      .setFontColor("#6C757D")
      .setRanges([range04])
      .build(),
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(ISNUMBER($E2),$E2<=" + juuten + ")")
      .setBackground("#FFEBEE")
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([range04])
      .build(),
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(ISNUMBER($G2),$G2<0)")
      .setBackground("#FFEBEE")
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([range04])
      .build(),
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(ISNUMBER($G2),$G2>0)")
      .setBackground("#FFFDE7")
      .setFontColor("#F57F17")
      .setRanges([range04])
      .build(),
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        "=AND(ISNUMBER($E2),$E2<=" + hacchu + ",$E2>" + juuten + ")",
      )
      .setBackground("#FFFDE7")
      .setFontColor("#F57F17")
      .setRanges([range04])
      .build(),
  );

  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(ISNUMBER($E2),$E2>=" + minaoshi + ")")
      .setBackground("#FFF3E0")
      .setFontColor("#E65100")
      .setRanges([range04])
      .build(),
  );

  // 販売実績なし・在庫あり（滞留在庫）
  rules04.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND($C2>0,$D2=0)")
      .setBackground("#F3E5F5")
      .setFontColor("#6A1B9A")
      .setRanges([range04])
      .build(),
  );

  s04.setConditionalFormatRules(rules04);

  // 05_差分確認
  var rules05 = [];
  var range05 = s05.getRange("A2:N" + MAX_DATA_ROWS);

  rules05.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(ISNUMBER($E2),$E2<=" + juuten + ")")
      .setBackground("#FFEBEE")
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([range05])
      .build(),
  );

  rules05.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(ISNUMBER($G2),$G2<0)")
      .setBackground("#FFEBEE")
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([range05])
      .build(),
  );

  rules05.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(ISNUMBER($G2),$G2>0)")
      .setBackground("#FFFDE7")
      .setFontColor("#F57F17")
      .setRanges([range05])
      .build(),
  );

  rules05.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(ISNUMBER($E2),$E2>=" + minaoshi + ")")
      .setBackground("#E3F2FD")
      .setFontColor("#1565C0")
      .setRanges([range05])
      .build(),
  );

  s05.setConditionalFormatRules(rules05);

  // 06_入出庫台帳
  var rules06 = [];
  var range06 = s06.getRange("A2:G" + MAX_DATA_ROWS);

  rules06.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="発注"')
      .setBackground("#E3F2FD")
      .setFontColor("#1565C0")
      .setRanges([range06])
      .build(),
  );

  rules06.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="入荷"')
      .setBackground("#E8F5E9")
      .setFontColor("#2E7D32")
      .setRanges([range06])
      .build(),
  );

  rules06.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="出荷"')
      .setBackground("#FFF3E0")
      .setFontColor("#E65100")
      .setRanges([range06])
      .build(),
  );

  rules06.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="棚卸"')
      .setBackground("#F3E5F5")
      .setFontColor("#6A1B9A")
      .setRanges([range06])
      .build(),
  );

  s06.setConditionalFormatRules(rules06);

  // 03_月販集計
  var rules03 = [];
  var range03 = s03.getRange("A2:F" + MAX_DATA_ROWS);

  rules03.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(NOT(ISBLANK($A2)),ISNUMBER($F2),$F2=0)")
      .setBackground("#FFF3E0")
      .setFontColor("#E65100")
      .setRanges([range03])
      .build(),
  );

  s03.setConditionalFormatRules(rules03);

  SpreadsheetApp.getUi().alert(
    "🎨 条件付き書式設定完了",
    "全シートの条件付き書式を再設定しました。\n\n" +
      "04_在庫日数_運用: 7件（滞留在庫検知追加）\n" +
      "05_差分確認: 4件\n" +
      "06_入出庫台帳: 4件\n" +
      "03_月販集計: 1件\n\n" +
      "重点管理閾値: " +
      juuten +
      "日\n" +
      "見直し閾値: " +
      minaoshi +
      "日\n" +
      "発注推奨閾値: " +
      hacchu +
      "日",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/*************************************************
 * 条件付き書式一覧出力・リセット
 *************************************************/
function listAllConditionalFormats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listName = "★条件付き書式一覧";
  var listSheet = ss.getSheetByName(listName);

  if (listSheet) ss.deleteSheet(listSheet);
  listSheet = ss.insertSheet(listName);

  var headers = ["シート名", "優先順位", "適用範囲", "条件種別", "条件値"];
  listSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  listSheet
    .getRange(1, 1, 1, headers.length)
    .setBackground("#37474F")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  var allRows = [];
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getName() === listName) continue;

    var rules = sheet.getConditionalFormatRules();

    if (!rules || rules.length === 0) {
      allRows.push([sheet.getName(), "-", "-", "条件付き書式なし", "-"]);
      continue;
    }

    for (var j = 0; j < rules.length; j++) {
      var rule = rules[j];
      var ranges = rule.getRanges();
      var rangeText = [];

      for (var k = 0; k < ranges.length; k++) {
        rangeText.push(ranges[k].getA1Notation());
      }

      var conditionType = "-";
      var conditionText = "-";
      var booleanCondition = rule.getBooleanCondition();
      var gradientCondition = rule.getGradientCondition();

      if (booleanCondition) {
        conditionType = String(booleanCondition.getCriteriaType());
        var criteriaValues = booleanCondition.getCriteriaValues();
        var valueText = [];

        if (criteriaValues && criteriaValues.length > 0) {
          for (var m = 0; m < criteriaValues.length; m++) {
            valueText.push(String(criteriaValues[m]));
          }
        }

        conditionText = valueText.join(", ");
      } else if (gradientCondition) {
        conditionType = "GRADIENT";
        conditionText = "グラデーション";
      }

      allRows.push([
        sheet.getName(),
        j + 1,
        rangeText.join(", "),
        conditionType,
        conditionText,
      ]);
    }
  }

  if (allRows.length > 0) {
    listSheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
  }

  listSheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert(
    "📋 出力完了",
    "「★条件付き書式一覧」シートに出力しました。",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function resetAllConditionalFormats() {
  var sheets = getRequiredSheets_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  sheets.s03.clearConditionalFormatRules();
  sheets.s04.clearConditionalFormatRules();
  sheets.s05.clearConditionalFormatRules();
  sheets.s06.clearConditionalFormatRules();

  var s08 = ss.getSheetByName("08_月次履歴");
  if (s08) s08.clearConditionalFormatRules();

  SpreadsheetApp.getUi().alert(
    "🧹 リセット完了",
    "全シートの条件付き書式をクリアしました。",
    SpreadsheetApp.getUi().ButtonSet.OK,
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
  var label = year + "/" + ("0" + month).slice(-2);

  return { year: year, month: month, start: start, end: end, label: label };
}

function parseDateValue_(value) {
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value.getTime())
  ) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === "string") {
    var str = trimString_(value);
    if (str === "") return null;

    var m1 = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m1) {
      return new Date(
        parseInt(m1[1], 10),
        parseInt(m1[2], 10) - 1,
        parseInt(m1[3], 10),
      );
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

  var t = new Date(
    dateObj.getFullYear(),
    dateObj.getMonth(),
    dateObj.getDate(),
  ).getTime();
  var s = new Date(
    startDate.getFullYear(),
    startDate.getMonth(),
    startDate.getDate(),
  ).getTime();
  var e = new Date(
    endDate.getFullYear(),
    endDate.getMonth(),
    endDate.getDate(),
  ).getTime();

  return t >= s && t <= e;
}

function isYesFlag_(value) {
  var s = trimString_(value).toLowerCase();
  return s === "yes" || s === "y" || s === "true" || s === "1";
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
  if (value === null || value === "" || typeof value === "undefined") {
    return typeof defaultValue === "undefined" ? 0 : defaultValue;
  }

  var n = Number(value);
  if (isNaN(n)) {
    return typeof defaultValue === "undefined" ? 0 : defaultValue;
  }

  return n;
}

function trimString_(value) {
  if (value === null || typeof value === "undefined") return "";
  return String(value).replace(/^\s+|\s+$/g, "");
}

function round1_(value) {
  return Math.round(toNumber_(value, 0) * 10) / 10;
}

function floor1_(value) {
  return Math.floor(toNumber_(value, 0));
}

function padRow_(row, totalCols) {
  var out = [];
  for (var i = 0; i < totalCols; i++) {
    out.push(i < row.length ? row[i] : "");
  }
  return out;
}

function buildBlankRow_(totalCols) {
  var row = [];
  for (var i = 0; i < totalCols; i++) {
    row.push("");
  }
  return row;
}

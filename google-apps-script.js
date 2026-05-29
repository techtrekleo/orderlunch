// Google Apps Script - 午餐訂購系統後端
// 每天的訂單存在各自的 Sheet (以日期命名，如 "2026-05-29")
// 菜單圖片網址存在 "設定" Sheet

// 你的 Google 試算表 ID（從試算表網址取得）
// 網址格式：https://docs.google.com/spreadsheets/d/【這段就是ID】/edit
const SPREADSHEET_ID = '1xVKp06mLHu50PppA8Baxm8c6uzyZ5dD52agCzYOw5Kc';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// 第一次使用前，在編輯器手動執行這個函式一次，跳出授權視窗後同意，即可解決權限問題
function 授權用() {
  const ss = getSpreadsheet();
  Logger.log(ss.getName());
}

function doGet(e) {
  const action = e.parameter.action;
  const date = e.parameter.date;

  if (action === 'getData') {
    return ContentService.createTextOutput(JSON.stringify(getData(date)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  const data = body.data;
  const date = body.date;

  switch (action) {
    case 'addOrder':
      addOrder(date, data);
      break;
    case 'deleteOrder':
      deleteOrder(date, data.id);
      break;
    case 'updateMenu':
      updateMenu(date, data.url, data.side);
      break;
    case 'clearOrders':
      clearOrders(date);
      break;
    case 'clearAll':
      clearAll(date);
      break;
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- 取得指定日期的所有資料 ---
function getData(date) {
  const ss = getSpreadsheet();

  // 讀取該日期的菜單圖片
  const settingsSheet = getOrCreateSettingsSheet(ss);
  const menuFront = getMenuUrl(settingsSheet, date, 'front');
  const menuBack = getMenuUrl(settingsSheet, date, 'back');

  // 讀取該日期的訂單
  const orders = [];
  const orderSheet = ss.getSheetByName(date);
  if (orderSheet) {
    const lastRow = orderSheet.getLastRow();
    if (lastRow >= 2) {
      const dataRange = orderSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      dataRange.forEach(function(row) {
        if (row[0]) {
          orders.push({
            id: String(row[0]),
            name: row[1],
            itemName: row[2],
            price: Number(row[3])
          });
        }
      });
    }
  }

  return {
    menuImageUrlFront: menuFront,
    menuImageUrlBack: menuBack,
    orders: orders
  };
}

// --- 新增訂單 ---
function addOrder(date, order) {
  const ss = getSpreadsheet();
  const sheet = getOrCreateOrderSheet(ss, date);
  sheet.appendRow([order.id, order.name, order.itemName, order.price]);
}

// --- 刪除訂單 ---
function deleteOrder(date, orderId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(date);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(orderId)) {
      sheet.deleteRow(i + 2);
      return;
    }
  }
}

// --- 更新菜單圖片 ---
function updateMenu(date, url, side) {
  const ss = getSpreadsheet();
  const settingsSheet = getOrCreateSettingsSheet(ss);
  setMenuUrl(settingsSheet, date, side || 'front', url);
}

// --- 清除指定日期訂單 ---
function clearOrders(date) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(date);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.deleteRows(2, lastRow - 1);
    }
  }
}

// --- 清除所有資料(含菜單) ---
function clearAll(date) {
  const ss = getSpreadsheet();

  // 刪除該日期的訂單 sheet
  const sheet = ss.getSheetByName(date);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.deleteRows(2, lastRow - 1);
    }
  }

  // 清除該日期的菜單
  const settingsSheet = getOrCreateSettingsSheet(ss);
  clearMenuUrls(settingsSheet, date);
}

// === 輔助函式 ===

// Google Sheets 會把日期字串自動轉成 Date 物件，需要統一格式再比對
function normalizeDate(value) {
  if (value instanceof Date) {
    var y = value.getFullYear();
    var m = String(value.getMonth() + 1).padStart(2, '0');
    var d = String(value.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(value);
}

function getOrCreateOrderSheet(ss, date) {
  var sheet = ss.getSheetByName(date);
  if (!sheet) {
    sheet = ss.insertSheet(date);
    sheet.appendRow(['ID', '姓名', '餐點', '價格']);
  }
  return sheet;
}

function getOrCreateSettingsSheet(ss) {
  var sheet = ss.getSheetByName('設定');
  if (!sheet) {
    sheet = ss.insertSheet('設定');
    sheet.appendRow(['日期', '面', '網址']);
    sheet.getRange('A:A').setNumberFormat('@');
  }
  return sheet;
}

function getMenuUrl(settingsSheet, date, side) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow < 2) return null;

  const data = settingsSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < data.length; i++) {
    if (normalizeDate(data[i][0]) === date && data[i][1] === side) {
      return data[i][2] || null;
    }
  }
  return null;
}

function setMenuUrl(settingsSheet, date, side, url) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow >= 2) {
    const data = settingsSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (normalizeDate(data[i][0]) === date && data[i][1] === side) {
        settingsSheet.getRange(i + 2, 3).setValue(url);
        return;
      }
    }
  }
  settingsSheet.appendRow([date, side, url]);
}

function clearMenuUrls(settingsSheet, date) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow < 2) return;

  const data = settingsSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (normalizeDate(data[i][0]) === date) {
      settingsSheet.deleteRow(i + 2);
    }
  }
}

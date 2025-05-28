/**
 * 蝦皮組合配對上傳功能
 * 將 Google Sheets 中的組合資料上傳到資料庫
 */

/**
 * 主要上傳函數 - 讀取工作表資料並顯示確認對話框
 */
function uploadToDatabaseV4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("蝦皮組合配對1區");
  
  const rawDate = sheet.getRange("B1").getValue();
  const 日期 = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const 場次 = sheet.getRange("B2").getValue();
  const 購物車 = sheet.getRange("B3").getValue();
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 4) {
    SpreadsheetApp.getUi().alert("❌ 沒有找到任何組合資料（第5列以下）。請確認是否有填入資料！");
    return;
  }
  
  const header = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).getValues();
  
  const records = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const record = [日期, 場次, 購物車];
    let hasData = false;
    
    for (let j = 1; j < row.length; j += 2) {
      const 商品 = row[j];
      const 數量 = row[j + 1];
      if (商品 && 數量 && !isNaN(數量) && 數量 > 0) {
        hasData = true;
      }
    }
    
    if (hasData) {
      record.push(...row);
      records.push(record);
    }
  }
  
  if (records.length === 0) {
    SpreadsheetApp.getUi().alert("❌ 沒有有效的資料可上傳（所有組合皆為空）！");
    return;
  }
  
  const template = HtmlService.createTemplateFromFile("ConfirmUploadDialog");
  template.serializedRecords = JSON.stringify(records);
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(600).setHeight(600),
    "上傳前確認"
  );
}

/**
 * 從 JSON 字串確認上傳到資料庫
 * @param {string} rawJson - 序列化的記錄資料
 */
function confirmUploadToDatabaseFromString(rawJson) {
  try {
    const records = JSON.parse(rawJson);
    if (!Array.isArray(records)) throw new Error("資料格式錯誤");
    
    confirmUploadToDatabase(records);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ 上傳失敗：" + e.message);
    throw e;
  }
}

/**
 * 確認上傳資料到資料庫
 * @param {Array} records - 要上傳的記錄陣列
 */
function confirmUploadToDatabase(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = SpreadsheetApp.openById("1bpetUBRQ35ijRoFUKkiHU9PaHmIn2o-Cx93cnexUL7U").getSheetByName("資料庫");
  const logSheet = ss.getSheetByName("配對紀錄");
  const ui = SpreadsheetApp.getUi();
  const now = new Date();
  
  if (!records || !Array.isArray(records) || records.length === 0) {
    ui.alert("❌ 沒有有效資料可寫入！");
    return;
  }
  
  const rawDate = records[0][0];
  const 場次 = records[0][1];
  const 日期 = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  Logger.log(`✅ 嘗試上傳資料：日期=${日期}、場次=${場次}，共 ${records.length} 筆`);
  
  // 🧯 保險：取得資料庫中的場次清單，若沒資料則略過重複判斷
  let 是否重複 = false;
  const lastRow = dbSheet.getLastRow();
  if (lastRow > 1) {
    const dbRange = dbSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    是否重複 = dbRange.some(row => {
      const db日期 = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy/MM/dd");
      return db日期 === 日期 && row[1] === 場次;
    });
  }
  
  if (是否重複) {
    ui.alert(`❌ 日期「${日期}」場次「${場次}」已存在，請勿重複上傳！`);
    return;
  }
  
  // ✅ 寫入資料庫，只保留前 40 欄避免爆欄
  records.forEach(r => dbSheet.appendRow(r.slice(0, 40)));
  
  // ✅ 寫入記錄
  logSheet.appendRow([now, "✅ 成功", `上傳 ${records.length} 筆資料：${日期} / ${場次}`]);
  
  ui.alert(`✅ 成功寫入 ${records.length} 筆資料到資料庫！`);
}

# 蝦皮組合上傳工具 (Shopee Combo Uploader)

這是一個 Google Apps Script 工具，用於將 Google Sheets 中的蝦皮商品組合配對資料上傳到指定的資料庫工作表。

## 功能特色

- ✅ **資料驗證**：自動檢查資料完整性，確保只上傳有效的組合資料
- 🔍 **重複檢查**：防止相同日期和場次的資料重複上傳
- 📋 **預覽確認**：上傳前顯示詳細的資料預覽，讓使用者確認內容
- 📊 **自動記錄**：所有上傳操作都會記錄到配對紀錄工作表
- 🧹 **自動清除**：成功上傳後自動清除已上傳的資料，保留下拉選單設定
- 🎯 **錯誤處理**：完整的錯誤處理機制，提供清楚的錯誤訊息

## 檔案結構

```
shopee-combo-uploader/
├── upload.gs                    # 主要上傳功能程式碼
├── ConfirmUploadDialog.html      # 上傳確認對話框介面
└── README.md                    # 說明文件
```

## 使用方法

### 1. 設定工作表

確保你的 Google Sheets 包含以下工作表：

- **蝦皮組合配對1區**：主要資料來源工作表
  - B1：日期
  - B2：場次
  - B3：購物車
  - 第4列：標題列（商品名稱和數量）
  - 第5列以下：實際組合資料

- **配對紀錄**：用於記錄上傳操作的日誌

### 2. 執行上傳

1. 開啟 Google Sheets
2. 點選「擴充功能」→「Apps Script」
3. 執行 `uploadToDatabaseV4()` 函數
4. 在彈出的確認對話框中檢查資料
5. 點選「✅ 確認上傳」完成上傳
6. 系統會自動清除已上傳的資料，準備下一次使用

### 3. 自動清除機制

上傳成功後，系統會自動執行以下清除動作：
- **清除內容**：B1~B3 和第5列以下的所有資料
- **保留設定**：下拉選單、格式、資料驗證規則
- **即用性**：清除後可立即開始下一批資料輸入

### 4. 資料格式要求 資料格式要求

- 商品名稱和數量需要成對出現
- 數量必須是大於 0 的數字
- 至少要有一個有效的商品組合才能上傳

## 主要函數說明

### `uploadToDatabaseV4()`
主要入口函數，負責：
- 讀取工作表資料
- 驗證資料格式
- 顯示確認對話框

### `confirmUploadToDatabaseFromString(rawJson)`
處理 HTML 對話框傳來的 JSON 資料，進行格式驗證。

### `confirmUploadToDatabase(records)`
執行實際的資料庫寫入操作：
- 檢查重複資料
- 寫入到目標資料庫
- 記錄操作日誌
- 自動清除已上傳的資料

### `clearUploadedData()`
自動清除功能：
- 清除 B1~B3 欄位（日期、場次、購物車）
- 清除第5列以下的所有資料內容
- 保留下拉選單和格式設定

## 資料庫設定

目標資料庫 ID：`1bpetUBRQ35ijRoFUKkiHU9PaHmIn2o-Cx93cnexUL7U`

**注意**：請確保腳本有權限訪問目標資料庫工作表。

## 錯誤處理

系統會自動處理以下情況：
- 沒有找到有效資料
- 資料格式錯誤
- 重複上傳檢查
- 網路連線問題

所有錯誤都會顯示清楚的中文提示訊息。

## 安裝步驟

1. **建立新的 Google Apps Script 專案**
   ```
   前往 https://script.google.com
   點選「新增專案」
   ```

2. **複製程式碼**
   - 將 `upload.gs` 的內容貼到 `Code.gs` 中
   - 新增 HTML 檔案 `ConfirmUploadDialog.html`

3. **設定權限**
   - 第一次執行時會要求授權
   - 允許腳本存取 Google Sheets 和顯示對話框

4. **測試功能**
   - 準備測試資料
   - 執行 `uploadToDatabaseV4()` 函數

## 注意事項

⚠️ **重要提醒**：
- 使用前請備份重要資料
- 確認目標資料庫的存取權限
- 建議先用測試資料進行驗證

## 版本資訊

- **版本**：V4
- **最後更新**：2025年5月
- **相容性**：Google Apps Script

## 授權

此專案採用 MIT 授權條款。

## 聯絡方式

如有問題或建議，請透過 GitHub Issues 回報。

---

**🎯 讓蝦皮商品組合管理更輕鬆！**

/**
 * 公關活動報價 → Google Drive 歸檔
 * 部署設定：
 *   1. 開 https://script.google.com → New project
 *   2. 把整份檔案內容貼上、儲存（檔名隨意）
 *   3. 點右上角 Deploy → New deployment
 *      - Type: Web app
 *      - Description: 報價單歸檔
 *      - Execute as: Me (beyouwedding@gmail.com)
 *      - Who has access: Anyone
 *   4. 第一次部署會跳授權 → Allow（同意 Drive 與 Spreadsheet 權限）
 *   5. 複製拿到的 Web app URL，貼回報價工具的 Settings → Apps Script 網址
 *
 * 收到的 payload（form POST，欄位 name="payload"，值為 JSON 字串）：
 *   {
 *     quoteNumber, quoteDate, validUntil,
 *     clientCompany, clientContact, eventName, eventDate,
 *     subtotal, tax, total,
 *     vendorPdfBase64, clientPdfBase64,
 *     vendorHtml, clientHtml,
 *     stateJson, baseFileName
 *   }
 *
 * 行為：
 *   - 在 ROOT_FOLDER_ID 下建立/找到 客戶名 子資料夾
 *   - 存入 廠商版 PDF / 業主版 PDF / 廠商版 HTML / 業主版 HTML / 原始 JSON
 *   - 在根目錄維護一個「報價單總覽」Google Sheet，每次歸檔追加一列
 */

// ⚠️ 你的 Drive 根資料夾 ID（從分享網址末段擷取）
const ROOT_FOLDER_ID = '1VJngG53kL1mg79ISxPA8oVJVhYWdYNPA'
const SHEET_NAME = '報價單總覽'

function doPost(e) {
  try {
    const payload = JSON.parse(e.parameter.payload)
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID)

    // 1. 客戶資料夾（同名就重用，不存在就建立）
    const clientName = sanitize(payload.clientCompany || '未命名客戶')
    const clientFolder = getOrCreateFolder(root, clientName)

    const base = sanitize(payload.baseFileName || payload.quoteNumber)
    const out = {}

    // 2. 廠商版 PDF
    if (payload.vendorPdfBase64) {
      out.vendorPdfUrl = saveFile(
        clientFolder,
        `${base}_廠商版.pdf`,
        Utilities.base64Decode(payload.vendorPdfBase64),
        'application/pdf'
      )
    }

    // 3. 業主版 PDF
    if (payload.clientPdfBase64) {
      out.clientPdfUrl = saveFile(
        clientFolder,
        `${base}_業主版.pdf`,
        Utilities.base64Decode(payload.clientPdfBase64),
        'application/pdf'
      )
    }

    // 4. 廠商版 HTML
    if (payload.vendorHtml) {
      out.vendorHtmlUrl = saveFile(
        clientFolder,
        `${base}_廠商版.html`,
        payload.vendorHtml,
        'text/html'
      )
    }

    // 5. 業主版 HTML
    if (payload.clientHtml) {
      out.clientHtmlUrl = saveFile(
        clientFolder,
        `${base}_業主版.html`,
        payload.clientHtml,
        'text/html'
      )
    }

    // 6. 原始資料 JSON（可日後重新載入到報價工具編輯）
    if (payload.stateJson) {
      out.stateUrl = saveFile(
        clientFolder,
        `${base}_資料.json`,
        payload.stateJson,
        'application/json'
      )
    }

    // 7. 寫入「報價單總覽」Sheet
    appendToMasterSheet(root, payload, out, clientFolder)

    return jsonOut({ ok: true, folder: clientFolder.getUrl(), files: out })
  } catch (err) {
    return jsonOut({ ok: false, error: String(err), stack: err.stack || '' })
  }
}

function doGet() {
  return ContentService.createTextOutput(
    '✅ 報價單歸檔服務運作中\n根資料夾：' + ROOT_FOLDER_ID + '\n總覽 Sheet：' + SHEET_NAME
  )
}

// ─── helpers ─────────────────────────────────────────────────────

function saveFile(folder, name, content, mime) {
  // 同名檔案存在就先刪除（避免重複），再寫新版
  const it = folder.getFilesByName(name)
  while (it.hasNext()) it.next().setTrashed(true)
  const blob = Utilities.newBlob(content, mime, name)
  return folder.createFile(blob).getUrl()
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name)
  return it.hasNext() ? it.next() : parent.createFolder(name)
}

function getOrCreateMasterSheet(parent) {
  const it = parent.getFilesByName(SHEET_NAME)
  let ss
  if (it.hasNext()) {
    ss = SpreadsheetApp.open(it.next())
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME)
    DriveApp.getFileById(ss.getId()).moveTo(parent)
    initMasterSheet(ss)
  }
  return ss
}

function initMasterSheet(ss) {
  const sheet = ss.getActiveSheet()
  sheet.setName('歸檔記錄')
  const headers = [
    '歸檔時間', '報價單編號', '客戶名稱', '聯絡人',
    '活動名稱', '活動日期', '報價日期', '有效期限',
    '小計(NT$)', '稅(NT$)', '總計(NT$)',
    '廠商版 PDF', '業主版 PDF', '廠商版 HTML', '業主版 HTML',
    '原始資料 JSON', '客戶資料夾'
  ]
  sheet.appendRow(headers)
  sheet.setFrozenRows(1)
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a3560')
    .setFontColor('white')
  sheet.setColumnWidth(1, 150)
  sheet.setColumnWidth(2, 140)
  sheet.setColumnWidth(3, 160)
  for (let i = 12; i <= 17; i++) sheet.setColumnWidth(i, 110)
}

function appendToMasterSheet(root, payload, out, clientFolder) {
  const ss = getOrCreateMasterSheet(root)
  const sheet = ss.getActiveSheet()
  const linkOrEmpty = (url, label) => url ? `=HYPERLINK("${url}","${label}")` : ''
  sheet.appendRow([
    new Date(),
    payload.quoteNumber || '',
    payload.clientCompany || '',
    payload.clientContact || '',
    payload.eventName || '',
    payload.eventDate || '',
    payload.quoteDate || '',
    payload.validUntil || '',
    Number(payload.subtotal) || 0,
    Number(payload.tax) || 0,
    Number(payload.total) || 0,
    linkOrEmpty(out.vendorPdfUrl, '開啟'),
    linkOrEmpty(out.clientPdfUrl, '開啟'),
    linkOrEmpty(out.vendorHtmlUrl, '開啟'),
    linkOrEmpty(out.clientHtmlUrl, '開啟'),
    linkOrEmpty(out.stateUrl, '下載'),
    linkOrEmpty(clientFolder.getUrl(), '進入')
  ])
}

function sanitize(name) {
  return String(name).replace(/[\\/:*?"<>|]/g, '_').trim() || '未命名'
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}

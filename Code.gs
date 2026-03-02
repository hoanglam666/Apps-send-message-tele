const BOT_TOKEN = "YOUR_TOKEN";

const COL = {
  CHAT_ID: 1,
  MESSAGE: 2,
  STATUS: 3,
  CREATED_TIME: 4,
  DELAY_MINUTES: 5
};

const TELEGRAM = {
  LAST_UPDATE_KEY: "TELEGRAM_LAST_UPDATE_ID",
  DATA_SHEET_NAME: "data"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Telegram Alerts")
    .addItem("Gửi ngay dòng đang chọn", "sendSelectedPendingNow")
    .addItem("Gửi ngay tất cả PENDING", "sendAllPendingNow")
    .addSeparator()
    .addItem("Đọc dữ liệu Hàng dư từ group", "fetchGroupDataFromTelegram")
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range || e.range.getRow() === 1) {
    return;
  }

  const editedColumn = e.range.getColumn();
  const newValue = String(e.value || "").trim().toUpperCase();
  const oldValue = String(e.oldValue || "").trim().toUpperCase();

  if (editedColumn === COL.STATUS && newValue === "PENDING" && oldValue !== "PENDING") {
    e.range.getSheet().getRange(e.range.getRow(), COL.CREATED_TIME).setValue(new Date());
  }
}

function getBotToken_() {
  const scriptToken = PropertiesService.getScriptProperties().getProperty("BOT_TOKEN");
  const token = (scriptToken || BOT_TOKEN || "").trim();

  if (!token || token === "YOUR_TOKEN") {
    throw new Error("Thiếu BOT_TOKEN. Hãy cập nhật hằng BOT_TOKEN hoặc Script Properties BOT_TOKEN.");
  }

  return token;
}

function telegramApiGet_(method, queryParams) {
  const token = getBotToken_();
  const query = queryParams
    ? "?" + Object.keys(queryParams).map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(queryParams[key])}`).join("&")
    : "";
  const url = `https://api.telegram.org/bot${token}/${method}${query}`;

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true
  });
  const statusCode = response.getResponseCode();
  const content = response.getContentText();

  if (statusCode !== 200) {
    throw new Error(`Telegram API lỗi ${statusCode}: ${content}`);
  }

  return JSON.parse(content);
}

function sendTelegram(chatId, message) {
  const token = getBotToken_();
  const url = `https://api.telegram.org/bot${token}/sendMessage`;

  const payload = {
    chat_id: String(chatId).trim(),
    text: message,
    parse_mode: "HTML"
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();

  if (statusCode !== 200) {
    throw new Error(`Telegram API lỗi ${statusCode}: ${response.getContentText()}`);
  }
}

function processMessages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const rowIndex = i + 1;
    const status = String(data[i][COL.STATUS - 1] || "").trim().toUpperCase();
    const createdTime = data[i][COL.CREATED_TIME - 1];
    const delayMinutes = Number(data[i][COL.DELAY_MINUTES - 1]);

    if (!isValidPendingRow_(data[i], status)) {
      continue;
    }

    const effectiveCreatedTime = createdTime instanceof Date ? createdTime : new Date();

    if (!(createdTime instanceof Date)) {
      sheet.getRange(rowIndex, COL.CREATED_TIME).setValue(effectiveCreatedTime);
    }

    const diffMinutes = (now - effectiveCreatedTime) / 1000 / 60;

    if (!isNaN(delayMinutes) && diffMinutes >= delayMinutes) {
      sendRow_(sheet, rowIndex, data[i]);
    }
  }
}

function sendSelectedPendingNow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rowIndex = sheet.getActiveRange().getRow();

  if (rowIndex <= 1) {
    throw new Error("Vui lòng chọn một dòng dữ liệu (không phải header).");
  }

  const rowData = sheet.getRange(rowIndex, 1, 1, COL.DELAY_MINUTES).getValues()[0];
  const status = String(rowData[COL.STATUS - 1] || "").trim().toUpperCase();

  if (!isValidPendingRow_(rowData, status)) {
    throw new Error("Dòng đang chọn không hợp lệ hoặc không ở trạng thái PENDING.");
  }

  sendRow_(sheet, rowIndex, rowData);
}

function sendAllPendingNow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowIndex = i + 1;
    const status = String(data[i][COL.STATUS - 1] || "").trim().toUpperCase();

    if (isValidPendingRow_(data[i], status)) {
      sendRow_(sheet, rowIndex, data[i]);
    }
  }
}

function fetchGroupDataFromTelegram() {
  const scriptProps = PropertiesService.getScriptProperties();
  const lastUpdateId = Number(scriptProps.getProperty(TELEGRAM.LAST_UPDATE_KEY) || 0);
  const apiResult = telegramApiGet_("getUpdates", { offset: lastUpdateId + 1, timeout: 0 });

  if (!apiResult.ok || !Array.isArray(apiResult.result)) {
    throw new Error(`Không lấy được updates: ${JSON.stringify(apiResult)}`);
  }

  const rows = [];
  let maxUpdateId = lastUpdateId;

  apiResult.result.forEach((update) => {
    maxUpdateId = Math.max(maxUpdateId, Number(update.update_id || 0));
    const message = update.message;

    if (!message || !message.chat || !message.text) {
      return;
    }

    const chatType = String(message.chat.type || "").toLowerCase();
    if (chatType !== "group" && chatType !== "supergroup") {
      return;
    }

    const entries = parseHangDuEntries_(message.text);
    entries.forEach((entry) => {
      rows.push([entry.serial, "", entry.groupCode]);
    });
  });

  if (rows.length > 0) {
    const sheet = getOrCreateDataSheet_();
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);
  }

  scriptProps.setProperty(TELEGRAM.LAST_UPDATE_KEY, String(maxUpdateId));
}

function parseHangDuEntries_(text) {
  if (!text || text.toLowerCase().indexOf("hàng dư") === -1) {
    return [];
  }

  const regex = /(\d{6})\s*-\s*(\d{1,2})/g;
  const entries = [];
  let match;

  while ((match = regex.exec(text)) !== null) {
    entries.push({
      serial: match[1],
      groupCode: match[2].padStart(2, "0")
    });
  }

  return entries;
}

function getOrCreateDataSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(TELEGRAM.DATA_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(TELEGRAM.DATA_SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([["Data_6_So", "Message", "Ma_Group"]]);
  }

  return sheet;
}

function isValidPendingRow_(rowData, status) {
  const chatId = rowData[COL.CHAT_ID - 1];
  const message = rowData[COL.MESSAGE - 1];

  return (
    status === "PENDING" &&
    String(chatId).trim() !== "" &&
    String(message).trim() !== ""
  );
}

function sendRow_(sheet, rowIndex, rowData) {
  const chatId = rowData[COL.CHAT_ID - 1];
  const message = rowData[COL.MESSAGE - 1];

  try {
    sendTelegram(chatId, `📦 <b>CẢNH BÁO HỆ THỐNG</b>\n\n${message}`);
    sheet.getRange(rowIndex, COL.STATUS).setValue("SENT");
  } catch (error) {
    sheet.getRange(rowIndex, COL.STATUS).setValue("ERROR");
    sheet.getRange(rowIndex, COL.STATUS).setNote(String(error.message || error));
  }
}

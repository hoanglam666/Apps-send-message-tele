const BOT_TOKEN = "YOUR_TOKEN";

const COL = {
  CHAT_ID: 1,
  MESSAGE: 2,
  STATUS: 3,
  CREATED_TIME: 4,
  DELAY_MINUTES: 5
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Telegram Alerts")
    .addItem("Gửi ngay dòng đang chọn", "sendSelectedPendingNow")
    .addItem("Gửi ngay tất cả PENDING", "sendAllPendingNow")
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

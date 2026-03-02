const BOT_TOKEN = "YOUR_TOKEN";

function sendTelegram(chatId, message) {
  const url = `https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`;

  const payload = {
    chat_id: chatId,
    text: message,
    parse_mode: "HTML"
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(url, options);
}

function processMessages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const chatId = data[i][0];
    const message = data[i][1];
    const status = data[i][2];
    const createdTime = data[i][3];
    const delayMinutes = Number(data[i][4]);

    if (
      status === "PENDING" &&
      createdTime instanceof Date &&
      !isNaN(delayMinutes) &&
      String(chatId).trim() !== "" &&
      String(message).trim() !== ""
    ) {
      const diffMinutes = (now - createdTime) / 1000 / 60;

      if (diffMinutes >= delayMinutes) {
        try {
          sendTelegram(chatId, `📦 <b>CẢNH BÁO HỆ THỐNG</b>\n\n${message}`);
          sheet.getRange(i + 1, 3).setValue("SENT");
        } catch (e) {
          sheet.getRange(i + 1, 3).setValue("ERROR");
        }
      }
    }
  }
}

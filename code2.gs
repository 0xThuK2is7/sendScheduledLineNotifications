// แบบมี Sticker
// หา PackageId กับ StickerID ได้ที่ https://developers.line.biz/en/docs/messaging-api/sticker-list/#sticker-definitions
// ส่งแบบหลายกลุ่ม
function sendScheduledLineNotifications() {
  const tokens = ['YOUR_LINE_ACCESS_TOKEN_1', 'YOUR_LINE_ACCESS_TOKEN_2', 'YOUR_LINE_ACCESS_TOKEN_3'];
  const spreadsheetId = 'YOUR_SPREADSHEET_ID';
  const sheetName = 'YOUR_SHEET_NAME';
  const timeZone = 'GMT+7'; // Adjust to your desired time zone

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); // Start from row 2 and get columns A to H

  const now = new Date();
  const today = Utilities.formatDate(now, timeZone, 'dd/MM/yyyy');
  const currentTime = Utilities.formatDate(now, timeZone, 'HH:mm');

  for (const row of dataRange) {
    const date = Utilities.formatDate(row[1], timeZone, 'dd/MM/yyyy');
    const originalTime = row[2]; // Original time from the sheet in "HH:mm:ss" format

    // Adjust the timestamp date, month, and year while keeping the time unchanged
    const adjustedTimestamp = new Date(originalTime);
    adjustedTimestamp.setFullYear(now.getFullYear());
    adjustedTimestamp.setMonth(now.getMonth());
    adjustedTimestamp.setDate(now.getDate());

    const adjustedTime = Utilities.formatDate(adjustedTimestamp, timeZone, 'HH:mm');
    const stickerPackageId = row[5].toString();
    const stickerId = row[6].toString();
    const msg = row[3];

    const imgUrl = row[4];
    const imgID = extractImageID(imgUrl); // Extract the image ID from the URL

  for (const token of tokens) {
      if (today === date && currentTime === adjustedTime) {
        const message = `\n${msg}`;
        const image = null;

        if (imgID) {
          image = DriveApp.getFileById(imgID).getBlob();
        }
        sendLineNotify(message, image, token, stickerPackageId, stickerId);
      }
      console.log(today, date, currentTime, adjustedTime, stickerPackageId, stickerId, imgID);
    }
  }
}

function extractImageID(url) {
  const matches = url.match(/[-\w]{25,}/); // Extract the image ID from the URL
  return matches ? matches[0] : null;
}

function sendLineNotify(message, image, token, stickerPackageId, stickerId) {
  const options = {
    method: 'post',
    payload: {
      message: message,
      imageFile: image,
      stickerPackageId: stickerPackageId,
      stickerId: stickerId,
    },
    headers: { Authorization: `Bearer ${token}` },
  };
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', options);
}


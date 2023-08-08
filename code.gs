function sendScheduledLineNotifications() {
  const token = 'YOUR_LINE_ACCESS_TOKEN';
  const spreadsheetId = 'YOUR_SPREADSHEET_ID';
  const sheetName = 'YOUR_SHEET_NAME';
  const timeZone = 'GMT+7'; // Adjust to your desired time zone

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // Start from row 2 and get columns A to E

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

    const msg = row[3];
    const imgUrl = row[4];
    const imgID = imgUrl.split('https://drive.google.com/open?id=')[1];

    if (today === date && currentTime === adjustedTime) {
      const message = `\nแจ้งเตือน : ${msg}`;
      const image = DriveApp.getFileById(imgID).getBlob();
      sendLineNotify(message, image, token);
    }
    console.log(today,date,currentTime,adjustedTime);
  }
}

function sendLineNotify(message, image, token) {
  const options = {
    method: 'post',
    payload: {
      message: message,
      imageFile: image,
    },
    headers: { Authorization: `Bearer ${token}` },
  };
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', options);
}

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
    const originalTime = Utilities.formatDate(row[2], timeZone, 'HH:mm'); // Original time from the sheet
    const timeParts = originalTime.split(':');
    const hours = parseInt(timeParts[0]);
    const minutes = parseInt(timeParts[1]) - 17; // Subtract 17 minutes
    const adjustedTime = `${hours}:${minutes < 10 ? '0' : ''}${minutes}`; // Format adjusted time

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

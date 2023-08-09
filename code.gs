// แบบไม่มี Sticker
function sendScheduledLineNotifications() {
  const token = 'YOUR_LINE_ACCESS_TOKEN';
  const spreadsheetId = 'YOUR_SPREADSHEET_ID';
  const sheetName = 'YOUR_SHEET_NAME';
  const timeZone = 'GMT+7'; // Adjust to your desired time zone

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // 5 หมายถึง columns A ถึง E

  const now = new Date();
  const today = Utilities.formatDate(now, timeZone, 'dd/MM/yyyy');
  const currentTime = Utilities.formatDate(now, timeZone, 'HH:mm');

  for (const row of dataRange) {
    // column A = row[0], column B = row[1], column C = row[2], ...
    const date = Utilities.formatDate(row[1], timeZone, 'dd/MM/yyyy'); // column B
    const originalTime = row[2]; // colum C

    // Adjust the timestamp date, month, and year while keeping the time unchanged
    const adjustedTimestamp = new Date(originalTime);
    adjustedTimestamp.setFullYear(now.getFullYear());
    adjustedTimestamp.setMonth(now.getMonth());
    adjustedTimestamp.setDate(now.getDate());

    const adjustedTime = Utilities.formatDate(adjustedTimestamp, timeZone, 'HH:mm');

    const msg = row[3]; // column D
    const imgUrl = row[4]; // column E
    const imgID = extractImageID(imgUrl); // Extract the image ID from the URL

    if (today === date && currentTime === adjustedTime) {
        const message = `\n${msg}`;
        const image = null;

        if (imgID) {
          image = DriveApp.getFileById(imgID).getBlob();
        }
        sendLineNotify(message, image, token);
    }
    console.log(today,date,currentTime,adjustedTime);
  }
}

function extractImageID(url) {
  const matches = url.match(/[-\w]{25,}/); // Extract the image ID from the URL
  return matches ? matches[0] : null;
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

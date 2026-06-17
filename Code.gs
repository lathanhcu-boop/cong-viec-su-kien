// Thay đổi doGet để phục vụ như một API thay vì chỉ trả về file HTML
function doGet(e) {
  var action = e.parameter.action;

  if (action === "getMonthData") {
    var contractor = e.parameter.contractor;
    var type = e.parameter.tab;
    var year = e.parameter.year;
    var month = e.parameter.month;
    
    var data = getMonthData(contractor, type, year, month);
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Nếu không có action, trả về trang HTML (để phòng hờ khi mở trên trình duyệt)
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Công việc Sự kiện - The Dewey Schools THT')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Xử lý dữ liệu gửi lên từ fetch POST
function doPost(e) {
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  
  if (action === "saveData") {
    var result = saveData(params.contractor, params.tab, params.dateStr, params.content);
    return ContentService.createTextOutput(JSON.stringify({success: true, message: result}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  if (!sheet) {
    sheet = ss.insertSheet('Data');
    sheet.appendRow(['Nhà thầu', 'Loại', 'Ngày', 'Nội dung']);
  }
  return sheet;
}

function saveData(contractor, type, dateStr, content) {
  var sheet = initSheet();
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][2];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    if (data[i][0] === contractor && data[i][1] === type && rowDate === dateStr) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > -1) {
    sheet.getRange(foundRow, 4).setValue(content);
  } else {
    sheet.appendRow([contractor, type, dateStr, content]);
  }
  return "Đã lưu thông tin thành công!";
}

function getMonthData(contractor, type, year, month) {
  var sheet = initSheet();
  var data = sheet.getDataRange().getValues();
  var result = {};
  var prefix = year + "-" + String(month).padStart(2, '0');

  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][2];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      rowDate = String(rowDate);
    }
    
    if (data[i][0] === contractor && data[i][1] === type && rowDate.startsWith(prefix)) {
      result[rowDate] = data[i][3];
    }
  }
  return result;
}

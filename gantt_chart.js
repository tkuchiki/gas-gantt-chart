var colors ={
  "black": "#000000",
  "blue":  "#0000ff",
  "gray":  "#808080",
  "red":   "#ff0000"
};

var days      = ['日', '月', '火', '水', '木', '金', '土'];
var lastDates = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

var startDateColumn = 3; // 日の開始列
var dateRow         = 1; // 日の行
var dayRow          = 2; // 曜日の行
var startChartRow   = 3; // ガントチャートの開始列

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "日、曜日を設定",
    functionName : "setDate"
  },
  {
    name : "休日列を塗りつぶす",
    functionName : "setHolidayBGColor"
  }];
  sheet.addMenu("関数", entries);
};

function getActiveSheet() {
  return SpreadsheetApp.getActiveSheet(); 
}

function setHolidayBGColor() {
  var sheet      = getActiveSheet();
  var lastColumn = sheet.getLastColumn();
  
  for (i = startDateColumn; i <= lastColumn; i++) {
    if (sheet.getRange(dayRow, i).getValue().search(/(土|日|祝|休)/) != -1) {
      sheet.getRange(startChartRow, i, sheet.getLastRow() - dayRow).setBackgroundColor(getColorCode('gray'));
    }
  }
}

function setDate() {
  var yearMonth = Browser.inputBox('年月(YYYY/MM) を入力してください', Browser.Buttons.OK_CANCEL);
  var dates     = yearMonth.split('/');
  
  if (dates.length == 1) {
    Browser.msgBox(dates[0]  + " : データが不正な形式です。YYYY/MM の形式で入力してください。");
    return;
  }
  
  var year     = dates[0];
  var month    = parseInt(dates[1], 10);
  var lastDate =  0;
  
  if (month == 2 && checkLeapYear(year)) {
    lastDate = 29;
  }
  else {
    lastDate = lastDates[month];
  }
    
  var sheet = getActiveSheet();
  
  for (i = 1; i <= lastDate; i++) {
    sheet.getRange(dateRow, i + startDateColumn - 1).setValue(i);
    var date   = new Date(year, month - 1, i);
    var dayNum = date.getDay();

    var color = "black";
    if (dayNum == 6) {
      color = "blue";
    }
    else if (dayNum == 0) {
      color = "red";
    }
    
    sheet.getRange(dayRow, i + startDateColumn - 1).setValue(days[dayNum]).setFontColor(getColorCode(color));
  }
  
  // 1日だけ M/D 表記にする
  sheet.getRange(dateRow, startDateColumn).setValue(month + "/" + 1);
}

function getColorCode(name) {
  return colors[name];
}

function checkLeapYear(year) {
  return (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
}

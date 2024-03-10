// SPREADSHEETのIDを記入
const TARGET_SPREADSHEET_ID = 'xxx';

// 学期の全Lesson数（予定）を記入
const all = "３";

const SHEET_NAMES = ['途中経過', '学年の成績予想'];
const TARGET_SPREADSHEET = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
const TARGET_SHEETS = SHEET_NAMES.map(sheetName => TARGET_SPREADSHEET.getSheetByName(sheetName));

let user = Session.getActiveUser();
let email = user.getEmail();
let row = findRow(TARGET_SHEETS[0], email, 1);
function findRow(sheet, val, col) {
  let dat = sheet.getDataRange().getValues();
  for (let i = 1; i < dat.length; i++) {
    if (dat[i][col - 1] === val) {
      return i + 1;
    }
  }
  return 0;
}
// -----------------------------------------------------
// アプリを開いた時に実行する関数（index.htmlを表示する） *gs関数1
// -----------------------------------------------------
function doGet(e) {
  let page = e.parameter.page || 'index';
  let html = HtmlService.createTemplateFromFile(page);

  // 学期の全Lesson数（予定）を記入
  html.all = all;

  html.title = TARGET_SHEETS[0].getRange(1, 1).getValue();
  html.semester = TARGET_SHEETS[0].getRange(1, 2).getValue();
  html.term = TARGET_SHEETS[0].getRange(1, 6).getValue();
  html.user_number = TARGET_SHEETS[0].getRange(row, 2).getValue();
  html.user_name = TARGET_SHEETS[0].getRange(row, 3).getValue();

  for (let i = 0; i < SHEET_NAMES.length - 1; i++) {
    html['gt_' + (i)] = TARGET_SHEETS[i].getRange(1, 4).getValue();
    html['kMaxScore_' + (i)] = TARGET_SHEETS[i].getRange(1, 7).getValue();
    html['tMaxScore_' + (i)] = TARGET_SHEETS[i].getRange(1, 8).getValue();
    html['aMaxScore_' + (i)] = TARGET_SHEETS[i].getRange(1, 9).getValue();
    html['lesson_' + (i)] = TARGET_SHEETS[i].getRange(1, 15).getValue();
    html['kRawScore_' + (i)] = TARGET_SHEETS[i].getRange(row, 4).getValue();
    html['tRawScore_' + (i)] = TARGET_SHEETS[i].getRange(row, 5).getValue();
    html['aRawScore_' + (i)] = TARGET_SHEETS[i].getRange(row, 6).getValue();
    html['kRawPercentage_' + (i)] = TARGET_SHEETS[i].getRange(row, 7).getValue();
    html['tRawPercentage_' + (i)] = TARGET_SHEETS[i].getRange(row, 8).getValue();
    html['aRawPercentage_' + (i)] = TARGET_SHEETS[i].getRange(row, 9).getValue();
    html['kRawPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 10).getValue();
    html['tRawPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 11).getValue();
    html['aRawPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 12).getValue();
    html['rawRating_' + (i)] = TARGET_SHEETS[i].getRange(row, 13).getValue();
    html['absence_' + (i)] = TARGET_SHEETS[i].getRange(row, 14).getValue();
    html['kExpectedPercentage_' + (i)] = TARGET_SHEETS[i].getRange(row, 15).getValue();
    html['tExpectedPercentage_' + (i)] = TARGET_SHEETS[i].getRange(row, 16).getValue();
    html['aExpectedPercentage_' + (i)] = TARGET_SHEETS[i].getRange(row, 17).getValue();
    html['kExpectedPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 18).getValue();
    html['tExpectedPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 19).getValue();
    html['aExpectedPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 20).getValue();
    html['expectedRating_' + (i)] = TARGET_SHEETS[i].getRange(row, 21).getValue();
    html['kNextPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 22).getValue();
    html['tNextPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 23).getValue();
    html['aNextPerspective_' + (i)] = TARGET_SHEETS[i].getRange(row, 24).getValue();
    html['kNextScore_' + (i)] = TARGET_SHEETS[i].getRange(row, 25).getValue();
    html['tNextScore_' + (i)] = TARGET_SHEETS[i].getRange(row, 26).getValue();
    html['aNextScore_' + (i)] = TARGET_SHEETS[i].getRange(row, 27).getValue();
  }

  html.kPerspective_final = TARGET_SHEETS[SHEET_NAMES.length - 1].getRange(row, 4).getValue();
  html.tPerspective_final = TARGET_SHEETS[SHEET_NAMES.length - 1].getRange(row, 5).getValue();
  html.aPerspective_final = TARGET_SHEETS[SHEET_NAMES.length - 1].getRange(row, 6).getValue();
  html.rating_final = TARGET_SHEETS[SHEET_NAMES.length - 1].getRange(row, 7).getValue();
  html.absence_final = TARGET_SHEETS[SHEET_NAMES.length - 1].getRange(row, 8).getValue();

  const htmlEvl = html.evaluate();
  htmlEvl.setTitle("途中経過");
  return htmlEvl; 
}

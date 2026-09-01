const S_LOGIN = 'تسجيل الدخول';
const S_RESP = 'الردود';
const S_DONE = 'تم التعديل';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('منصة الإدارة المدرسية')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function sanitizeForJSON(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  }
  return value === null || value === undefined ? "" : String(value);
}

function getQuestions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_RESP);
  if (!sheet) return[];
  var vals = sheet.getRange(1, 3, 1, 15).getValues()[0]; 
  return vals.map(sanitizeForJSON);
}

function checkSchoolLogin(schoolId) {
  if (!schoolId) return false;
  var searchId = String(schoolId).trim().toLowerCase(); 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LOGIN);
  var data = sheet.getRange("A:A").getValues();
  
  for (var i = 1; i < data.length; i++) {
    var dbId = String(data[i][0]).trim().toLowerCase();
    if (dbId === searchId && searchId !== "") {
      return true;
    }
  }
  return false;
}

function checkAdminLogin(adminCode) {
  if (!adminCode) return false;
  var searchCode = String(adminCode).trim().toLowerCase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LOGIN);
  var data = sheet.getRange("F:F").getValues();
  
  for (var i = 1; i < data.length; i++) {
    var dbCode = String(data[i][0]).trim().toLowerCase();
    if (dbCode === searchCode && searchCode !== "") {
      return true;
    }
  }
  return false;
}

function submitResponse(schoolId, answers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(S_RESP);
  
  var timestamp = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var uuid = Utilities.getUuid(); 
  
  var rowData = [timestamp, schoolId].concat(answers);
  rowData.push(uuid); 
  
  sheet.appendRow(rowData);
  return true;
}

function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var respSheet = ss.getSheetByName(S_RESP);
  var doneSheet = ss.getSheetByName(S_DONE);
  
  if (!doneSheet) doneSheet = ss.insertSheet(S_DONE);

  var lastRowResp = respSheet.getLastRow();
  if (lastRowResp < 2) return { headers: [], pending:[] };
  
  var respData = respSheet.getRange(1, 1, lastRowResp, 18).getValues();
  var displayDates = respSheet.getRange(1, 1, lastRowResp, 1).getDisplayValues();
  var headers = respData[0].map(sanitizeForJSON);
  
  var lastRowDone = doneSheet.getLastRow();
  var doneIds =[];
  if (lastRowDone >= 1) {
    var doneData = doneSheet.getRange(1, 1, lastRowDone, 18).getValues();
    for (var i = 0; i < doneData.length; i++) {
      if (doneData[i][17]) doneIds.push(String(doneData[i][17]));
    }
  }

  var pending =[];
  for (var i = 1; i < respData.length; i++) {
    var uuid = respData[i][17] ? String(respData[i][17]) : "";
    var schoolId = String(respData[i][1]).trim();
    
    if (uuid && doneIds.indexOf(uuid) === -1 && schoolId !== "") {
      pending.push({
        timestamp: displayDates[i][0],
        schoolId: schoolId,
        answers: respData[i].slice(2, 17).map(sanitizeForJSON),
        uuid: uuid
      });
    }
  }
  
  return { headers: headers.slice(2, 17), pending: pending };
}

function markAsCompleted(uuid, empName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var respSheet = ss.getSheetByName(S_RESP);
  var doneSheet = ss.getSheetByName(S_DONE);
  
  var lastRow = respSheet.getLastRow();
  if(lastRow < 2) return false;
  
  var respData = respSheet.getRange(1, 1, lastRow, 18).getValues();
  var targetRow = null;
  
  for (var i = 1; i < respData.length; i++) {
    if (String(respData[i][17]) === String(uuid)) {
      targetRow = respData[i].slice(); 
      break;
    }
  }
  
  if (targetRow) {
    while (targetRow.length < 18) targetRow.push(""); 
    targetRow.push(empName);
    targetRow.push(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));
    doneSheet.appendRow(targetRow);
    return true;
  }
  return false;
}
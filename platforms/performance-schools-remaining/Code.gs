function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('منصة دورة الأداء')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getMetadata() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('المراحل الدراسية');
  var data = sheet.getDataRange().getValues();

  var supervisors = []; 
  var schoolsData = []; 

  for(var i=1; i<data.length; i++) {
    if(data[i][1]) {
      supervisors.push({name: data[i][1].toString().trim(), phase: data[i][2].toString().trim()});
    }
    if(data[i][13]) {
      schoolsData.push({
        school: data[i][13].toString().trim(),
        sector: data[i][14].toString().trim(),
        gender: data[i][15].toString().trim()
      });
    }
  }
  return { supervisors: supervisors, schools: schoolsData };
}

function verifyLeadershipLogin(code) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('المراحل الدراسية');
  var lastRow = sheet.getLastRow();
  
  var unitHeads = sheet.getRange("F2:H" + lastRow).getValues();
  var deptHeads = sheet.getRange("J2:K" + lastRow).getValues();
  var cleanCode = code.toString().trim();

  for (var j = 0; j < deptHeads.length; j++) {
    if (deptHeads[j][0] && deptHeads[j][1].toString().trim() === cleanCode) {
      return { success: true, role: "dept_head", phase: "الكل", name: deptHeads[j][0] };
    }
  }

  for (var i = 0; i < unitHeads.length; i++) {
    if (unitHeads[i][0] && unitHeads[i][2].toString().trim() === cleanCode) {
      return { success: true, role: "unit_head", phase: unitHeads[i][1].toString().trim(), name: unitHeads[i][0] };
    }
  }
  return { success: false };
}

function getDashboardData(filterPhase, role) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var basicData = ss.getSheetByName('البيانات الاساسية').getDataRange().getValues();
  var remData = ss.getSheetByName('المدارس المتبقية').getDataRange().getValues();

  var basic = [];
  var rem = [];

  for(var i=1; i<basicData.length; i++) {
    var phaseFull = basicData[i][3] ? basicData[i][3].toString().trim() : "";
    var isMatch = (role === "dept_head") || phaseFull.indexOf(filterPhase) !== -1;
    if(isMatch) {
      basic.push({
         sup: basicData[i][2] ? basicData[i][2].toString().trim() : "",
         phase: phaseFull,
         assigned: Number(basicData[i][6]) || 0,
         visited: Number(basicData[i][7]) || 0,
         remaining: Number(basicData[i][8]) || 0
      });
    }
  }

  for(var j=1; j<remData.length; j++) {
    var phaseFull = remData[j][6] ? remData[j][6].toString().trim() : ""; 
    var isMatch = (role === "dept_head") || phaseFull.indexOf(filterPhase) !== -1;
    if(isMatch) {
       rem.push({
          sup: remData[j][2] ? remData[j][2].toString().trim() : "",
          school: remData[j][3] ? remData[j][3].toString().trim() : "",
          sector: remData[j][4] ? remData[j][4].toString().trim() : "غير محدد", // الاستدعاء حصرياً من العمود E
          gender: remData[j][5] ? remData[j][5].toString().trim() : "",
          method: remData[j][8] ? remData[j][8].toString().trim() : "" 
       });
    }
  }

  return { basic: basic, rem: rem };
}

function processForm(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var basicSheet = ss.getSheetByName('البيانات الاساسية');
  var remainingSheet = ss.getSheetByName('المدارس المتبقية');

  var timestamp = new Date();
  var cleanName = payload.supervisorName.trim().replace(/\s+/g, '_');
  var shortId = Math.random().toString(36).substring(2, 7); 
  var uniqueId = cleanName + "_" + shortId;

  basicSheet.appendRow([
    uniqueId, timestamp, payload.supervisorName, payload.studyPhase, 
    payload.perfCycle, payload.gender, payload.assignedCount, 
    payload.visitedCount, payload.remainingCount
  ]);

  if (payload.remainingSchools && payload.remainingSchools.length > 0) {
    payload.remainingSchools.forEach(function(school) {
      remainingSheet.appendRow([
        uniqueId, timestamp, payload.supervisorName, school.schoolName, 
        school.sector, school.gender, payload.studyPhase, 
        payload.perfCycle, school.methodUsed, school.visitDate, school.day
      ]);
    });
  }
  return "تم حفظ البيانات بنجاح!";
}
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('منصة رفع التقييم للمدارس (المرحلة الثانية)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function extractValidCode(val) {
  if (val === undefined || val === null || val === "") return "";
  var str = String(val).trim();
  if (/^\d+$/.test(str) && str.length < 4) str = str.padStart(4, '0');
  return str;
}

function getSheetSafe(ss, targetName) {
  var sheets = ss.getSheets();
  var cleanTarget = targetName.replace(/\s/g, '').replace(/ة/g, 'ه').replace(/[أإآ]/g, 'ا');
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().replace(/\s/g, '').replace(/ة/g, 'ه').replace(/[أإآ]/g, 'ا');
    if (sName === cleanTarget) return sheets[i];
  }
  return null;
}

function verifyLogin(passcode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetSafe(ss, "تسجيل الدخول");
    if (!sheet) return { error: 'ورقة "تسجيل الدخول" غير موجودة.' };
    
    var data = sheet.getDataRange().getValues();
    passcode = String(passcode).trim();
    
    for (var i = 1; i < data.length; i++) {
      var dashCode = extractValidCode(data[i][10]);
      if (dashCode !== "" && dashCode.slice(-4) === passcode) return { role: 'dashboard' };
      
      var code1 = extractValidCode(data[i][3]);
      if (code1 !== "" && code1.slice(-4) === passcode) return { role: 'supervisor', name: String(data[i][1]).trim(), stage: String(data[i][2]).trim() };
      
      var code2 = extractValidCode(data[i][7]);
      if (code2 !== "" && code2.slice(-4) === passcode) return { role: 'supervisor', name: String(data[i][5]).trim(), stage: String(data[i][6]).trim() };
    }
    return { error: 'الكود غير صحيح أو غير مسجل بالنظام' };
  } catch (e) { return { error: 'حدث خطأ داخلي: ' + e.message }; }
}

function getFormData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var respSheet = getSheetSafe(ss, "الردود");
    var dataSheet = getSheetSafe(ss, "البيانات الأساسية");
    var loginSheet = getSheetSafe(ss, "تسجيل الدخول");
    
    if (!respSheet || !dataSheet || !loginSheet) return { error: 'يرجى التأكد من وجود أوراق العمل.' };
    
    var lastCol = respSheet.getLastColumn();
    if (lastCol === 0) return { error: 'ورقة "الردود" فارغة! يرجى كتابة الأسئلة في الصف الأول.' };
    
    var headers = respSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var schoolData = dataSheet.getDataRange().getValues();
    var schools =[];
    for(var i = 1; i < schoolData.length; i++) {
      if(schoolData[i][2]) { 
        schools.push({ name: String(schoolData[i][2]).trim(), gender: schoolData[i][3] ? String(schoolData[i][3]).trim() : "" });
      }
    }

    var loginData = loginSheet.getDataRange().getValues();
    var supervisors =[];
    for(var i = 1; i < loginData.length; i++) {
      var name1 = loginData[i][1] ? String(loginData[i][1]).trim() : "";
      var name2 = loginData[i][5] ? String(loginData[i][5]).trim() : "";
      if (name1 && supervisors.indexOf(name1) === -1) supervisors.push(name1);
      if (name2 && supervisors.indexOf(name2) === -1) supervisors.push(name2);
    }
    
    return { headers: headers, schools: schools, supervisors: supervisors };
  } catch (e) { return { error: 'تعذر جلب البيانات: ' + e.message }; }
}

function submitResponse(formData, fileData, loggedInSupervisor, stage) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetSafe(ss, "الردود");
    if (!sheet) return "خطأ: لم يتم العثور على ورقة الردود.";

    var driveLink = formData['رابط_درايف'] ? String(formData['رابط_درايف']).trim() : ""; 
    var finalFileUrl = driveLink; // تعيين الرابط كقيمة افتراضية للطوارئ
    var schoolName = formData['اسم المدرسة'] || 'مدرسة غير محددة';
    var timestamp = new Date(); 
    var finalSupervisorFolder = formData['اسم المشرف المقيم'] || loggedInSupervisor;

    // --- معالجة الرفع والتخزين السحابي بطريقة مضادة للأعطال (Fail-safe) ---
    if ((fileData && fileData.bytes) || driveLink) {
      try {
        var ssId = ss.getId();
        var parentFolder;
        try {
          var parents = DriveApp.getFileById(ssId).getParents();
          parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
        } catch(e) { parentFolder = DriveApp.getRootFolder(); }

        var mainFolder = getOrCreateFolder("منصة رفع التقييم للمدارس (المرحلة الثانية)", parentFolder);
        var stageFolder = getOrCreateFolder(stage, mainFolder);
        var supFolder = getOrCreateFolder(finalSupervisorFolder, stageFolder);
        var schoolFolder = getOrCreateFolder(schoolName, supFolder);
        var finalFolder = schoolFolder.getFiles().hasNext() 
          ? getOrCreateFolder(Utilities.formatDate(timestamp, "GMT+3", "yyyy-MM-dd_HH-mm"), schoolFolder) 
          : schoolFolder;

        // رفع الملف من الجهاز
        if (fileData && fileData.bytes) {
          var mime = fileData.mimeType || "application/octet-stream";
          var blob = Utilities.newBlob(Utilities.base64Decode(fileData.bytes), mime, fileData.filename);
          var createdFile = finalFolder.createFile(blob);
          finalFileUrl = createdFile.getUrl();
        } 
        // أو نسخ رابط درايف
        else if (driveLink) {
          var idMatch = driveLink.match(/[-\w]{25,}/);
          if (idMatch) {
            try {
              var fileToCopy = DriveApp.getFileById(idMatch[0]);
              finalFileUrl = fileToCopy.makeCopy(fileToCopy.getName(), finalFolder).getUrl();
            } catch(copyErr) {
              // إذا كان الملف محمي ضد النسخ أو لا يوجد صلاحية، احتفظ بالرابط الأصلي ولا توقف العملية
              finalFileUrl = driveLink;
            }
          }
        }
      } catch(driveErr) {
        // إذا فشل نظام الدرايف تماماً، نرفع كحل بديل وننقذ الموقف
        if (fileData && fileData.bytes) {
          try {
            var blobFallback = Utilities.newBlob(Utilities.base64Decode(fileData.bytes), fileData.mimeType || "application/octet-stream", fileData.filename);
            finalFileUrl = DriveApp.createFile(blobFallback).getUrl();
          } catch(fatalErr) {
            throw new Error("تأكد من توفر مساحة كافية في حساب جوجل درايف الخاص بك.");
          }
        }
      }
    } else {
      throw new Error("يرجى إرفاق ملف أو وضع رابط صالح.");
    }

    // --- تخزين البيانات مع ضمان العمودين F و G (الهوية والمرحلة) ---
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowData = new Array(Math.max(headers.length, 7)).fill(""); // ضمان وجود 7 أعمدة على الأقل لتخزين F و G

    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim();
      
      // التخزين الإجباري والدقيق
      if (h === 'تاريخ ووقت الرد' || h === 'التاريخ') rowData[i] = timestamp;
      else if (i === 5) rowData[i] = loggedInSupervisor; // العمود F = اسم المشرف حصراً
      else if (i === 6) rowData[i] = stage;              // العمود G = المرحلة حصراً
      else if (h === 'ملف الاستمارة') rowData[i] = finalFileUrl;
      else if (h === 'الجنس' || h === 'جنس المدرسة') rowData[i] = formData['الجنس'] || "";
      else if (h === 'اسم المدرسة' || h === 'المدرسة') rowData[i] = formData['اسم المدرسة'] || "";
      else rowData[i] = formData[h] || ""; 
    }

    sheet.appendRow(rowData); 
    return "تم الإرسال والاعتماد بنجاح ✔️";
    
  } catch (err) {
    return "❌ حدث خطأ: " + err.message;
  }
}

function getOrCreateFolder(name, parent) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. حساب المدارس المستهدفة
  var dataSheet = getSheetSafe(ss, "البيانات الأساسية");
  var totalSchoolsCount = 0;
  if (dataSheet) {
    var sData = dataSheet.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) { if (sData[i][2]) totalSchoolsCount++; }
  }

  // 2. تحليل الإنجاز بناءً على الهوية (العمود F والمرحلة G)
  var respSheet = getSheetSafe(ss, "الردود");
  var respData = respSheet.getDataRange().getValues();
  var respHeaders = respData[0] ||[];
  
  var schoolIndex = respHeaders.indexOf('اسم المدرسة');
  if (schoolIndex === -1) schoolIndex = 2; // افتراضي
  
  var uniqueSchools = {}; 
  var stageStats = {}; 
  var activeSupervisors = {}; // أسماء المشرفين المنجزين
  
  if (respData.length > 1) {
    for (var i = 1; i < respData.length; i++) {
      var school = respData[i][schoolIndex];
      var supName = String(respData[i][5]).trim(); // العمود F (اسم المشرف الذي سجل الدخول)
      var stage = String(respData[i][6]).trim();   // العمود G (المرحلة)
      
      if (supName) activeSupervisors[supName] = true;
      if (school && !uniqueSchools[school]) { 
        uniqueSchools[school] = true; 
        if (stage) stageStats[stage] = (stageStats[stage] || 0) + 1; 
      }
    }
  }
  
  // 3. مقارنة ذكية (Cross-Reference) لتحديد من لم يبدأ
  var loginSheet = getSheetSafe(ss, "تسجيل الدخول");
  var loginData = loginSheet.getDataRange().getValues();
  var completedList =[]; 
  var pendingList =[];
  
  for (var i = 1; i < loginData.length; i++) {
    var name1 = loginData[i][1] ? String(loginData[i][1]).trim() : "";
    if (name1) { activeSupervisors[name1] ? completedList.push(name1) : pendingList.push(name1); }
    
    var name2 = loginData[i][5] ? String(loginData[i][5]).trim() : "";
    if (name2) { activeSupervisors[name2] ? completedList.push(name2) : pendingList.push(name2); }
  }
  
  return { 
    stageStats: stageStats, 
    completedCount: [...new Set(completedList)].length, 
    pendingCount: [...new Set(pendingList)].length, 
    evaluatedSchoolsCount: Object.keys(uniqueSchools).length, 
    totalSchoolsCount: totalSchoolsCount,
    completedList: [...new Set(completedList)], 
    pendingList: [...new Set(pendingList)] 
  };
}
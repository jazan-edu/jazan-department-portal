/**
 * منصة إدارة المرفقات والتقارير الإشرافية - النسخة النهائية المحدثة
 */

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('منصة رفع شواهد الأداء')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. نظام التحقق
function authenticateUser(code, portalType) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('تسجيل الدخول');
  const data = sheet.getDataRange().getValues();
  const targetCode = String(code).trim();
  
  for (let i = 1; i < data.length; i++) {
    if (portalType === 'supervisor') {
      if (String(data[i][8]).slice(-4) === targetCode && data[i][7]) {
        return { name: data[i][7], type: 'supervisor', isSpecial: true };
      }
      if (String(data[i][5]).slice(-4) === targetCode && data[i][4]) {
        return { name: data[i][4], type: 'supervisor', isSpecial: false };
      }
    } else if (portalType === 'unit_head') {
      if (String(data[i][2]).slice(-4) === targetCode && data[i][1]) {
        return { name: data[i][1], type: 'unit_head' };
      }
    } else if (portalType === 'leader') {
      if (String(data[i][11]).slice(-4) === targetCode && data[i][10]) {
        return { name: data[i][10], type: 'leader' };
      }
    }
  }
  return null;
}

// 2. جلب الأسئلة ديناميكياً
function getQuestions(user) {
  const ss = SpreadsheetApp.openById(SS_ID);
  
  if (user.type === 'supervisor') {
    const sheet = ss.getSheetByName('ردود المشرفين');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    let targetIndices = user.isSpecial ? [2, 3, 4, 5, 6, 8, 9] : [2, 3, 4, 5, 6, 7, 9];
    return targetIndices.map(idx => ({ 
      index: idx, 
      title: headers[idx] || ("سؤال " + (idx+1)),
      expectedType: 'number',
      isRequired: idx === 9 ? false : true 
    }));
  } else {
    const sheet = ss.getSheetByName('ردود رؤساء الوحدات');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let targetIndices =[2, 3, 4, 5, 6, 7, 8];
    return targetIndices.map(idx => ({ 
      index: idx, 
      title: headers[idx] || ("سؤال " + (idx+1)),
      expectedType: 'number',
      isRequired: idx === 8 ? false : true 
    }));
  }
}

// 3. معالجة الإرسال وتخزين البيانات (معدلة جذرياً حسب الهيكل الهرمي المطلوب)
function submitFormData(formData, filesData) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const user = formData.user;
  const timestamp = new Date();
  
  const responseSheetName = user.type === 'supervisor' ? 'ردود المشرفين' : 'ردود رؤساء الوحدات';
  const fileSheetName = user.type === 'supervisor' ? 'ملفات المشرفين' : 'ملفات رؤساء الوحدات';
  
  const responseSheet = ss.getSheetByName(responseSheetName);
  let fileSheet = ss.getSheetByName(fileSheetName);
  
  if (!fileSheet) {
    fileSheet = ss.insertSheet(fileSheetName);
    fileSheet.appendRow(['الطابع الزمني', 'الاسم', 'السؤال', 'اسم الملف / المصدر', 'رابط الملف', 'نوع المشاركة']);
    fileSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#BDB76B").setFontColor("#000000");
  }
  
  const lastCol = responseSheet.getLastColumn();
  const newRow = new Array(lastCol).fill("");
  newRow[0] = timestamp; 
  newRow[1] = user.name; 
  
  // ----------------------------------------------------------------------
  // بداية المنطق الجديد للهيكل التنظيمي للمجلدات
  // ----------------------------------------------------------------------
  
  // دالة مساعدة للتحقق من وجود المجلد لعدم التكرار
  function getOrCreateFolder(parentFolder, folderName) {
    const folders = parentFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
  }

  // 1. تحديد مكان ملف الشيت الحالي ليكون هو الأساس
  const ssFile = DriveApp.getFileById(SS_ID);
  const parentFolders = ssFile.getParents();
  const mainParent = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  
  // 2. المجلد الرئيسي للمشروع (بجوار ملف الشيت)
  const rootFolder = getOrCreateFolder(mainParent, "مرفقات المنصة المحدثة");

  // 3. المجلدات الثابتة للتصنيف
  const unitHeadsFolder = getOrCreateFolder(rootFolder, "رؤساء الوحدات");
  const supervisorsFolder = getOrCreateFolder(rootFolder, "المشرفين والمشرفات");

  // 4. تحديد المجلد المستهدف بناءً على نوع المستخدم
  const categoryFolder = (user.type === 'supervisor') ? supervisorsFolder : unitHeadsFolder;

  // 5. مجلد باسم المستخدم (التحقق من وجوده)
  const userFolder = getOrCreateFolder(categoryFolder, user.name);

  // 6. مجلد تاريخ ووقت الرد (يتم إنشاؤه جديد دائماً لتمييز الردود المتعددة)
  const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH-mm-ss");
  const submissionFolder = userFolder.createFolder("رد بتاريخ " + dateStr);

  // تخزين مؤقت لمجلدات الأسئلة داخل هذا الرد (لزيادة السرعة ودعم متصفح Safari)
  let questionFolders = {};

  // 7. معالجة وتخزين الملفات
  filesData.forEach(fileObj => {
    let fileUrl = "";
    let savedFileName = "";
    
    // تنظيف اسم السؤال ليكون صالحاً كاسم مجلد
    let cleanTitle = String(fileObj.qTitle).replace(/[\\/:*?"<>|]/g, "_");
    let qFolderName = fileObj.qIndex + " - " + cleanTitle;
    
    // إنشاء مجلد خاص بالسؤال الداخلي إذا لم يتم إنشاؤه في هذه الدفعة
    if (!questionFolders[qFolderName]) {
      questionFolders[qFolderName] = submissionFolder.createFolder(qFolderName);
    }
    let currentQFolder = questionFolders[qFolderName];
    
    if (fileObj.source === 'device') {
      let blob = Utilities.newBlob(Utilities.base64Decode(fileObj.base64), fileObj.mimeType, user.name + "_" + fileObj.fileName);
      // رفع الملف فعلياً داخل مجلد السؤال المخصص
      let file = currentQFolder.createFile(blob); 
      fileUrl = file.getUrl(); // أخذ الرابط بعد التأكد من نجاح الرفع
      savedFileName = fileObj.fileName;
    } else {
      fileUrl = fileObj.driveUrl;
      savedFileName = "مرفوع عبر رابط درايف";
    }
    
    // تسجيل بيانات الملف في الشيت
    fileSheet.appendRow([timestamp, user.name, fileObj.qTitle, savedFileName, fileUrl, fileObj.partType || ""]);
  });
  
  // ----------------------------------------------------------------------
  // نهاية المنطق الجديد
  // ----------------------------------------------------------------------

  formData.answers.forEach(ans => {
    newRow[ans.qIndex] = ans.value;
  });

  responseSheet.appendRow(newRow);
  return true;
}

// 4. دالة جلب بيانات لوحة القيادات
function getLeaderDashboardData() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let dashboardData = { supervisors: [], unitHeads:[] };

  function processData(responseSheetName, fileSheetName, userType) {
    const rSheet = ss.getSheetByName(responseSheetName);
    const fSheet = ss.getSheetByName(fileSheetName);
    if (!rSheet) return;

    const rData = rSheet.getDataRange().getValues();
    const headers = rData[0];
    
    let latestResponses = {};
    for (let i = 1; i < rData.length; i++) {
      let row = rData[i];
      let timestamp = new Date(row[0]).getTime();
      let name = String(row[1]).trim();
      if (!name) continue;
      
      if (!latestResponses[name] || timestamp > latestResponses[name].timestamp) {
        latestResponses[name] = { 
          timestamp: timestamp, 
          dateStr: row[0], 
          answers: row.slice(2).map((val, idx) => ({ question: headers[idx+2], answer: val }))
        };
      }
    }

    let filesData = {};
    if (fSheet) {
      const fData = fSheet.getDataRange().getValues();
      for (let i = 1; i < fData.length; i++) {
        let timestamp = new Date(fData[i][0]).getTime();
        let name = String(fData[i][1]).trim();
        let qTitle = fData[i][2];
        let fileUrl = fData[i][4];
        let partType = fData[i][5] || '';
        
        if (latestResponses[name] && latestResponses[name].timestamp === timestamp) {
          if (!filesData[name]) filesData[name] = [];
          filesData[name].push({ title: qTitle, url: fileUrl, type: partType });
        }
      }
    }


        for (let name in latestResponses) {
          let userData = {
            name: name,
            date: Utilities.formatDate(new Date(latestResponses[name].timestamp), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
            answers: latestResponses[name].answers.filter(a => a.answer !== ""),
            files: filesData[name] ||[]
          };
          
          if(userType === 'supervisor') dashboardData.supervisors.push(userData);
          else dashboardData.unitHeads.push(userData);
        }
      }

  processData('ردود المشرفين', 'ملفات المشرفين', 'supervisor');
  processData('ردود رؤساء الوحدات', 'ملفات رؤساء الوحدات', 'unit_head');

  return dashboardData;
}

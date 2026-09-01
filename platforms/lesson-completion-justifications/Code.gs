function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('منصة المتابعة المدرسية')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// دالة جلب بيانات المدرسة لضمان التطابق التام وجلب كافة السجلات لنفس الرقم
function getSchoolData(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('البيانات الاساسية');
  const data = sheet.getDataRange().getValues();
  const searchId = String(id).trim().toLowerCase(); 
  
  let records = []; 
  
  for (let i = 1; i < data.length; i++) {
    const currentId = String(data[i][15]).trim().toLowerCase(); // العمود P 
    
    if (currentId === searchId && searchId !== '') { 
      records.push({
        colA: data[i][0] === '' ? '' : parseValue(data[i][0]), 
        colC: data[i][2] === '' ? '' : parseValue(data[i][2]), 
        colF: data[i][5] === '' ? '' : parseValue(data[i][5]), 
        colG: parseInt(data[i][6]) || 0, 
        day: data[i][12], // العمود M (التاريخ واليوم)
        name: data[i][14], 
        id: data[i][15], 
        gender: data[i][16], 
        level: data[i][17], 
        type: data[i][18], 
        gov: data[i][19] 
      });
    }
  }
  
  if (records.length > 0) {
    return { found: true, records: records };
  } else {
    return { found: false };
  }
}

function parseValue(val) {
  if (val === '' || val === null) return '';
  if (typeof val === 'number') return Number((val * 100).toFixed(2));
  if (typeof val === 'string') return parseFloat(val.replace('%', ''));
  return val; 
}

function getFormHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('الردود');
  return sheet.getRange(1, 1, 1, 4).getValues()[0];
}

function submitResponse(dataArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('الردود');
  
  for (let i = 0; i < dataArray.length; i++) {
    let dataObj = dataArray[i];
    let row = [
      dataObj.justA || '', 
      dataObj.justC || '', 
      dataObj.justF || '', 
      dataObj.justG || '', 
      '', 
      '', 
      dataObj.schoolDay, // التاريخ / اليوم
      dataObj.schoolName, 
      dataObj.schoolId,  // الرقم الوزاري
      dataObj.schoolGender, 
      dataObj.schoolLevel, 
      dataObj.schoolType, 
      dataObj.schoolGov 
    ];
    sheet.appendRow(row);
  }
  
  // تفعيل التحديث الآلي الفوري للمتأخرين بمجرد استلام الرد
  syncPendingSchools();
  return true;
}

// مزامنة المدارس المتاخرة بناءً على دقة (الرقم الوزاري + اليوم)
function syncPendingSchools() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('البيانات الاساسية');
  const respSheet = ss.getSheetByName('الردود');
  
  if (!masterSheet || !respSheet) return;
  
  const masterData = masterSheet.getDataRange().getValues();
  const respData = respSheet.getDataRange().getValues();
  
  // 1. تجميع الردود المستلمة باستخدام مفتاح مركب (الرقم + اليوم)
  let submittedRecords = new Set();
  for (let i = 1; i < respData.length; i++) {
    let id = String(respData[i][8]).trim().toLowerCase(); // العمود I في الردود
    let date = String(respData[i][6]).trim();             // العمود G في الردود
    if (id && date) submittedRecords.add(id + '||' + date);
  }
  
  let pendingRows = [];
  
  // 2. التحقق من السجلات الأساسية (مقارنة المفتاح المركب)
  for (let i = 1; i < masterData.length; i++) {
    let mId = String(masterData[i][15]).trim().toLowerCase(); // العمود P
    let mDate = String(masterData[i][12]).trim();             // العمود M
    
    if (!mId || !mDate) continue;
    
    // الإضافة التلقائية: إذا لم تقم المدرسة بإرسال رد لهذا (اليوم) تحديداً
    if (!submittedRecords.has(mId + '||' + mDate)) {
      let mName = masterData[i][14] || '';      
      let mLevel = masterData[i][17] || '';     
      let mGov = masterData[i][19] || '';       
      let mSector = masterData[i][18] || '';    
      
      // تجهيز الصف حسب الترتيب المطلوب بالضبط
      pendingRows.push([mId, mName, mLevel, mGov, mSector, mDate]);
    }
  }
  
  // 3. التحديث الديناميكي لصفحة المدارس المتاخرة
  let pendingSheet = ss.getSheetByName('المدارس المتاخرة');
  if (!pendingSheet) pendingSheet = ss.insertSheet('المدارس المتاخرة');
  
  // مسح البيانات القديمة لضمان "الحذف الديناميكي" للصفوف التي تم الإجابة عليها
  pendingSheet.clear(); 
  
  // كتابة العناوين الجديدة شاملة "اليوم"
  let headers = ['الرقم الوزاري', 'اسم المدرسة', 'المرحلة الدراسية', 'المحافظة', 'القطاع', 'اليوم'];
  pendingSheet.appendRow(headers);
  pendingSheet.getRange(1, 1, 1, headers.length).setBackground('#005a3c').setFontColor('white').setFontWeight('bold');
  
  // الترحيل السريع للبيانات باستخدام المصفوفات
  if (pendingRows.length > 0) {
    pendingSheet.getRange(2, 1, pendingRows.length, headers.length).setValues(pendingRows);
  }
}

function loginLeader(code) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('القيادات');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] == code) return { role: 'dept_head' };
    if (data[i][3] == code) return { role: 'unit_head', level: data[i][2] }; 
  }
  return { role: null };
}

// الإحصائيات المتقدمة للقيادات
function getLeaderData(role, level) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  syncPendingSchools(); 
  
  const respSheet = ss.getSheetByName('الردود');
  const respData = respSheet.getDataRange().getDisplayValues();
  const headers = respData.length > 0 ? respData[0].filter((_, idx) => idx !== 4 && idx !== 5) : [];
  
  let results = [];
  let submittedRecords = new Set(); // مفتاح مركب للردود
  
  for (let i = 1; i < respData.length; i++) {
    let rowLevel = String(respData[i][10] || '');
    let rowId = String(respData[i][8]).trim().toLowerCase();
    let rowDate = String(respData[i][6]).trim();
    
    if (rowId) submittedRecords.add(rowId + '||' + rowDate);
    
    if (role === 'dept_head' || (role === 'unit_head' && rowLevel.includes(level))) {
      results.push(respData[i].filter((_, idx) => idx !== 4 && idx !== 5));
    }
  }

  const masterSheet = ss.getSheetByName('البيانات الاساسية');
  const masterData = masterSheet.getDataRange().getValues();
  
  let progressStats = {
    all: { totalRecords: 0, uniqueSchools: 0, duplicatesCount: 0, received: 0, pending: [] },
    primary: { totalRecords: 0, uniqueSchools: 0, duplicatesCount: 0, received: 0, pending: [] },
    middle: { totalRecords: 0, uniqueSchools: 0, duplicatesCount: 0, received: 0, pending: [] },
    high: { totalRecords: 0, uniqueSchools: 0, duplicatesCount: 0, received: 0, pending: [] }
  };

  let schoolMap = {};
  
  for (let i = 1; i < masterData.length; i++) {
    let mId = String(masterData[i][15]).trim().toLowerCase();
    let mName = masterData[i][14];             
    let mLevel = String(masterData[i][17]).trim(); 
    let mDate = String(masterData[i][12]).trim();
    
    if (!mId) continue; 

    if (!schoolMap[mId]) {
      schoolMap[mId] = { id: mId, name: mName, level: mLevel, dates: new Set(), totalRecords: 0 };
    }
    schoolMap[mId].dates.add(mDate);
    schoolMap[mId].totalRecords++;
  }

  for (let key in schoolMap) {
    let s = schoolMap[key];
    
    let submittedCount = 0;
    s.dates.forEach(d => {
      if (submittedRecords.has(s.id + '||' + d)) submittedCount++;
    });

    let isFullySubmitted = (submittedCount >= s.dates.size);

    let pendingObj = {
      id: s.id, 
      name: s.name, 
      level: s.level, 
      hasDuplicates: s.totalRecords > 1, 
      totalRecords: s.totalRecords,
      missingCount: s.dates.size - submittedCount 
    };

    function updateStats(cat) {
      cat.totalRecords += s.totalRecords;
      cat.uniqueSchools += 1;
      cat.duplicatesCount += (s.totalRecords - 1);
      if (isFullySubmitted) {
        cat.received += 1;
      } else {
        cat.pending.push(pendingObj);
      }
    }

    updateStats(progressStats.all);

    if (s.level.indexOf('بتدائ') !== -1) {
      updateStats(progressStats.primary);
    } else if (s.level.indexOf('متوسط') !== -1) {
      updateStats(progressStats.middle);
    } else if (s.level.indexOf('ثانو') !== -1) {
      updateStats(progressStats.high);
    }
  }

  return { 
    headers: headers, 
    rows: results, 
    progress: progressStats,
    role: role,
    level: level
  };
}
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function getSheetSafe(ss, names) {
  for (let name of names) {
    let sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  return null;
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('حصر المواد الكيميائية - TBC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 1. جلب البيانات الأساسية والأسئلة بشكل ديناميكي
function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const mainSheet = getSheetSafe(ss,['البيانات الاساسية', 'البيانات الأساسية']);
  
  // جلب كافة العناوين من الصف الأول (بما فيها أسئلة السلامة المرورية)
  const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0].map(String);
  const itemsData = mainSheet.getRange('I2:I').getValues().flat().filter(String);
  const uniqueItems =[...new Set(itemsData)];
  
  return { headers: headers, items: uniqueItems };
}

// 2. تسجيل دخول المدرسة
function loginSchool(ministerialNumber) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const mainSheet = getSheetSafe(ss,['البيانات الاساسية', 'البيانات الأساسية']);
  const data = mainSheet.getRange('A2:C').getValues(); 
  
  const searchNum = String(ministerialNumber).trim();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][2]).trim() === searchNum) {
      return {
        success: true, sector: data[i][0], schoolName: data[i][1], minNum: searchNum
      };
    }
  }
  return { success: false };
}

// 3. ترحيل البيانات الديناميكي (شامل السلامة المرورية والشرط الجديد)
function submitData(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let respSheet = getSheetSafe(ss,['الردود']);
  
  if (!respSheet || respSheet.getLastColumn() === 0) {
    if (!respSheet) respSheet = ss.insertSheet('الردود');
    const mainSheet = getSheetSafe(ss, ['البيانات الاساسية', 'البيانات الأساسية']);
    const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0].map(String);
    respSheet.appendRow(['وقت الرد', 'الإيميل'].concat(mainHeaders));
  }

  const headers = respSheet.getRange(1, 1, 1, respSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const getIdx = (names) => {
    for (let name of names) {
      let idx = headers.findIndex(h => h === name || h.includes(name));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  // سحب البريد الإلكتروني للمستخدم الفعلي
  const userEmail = Session.getActiveUser().getEmail() || "غير متوفر";
  // التوقيت بتوقيت السعودية
  const timestamp = Utilities.formatDate(new Date(), "Asia/Riyadh", "yyyy/MM/dd HH:mm:ss");

  const cTime = getIdx(['وقت', 'تاريخ']); 
  const cEmail = getIdx(['إيميل', 'ايميل', 'بريد', 'البريد']); 
  const cSector = getIdx(['القطاع']);
  const cSchool = getIdx(['المدرسة']);
  const cMinNum = getIdx(['الرقم الوزاري']);
  const cGender = getIdx(['الجنس']);
  const cBuilding = getIdx(['نوع المبنى']);
  const cName = getIdx(['اسم مدخل البيانات']);
  const cPhone = getIdx(['جوال']);
  const cCount = getIdx(['عدد المواد']);
  const cItem = getIdx(['اسم الماده', 'اسم المادة']);
  const cUnit = getIdx(['وحدة القياس']);
  const cCap = getIdx(['سعة العبوة']);
  const cQty = getIdx(['الكمية']);
  const cHasChem = getIdx(['هل توجد مواد كيميائية متوفرة', 'مواد كيميائية متوفرة']);
  
  const baseInfo = formData.baseInfo;
  const trafficInfo = formData.trafficInfo ||[];
  const items = formData.items; 
  const rowsToInsert =[];
  const numCols = headers.length;

  const createRow = (itemObj) => {
    let row = new Array(numCols).fill("");
    if(cTime !== -1) row[cTime] = timestamp;
    if(cEmail !== -1) row[cEmail] = userEmail;
    if(cSector !== -1) row[cSector] = baseInfo.sector;
    if(cSchool !== -1) row[cSchool] = baseInfo.schoolName;
    if(cMinNum !== -1) row[cMinNum] = baseInfo.minNum;
    if(cGender !== -1) row[cGender] = baseInfo.d;
    if(cBuilding !== -1) row[cBuilding] = baseInfo.e;
    if(cName !== -1) row[cName] = baseInfo.f;
    if(cPhone !== -1) row[cPhone] = baseInfo.g;
    if(cCount !== -1) row[cCount] = baseInfo.h;
    if(cHasChem !== -1) row[cHasChem] = baseInfo.hasChemicals;
    
    // ربط أسئلة السلامة المرورية
    trafficInfo.forEach(t => {
      let tIdx = getIdx([t.header]);
      if(tIdx !== -1) row[tIdx] = t.val;
    });

    if(itemObj) {
      if(cItem !== -1) row[cItem] = itemObj.i;
      if(cUnit !== -1) row[cUnit] = itemObj.j;
      if(cCap !== -1) row[cCap] = itemObj.k;
      if(cQty !== -1) row[cQty] = itemObj.l;
    }
    return row;
  };

  if (items && items.length > 0) {
    items.forEach(item => rowsToInsert.push(createRow(item)));
  } else {
    rowsToInsert.push(createRow(null));
  }

  if (rowsToInsert.length > 0) {
    respSheet.getRange(respSheet.getLastRow() + 1, 1, rowsToInsert.length, numCols).setValues(rowsToInsert);
  }
  return true;
}

function verifyLogin(passcode) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const loginSheet = ss.getSheetByName('تسجيل الدخول');
  if (!loginSheet) return false;
  return loginSheet.getRange('D2:D').getValues().flat().filter(String).map(String).includes(String(passcode));
}

// 5. جلب بيانات الداش بورد (تم إصلاحه جذرياً باستخدام getDisplayValues لمنع الانهيار)
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const mainSheet = getSheetSafe(ss,['البيانات الاساسية', 'البيانات الأساسية']);
    const respSheet = getSheetSafe(ss,['الردود']);
    
    let allSchoolsMaster =[];
    let mainHeaders =[];
    let totalSchoolsCount = 0;

    // حماية إضافية: التأكد من وجود بيانات قبل الجلب لمنع انهيار الخادم
    if (mainSheet) {
      const lastRow = mainSheet.getLastRow();
      if (lastRow > 1) {
        allSchoolsMaster = mainSheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues().filter(r => String(r[2]).trim() !== '');
        totalSchoolsCount = allSchoolsMaster.length;
      }
      const lastCol = mainSheet.getLastColumn();
      if (lastCol > 0) {
        mainHeaders = mainSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(String);
      }
    }
    
    if (!respSheet || respSheet.getLastRow() <= 1) {
      return { total: totalSchoolsCount, respondedCount: 0, list:[], missingList: allSchoolsMaster.map(r => ({name: r[1], minNum: r[2]})), details: {} };
    }
    
    const dataRange = respSheet.getDataRange();
    const respValues = dataRange.getValues(); 
    const respDisplayValues = dataRange.getDisplayValues(); 
    
    const headers = respDisplayValues[0].map(h => String(h).trim());
    const respRows = respValues.slice(1);
    const respDisplayRows = respDisplayValues.slice(1);
    
    const getCol = (names) => {
      for (let name of names) {
        let idx = headers.findIndex(h => h === name || h.includes(name));
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const idxMinNum = getCol(['الرقم الوزاري']);
    const idxSchool = getCol(['المدرسة']);
    const idxSector = getCol(['القطاع']);
    const idxGender = getCol(['الجنس']);
    const idxBuilding = getCol(['نوع المبنى']);
    const idxEntryName = getCol(['اسم مدخل البيانات']);
    const idxPhone = getCol(['جوال']);
    const idxTime = getCol(['وقت', 'تاريخ']);
    const idxEmail = getCol(['إيميل', 'ايميل', 'بريد', 'البريد']);
    const idxHasChem = getCol(['هل توجد مواد كيميائية متوفرة', 'مواد كيميائية متوفرة']);
    
    const idxItemName = getCol(['اسم الماده', 'اسم المادة']);
    const idxUnit = getCol(['وحدة القياس']);
    const idxCapacity = getCol(['سعة العبوة']);
    const idxQuantity = getCol(['الكمية']);

    const trafficHeaders =[];
    for(let i = 12; i <= 19; i++) {
      if(mainHeaders[i]) trafficHeaders.push(mainHeaders[i]);
    }

    const groups = {};
    
    respRows.forEach((row, index) => {
      const minNum = String(respDisplayRows[index][idxMinNum]).trim();
      if(!minNum) return;
      
      let timeVal = index; 
      let displayTime = "غير متوفر";

      if (idxTime !== -1) {
        displayTime = respDisplayRows[index][idxTime] || "غير متوفر";
        let t = row[idxTime];
        if (t instanceof Date) {
          timeVal = t.getTime();
        } else {
          let parsed = new Date(String(t).replace(/-/g, '/')).getTime();
          if (!isNaN(parsed)) timeVal = parsed; 
        }
      }
      
      if (!groups[minNum] || timeVal > groups[minNum].timeVal) {
        groups[minNum] = { timeVal: timeVal, displayTime: displayTime, displayRow: respDisplayRows[index], items:[], traffic:[] };
      }
    });
    
    respRows.forEach((row, index) => {
      const minNum = String(respDisplayRows[index][idxMinNum]).trim();
      if(!minNum) return;
      
      let currentTimeVal = index;
      if (idxTime !== -1) {
        let t = row[idxTime];
        if (t instanceof Date) currentTimeVal = t.getTime();
        else {
          let parsed = new Date(String(t).replace(/-/g, '/')).getTime();
          if (!isNaN(parsed)) currentTimeVal = parsed;
        }
      }
      
      if (currentTimeVal === groups[minNum].timeVal) {
        let dRow = respDisplayRows[index]; 
        
        if (idxItemName !== -1 && dRow[idxItemName] && String(dRow[idxItemName]).trim() !== '') {
          groups[minNum].items.push({
            name: dRow[idxItemName], unit: idxUnit !== -1 ? dRow[idxUnit] : '', cap: idxCapacity !== -1 ? dRow[idxCapacity] : '', qty: idxQuantity !== -1 ? dRow[idxQuantity] : ''
          });
        }
        
        if (groups[minNum].traffic.length === 0) {
          trafficHeaders.forEach(th => {
            let cIdx = getCol([th]);
            groups[minNum].traffic.push({
              question: th, answer: cIdx !== -1 ? dRow[cIdx] : 'لم يُجب'
            });
          });
        }
      }
    });
    
    const respondedList =[];
    const details = {};
    
    for (let minNum in groups) {
      const g = groups[minNum];
      const dRow = g.displayRow;
      
      const mappedBase = {
        time: g.displayTime,
        email: idxEmail !== -1 ? dRow[idxEmail] : 'غير متوفر',
        hasChemicals: idxHasChem !== -1 ? dRow[idxHasChem] : 'غير مدخل',
        sector: idxSector !== -1 ? dRow[idxSector] : 'غير مدخل',
        schoolName: idxSchool !== -1 ? dRow[idxSchool] : 'غير مدخل',
        gender: idxGender !== -1 ? dRow[idxGender] : '',
        building: idxBuilding !== -1 ? dRow[idxBuilding] : '',
        entryName: idxEntryName !== -1 ? dRow[idxEntryName] : '',
        phone: idxPhone !== -1 ? dRow[idxPhone] : ''
      };
      respondedList.push({ name: mappedBase.schoolName, minNum: minNum });
      details[minNum] = { base: mappedBase, items: g.items, traffic: g.traffic };
    }

    const respondedKeys = Object.keys(groups);
    const missingList = allSchoolsMaster
      .filter(row => !respondedKeys.includes(String(row[2]).trim()))
      .map(row => ({ name: row[1], minNum: String(row[2]).trim() }));
    
    return { total: totalSchoolsCount, respondedCount: respondedList.length, list: respondedList, missingList: missingList, details: details };
  
  } catch (error) {
    // التقاط أي خطأ في الخادم وإرساله كرسالة نصية لكي تراه في المتصفح ولا يعلق
    throw new Error("حدث خطأ في الخادم: " + error.message);
  }
}
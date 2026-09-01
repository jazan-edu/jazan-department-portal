// أسماء أوراق العمل الثابتة
const SHEET_BASIC_DATA = 'البيانات الاساسية';
const SHEET_RESPONSES = 'الردود';
const SHEET_LOGIN = 'تسجيل الدخول';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('منصة متابعة حالة التقويم الذاتي')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// التحقق من الرقم الوزاري
function verifySchool(ministerialNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BASIC_DATA);
  const data = sheet.getDataRange().getValues();
  
  const input = String(ministerialNum).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const dbValue = String(data[i][0]).trim().toLowerCase();
    if (dbValue === input) {
      return {
        success: true,
        schoolName: data[i][1],
        principalName: data[i][2],
        stage: data[i][3]
      };
    }
  }
  return { success: false, message: 'الرقم الوزاري غير موجود.' };
}

// إرسال النموذج وترحيل البيانات مع الحفاظ على ترتيب الأعمدة
function submitForm(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const basicSheet = ss.getSheetByName(SHEET_BASIC_DATA);
  const responsesSheet = ss.getSheetByName(SHEET_RESPONSES);
  
  const basicData = basicSheet.getDataRange().getValues();
  let fullBasicRow =[];
  const inputNum = String(formData.ministerialNum).trim().toLowerCase();
  
  for (let i = 1; i < basicData.length; i++) {
    if (String(basicData[i][0]).trim().toLowerCase() === inputNum) {
      fullBasicRow = basicData[i].slice(0, 15); // الأعمدة من A إلى O
      break;
    }
  }
  
  if (fullBasicRow.length === 0) return { success: false, message: 'خطأ في جلب البيانات.' };

  // توجيه البيانات بدقة حسب الأعمدة المطلوبة
  const finalRow =[
    ...fullBasicRow,
    formData.q1_date,               // العمود P
    formData.q2_status,             // العمود Q
    formData.problemType || '',     // العمود R
    formData.ticketRaised || '',    // العمود S
    formData.ticketNum || '',       // العمود T
    formData.otherReason || '',     // العمود U (تخصيص للسبب 'أخرى')
    formData.startDate || '',       // العمود V
    formData.noProblemReason || '', // العمود W حصراً (سبب عدم اكتمال نموذج الجاهزية)
    formData.endDate || ''          // العمود X
  ];
  
  responsesSheet.appendRow(finalRow);
  return { success: true };
}

// دخول القيادات
function loginLeader(code) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOGIN);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 2; i < data.length; i++) {
    if (data[i][4] == code) { 
      return { success: true, role: 'unit', name: data[i][2], stage: data[i][3] };
    }
    if (data[i][6] == code) {
      return { success: true, role: 'head', name: data[i][5], stage: 'الكل' };
    }
  }
  return { success: false, message: 'كود الدخول غير صحيح.' };
}

// جلب الإحصائيات (محدثة لدعم قائمتي الردود وعدم الردود)
function getDashboardData(stageFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const basicData = ss.getSheetByName(SHEET_BASIC_DATA).getDataRange().getValues();
  const responsesData = ss.getSheetByName(SHEET_RESPONSES).getDataRange().getValues();
  
  const respondedSet = new Set();
  for (let i = 1; i < responsesData.length; i++) {
    respondedSet.add(String(responsesData[i][0]).trim().toLowerCase());
  }
  
  let stats = { total: 0, responded: 0, missing: 0, missingList: [], respondedList:[] };
  
  for (let i = 1; i < basicData.length; i++) {
    const minNum = String(basicData[i][0]).trim().toLowerCase();
    const schoolName = basicData[i][1];
    const dbStage = basicData[i][3];
    const authority = basicData[i][4] || ''; // السطر الجديد: جلب السلطة من العمود E (الفهرس 4)
    
    if (isStageMatch(dbStage, stageFilter)) {
      stats.total++;
     if (respondedSet.has(minNum)) {
        stats.responded++;
        // التعديل: إرسال البيانات كمصفوفة [رقم، اسم، سلطة]
        stats.respondedList.push([minNum, schoolName, authority]);
      } else {
        stats.missing++;
        // التعديل: إرسال البيانات كمصفوفة [رقم، اسم، سلطة]
        stats.missingList.push([minNum, schoolName, authority]);
      }
    }
  }
  return stats;
}

// منطق الفرز الذكي (محدث ليعالج كافة احتمالات المرحلة الابتدائية)
function isStageMatch(dbStage, filterStage) {
  if (!filterStage || filterStage === 'الكل') return true;
  if (!dbStage) return false;
  
  const normalize = (str) => {
    let s = str.toString().trim();
    // البحث عن جزء من الكلمة للتعامل مع (ابتدائي، إبتدائية، بتدائي، الابتدائية...)
    if (s.includes("بتدائي")) return "ابتدائي";
    if (s.includes("متوسط")) return "متوسط";
    if (s.includes("ثانوي")) return "ثانوي";
    return s;
  };

  return normalize(dbStage) === normalize(filterStage);
}
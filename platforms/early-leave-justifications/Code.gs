function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('منصة حصر تعليق الدراسة')
    // هذا السطر مهم جداً لتوافق الواجهة مع مقاس شاشة الجوال
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// جلب البيانات الأولية من صفحة "تسجيل الدخول" والتحقق من الرقم الوزاري
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // التأكد من أن اسم الصفحة مطابق لاسم الشيت الخاص بك
  const sheet = ss.getSheetByName("تسجيل الدخول"); 
  if (!sheet) return []; // لتجنب الأخطاء إذا كانت الصفحة غير موجودة
  
  const data = sheet.getDataRange().getValues();
  data.shift(); // حذف صف العناوين الأول
  
  return data.map(r => ({
    schoolName: String(r[0] || "").trim(), // العمود A (اسم المدرسة)
    sector: String(r[1] || "").trim(),     // العمود B (القطاع)
    gender: String(r[2] || "").trim(),     // العمود C (الجنس)
    // العمود E (الرقم الوزاري) وهو الفهرس رقم 4، تم تحويله لنص لضمان قبول الأرقام والحروف
    ministerialId: String(r[4] || "").trim() 
  })).filter(item => item.ministerialId !== ""); // استبعاد الصفوف الفارغة
}

// دالة ترحيل البيانات إلى صفحة "بيانات الإنصراف"
function processForm(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("بيانات الإنصراف");
  
  // إنشاء الصفحة إذا لم تكن موجودة مع العناوين
  if (!sheet) {
    sheet = ss.insertSheet("بيانات الإنصراف");
    const headers = [
      "الرقم الوزاري", "اسم المدرسة", "الجنس", "القطاع", "اسم المدير /ة", "رقم جوال المدير /ة", 
      "المرحلة الدراسية", "إجمالي عدد المعلمين /ات", "عدد المعلمين /ات الغائبين", "إجمالي عدد الطلاب", 
      "عدد الطلاب /ات الغائبين /ات", "التاريخ", "زمن انصراف الطلاب /ات", "تم إرسال رسالة؟", "نص الرسالة", 
      "زمن انصراف المعلمين /ات", "سبب الانصراف", "وصف المشكلة", "رقم البلاغ", "سبب آخر"
    ];
    sheet.appendRow(headers);
  }

  sheet.appendRow([
    formData.ministerialId,
    formData.schoolName,
    formData.gender,
    formData.sector,
    formData.managerName,
    formData.managerPhone,
    formData.levels,
    formData.teachersCount,
    formData.absentTeachers,
    formData.studentsCount,
    formData.absentStudents,
    formData.date,
    formData.studentTime,
    formData.isSent,
    formData.msgContent,
    formData.teacherTime,
    formData.reason,
    formData.problemDesc,
    formData.ticketId,
    formData.otherReason
  ]);
  
  return "تم تسجيل وإرسال البيانات بنجاح!";
}
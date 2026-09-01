function processAttendanceFile(base64Data, fileName) {

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    MimeType.MICROSOFT_EXCEL,
    fileName
  );

  const tempFile = DriveApp.createFile(blob);

  const resource = {
    title: tempFile.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  };

  const converted = Drive.Files.copy(resource, tempFile.getId());

  const tempSpreadsheet = SpreadsheetApp.openById(converted.id);
  const sheet = tempSpreadsheet.getSheets()[0];

  const data = sheet.getDataRange().getValues();

  const mainSS = SpreadsheetApp.getActiveSpreadsheet();
  const absSheet = mainSS.getSheetByName("Absences");
  const empSheet = mainSS.getSheetByName("Employees");

  const employees = empSheet.getDataRange().getValues();

  const empMap = {};

  for (let i = 1; i < employees.length; i++) {
    empMap[String(employees[i][0]).trim()] = {
      mobile: employees[i][2],
      email: employees[i][3]
    };
  }

  // مسح البيانات القديمة من شيت الغياب والتأخر
  absSheet.getRange(2, 1,
    Math.max(absSheet.getLastRow(), 2),
    6).clearContent();

  let output = [];

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const nationalId = String(row[0]).trim();
    const fullName   = row[1];
    const dateStr    = String(row[3]).trim(); // عمود التاريخ (D)
    const baseStatus = String(row[4]).trim(); // عمود حالة التحضير (E)
    
    // تخطي أسطر العناوين المكررة نهائياً
    if (nationalId === "السجل المدني" || nationalId === "id" || fullName === "الإسم" || fullName === "الاسم") {
      continue;
    }

    // جلب التوقيت بشكل نصي نظيف ومسح الفراغات
    const checkInTime  = String(row[6]).trim(); // عمود توقيت الحضور (G)
    const checkOutTime = String(row[8]).trim(); // عمود توقيت الانصراف (I)

    // 1. استبعاد أيام الإجازة الرسمية (الجمعة والسبت) فوراً
    if (dateStr.includes("الجمعة") || dateStr.includes("السبت")) {
      continue;
    }

    let currentStatus = "";

    // 2. فحص الحالات الثلاث وتحديد الوقت المناسب لكل حالة
    if (baseStatus === "غياب") {
      currentStatus = "غياب";
    } 
    else if (checkInTime && checkInTime !== "-" && checkInTime > "07:30") {
      currentStatus = "تأخر";
    } 
    else if (checkOutTime && checkOutTime !== "-" && checkOutTime < "14:30") {
      currentStatus = "انصراف مبكر";
    }

    // إذا تحققت أي حالة من الحالات الثلاث، نقوم بتسجيل السجل
    if (currentStatus !== "") {
      const employee = empMap[nationalId] || {};

      // 🌟 نرسل المصفوفة متكاملة: [السجل, الاسم, التاريخ, الحالة, وقت الحضور, وقت الانصراف, الجوال, البريد]
      output.push([
        nationalId,
        fullName,
        dateStr, 
        currentStatus,
        (checkInTime && checkInTime !== "-") ? checkInTime : "-",
        (checkOutTime && checkOutTime !== "-") ? checkOutTime : "-",
        employee.mobile || "",
        employee.email || ""
      ]);
    }
  }

  // تسجيل البيانات الأساسية في شيت الاكسل للرجوع إليها
  if (output.length > 0) {
    let sheetOutput = output.map(r => [r[0], r[1], r[2], r[3], r[6], r[7]]);
    absSheet.getRange(2, 1, sheetOutput.length, 6).setValues(sheetOutput);
  }

  DriveApp.getFileById(tempFile.getId()).setTrashed(true);
  DriveApp.getFileById(converted.id).setTrashed(true);

  return {
    totalAbsences: output.length,
    records: output
  };
}
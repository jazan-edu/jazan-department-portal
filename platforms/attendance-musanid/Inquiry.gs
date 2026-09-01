function createAndSendInquiries(selectedRecords) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName("InquiryTemplate");
  const logSheet = ss.getSheetByName("المساءلات");

  // إذا لم يتم تمرير سجلات محددة (إرسال الكل)، نقوم بقراءة الشيت بأكمله كما كان سابقاً
  let data = [];
  if (selectedRecords && selectedRecords.length > 0) {
    data = selectedRecords; 
  } else {
    const absSheet = ss.getSheetByName("Absences");
    data = absSheet.getDataRange().getValues();
    data.shift(); // إزالة السطر الأول (العناوين) إذا كنا نقرأ من الشيت مباشرة
  }

  const folderName = "Absence_Inquiries";
  let folderIterator = DriveApp.getFoldersByName(folderName);
  let folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(folderName);

  // المتغير i يبدأ من 0 لأن البيانات الممررة مصفوفة نظيفة، أو تم عمل shift للرأس
  for(let i = 0; i < data.length; i++){

    const nationalId  = data[i][0];
    const fullName    = data[i][1];
    const absenceDate = data[i][2];
    const email       = data[i][5]; // ترتيب العمود السادس (البريد)

    if(!email) continue;

    const refNo =
      "ADM-JZN-" +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy") +
      "-" +
      Utilities.formatString("%04d", i + 1);

    template.getRange("C10").setValue(fullName);
    template.getRange("G10").setValue(nationalId);
    template.getRange("F12").setValue(absenceDate);
    template.getRange("I4").setValue(refNo);
    template.getRange("I5").setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));

    SpreadsheetApp.flush();

    const pdfBlob = exportSheetToPdf_(ss.getId(), template.getSheetId(), fullName, refNo);
    const pdfFile = folder.createFile(pdfBlob);

    GmailApp.sendEmail(
      email,
      "مساءلة نظامية بخصوص غياب",
      "المكرم الزميل, تجدون مرفقاً طيه مساءلة رسمية , نأمل الرد عليها وإعادة إرسالها مرة أخرى على هذا البريد ",
      { attachments:[pdfBlob] }
    );

    logSheet.appendRow([
      refNo,
      nationalId,
      fullName,
      absenceDate,
      email,
      "تم الإرسال",
      pdfFile.getUrl()
    ]);
  }

  return "تم إنشاء وإرسال المساءلات للمحدد بنجاح";
}

// دالة تحويل الـ PDF تبقى كما هي بالأسفل بالداخل أو الخارج
function exportSheetToPdf_(spreadsheetId, sheetId, name, refNo){
  const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?";

  const params = {
    exportFormat: "pdf",
    format: "pdf",
    size: "A4",             // حجم الورقة
    portrait: true,          // وضع طولي
    fitw: true,              // احتواء العرض بالكامل ليملأ الورقة
    fith: false,             // جعلناه false حتى لا يضغط الجدول عمودياً، وسيعتمد على حذف الصفوف الزائدة من الشيت
    sheetnames: false,
    printtitle: false,
    pagenumbers: false,
    gridlines: false,
    
    // إعدادات التوسيط والهوامش
    cpx: "CENTER",           // توسيط المحتوى أفقياً
    cpy: "TOP",              // محاذاة المحتوى للأعلى
    
    gid: sheetId
  };

  const query = Object.keys(params).map(k => k + "=" + params[k]).join("&");
  const token = ScriptApp.getOAuthToken();
  
  const response = UrlFetchApp.fetch(url + query, {
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });

  return response.getBlob().setName(refNo + "_" + name + ".pdf");
}
// دالة إرسال الكل القديمة يتم تعديلها لترسل null حتى تعمل بشكل طبيعي لكامل الشيت
function sendInquiries(){
  google.script.run
    .withSuccessHandler(function(msg){
      alert(msg);
    })
    .createAndSendInquiries(null);
}
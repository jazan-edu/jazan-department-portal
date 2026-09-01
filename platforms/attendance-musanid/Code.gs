function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("نظام إدارة الغياب")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

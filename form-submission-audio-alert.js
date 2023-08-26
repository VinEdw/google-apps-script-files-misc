function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Open Sidebar', 'createSidebar')
    .addToUi();
}

function createSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Submission Alert Sidebar");
  const ui = SpreadsheetApp.getUi();
  ui.showSidebar(html);
}

function getResponseCount() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SS.getSheetByName("Form Responses 1");
  const formUrl = sheet.getFormUrl();
  const form = FormApp.openByUrl(formUrl);
  const responses = form.getResponses();
  const responseCount = responses.length;
  return responseCount;
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const clearButtonRange = sheet.getRange('ClearButton');
  if (e.range.getA1Notation() === clearButtonRange.getA1Notation() && range.isChecked()) {
    SpreadsheetApp.getActiveSheet().getRangeList(['PollResponses', 'TotalParticipants', 'Extras']).clearContent();
    range.uncheck();
  }
}
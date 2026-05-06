 const sheet = SpreadsheetApp.getActiveSheet(); // Get the active sheet

function clearRange(rangeToClear) {
  const range = sheet.getRange(rangeToClear); 
  range.setValue("");
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Condition: Only run if cell 'A1' on 'Sheet1' is edited
  if (sheet.getName() === SHEET_NAME && range.getA1Notation() === CLEAR_BUTTON) {
    clearRange(CLEAR_RANGE);
    const BUTTON_TO_RESET = sheet.getRange("CLEAR_BUTTON");
    BUTTON_TO_RESET.setValue(false)
  }
}


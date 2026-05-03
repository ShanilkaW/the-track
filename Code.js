 const sheet = SpreadsheetApp.getActiveSheet(); // Get the active sheet

function setRangeToFalse(rangeToFalse) {
  const range = sheet.getRange(rangeToFalse); 
  range.setValue(false); 
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Condition: Only run if cell 'A1' on 'Sheet1' is edited
  if (sheet.getName() === "Daily Routines (M)" && range.getA1Notation() === "C34") {
    setRangeToFalse("D3:F31");
    const cell2 = sheet.getRange("C34");
    cell2.setValue(false)
  }
}


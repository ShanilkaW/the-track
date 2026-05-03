// This will give sheet
function getSheet(sheet){
  try{
    let selectedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet)
    return selectedSheet
  } 
  
  catch(error){
    console.log("can not get the sheet due to " + error.stack)
  }
}

// This will give range in any sheet
function getRange(sheet,range) {
  try{
    let dataRange = getSheet(sheet).getRange(range)
    return dataRange
  } 
  
  catch(error){
    console.log("can not get the range due to " + error.stack)
  }
}

// This will give values of any cell or range in any sheet
function getRangeValues(sheet,range) {
  try{
    let dataRange = getRange(sheet,range)
    let values = dataRange.getValues()
    return values
  } 
  
  catch(error){
    console.log("can not get values in the range due to " + error.stack)
  }
}

// This will read configs sheet and return range based on the value cell's formula.
// Way to use dynamic ranges in code
function getRangeFromConfig(key){

  //finding formula against the key
  const rangeAsFormula = getConfigFormula(key)

  //removing = sign in the start from formula
  const rangeWithFullA1Notation = rangeAsFormula.substring(1)

  //check against a reg ex and split sheet range
  const regex = /^'([^']+)'!([A-Z0-9:]+)$/
  const match = rangeWithFullA1Notation.match(regex);
  const sheet = match[1]
  const range = match[2]

  //get and return range
  return getRange(sheet,range)
}

// This will give values of any cell or range in any sheet
function setRangeValues(sheet,range,values) {
  try{
    let dataRange = getRange(sheet,range)
    dataRange.setValues(values)
  } 
  
  catch(error){
    console.log("can not set values in the range due to " + error.stack)
  }
}



//----------------------------Tests----------------------------

function runAllTests(){
  getRange_test_correctCell("Tests", "A1")
  getRange_test_correctRange("Tests", "A1:B3")
  getRange_test_incorrectCell("Tests", "B")
  getRange_test_incorrectRange("Tests", "A1:1B")
}

function readConfigValue_test(){
  let valueFromConfig = readConfigValue("sdf33333")
} 

function getRange_test_correctCell(sheet,cell){
  console.log("Testing with correct cell")
  let correctCell = getRange(sheet,cell)
  console.log(correctCell)
  console.log("-------Testing with correct cell is Over-------")
}

function getRange_test_correctRange(sheet,cell){
  console.log("Testing with correct range")
  let correctRange = getRange(sheet,cell)
  console.log(correctRange)
  console.log("-------Testing with correct range is Over-------")
}

function getRange_test_incorrectCell(sheet,cell){
  console.log("Testing with incorrect cell")
  let incorrectCell = getRange(sheet,cell)
  console.log(incorrectCell)
  console.log("-------Testing with incorrect cell is Over-------")
}

function getRange_test_incorrectRange(sheet,cell){
  console.log("Testing with incorrect range")
  let incorrectRange = getRange(sheet,cell)
  console.log(incorrectRange)
  console.log("-------Testing with incorrect range is Over-------")
}



















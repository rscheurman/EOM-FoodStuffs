//Sheet Refs
var rawFiltered = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RawDataFiltered')
var fsniReport = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FSNI')
var fsniCopyReport = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of FSNI')
var fssiCopyReport = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of FSSI')
var fssiReport = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FSSI')
var rawImport = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RawDataFromXero')

//Data Refs
var rawImportData = rawImport.getRange(2, 20, rawImport.getLastRow(),1)
var rawFilteredData = rawFiltered.getRange(2,1,rawFiltered.getLastRow(),8).getValues();
var fsniReportData = []
var fssiReportData =[]

//Functions
function main() {
  formatReference();
  for(i=0; i<rawFilteredData.length; i++) {
    if(rawFilteredData[i][6].includes('FSNI') && rawFilteredData[i][7] != 'Duplicate') {
      rawFilteredData[i].pop()
      fsniReportData.push(rawFilteredData[i])
    } else if(rawFilteredData[i][6].includes('FSSI') && rawFilteredData[i][7] != 'Duplicate') {
      rawFilteredData[i].pop()
      fssiReportData.push(rawFilteredData[i])
    }
  }
  clearOldValues()
  fsniCopyReport.getRange(3,1,fsniReportData.length,7).setValues(fsniReportData)
  fssiCopyReport.getRange(3,1,fssiReportData.length,7).setValues(fssiReportData)
  formatting()
  sorting()
}

function formatting() {
  //format reference and invoice as text and left-align
  fssiCopyReport.getRange(3,3,fssiCopyReport.getLastRow()).setNumberFormat('@')
  fsniCopyReport.getRange(3,3,fsniCopyReport.getLastRow()).setNumberFormat('@')
  fssiCopyReport.getRange(3,3,fssiCopyReport.getLastRow()).setHorizontalAlignment('left')
  fsniCopyReport.getRange(3,3,fsniCopyReport.getLastRow()).setHorizontalAlignment('left')
  //format dates as dd-MMM-yy .setNumberFormat('d"-"mmm"-"yy');
  fsniCopyReport.getRange(3,4,fsniCopyReport.getLastRow()).setNumberFormat('d"-"mmm"-"yy')
  fssiCopyReport.getRange(3,4,fssiCopyReport.getLastRow()).setNumberFormat('d"-"mmm"-"yy')
  fssiCopyReport.getRange(3,5,fssiCopyReport.getLastRow()).setNumberFormat('d"-"mmm"-"yy')
  fsniCopyReport.getRange(3,5,fsniCopyReport.getLastRow()).setNumberFormat('d"-"mmm"-"yy')
  
  //format amount as currency
  fssiCopyReport.getRange(3,3,fssiCopyReport.getLastRow()).setNumberFormat('"$"#,##0.00')
  fsniCopyReport.getRange(3,3,fsniCopyReport.getLastRow()).setNumberFormat('"$"#,##0.00')
}

function sorting() {
  fssiCopyReport.sort(1)
  fsniCopyReport.sort(1)
}

function clearOldValues() {
  fsniCopyReport.getRange(3,1,fsniCopyReport.getLastRow(),7).clearContent()
  fssiCopyReport.getRange(3,1,fssiCopyReport.getLastRow(),7).clearContent()
}

function formatReference() {
  rawImportData.setNumberFormat('@')
}





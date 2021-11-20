function formatProtectedSheet() {
  var uploadSheetId = PropertiesService.getScriptProperties().getProperty('lastUploadedSheet');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uploadSheet = getSheetById(uploadSheetId);

  // Remove extra space //
  setHeaderRow(uploadSheet);
  deleteExtraColumns(uploadSheet);
  fillColumn(uploadSheetId);
  applyFilter(uploadSheet);
  deleteExtraRows(uploadSheet, getFilteredRowRanges(uploadSheet));

  // Combine rows by attorney name //
  combineRowsByProBono(uploadSheetId);
  sortData(uploadSheet, 'primaryProBono');

  // Create primary key and exceptions object //
  addPrimaryKeys(uploadSheetId);
  setExceptionsProperty(event='upload', type='email');

  // Save sheet properties //
  saveSheetProperties();
  
  Logger.log('Finished formatting uploaded sheet ' + uploadSheetId);
  return true;
}

function validateData() {
  var uploadSheetId = PropertiesService.getScriptProperties().getProperty('lastUploadedSheet');
  var uploadSheet = getSheetById(sheetId);

  // Check case column formatting //
  var [caseColIdx, casesRange, cases] = getColumnCustom(uploadSheetId, 'primaryCase', event='upload');
  for (var i=0; i < cases.length; i++) {
    var row = i + 2;
    var caseName = cases[i].trim();

    // TODO need to handle begin/end/multiple semi-colons //
    while (caseName.endsWith(';')) {
      caseName = caseName.slice(0, caseName.length-1);
    }
    while (caseName.startsWith(';')) {
      caseName = caseName.slice(1, caseName.length);
    }
    if (caseName.match(/;{2,}/)) {
      var groups = caseName.match(/([A-Z -,1-9]*)(;{2,})([A-Z -,1-9]*)/i);
      caseName = groups[1] + ';' + groups[3];
    }
    uploadSheet.getRange(row, caseColIdx + 1, 1, 1).setValue(caseName);
  }
  
  // Check email column formatting //
  var [peIdx, peRange, primaryEmails] = getColumnCustom(uploadSheetId, 'primaryEmail', event='upload');
  //var [meIdx, meRange, managerEmails] = getColumnCustom(uploadSheetId, 'managerEmail', event='upload');     // comment out unless manager emails set on upload //
  isInvalidCell(sheetId, peRange.getA1Notation(), 'upload');
  //isInvalidCell(sheetId, meRange.getA1Notation(), 'upload');

  Logger.log('Finished validating uploaded sheet ' + sheetId);
  return true;
}


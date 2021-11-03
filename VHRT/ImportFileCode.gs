function formatProtectedSheet() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('lastUploadedSheet');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uploadSheet = getSheetById(sheetId);

  // Remove extra space //
  setHeaderRow(uploadSheet);
  deleteExtraColumns(uploadSheet);
  fillColumn(uploadSheet);
  applyFilter(uploadSheet);
  deleteExtraRows(uploadSheet, getFilteredRowRanges(uploadSheet));

  // Combine rows by attorney name //
  combineRowsByProBono(uploadSheet);
  sortData(uploadSheet, 'primaryProBono');

  // Create primary key and exceptions object //
  addPrimaryKeys(uploadSheet);
  setExceptionsProperty(event='upload', type='email');

  // Save sheet properties //
  saveSheetProperties();
  
  Logger.log('Finished formatting uploaded sheet ' + sheetId);
  return true;
}

function validateData() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('lastUploadedSheet');
  var uploadSheet = getSheetById(sheetId);

  // Check case column formatting //
  var [caseColIdx, casesRange, cases] = getColumnCustom(uploadSheet, 'primaryCase');
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
  var [peIdx, peRange, primaryEmails] = getColumnCustom(uploadSheet, 'primaryEmail');
  //var [meIdx, meRange, managerEmails] = getColumnCustom(uploadSheet, 'managerEmail');     // comment out unless manager emails set on upload //
  isInvalidCell(sheetId, peRange.getA1Notation(), 'upload');
  //isInvalidCell(sheetId, meRange.getA1Notation(), 'upload');

  Logger.log('Finished validating uploaded sheet ' + sheetId);
  return true;
}


/* FRAMEWORK */

function doGet(request) {
  return HtmlService.createTemplateFromFile('Index').evaluate();
  //return output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {  // Includes JavaScript.html file //
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}


/* AUTHENTICATION AND DATA RETRIEVAL */

function verifyRegisteredVolunteer([pk, email]) {
  deleteTempProperties();
  var isVerified = false;
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');

  if (protectedSheetId) {
    var protectedSheet = getSheetById(protectedSheetId);
    var protectedData = protectedSheet.getDataRange().getDisplayValues();
    var headers = protectedData[0];

    var [meIdx, meRange, managerEmails] = getColumnCustom(protectedSheet, 'managerEmail');
    var [peIdx, peRange, primaryEmails] = getColumnCustom(protectedSheet, 'primaryEmail');
    var [pkIdx, pkRange, primaryKeys] = getColumnCustom(protectedSheet, 'primaryKey');

    if (primaryKeys.includes(pk)) {
      if (primaryEmails.includes(email)) {
        isVerified = true;
      } else if (managerEmails.includes(email)) {
        isVerified = true;
        var cacheManagerEmailIdx = getManagerIdx(email);
        PropertiesService.getScriptProperties().setProperty('ManagerEmailIdx', JSON.stringify(cacheManagerEmailIdx));
      }
    }
  }

  Logger.log('User opened protected sheet: ' + protectedSheetId);
  Logger.log(JSON.stringify({'isVerified':isVerified, 'primaryKey':pk, 'email':email}));
  return [isVerified, pk, email];
}

function getTableData(pk) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tableCols = JSON.parse(PropertiesService.getScriptProperties().getProperty('reportColumns'));
  
  setTableProperties(pk);
  var caseNames = JSON.parse(PropertiesService.getScriptProperties().getProperty('caseNames'));
  var firmNameDict = addFirmNameDict(pk);

  return [tableCols, caseNames, firmNameDict, pk];
}


/* LOGGING INPUTS TO MAIN REPORT */

function updateAggregateReport([userInputData, pk]) {
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId));
  var reportSheetId = sheetProperties['reportSheet'];
  var reportSheet = getSheetById(reportSheetId);

  var isUpdated;
  var alertMessage = 'Hours successfully reported!';
  var now = Utilities.formatDate(new Date(), 'America/Chicago', 'yyyy-MM-dd HH:mm:ss');
  var casePKs = PropertiesService.getScriptProperties().getProperty('casePKs');
  casePKs = casePKs.length ? JSON.parse(casePKs) : false;

  for (var r=0; r < userInputData.length; r++) {
    var row = userInputData[r];
    if (casePKs) {
      row.push(casePKs[r]);  // sets attorney PK associated with case name from input table -- presumes caseNames and casePKs arrays are in mutual order //
    } else {
      row.push(pk);
    }
    row.push(now);
    reportSheet.appendRow(row);
  }

  isUpdated = checkReportSheetUpdated(reportSheet, userInputData);

  if (!isUpdated) {
    alertMessage = 'Please check your formatting and try again.  If the problem persists, contact lfaulkner@hrionline.org.'
    Logger.log('Data inputs not logged to reporting aggregate sheet');
    return [isUpdated, pk, alertMessage];
  }

  Logger.log('Aggregate report updated: ' + isUpdated.toString());
  Logger.log( 
    casePKs ? 
    '#(' + Object.values(casePKs).filter( (i,n) => {return Object.values(casePKs).indexOf(i) == n} ).toString() + ')' : 
    '#' + pk 
  );
  return [isUpdated, pk, alertMessage];
}


/* RELOADING PAGE */

function reloadPage(request) {
  return 'https://script.google.com/macros/s/AKfycbwQHxFpty2QGxuMLxOG3iRWfQL9KSv2w64uu8fM7nw/dev' // Change to ScriptApp.getService().getUrl() when deployed //
}

function logUserPageReload() {
  Logger.log('User cancelled report submission and reloaded page');
}


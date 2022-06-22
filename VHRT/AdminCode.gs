// Add-on Functions //
// _____________________________________________________________________________________________________________________________________ //

function buildImportSideBar() {
  var ui = SpreadsheetApp.getUi();

  var htmlOutput = HtmlService
    .createTemplateFromFile('ImportFile')
    .evaluate()
    .setTitle('Import a File');

  ui.showSidebar(htmlOutput);
}

function buildSetupSideBar() {
  var ui = SpreadsheetApp.getUi();

  var htmlOutput = HtmlService
    .createTemplateFromFile('Admin')
    .evaluate()
    .setTitle('Activate Sheets for Reporting');

  ui.showSidebar(htmlOutput);
}

function resetExtension() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  // Delete script properties except sheetProperties //
  var keep = [];
  for (var sheet of sheets) {
    keep.push(sheet.getSheetId().toString());
  };
  var allProperties = PropertiesService.getScriptProperties().getProperties();
  for (var k of Object.keys(allProperties)) {
    keep.includes(k) ? null : PropertiesService.getScriptProperties().deleteProperty(k);
  }

  // Remove settings //
  removeDataValidations();
  for (var sheet of sheets) {
    sheet.setTabColor(null);
  }
  
  raiseAlert('Success!', 'Volunteer Hours Reporting Tool was reset. Refresh your browser, then run "Activate Sheets" to continue.')
}


// Build Active Sheets //
// _____________________________________________________________________________________________________________________________________ //

function convertFileUpload(file) {
  var data = Utilities.parseCsv(file);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var numSheets = ss.getNumSheets();
  var now = Utilities.formatDate(new Date(), 'America/Chicago', 'MM-dd-yyyy HH:mm:ss');
  var uploadSheet = ss.insertSheet().setName('Uploaded ' + now);
  var uploadSheetId = uploadSheet.getSheetId().toString();
  
  uploadSheet.getRange(1,1, data.length, data[0].length).setValues(data);
  setScriptProperty('lastUploadedSheet', uploadSheetId);

  Logger.log('Uploaded ' + file);
  Logger.log('Created new sheet: ' + ss.getSheets().some( sheet => sheet.getSheetName() == uploadSheet.getSheetName()));
  Logger.log('Set property "lastUploadedSheet": ' + Boolean(PropertiesService.getScriptProperties().getProperty('lastUploadedSheet')));

  return uploadSheet.getSheetName();
}


function finalizeActiveSheets(eventObj) {
  var event = Object.keys(eventObj)[0];
  var reportHeaders = Object.values(eventObj)[0];
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var protectedSheet = getSheetById(protectedSheetId);
  var protectedSheetName = protectedSheet.getSheetName();
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId));
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (event == 'insert') {
    resetActiveReport(protectedSheetId);
    var reportSheetId = insertActiveReport(protectedSheetId, reportHeaders);
    var reportSheetName = getSheetById(reportSheetId).getSheetName();
    insertActiveProtectedSheet(protectedSheetId);
  }  

  if (event == 'update') {
    var reportSheetId = sheetProperties['reportSheet'];
    var reportSheet = getSheetById(reportSheetId);
    var reportSheetName = reportSheet.getSheetName();

    // Set report columns from current active report sheet //
    var reportColumns = reportSheet.getRange(1, 1, 1, reportSheet.getLastColumn()).getValues()[0];
    reportColumns = reportColumns.filter( (val, idx, arr) => {
      return !(['PRIMARY KEY', 'Timestamp'].includes(val));
    });
    setScriptProperty('reportColumns', reportColumns);

    // Set data validations and formatting //
    setSheetDataValidations(protectedSheetId, ['primaryKey', 'primaryCase'], headers=true);
    setSheetDataValidations(reportSheetId, [], headers=true);
    formatSheet(protectedSheetId, 'protected');
    formatSheet(reportSheetId, 'report');
  }

  var activeSheetObj = {};
  activeSheetObj[event] = [protectedSheetName, reportSheetName];
  
  return activeSheetObj;
}


// Edit Active Sheets //
// _____________________________________________________________________________________________________________________________________ //

function closeActiveSheets() {
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  resetActiveReport(protectedSheetId);
  resetActiveProtectedSheet(protectedSheetId);
  
  var isReset = !(
    JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId))['reportSheet'] && 
    PropertiesService.getScriptProperties().getProperty('protectedSheet')
  );
  Logger.log('Active sheets closed: ' + isReset.toString());
  
  // added to handle user CTRL-Z on sheet deletion //  NEED THIS?

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rs = ss.getSheetByName('Report Summary');
  if (rs) {
    Logger.log('Active sheets closed: missing script property for Report Summary sheet');  // Must log before delete //
    rs.setName(reportRenamed);
  } 
}


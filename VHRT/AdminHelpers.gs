// Building Add-on html //

function getActiveSheetProperties(sheetName=null) {
  var activeSheetExists;

  if (!sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
    var activeSheetName = sheetId ? getSheetById(sheetId).getSheetName() : null;

    if (!activeSheetName) {
      return null;
    } else {
      var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
      activeSheetExists = true;
    }

  } else {
    var activeSheetName = sheetName;
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName);
    var sheetId = activeSheet.getSheetId();
    var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
    activeSheetExists = false;

    PropertiesService.getScriptProperties().setProperty('protectedSheet', sheetId.toString());
    setExceptionsProperty(event="setup", type="email");
  }

  var entries = Object.entries(sheetProperties);

  var formatEntries = {
    'primaryKey':'<b>Primary keys:</b><br>',
    'primaryCase':'<b>Case names:</b><br>',
    'primaryFirm':'<b>Law firms:</b><br>',
    'primaryEmail':'<b>Attorney emails:</b><br>',
    'managerEmail':'<b>Office manager emails:</b><br>',
    'primaryProBono':'<b>Attorney names:</b><br>'
  };
  
  var formattedProperties = [];
  for (var entry of entries) {
    var propertyKey = entry[0];
    var property = entry[1];
    if (Object.keys(formatEntries).includes(propertyKey)) {
      formattedProperties.push(formatEntries[propertyKey] + '"' + entry[1] + '"');
    }
  }

  var activeSheetProperties = {};
  activeSheetProperties[activeSheetName] = formattedProperties;
  activeSheetProperties.activeSheetExists = activeSheetExists;

  if (!activeSheetExists) {
    var now = Utilities.formatDate(new Date(), 'America/Chicago', 'MM-dd-yyyy HH:mm:ss');
    activeSheet.setName('Protected ' + now);
  }
      
  return activeSheetProperties;
}

function createRadioButtons(elementId) {
  var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var radios = {};
  var result = {};

  if (elementId == 'b-sel-sheet') {
    radios['sheets'] = ss.getSheets().map(s => s.getSheetName());

  } else {
    var dateRangeString = getDateRange(); 
    var autoReportCols = ['Attorney', 'Hours spent on case between ' + dateRangeString, 'Billing Rate (hr)'];
    var manReportCols = function() {
      var headers = getSheetById(sheetId).getDataRange().getValues()[0];
      headers.splice(headers.indexOf(sheetProperties['primaryKey']), 1);
      return headers;
    }

    radios['autoCols'] = autoReportCols;
    radios['manCols'] = manReportCols();
  }
  
  result[elementId] = radios;

  Logger.log(result);
  return result;
}


// Building admin sheets //
function getSheetById(sheetId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet;
  for (s of ss.getSheets()) {
    if (s.getSheetId() == sheetId) {
      sheet = s;
      break
    }
  }
  return sheet;
}

function reorderCols(colsFirst, colArr) {
  var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
  var count = 0;
  for (var col of colsFirst) {
    val = sheetProperties[col];
    if (val) {
      var colIdx = colArr.indexOf(val);
      colArr.splice(colIdx, 1);
      colArr.splice(count, 0, val);
      count += 1;
    } else {
      Logger.log(col + ' is not a set script property');
    }
  }
  return colArr;
}


// Permissions & Validations //

function setPermissions(activeSheet) {
   var protection = activeSheet.protect().setDescription('Admin Use Only');
   var admin = PropertiesService.getScriptProperties().getProperty('adminUser');
   var superUser = PropertiesService.getScriptProperties().getProperty('superUser');

   protection.addEditors([admin, superUser]);
   Logger.log(activeSheet.getSheetName() + ' protected for admin use only');   
}

function setDataValidation(activeSheetId, range) {
  var activeSheet = getSheetById(activeSheetId);
  var a1range = range.getA1Notation();
  var a1List = getCommaSepRange(a1range);
  var validations = [];
    
  for (var i=0; i < a1List.length; i++) {
    var cellText = activeSheet.getRange(a1List[i]).getDisplayValue();
    var ruleContainsText = SpreadsheetApp.newDataValidation().requireTextEqualTo(cellText).setAllowInvalid(false).build();
    validations.push([ruleContainsText]);
  }

  try {
    // 2D array for column ranges //
    activeSheet.getRange(a1range).setDataValidations(validations);
  } catch (e) {
    // 1D array for row ranges //
    activeSheet.getRange(a1range).setDataValidations([validations.flat()]);
  }
  
  Logger.log('Set active sheet ' + activeSheet.getSheetName() + ' validations: ' + range.getA1Notation());
}


// Toast helpers //

function logToDebugger(object) {
  Logger.log(object);
}

function raiseAlert(alertTitle, alertString, buttons='OK') {
  var ui = SpreadsheetApp.getUi();
  var enums = {
    'OK':ui.ButtonSet.OK, 
    'OK_CANCEL':ui.ButtonSet.OK_CANCEL,
    'YES_NO':ui.ButtonSet.YES_NO,
    'YES_NO_CANCEL':ui.ButtonSet.YES_NO_CANCEL
  }
  
  var response = ui.alert(alertTitle, alertString, enums[buttons]);
  return response;
}


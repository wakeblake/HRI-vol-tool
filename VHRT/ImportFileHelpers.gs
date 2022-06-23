/* Properties helpers */

function selectColumnName(elementId, event="upload") {
  var idDict = {'ccb':'primaryCase', 'ecb':'primaryEmail', 'lfb':'primaryFirm', 'pbb':'primaryProBono', 'meb':'managerEmail'};
  var sheetId = event == "upload" ? 
    PropertiesService.getScriptProperties().getProperty('lastUploadedSheet') : 
    PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetById(sheetId); 

  var values = sheet.getActiveRange().getValues();
  var value = values[0][0];
  var headerRow = sheet.getActiveRange().getRow();

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'You\'ve selected "' + value + '"',
    'Is this correct?', 
    ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var propertyName = idDict[elementId];
    var data = value ? value : idDict[elementId];

    setScriptProperty(propertyName, data);
    setScriptProperty('headerRow', headerRow);
    Logger.log('Set attribute ' + propertyName + ': ' + PropertiesService.getScriptProperties().getProperty(propertyName));
    return true;
  }

  return false;
}

function saveSheetProperties(sheetKey='lastUploadedSheet') {
  var sheetId = PropertiesService.getScriptProperties().getProperty(sheetKey);
  var properties = PropertiesService.getScriptProperties().getProperties();
  var keepProperties = ['primaryFirm','primaryKey','primaryEmail','primaryProBono','primaryCase', 'managerEmail'];
  var sheetProperties = PropertiesService.getScriptProperties().getProperty(sheetId);

  sheetProperties ? null : setScriptProperty(sheetId, {});
  sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));

  Object.entries(properties).forEach(
    e => keepProperties.includes(e[0]) ? sheetProperties[e[0]] = e[1] : null
  );

  setScriptProperty(sheetId, sheetProperties);
  var isSaved = Boolean(JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId)));
  Logger.log('Saved sheet properties: ' + isSaved);
}

function setExceptionsProperty(event, type) {
  var sheetId = (event == 'upload') ? 
    PropertiesService.getScriptProperties().getProperty('lastUploadedSheet') : 
    PropertiesService.getScriptProperties().getProperty('protectedSheet');

  var exceptions = JSON.parse(PropertiesService.getScriptProperties().getProperty('exceptions'));

  if (!exceptions[type]) {
    exceptions[type] = {};
    setScriptProperty('exceptions', exceptions);
    Logger.log('Created attribute "exceptions", type "' + type + '"')   
    return;
  }
}


/* Formatting helpers */

function getColumnCustom(sheet, colNamePropertyKey, event='upload') {
  try {
    var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheet.getSheetId()));
    var colName = sheetProperties[colNamePropertyKey];
  } catch (err) {
    // TODO check if this is needed //
    var colName = PropertiesService.getScriptProperties().getProperty(colNamePropertyKey);
  }
  //if (event == 'upload') {
  //  var colName = PropertiesService.getScriptProperties().getProperty(colNamePropertyKey);
  //} else {
  //  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheet.getSheetId()));
  //  var colName = sheetProperties[colNamePropertyKey];
  //}
  var data = sheet.getDataRange().getValues();
  var colIdx = data[0].indexOf(colName);
  var colRange = sheet.getRange(2, colIdx + 1, data.length-1,1);  // column range excluding header //
  var colData = colRange.getValues().flat();   // column data excluding header as 1D list //
  return [colIdx, colRange, colData];
}

function applyFilter(sheet=getSheetById(PropertiesService.getScriptProperties().getProperty('lastUploadedSheet'))) {
  var primaryCase = PropertiesService.getScriptProperties().getProperty('primaryCase');
  var primaryProBono = PropertiesService.getScriptProperties().getProperty('primaryProBono');
  var cols = sheet.getDataRange().getValues()[0];
  var caseIdx = cols.indexOf(primaryCase);
  var ppbIdx = cols.indexOf(primaryProBono);
  var filter = sheet.getDataRange().createFilter();
  var filterCriteriaNotEmpty = SpreadsheetApp.newFilterCriteria().whenCellNotEmpty();
  var filterCriteriaEmpty = SpreadsheetApp.newFilterCriteria().whenCellEmpty();
  var filterName = filter.setColumnFilterCriteria(ppbIdx + 1, filterCriteriaNotEmpty); 
  var filterCase = filter.setColumnFilterCriteria(caseIdx + 1, filterCriteriaEmpty);
}

function getFilteredRowRanges(sheet) {
  var visibleA1RowRanges = [];
  var sheetId = sheet.getSheetId();
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();

  var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';   
  var lastColNum = sheet.getLastColumn();
  var lastColLetter = alphabet[lastColNum - 1];

  var fields = "sheets(data(rowMetadata(hiddenByFilter)),properties/sheetId)";
  var sheets = Sheets.Spreadsheets.get(ssId, {fields: fields}).sheets;
  for (var obj of sheets) {
    if (obj.properties.sheetId == sheetId) {
      var data = obj.data;
      var rows = data[0].rowMetadata;
      for (var i=1; i < rows.length; i++) {
        var visibleRowNum = rows[i].hiddenByFilter ? false : i+1;
        if (visibleRowNum) {
          var rowRange = 'A' + visibleRowNum + ':' + lastColLetter + visibleRowNum;   // sheet.getRange() is too slow //
          visibleA1RowRanges.push(rowRange);
        }
      }
      Logger.log(visibleA1RowRanges);
      return visibleA1RowRanges;
    }
  }
}

function fillColumn(sheet) {
  var [ppbIdx, primaryProBonoRange, primaryProBono] = getColumnCustom(sheet, 'primaryProBono');
  var fillCol = [];
  var name = primaryProBono[0];
  primaryProBono.forEach( row => {
    name = row ? row : name;
    row ? fillCol.push([row]) : fillCol.push([name]);
  });
  
  primaryProBonoRange.setValues(fillCol);
}

function setHeaderRow(sheet) {
  var headerRow = PropertiesService.getScriptProperties().getProperty('headerRow');
  if (!(headerRow == 1)) {
    sheet.deleteRows(1, headerRow-1);
  }
}

function sortData(sheet, sortOnProperty, order=true) {
  // assumes headers on top row //
  var column = PropertiesService.getScriptProperties().getProperty(sortOnProperty)
  var headers = sheet.getDataRange().getValues()[0];
  var sortColIdx = headers.indexOf(column);
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());  // exclude headers //
  dataRange.sort({column: sortColIdx + 1, ascending: order});
}

function deleteExtraRows(sheet, a1RowRanges) {
  var rangeList = sheet.getRangeList(a1RowRanges);
  rangeList.clearContent();
  sheet.getFilter() ? sheet.getFilter().remove() : null;

  sortData(sheet, 'primaryProBono');
  var numSortedRows = sheet.getLastRow();
  var numBlankRows = sheet.getMaxRows() - numSortedRows;
  sheet.deleteRows(numSortedRows + 1, numBlankRows);
}

function deleteExtraColumns(sheet) {
  var headers = sheet.getDataRange().getValues()[0];
  var keepProperties = ['primaryCase', 'primaryEmail', 'primaryProBono', 'primaryFirm', 'managerEmail'];
  var keepColNames = keepProperties.map( prop => PropertiesService.getScriptProperties().getProperty(prop) );
  var deleteColNums = headers.map( 
    colName => keepColNames.includes(colName) ? false : headers.indexOf(colName) + 1
  ).filter( colName => colName);

  deleteColNums.reverse().forEach( (i) => sheet.deleteColumn(i));  // must delete in reverse order to preserve column index //
}

function combineRowsByProBono(sheet) {
  var [caseColIdx, caseRange, cases] = getColumnCustom(sheet, 'primaryCase');
  var [ppbIdx, primaryProBonoRange, primaryProBono] = getColumnCustom(sheet, 'primaryProBono');
  var ppbDict = {};
  var currRow = 2
  for (var name of primaryProBono) {
    ppbDict[name] ? ppbDict[name] = ppbDict[name] + 1 : ppbDict[name] = 1
  }
  for (var key of Object.keys(ppbDict)) {
    var firstRow = primaryProBono.indexOf(key) + 2;
    var lastRow = primaryProBono.lastIndexOf(key) + 2;
    if (firstRow == lastRow) {
      currRow += 1;
      continue;
    }
    var casesSubset = sheet.getRange(currRow, caseColIdx + 1, (lastRow-firstRow) + 1, 1).getValues().flat();
    var subsetLength = casesSubset.length;
    var casesCombined = casesSubset.join(';');
    var firstRowValues = sheet.getRange(currRow, 1, 1, sheet.getLastColumn()).getValues().flat();

    firstRowValues.splice(caseColIdx, 1, casesCombined);
    sheet.deleteRows(currRow, subsetLength);
    sheet.appendRow(firstRowValues);
  }
}

function addPrimaryKeys(sheet) {
  var [ppbIdx, primaryProBonoRange, primaryProBono] = getColumnCustom(sheet, 'primaryProBono');
  var data = sheet.getDataRange().getValues();
  var primaryKeys = [];
  var primaryKeyName = 'PRIMARY KEY';

  for (var name of primaryProBono) { 
    var key = generatePk('0123456789', 9);
    key = primaryKeys.flat().includes(key) ? generatePk('0123456789', 9) : key;
    primaryKeys.push(key);
  }

  for (var i=0; i < data.length; i++) {
    if (i == 0) {
      data[i].splice(0, 0, primaryKeyName);
      continue;
    }
    data[i].splice(0, 0, primaryKeys[i - 1]);  // data includes header row but pk does not //
  }

  var updateRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn() + 1);
  updateRange.setValues(data);

  setScriptProperty('primaryKey', primaryKeyName);
}

function generatePk(chars, len) {
  var key = '';
  for (var i=0; i < len; i++) {
    if (i % 3 == 0 && i != 0) {
      key += '-';
    }
    key += chars.charAt(Math.floor(Math.random() * len));
  }
  return key
};

function getDateRange(){
  var now = new Date();
  var year = now.getFullYear();
  var lastYear = year - 1;
  var nextYear = year + 1;
  var cutOff = new Date(year + '-07-01T12:00:00');
  if (now >= cutOff) {
    return '07/01/' + year + ' - 06/30/' + nextYear;
  } else {
    return '07/01/' + lastYear + ' - 06/30/' + year;
  }
}





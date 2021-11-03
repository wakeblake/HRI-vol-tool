// onEdit trigger helper //

function isInvalidCell(sheetId=PropertiesService.getScriptProperties().getProperty('lastUploadedSheet'), a1range, event) {
  var adminLoggerUrl = PropertiesService.getScriptProperties().getProperty('adminLoggerUrl');
  var exceptions = JSON.parse(PropertiesService.getScriptProperties().getProperty('exceptions'));
  var sheet = getSheetById(sheetId);
  var a1List = getCommaSepRange(a1range);

  // check for and highlight invalid emails //

  var noErrors = [];
  var displayErrors = {};
  var headers = sheet.getDataRange().getValues()[0];
  var email;

  for (var a1 of a1List) {
    email = sheet.getRange(a1).getValue();
    
    // handles edit events //
    if (event == 'edit') {
      var primaryEmail = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId))['primaryEmail'];
      var managerEmail = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId))['managerEmail'];
      email = (sheet.getRange(a1).getColumn() == headers.indexOf(primaryEmail) + 1) || 
              (sheet.getRange(a1).getColumn() == headers.indexOf(managerEmail) + 1) ? 
              sheet.getRange(a1).getValue() : 'NA';
    }

    // handles upload events - assumes data type email //
    if (email == 'NA') {
      continue;

    } else if (email && exceptions['email'].includes(email)) {
      continue;

    } else {
      var loggerSheet = SpreadsheetApp.openByUrl(adminLoggerUrl).getSheets()[0];
      var inputCell = loggerSheet.getRange(1, 1);
      var isEmail = loggerSheet.getRange(1,2); 

      inputCell.setValue(email);

      if(!isEmail.getValue()) {

        /* TODO clean this up 
        var lastRow = loggerSheet.getLastRow();
        var insertRange = loggerSheet.getRange(lastRow + 1, 1, 1, 4);
        var now = Utilities.formatDate(new Date(), 'America/Chicago', 'yyyy-MM-dd HH:mm:ss');
        loggerSheet.getRange(insertRange.getA1Notation()).setValues([[now, email, isEmail.getDisplayValue(), event]]);
        */

        email ? displayErrors[a1] = email : displayErrors[a1] = '{NULL}';

      } else {
        noErrors.push(a1);
      }
    }
  }
  
  // (un)highlight errors //
  if (Object.keys(displayErrors).length) {
    sheet.getRangeList(Object.keys(displayErrors)).setBackground('#FFB6C1');
  }
  
  if (noErrors.length) {
    sheet.getRangeList(noErrors).setBackground('#FFFFFF');
  }

  // raise alert for errors and save exceptions on user direction //

  if (Object.keys(displayErrors).length) {
    var errors = Object.values(displayErrors).join(', ');
    var cells = Object.keys(displayErrors).join(', ');
    var ui = SpreadsheetApp.getUi();
    var response = raiseAlert(
      'Possible invalid emails highlighted in cells ' + cells + ':',
      errors + '\r\n' + '\r\n' +
      'Click "OK" to return to the sheet or "CANCEL" to ignore this alert for these emails in the future.',
      buttons='OK_CANCEL'
    );
    
    if (response == ui.Button.CANCEL) {
      Object.entries(displayErrors).forEach(e => {
        if (!(e[1] == '{NULL}')) {
          exceptions['email'].push(e[1]);
          sheet.getRange(e[0]).setBackground('#FFFFFF');
        }
      })
      PropertiesService.getScriptProperties().setProperty('exceptions', JSON.stringify(exceptions));
    }
  }
}



/* Data retreival helpers */

function setTableProperties(pk) {   // Assumes cases per attorney grouped by attorney in same cell in sheet //
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId));
  var managerEmailIdxJSON = PropertiesService.getScriptProperties().getProperty('managerEmailIdx');

  var protectedSheet = getSheetById(protectedSheetId);
  var [caseNameIdx, caseNameRange, cases] = getColumnCustom(protectedSheet, 'primaryCase');
  var [pkIdx, pkRange, primaryKeys] = getColumnCustom(protectedSheet, 'primaryKey');

  if (managerEmailIdxJSON) {
    var managerEmailIdx = JSON.parse(managerEmailIdxJSON);
    var caseNames = [];
    var casePKs = [];

    for(var i of managerEmailIdx) {
      var caseNamesString = protectedSheet.getRange(i, caseNameIdx+1).getValue();
      caseNamesString.split(';').forEach( name => {
        caseNames.push(name);
        casePKs.push(protectedSheet.getRange(i, pkIdx+1).getValue());
      })
    }

  } else {
    var rowIdx = primaryKeys.indexOf(pk);
    var caseNamesString = protectedSheet.getRange(rowIdx+2, caseNameIdx+1).getValue();
    var caseNames = caseNamesString.split(';');
  }

  PropertiesService.getScriptProperties().setProperties({'caseNames': JSON.stringify(caseNames), 'casePKs': JSON.stringify(casePKs)});

  Logger.log('Set table property "caseNames": ' + PropertiesService.getScriptProperties().getProperty('caseNames'));
  Logger.log('Set table property "casePKs": ' + PropertiesService.getScriptProperties().getProperty('casePKs'));
}

function getManagerIdx(email){
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = PropertiesService.getScriptProperties().getProperty(protectedSheetId);
  var protectedSheet = getSheetById(protectedSheetId);

  var managerEmailCol = sheetProperties['managerEmail']; 
  var [meIdx, meRange, managerEmails] = getColumnCustom(protectedSheet, 'managerEmail');
  var isManagerIdx = [];
  for (var i=0; i < managerEmails.length; i++) {
    if(managerEmails[i] == email) {
      isManagerIdx.push(i+2);
    }
  }
  Logger.log('User is office manager: ' + JSON.stringify(isManagerIdx));
  return isManagerIdx;
}

function addFirmNameDict(pk) {
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId));
  var protectedSheet = getSheetById(protectedSheetId);

  var [pkIdx, pkRange, primaryKeys] = getColumnCustom(protectedSheet, 'primaryKey');
  var [primaryFirmIdx ,primaryFirmRange, firms] = getColumnCustom(protectedSheet, 'primaryFirm');

  var rowIdx = primaryKeys.indexOf(pk);
  var primaryFirmCol = sheetProperties['primaryFirm'];
  var firmNameDict = {};
  firmNameDict[primaryFirmCol] = protectedSheet.getRange(rowIdx+2, primaryFirmIdx+1).getValue();

  Logger.log('Added firmNameDict: ' + JSON.stringify(firmNameDict));
  return firmNameDict;
}


/* Logging user inputs helpers */

function checkReportSheetUpdated(reportSheet, userInputData) {
  var isUpdated = true;
  var reportData = reportSheet.getDataRange().getDisplayValues();
  for (var i=0; i < userInputData.length; i++) {
    for(var j = reportData.length-1; j > 0; j--) {
      if (JSON.stringify(userInputData[i]) == JSON.stringify(reportData[j])) {
        break;
      }
    }
    if (isUpdated == true) {
      continue
    };
    isUpdated = false;
    break
  }
  return isUpdated;
}

function getCommaSepRange(a1range) {
  var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  var re = /([A-Z]+)(\d+)/;
  var lst = [];
  var a1List = [];

  a1range.split(':').forEach( i => lst.push([i.match(re)[1], i.match(re)[2]]) );

  var fr = Number(lst[0][1]);
  var lr = lst.length > 1 ? Number(lst[1][1]) : fr;
  var fc = lst[0][0];
  var lc = lst.length > 1 ? lst[1][0] : fc;
  
  for (i=fr; i < lr + 1; i++) {
    for (j=alphabet.indexOf(fc); j < alphabet.indexOf(lc) + 1; j++) {
      a1List.push(alphabet[j] + i);
    }
  }

  return a1List;
}


// General admin maintenance //
function emailOnFailedLogin([email, pk]) {
  var adminUser = PropertiesService.getScriptProperties().getProperty('adminUser');
  var superUser = PropertiesService.getScriptProperties().getProperty('superUser');
  var now = Utilities.formatDate(new Date(), 'America/Chicago', 'MM-dd-yyyy HH:mm:ss');
  GmailApp.sendEmail(
    adminUser, 
    'Failed User Login', 
    'Failed user ' + email + ' attempted login with key ' + pk + ' at ' + now + '.', 
    {from: superUser, replyTo: email}
  );
}

function emailUserSubmission([userInputData, pk]) {
  var adminUser = PropertiesService.getScriptProperties().getProperty('adminUser');
  var superUser = PropertiesService.getScriptProperties().getProperty('superUser');
  var now = Utilities.formatDate(new Date(), 'America/Chicago', 'MM-dd-yyyy HH:mm:ss');
  var reformatDataStr = '';
  var reportColumns = JSON.parse(PropertiesService.getScriptProperties().getProperty('reportColumns'));
  var html;

  userInputData.forEach(row => {
    reformatDataStr = reformatDataStr + '<tr style="border: 1px solid black">';
    for (var i of row) {
      reformatDataStr = reformatDataStr + '<td style="border: 1px solid black">' + i + '</td>';
    }
    reformatDataStr = reformatDataStr + '</tr>';
  });

  reportColumns = reportColumns.map( c => '<th style="border: 1px solid black">' + c + '</th>');
  reportColumns = reportColumns.join('');

  html = 
    '<p>' + 'Volunteer ' + pk + ' reported the following at ' + now + ':</p><br>' +
    '<table style="border: 1px solid black">' + 
      '<thead style="border: 1px solid black">' + 
        '<tr style="border: 1px solid black">' + 
          reportColumns + 
        '</tr>' + 
      '</thead>' + 
      '<tbody>' +
        reformatDataStr + 
      '</tbody>' + 
    '</table>'

  GmailApp.sendEmail(
    adminUser, 
    'Volunteer Has Reported Hours', 
    '',
    {from: superUser,htmlBody: html}
  );
}

function deleteTempProperties() {
  PropertiesService.getScriptProperties().deleteProperty('managerEmailIdx');
  PropertiesService.getScriptProperties().deleteProperty('casePKs');
  PropertiesService.getScriptProperties().deleteProperty('caseNames');
}


// Error toast helpers //

function logErrorFromHTML(pk) {
  Logger.log('ERROR: Failure to save user ' + pk + ' input data');
}




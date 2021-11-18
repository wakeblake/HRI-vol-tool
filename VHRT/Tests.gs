/* Unit Tests -- Run ONLY when active sheets are set */

const QUnit = QUnitGS2.QUnit;
const TESTS = [
    userAuthentication,
    dataRetrievalHelpers,

] 

/*
function doGet(request) {
  QUnitGS2.init();

  TESTS.forEach((testFunc) => {
    testFunc();
  })

  QUnit.start();
  return QUnitGS2.getHtml();
}

function getResultsFromServer() {
  return QUnitGS2.getResultsFromServer();
}
*/

function dataRetrievalHelpers() {
  QUnit.module('Data retrieval helpers');

  var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var meColName = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId))['managerEmail'];
  var testSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(0);
  testSheet.setName('HRI_TESTS');

  var testData = [];
  testData.push(['Col1', meColName, 'Col3', 'Col4', 'Col5']);
  for (var i=1; i < 11; i++) {
    testData.push([i*1, i*2, i*3, i*4, i*5]);
  }
  testSheet.getRange(1,1,11,5).setValues(testData);
  testSheet.getRange(2,2,3,1).setValues([['wakeblake@gmail.com'], ['wakeblake@gmail.com'], ['wakeblake@gmail.com']])
  PropertiesService.getScriptProperties().setProperty('testColumnName', 'Col4');

  QUnit.test('Get column data', assert => {
    assert.equal(getColumnCustom(testSheet, 'testColumnName'), [3, testSheet.getRange(2, 4, testSheet.getLastRow(), 1), [4, 8, 12, 16, 20, 24, 28, 32, 36, 40]]);
  })

  QUnit.test('Get manager indices', assert => {
    assert.equal(getManagerIdx('wakeblake@gmail.com'), [2,3,4]);
  })
}


function userAuthentication() {
  QUnit.module('Authentication (random draws)');
  var protectedSheet = getSheetById(PropertiesService.getScriptProperties().getProperty('protectedSheet'));
  var [peIdx, peRange, primaryEmails] = getColumnCustom(protectedSheet, 'primaryEmail', 'test');
  var [pkIdx, pkRange, primaryKeys] = getColumnCustom(protectedSheet, 'primaryKey', 'test');

  for (var i=0; i < 5; i++) {
    var n = Math.floor(Math.random() * primaryKeys.length);
    var pk = primaryKeys[n];
    var email = primaryEmails[n];

    QUnit.test('Active user pk and email', assert => {
      assert.ok(verifyRegisteredVolunteer([pk, email])[0], 'Active user was not verified');
    })

    QUnit.test('Active user wrong pk', assert => {
      var pkErr = '000-000-001';
      assert.ok(!verifyRegisteredVolunteer([pkErr, email])[0], 'Active user verified with failed login')
    })

    QUnit.test('Unknown user and pk', assert => {
      emailErr = 'noemail@noemail.net';
      assert.ok(!verifyRegisteredVolunteer([pk, emailErr])[0], 'Unknown user was verified');
    })
  }
}


function endTestReset() {
  var testSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HRI_TESTS');
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(testSheet);
  PropertiesService.getScriptProperties().deleteProperty('testColumnName');
}


function test() {
  //PropertiesService.getScriptProperties().setProperty('managerEmail', 'Manager Emails');
  //console.log(PropertiesService.getScriptProperties().getProperties());
  console.log(getManagerIdx('wakeblake@gmail.com'));
  console.log(SpreadsheetApp.getActiveSheet().getSheetName());
  
}

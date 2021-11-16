const QUnit = QUnitGS2.QUnit;
const TESTS = [
    userAuthentication,
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
  var sheet = Sheets.newSpreadsheet().sheets()[0];
  var testData = [];
  testData.push(['Col1', 'Col2', 'Col3', 'Col4', 'Col5']);
  for (var i=1; i < 11; i++) {
    testData.push([i*1, i*2, i*3, i*4, i*5]);
  }
  sheet.setValues(testData);
  PropertiesService.getScriptProperties().setProperty('testColumnName', 'Col4')

  QUnit.test('Get column data', assert => {
    assert.ok(getColumnCustom(sheet, 'testColumnName') == [3, sheet.getRange(2, 4, sheet.getLastRow(), 1), [4, 8, 12, 16, 20, 24, 28, 32, 36, 40]]);
  })
  
}

function userAuthentication() {
  QUnit.module('Authentication (random draws)');
  var protectedSheet = getSheetById(PropertiesService.getScriptProperties().getProperty('protectedSheet'));
  var [peIdx, peRange, primaryEmails] = getColumnCustom(protectedSheet, 'primaryEmail');
  var [pkIdx, pkRange, primaryKeys] = getColumnCustom(protectedSheet, 'primaryKey');

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

/* Unit Tests -- Run ONLY when active sheets are set */

// Set up QUnit and run tests //

const QUnit = QUnitGS2.QUnit;
const TESTS = [
    dataRetrievalHelpers,
    userAuthentication,
] 


function doGet(request) {
  QUnit.config.autostart = false;

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

QUnit.done( details => {
  endTestReset();
})


// QUnit Modules //

function dataRetrievalHelpers() {
  QUnit.module('Data retrieval helpers');

  var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId))
  var testSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(0);
  var testSheetId = testSheet.getSheetId().toString();
  PropertiesService.getScriptProperties().setProperty(testSheetId, JSON.stringify(sheetProperties))
  testSheet.setName('HRI_TESTS');

  var testProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(testSheetId));
  var testData = [];
  testData.push([testProperties['primaryKey'], testProperties['managerEmail'], 'Col3', testProperties['primaryCase'], testProperties['primaryFirm']]);
  for (var i=1; i < 11; i++) {
    testData.push([i*1, i*2, i*3, i*4, i*5]);
  }
  testSheet.getRange(1,1,11,5).setValues(testData);
  testSheet.getRange(2,2,3,1).setValues([['wakeblake@gmail.com'], ['wakeblake@gmail.com'], ['wakeblake@gmail.com']])

  QUnit.test('Get sheet by ID', assert => {
    assert.deepEqual(getSheetById(testSheet.getSheetId()), testSheet);
  })

  QUnit.test('Get column data', assert => {
    assert.deepEqual(getColumnCustom(testSheetId, 'primaryCase', event='test'), [3, testSheet.getRange(2, 4, testSheet.getLastRow(), 1), [4, 8, 12, 16, 20, 24, 28, 32, 36, 40]]);
  })

  QUnit.test('Get manager indices', assert => {
    assert.deepEqual(getManagerIdx(testSheetId, 'wakeblake@gmail.com', event='test'), [2,3,4]);
  })

  QUnit.test('Get firm name from pk', assert => {
    assert.equal(addFirmObj(testSheetId, 7)[testProperties['primaryFirm']], 35);
  })
}

function userAuthentication() {
  QUnit.module('Authentication (random draws)');
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var [peIdx, peRange, primaryEmails] = getColumnCustom(protectedSheetId, 'primaryEmail', 'test');
  var [pkIdx, pkRange, primaryKeys] = getColumnCustom(protectedSheetId, 'primaryKey', 'test');

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


// Test helpers //

function endTestReset() {
  var testSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HRI_TESTS');
  var testSheetId = testSheet.getSheetId().toString();
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(testSheet);
}

function testBuildHelper() {
  console.log(PropertiesService.getScriptProperties().getProperties());
}

/* Unit Tests -- Run ONLY when active sheets are set */

// Set up QUnit//

const QUnit = QUnitGS2.QUnit;
const TESTS = [
    dataRetrievalHelpers,
    userAuthentication,
    dataRetrieval,
] 


// Run tests - COMMENT OUT doGet() TO MAKE WEB APP LIVE //
/*
function doGet(request) {
  setUpTestData();

  QUnit.config.autostart = false;

  QUnitGS2.init();

  TESTS.forEach((testFunc) => {
    testFunc();
  })

  QUnit.start();
  return QUnitGS2.getHtml();
}
*/
function getResultsFromServer() {
  return QUnitGS2.getResultsFromServer();
}

QUnit.done( details => {
  endTestReset();
})


// Set up test data //

function setUpTestData() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
  var testSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(0);
  var testSheetId = testSheet.getSheetId().toString();
  PropertiesService.getScriptProperties().setProperty(testSheetId, JSON.stringify(sheetProperties));
  testSheet.setName('HRI_TESTS');

  var testProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(testSheetId));
  var testData = [];
  testData.push([testProperties['primaryKey'], testProperties['managerEmail'], 'Col3', testProperties['primaryCase'], testProperties['primaryFirm']]);
  for (var i=1; i < 11; i++) {
    testData.push([i*1, i*2, i*3, i*4, i*5]);
    testData[i] = testData[i].map(cell => 'a' + cell.toString());
  }
  testSheet.getRange(1,1,11,5).setValues(testData);
  testSheet.getRange(2,2,3,1).setValues([['wakeblake@gmail.com'], ['wakeblake@gmail.com'], ['wakeblake@gmail.com']])
}

function getTestData() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
  var testSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HRI_TESTS');
  var testSheetId = testSheet.getSheetId().toString();
  var testProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(testSheetId));

  return [sheetId, sheetProperties, testSheet, testSheetId, testProperties];
}


// QUnit Modules //

function dataRetrievalHelpers() {
  var [sheetId, sheetProperties, testSheet, testSheetId, testProperties] = getTestData();
  QUnit.module('Data retrieval helpers');

  QUnit.test('Get sheet by ID', assert => {
    assert.deepEqual(getSheetById(testSheet.getSheetId()), testSheet);
  })

  QUnit.test('Get column data', assert => {
    assert.deepEqual(
      getColumnCustom(testSheetId, 'primaryCase', event='test'), 
      [3, testSheet.getRange(2, 4, testSheet.getLastRow(), 1), ['a4','a8','a12','a16','a20','a24','a28','a32','a36','a40']]
    );
  })

  QUnit.test('Get comma separated range of cells in A1Notation', assert => {
    assert.deepEqual(getCommaSepRange('B4:D10'), ['B4','C4','D4','B5','C5','D5','B6','C6','D6','B7','C7','D7','B8','C8','D8','B9','C9','D9','B10','C10','D10']);
  })

  QUnit.test('Get manager indices', assert => {
    assert.deepEqual(getManagerIdx(testSheetId, 'wakeblake@gmail.com', event='test'), [2,3,4]);
  })

  QUnit.test('Get firm name from pk', assert => {
    assert.equal(addFirmObj(testSheetId, 'a7')[testProperties['primaryFirm']], 'a35');
  })
}

function userAuthentication() {
  QUnit.module('Authentication (random draws)');
  var protectedSheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var [peIdx, peRange, primaryEmails] = getColumnCustom(protectedSheetId, 'primaryEmail', 'test');
  var [pkIdx, pkRange, primaryKeys] = getColumnCustom(protectedSheetId, 'primaryKey', 'test');

  var testSet = [];
  for (var i=0; i < 5; i++) {
    var n = Math.floor(Math.random() * primaryKeys.length);
    var pk = primaryKeys[n];
    var email = primaryEmails[n];
    testSet.push([pk, email]);
  }

  QUnit.test('Active user pk and email', assert => {
    assert.ok(verifyRegisteredVolunteer(testSet[0])[0], 'Active user was not verified');
    assert.ok(verifyRegisteredVolunteer(testSet[1])[0], 'Active user was not verified');
    assert.ok(verifyRegisteredVolunteer(testSet[2])[0], 'Active user was not verified');
    assert.ok(verifyRegisteredVolunteer(testSet[3])[0], 'Active user was not verified');
    assert.ok(verifyRegisteredVolunteer(testSet[4])[0], 'Active user was not verified');
  })

  QUnit.test('Active user wrong pk', assert => {
    var pkErr = '000-000-001';
    assert.ok(!verifyRegisteredVolunteer([pkErr, testSet[0][1]])[0], 'Active user verified with failed login');
    assert.ok(!verifyRegisteredVolunteer([pkErr, testSet[1][1]])[0], 'Active user verified with failed login');
    assert.ok(!verifyRegisteredVolunteer([pkErr, testSet[2][1]])[0], 'Active user verified with failed login');
    assert.ok(!verifyRegisteredVolunteer([pkErr, testSet[3][1]])[0], 'Active user verified with failed login');
    assert.ok(!verifyRegisteredVolunteer([pkErr, testSet[4][1]])[0], 'Active user verified with failed login');
  })

  QUnit.test('Unknown user and pk', assert => {
    emailErr = 'noemail@noemail.net';
    assert.ok(!verifyRegisteredVolunteer([testSet[0][0], emailErr])[0], 'Unknown user was verified');
    assert.ok(!verifyRegisteredVolunteer([testSet[1][0], emailErr])[0], 'Unknown user was verified');
    assert.ok(!verifyRegisteredVolunteer([testSet[2][0], emailErr])[0], 'Unknown user was verified');
    assert.ok(!verifyRegisteredVolunteer([testSet[3][0], emailErr])[0], 'Unknown user was verified');
    assert.ok(!verifyRegisteredVolunteer([testSet[4][0], emailErr])[0], 'Unknown user was verified');
  })
}

function dataRetrieval() {
  var [sheetId, sheetProperties, testSheet, testSheetId, testProperties] = getTestData();
  QUnit.module('Set up user data table');
  
  QUnit.test('Set table properties for attorney user', assert => {
    var pk = 'a9';
    setTableProperties(testSheetId, pk);
    var caseNames = PropertiesService.getScriptProperties().getProperty('caseNames') ?
                    JSON.parse(PropertiesService.getScriptProperties().getProperty('caseNames')) :
                    PropertiesService.getScriptProperties().getProperty('caseNames');
    var casePKs = PropertiesService.getScriptProperties().getProperty('casePKs') ? 
                  JSON.parse(PropertiesService.getScriptProperties().getProperty('casePKs')) :
                  PropertiesService.getScriptProperties().getProperty('casePKs');
    assert.deepEqual(caseNames, ['a36']);
    assert.deepEqual(casePKs, '');
  })

  QUnit.test('Get table data for attorney user', assert => {
    var tableCols = JSON.parse(PropertiesService.getScriptProperties().getProperty('reportColumns'));
    assert.deepEqual(
      getTableData('a5', sheetId=testSheetId), 
      [tableCols, ['a20'], {'Organization Name':'a25'}, 'a5']
    );
  })

  QUnit.test('Set table properties for manager user', assert => {
    PropertiesService.getScriptProperties().setProperty('managerEmailIdx', JSON.stringify([2,3,4]));
    var pk = 'a3';
    setTableProperties(testSheetId, pk);
    var caseNames = PropertiesService.getScriptProperties().getProperty('caseNames') ?
                    JSON.parse(PropertiesService.getScriptProperties().getProperty('caseNames')) :
                    PropertiesService.getScriptProperties().getProperty('caseNames');
    var casePKs = PropertiesService.getScriptProperties().getProperty('casePKs') ? 
                  JSON.parse(PropertiesService.getScriptProperties().getProperty('casePKs')) :
                  PropertiesService.getScriptProperties().getProperty('casePKs');
    assert.deepEqual(caseNames, ['a4','a8','a12']);
    assert.deepEqual(casePKs, ['a1','a2','a3']);
  })

  QUnit.test('Get table data for manager user', assert => {
    var tableCols = JSON.parse(PropertiesService.getScriptProperties().getProperty('reportColumns'));
    assert.deepEqual(
      getTableData('a2', sheetId=testSheetId), 
      [tableCols, ['a4','a8','a12'], {'Organization Name':'a10'}, 'a2']
    );
  })
}


// Test helpers //

function endTestReset() {
  var testSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HRI_TESTS');
  var testSheetId = testSheet.getSheetId().toString();
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(testSheet);
  PropertiesService.getScriptProperties().deleteProperty('caseNames');
  PropertiesService.getScriptProperties().deleteProperty('casePKs');
  PropertiesService.getScriptProperties().deleteProperty('managerEmailIdx');
}

function testBuildHelper() {
  //var testSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HRI_TESTS');
  //var testSheetId = testSheet.getSheetId().toString();
  console.log(PropertiesService.getScriptProperties().getProperties());
}



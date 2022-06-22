function test() {
  //console.log(PropertiesService.getUserProperties().getProperties())
  console.log(PropertiesService.getScriptProperties().getProperties());
  //console.log(PropertiesService.getUserProperties().getProperties());
  //PropertiesService.getScriptProperties().setProperty('protectedSheet', '');
}

function test2() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  //PropertiesService.getScriptProperties().deleteProperty('exceptions');
  //PropertiesService.getUserProperties().deleteAllProperties();
}

function test3() {
  sheet = getSheetById('11026088');
  var [keyIdx, keyRange, keyEmails] = getColumnCustom(sheet, 'primaryEmail');
  console.log(keyRange.getA1Notation());
  var a1List = getCommaSepRange(keyRange.getA1Notation());
  console.log(a1List);
}

function getAuthToken() {
  var token = ScriptApp.getOAuthToken();
  console.log(token);
}



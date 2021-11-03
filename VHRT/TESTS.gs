function test() {
  //console.log(PropertiesService.getUserProperties().getProperties())
  console.log(PropertiesService.getScriptProperties().getProperties());
  //console.log(PropertiesService.getUserProperties().getProperties());
  //PropertiesService.getScriptProperties().setProperty('protectedSheet', '');
}

function test2() {
  //PropertiesService.getScriptProperties().deleteAllProperties();
  //PropertiesService.getScriptProperties().deleteProperty('exceptions');
  PropertiesService.getUserProperties().deleteAllProperties();
}

function test3() {
  //var sheetProps = JSON.parse(PropertiesService.getScriptProperties().getProperty('202187291'))
  //delete sheetProps['reportSheet'];
  //console.log(sheetProps);
  var sheetId = PropertiesService.getScriptProperties().getProperty('lastUploadedSheet');
  var sheet = getSheetById(sheetId);
  console.log(sheet.getRangeList(['A1']).getRanges().map(r => r.getValue()));
}

function getAuthToken() {
  var token = ScriptApp.getOAuthToken();
  console.log(token);
}

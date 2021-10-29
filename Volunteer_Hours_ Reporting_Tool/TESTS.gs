function test() {
  console.log(PropertiesService.getScriptProperties().getProperties());
  //console.log(PropertiesService.getUserProperties().getProperties());
  //var sheetId = PropertiesService.getScriptProperties().getProperty('lastUploadedSheet');
  //PropertiesService.getScriptProperties().setProperty('protectedSheet', '');
}

function test2() {
  //PropertiesService.getScriptProperties().deleteAllProperties();
  PropertiesService.getScriptProperties().deleteProperty('exceptions');
}

function getAuthToken() {
  var token = ScriptApp.getOAuthToken();
  console.log(token);
}

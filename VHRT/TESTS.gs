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
var reportColumns = [1,2,3,4,5];
reportColumns = reportColumns.map( c => '<th>' + c + '</th>');
console.log(reportColumns);
}

function getAuthToken() {
  var token = ScriptApp.getOAuthToken();
  console.log(token);
}



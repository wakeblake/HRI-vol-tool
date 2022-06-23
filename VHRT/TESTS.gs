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
  var pk = '616-680-230';
  var [keyIdx, keyRange, primaryKeys] = getColumnCustom(sheet, 'primaryKey');
  var [mgrIdx, mgrRange, mgrEmails] = getColumnCustom(sheet, 'managerEmail');
  var dict = {}
  primaryKeys.map( (element, i) => {
    return dict[element] = mgrEmails[i];
  });
  console.log(dict[pk]);
}

function getAuthToken() {
  var token = ScriptApp.getOAuthToken();
  console.log(token);
}



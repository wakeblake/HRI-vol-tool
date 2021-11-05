/*  Brute force looping through alphabet to find Google-approved valid email domains.  
    Checks email against Sheets isEmail() cell formula. Format on Sheet:
    Row: [A1: Email][B1: Email][C1: isEmail(A1)][D1: isEmail(B1)]
*/

// Use this to run functions asynchronously via user interface //
function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('dom3_1', 'dom3_1');
  menu.addItem('dom3_2', 'dom3_2');
  menu.addItem('dom3_3', 'dom3_3');
  menu.addItem('dom3_4', 'dom3_4');
  menu.addItem('dom3_5', 'dom3_5');
  menu.addItem('dom2', 'dom2');
  menu.addToUi();
}

// partition alpha among functions - time limits exceeded after 5 partitions, run one init pos letter at a time //
function dom3_1(pos=0) {
  var loggerSheet = SpreadsheetApp.getActiveSheet();
  var inputCells = loggerSheet.getRange(1,1,1,2);
  var isEmails = loggerSheet.getRange(1,3,1,2);
  var alpha = 'abcdefghijklmnopqrstuvwxyz';
  var [fhalf, bhalf] = [ 'abcdefghijklm', 'nopqrstuvwxyz' ]

  var email = 'blake.holleman@gmail.';
  var keep = [];

  for (var i=pos; i < 5; i++) {
    for (var c2 of alpha) {
      for (var j=0; j < fhalf.length; j++) {
        inputCells.setValues([[email + alpha[i] + c2 + fhalf[j], email + alpha[i] + c2 + bhalf[j]]]);
        var [[e1, e2]] = isEmails.getValues();
        if (e1) {
          keep.push(alpha[i] + c2 + fhalf[j]);
        } else if (e2) {
          keep.push(alpha[i] + c2 + bhalf[j]);
        }
      }
    }
    break;
  }
  Logger.log(keep);
}

function dom3_2(pos=5) {
  var loggerSheet = SpreadsheetApp.getActiveSheet();
  var inputCells = loggerSheet.getRange(2,1,1,2);
  var isEmails = loggerSheet.getRange(2,3,1,2);
  var alpha = 'abcdefghijklmnopqrstuvwxyz';
  var [fhalf, bhalf] = [ 'abcdefghijklm', 'nopqrstuvwxyz' ]

  var email = 'blake.holleman@gmail.';
  var keep = [];

  for (var i=pos; i < 10; i++) {
    for (var c2 of alpha) {
      for (var j=0; j < fhalf.length; j++) {
        inputCells.setValues([[email + alpha[i] + c2 + fhalf[j], email + alpha[i] + c2 + bhalf[j]]]);
        var [[e1, e2]] = isEmails.getValues();
        if (e1) {
          keep.push(alpha[i] + c2 + fhalf[j]);
        } else if (e2) {
          keep.push(alpha[i] + c2 + bhalf[j]);
        }
      }
    }
    break;
  }
  Logger.log(keep);
}

function dom3_3(pos=10) {
  var loggerSheet = SpreadsheetApp.getActiveSheet();
  var inputCells = loggerSheet.getRange(3,1,1,2);
  var isEmails = loggerSheet.getRange(3,3,1,2);
  var alpha = 'abcdefghijklmnopqrstuvwxyz';
  var [fhalf, bhalf] = [ 'abcdefghijklm', 'nopqrstuvwxyz' ]

  var email = 'blake.holleman@gmail.';
  var keep = [];

  for (var i=pos; i < 15; i++) {
    for (var c2 of alpha) {
      for (var j=0; j < fhalf.length; j++) {
        inputCells.setValues([[email + alpha[i] + c2 + fhalf[j], email + alpha[i] + c2 + bhalf[j]]]);
        var [[e1, e2]] = isEmails.getValues();
        if (e1) {
          keep.push(alpha[i] + c2 + fhalf[j]);
        } else if (e2) {
          keep.push(alpha[i] + c2 + bhalf[j]);
        }
      }
    }
    break;
  }
  Logger.log(keep);
}

function dom3_4(pos=15) {
  var loggerSheet = SpreadsheetApp.getActiveSheet();
  var inputCells = loggerSheet.getRange(4,1,1,2);
  var isEmails = loggerSheet.getRange(4,3,1,2);
  var alpha = 'abcdefghijklmnopqrstuvwxyz';
  var [fhalf, bhalf] = [ 'abcdefghijklm', 'nopqrstuvwxyz' ]

  var email = 'blake.holleman@gmail.';
  var keep = [];

  for (var i=pos; i < 20; i++) {
    for (var c2 of alpha) {
      for (var j=0; j < fhalf.length; j++) {
        inputCells.setValues([[email + alpha[i] + c2 + fhalf[j], email + alpha[i] + c2 + bhalf[j]]]);
        var [[e1, e2]] = isEmails.getValues();
        if (e1) {
          keep.push(alpha[i] + c2 + fhalf[j]);
        } else if (e2) {
          keep.push(alpha[i] + c2 + bhalf[j]);
        }
      }
    }
    break;
  }
  Logger.log(keep);
}

function dom3_5(pos=20) {
  var loggerSheet = SpreadsheetApp.getActiveSheet();
  var inputCells = loggerSheet.getRange(5,1,1,2);
  var isEmails = loggerSheet.getRange(5,3,1,2);
  var alpha = 'abcdefghijklmnopqrstuvwxyz';
  var [fhalf, bhalf] = [ 'abcdefghijklm', 'nopqrstuvwxyz' ]

  var email = 'blake.holleman@gmail.';
  var keep = [];

  for (var i=pos; i < 25; i++) {
    for (var c2 of alpha) {
      for (var j=0; j < fhalf.length; j++) {
        inputCells.setValues([[email + alpha[i] + c2 + fhalf[j], email + alpha[i] + c2 + bhalf[j]]]);
        var [[e1, e2]] = isEmails.getValues();
        if (e1) {
          keep.push(alpha[i] + c2 + fhalf[j]);
        } else if (e2) {
          keep.push(alpha[i] + c2 + bhalf[j]);
        }
      }
    }
    break;
  }
  Logger.log(keep);
}

function dom2() {
  var loggerSheet = SpreadsheetApp.getActiveSheet();
  var inputCells = loggerSheet.getRange(1,1,1,2);
  var isEmails = loggerSheet.getRange(1,3,1,2);
  var alpha = 'abcdefghijklmnopqrstuvwxyz';
  var [fhalf, bhalf] = [ 'abcdefghijklm', 'nopqrstuvwxyz' ]

  var email = 'blake.holleman@gmail.';
  var keep = [];

  for (var c1 of alpha) {
    for (var i=0; i < fhalf.length; i++) {
      inputCells.setValues([[email + c1 + fhalf[i], email + c1 + bhalf[i]]]);
      var [[e1, e2]] = isEmails.getValues();
      if (e1) {
        keep.push(c1 + fhalf[i]);
      } else if (e2) {
        keep.push( c1 + bhalf[i]);
      }
    }
  }
  Logger.log(keep);
}

function splitAlpha() {
  var alpha = 'abcdefghijklmnopqrstuvwxyz';
  var half = alpha.length/2;
  var fhalf = alpha.slice(0,half);
  var bhalf = alpha.slice(half);
  console.log([fhalf, bhalf]);
  return [fhalf, bhalf];
}

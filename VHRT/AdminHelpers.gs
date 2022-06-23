// Building Add-on html //
// _____________________________________________________________________________________________________________________________________ //

function createRadioButtons(elementId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var radios = {};
  var result = {};

  // Handle protected sheet selection //
  if (elementId == 'b-sel-sheet') {
    radios['sheets'] = ss.getSheets().map(s => s.getSheetName());
  
  // Get protected sheet properties //
  } else {
    var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet'); 
    var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
    var reportSheetId = sheetProperties['reportSheet'];

    // Handle continue with current active report //
    if (reportSheetId) {
      radios['updateActiveSheets'] = 'Updating active sheets';
      elementId = 'sel-update'

    // Handle create new active report //
    } else {
      var dateRangeString = getDateRange(); 
      var autoReportCols = [sheetProperties['primaryCase'], sheetProperties['primaryFirm'], 'Attorneys (First and Last)', 'Hours spent on case between ' + dateRangeString, 'Billing Rate (hr)'];
      var manReportCols = function() {
        var headers = getSheetById(sheetId).getDataRange().getValues()[0];
        headers.splice(headers.indexOf(sheetProperties['primaryKey']), 1);
        headers.splice(headers.indexOf(sheetProperties['primaryCase']), 1);
        headers.splice(headers.indexOf(sheetProperties['primaryFirm']), 1);
        headers.splice(headers.indexOf(sheetProperties['primaryProBono']), 1);
        return headers;
      }
      radios['autoCols'] = autoReportCols;
      radios['manCols'] = manReportCols();
    }
  }
  
  result[elementId] = radios;
  Logger.log(result);
  return result;
}

function getActiveSheetProperties(sheetName=null) {
  var activeSheetExists;

  // Get/Set protected sheet //
  if (!sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
    var activeSheetName = sheetId ? getSheetById(sheetId).getSheetName() : null;

    if (!activeSheetName) {
      return null;
    } else {
      var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
      activeSheetExists = true;
    }

  } else {
    var activeSheetName = sheetName;
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName);
    var sheetId = activeSheet.getSheetId().toString();
    var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
    activeSheetExists = false;

    setScriptProperty('protectedSheet', sheetId);
    setExceptionsProperty(event="setup", type="email");
  }

  // Return protected sheet properties to client //
  var entries = Object.entries(sheetProperties);

  var formatEntries = {
    'primaryKey':'<b>Primary keys:</b><br>',
    'primaryCase':'<b>Case names:</b><br>',
    'primaryFirm':'<b>Law firms:</b><br>',
    'primaryEmail':'<b>Attorney emails:</b><br>',
    'managerEmail':'<b>Office manager emails:</b><br>',
    'primaryProBono':'<b>Attorney names:</b><br>'
  };
  
  var formattedProperties = [];
  for (var entry of entries) {
    var propertyKey = entry[0];
    var property = entry[1];
    if (Object.keys(formatEntries).includes(propertyKey)) {
      formattedProperties.push(formatEntries[propertyKey] + '"' + entry[1] + '"');
    }
  }

  var activeSheetProperties = {};
  activeSheetProperties[activeSheetName] = formattedProperties;
  activeSheetProperties.activeSheetExists = activeSheetExists;

  //if (!activeSheetExists) {
  //  var now = Utilities.formatDate(new Date(), 'America/Chicago', 'MM-dd-yyyy HH:mm:ss');
  //  activeSheet.setName('Protected ' + now);
  //}
      
  return activeSheetProperties;
}


// Building admin sheets //
// _____________________________________________________________________________________________________________________________________ //

function formatSheet(sheetId, type) {
  // set sheet name and tab color //
  if (type == 'protected') {
    var protectedSheet = getSheetById(sheetId);
    var now = Utilities.formatDate(new Date(), 'America/Chicago', 'MM-dd-yyyy HH:mm:ss');  
    var protectedSheetName = 'Protected ' + now;
    protectedSheet.setName('*ACTIVE* ' + protectedSheetName);
    protectedSheet.setTabColor('blue');
  }

  if (type == 'report') {
    var reportSheet = getSheetById(sheetId);
    reportSheet.setName('*ACTIVE* ' + 'Report ' + Utilities.formatDate(new Date(), 'America/Chicago', 'MM-dd-yyyy HH:mm:ss'));
    reportSheet.setTabColor('blue');
  }
}

function getSheetById(sheetId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet;
  for (s of ss.getSheets()) {
    if (s.getSheetId() == sheetId) {
      sheet = s;
      break
    }
  }
  return sheet;
}

function insertActiveProtectedSheet(protectedSheetId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var protectedSheet = getSheetById(protectedSheetId);
  var protectedSheetName = protectedSheet.getSheetName();
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId));  
 
  //ss.setActiveSheet(protectedSheet);
  setPermissions(protectedSheet);
  setSheetDataValidations(protectedSheetId, ['primaryKey', 'primaryCase'], headers=true)  
  formatSheet(protectedSheetId, 'protected');
}

function insertActiveReport(protectedSheetId, reportHeaders) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId));
  var reorderFirst = ['primaryCase','primaryFirm'];
  var reportColumns = reorderCols(reorderFirst, reportHeaders);
  
  // Insert new report summary sheet //
  reportSheetId = ss.insertSheet('Report', 1).getSheetId();
  sheetProperties['reportSheet'] = reportSheetId;
  reportSheet = getSheetById(reportSheetId);
  var reportSheetName = reportSheet.getSheetName();

  // Set internal report columns for summary sheet //
  var internalReportCols = reportHeaders.slice();
  internalReportCols.push(sheetProperties['primaryKey']);
  internalReportCols.push('Timestamp');
  reportSheet.getRange(1, 1, 1, internalReportCols.length).setValues([reorderCols(reorderFirst, internalReportCols)]);
 
  setPermissions(reportSheet);
  setSheetDataValidations(reportSheetId, [], headers=true);
  formatSheet(reportSheetId, 'report');

  // Save script properties //
  setScriptProperty(protectedSheetId, sheetProperties);
  setScriptProperty('reportColumns', reportColumns);

  // Debug //
  isReportSheet = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId)).hasOwnProperty('reportSheet');
  isReportSheet ? Logger.log('Saved report sheet: ' + reportSheetId) : Logger.log('Saved report sheet: ' + reportSheetId);
  isReportColumns = PropertiesService.getScriptProperties().getProperty('reportColumns');
  isReportColumns ? Logger.log('Saved user-facing report columns: ' + JSON.stringify(isReportColumns)) : Logger.log('Saved user-facing report columns: ' + Boolean(isReportColumns));
  return reportSheetId;
}

function reorderCols(colsFirst, colArr) {
  var sheetId = PropertiesService.getScriptProperties().getProperty('protectedSheet');
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
  var count = 0;
  for (var col of colsFirst) {
    val = sheetProperties[col];
    if (val && colArr.includes(val)) {
      var colIdx = colArr.indexOf(val);
      colArr.splice(colIdx, 1);
      colArr.splice(count, 0, val);
      count += 1;
    } else {
      Logger.log(col + ' is either not a set script property or user-selected reporting column');
    }
  }
  return colArr;
}

function resetActiveProtectedSheet(protectedSheetId) {
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId));
  var protectedSheet = getSheetById(protectedSheetId);
  var protectedSheetName = protectedSheet.getSheetName();
  
  protectedSheet.getDataRange().clearDataValidations();
  protectedSheet.setName(protectedSheetName.split('*ACTIVE* ')[1]);
  protectedSheet.setTabColor(null);
  delete sheetProperties['reportSheet'];
  
  setScriptProperty('protectedSheet', '');
  setScriptProperty(protectedSheetId, sheetProperties);
}

function resetActiveReport(protectedSheetId) {
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(protectedSheetId));
  var reportSheetId = sheetProperties['reportSheet'];
  var reportSheet = getSheetById(reportSheetId);
  
  if (reportSheet) {
    var reportRenamed = 'Report Closed ~ ' + Utilities.formatDate(new Date(), 'America/Chicago', 'MM-dd-yyyy HH:mm:ss');
    reportSheet.setName(reportRenamed);
    reportSheet.setTabColor(null);
  }
  delete sheetProperties['reportSheet'];

  setScriptProperty(protectedSheetId, sheetProperties);
  setScriptProperty('reportColumns', []);
}

function setScriptProperty(propertyName, data) {
  var string_data = data;
  if (typeof data == 'object') {
    string_data = JSON.stringify(data);    
  }
  PropertiesService.getScriptProperties().setProperty(propertyName, string_data);
}


// Permissions & Validations //
// _____________________________________________________________________________________________________________________________________ //

function checkEmails(sheetProperty='protectedSheet') {
  var sheetId = PropertiesService.getScriptProperties().getProperty(sheetProperty);
  if (!sheetId) {
    raiseAlert('Error', 'You must activate a protected sheet before validating emails');
    return null;
  }

  // Get script properties //
  var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId))
  var primaryEmail = sheetProperties['primaryEmail'];
  var managerEmail = sheetProperties['managerEmail'];
  var exceptions = JSON.parse(PropertiesService.getScriptProperties().getProperty('exceptions'));
  
  var sheet = getSheetById(sheetId);
  for (var colKey of ['primaryEmail', 'managerEmail']) {
    var [keyIdx, keyRange, keyEmails] = getColumnCustom(sheet, colKey);
    var a1List = getCommaSepRange(keyRange.getA1Notation());  
    var noErrors = [];
    var displayErrors = {};
    var headers = sheet.getDataRange().getValues()[0];
    var emailRe = createEmailRegex();
    var email;
    var isEmail;

    // Categorize errors for display //
    for (var a1 of a1List) {
      email = sheet.getRange(a1).getValue();
      isEmail = emailRe.test(email);
      if (isEmail) {
        noErrors.push(a1);
        delete exceptions['email'][a1];
      } else if (exceptions['email'][a1] == email) {
        noErrors.push(a1);
      } else {
        email ? displayErrors[a1] = email : displayErrors[a1] = '{NULL}';
      }
    }

    // Set cell formatting //
    if (noErrors.length) {
      sheet.getRangeList(noErrors).setBackground('#FFFFFF');
      setScriptProperty('exceptions', exceptions);
    }
    if (Object.keys(displayErrors).length) {
      sheet.getRangeList(Object.keys(displayErrors)).setBackground('#FFB6C1');
      
      // Raise alert for errors and save exceptions on user direction //
      checkEmailHandleError(sheetId, displayErrors);
    }
  }
}

function checkEmailHandleError(sheetId, displayErrorsObj) {
  var sheet = getSheetById(sheetId);
  var exceptions = JSON.parse(PropertiesService.getScriptProperties().getProperty('exceptions'));
  var errorKeys = Object.keys(displayErrorsObj);
  var errorVals = Object.values(displayErrorsObj);

  if (errorKeys) {
    var errors = errorVals.join(', ');
    var a1Cells = errorKeys.join(', ');
    var ui = SpreadsheetApp.getUi();
    var response = raiseAlert(
      'Possible invalid emails highlighted in cells ' + a1Cells + ':',
      errors + '\r\n' + '\r\n' +
      'Click "YES" to save these cell values as exceptions and ignore this alert in the future.',
      buttons='YES_NO'
    );
    
    if (response == ui.Button.YES) {
      Logger.log(displayErrorsObj);
      Object.entries(displayErrorsObj).forEach(e => {
        exceptions['email'][e[0]] = e[1];
        sheet.getRange(e[0]).setBackground('#FFFFFF');
      });
      setScriptProperty('exceptions', exceptions);
    }
  }
}

function createEmailRegex() {
  // validDoms scraped from Google-approved valid email domains - checked against Sheets isEmail() cell formula //
  var emailRe = '^[a-z0-9_\\-+.]+@[a-z0-9_\\-.]+';
  var validDoms2Char = 
    `ad, ae, af, ag, ai, al, am, ao, aq, ar, as, at, au, aw, ax, az, ba, bb, bd, be, bf, bg, bh, bi, bj, bl, bm, bn, bo, bq, br, bs, 
    bt, bv, bw, by, bz, ca, cc, cd, cf, cg, ch, ci, ck, cl, cm, cn, co, cr, cu, cv, cw, cx, cy, cz, de, dj, dk, dm, do, dz, ec, ee, eg, 
    eh, er, es, et, fi, fj, fk, fm, fo, fr, ga, gb, gd, ge, gf, gg, gh, gi, gl, gm, gn, gp, gq, gr, gs, gt, gu, gw, gy, hk, hm, hn, hr, 
    ht, hu, id, ie, il, im, in, io, iq, ir, is, it, je, jm, jo, jp, ke, kg, kh, ki, km, kn, kp, kr, kw, ky, kz, la, lb, lc, li, lk, lr, 
    ls, lt, lu, lv, ly, ma, mc, md, me, mf, mg, mh, mk, ml, mm, mn, mo, mp, mq, mr, ms, mt, mu, mv, mw, mx, my, mz, na, nc, ne, nf, ng, 
    ni, nl, no, np, nr, nu, nz, om, pa, pe, pf, pg, ph, pk, pl, pm, pn, pr, ps, pt, pw, py, qa, re, ro, rs, ru, rw, sa, sb, sc, sd, se, 
    sg, sh, si, sj, sk, sl, sm, sn, so, sr, ss, st, sv, sx, sy, sz, tc, td, tf, tg, th, tj, tk, tl, tm, tn, to, tr, tt, tv, tw, tz, ua, 
    ug, uk, um, us, uy, uz, va, vc, ve, vg, vi, vn, vu, wf, ws, ye, yt, za, zm, zw`
  var validDoms3Char = 'gov, com, net, edu, org';
  var validDoms = '(' + validDoms2Char + ', ' + validDoms3Char + ')';

  validDomsRe = validDoms.split(', ').join('|');
  emailRe = new RegExp(emailRe + validDomsRe + '$', 'i');
  return emailRe;
}

function removeDataValidations(sheetId) {
  var sheet = getSheetById(sheetId);
  if (sheet) {
    var sheetProperties = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId));
    if (sheetProperties) {
      sheet.getDataRange().clearDataValidations();
      var reportSheetId = sheetProperties['reportSheet'];
      var reportSheet = getSheetById(reportSheetId);
      if (reportSheet) {
        reportSheet.getDataRange().clearDataValidations();
      }
    }
  }  

  Logger.log('Removed data validations from active sheets');
}

function setPermissions(activeSheet) {
   var protection = activeSheet.protect().setDescription('Admin Use Only');
   var admin = PropertiesService.getScriptProperties().getProperty('adminUser');
   //var superUser = PropertiesService.getScriptProperties().getProperty('superUser');

   protection.addEditors([admin]);
   Logger.log(activeSheet.getSheetName() + ' protected for admin use only');   
}

function setDataValidation(activeSheetId, range) {
  var activeSheet = getSheetById(activeSheetId);
  var a1range = range.getA1Notation();
  var a1List = getCommaSepRange(a1range);
  var validations = [];
    
  for (var i=0; i < a1List.length; i++) {
    var cellText = activeSheet.getRange(a1List[i]).getDisplayValue();
    var ruleContainsText = SpreadsheetApp.newDataValidation().requireTextEqualTo(cellText).setAllowInvalid(false).build();
    validations.push([ruleContainsText]);
  }

  try {
    // 2D array for column ranges //
    activeSheet.getRange(a1range).setDataValidations(validations);
  } catch (e) {
    // 1D array for row ranges //
    activeSheet.getRange(a1range).setDataValidations([validations.flat()]);
  }
  
  Logger.log('Set active sheet ' + activeSheet.getSheetName() + ' validations: ' + range.getA1Notation());
}

function setSheetDataValidations(sheetId, colKeysList, headers=true) {
  sheet = getSheetById(sheetId);
  if (headers) {
    var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    setDataValidation(sheetId, headersRange);
  }

  for (var colKey of colKeysList) {
    var [keyIdx, keyRange, keyValues] = getColumnCustom(sheet, colKey);
    setDataValidation(sheetId, keyRange);
  }
}


// Toast helpers //
// _____________________________________________________________________________________________________________________________________ //

function logToDebugger(object) {
  Logger.log(object);
}

function raiseAlert(alertTitle, alertString, buttons='OK') {
  var ui = SpreadsheetApp.getUi();
  var enums = {
    'OK':ui.ButtonSet.OK, 
    'OK_CANCEL':ui.ButtonSet.OK_CANCEL,
    'YES_NO':ui.ButtonSet.YES_NO,
    'YES_NO_CANCEL':ui.ButtonSet.YES_NO_CANCEL
  }
  
  var response = ui.alert(alertTitle, alertString, enums[buttons]);
  return response;
}


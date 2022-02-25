// onEdit trigger helper //

function isInvalidCell(sheetId=PropertiesService.getScriptProperties().getProperty('lastUploadedSheet'), a1range, event) {
  var adminLoggerUrl = PropertiesService.getScriptProperties().getProperty('adminLoggerUrl');
  var exceptions = JSON.parse(PropertiesService.getScriptProperties().getProperty('exceptions'));
  var sheet = getSheetById(sheetId);
  var a1List = getCommaSepRange(a1range);

  // check for and highlight invalid emails //

  var noErrors = [];
  var displayErrors = {};
  var headers = sheet.getDataRange().getValues()[0];
  var emailRe = createEmailRegex();
  var email;

  for (var a1 of a1List) {
    email = sheet.getRange(a1).getValue();
    
    // handles edit events //
    if (event == 'edit') {
      try {
        var primaryEmail = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId))['primaryEmail'];
        var managerEmail = JSON.parse(PropertiesService.getScriptProperties().getProperty(sheetId))['managerEmail'];
      } catch (e) {
        // catch handles edits to sheet during upload process //
        var primaryEmail = PropertiesService.getScriptProperties().getProperty('primaryEmail');
        var managerEmail = PropertiesService.getScriptProperties().getProperty('managerEmail');    
      }

      email = (sheet.getRange(a1).getColumn() == headers.indexOf(primaryEmail) + 1) || 
              (sheet.getRange(a1).getColumn() == headers.indexOf(managerEmail) + 1) ? 
              sheet.getRange(a1).getValue() : 'NA';
    }

    // handles upload events - assumes data type email //
    if (email == 'NA') {
      continue;

    } else if (email && exceptions['email'].includes(email)) {
      continue;

    } else {

      /* DEPRECATED
      var loggerSheet = SpreadsheetApp.openByUrl(adminLoggerUrl).getSheets()[0];
      var inputCell = loggerSheet.getRange(1, 1);
      var isEmail = loggerSheet.getRange(1,2); 
      inputCell.setValue(email);
      
      if(!isEmail.getValue()) {

        // TODO clean this up //
        var lastRow = loggerSheet.getLastRow();
        var insertRange = loggerSheet.getRange(lastRow + 1, 1, 1, 4);
        var now = Utilities.formatDate(new Date(), 'America/Chicago', 'yyyy-MM-dd HH:mm:ss');
        loggerSheet.getRange(insertRange.getA1Notation()).setValues([[now, email, isEmail.getDisplayValue(), event]]);
      */

      var isEmail = emailRe.test(email);

      if (!isEmail) {
        email ? displayErrors[a1] = email : displayErrors[a1] = '{NULL}';
      } else {
        noErrors.push(a1);
      }
    }
  }
  
  // (un)highlight errors //
  if (Object.keys(displayErrors).length) {
    sheet.getRangeList(Object.keys(displayErrors)).setBackground('#FFB6C1');
  }

  if (noErrors.length) {
    sheet.getRangeList(noErrors).setBackground('#FFFFFF');
  }

  // raise alert for errors and save exceptions on user direction //

  if (Object.keys(displayErrors).length) {
    var errors = Object.values(displayErrors).join(', ');
    var cells = Object.keys(displayErrors).join(', ');
    var ui = SpreadsheetApp.getUi();
    var response = raiseAlert(
      'Possible invalid emails highlighted in cells ' + cells + ':',
      errors + '\r\n' + '\r\n' +
      'Click "OK" to return to the sheet or "CANCEL" to ignore this alert for these emails in the future.',
      buttons='OK_CANCEL'
    );
    
    if (response == ui.Button.CANCEL) {
      Object.entries(displayErrors).forEach(e => {
        if (!(e[1] == '{NULL}')) {
          exceptions['email'].push(e[1]);
          sheet.getRange(e[0]).setBackground('#FFFFFF');
        }
      })
      PropertiesService.getScriptProperties().setProperty('exceptions', JSON.stringify(exceptions));
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




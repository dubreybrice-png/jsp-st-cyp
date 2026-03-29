/* ═══════════════════════════════════════════════════════
   POINT D'ENTRÉE — JSP St Cyprien
   ═══════════════════════════════════════════════════════ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('JSP St Cyprien — Suivi Activité')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ═══════════════════════════════════════════════════════
   AUTHENTIFICATION
   ═══════════════════════════════════════════════════════ */

function loginJSP(login, pwd) {
  var ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(Config.SHEETS.JSP);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(login).trim() &&
        String(data[i][2]).trim() === String(pwd).trim()) {
      return {
        success: true,
        identite: data[i][0],
        login: String(data[i][1]),
        section: String(data[i][3])
      };
    }
  }
  return { success: false, message: 'Identifiant ou mot de passe incorrect.' };
}

function loginReferent(login, pwd) {
  var ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(Config.SHEETS.REFERENTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(login).trim() &&
        String(data[i][2]).trim() === String(pwd).trim()) {
      return {
        success: true,
        identite: data[i][0],
        login: String(data[i][1])
      };
    }
  }
  return { success: false, message: 'Identifiant ou mot de passe incorrect.' };
}

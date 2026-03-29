/* ═══════════════════════════════════════════════════════
   SETUP SPREADSHEET — Exécuter une seule fois
   Crée le Google Sheet avec tous les onglets et données
   ═══════════════════════════════════════════════════════ */

function setupSpreadsheet() {
  var ss = SpreadsheetApp.create('JSP St Cyp — Suivi Activité');
  var defaultSheets = ss.getSheets();

  /* ────── RÉFÉRENTS ────── */
  var refsSheet = ss.insertSheet('Référents');
  refsSheet.appendRow(['Identité', 'Login', 'Mot de passe']);
  var referents = [
    ['Jérome Casenove',  '66000',    '66000'],
    ['Sophie Martin',    'smartin',  '66000'],
    ['Philippe Dubois',  'pdubois',  '66000'],
    ['Marie Laurent',    'mlaurent', '66000'],
    ['François Garcia',  'fgarcia',  '66000'],
    ['Isabelle Roux',    'iroux',    '66000'],
    ['Pierre Bonnet',    'pbonnet',  '66000'],
    ['Catherine Blanc',  'cblanc',   '66000'],
    ['Alain Faure',      'afaure',   '66000'],
    ['Nathalie Girard',  'ngirard',  '66000']
  ];
  for (var i = 0; i < referents.length; i++) refsSheet.appendRow(referents[i]);
  refsSheet.setFrozenRows(1);
  refsSheet.getRange('1:1').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');

  /* ────── SECTIONS ────── */
  var secSheet = ss.insertSheet('Sections');
  secSheet.appendRow(['Section', 'Référent(s)']);
  var sections = [
    ['JSP1', 'Jérome Casenove, Sophie Martin'],
    ['JSP2', 'Philippe Dubois, Marie Laurent'],
    ['JSP3', 'François Garcia, Isabelle Roux'],
    ['JSP4', 'Pierre Bonnet, Catherine Blanc']
  ];
  for (var i = 0; i < sections.length; i++) secSheet.appendRow(sections[i]);
  secSheet.setFrozenRows(1);
  secSheet.getRange('1:1').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');

  /* ────── LISTE JSP ────── */
  var jspSheet = ss.insertSheet('Liste JSP');
  jspSheet.appendRow(['Identité', 'Login', 'Mot de passe', 'Section']);
  var jsps = [
    ['Lucas Martin',  'lmartin',  '66000', 'JSP1'],
    ['Emma Dupont',   'edupont',  '66000', 'JSP1'],
    ['Hugo Bernard',  'hbernard', '66000', 'JSP1'],
    ['Léa Moreau',    'lmoreau',  '66000', 'JSP2'],
    ['Nathan Petit',  'npetit',   '66000', 'JSP2'],
    ['Chloé Robert',  'crobert',  '66000', 'JSP2'],
    ['Théo Richard',  'trichard', '66000', 'JSP3'],
    ['Jade Durand',   'jdurand',  '66000', 'JSP3'],
    ['Enzo Laurent',  'elaurent', '66000', 'JSP4'],
    ['Manon Simon',   'msimon',   '66000', 'JSP4']
  ];
  for (var i = 0; i < jsps.length; i++) jspSheet.appendRow(jsps[i]);
  jspSheet.setFrozenRows(1);
  jspSheet.getRange('1:1').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');

  // Data validation: colonne Section = dropdown depuis onglet Sections
  var sectionRange = secSheet.getRange('A2:A' + (sections.length + 1));
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sectionRange)
    .setAllowInvalid(false)
    .build();
  jspSheet.getRange('D2:D100').setDataValidation(rule);

  /* ────── ÉVÉNEMENTS ────── */
  var evtSheet = ss.insertSheet('Événements');
  evtSheet.appendRow(['ID', 'Date', 'Nom', 'Heure Début', 'Heure Fin', 'Lieu', 'Sections', 'Créé par']);
  evtSheet.setFrozenRows(1);
  evtSheet.getRange('1:1').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');

  /* ────── PRÉSENCES ────── */
  var presSheet = ss.insertSheet('Présences');
  presSheet.appendRow(['EventID', 'Date', 'Login JSP', 'Nom JSP', 'Section', 'Présent', 'Absence Signalée', 'Motif']);
  presSheet.setFrozenRows(1);
  presSheet.getRange('1:1').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');

  /* ────── Nettoyage ────── */
  for (var i = 0; i < defaultSheets.length; i++) {
    try { ss.deleteSheet(defaultSheets[i]); } catch(e) {}
  }

  // Auto-resize
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    try {
      for (var c = 1; c <= allSheets[i].getLastColumn(); c++) {
        allSheets[i].autoResizeColumn(c);
      }
    } catch(e) {}
  }

  Logger.log('╔═══════════════════════════════════════════════════╗');
  Logger.log('║  Spreadsheet créé avec succès !                   ║');
  Logger.log('║  ID : ' + ss.getId());
  Logger.log('║  URL: ' + ss.getUrl());
  Logger.log('║                                                   ║');
  Logger.log('║  → Copiez l\'ID dans Config.js > SPREADSHEET_ID   ║');
  Logger.log('╚═══════════════════════════════════════════════════╝');

  return ss.getId();
}

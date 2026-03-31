/* ═══════════════════════════════════════════════════════
   SETUP SPREADSHEET — Exécuter une seule fois
   Crée le Google Sheet avec tous les onglets et données
   ═══════════════════════════════════════════════════════ */

function setupSpreadsheet() {
  var ss = SpreadsheetApp.create('JSP St Cyp — Suivi Activité');
  var defaultSheets = ss.getSheets();

  /* ────── RÉFÉRENTS ────── */
  var refsSheet = ss.insertSheet('Référents');
  refsSheet.appendRow(['Identité', 'Login', 'Mot de passe', 'Email']);
  var referents = [
    ['Jérome Casenove',  '66000',    '66000', 'j.casenove@sdis66.fr'],
    ['Sophie Martin',    'smartin',  '66000', 's.martin@sdis66.fr'],
    ['Philippe Dubois',  'pdubois',  '66000', 'p.dubois@sdis66.fr'],
    ['Marie Laurent',    'mlaurent', '66000', 'm.laurent@sdis66.fr'],
    ['François Garcia',  'fgarcia',  '66000', 'f.garcia@sdis66.fr'],
    ['Isabelle Roux',    'iroux',    '66000', 'i.roux@sdis66.fr'],
    ['Pierre Bonnet',    'pbonnet',  '66000', 'p.bonnet@sdis66.fr'],
    ['Catherine Blanc',  'cblanc',   '66000', 'c.blanc@sdis66.fr'],
    ['Alain Faure',      'afaure',   '66000', 'a.faure@sdis66.fr'],
    ['Nathalie Girard',  'ngirard',  '66000', 'n.girard@sdis66.fr']
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
  jspSheet.appendRow(['Identité', 'Login', 'Mot de passe', 'Section', 'Email']);
  var jsps = [
    ['Lucas Martin',  'lmartin',  '66000', 'JSP1', 'l.martin@example.com'],
    ['Emma Dupont',   'edupont',  '66000', 'JSP1', 'e.dupont@example.com'],
    ['Hugo Bernard',  'hbernard', '66000', 'JSP1', 'h.bernard@example.com'],
    ['Léa Moreau',    'lmoreau',  '66000', 'JSP2', 'l.moreau@example.com'],
    ['Nathan Petit',  'npetit',   '66000', 'JSP2', 'n.petit@example.com'],
    ['Chloé Robert',  'crobert',  '66000', 'JSP2', 'c.robert@example.com'],
    ['Théo Richard',  'trichard', '66000', 'JSP3', 't.richard@example.com'],
    ['Jade Durand',   'jdurand',  '66000', 'JSP3', 'j.durand@example.com'],
    ['Enzo Laurent',  'elaurent', '66000', 'JSP4', 'e.laurent@example.com'],
    ['Manon Simon',   'msimon',   '66000', 'JSP4', 'm.simon@example.com']
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

/* ═══════════════════════════════════════════════════════
   MIGRATION — Ajouter les colonnes Email
   Exécuter si le spreadsheet existait avant la v2
   ═══════════════════════════════════════════════════════ */
function addEmailColumns() {
  var ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);

  var jspSheet = ss.getSheetByName(Config.SHEETS.JSP);
  var jspHeaders = jspSheet.getRange(1, 1, 1, jspSheet.getLastColumn()).getValues()[0];
  if (jspHeaders.indexOf('Email') === -1) {
    var col = jspSheet.getLastColumn() + 1;
    jspSheet.getRange(1, col).setValue('Email').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    Logger.log('Colonne Email ajoutée à Liste JSP (colonne ' + col + ')');
  } else {
    Logger.log('Colonne Email déjà présente dans Liste JSP');
  }

  var refSheet = ss.getSheetByName(Config.SHEETS.REFERENTS);
  var refHeaders = refSheet.getRange(1, 1, 1, refSheet.getLastColumn()).getValues()[0];
  if (refHeaders.indexOf('Email') === -1) {
    var col = refSheet.getLastColumn() + 1;
    refSheet.getRange(1, col).setValue('Email').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    Logger.log('Colonne Email ajoutée à Référents (colonne ' + col + ')');
  } else {
    Logger.log('Colonne Email déjà présente dans Référents');
  }

  Logger.log('\n✅ Migration terminée ! Remplissez les emails dans le spreadsheet.');
}

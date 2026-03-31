/* ═══════════════════════════════════════════════════════
   SERVICE DE DONNÉES — JSP St Cyprien
   Toutes les opérations CRUD spreadsheet
   ═══════════════════════════════════════════════════════ */

function getSS_() {
  return SpreadsheetApp.openById(Config.SPREADSHEET_ID);
}

function formatDate_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, 'Europe/Paris', 'dd/MM/yyyy');
  var s = String(d);
  /* Convertir yyyy-MM-dd en dd/MM/yyyy si besoin */
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return m[3] + '/' + m[2] + '/' + m[1];
  return s;
}

function formatTime_(t) {
  if (t instanceof Date) return Utilities.formatDate(t, 'Europe/Paris', 'HH:mm');
  var s = String(t || '');
  /* Si déjà au format HH:mm ou HH:mm:ss, ne garder que HH:mm */
  var m = s.match(/(\d{1,2}:\d{2})/);
  if (m) return m[1];
  return s;
}

/* ═══════════════════════════════════════════════════════
   CHARGEMENT PAGE JSP
   ═══════════════════════════════════════════════════════ */
function getJSPPageData(jspLogin) {
  var ss = getSS_();

  /* Info JSP */
  var jspSheet = ss.getSheetByName(Config.SHEETS.JSP);
  var jspData = jspSheet.getDataRange().getValues();
  var jsp = null;
  for (var i = 1; i < jspData.length; i++) {
    if (String(jspData[i][1]).trim() === String(jspLogin).trim()) {
      jsp = { identite: jspData[i][0], login: String(jspData[i][1]), section: String(jspData[i][3]) };
      break;
    }
  }
  if (!jsp) return { error: 'JSP non trouvé' };

  /* Événements de la section */
  var evtSheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var evtData = evtSheet.getDataRange().getValues();
  var events = [];
  for (var i = 1; i < evtData.length; i++) {
    if (!evtData[i][0]) continue;
    var sections = String(evtData[i][6]).split(',').map(function(s) { return s.trim(); });
    if (sections.indexOf(jsp.section) !== -1) {
      events.push({
        id: String(evtData[i][0]),
        date: formatDate_(evtData[i][1]),
        nom: String(evtData[i][2]),
        heureDebut: formatTime_(evtData[i][3]),
        heureFin: formatTime_(evtData[i][4]),
        lieu: String(evtData[i][5]),
        sections: sections
      });
    }
  }

  /* Présences de ce JSP */
  var presSheet = ss.getSheetByName(Config.SHEETS.PRESENCES);
  var presData = presSheet.getDataRange().getValues();
  var attendance = {};
  for (var i = 1; i < presData.length; i++) {
    if (String(presData[i][2]).trim() === jspLogin.trim()) {
      attendance[String(presData[i][0])] = {
        present: String(presData[i][5]),
        signale: String(presData[i][6]),
        motif: presData[i][7] || ''
      };
    }
  }

  /* Stats */
  var presences = 0, totalAbsences = 0, absSignalees = 0, absNonSignalees = 0;
  for (var i = 0; i < events.length; i++) {
    var att = attendance[events[i].id];
    if (att) {
      if (att.present === 'OUI') presences++;
      else {
        totalAbsences++;
        if (att.signale === 'OUI') absSignalees++;
        else absNonSignalees++;
      }
    }
  }
  var eventsDone = presences + totalAbsences;
  var tauxAbsence = eventsDone > 0 ? Math.round(totalAbsences / eventsDone * 1000) / 10 : 0;

  return {
    jsp: jsp,
    events: events,
    attendance: attendance,
    stats: {
      totalEvents: events.length,
      eventsDone: eventsDone,
      presences: presences,
      totalAbsences: totalAbsences,
      absSignalees: absSignalees,
      absNonSignalees: absNonSignalees,
      tauxAbsence: tauxAbsence
    }
  };
}

/* ═══════════════════════════════════════════════════════
   CHARGEMENT PAGE RESPONSABLE
   ═══════════════════════════════════════════════════════ */
function getReferentPageData(refLogin) {
  var ss = getSS_();

  /* Toutes les sections */
  var secSheet = ss.getSheetByName(Config.SHEETS.SECTIONS);
  var secData = secSheet.getDataRange().getValues();
  var allSections = [];
  for (var i = 1; i < secData.length; i++) {
    if (secData[i][0]) allSections.push(String(secData[i][0]));
  }

  /* Tous les événements */
  var evtSheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var evtData = evtSheet.getDataRange().getValues();
  var events = [];
  for (var i = 1; i < evtData.length; i++) {
    if (!evtData[i][0]) continue;
    events.push({
      id: String(evtData[i][0]),
      date: formatDate_(evtData[i][1]),
      nom: String(evtData[i][2]),
      heureDebut: formatTime_(evtData[i][3]),
      heureFin: formatTime_(evtData[i][4]),
      lieu: String(evtData[i][5]),
      sections: String(evtData[i][6]).split(',').map(function(s) { return s.trim(); }),
      creePar: String(evtData[i][7] || '')
    });
  }

  return { sections: allSections, events: events };
}

/* ═══════════════════════════════════════════════════════
   GESTION ÉVÉNEMENTS
   ═══════════════════════════════════════════════════════ */
function createEvent(eventData) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var id = new Date().getTime().toString();
  sheet.appendRow([
    id,
    eventData.date,
    eventData.nom,
    eventData.heureDebut,
    eventData.heureFin,
    eventData.lieu,
    eventData.sections.join(', '),
    eventData.creePar || ''
  ]);

  /* Notification par mail si demandé */
  if (eventData.notifyByMail) {
    try {
      sendEventNotification_(id, eventData);
    } catch(e) {
      Logger.log('Erreur envoi notification mail: ' + e.message);
    }
  }

  return { success: true, id: id };
}

function deleteEvent(eventId) {
  var ss = getSS_();

  /* Supprimer événement */
  var sheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(eventId)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  /* Supprimer les présences associées */
  var presSheet = ss.getSheetByName(Config.SHEETS.PRESENCES);
  var presData = presSheet.getDataRange().getValues();
  for (var i = presData.length - 1; i >= 1; i--) {
    if (String(presData[i][0]) === String(eventId)) {
      presSheet.deleteRow(i + 1);
    }
  }
  return { success: true };
}

/* ═══════════════════════════════════════════════════════
   GESTION PRÉSENCES
   ═══════════════════════════════════════════════════════ */
function getAttendanceList(eventId) {
  var ss = getSS_();

  /* Événement */
  var evtSheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var evtData = evtSheet.getDataRange().getValues();
  var event = null;
  for (var i = 1; i < evtData.length; i++) {
    if (String(evtData[i][0]) === String(eventId)) {
      event = {
        id: String(evtData[i][0]),
        date: formatDate_(evtData[i][1]),
        nom: String(evtData[i][2]),
        heureDebut: formatTime_(evtData[i][3]),
        heureFin: formatTime_(evtData[i][4]),
        lieu: String(evtData[i][5]),
        sections: String(evtData[i][6]).split(',').map(function(s) { return s.trim(); })
      };
      break;
    }
  }
  if (!event) return { error: 'Événement non trouvé' };

  /* JSP des sections concernées */
  var jspSheet = ss.getSheetByName(Config.SHEETS.JSP);
  var jspData = jspSheet.getDataRange().getValues();
  var jsps = [];
  for (var i = 1; i < jspData.length; i++) {
    if (event.sections.indexOf(String(jspData[i][3])) !== -1) {
      jsps.push({
        identite: String(jspData[i][0]),
        login: String(jspData[i][1]),
        section: String(jspData[i][3])
      });
    }
  }

  /* Présences existantes */
  var presSheet = ss.getSheetByName(Config.SHEETS.PRESENCES);
  var presData = presSheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < presData.length; i++) {
    if (String(presData[i][0]) === String(eventId)) {
      existing[String(presData[i][2]).trim()] = {
        present: String(presData[i][5]),
        signale: String(presData[i][6]),
        motif: String(presData[i][7] || '')
      };
    }
  }

  /* Fusionner */
  for (var i = 0; i < jsps.length; i++) {
    var ex = existing[jsps[i].login];
    if (ex) {
      jsps[i].present = ex.present;
      jsps[i].signale = ex.signale;
      jsps[i].motif = ex.motif;
    } else {
      jsps[i].present = 'OUI';
      jsps[i].signale = '';
      jsps[i].motif = '';
    }
  }

  jsps.sort(function(a, b) { return a.identite.localeCompare(b.identite); });

  /* Toutes les sections disponibles (pour le sélecteur dans l'appel) */
  var secSheet = ss.getSheetByName(Config.SHEETS.SECTIONS);
  var secData = secSheet.getDataRange().getValues();
  var allSections = [];
  for (var i = 1; i < secData.length; i++) {
    if (secData[i][0]) allSections.push(String(secData[i][0]));
  }

  return { event: event, jsps: jsps, allSections: allSections };
}

function saveAttendance(eventId, attendanceList) {
  var ss = getSS_();
  var presSheet = ss.getSheetByName(Config.SHEETS.PRESENCES);
  var presData = presSheet.getDataRange().getValues();

  /* Date de l'événement */
  var evtSheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var evtData = evtSheet.getDataRange().getValues();
  var eventDate = '';
  for (var i = 1; i < evtData.length; i++) {
    if (String(evtData[i][0]) === String(eventId)) {
      eventDate = formatDate_(evtData[i][1]);
      break;
    }
  }

  /* Map des lignes existantes */
  var existingRows = {};
  for (var i = 1; i < presData.length; i++) {
    if (String(presData[i][0]) === String(eventId)) {
      existingRows[String(presData[i][2]).trim()] = {
        row: i + 1,
        signale: String(presData[i][6]),
        motif: String(presData[i][7] || '')
      };
    }
  }

  for (var i = 0; i < attendanceList.length; i++) {
    var a = attendanceList[i];
    var ex = existingRows[a.login];

    if (ex) {
      /* Mise à jour ligne existante */
      presSheet.getRange(ex.row, 6).setValue(a.present);
      if (a.present === 'NON' && ex.signale === 'OUI') {
        /* Garder le flag signalé et motif */
      } else if (a.present === 'NON') {
        presSheet.getRange(ex.row, 7).setValue('NON');
      } else {
        presSheet.getRange(ex.row, 7).setValue('');
        presSheet.getRange(ex.row, 8).setValue('');
      }
    } else {
      /* Nouvelle ligne */
      var signale = a.present === 'NON' ? 'NON' : '';
      presSheet.appendRow([
        eventId, eventDate, a.login, a.identite, a.section,
        a.present, signale, ''
      ]);
    }
  }
  return { success: true };
}

function signalAbsence(jspLogin, eventId, motif) {
  var ss = getSS_();
  var presSheet = ss.getSheetByName(Config.SHEETS.PRESENCES);
  var presData = presSheet.getDataRange().getValues();

  /* Chercher un enregistrement existant */
  for (var i = 1; i < presData.length; i++) {
    if (String(presData[i][0]) === String(eventId) &&
        String(presData[i][2]).trim() === jspLogin.trim()) {
      presSheet.getRange(i + 1, 6).setValue('NON');
      presSheet.getRange(i + 1, 7).setValue('OUI');
      presSheet.getRange(i + 1, 8).setValue(motif);
      return { success: true };
    }
  }

  /* Créer un nouvel enregistrement */
  var jspSheet = ss.getSheetByName(Config.SHEETS.JSP);
  var jspData = jspSheet.getDataRange().getValues();
  var jspInfo = null;
  for (var i = 1; i < jspData.length; i++) {
    if (String(jspData[i][1]).trim() === jspLogin.trim()) {
      jspInfo = { identite: jspData[i][0], section: String(jspData[i][3]) };
      break;
    }
  }

  var evtSheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var evtData = evtSheet.getDataRange().getValues();
  var eventDate = '';
  for (var i = 1; i < evtData.length; i++) {
    if (String(evtData[i][0]) === String(eventId)) {
      eventDate = formatDate_(evtData[i][1]);
      break;
    }
  }

  presSheet.appendRow([
    eventId, eventDate, jspLogin,
    jspInfo ? jspInfo.identite : '',
    jspInfo ? jspInfo.section : '',
    'NON', 'OUI', motif
  ]);
  return { success: true };
}

/* ═══════════════════════════════════════════════════════
   BILAN GLOBAL
   ═══════════════════════════════════════════════════════ */
function getBilanData() {
  var ss = getSS_();

  /* Sections */
  var secSheet = ss.getSheetByName(Config.SHEETS.SECTIONS);
  var secData = secSheet.getDataRange().getValues();
  var sections = [];
  for (var i = 1; i < secData.length; i++) {
    if (secData[i][0]) sections.push(String(secData[i][0]));
  }

  /* JSP */
  var jspSheet = ss.getSheetByName(Config.SHEETS.JSP);
  var jspData = jspSheet.getDataRange().getValues();

  /* Événements → par section */
  var evtSheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var evtData = evtSheet.getDataRange().getValues();
  var eventsBySection = {};
  for (var i = 1; i < evtData.length; i++) {
    if (!evtData[i][0]) continue;
    var secs = String(evtData[i][6]).split(',').map(function(s) { return s.trim(); });
    for (var j = 0; j < secs.length; j++) {
      if (!eventsBySection[secs[j]]) eventsBySection[secs[j]] = [];
      eventsBySection[secs[j]].push(String(evtData[i][0]));
    }
  }

  /* Présences */
  var presSheet = ss.getSheetByName(Config.SHEETS.PRESENCES);
  var presData = presSheet.getDataRange().getValues();
  var attMap = {};
  for (var i = 1; i < presData.length; i++) {
    var login = String(presData[i][2]).trim();
    var evtId = String(presData[i][0]);
    if (!attMap[login]) attMap[login] = {};
    attMap[login][evtId] = {
      present: String(presData[i][5]),
      signale: String(presData[i][6])
    };
  }

  /* Construction du bilan par section */
  var bilan = {};
  for (var s = 0; s < sections.length; s++) {
    var sec = sections[s];
    bilan[sec] = [];
    var sectionEvents = eventsBySection[sec] || [];

    for (var i = 1; i < jspData.length; i++) {
      if (String(jspData[i][3]) !== sec) continue;
      var login = String(jspData[i][1]).trim();
      var identite = String(jspData[i][0]);
      var pres = 0, abs = 0, absS = 0, absNS = 0, withAtt = 0;

      for (var e = 0; e < sectionEvents.length; e++) {
        var att = attMap[login] && attMap[login][sectionEvents[e]];
        if (att) {
          withAtt++;
          if (att.present === 'OUI') pres++;
          else {
            abs++;
            if (att.signale === 'OUI') absS++;
            else absNS++;
          }
        }
      }

      bilan[sec].push({
        identite: identite,
        login: login,
        totalEvents: sectionEvents.length,
        eventsWithAttendance: withAtt,
        presences: pres,
        absences: abs,
        absSignalees: absS,
        absNonSignalees: absNS,
        tauxAbsence: withAtt > 0 ? Math.round(abs / withAtt * 1000) / 10 : 0
      });
    }

    bilan[sec].sort(function(a, b) { return a.identite.localeCompare(b.identite); });
  }

  return { sections: sections, bilan: bilan };
}

/* ═══════════════════════════════════════════════════════
   APPEL — Rafraîchir la liste JSP pour des sections données
   ═══════════════════════════════════════════════════════ */
function refreshRollCallJSPs(eventId, selectedSections) {
  var ss = getSS_();

  /* JSP des sections sélectionnées */
  var jspSheet = ss.getSheetByName(Config.SHEETS.JSP);
  var jspData = jspSheet.getDataRange().getValues();
  var jsps = [];
  for (var i = 1; i < jspData.length; i++) {
    if (selectedSections.indexOf(String(jspData[i][3])) !== -1) {
      jsps.push({
        identite: String(jspData[i][0]),
        login: String(jspData[i][1]),
        section: String(jspData[i][3])
      });
    }
  }

  /* Présences existantes pour cet événement */
  var presSheet = ss.getSheetByName(Config.SHEETS.PRESENCES);
  var presData = presSheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < presData.length; i++) {
    if (String(presData[i][0]) === String(eventId)) {
      existing[String(presData[i][2]).trim()] = {
        present: String(presData[i][5]),
        signale: String(presData[i][6]),
        motif: String(presData[i][7] || '')
      };
    }
  }

  /* Fusionner : si absence signalée → pré-cocher absent */
  for (var i = 0; i < jsps.length; i++) {
    var ex = existing[jsps[i].login];
    if (ex) {
      jsps[i].present = ex.present;
      jsps[i].signale = ex.signale;
      jsps[i].motif = ex.motif;
    } else {
      jsps[i].present = 'OUI';
      jsps[i].signale = '';
      jsps[i].motif = '';
    }
  }

  jsps.sort(function(a, b) { return a.identite.localeCompare(b.identite); });
  return { jsps: jsps };
}

/* ═══════════════════════════════════════════════════════
   MISE À JOUR DES SECTIONS D'UN ÉVÉNEMENT
   ═══════════════════════════════════════════════════════ */
function updateEventSections(eventId, newSections) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(Config.SHEETS.EVENTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(eventId)) {
      sheet.getRange(i + 1, 7).setValue(newSections.join(', '));
      return { success: true };
    }
  }
  return { error: 'Événement non trouvé' };
}

/* ═══════════════════════════════════════════════════════
   NOTIFICATION PAR MAIL — Envoi lors de la création
   ═══════════════════════════════════════════════════════ */
function sendEventNotification_(eventId, eventData) {
  var ss = getSS_();
  var webAppUrl = ScriptApp.getService().getUrl();
  var eventLink = webAppUrl + '?event=' + eventId;

  /* Formater la date */
  var parts = eventData.date.split('-');
  var moisNoms = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
  var dateLabel = parseInt(parts[2]) + ' ' + moisNoms[parseInt(parts[1]) - 1] + ' ' + parts[0];

  /* Récupérer les emails des JSP concernés (col E = index 4) */
  var jspSheet = ss.getSheetByName(Config.SHEETS.JSP);
  var jspData = jspSheet.getDataRange().getValues();
  var recipients = [];
  for (var i = 1; i < jspData.length; i++) {
    var jspSection = String(jspData[i][3]);
    if (eventData.sections.indexOf(jspSection) !== -1) {
      var email = jspData[i][4] ? String(jspData[i][4]).trim() : '';
      if (email && email.indexOf('@') !== -1) {
        recipients.push(email);
      }
    }
  }

  /* Récupérer les emails des référents des sections concernées */
  var secSheet = ss.getSheetByName(Config.SHEETS.SECTIONS);
  var secData = secSheet.getDataRange().getValues();
  var refNames = [];
  for (var i = 1; i < secData.length; i++) {
    if (eventData.sections.indexOf(String(secData[i][0])) !== -1) {
      var names = String(secData[i][1]).split(',');
      for (var j = 0; j < names.length; j++) {
        var name = names[j].trim();
        if (name && refNames.indexOf(name) === -1) refNames.push(name);
      }
    }
  }
  var refSheet = ss.getSheetByName(Config.SHEETS.REFERENTS);
  var refData = refSheet.getDataRange().getValues();
  for (var i = 1; i < refData.length; i++) {
    var refName = String(refData[i][0]).trim();
    if (refNames.indexOf(refName) !== -1) {
      var email = refData[i][3] ? String(refData[i][3]).trim() : '';
      if (email && email.indexOf('@') !== -1 && recipients.indexOf(email) === -1) {
        recipients.push(email);
      }
    }
  }

  if (recipients.length === 0) {
    Logger.log('Notification mail : aucun destinataire avec email trouvé');
    return;
  }

  var subject = '🔥 JSP St Cyprien — Nouvel événement : ' + eventData.nom;

  var htmlBody = '<!DOCTYPE html>'
    + '<html><body style="margin:0;padding:0;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f1f5f9;padding:30px 0;">'
    + '<tr><td align="center">'
    + '<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">'
    /* En-tête rouge */
    + '<tr><td style="background:linear-gradient(135deg,#dc2626,#991b1b);padding:35px 30px;text-align:center;border-radius:16px 16px 0 0;">'
    + '<h1 style="color:#ffffff;margin:0;font-size:28px;">🔥 JSP St Cyprien</h1>'
    + '<p style="color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:15px;">Nouvel événement programmé</p>'
    + '</td></tr>'
    /* Corps */
    + '<tr><td style="background:#ffffff;padding:35px 30px;border-left:1px solid #e2e8f0;border-right:1px solid #e2e8f0;">'
    + '<h2 style="color:#1e293b;margin:0 0 24px;font-size:22px;">' + eventData.nom + '</h2>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:28px;">'
    + '<tr><td style="padding:12px 0;border-bottom:1px solid #f1f5f9;"><span style="font-size:18px;">📅</span><span style="color:#1e293b;font-size:15px;font-weight:600;margin-left:12px;">' + dateLabel + '</span></td></tr>'
    + '<tr><td style="padding:12px 0;border-bottom:1px solid #f1f5f9;"><span style="font-size:18px;">🕐</span><span style="color:#1e293b;font-size:15px;margin-left:12px;">' + eventData.heureDebut + ' — ' + eventData.heureFin + '</span></td></tr>'
    + '<tr><td style="padding:12px 0;border-bottom:1px solid #f1f5f9;"><span style="font-size:18px;">📍</span><span style="color:#1e293b;font-size:15px;margin-left:12px;">' + (eventData.lieu || 'CIS St Cyprien') + '</span></td></tr>'
    + '<tr><td style="padding:12px 0;"><span style="font-size:18px;">👥</span><span style="color:#1e293b;font-size:15px;margin-left:12px;">' + eventData.sections.join(', ') + '</span></td></tr>'
    + '</table>'
    /* Bouton CTA */
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:10px 0 24px;">'
    + '<a href="' + eventLink + '" style="display:inline-block;background:#2563eb;color:#ffffff;padding:16px 32px;border-radius:10px;text-decoration:none;font-weight:700;font-size:15px;line-height:1.5;">'
    + 'Cliquez ici pour accéder à l&#39;événement<br>et signaler votre présence ou absence'
    + '</a></td></tr></table>'
    + '<p style="color:#94a3b8;font-size:12px;text-align:center;margin:0;">'
    + 'Si le bouton ne fonctionne pas, copiez ce lien :<br>'
    + '<a href="' + eventLink + '" style="color:#2563eb;word-break:break-all;font-size:11px;">' + eventLink + '</a></p>'
    + '</td></tr>'
    /* Pied de page */
    + '<tr><td style="background:#f8fafc;padding:20px 30px;text-align:center;border-radius:0 0 16px 16px;border:1px solid #e2e8f0;border-top:none;">'
    + '<p style="color:#94a3b8;font-size:13px;margin:0;">🔥 JSP St Cyprien — Suivi d&#39;activité</p>'
    + '<p style="color:#cbd5e1;font-size:11px;margin:6px 0 0;">Cet email a été envoyé automatiquement.</p>'
    + '</td></tr>'
    + '</table></td></tr></table></body></html>';

  for (var i = 0; i < recipients.length; i++) {
    try {
      MailApp.sendEmail({
        to: recipients[i],
        subject: subject,
        htmlBody: htmlBody
      });
    } catch(e) {
      Logger.log('Erreur envoi mail à ' + recipients[i] + ': ' + e.message);
    }
  }
  Logger.log('Notifications envoyées à ' + recipients.length + ' destinataire(s)');
}

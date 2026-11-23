/**
 * iPad Pano: Datendienst (Google Apps Script, an eine Google-Tabelle gebunden).
 * Liefert: Google-Kalender-Termine, Notizen, Countdowns, Fotos aus einem Drive-Ordner.
 */

var BLATT = { notizen: 'Notizen', countdown: 'Countdown', einst: 'Einstellungen' };

/** Einmal manuell ausführen: legt die Blätter und den Token an. */
function einrichten() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var n = blattHolen(ss, BLATT.notizen);
  if (n.getLastRow() === 0) {
    n.getRange(1, 1, 3, 2).setValues([['Notiz', 'Erledigt'], ['Beispielnotiz', false], ['Zweite Notiz', false]]);
    n.getRange('B2:B200').insertCheckboxes();
  }

  var c = blattHolen(ss, BLATT.countdown);
  if (c.getLastRow() === 0) {
    c.getRange(1, 1, 2, 3).setValues([['Name', 'Datum', 'Jährlich'], ['Silvester', new Date(new Date().getFullYear(), 11, 31), true]]);
    c.getRange('B2:B200').setNumberFormat('dd.MM.yyyy');
    c.getRange('C2:C200').insertCheckboxes();
  }

  var e = blattHolen(ss, BLATT.einst);
  if (e.getLastRow() === 0) {
    e.getRange(1, 1, 6, 3).setValues([
      ['Schlüssel', 'Wert', 'Erklärung'],
      ['token', Utilities.getUuid().replace(/-/g, ''), 'Geheimer Schlüssel. Wird bei der iPad-Einrichtung eingegeben.'],
      ['stadt', '', 'Stadt für das Wetter (z. B. Köln). Leer = Standort des iPads.'],
      ['foto_ordner_id', '', 'ID des Google-Drive-Ordners (in der Adresse nach /folders/).'],
      ['kalender', '', 'Kalender-IDs, mit Komma getrennt. Leer = alle sichtbaren Kalender.'],
      ['tage', 14, 'Wie viele Tage im Voraus angezeigt werden.']
    ]);
  }
  SpreadsheetApp.getUi().alert('Einrichtung fertig. Den Token im Blatt "Einstellungen" notieren.');
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var einst = einstellungenLesen();
  if (!einst.token || p.k !== einst.token) return json({ hata: 'yetki' });
  try {
    if (p.foto) return json(fotoHolen(p.foto, einst));
    return json(daten(einst));
  } catch (err) {
    return json({ hata: String(err) });
  }
}

function daten(einst) {
  return {
    zaman: new Date().toISOString(),
    sehir: einst.stadt || '',
    etkinlikler: termine(einst),
    notlar: notizen(),
    sayimlar: countdowns(),
    fotolar: fotoListe(einst)
  };
}

function termine(einst) {
  var tz = Session.getScriptTimeZone();
  var start = new Date(); start.setHours(0, 0, 0, 0);
  var tage = Number(einst.tage) || 14;
  var ende = new Date(start.getTime() + tage * 864e5);

  var kalender;
  if (einst.kalender) {
    kalender = String(einst.kalender).split(',').map(function (id) {
      return CalendarApp.getCalendarById(id.trim());
    }).filter(Boolean);
  } else {
    kalender = CalendarApp.getAllCalendars().filter(function (k) { return !k.isHidden(); });
  }

  var liste = [];
  kalender.forEach(function (k) {
    var farbe = k.getColor();
    k.getEvents(start, ende).forEach(function (ev) {
      var ganz = ev.isAllDayEvent();
      liste.push({
        t: ev.getTitle() || '(Ohne Titel)',
        g: ganz,
        b: ganz ? Utilities.formatDate(ev.getAllDayStartDate(), tz, 'yyyy-MM-dd') : ev.getStartTime().toISOString(),
        s: ganz ? Utilities.formatDate(ev.getAllDayEndDate(), tz, 'yyyy-MM-dd') : ev.getEndTime().toISOString(),
        r: farbe
      });
    });
  });
  liste.sort(function (x, y) { return x.b < y.b ? -1 : x.b > y.b ? 1 : 0; });
  return liste.slice(0, 60);
}

function notizen() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BLATT.notizen);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim() !== '' && r[1] !== true; })
    .map(function (r) { return String(r[0]).trim(); });
}

function countdowns() {
  var tz = Session.getScriptTimeZone();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BLATT.countdown);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return String(r[0]).trim() !== '' && r[1] instanceof Date; })
    .map(function (r) {
      return { ad: String(r[0]).trim(), tarih: Utilities.formatDate(r[1], tz, 'yyyy-MM-dd'), yillik: r[2] === true };
    });
}

function fotoListe(einst) {
  if (!einst.foto_ordner_id) return [];
  var cache = CacheService.getScriptCache();
  var schl = 'fotos_' + einst.foto_ordner_id;
  var gespeichert = cache.get(schl);
  if (gespeichert) return JSON.parse(gespeichert);

  var ids = [];
  var it = DriveApp.getFolderById(einst.foto_ordner_id).getFiles();
  while (it.hasNext() && ids.length < 300) {
    var f = it.next();
    if (/^image\//.test(f.getMimeType())) ids.push(f.getId());
  }
  cache.put(schl, JSON.stringify(ids), 600);
  return ids;
}

/** Liefert nur Fotos aus dem eingestellten Ordner, verkleinert auf ca. 1600 px. */
function fotoHolen(id, einst) {
  if (fotoListe(einst).indexOf(id) === -1) return { hata: 'foto' };
  var auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var blob = null;
  try {
    var meta = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink', { headers: auth }
    ).getContentText());
    if (meta.thumbnailLink) {
      var r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s1600'), { headers: auth, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) blob = r.getBlob();
    }
  } catch (e) { /* ohne Vorschaubild: Originaldatei */ }
  if (!blob) blob = DriveApp.getFileById(id).getBlob();
  return { mime: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) };
}

function einstellungenLesen() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BLATT.einst);
  var o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0]) o[String(r[0]).trim()] = typeof r[1] === 'string' ? r[1].trim() : r[1];
  });
  return o;
}

function blattHolen(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

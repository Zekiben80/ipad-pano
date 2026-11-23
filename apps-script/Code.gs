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
    if (p.foto) return json(fotoHolen(p.foto, einst, p.gr));
    return json(daten(einst));
  } catch (err) {
    return json({ hata: String(err) });
  }
}

var BEREICHE = ['kalender', 'wetter', 'fotos', 'countdown'];

/** Admin-Seite: Änderungen speichern. Body (text/plain): {"k":TOKEN, "anzeige":{...}, "stadt":"...", "countdowns":[...]} */
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return json({ hata: 'format' }); }
  var einst = einstellungenLesen();
  if (!einst.token || body.k !== einst.token) return json({ hata: 'yetki' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (body.anzeige) {
      BEREICHE.forEach(function (n) {
        if (n in body.anzeige) einstellungSetzen('zeige_' + n, body.anzeige[n] === true, 'Auf dem Pano anzeigen (Admin-Seite).');
      });
    }
    if (typeof body.stadt === 'string') einstellungSetzen('stadt', body.stadt.trim().slice(0, 80));
    if (Array.isArray(body.countdowns)) countdownsSchreiben(body.countdowns);
    if (body.fotoNeu) fotoSpeichern(body.fotoNeu, einst);
    if (body.fotoLoeschen) fotoLoeschen(body.fotoLoeschen, einst);
    SpreadsheetApp.flush();
    return json(daten(einstellungenLesen()));
  } catch (err) {
    return json({ hata: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function einstellungSetzen(name, wert, erklaerung) {
  var sh = blattHolen(SpreadsheetApp.getActiveSpreadsheet(), BLATT.einst);
  var n = sh.getLastRow();
  var namen = n >= 2 ? sh.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  var i = namen.indexOf(name);
  var zeile = i === -1 ? n + 1 : i + 2;
  if (i === -1) sh.getRange(zeile, 1, 1, 3).setValues([[name, '', erklaerung || '']]);
  if (typeof wert === 'boolean') sh.getRange(zeile, 2).insertCheckboxes();
  sh.getRange(zeile, 2).setValue(wert);
}

/** Foto von der Admin-Seite (Base64-JPEG, bereits verkleinert) im Foto-Ordner speichern. */
function fotoSpeichern(foto, einst) {
  if (!einst.foto_ordner_id) throw new Error('Kein Foto-Ordner eingestellt');
  if (!foto.data || foto.data.length > 15e6) throw new Error('Foto ungültig oder zu groß');
  var name = String(foto.name || 'Foto').replace(/[\\\/:*?"<>|]/g, '').slice(0, 80).replace(/\.\w+$/, '') + '.jpg';
  var blob = Utilities.newBlob(Utilities.base64Decode(foto.data), 'image/jpeg', name);
  DriveApp.getFolderById(einst.foto_ordner_id).createFile(blob);
  CacheService.getScriptCache().remove('fotos_' + einst.foto_ordner_id);
}

/** Foto in den Drive-Papierkorb verschieben (dort 30 Tage wiederherstellbar). */
function fotoLoeschen(id, einst) {
  if (fotoListe(einst).indexOf(id) === -1) throw new Error('Foto nicht gefunden');
  DriveApp.getFileById(id).setTrashed(true);
  CacheService.getScriptCache().remove('fotos_' + einst.foto_ordner_id);
}

function countdownsSchreiben(liste) {
  var sh = blattHolen(SpreadsheetApp.getActiveSpreadsheet(), BLATT.countdown);
  var zeilen = liste.filter(function (x) {
    return x && String(x.ad || '').trim() && /^\d{4}-\d{2}-\d{2}$/.test(x.tarih);
  }).slice(0, 100).map(function (x) {
    var p = x.tarih.split('-');
    return [String(x.ad).trim().slice(0, 60), new Date(+p[0], +p[1] - 1, +p[2]), x.yillik === true];
  });
  var n = sh.getLastRow();
  if (n >= 2) sh.getRange(2, 1, n - 1, 3).clearContent();
  if (zeilen.length) {
    sh.getRange(2, 1, zeilen.length, 3).setValues(zeilen);
    sh.getRange(2, 2, zeilen.length, 1).setNumberFormat('dd.MM.yyyy');
  }
}

function daten(einst) {
  var anzeige = {};
  BEREICHE.forEach(function (n) { anzeige[n] = einst['zeige_' + n] !== false; });
  return {
    zaman: new Date().toISOString(),
    anzeige: anzeige,
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

/** Liefert nur Fotos aus dem eingestellten Ordner, verkleinert (Standard 1600 px, Vorschau z. B. 300 px). */
function fotoHolen(id, einst, gr) {
  if (fotoListe(einst).indexOf(id) === -1) return { hata: 'foto' };
  var groesse = Math.min(1600, Math.max(200, Number(gr) || 1600));
  var vorschau = groesse < 1600;
  var auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var blob = null;
  try {
    var meta = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink', { headers: auth }
    ).getContentText());
    if (meta.thumbnailLink) {
      var r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s' + groesse), { headers: auth, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) blob = r.getBlob();
    }
  } catch (e) { /* ohne Vorschaubild: Originaldatei */ }
  if (!blob && vorschau) return { hata: 'vorschau' };
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

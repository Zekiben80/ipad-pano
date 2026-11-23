/**
 * iPad Pano: veri servisi (Google Apps Script, Google Tablosuna bağlı).
 * Verir: Google Takvim etkinlikleri, Notlar, Geri sayımlar, Drive klasöründeki fotoğraflar.
 * Kurulum adımları: KURULUM.md
 */

var SAYFA = { notlar: 'Notlar', sayim: 'GeriSayim', ayar: 'Ayarlar' };

/** Bir kez elle çalıştır: sayfaları ve anahtarı oluşturur. */
function kurulum() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var n = sayfaAl(ss, SAYFA.notlar);
  if (n.getLastRow() === 0) {
    n.getRange(1, 1, 3, 2).setValues([['Not', 'Bitti'], ['Örnek not', false], ['İkinci not', false]]);
    n.getRange('B2:B200').insertCheckboxes();
  }

  var s = sayfaAl(ss, SAYFA.sayim);
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 2, 3).setValues([['Ad', 'Tarih', 'Her yıl'], ['Yılbaşı', new Date(new Date().getFullYear(), 11, 31), true]]);
    s.getRange('B2:B200').setNumberFormat('dd.MM.yyyy');
    s.getRange('C2:C200').insertCheckboxes();
  }

  var a = sayfaAl(ss, SAYFA.ayar);
  if (a.getLastRow() === 0) {
    a.getRange(1, 1, 6, 3).setValues([
      ['Anahtar', 'Değer', 'Açıklama'],
      ['token', Utilities.getUuid().replace(/-/g, ''), 'Gizli anahtar. iPad kurulumunda girilir.'],
      ['sehir', '', 'Hava durumu için şehir (ör. Köln). Boşsa iPad konumu kullanılır.'],
      ['foto_klasor_id', '', 'Google Drive klasör kimliği (adres çubuğunda /folders/ sonrası).'],
      ['takvimler', '', 'Virgülle takvim kimlikleri. Boşsa tüm görünür takvimler.'],
      ['gun_sayisi', 14, 'Kaç gün ilerisi gösterilsin.']
    ]);
  }
  SpreadsheetApp.getUi().alert('Kurulum tamam. Ayarlar sayfasındaki token değerini not al.');
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var ayar = ayarlariOku();
  if (!ayar.token || p.k !== ayar.token) return json({ hata: 'yetki' });
  try {
    if (p.foto) return json(fotoGetir(p.foto, ayar));
    return json(veri(ayar));
  } catch (err) {
    return json({ hata: String(err) });
  }
}

function veri(ayar) {
  return {
    zaman: new Date().toISOString(),
    sehir: ayar.sehir || '',
    etkinlikler: etkinlikler(ayar),
    notlar: notlar(),
    sayimlar: sayimlar(),
    fotolar: fotoListesi(ayar)
  };
}

function etkinlikler(ayar) {
  var tz = Session.getScriptTimeZone();
  var bas = new Date(); bas.setHours(0, 0, 0, 0);
  var gun = Number(ayar.gun_sayisi) || 14;
  var bit = new Date(bas.getTime() + gun * 864e5);

  var takvimler;
  if (ayar.takvimler) {
    takvimler = String(ayar.takvimler).split(',').map(function (id) {
      return CalendarApp.getCalendarById(id.trim());
    }).filter(Boolean);
  } else {
    takvimler = CalendarApp.getAllCalendars().filter(function (c) { return !c.isHidden(); });
  }

  var sonuc = [];
  takvimler.forEach(function (c) {
    var renk = c.getColor();
    c.getEvents(bas, bit).forEach(function (ev) {
      var tumGun = ev.isAllDayEvent();
      sonuc.push({
        t: ev.getTitle() || '(Başlıksız)',
        g: tumGun,
        b: tumGun ? Utilities.formatDate(ev.getAllDayStartDate(), tz, 'yyyy-MM-dd') : ev.getStartTime().toISOString(),
        s: tumGun ? Utilities.formatDate(ev.getAllDayEndDate(), tz, 'yyyy-MM-dd') : ev.getEndTime().toISOString(),
        r: renk
      });
    });
  });
  sonuc.sort(function (x, y) { return x.b < y.b ? -1 : x.b > y.b ? 1 : 0; });
  return sonuc.slice(0, 60);
}

function notlar() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SAYFA.notlar);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim() !== '' && r[1] !== true; })
    .map(function (r) { return String(r[0]).trim(); });
}

function sayimlar() {
  var tz = Session.getScriptTimeZone();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SAYFA.sayim);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return String(r[0]).trim() !== '' && r[1] instanceof Date; })
    .map(function (r) {
      return { ad: String(r[0]).trim(), tarih: Utilities.formatDate(r[1], tz, 'yyyy-MM-dd'), yillik: r[2] === true };
    });
}

function fotoListesi(ayar) {
  if (!ayar.foto_klasor_id) return [];
  var cache = CacheService.getScriptCache();
  var anahtar = 'fotolar_' + ayar.foto_klasor_id;
  var kayit = cache.get(anahtar);
  if (kayit) return JSON.parse(kayit);

  var ids = [];
  var it = DriveApp.getFolderById(ayar.foto_klasor_id).getFiles();
  while (it.hasNext() && ids.length < 300) {
    var f = it.next();
    if (/^image\//.test(f.getMimeType())) ids.push(f.getId());
  }
  cache.put(anahtar, JSON.stringify(ids), 600);
  return ids;
}

/** Yalnızca ayarlı klasördeki bir fotoğrafı, ~1600 px küçültülmüş olarak verir. */
function fotoGetir(id, ayar) {
  if (fotoListesi(ayar).indexOf(id) === -1) return { hata: 'foto' };
  var yetki = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var blob = null;
  try {
    var meta = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink', { headers: yetki }
    ).getContentText());
    if (meta.thumbnailLink) {
      var r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s1600'), { headers: yetki, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) blob = r.getBlob();
    }
  } catch (e) { /* küçük resim yoksa asıl dosyaya düş */ }
  if (!blob) blob = DriveApp.getFileById(id).getBlob();
  return { mime: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) };
}

function ayarlariOku() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SAYFA.ayar);
  var o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0]) o[String(r[0]).trim()] = typeof r[1] === 'string' ? r[1].trim() : r[1];
  });
  return o;
}

function sayfaAl(ss, ad) { return ss.getSheetByName(ad) || ss.insertSheet(ad); }

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// D-Buchung: Automatische Gmail-Anfragen-Erkennung
// Laeuft 1x taeglich auf dem Gmail-Konto "D Buchung"
// ============================================================

// --- KONFIGURATION ---
var PROCESSED_LABEL = 'AB-Verarbeitet';
var SHEET_NAME = 'Anfragen';
var GMAIL_SEARCH_DAYS = 14;      // wie weit zurueck nach unverarbeiteten Mails suchen
var SHARED_TOKEN = 'dv8-Kx92mPqR7w'; // einfacher Zugriffsschutz, gleicher Token wie in D-Buchung
var CLAUDE_MODEL = 'claude-sonnet-4-6';

function getApiKey(){
  return PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
}

// === Haupt-Funktion: per Trigger 1x taeglich aufrufen ===
function dailyCheckInquiries(){
  var label = getOrCreateLabel(PROCESSED_LABEL);
  var sheet = getSheet();
  var threads = GmailApp.search('in:inbox -label:' + PROCESSED_LABEL + ' newer_than:' + GMAIL_SEARCH_DAYS + 'd');
  Logger.log('Gefundene Threads: ' + threads.length);
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    try {
      var msgs = thread.getMessages();
      var msg = msgs[msgs.length - 1]; // neueste Nachricht im Thread
      var subject = msg.getSubject();
      var body = msg.getPlainBody();
      var fromEmail = msg.getFrom();
      var data = extractBookingData(subject, body, fromEmail);
      if (data) {
        appendToSheet(sheet, data, thread.getId());
        Logger.log('Anfrage erkannt von: ' + data.gast);
      } else {
        Logger.log('Keine Buchungsanfrage erkannt (oder Fehler): ' + subject);
      }
    } catch (e) {
      Logger.log('Fehler bei Thread: ' + e.message);
    }
    thread.addLabel(label); // immer als verarbeitet markieren, auch wenn keine Anfrage erkannt wurde
  }
}

function getOrCreateLabel(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ThreadID','Datum','Objekt','Gast','Email','Telefon','Anreise','Abreise','Personen','Hunde','Quelle','Hinweis','Importiert']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// === Mailtext per Claude API analysieren ===
function extractBookingData(subject, body, fromEmail) {
  var apiKey = getApiKey();
  if (!apiKey) { Logger.log('FEHLER: ANTHROPIC_API_KEY fehlt in den Script-Properties!'); return null; }

  var prompt = "Du bekommst eine E-Mail an eine Vermietung von Ferienwohnungen auf Ruegen (Moenchgut, Ostsee). " +
    "Die Wohnungen heissen: Duenenvilla 3 (DV3), Duenenvilla 5 (DV5), Strandvilla 8a (SV8a), Strandvilla 8b (SV8b). " +
    "Pruefe ob es eine BUCHUNGSANFRAGE eines Gastes ist (nicht Werbung, Newsletter, Rechnung, interne Mail o.ae.).\n\n" +
    "Antworte AUSSCHLIESSLICH mit einem JSON-Objekt, OHNE Markdown-Codeblock, OHNE Erklaerung, exakt in dieser Form:\n" +
    '{"ist_buchungsanfrage": true, "objekt": "dv3", "gast": "Vorname Nachname", "anreise": "YYYY-MM-DD", "abreise": "YYYY-MM-DD", "personen": 2, "hunde": 0, "email": "", "telefon": "", "quelle": "", "hinweis": ""}\n\n' +
    "Regeln:\n" +
    "- objekt: dv3, dv5, sv8a, sv8b oder \"unklar\" wenn nicht erkennbar\n" +
    "- anreise/abreise: Format YYYY-MM-DD. Wenn kein Jahr angegeben ist, nimm das naechste passende zukuenftige Jahr an.\n" +
    "- email/telefon: aus der Mail extrahieren falls vorhanden, sonst leer lassen\n" +
    "- quelle: z.B. Name der Buchungsplattform falls erkennbar, sonst leer\n" +
    "- hinweis: kurzer Hinweis (max 1 Satz) falls etwas unsicher/unklar geschaetzt wurde, sonst leer lassen\n" +
    "- Falls KEINE Buchungsanfrage: {\"ist_buchungsanfrage\": false}\n\n" +
    "Absender: " + fromEmail + "\n" +
    "Betreff: " + subject + "\n\n" +
    "Text:\n" + body.substring(0, 4000);

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: 500,
    messages: [{ role: 'user', content: prompt }]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  var code = resp.getResponseCode();
  if (code !== 200) { Logger.log('Claude API Fehler ' + code + ': ' + resp.getContentText()); return null; }

  var json = JSON.parse(resp.getContentText());
  if (!json.content || !json.content[0] || !json.content[0].text) return null;

  var text = json.content[0].text.trim();
  text = text.replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();

  var data;
  try { data = JSON.parse(text); } catch (e) { Logger.log('JSON-Parse-Fehler: ' + text); return null; }

  if (!data.ist_buchungsanfrage) return null;
  if (!data.email) data.email = extractEmailFromHeader(fromEmail);
  return data;
}

function extractEmailFromHeader(fromHeader) {
  var m = fromHeader.match(/<(.+)>/);
  return m ? m[1] : fromHeader;
}

function appendToSheet(sheet, data, threadId) {
  // Duplikat-Schutz: gleicher Thread nicht doppelt eintragen
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === threadId) return;
  }
  sheet.appendRow([
    threadId,
    new Date(),
    data.objekt || 'unklar',
    data.gast || '',
    data.email || '',
    data.telefon || '',
    data.anreise || '',
    data.abreise || '',
    data.personen || 2,
    data.hunde || 0,
    data.quelle || 'E-Mail',
    data.hinweis || '',
    false // Importiert
  ]);
}

// ============================================================
// WEB APP ENDPUNKT (fuer D-Buchung)
// ============================================================
function doGet(e) {
  if (!e.parameter.token || e.parameter.token !== SHARED_TOKEN) {
    return json_({ error: 'unauthorized' });
  }
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[12]) continue; // bereits importiert -> nicht erneut liefern
    out.push({
      id: r[0], row: i + 1, objekt: r[2], gast: r[3], email: r[4], telefon: r[5],
      von: fmtDate_(r[6]), bis: fmtDate_(r[7]), pers: r[8], hunde: r[9], quelle: r[10], hinweis: r[11]
    });
  }
  return json_(out);
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad request' }); }
  if (!body.token || body.token !== SHARED_TOKEN) return json_({ error: 'unauthorized' });

  if (body.action === 'markImported' && body.rows && body.rows.length) {
    var sheet = getSheet();
    body.rows.forEach(function (rowNum) {
      sheet.getRange(rowNum, 13).setValue(true); // Spalte 13 = Importiert
    });
    return json_({ ok: true, count: body.rows.length });
  }
  return json_({ error: 'unknown action' });
}

function fmtDate_(v) {
  if (!v) return '';
  if (typeof v === 'string') return v;
  return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// === Zum manuellen Testen: einmal in der Apps-Script-Oberflaeche ausfuehren ===
function testRun() {
  dailyCheckInquiries();
}

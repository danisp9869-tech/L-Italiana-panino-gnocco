/* ─────────────────────────────────────────────────────────────────────────
   L'ITALIANA — Endpoint prenotazioni (Google Apps Script)
   Scrive la prenotazione sul Foglio Google e invia l'email di riepilogo al
   ristorante. Facoltativa (e disattivata di default) la conferma al cliente.

   COME PUBBLICARLO
   1. Apri il Foglio Google dedicato → Estensioni → Apps Script.
   2. Incolla TUTTO questo codice (sostituendo l'esistente) e Salva.
   3. Esegui una volta la funzione  verificaFoglio  (menu Esegui) e concedi
      i permessi richiesti (Fogli + Gmail).
   4. Distribuisci → Nuova distribuzione → Tipo: App web
        · Esegui come:  Io (il tuo account)
        · Chi ha accesso:  Chiunque
      Copia l'URL che finisce con  /exec.
   5. Incolla quell'URL nella landing, in  index.html , alla riga
        var SHEET_URL = "";
      (dentro lo script del form) e ripubblica la landing.
   ───────────────────────────────────────────────────────────────────────── */

// ── CONFIGURAZIONE ────────────────────────────────────────────────────────
var SHEET_ID          = '1NukO7lMPwc2WDajw_5GxWUjXFSlm73zUUbBoP2kfjxM';
var RESTAURANT_EMAIL  = 'federicogiampa@gmail.com';   // riceve la notifica di prenotazione

// Mittente delle email in uscita: INTERRUTTORE della conferma al cliente.
// Di default le email partono dall'account che possiede lo script. Per farle
// partire da un altro indirizzo serve un ALIAS VERIFICATO di quell'account
// (Gmail → Impostazioni → Account → "Invia messaggi come"). Finché è vuoto,
// la conferma al CLIENTE non viene inviata (parte solo la notifica al ristorante).
var MITTENTE_ALIAS    = '';

var RESTAURANT_NOME      = "L'Italiana";
var RESTAURANT_TEL       = '+39 0173 216770';
var RESTAURANT_INDIRIZZO = "Via Roma 45, 12040 Piobesi d'Alba (CN)";
var RESTAURANT_MAPS      = 'https://www.google.com/maps/search/?api=1&query=Via+Roma+45+Piobesi+d%27Alba+CN';

// Colonne del foglio (in quest'ordine). Comprende tutti i dati richiesti:
// Nome, Cognome, Data, Coperti, Telefono, Email, Privacy, Marketing.
var INTESTAZIONI = ['Data', 'Ora', 'Coperti', 'Nome', 'Telefono', 'Email', 'Richieste', 'Privacy', 'Marketing', 'Offerta', 'Creata'];

var TZ = 'Europe/Rome';

// ── UTILITÀ DATE ──────────────────────────────────────────────────────────
// "2026-09-05" → "05/09/2026". La landing manda la data in ISO (input date).
function dataIta_(v) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v || ''));
  return m ? m[3] + '/' + m[2] + '/' + m[1] : (v || '');
}
function creataIta_() {
  return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
}

// ── RICEZIONE PRENOTAZIONE ────────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var p  = e.parameter || {};
    var sheet = ss.getSheetByName('Prenotazioni') || ss.insertSheet('Prenotazioni');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(INTESTAZIONI);
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      dataIta_(p.data), p.ora || '', p.persone || '',
      p.nome || '', p.telefono || '', p.email || '',
      p.richieste || '', p.privacy || '', p.marketing || '',
      p.offerta || '', creataIta_()
    ]);

    // Notifica al ristorante (riepilogo completo)
    try {
      var nomeCompleto = (p.nome || '').trim();
      var subj = '🍔 ' + (p.offerta ? p.offerta + ' — ' : '') + 'Nuova prenotazione — ' +
                 nomeCompleto + ' · ' + dataIta_(p.data) + ' ' + (p.ora || '') + ' · ' + (p.persone || '') + 'p';
      var html =
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;line-height:1.5">' +
          '<h2 style="margin:0 0 4px">Nuova prenotazione — ' + RESTAURANT_NOME + '</h2>' +
          '<p style="margin:0 0 14px;color:#666">Ricevuta dal sito delle prenotazioni</p>' +
          '<table cellpadding="7" style="border-collapse:collapse;border:1px solid #eee">' +
            trEmail_('Offerta', p.offerta) +
            trEmail_('Data', dataIta_(p.data)) +
            trEmail_('Orario', p.ora) +
            trEmail_('Coperti', p.persone) +
            trEmail_('Nome', p.nome) +
            trEmail_('Telefono', p.telefono) +
            trEmail_('Email', p.email) +
            trEmail_('Richieste', p.richieste) +
            trEmail_('Consenso privacy', p.privacy) +
            trEmail_('Consenso marketing', p.marketing) +
            trEmail_('Ricevuta il', creataIta_()) +
          '</table>' +
        '</div>';
      var opts = opzioniInvio_({ htmlBody: html });
      if (p.email && p.email.indexOf('@') > 0) opts.replyTo = p.email;   // rispondi = scrivi al cliente
      MailApp.sendEmail(RESTAURANT_EMAIL, subj, 'Nuova prenotazione ricevuta.', opts);
    } catch (mailErr) {}

    // Conferma al cliente (facoltativa: solo se è impostato un alias verificato)
    try { inviaConfermaCliente_(p); } catch (clienteErr) {}

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ── EMAIL: UTILITÀ MITTENTE ───────────────────────────────────────────────
var _alias = null;
function aliasUsabile_(indirizzo) {
  if (!indirizzo) return false;
  if (_alias === null) { try { _alias = GmailApp.getAliases(); } catch (e) { _alias = []; } }
  return _alias.indexOf(indirizzo) !== -1;
}
function opzioniInvio_(extra) {
  var o = extra || {};
  o.name = 'Prenotazioni ' + RESTAURANT_NOME;
  if (aliasUsabile_(MITTENTE_ALIAS)) o.from = MITTENTE_ALIAS;
  return o;
}
function emailValida_(v) {
  return /^[^@\s]+@[^@\s.]+\.[^@\s]+$/.test(String(v || '').trim());
}
function trEmail_(label, value) {
  if (!value) return '';
  return '<tr><td style="border:1px solid #eee;color:#666;font-weight:bold">' + label +
         '</td><td style="border:1px solid #eee">' + value + '</td></tr>';
}

// ── CONFERMA AL CLIENTE (facoltativa) ─────────────────────────────────────
function inviaConfermaCliente_(p) {
  var dest = String(p.email || '').trim();
  if (!emailValida_(dest)) return false;
  if (!aliasUsabile_(MITTENTE_ALIAS)) return false;   // senza alias verificato non si invia

  var primoNome = String(p.nome || '').trim().split(' ')[0];
  var ciao = primoNome ? 'Ciao ' + primoNome + ' ☺️' : 'Ciao ☺️';
  var data = dataIta_(p.data);
  var quando = data + (p.ora ? ' alle ' + p.ora : '');
  var oggetto = 'Prenotazione confermata da ' + RESTAURANT_NOME + (data ? ' — ' + data : '');
  var righe = [
    ciao,
    'La tua prenotazione è confermata: il tuo posto è riservato. 🎉',
    'Ti aspetta uno gnocco fritto con lardo alle erbe in omaggio per ogni Maxi Montanaro. 🍔',
    (quando ? '📅 ' + quando + '\n' : '') + '📍 ' + RESTAURANT_INDIRIZZO,
    "Se non dovessi riuscire a venire, faccelo sapere in anticipo così liberiamo il tavolo per qualcun altro.",
    'Ci vediamo presto!'
  ];

  var linkMaps = '<a href="' + RESTAURANT_MAPS + '" style="color:#ffb200">' + RESTAURANT_INDIRIZZO + '</a>';
  var paragrafi = righe.map(function (r) {
    return '<p style="margin:0 0 14px">' +
      r.split(RESTAURANT_INDIRIZZO).join(linkMaps).split('\n').join('<br>') + '</p>';
  }).join('');

  var nomeCompleto = (p.nome || '').trim();
  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;line-height:1.55;max-width:520px">' +
      paragrafi +
      '<p style="margin:22px 0 8px;font-weight:bold">Riepilogo della prenotazione</p>' +
      '<table cellpadding="7" style="border-collapse:collapse;border:1px solid #eee">' +
        trEmail_('Data', data) + trEmail_('Orario', p.ora) + trEmail_('Coperti', p.persone) +
        trEmail_('A nome di', nomeCompleto) + trEmail_('Telefono', p.telefono) +
        trEmail_('Richieste', p.richieste) +
      '</table>' +
      '<p style="margin:16px 0 0;color:#666;font-size:13.5px">Per modifiche rispondi a questa email o chiamaci allo ' +
        '<a href="tel:' + RESTAURANT_TEL.replace(/\s/g, '') + '" style="color:#ffb200">' + RESTAURANT_TEL + '</a>.</p>' +
    '</div>';

  var testo = righe.join('\n\n') + '\n\n— Riepilogo —\n' +
    (data ? 'Data: ' + data + '\n' : '') + (p.ora ? 'Orario: ' + p.ora + '\n' : '') +
    (p.persone ? 'Coperti: ' + p.persone + '\n' : '') + (nomeCompleto ? 'A nome di: ' + nomeCompleto + '\n' : '') +
    (p.telefono ? 'Telefono: ' + p.telefono + '\n' : '') +
    '\nPer modifiche rispondi a questa email o chiamaci allo ' + RESTAURANT_TEL + '.';

  MailApp.sendEmail(dest, oggetto, testo, opzioniInvio_({ htmlBody: html, replyTo: RESTAURANT_EMAIL }));
  return true;
}

// ── DIAGNOSTICA (da eseguire a mano dall'editor) ──────────────────────────
function verificaFoglio() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Prenotazioni');
  var msg =
    'Foglio raggiunto: ' + ss.getName() +
    '\nScheda "Prenotazioni": ' + (sheet ? 'trovata, ' + sheet.getLastRow() + ' righe' : 'ASSENTE (verrà creata alla prima prenotazione)') +
    '\nEmail in uscita da: ' + Session.getEffectiveUser().getEmail() +
    '\nNotifica al ristorante: ' + RESTAURANT_EMAIL +
    '\nInvii disponibili oggi: ' + MailApp.getRemainingDailyQuota();
  Logger.log(msg);
  return msg;
}

function chiMandaLeMail() {
  var account = Session.getEffectiveUser().getEmail();
  var alias; try { alias = GmailApp.getAliases(); } catch (e) { alias = ['(permesso Gmail non concesso)']; }
  var msg =
    'Account dello script:   ' + account +
    '\nAlias verificati:       ' + (alias.length ? alias.join(', ') : '(nessuno)') +
    '\nMITTENTE_ALIAS:         ' + (MITTENTE_ALIAS || '(vuoto)') +
    '\nLe email partiranno da: ' + (aliasUsabile_(MITTENTE_ALIAS) ? MITTENTE_ALIAS : account) +
    '\nConferma al cliente:    ' + (aliasUsabile_(MITTENTE_ALIAS) ? 'ATTIVA' : 'BLOCCATA (parte solo la notifica al ristorante)') +
    '\nInvii disponibili oggi: ' + MailApp.getRemainingDailyQuota();
  Logger.log(msg);
  return msg;
}

function doGet() {
  return ContentService.createTextOutput("L'Italiana — endpoint prenotazioni attivo · v1");
}

// DAMIR HOTEL — Google Apps Script form endpoint
// ============================================================
// SETUP (one-time, ~5 minutes):
//
// 1. Open https://sheets.google.com, create a new spreadsheet.
//    Rename it "Guest Inquiries".
//    In row 1 add these headers (one per cell, A through J):
//      Timestamp | Name | Email | Phone | Guests |
//      Check-in  | Check-out | Room | Message | Source
//
// 2. In the spreadsheet: Extensions → Apps Script
//    Delete all default code and paste this entire file.
//
// 3. Deploy → New deployment → Web app
//      Execute as:       Me
//      Who has access:   Anyone
//    Click Deploy → copy the Web app URL.
//    IMPORTANT: "Who has access" must be "Anyone" (not "Anyone in organisation")
//    or the form will silently redirect guests to a Google login page.
//
// 4. Open Design/assets/hotel.js and set:
//      formEndpoint: 'PASTE_YOUR_WEB_APP_URL_HERE',
//
// 5. Set Script Properties (Project Settings → Script Properties):
//      ALLOWED_ORIGINS  →  your-domain.com,www.your-domain.com
//    Without this, requests are accepted from any origin (dev mode).
//    Comma-separated, no spaces, no protocol (e.g. "damirhotel.uz,www.damirhotel.uz").
//
// 6. (Optional) visit the URL in a browser — you should see:
//    {"ok":true,"service":"Damir Hotel form endpoint"}
// ============================================================

// ── rate-limit config ──────────────────────────────────────
var RATE_LIMIT_MAX    = 20;    // max accepted submissions per window (globally)
var RATE_LIMIT_WINDOW = 3600;  // window size in seconds (1 hour)
var MAX_FIELD_LEN     = 1000;  // max characters per text field

function doPost(e) {
  try {
    // 1. Rate limit — atomic via LockService, counted via CacheService
    if (!_checkRateLimit()) {
      return _jsonOut({ ok: false, error: 'rate_limited' });
    }

    var data = JSON.parse(e.postData.contents);

    // 2. Honeypot: bots fill the hidden website field; silently succeed
    if (data.website) return _jsonOut({ ok: true });

    // 3. Origin check — client sends window.location.hostname as _origin.
    //    If ALLOWED_ORIGINS is set in Script Properties, the value must match.
    //    If ALLOWED_ORIGINS is not set, any origin is accepted (dev / migration mode).
    if (!_originAllowed(data._origin)) {
      Logger.log('Blocked request from origin: ' + data._origin);
      return _jsonOut({ ok: false, error: 'forbidden' });
    }

    // 4. Input validation
    var validationError = _validate(data);
    if (validationError) return _jsonOut({ ok: false, error: validationError });

    // --- append row to sheet ---
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([
      new Date(),
      data.name     || '',
      data.email    || '',
      data.phone    || '',
      data.guests   || '',
      data.checkin  || '',
      data.checkout || '',
      data.room     || '',
      data.message  || '',
      data.source   || 'contact-form'
    ]);

    // --- email the hotel ---
    var to      = 'damirhotel@gmail.com';
    var subject = 'New inquiry — ' + (data.name || 'guest') +
                  (data.checkin  ? ', ' + data.checkin  : '') +
                  (data.checkout ? ' – ' + data.checkout : '');

    var body = [
      'New reservation inquiry via Damir Hotel website',
      '',
      'Name:       ' + (data.name     || '—'),
      'Email:      ' + (data.email    || '—'),
      'Phone:      ' + (data.phone    || '—'),
      'Guests:     ' + (data.guests   || '—'),
      'Check-in:   ' + (data.checkin  || '—'),
      'Check-out:  ' + (data.checkout || '—'),
      'Room:       ' + (data.room     || 'No preference'),
      '',
      'Message:',
      data.message || '—',
      '',
      '──────────────────────────────',
      'Reply to this email to respond directly to the guest.'
    ].join('\n');

    MailApp.sendEmail({
      to:      to,
      subject: subject,
      body:    body,
      replyTo: data.email || to
    });

    // Confirmation to the guest (only if they gave an email)
    if (data.email) {
      var guestBody = [
        'Dear ' + (data.name || 'Guest') + ',',
        '',
        'Thank you for reaching out to Damir Hotel. We have received your reservation inquiry and our team will reply within a few hours to confirm availability.',
        '',
        'Your request summary:',
        '  Check-in:   ' + (data.checkin  || '—'),
        '  Check-out:  ' + (data.checkout || '—'),
        '  Guests:     ' + (data.guests   || '—'),
        '  Room:       ' + (data.room     || 'No preference'),
        '',
        'If you have any urgent questions, you can reach us directly:',
        '  Phone: +998 55 106 07 07',
        '  Email: damirhotel@gmail.com',
        '',
        'Warm regards,',
        'Damir Hotel · Рауаж 4, 230100 Nukus'
      ].join('\n');

      MailApp.sendEmail({
        to:      data.email,
        subject: 'We received your inquiry — Damir Hotel',
        body:    guestBody,
        replyTo: to
      });
    }

    return _jsonOut({ ok: true });

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return _jsonOut({ ok: false, error: 'server_error' });
  }
}

// Health check — GET request confirms the endpoint is live
function doGet() {
  return _jsonOut({ ok: true, service: 'Damir Hotel form endpoint' });
}

// ── helpers ────────────────────────────────────────────────

function _jsonOut(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// Atomic per-hour counter using CacheService + LockService.
// Returns true if the request is within the limit, false if it should be rejected.
function _checkRateLimit() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(3000); } catch (e) { return false; }
  try {
    var cache = CacheService.getScriptCache();
    var key   = 'rl_' + Math.floor(Date.now() / 1000 / RATE_LIMIT_WINDOW);
    var count = parseInt(cache.get(key) || '0', 10);
    if (count >= RATE_LIMIT_MAX) return false;
    cache.put(key, String(count + 1), RATE_LIMIT_WINDOW + 60);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// Checks data._origin (window.location.hostname sent by the client) against
// the ALLOWED_ORIGINS Script Property. If the property is not set, all origins
// are accepted so existing deployments keep working during rollout.
function _originAllowed(origin) {
  var raw = PropertiesService.getScriptProperties().getProperty('ALLOWED_ORIGINS');
  if (!raw) return true; // property not configured — open mode (dev / rollout)
  if (!origin) return false;
  var allowed = raw.split(',').map(function (s) { return s.trim().toLowerCase(); });
  return allowed.indexOf(origin.toLowerCase()) !== -1;
}

// Basic field validation. Returns an error string or null when valid.
function _validate(data) {
  if (!data.name  || !String(data.name).trim())  return 'name_required';
  if (!data.email || !String(data.email).trim())  return 'email_required';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) return 'email_invalid';

  var textFields = ['name', 'email', 'phone', 'message', 'room', 'source'];
  for (var i = 0; i < textFields.length; i++) {
    var v = data[textFields[i]];
    if (v && String(v).length > MAX_FIELD_LEN) return 'field_too_long';
  }
  return null;
}

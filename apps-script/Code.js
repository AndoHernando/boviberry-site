// ============================================================
// BOVIBERRY — GOOGLE APPS SCRIPT BACKEND (v3)
// Reviews · Commissions · Waitlist · Moodboards · Admin auth
// ============================================================
//
// SETUP — run once from the Google account that owns the studio
//
// 1. Paste ALL of this code, replacing everything in the editor
// 2. Check SPREADSHEET_ID below matches your sheet
//      (the long string in the sheet URL between /d/ and /edit)
// 3. Set the admin password:
//      - put it in ADMIN_PASSWORD_TO_SET below, save
//      - run "setAdminPassword" from the function dropdown
//      - clear it back to '' and save again
// 4. Deploy → New deployment → Web app
//      Execute as: Me    ·    Who has access: Anyone
// 5. Copy the /exec URL into the four site files that use it:
//      index.html · reviews.html · moodboard.html · admin.html
//
// ============================================================

// ── CONFIG ──
// Where new submissions get announced (your own inbox).
var NOTIFY_EMAIL = 'Boviberrystudios@gmail.com';

// Identity on the auto-reply clients receive.
//
// SEND_AS should stay BLANK when this script is owned by the BoviBerry
// account — mail already goes out from the owner, so nothing extra is needed.
// Only fill it in if a DIFFERENT account owns the script, and only after this
// address is verified there under Gmail → Settings → Accounts and Import →
// "Send mail as". Run listMyAliases() any time to check the current state.
var SENDER_NAME = 'BoviBerry';
var REPLY_TO    = 'Boviberrystudios@gmail.com';
var SEND_AS     = '';
var SITE_URL    = 'https://boviberry.com';

// ── SPAM PROTECTION ──
// Cloudflare Turnstile (free, invisible CAPTCHA). Leave blank to disable —
// everything still works, just without the challenge.
// Get keys at: Cloudflare dashboard → Turnstile → Add site
var TURNSTILE_SECRET = '';

// Submissions scoring at or above this are flagged for review (not rejected).
var SPAM_THRESHOLD = 40;

// The spreadsheet everything is stored in.
var SPREADSHEET_ID = '1tpezSsLBbe3Ytx7Es4KO2iiTQnyQzpAyP97WP3rk6zc';

// Only used by setAdminPassword(). Leave blank except when changing it.
var ADMIN_PASSWORD_TO_SET = '';
// ============================================================


// ── CACHING (avoids reopening the spreadsheet on every call) ──
var _ss = null;
var _sheetCache = {};
var _settingsCache = null;

function getSpreadsheet() {
  if (_ss) return _ss;
  _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets(_ss);
  return _ss;
}

function ensureSheets(ss) {
  var defs = {
    'Reviews':          ['id','name','service','rating','review','date','approved','featured'],
    'Commissions':      ['id','name','email','phone','socialPlatform','socialHandle','service','budget','vision','references','timeline','referral','date','status','flag','flagReason'],
    'Waitlist':         ['id','name','email','service','date'],
    'Settings':         ['key','value'],
    'Moodboards':       ['id','name','email','links','details','colors','date','flag','flagReason'],
    'MoodboardImages':  ['id','moodboardId','imageData'],
    'Catalog':          ['id','label','price','salePrice','onSale','active','note',
                         'kind','formGroup','turnaround','minDays','stripeLink','sortOrder']
  };
  for (var name in defs) {
    var sh = ss.getSheetByName(name);
    if (sh) ensureColumns(sh, defs[name]);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(defs[name]);
      sh.getRange(1, 1, 1, defs[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
      if (name === 'Settings') sh.appendRow(['waitlist', 'false']);
      if (name === 'Catalog')  seedCatalog(sh);
    }
  }
}



/**
 * Adds any newly-introduced columns to a sheet that already exists, so the
 * schema can grow without a manual migration. Existing rows simply read back
 * empty for the new fields.
 */
function ensureColumns(sheet, headers) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) { sheet.appendRow(headers); return; }
  var current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = [];
  for (var i = 0; i < headers.length; i++) {
    if (current.indexOf(headers[i]) === -1) missing.push(headers[i]);
  }
  if (!missing.length) return;
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
}

/**
 * Verifies a Cloudflare Turnstile token.
 * Fails OPEN on error — losing a real commission is worse than letting one
 * spam entry through, and it still gets scored below either way.
 */
function verifyTurnstile(token) {
  if (!TURNSTILE_SECRET) return true;          // not configured
  if (!token) return false;                    // configured but nothing sent
  try {
    var res = UrlFetchApp.fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
      method: 'post',
      payload: { secret: TURNSTILE_SECRET, response: token },
      muteHttpExceptions: true
    });
    return !!JSON.parse(res.getContentText()).success;
  } catch (e) {
    return true;
  }
}

/**
 * Scores a submission for spam signals. Returns { score, reasons }.
 * Nothing is ever rejected on this — it only flags for her review.
 */
function spamScore(p) {
  var score = 0, reasons = [];

  // Hidden field only a bot would fill
  if (p.hp) { score += 100; reasons.push('Filled a hidden field'); }

  // Humans don't complete a ten-field form in three seconds
  if (p.elapsed !== undefined && p.elapsed !== '' && Number(p.elapsed) < 3000) {
    score += 40; reasons.push('Submitted in ' + (Number(p.elapsed) / 1000).toFixed(1) + 's');
  }

  // Reserved fictional phone prefix
  var digits = String(p.phone || '').replace(/\D/g, '');
  if (/^1?555/.test(digits)) { score += 35; reasons.push('555 test phone number'); }

  // Gmail ignores dots, so scattered dots generate "unique" throwaway addresses
  var em = String(p.email || '').toLowerCase();
  var g = em.match(/^([^@]+)@gmail\.com$/);
  if (g && (g[1].match(/\./g) || []).length >= 3) {
    score += 30; reasons.push('Unusual dot pattern in Gmail address');
  }

  // A brief with no actual words in it
  var v = String(p.vision || p.details || '').trim();
  if (v && /^[\d\s()+.\-]+$/.test(v)) { score += 40; reasons.push('Brief contains no words'); }
  else if (v && v.length < 15) { score += 15; reasons.push('Brief unusually short'); }

  // Links where links don't belong
  if (/https?:\/\//i.test(String(p.name || ''))) { score += 50; reasons.push('Link in the name field'); }

  // Picked a subscription but claims a budget below its price
  var svc = String(p.service || '').toLowerCase();
  var bud = String(p.budget || '').toLowerCase();
  if (svc.indexOf('tier') > -1 && (bud.indexOf('under $50') > -1)) {
    score += 20; reasons.push('Budget contradicts the selected service');
  }

  return { score: score, reasons: reasons };
}

function getSheet(name) {
  if (_sheetCache[name]) return _sheetCache[name];
  _sheetCache[name] = getSpreadsheet().getSheetByName(name);
  return _sheetCache[name];
}

function rowsToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = String(data[i][j] || '');
    out.push(obj);
  }
  return out;
}

function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}



/* ── CATALOG ──
   Prices, sale status and visibility for everything sold on the site.
   Seeded with the values that are hardcoded in the HTML, so the sheet and
   the site agree from day one. Editing a row here changes the live site;
   the HTML keeps its own copy as a fallback if this ever can't be reached. */
var CATALOG_DEFAULTS = [
  // ── Botanical Harvest — streaming (YouTube / Twitch) ──
  ['stream-blueberry',  'Botanical Harvest — Blueberry',  '$70',  '', 'false', 'true', '/month'],
  ['stream-raspberry',  'Botanical Harvest — Raspberry',  '$85',  '', 'false', 'true', '/month'],
  ['stream-cherry',     'Botanical Harvest — Cherry',     '$100',  '', 'false', 'true', '/month'],

  // ── Farmers Market — social media (Instagram / TikTok / Pinterest) ──
  ['social-blueberry',  'Farmers Market — Blueberry',     '$80',  '', 'false', 'true', '/month'],
  ['social-raspberry',  'Farmers Market — Raspberry',     '$95',  '', 'false', 'true', '/month'],
  ['social-cherry',     'Farmers Market — Cherry',        '$110',  '', 'false', 'true', '/month'],

  // ── Logo design ──
  ['logo-wordmark',     'Logo — Word mark',               '$199', '', 'false', 'true', ''],
  ['logo-letterforms',  'Logo — Letterforms',             '$249', '', 'false', 'true', ''],
  ['logo-pictorial',    'Logo — Pictorial',               '$299', '', 'false', 'true', ''],
  ['logo-abstract',     'Logo — Abstract',                '$299', '', 'false', 'true', ''],
  ['logo-emblem',       'Logo — Emblem',                  '$399', '', 'false', 'true', ''],
  ['logo-mascot',       'Logo — Mascot',                  '$499', '', 'false', 'true', ''],

  // ── Individual projects ──
  ['punch-printable',   'Punch card — Printable',         '$29',  '', 'false', 'true', ''],
  ['punch-digital',     'Punch card — Digital wallet',    '$69',  '', 'false', 'true', ''],

  // ── Add-ons ──
  ['addon-moodboard',   'Add-on — Tailored mood board',   '$25',  '', 'false', 'true', ''],
  ['addon-revision',    'Add-on — Full revision round',   '$40',  '', 'false', 'true', ''],
  ['addon-rush',        'Add-on — Rush delivery',         '+25%', '', 'false', 'true', ''],

  // ── Kits ──
  ['kit-starter',       'Kit — Starter',                  '',     '', 'false', 'true', 'quoted individually'],
  ['kit-creator-pro',   'Kit — Creator Pro',              '',     '', 'false', 'true', 'quoted individually'],
  ['kit-brand-identity','Kit — Brand Identity',           '',     '', 'false', 'true', 'quoted individually']
];

function seedCatalog(sheet) {
  sheet.getRange(2, 1, CATALOG_DEFAULTS.length, 7).setValues(CATALOG_DEFAULTS);
}

/** Adds any newly-introduced default rows without touching her edits. */

/**
 * Metadata that drives the commission form. Keyed by catalog id.
 *   kind      — subscription | fixed | quote  (decides how the form behaves)
 *   formGroup — the heading it sits under in the service dropdown
 *   turnaround/minDays — powers the "that's sooner than usual" hint
 *   sortOrder — display order within its group
 */
var CATALOG_META = {
  // id: [kind, formGroup, turnaround, minDays, stripeLink, sortOrder]
  'stream-blueberry':   ['subscription','Streaming — YouTube / Twitch','', 0, '', 10],
  'stream-raspberry':   ['subscription','Streaming — YouTube / Twitch','', 0, '', 11],
  'stream-cherry':      ['subscription','Streaming — YouTube / Twitch','', 0, '', 12],
  'social-blueberry':   ['subscription','Social — Instagram / TikTok / Pinterest','', 0, '', 20],
  'social-raspberry':   ['subscription','Social — Instagram / TikTok / Pinterest','', 0, '', 21],
  'social-cherry':      ['subscription','Social — Instagram / TikTok / Pinterest','', 0, '', 22],
  'logo-wordmark':      ['fixed','Logo design','7–14 business days', 7, '', 30],
  'logo-letterforms':   ['fixed','Logo design','7–14 business days', 7, '', 31],
  'logo-pictorial':     ['fixed','Logo design','7–14 business days', 7, '', 32],
  'logo-abstract':      ['fixed','Logo design','7–14 business days', 7, '', 33],
  'logo-emblem':        ['fixed','Logo design','7–14 business days', 7, '', 34],
  'logo-mascot':        ['fixed','Logo design','10–14 business days', 10, '', 35],
  'punch-printable':    ['fixed','Individual projects','3–5 business days', 3, '', 40],
  'punch-digital':      ['fixed','Individual projects','5–7 business days', 5, '', 41],
  'addon-moodboard':    ['fixed','Add-ons','', 0, '', 50],
  'addon-revision':     ['fixed','Add-ons','', 0, '', 51],
  'addon-rush':         ['quote','', '', 0, '', 52],
  'kit-starter':        ['quote','Brand & creator kits','5–7 business days', 5, '', 60],
  'kit-creator-pro':    ['quote','Brand & creator kits','7–10 business days', 7, '', 61],
  'kit-brand-identity': ['quote','Brand & creator kits','10–14 business days', 10, '', 62]
};

/**
 * Fills in form metadata for existing catalog rows without touching prices,
 * sale flags or visibility. Safe to re-run — only writes empty cells.
 * Run once after deploying the wider schema.
 */
function syncCatalogMeta() {
  var sheet = getSheet('Catalog');
  var data = sheet.getDataRange().getValues();
  var hdr = data[0];
  var cols = {};
  ['kind','formGroup','turnaround','minDays','stripeLink','sortOrder'].forEach(function (k) {
    cols[k] = hdr.indexOf(k) + 1;
  });
  if (!cols.kind) { Logger.log('Columns missing — deploy a new version first.'); return; }

  var order = ['kind','formGroup','turnaround','minDays','stripeLink','sortOrder'];
  var filled = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0]);
    var meta = CATALOG_META[id];
    if (!meta) continue;
    for (var c = 0; c < order.length; c++) {
      var col = cols[order[c]];
      if (!col) continue;
      if (String(data[i][col - 1] || '') === '') {
        sheet.getRange(i + 1, col).setValue(meta[c]);
        filled++;
      }
    }
  }
  Logger.log(filled ? 'Filled ' + filled + ' empty metadata cell(s).' : 'Catalog metadata already complete.');
}

function syncCatalogDefaults() {
  var sheet = getSheet('Catalog');
  var existing = {};
  rowsToObjects(sheet).forEach(function (r) { existing[r.id] = true; });
  var added = 0;
  for (var i = 0; i < CATALOG_DEFAULTS.length; i++) {
    if (!existing[CATALOG_DEFAULTS[i][0]]) { sheet.appendRow(CATALOG_DEFAULTS[i]); added++; }
  }
  Logger.log(added ? 'Added ' + added + ' new catalog row(s).' : 'Catalog already up to date.');
}


// ── SETTINGS ──
function loadSettings() {
  if (_settingsCache) return _settingsCache;
  var data = getSheet('Settings').getDataRange().getValues();
  _settingsCache = {};
  for (var i = 1; i < data.length; i++) _settingsCache[data[i][0]] = String(data[i][1]);
  return _settingsCache;
}

function getSetting(key) {
  var s = loadSettings();
  return s[key] !== undefined ? s[key] : '';
}

function setSetting(key, value) {
  var sheet = getSheet('Settings');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      if (_settingsCache) _settingsCache[key] = String(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
  if (_settingsCache) _settingsCache[key] = String(value);
}


// ── AUTH ──
function sha256Hex(str) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    hex += ('0' + (bytes[i] & 0xFF).toString(16)).slice(-2);
  }
  return hex;
}

/**
 * RUN THIS MANUALLY from the editor to set or change the admin password.
 * Set ADMIN_PASSWORD_TO_SET at the top first, run this, then blank it out again.
 */
function setAdminPassword() {
  if (!ADMIN_PASSWORD_TO_SET) {
    Logger.log('Set ADMIN_PASSWORD_TO_SET at the top of the file first, then run this again.');
    return;
  }
  setSetting('adminPassHash', sha256Hex(ADMIN_PASSWORD_TO_SET));
  setSetting('adminTokens', '[]');
  Logger.log('Admin password saved. Now blank out ADMIN_PASSWORD_TO_SET and save the file.');
}

function adminLogin(hash) {
  var stored = getSetting('adminPassHash');
  if (!stored) return { ok: false, error: 'No admin password set. Run setAdminPassword() in Apps Script.' };
  if (!hash || hash !== stored) return { ok: false, error: 'Incorrect password.' };

  var tokens = [];
  try { tokens = JSON.parse(getSetting('adminTokens') || '[]'); } catch (e) { tokens = []; }
  var now = Date.now();
  tokens = tokens.filter(function (t) { return t.exp > now; });   // drop expired
  var token = Utilities.getUuid().replace(/-/g, '');
  tokens.push({ t: token, exp: now + 12 * 60 * 60 * 1000 });      // 12 hour session
  if (tokens.length > 5) tokens = tokens.slice(-5);               // cap devices
  setSetting('adminTokens', JSON.stringify(tokens));
  return { ok: true, token: token };
}

function validToken(token) {
  if (!token) return false;
  var tokens = [];
  try { tokens = JSON.parse(getSetting('adminTokens') || '[]'); } catch (e) { return false; }
  var now = Date.now();
  for (var i = 0; i < tokens.length; i++) {
    if (tokens[i].t === token && tokens[i].exp > now) return true;
  }
  return false;
}

function adminLogout(token) {
  var tokens = [];
  try { tokens = JSON.parse(getSetting('adminTokens') || '[]'); } catch (e) { tokens = []; }
  tokens = tokens.filter(function (t) { return t.t !== token; });
  setSetting('adminTokens', JSON.stringify(tokens));
}


// ── EMAIL ──
function notify(subject, body) {
  try {
    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, body: body });
  } catch (e) { /* never let email failure break a submission */ }
}

/* ============================================================
   STUDIO NOTIFICATION EMAILS
   Sent to her when something comes in. Built to be scannable on a
   phone: the facts she needs to quote a job sit at the top, the
   long-form brief sits below, and the actions are one tap away.
   ============================================================ */

function adminShell(eyebrow, title, rowsHtml, extraHtml, ctaHtml) {
  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
  '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body style="margin:0;padding:0;background-color:#f5efe3;">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
         'style="background-color:#f5efe3;padding:24px 12px;"><tr><td align="center">' +

    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
           'style="max-width:600px;background-color:#ffffff;border:1px solid #e0d6c4;">' +

      '<tr><td style="height:4px;background-color:#2c7449;font-size:0;line-height:0;">&nbsp;</td></tr>' +

      '<tr><td style="padding:26px 32px 0 32px;">' +
        '<div style="font-family:Helvetica,Arial,sans-serif;font-size:10px;letter-spacing:3px;' +
                    'text-transform:uppercase;color:#9bd16b;font-weight:bold;">' + eyebrow + '</div>' +
        '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:25px;color:#2c2419;' +
                    'padding-top:6px;line-height:1.25;">' + title + '</div>' +
      '</td></tr>' +

      '<tr><td style="padding:20px 32px 0 32px;">' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
               'style="font-family:Helvetica,Arial,sans-serif;font-size:14px;">' +
          rowsHtml +
        '</table>' +
      '</td></tr>' +

      (extraHtml || '') +

      (ctaHtml ? '<tr><td style="padding:24px 32px 0 32px;">' + ctaHtml + '</td></tr>' : '') +

      '<tr><td style="padding:26px 32px 28px 32px;">' +
        '<div style="height:1px;background-color:#e0d6c4;font-size:0;line-height:0;margin-bottom:14px;">&nbsp;</div>' +
        '<div style="font-family:Helvetica,Arial,sans-serif;font-size:11px;color:#b0a18a;">' +
          'BoviBerry Studio · sent automatically from ' +
          '<a href="' + SITE_URL + '" style="color:#9a8b74;text-decoration:none;">boviberry.com</a>' +
        '</div>' +
      '</td></tr>' +

    '</table>' +
  '</td></tr></table></body></html>';
}

/** One label/value line in the summary table. */
function row(label, value, opts) {
  if (value === undefined || value === null || value === '') return '';
  var v = String(value);
  if (!opts) opts = {};
  var valStyle = 'padding:7px 0;color:#2c2419;line-height:1.5;' +
                 (opts.strong ? 'font-weight:bold;' : '');
  return '<tr>' +
    '<td width="108" valign="top" style="padding:7px 12px 7px 0;color:#9a8b74;' +
        'font-size:11px;letter-spacing:1.2px;text-transform:uppercase;white-space:nowrap;">' +
      escHtml(label) + '</td>' +
    '<td valign="top" style="' + valStyle + '">' + (opts.raw ? v : escHtml(v)) + '</td>' +
  '</tr>';
}

/** A block of long-form text — the brief, references, review body. */
function panel(label, text, accent) {
  if (!text) return '';
  return '<tr><td style="padding:20px 32px 0 32px;">' +
    '<div style="font-family:Helvetica,Arial,sans-serif;font-size:10px;letter-spacing:2px;' +
                'text-transform:uppercase;color:#9a8b74;padding-bottom:8px;">' + escHtml(label) + '</div>' +
    '<div style="font-family:Helvetica,Arial,sans-serif;font-size:15px;line-height:24px;' +
                'color:#3d3226;background-color:#f7f3ea;border-left:3px solid ' + (accent || '#9bd16b') + ';' +
                'padding:14px 16px;white-space:pre-wrap;word-break:break-word;">' +
      escHtml(text) +
    '</div></td></tr>';
}

/** Buttons. Table-based so Outlook renders them. */
function ctaButtons(primaryLabel, primaryUrl, secondaryLabel, secondaryUrl) {
  var btn = function (label, url, bg, color, border) {
    return '<td style="padding-right:10px;">' +
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>' +
      '<td style="background-color:' + bg + ';border:1px solid ' + border + ';">' +
      '<a href="' + url + '" style="display:inline-block;padding:11px 20px;' +
         'font-family:Helvetica,Arial,sans-serif;font-size:12px;letter-spacing:1.5px;' +
         'text-transform:uppercase;color:' + color + ';text-decoration:none;">' + label + '</a>' +
      '</td></tr></table></td>';
  };
  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>' +
    btn(primaryLabel, primaryUrl, '#2c7449', '#ffffff', '#2c7449') +
    (secondaryUrl ? btn(secondaryLabel, secondaryUrl, '#ffffff', '#2c7449', '#c9bda6') : '') +
  '</tr></table>';
}

function adminSend(subject, textBody, htmlBody) {
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: subject,
      body: textBody,
      htmlBody: htmlBody,
      name: 'BoviBerry Studio'
    });
  } catch (e) { /* submission already saved — never block on email */ }
}

function telHref(phone) {
  var d = String(phone || '').replace(/[^\d+]/g, '');
  return d ? 'tel:' + d : '';
}


function flagBanner(flag) {
  if (!flag) return '';
  return '<tr><td style="padding:20px 32px 0 32px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
           'style="background-color:#fdecec;border-left:3px solid #c94c4c;">' +
    '<tr><td style="padding:13px 16px;font-family:Helvetica,Arial,sans-serif;font-size:14px;' +
                   'line-height:21px;color:#8a3030;">' +
      '<strong>Possible spam.</strong> ' + escHtml(flag.reasons.join(' · ')) +
      '<br><span style="font-size:12px;color:#a86a6a;">No auto-reply was sent. ' +
      'Mark it as legitimate in the studio panel if this is a real enquiry.</span>' +
    '</td></tr></table></td></tr>';
}

// ── NEW COMMISSION ──
function notifyCommission(p, unclaimed, flag) {
  var name = p.name || 'Someone';
  var hasPhone = p.phone && p.phone !== 'Not provided';
  var hasSocial = p.socialPlatform && p.socialPlatform !== '—';

  var rows =
    row('Service', p.service, { strong: true }) +
    row('Budget', p.budget) +
    row('Timeline', p.timeline) +
    row('Found via', p.referral) +
    '<tr><td colspan="2" style="padding:10px 0;"><div style="height:1px;background-color:#eee6d8;font-size:0;">&nbsp;</div></td></tr>' +
    row('Email', '<a href="mailto:' + escHtml(p.email) + '" style="color:#2c7449;text-decoration:none;">' +
                 escHtml(p.email) + '</a>', { raw: true }) +
    (hasPhone ? row('Phone', '<a href="' + telHref(p.phone) + '" style="color:#2c7449;text-decoration:none;">' +
                 escHtml(p.phone) + '</a>', { raw: true }) : '') +
    (hasSocial ? row(p.socialPlatform, p.socialHandle) : '') +
    row('Received', p.date);

  var warn = (unclaimed >= 5)
    ? '<tr><td style="padding:20px 32px 0 32px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
             'style="background-color:#fdf3e0;border-left:3px solid #ffaa1b;">' +
      '<tr><td style="padding:13px 16px;font-family:Helvetica,Arial,sans-serif;font-size:14px;' +
                     'line-height:21px;color:#7a5a20;">' +
        '<strong>' + unclaimed + ' unclaimed commissions.</strong> Consider switching the waitlist on ' +
        'while you work through the queue.' +
      '</td></tr></table></td></tr>'
    : '';

  var extra = flagBanner(flag) + warn +
    panel('The brief', p.vision) +
    (p.references && p.references !== 'None provided' ? panel('Reference links', p.references, '#c9bda6') : '');

  var cta = ctaButtons('Reply to ' + escHtml(String(name).split(' ')[0]),
                       'mailto:' + encodeURIComponent(p.email) + '?subject=' +
                         encodeURIComponent('Re: your BoviBerry commission'),
                       'Open studio', SITE_URL + '/admin.html');

  var text =
    'NEW COMMISSION — ' + name + '\n\n' +
    'Service: ' + (p.service || '—') + '\n' +
    'Budget: ' + (p.budget || '—') + '\n' +
    'Timeline: ' + (p.timeline || '—') + '\n' +
    'Found via: ' + (p.referral || '—') + '\n\n' +
    'Email: ' + (p.email || '—') + '\n' +
    'Phone: ' + (p.phone || 'Not provided') + '\n' +
    (hasSocial ? p.socialPlatform + ': ' + (p.socialHandle || '—') + '\n' : '') +
    '\nBRIEF\n' + (p.vision || '—') + '\n\n' +
    'REFERENCES\n' + (p.references || 'None provided') + '\n\n' +
    (unclaimed >= 5 ? 'NOTE: ' + unclaimed + ' unclaimed commissions.\n\n' : '') +
    SITE_URL + '/admin.html';

  adminSend((flag ? '[Possible spam] ' : '') + 'New commission — ' + name + ' · ' + (p.service || 'Custom'),
            text, adminShell('New commission', escHtml(name), rows, extra, cta));
}

// ── NEW MOOD BOARD ──
function notifyMoodboard(p, flag) {
  var name = p.name || 'Someone';
  var imgCount = (p.images || []).length;

  var swatches = '';
  if (p.colors && p.colors !== 'None') {
    var list = String(p.colors).split(',');
    swatches = '<tr><td style="padding:20px 32px 0 32px;">' +
      '<div style="font-family:Helvetica,Arial,sans-serif;font-size:10px;letter-spacing:2px;' +
                  'text-transform:uppercase;color:#9a8b74;padding-bottom:10px;">Colour palette</div>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>';
    for (var i = 0; i < list.length; i++) {
      var hex = list[i].trim();
      if (!/^#[0-9a-fA-F]{3,8}$/.test(hex)) continue;
      swatches += '<td style="padding-right:8px;" align="center">' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>' +
        '<td bgcolor="' + hex + '" width="52" height="52" ' +
            'style="background-color:' + hex + ';border:1px solid #d8cdb8;font-size:0;line-height:0;">&nbsp;</td>' +
        '</tr><tr><td align="center" style="font-family:Helvetica,Arial,sans-serif;font-size:10px;' +
            'color:#9a8b74;padding-top:5px;">' + escHtml(hex) + '</td></tr></table></td>';
    }
    swatches += '</tr></table></td></tr>';
  }

  var rows =
    row('From', name, { strong: true }) +
    row('Email', '<a href="mailto:' + escHtml(p.email) + '" style="color:#2c7449;text-decoration:none;">' +
                 escHtml(p.email) + '</a>', { raw: true }) +
    row('Images', imgCount ? imgCount + ' uploaded' : 'None') +
    row('Received', p.date);

  var links = '';
  if (p.links && p.links !== 'None') {
    var ls = String(p.links).split('\n').filter(function (l) { return l.trim(); });
    links = '<tr><td style="padding:20px 32px 0 32px;">' +
      '<div style="font-family:Helvetica,Arial,sans-serif;font-size:10px;letter-spacing:2px;' +
                  'text-transform:uppercase;color:#9a8b74;padding-bottom:8px;">Inspiration links</div>';
    for (var j = 0; j < ls.length; j++) {
      links += '<div style="padding:5px 0;border-bottom:1px solid #eee6d8;">' +
        '<a href="' + escHtml(ls[j].trim()) + '" style="font-family:Helvetica,Arial,sans-serif;' +
           'font-size:13px;color:#2c7449;text-decoration:none;word-break:break-all;">' +
           escHtml(ls[j].trim()) + '</a></div>';
    }
    links += '</td></tr>';
  }

  var extra = flagBanner(flag) +
    (p.details && p.details !== 'None' ? panel('Their vision', p.details) : '') +
    links + swatches +
    (imgCount ? '<tr><td style="padding:18px 32px 0 32px;' +
      'font-family:Helvetica,Arial,sans-serif;font-size:13px;color:#9a8b74;line-height:20px;">' +
      'The ' + imgCount + ' reference image' + (imgCount === 1 ? '' : 's') +
      ' are in the studio panel under Moodboards.</td></tr>' : '');

  var cta = ctaButtons('Open studio', SITE_URL + '/admin.html',
                       'Reply', 'mailto:' + encodeURIComponent(p.email));

  var text =
    'NEW MOOD BOARD — ' + name + '\n\n' +
    'Email: ' + (p.email || '—') + '\n' +
    'Images: ' + imgCount + '\n' +
    'Colours: ' + (p.colors || 'None') + '\n\n' +
    'VISION\n' + (p.details || 'None') + '\n\n' +
    'LINKS\n' + (p.links || 'None') + '\n\n' +
    SITE_URL + '/admin.html';

  adminSend((flag ? '[Possible spam] ' : '') + 'New mood board — ' + name, text,
            adminShell('New mood board', escHtml(name), rows, extra, cta));
}

// ── NEW REVIEW ──
function notifyReview(p) {
  var name = p.name || 'Someone';
  var n = parseInt(p.rating, 10) || 0;
  var starHtml = '<span style="color:#ffaa1b;font-size:17px;letter-spacing:2px;">' +
                 new Array(n + 1).join('★') +
                 '<span style="color:#ddd3c0;">' + new Array(6 - n).join('★') + '</span></span>';

  var rows =
    row('From', name, { strong: true }) +
    row('Rating', starHtml + ' <span style="color:#9a8b74;font-size:13px;">' + n + '/5</span>', { raw: true }) +
    row('Service', p.service) +
    row('Received', p.date);

  var cta = ctaButtons('Approve in studio', SITE_URL + '/admin.html');

  var text = 'NEW REVIEW — ' + name + '\n\nRating: ' + n + '/5\nService: ' +
             (p.service || '—') + '\n\n"' + (p.review || '') + '"\n\n' +
             'Approve it in the studio panel: ' + SITE_URL + '/admin.html';

  adminSend('New review — ' + name + ' · ' + n + '/5', text,
            adminShell('New review · awaiting approval', escHtml(name), rows,
                       panel('What they wrote', p.review, '#ffaa1b'), cta));
}

// ── NEW WAITLIST SIGNUP ──
function notifyWaitlist(p) {
  var name = p.name || 'Someone';
  var rows =
    row('Name', name, { strong: true }) +
    row('Email', '<a href="mailto:' + escHtml(p.email) + '" style="color:#2c7449;text-decoration:none;">' +
                 escHtml(p.email) + '</a>', { raw: true }) +
    row('Interested in', p.service) +
    row('Joined', p.date);

  var note = '<tr><td style="padding:20px 32px 0 32px;font-family:Helvetica,Arial,sans-serif;' +
    'font-size:14px;line-height:22px;color:#5c4f3d;">' +
    'When you reopen commissions, you can convert them straight into a job from the Waitlist tab.' +
    '</td></tr>';

  var cta = ctaButtons('Open studio', SITE_URL + '/admin.html',
                       'Reply', 'mailto:' + encodeURIComponent(p.email));

  var text = 'NEW WAITLIST SIGNUP — ' + name + '\n\nEmail: ' + (p.email || '—') +
             '\nInterested in: ' + (p.service || '—') + '\n\n' + SITE_URL + '/admin.html';

  adminSend('Waitlist signup — ' + name, text,
            adminShell('Waitlist signup', escHtml(name), rows, note, cta));
}

/**
 * Auto-reply sent to the client. Built as a table-based HTML email because
 * flexbox, grid and external stylesheets are unreliable across mail clients.
 * A plain-text version is included for clients that block HTML.
 */
function buildConfirmationEmail(name, service) {
  var firstName = String(name || '').trim().split(/\s+/)[0] || 'there';
  var what = service ? String(service) : 'your project';

  var text =
    'Hi ' + firstName + ',\n\n' +
    'Thanks so much for reaching out to BoviBerry! I\'ve received your request for ' + what +
    ' and I\'m excited to get started!\n\n' +
    'Before I put together a quote and timeline, I\'ll follow up within 1-2 business days with a few ' +
    'quick questions to make sure I nail down exactly what you\'re envisioning. In the meantime, if you ' +
    'have any reference images, mood boards, or other links handy, feel free to reply with those — ' +
    'they help a ton!\n\n' +
    'Talk soon!\n' +
    'Bovi B.\n\n' +
    '—\nBoviBerry · ' + SITE_URL;

  var html =
'<!DOCTYPE html><html><head><meta charset="utf-8">' +
'<meta name="viewport" content="width=device-width,initial-scale=1">' +
'<title>Thanks for reaching out</title></head>' +
'<body style="margin:0;padding:0;background-color:#f5efe3;">' +

'<div style="display:none;max-height:0;overflow:hidden;opacity:0;">' +
  'Got your request — I\'ll follow up within 1–2 business days.' +
'</div>' +

'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
       'style="background-color:#f5efe3;padding:32px 16px;">' +
'<tr><td align="center">' +

  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
         'style="max-width:560px;background-color:#ffffff;border:1px solid #e0d6c4;">' +

    // green rule
    '<tr><td style="height:4px;background-color:#2c7449;font-size:0;line-height:0;">&nbsp;</td></tr>' +

    // wordmark
    '<tr><td style="padding:36px 40px 0 40px;" align="center">' +
      '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:26px;letter-spacing:1px;color:#2c7449;">' +
        'BoviBerry' +
      '</div>' +
      '<div style="font-family:Helvetica,Arial,sans-serif;font-size:10px;letter-spacing:3px;' +
                  'text-transform:uppercase;color:#9a8b74;padding-top:8px;">' +
        'Art &amp; Design Studio' +
      '</div>' +
    '</td></tr>' +

    '<tr><td style="padding:28px 40px 0 40px;">' +
      '<div style="height:1px;background-color:#e0d6c4;font-size:0;line-height:0;">&nbsp;</div>' +
    '</td></tr>' +

    // body
    '<tr><td style="padding:28px 40px 8px 40px;font-family:Helvetica,Arial,sans-serif;' +
                   'font-size:16px;line-height:26px;color:#3d3226;">' +
      '<p style="margin:0 0 20px 0;">Hi ' + escHtml(firstName) + ',</p>' +
      '<p style="margin:0 0 20px 0;">' +
        'Thanks so much for reaching out to BoviBerry! I&rsquo;ve received your request for ' +
        '<strong style="color:#2c7449;">' + escHtml(what) + '</strong> and I&rsquo;m excited to get started!' +
      '</p>' +
      '<p style="margin:0 0 20px 0;">' +
        'Before I put together a quote and timeline, I&rsquo;ll follow up within ' +
        '<strong>1&ndash;2 business days</strong> with a few quick questions to make sure I nail down ' +
        'exactly what you&rsquo;re envisioning.' +
      '</p>' +
      '<p style="margin:0 0 24px 0;">' +
        'In the meantime, if you have any reference images, mood boards, or other links handy, ' +
        'feel free to reply with those &mdash; they help a ton!' +
      '</p>' +
      '<p style="margin:0 0 4px 0;">Talk soon!</p>' +
      '<p style="margin:0;font-family:Georgia,\'Times New Roman\',serif;font-size:19px;' +
                'font-style:italic;color:#2c7449;">Bovi B.</p>' +
    '</td></tr>' +

    // mood board nudge
    '<tr><td style="padding:24px 40px 0 40px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
             'style="background-color:#f7f3ea;border-left:3px solid #9bd16b;">' +
        '<tr><td style="padding:16px 18px;font-family:Helvetica,Arial,sans-serif;' +
                       'font-size:14px;line-height:22px;color:#5c4f3d;">' +
          'Want to show me the look you&rsquo;re after? You can build a mood board ' +
          '<a href="' + SITE_URL + '/moodboard.html" style="color:#2c7449;font-weight:bold;">right here</a> ' +
          '&mdash; images, links, colours, all in one place.' +
        '</td></tr>' +
      '</table>' +
    '</td></tr>' +

    // footer
    '<tr><td style="padding:30px 40px 34px 40px;" align="center">' +
      '<div style="height:1px;background-color:#e0d6c4;font-size:0;line-height:0;margin-bottom:20px;">&nbsp;</div>' +
      '<div style="font-family:Helvetica,Arial,sans-serif;font-size:12px;line-height:20px;color:#9a8b74;">' +
        '<a href="' + SITE_URL + '" style="color:#2c7449;text-decoration:none;">boviberry.com</a>' +
        '<br>' +
        '<a href="mailto:' + REPLY_TO + '" style="color:#9a8b74;text-decoration:none;">' + REPLY_TO + '</a>' +
      '</div>' +
    '</td></tr>' +

  '</table>' +

  '<div style="font-family:Helvetica,Arial,sans-serif;font-size:11px;color:#b0a18a;' +
              'padding-top:18px;max-width:560px;">' +
    'You received this because you submitted a request at boviberry.com.' +
  '</div>' +

'</td></tr></table></body></html>';

  return {
    subject: 'Thanks for reaching out, ' + firstName + ' — BoviBerry',
    text: text,
    html: html
  };
}

function escHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Sends the auto-reply. Never throws — a bad address must not lose a submission. */
function sendClientConfirmation(toEmail, name, service) {
  var addr = String(toEmail || '').trim();
  if (!addr || addr.indexOf('@') === -1) return;

  var mail = buildConfirmationEmail(name, service);
  var opts = {
    to: addr,
    subject: mail.subject,
    body: mail.text,
    htmlBody: mail.html,
    name: SENDER_NAME,
    replyTo: REPLY_TO
  };

  // Try the branded From address first
  if (SEND_AS) {
    try {
      opts.from = SEND_AS;
      MailApp.sendEmail(opts);
      return;
    } catch (e) {
      // Alias not verified yet — fall through rather than dropping the email
      delete opts.from;
    }
  }

  try {
    MailApp.sendEmail(opts);
  } catch (e) { /* silent — the submission itself already succeeded */ }
}

/**
 * Run this to check whether the send-as alias is live.
 * If Boviberrystudios@gmail.com isn't listed, finish verifying it in Gmail:
 * Settings → Accounts and Import → Send mail as → Add another email address.
 */
function listMyAliases() {
  var aliases = GmailApp.getAliases();
  Logger.log('Script owner: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Verified aliases: ' + (aliases.length ? aliases.join(', ') : '(none yet)'));
  Logger.log(aliases.indexOf(SEND_AS) > -1
    ? 'READY — emails will send from ' + SEND_AS
    : 'NOT READY — ' + SEND_AS + ' is not verified yet, so emails will send from the owner address.');
}

/** Preview the email in your own inbox without submitting a form. */
function testConfirmationEmail() {
  sendClientConfirmation(NOTIFY_EMAIL, 'Jane Doe', 'Logo — Mascot');
  Logger.log('Test confirmation sent to ' + NOTIFY_EMAIL);
}


// ── JSON RESPONSE ──
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
/**
 * Scores commissions and mood boards submitted before spam detection existed.
 * Run once from the editor: pick "rescanForSpam", click Run, check the log.
 *
 * Only sets flags — never deletes anything, and never re-flags something
 * you've already reviewed and dismissed by hand.
 */
function rescanForSpam() {
  var report = [];

  [['Commissions', 'vision'], ['Moodboards', 'details']].forEach(function (pair) {
    var sheetName = pair[0];
    var sheet = getSheet(sheetName);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) { report.push(sheetName + ': nothing to scan.'); return; }

    var hdr = data[0];
    var flagCol = hdr.indexOf('flag') + 1;
    var reasonCol = hdr.indexOf('flagReason') + 1;
    if (!flagCol || !reasonCol) {
      report.push(sheetName + ': flag columns missing — deploy a new version first.');
      return;
    }

    var flagged = 0;
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < hdr.length; j++) obj[hdr[j]] = String(data[i][j] || '');

      // leave anything already reviewed and cleared alone
      if (String(obj.flag).toLowerCase() === 'false' && obj.flagReason) continue;

      // historic rows have no timing or honeypot data — score on content only
      var res = spamScore({
        name: obj.name, email: obj.email, phone: obj.phone,
        service: obj.service, budget: obj.budget, vision: obj[pair[1]]
      });

      if (res.score >= SPAM_THRESHOLD) {
        sheet.getRange(i + 1, flagCol).setValue('true');
        sheet.getRange(i + 1, reasonCol).setValue(res.reasons.join(' · '));
        flagged++;
        report.push('   FLAGGED (' + res.score + ') ' + obj.name + ' — ' + res.reasons.join(' · '));
      } else {
        sheet.getRange(i + 1, flagCol).setValue('false');
      }
    }
    report.push(sheetName + ': ' + flagged + ' of ' + (data.length - 1) + ' flagged.');
  });

  var out = report.join('\n');
  Logger.log(out);
  return out;
}


// GET HANDLER
// ============================================================
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var action = p.action;

  // ---------- PUBLIC ----------
  if (action === 'getApproved') {
    return json(rowsToObjects(getSheet('Reviews')).filter(function (r) {
      return String(r.approved).toLowerCase() === 'true';
    }));
  }

  if (action === 'getFeatured') {
    return json(rowsToObjects(getSheet('Reviews')).filter(function (r) {
      return String(r.approved).toLowerCase() === 'true' &&
             String(r.featured).toLowerCase() === 'true';
    }));
  }

  if (action === 'getSettings') {
    return json({ waitlist: getSetting('waitlist') });
  }

  // Public — the site reads this to apply current prices and sale flags
  if (action === 'getCatalog') {
    return json(rowsToObjects(getSheet('Catalog')));
  }

  // Public — the pricing and home pages read this to render live prices
  
  if (action === 'adminLogin') {
    return json(adminLogin(p.h));
  }

  // ---------- ADMIN (token required) ----------
  if (!validToken(p.token)) {
    return json({ error: 'unauthorized' });
  }

  // One request instead of four — much faster admin load
  if (action === 'getAdminBundle') {
    return json({
      commissions: rowsToObjects(getSheet('Commissions')),
      reviews:     rowsToObjects(getSheet('Reviews')),
      waitlist:    rowsToObjects(getSheet('Waitlist')),
      moodboards:  rowsToObjects(getSheet('Moodboards')),
      catalog:     rowsToObjects(getSheet('Catalog')),
      waitlistOn:  getSetting('waitlist')
    });
  }

  // Images are heavy — loaded separately, only when the tab is opened
  if (action === 'getMoodboardImages') {
    var all = rowsToObjects(getSheet('MoodboardImages'));
    var byBoard = {};
    for (var i = 0; i < all.length; i++) {
      var mid = all[i].moodboardId;
      if (!byBoard[mid]) byBoard[mid] = [];
      byBoard[mid].push(all[i].imageData);
    }
    return json(byBoard);
  }

  if (action === 'getAll')         return json(rowsToObjects(getSheet('Reviews')));
  if (action === 'getCommissions') return json(rowsToObjects(getSheet('Commissions')));
  if (action === 'getWaitlist')    return json(rowsToObjects(getSheet('Waitlist')));
  if (action === 'getMoodboards')  return json(rowsToObjects(getSheet('Moodboards')));

  return json({ error: 'Unknown action' });
}


// ============================================================
// POST HANDLER
// ============================================================
function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ error: 'Invalid JSON' });
  }

  var action = payload.action;

  // ---------- PUBLIC SUBMISSIONS ----------
  if (action === 'submitReview') {
    verifyTurnstile(payload.cfToken);
    getSheet('Reviews').appendRow([
      Date.now().toString(),
      payload.name || '', payload.service || '', payload.rating || '5',
      payload.review || '', payload.date || new Date().toLocaleDateString(),
      'false', 'false'
    ]);
    notifyReview(payload);
    return json({ ok: true });
  }

  if (action === 'submitCommission') {
    var cScore = spamScore(payload);
    if (!verifyTurnstile(payload.cfToken)) {
      cScore.score += 60;
      cScore.reasons.push('Failed the CAPTCHA check');
    }
    var cFlagged = cScore.score >= SPAM_THRESHOLD;

    getSheet('Commissions').appendRow([
      Date.now().toString(),
      payload.name || '',
      payload.email || '',
      payload.phone || 'Not provided',
      payload.socialPlatform || '—',
      payload.socialHandle || '—',
      payload.service || '',
      payload.budget || '',
      payload.vision || '',
      payload.references || 'None provided',
      payload.timeline || '',
      payload.referral || '—',
      payload.date || new Date().toLocaleDateString(),
      'new',
      cFlagged ? 'true' : 'false',
      cScore.reasons.join(' · ')
    ]);
    var unclaimed = rowsToObjects(getSheet('Commissions')).filter(function (r) {
      return r.status === 'new';
    }).length;
    notifyCommission(payload, unclaimed, cFlagged ? cScore : null);
    // Flagged submissions get no auto-reply — sending to fabricated addresses
    // damages sender reputation over time.
    if (!cFlagged) sendClientConfirmation(payload.email, payload.name, payload.service);
    return json({ ok: true });
  }

  if (action === 'joinWaitlist') {
    verifyTurnstile(payload.cfToken);
    getSheet('Waitlist').appendRow([
      Date.now().toString(),
      payload.name || '', payload.email || '', payload.service || '',
      payload.date || new Date().toLocaleDateString()
    ]);
    notifyWaitlist(payload);
    return json({ ok: true });
  }

  if (action === 'submitMoodboard') {
    var mScore = spamScore(payload);
    if (!verifyTurnstile(payload.cfToken)) {
      mScore.score += 60;
      mScore.reasons.push('Failed the CAPTCHA check');
    }
    var mFlagged = mScore.score >= SPAM_THRESHOLD;

    var mbId = Date.now().toString();
    getSheet('Moodboards').appendRow([
      mbId,
      payload.name || '', payload.email || '',
      payload.links || 'None', payload.details || 'None', payload.colors || 'None',
      payload.date || new Date().toLocaleDateString(),
      mFlagged ? 'true' : 'false',
      mScore.reasons.join(' · ')
    ]);
    if (payload.images && payload.images.length) {
      var imgSheet = getSheet('MoodboardImages');
      var rows = [];
      for (var i = 0; i < payload.images.length; i++) {
        rows.push([mbId + '-' + i, mbId, payload.images[i]]);
      }
      imgSheet.getRange(imgSheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    }
    notifyMoodboard(payload, mFlagged ? mScore : null);
    if (!mFlagged) sendClientConfirmation(payload.email, payload.name, 'your mood board');
    return json({ ok: true });
  }

  // ---------- ADMIN ACTIONS (token required) ----------
  if (!validToken(payload.token)) {
    return json({ error: 'unauthorized' });
  }

  if (action === 'adminLogout') {
    adminLogout(payload.token);
    return json({ ok: true });
  }

  if (action === 'approveReview') {
    var r1 = findRowById(getSheet('Reviews'), payload.id);
    if (r1 > 0) getSheet('Reviews').getRange(r1, 7).setValue('true');
    return json({ ok: true });
  }

  if (action === 'featureReview') {
    var r2 = findRowById(getSheet('Reviews'), payload.id);
    if (r2 > 0) getSheet('Reviews').getRange(r2, 8).setValue(payload.feature ? 'true' : 'false');
    return json({ ok: true });
  }

  if (action === 'deleteReview') {
    var r3 = findRowById(getSheet('Reviews'), payload.id);
    if (r3 > 0) getSheet('Reviews').deleteRow(r3);
    return json({ ok: true });
  }

  if (action === 'updateCommissionStatus') {
    var r4 = findRowById(getSheet('Commissions'), payload.id);
    if (r4 > 0) getSheet('Commissions').getRange(r4, 14).setValue(payload.status);
    return json({ ok: true });
  }

  if (action === 'deleteCommission') {
    var r5 = findRowById(getSheet('Commissions'), payload.id);
    if (r5 > 0) getSheet('Commissions').deleteRow(r5);
    return json({ ok: true });
  }

  if (action === 'convertWaitlistToCommission') {
    var wSheet = getSheet('Waitlist');
    var wRow = findRowById(wSheet, payload.id);
    if (wRow > 0) {
      var row = wSheet.getDataRange().getValues()[wRow - 1];
      getSheet('Commissions').appendRow([
        Date.now().toString(), row[1], row[2], 'Not provided', '—', '—',
        row[3], '—', 'Converted from waitlist', 'None provided',
        '—', 'Waitlist', new Date().toLocaleDateString(), 'new'
      ]);
      wSheet.deleteRow(wRow);
    }
    return json({ ok: true });
  }

  if (action === 'updateCatalogItem') {
    var cSheet = getSheet('Catalog');
    var cRow = findRowById(cSheet, payload.id);
    if (cRow > 0) {
      // Resolve columns by header name so adding a field never needs a code change
      var cHdr = cSheet.getRange(1, 1, 1, cSheet.getLastColumn()).getValues()[0];
      Object.keys(payload).forEach(function (key) {
        if (key === 'action' || key === 'id' || key === 'token') return;
        var ci = cHdr.indexOf(key);
        if (ci === -1) return;
        var val = payload[key];
        if (typeof val === 'boolean') val = val ? 'true' : 'false';
        cSheet.getRange(cRow, ci + 1).setValue(val);
      });
    }
    return json({ ok: true });
  }

  if (action === 'clearFlag') {
    var fSheet = getSheet(payload.sheet === 'moodboard' ? 'Moodboards' : 'Commissions');
    var fRow = findRowById(fSheet, payload.id);
    if (fRow > 0) {
      var hdr = fSheet.getRange(1, 1, 1, fSheet.getLastColumn()).getValues()[0];
      var fi = hdr.indexOf('flag');
      if (fi > -1) fSheet.getRange(fRow, fi + 1).setValue('false');
    }
    return json({ ok: true });
  }

  if (action === 'addCatalogItem') {
    var aSheet = getSheet('Catalog');
    var newId = String(payload.id || '').trim().toLowerCase().replace(/[^a-z0-9-]/g, '-');
    if (!newId) return json({ error: 'An id is required' });
    if (findRowById(aSheet, newId) > 0) return json({ error: 'That id already exists' });
    aSheet.appendRow([
      newId,
      payload.label || newId,
      payload.price || '',
      '', 'false', 'true',
      payload.note || '',
      payload.kind || 'quote',
      payload.formGroup || '',
      payload.turnaround || '',
      payload.minDays || 0,
      payload.stripeLink || '',
      payload.sortOrder || 999
    ]);
    return json({ ok: true, id: newId });
  }

  if (action === 'deleteCatalogItem') {
    var dSheet = getSheet('Catalog');
    var dRow = findRowById(dSheet, payload.id);
    if (dRow > 0) dSheet.deleteRow(dRow);
    return json({ ok: true });
  }

  if (action === 'setWaitlist') {
    setSetting('waitlist', payload.enabled ? 'true' : 'false');
    return json({ ok: true });
  }

  if (action === 'deleteMoodboard') {
    var mbSheet = getSheet('Moodboards');
    var mbRow = findRowById(mbSheet, payload.id);
    if (mbRow > 0) mbSheet.deleteRow(mbRow);
    var imgSheet2 = getSheet('MoodboardImages');
    var imgData = imgSheet2.getDataRange().getValues();
    for (var k = imgData.length - 1; k >= 1; k--) {
      if (String(imgData[k][1]) === String(payload.id)) imgSheet2.deleteRow(k + 1);
    }
    return json({ ok: true });
  }

  return json({ error: 'Unknown action' });
}
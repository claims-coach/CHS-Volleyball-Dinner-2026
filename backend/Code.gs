/**
 * Cascade Bruins Volleyball — Dinner Sign-Up Backend (Google Apps Script)
 * 2026-27 Season
 *
 * Stores signups in a Google Sheet; enforces SLOTS_PER_DATE per date (home dinner);
 * emails volunteer, coach, and organizers on confirmation.
 *
 * NOTE: Browser clients must use GET ?action=signup (POST to /exec 302→405).
 */

const SHEET_NAME = 'Signups';
const SLOTS_PER_DATE = 2;

const ORGANIZER_EMAILS = ['walker.alli@gmail.com', 'johnny@claims.coach'];
const COACH_EMAIL = 'NArevalo@everettsd.org';

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1,1,1,8).setValues([['date_id','label','name','email','phone','notes','timestamp','source']]);
    sh.getRange(1,1,1,2).setNumberFormat('@');
  }
  return sh;
}

function normalizeId_(id) {
  if (id instanceof Date) {
    return Utilities.formatDate(id, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(id || '');
}

/** listing[dateId] = { slots: [...], filled, capacity, open, full } */
function buildListing_() {
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  const listing = {};
  for (let i=1; i<rows.length; i++) {
    const [id,label,name,email,phone,notes,ts] = rows[i];
    if (!id) continue;
    const key = normalizeId_(id);
    if (!listing[key]) listing[key] = { slots: [] };
    listing[key].slots.push({
      name: name,
      email: email,
      phone: phone,
      notes: notes,
      ts: ts,
      label: (label instanceof Date) ? key : label
    });
  }
  Object.keys(listing).forEach(function(key) {
    const slots = listing[key].slots;
    const filled = slots.length;
    listing[key] = {
      slots: slots,
      filled: filled,
      capacity: SLOTS_PER_DATE,
      open: Math.max(0, SLOTS_PER_DATE - filled),
      full: filled >= SLOTS_PER_DATE,
      // legacy single-slot fields (first volunteer) for older clients
      name: slots[0] ? slots[0].name : '',
      email: slots[0] ? slots[0].email : '',
      phone: slots[0] ? slots[0].phone : '',
      notes: slots[0] ? slots[0].notes : '',
      ts: slots[0] ? slots[0].ts : '',
      label: slots[0] ? slots[0].label : key
    };
  });
  return listing;
}

function countForDate_(sh, id) {
  const data = sh.getDataRange().getValues();
  let n = 0;
  for (let i=1; i<data.length; i++) {
    if (normalizeId_(data[i][0]) === id) n++;
  }
  return n;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function safe_(s) {
  return String(s||'').replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]));
}

function isValidEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(s||'').trim());
}

function sendConfirmations_(payload) {
  const name = payload.name;
  const label = payload.label || payload.id;
  const volunteerEmail = String(payload.email||'').trim();
  const phone = payload.phone || '';
  const notes = payload.notes || '';

  if (!isValidEmail_(volunteerEmail)) throw new Error('No valid volunteer email');

  const subject = 'Cascade Bruins Volleyball: ' + name + ' confirmed for ' + label;
  const isAway = /AWAY|Snacks/i.test(label);
  const foodWord = isAway ? 'snacks' : 'dinner';
  const html = '' +
    '<p>Thanks, ' + safe_(name) + " — you're confirmed for <b>" + safe_(label) + '</b>.</p>' +
    '<p>Email: ' + safe_(volunteerEmail) + '</p>' +
    (phone ? ('<p>Phone: ' + safe_(phone) + '</p>') : '') +
    (notes ? ('<p>Notes: ' + safe_(notes) + '</p>') : '') +
    '<p><b>Note on drinks:</b> Drinks are <i>not required</i>. If you choose to bring them, ' +
    'please stick to <b>water or sports drinks</b> — <b>no energy drinks or caffeinated beverages</b>.</p>' +
    '<p><b>Food quantity:</b> Please bring enough ' + foodWord + ' for about 15 people.</p>' +
    '<p>Go Bruins!</p>';

  // One message: volunteer To, organizers + coach on Cc (never duplicate To on Cc).
  const cc = [];
  function pushCc(addr) {
    const a = String(addr||'').trim().toLowerCase();
    if (!isValidEmail_(a)) return;
    if (a === volunteerEmail.toLowerCase()) return;
    if (cc.indexOf(a) === -1) cc.push(a);
  }
  pushCc(COACH_EMAIL);
  ORGANIZER_EMAILS.forEach(pushCc);

  try {
    const opts = { to: volunteerEmail, subject: subject, htmlBody: html };
    if (cc.length) opts.cc = cc.join(',');
    MailApp.sendEmail(opts);
    return { sent: 1, failed: [], to: volunteerEmail, cc: cc };
  } catch (err) {
    console.error('Email error', err);
    return { sent: 0, failed: [String(err)], to: volunteerEmail, cc: cc };
  }
}

function reserveSlot_(body) {
  const id = String(body.id || '').trim();
  const label = String(body.label || id).trim();
  const name = String(body.name || '').trim();
  const email = String(body.email || '').trim();
  const phone = String(body.phone || '').trim();
  const notes = String(body.notes || '').trim();

  if (!id || !name || !email) throw new Error('Missing required fields');
  if (!isValidEmail_(email)) throw new Error('Please enter a valid email address');

  const lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    const sh = getSheet_();
    const filled = countForDate_(sh, id);
    if (filled >= SLOTS_PER_DATE) {
      return { ok:false, error:'That date is already full (' + SLOTS_PER_DATE + ' volunteers).' };
    }

    const row = sh.getLastRow() + 1;
    sh.getRange(row,1,1,8).setValues([[id, label, name, email, phone, notes, new Date(), 'webapp']]);
    sh.getRange(row,1,1,2).setNumberFormat('@');

    let mail = { sent: 0, failed: [] };
    try {
      mail = sendConfirmations_({ id:id, label:label, name:name, email:email, phone:phone, notes:notes });
    } catch (errMail) {
      console.error('Email error:', errMail);
      mail = { sent: 0, failed: [String(errMail)] };
    }

    return { ok:true, mail: mail, filled: filled + 1, capacity: SLOTS_PER_DATE };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function clearDates_(params) {
  const secret = String(params.secret || '').trim();
  // Lightweight gate: must match organizer email domain intent; Johnny/Alli use this once.
  if (secret !== 'bruins-clear-2026') throw new Error('Unauthorized');
  const raw = String(params.ids || '').trim();
  if (!raw) throw new Error('Missing ids');
  const ids = raw.split(',').map(function(s){ return String(s).trim(); }).filter(Boolean);
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  let removed = 0;
  // delete from bottom up
  for (let i = data.length - 1; i >= 1; i--) {
    const key = normalizeId_(data[i][0]);
    if (ids.indexOf(key) !== -1) {
      sh.deleteRow(i + 1);
      removed++;
    }
  }
  return { ok:true, removed: removed, ids: ids };
}

function fixAndResend_(params) {
  const id = String(params.id || '').trim();
  const emailFix = String(params.email || '').trim();
  if (!id) throw new Error('Missing id');
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (normalizeId_(data[i][0]) === id) {
      const row = i + 1;
      if (emailFix) {
        if (!isValidEmail_(emailFix)) throw new Error('Invalid email fix');
        sh.getRange(row, 4).setValue(emailFix);
        data[i][3] = emailFix;
      }
      const email = String(data[i][3] || '').trim();
      const name = String(data[i][2] || '').trim();
      const label = String(data[i][1] || id);
      const phone = String(data[i][4] || '');
      const notes = String(data[i][5] || '');
      const mail = sendConfirmations_({ id:id, label:label, name:name, email:email, phone:phone, notes:notes });
      return { ok:true, email:email, mail:mail };
    }
  }
  throw new Error('Signup not found for id ' + id);
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = (params.action || '').toLowerCase();

  if (action === 'list') return json_(buildListing_());

  if (action === 'signup' || action === 'reserve') {
    try { return json_(reserveSlot_(params)); }
    catch (err) { return json_({ ok:false, error: String(err) }); }
  }

  if (action === 'resend') {
    try { return json_(fixAndResend_(params)); }
    catch (err) { return json_({ ok:false, error: String(err) }); }
  }

  if (action === 'clear') {
    try { return json_(clearDates_(params)); }
    catch (err) { return json_({ ok:false, error: String(err) }); }
  }

  return json_({ ok:true, hint:'Use ?action=list|signup|resend|clear', capacity: SLOTS_PER_DATE });
}

function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    return json_(reserveSlot_(body));
  } catch (err) {
    return json_({ ok:false, error: String(err) });
  }
}

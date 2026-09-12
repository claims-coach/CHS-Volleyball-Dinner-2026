/**
 * Cascade Bruins Volleyball — Dinner Sign‑Up Backend (Google Apps Script)
 * 2026-27 Season
 * 
 * Stores signups in a Google Sheet; enforces single slot per date; 
 * emails parent, coach, and organizer on confirmation.
 */

const SHEET_NAME = 'Signups';

const ORGANIZER_EMAILS = ['walker.alli@gmail.com', 'johnny@claims.coach'];
const COACH_EMAIL = 'NArevalo@everettsd.org';

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1,1,1,8).setValues([['date_id','label','name','email','phone','notes','timestamp','source']]);
  }
  return sh;
}

function buildListing_() {
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  const listing = {};
  for (let i=1; i<rows.length; i++) {
    const [id,label,name,email,phone,notes,ts] = rows[i];
    const normId = normalizeId_(id);
    if (normId) listing[normId] = { name, email, phone, notes, ts };
  }
  return listing;
}

function normalizeId_(id) {
  if (!id) return '';
  return String(id).trim();
}

function isValidEmail_(email) {
  if (!email || typeof email !== 'string') return false;
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email.trim());
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = (params.action || '').toLowerCase();
  
  if (action === 'list') {
    const listing = buildListing_();
    return ContentService.createTextOutput(JSON.stringify(listing))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'signup' || action === 'reserve') {
    return handleSignup_(params);
  }
  
  if (action === 'resend') {
    return handleResend_(params);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ok:true, hint:'Use ?action=list, ?action=signup, or ?action=resend'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSignup_(params) {
  try {
    const id = normalizeId_(params.id);
    const name = (params.name || '').trim();
    const email = (params.email || '').trim();
    const phone = (params.phone || '').trim();
    const notes = (params.notes || '').trim();
    const label = (params.label || id).trim();
    
    if (!id || !name || !email) throw new Error('Missing required fields');
    if (!isValidEmail_(email)) throw new Error('Invalid email format');

    const lock = LockService.getScriptLock();
    lock.tryLock(30000);

    const sh = getSheet_();
    const data = sh.getDataRange().getValues();
    
    for (let i=1; i<data.length; i++) {
      if (normalizeId_(data[i][0]) === id) {
        lock.releaseLock();
        return json_({ ok:false, error:'That date is already taken.' });
      }
    }
    
    const newRow = sh.getLastRow() + 1;
    sh.appendRow([id, label, name, email, phone, notes, new Date(), 'webapp']);
    sh.getRange(newRow, 1).setNumberFormat('@');

    const emailResult = sendConfirmationEmails_(name, email, id, phone, notes);

    lock.releaseLock();
    return json_({ ok:true, emailsSent: emailResult.sent, emailsFailed: emailResult.failed });
    
  } catch (err) {
    console.error('handleSignup_ error:', err);
    try { LockService.getScriptLock().releaseLock(); } catch(_){ }
    return json_({ ok:false, error: String(err) });
  }
}

function handleResend_(params) {
  try {
    const id = normalizeId_(params.id);
    const newEmail = (params.email || '').trim();
    
    if (!id) throw new Error('Missing id parameter');

    const lock = LockService.getScriptLock();
    lock.tryLock(30000);

    const sh = getSheet_();
    const data = sh.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i=1; i<data.length; i++) {
      if (normalizeId_(data[i][0]) === id) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      lock.releaseLock();
      return json_({ ok:false, error:'No signup found for that date.' });
    }
    
    const row = data[rowIndex];
    let name = row[2];
    let email = row[3];
    let phone = row[4];
    let notes = row[5];
    
    if (newEmail && isValidEmail_(newEmail)) {
      email = newEmail;
      sh.getRange(rowIndex + 1, 4).setValue(email);
    }

    const emailResult = sendConfirmationEmails_(name, email, id, phone, notes);

    lock.releaseLock();
    return json_({ ok:true, emailsSent: emailResult.sent, emailsFailed: emailResult.failed, updatedEmail: newEmail ? true : false });
    
  } catch (err) {
    console.error('handleResend_ error:', err);
    try { LockService.getScriptLock().releaseLock(); } catch(_){ }
    return json_({ ok:false, error: String(err) });
  }
}

function sendConfirmationEmails_(name, volunteerEmail, dateId, phone, notes) {
  const subject = `Cascade Bruins Volleyball: ${name} confirmed for ${dateId}`;
  const html = `
    <p>Thanks, ${safe_(name)} — you're confirmed for <b>${safe_(dateId)}</b>.</p>
    <p>Email: ${safe_(volunteerEmail)}</p>
    ${phone ? `<p>Phone: ${safe_(phone)}</p>` : ''}
    ${notes ? `<p>Notes: ${safe_(notes)}</p>` : ''}
    <p><b>Food quantity:</b> Please bring enough dinner/snacks for <b>about 15 people</b> (players + coaches).</p>
    <p><b>Note on drinks:</b> Drinks are <i>not required</i>. If you choose to bring them, 
    please stick to <b>water or sports drinks</b> — <b>no energy drinks or caffeinated beverages</b>.</p>
    <p>Go Bruins!</p>`;
  
  const allRecipients = [volunteerEmail, COACH_EMAIL, ...ORGANIZER_EMAILS].filter(r => isValidEmail_(r));
  const uniqueRecipients = [...new Set(allRecipients)];
  
  let sent = [];
  let failed = [];
  
  for (const recipient of uniqueRecipients) {
    try {
      MailApp.sendEmail({ to: recipient, subject, htmlBody: html });
      sent.push(recipient);
    } catch (err) {
      console.error('Email failed for ' + recipient + ':', err);
      failed.push(recipient);
    }
  }
  
  return { sent, failed };
}

function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    const id = normalizeId_(body.id);
    const name = (body.name || '').trim();
    const email = (body.email || '').trim();
    const phone = (body.phone || '').trim();
    const notes = (body.notes || '').trim();
    const label = (body.label || id).trim();
    
    if (!id || !name || !email) throw new Error('Missing required fields');
    if (!isValidEmail_(email)) throw new Error('Invalid email format');

    const lock = LockService.getScriptLock();
    lock.tryLock(30000);

    const sh = getSheet_();
    const data = sh.getDataRange().getValues();
    
    for (let i=1; i<data.length; i++) {
      if (normalizeId_(data[i][0]) === id) {
        lock.releaseLock();
        return json_({ ok:false, error:'That date is already taken.' });
      }
    }
    
    const newRow = sh.getLastRow() + 1;
    sh.appendRow([id, label, name, email, phone, notes, new Date(), 'webapp']);
    sh.getRange(newRow, 1).setNumberFormat('@');

    const emailResult = sendConfirmationEmails_(name, email, id, phone, notes);

    lock.releaseLock();
    return json_({ ok:true, emailsSent: emailResult.sent, emailsFailed: emailResult.failed });
    
  } catch (err) {
    console.error('doPost error:', err);
    try { LockService.getScriptLock().releaseLock(); } catch(_){ }
    return json_({ ok:false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function safe_(s) {
  return String(s||'').replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]));
}

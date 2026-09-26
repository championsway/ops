/*************************************************************************
 * CW — GOHIGHLEVEL CONTACTS
 *
 * Read-only. Writes results to Drive files for review. Sends nothing,
 * changes nothing in GHL.
 *
 *   cwGhl_Audit()              THE USEFUL ONE. Walks every member the DONE
 *                              engine tries to message and reports exactly
 *                              who it cannot resolve, and why.
 *   cwGhl_DumpContacts()       Full contact list -> Drive JSON + a CSV.
 *   cwGhl_Find('broome')       Quick lookup by name, email or phone.
 *
 * Relies on helpers already in the project: doneGhlApiKey_, doneGhlBaseUrl_,
 * doneGhlLocationId_, doneBuildEmailIndex_, doneFindContactByEmail_,
 * doneNameKey_, DONE_CFG.
 *************************************************************************/

var CWG_SS_ID = '1dwUGqqYqgAjQYyYMeFkA5UEVsCFTH01S89z93qNd1ys';
var CWG_TZ    = 'America/Chicago';
var CWG_MAX_PAGES = 60;                 // 100 per page; hard stop so it can't run away

// ---------------------------------------------------------------- fetch all
/**
 * Pages through every contact in the location. Returns an array of
 * {id, name, firstName, lastName, email, phone, tags, dateAdded}.
 */
function cwGhl_fetchAll_() {
  var base    = doneGhlBaseUrl_();
  var locId   = doneGhlLocationId_();
  var headers = {
    'Authorization': 'Bearer ' + doneGhlApiKey_(),
    'Version': DONE_CFG.ghlReadVersion
  };

  var all = [], startAfter = '', startAfterId = '', page = 0;

  while (page < CWG_MAX_PAGES) {
    var url = base + '/contacts/?locationId=' + encodeURIComponent(locId) + '&limit=100';
    if (startAfterId) url += '&startAfterId=' + encodeURIComponent(startAfterId);
    if (startAfter)   url += '&startAfter='   + encodeURIComponent(startAfter);

    var resp = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
    var code = resp.getResponseCode();
    if (code >= 400) {
      Logger.log('[cwGhl_fetchAll_] HTTP ' + code + ' on page ' + page + ': ' +
                 resp.getContentText().slice(0, 300));
      break;
    }

    var data     = JSON.parse(resp.getContentText());
    var contacts = data.contacts || [];
    if (!contacts.length) break;

    for (var i = 0; i < contacts.length; i++) {
      var c = contacts[i];
      var nm = String(c.contactName || ((c.firstName || '') + ' ' + (c.lastName || ''))).trim();
      all.push({
        id: c.id || '',
        name: nm,
        firstName: c.firstName || '',
        lastName: c.lastName || '',
        email: String(c.email || '').toLowerCase().trim(),
        phone: c.phone || '',
        tags: c.tags || [],
        dateAdded: c.dateAdded || ''
      });
    }

    var meta = data.meta || {};
    if (!meta.startAfterId && !meta.nextPageUrl) break;      // last page
    startAfterId = meta.startAfterId || '';
    startAfter   = meta.startAfter   || '';
    page++;
    Utilities.sleep(180);
  }

  Logger.log('[cwGhl_fetchAll_] ' + all.length + ' contacts over ' + (page + 1) + ' page(s)');
  return all;
}

// ------------------------------------------------------------------- audit
/**
 * The DONE engine reaches a member in two hops:
 *     sheet name -> email (doneBuildEmailIndex_)
 *     email      -> GHL contact (doneFindContactByEmail_)
 * A break at either hop means that member is silently skipped on every run,
 * logged only as "(no email)" or "(no GHL contact)".
 *
 * This walks the same two hops for everyone with a DONE cue and reports who
 * fails, where, and what a likely fix looks like.
 */
function cwGhl_Audit() {
  var ss = SpreadsheetApp.openById(CWG_SS_ID);
  var dc = ss.getSheetByName('Daily DONE Check');
  var emailIdx = doneBuildEmailIndex_(ss);

  var contacts = cwGhl_fetchAll_();

  // index GHL contacts by email and by normalized name
  var byEmail = {}, byName = {};
  for (var i = 0; i < contacts.length; i++) {
    var c = contacts[i];
    if (c.email) byEmail[c.email] = c;
    if (c.name) {
      var k = doneNameKey_(c.name);
      if (!byName[k]) byName[k] = [];
      byName[k].push(c);
    }
  }

  var lr = dc.getLastRow();
  var rows = dc.getRange(2, 1, lr - 1, 6).getValues();

  var ok = [], noEmail = [], noContact = [], suppressed = [];

  for (var r = 0; r < rows.length; r++) {
    var suppress = rows[r][0];
    var name = String(rows[r][1] || '').trim();
    var cue  = String(rows[r][3] || '').trim();
    if (!name || !cue) continue;

    var isSuppressed = suppress && String(suppress).trim() &&
                       String(suppress).trim().toLowerCase() !== 'false';
    if (isSuppressed) { suppressed.push(name); continue; }

    var key   = doneNameKey_(name);
    var email = emailIdx[key] || '';

    if (!email) {
      // is there a GHL contact under this name anyway? then the sheet is what's missing
      var guess = byName[key];
      noEmail.push({
        name: name,
        cue: cue,
        ghlNameMatch: guess ? guess.map(function (g) { return g.name + ' <' + (g.email || 'no email') + '>'; }) : [],
        fix: guess ? 'GHL has this name — copy the email into the sheet'
                   : 'No email in the sheet and no GHL contact under this name'
      });
      continue;
    }

    var contactId = '';
    try { contactId = doneFindContactByEmail_(email); } catch (e) {}

    if (!contactId) {
      var nameHit = byName[key];
      noContact.push({
        name: name,
        cue: cue,
        sheetEmail: email,
        ghlNameMatch: nameHit ? nameHit.map(function (g) { return g.name + ' <' + (g.email || 'no email') + '>'; }) : [],
        fix: nameHit ? 'GHL has this person under a DIFFERENT email — reconcile the two'
                     : 'No GHL contact for this email; the contact may not exist yet'
      });
      continue;
    }

    ok.push({ name: name, email: email, contactId: contactId });
    Utilities.sleep(120);
  }

  // duplicate emails in GHL — these make contact resolution nondeterministic
  var emailSeen = {}, dupes = [];
  for (var d = 0; d < contacts.length; d++) {
    var e = contacts[d].email;
    if (!e) continue;
    if (emailSeen[e]) dupes.push({ email: e, names: [emailSeen[e], contacts[d].name] });
    else emailSeen[e] = contacts[d].name;
  }

  var payload = {
    generated: Utilities.formatDate(new Date(), CWG_TZ, 'yyyy-MM-dd HH:mm'),
    ghlContactsTotal: contacts.length,
    membersResolvedOk: ok.length,
    membersSuppressed: suppressed.length,
    BROKEN_noEmailInSheet: noEmail,
    BROKEN_noGhlContact: noContact,
    duplicateEmailsInGhl: dupes,
    resolvedOk: ok,
    suppressedNames: suppressed
  };

  var f = DriveApp.createFile('cw_ghl_audit.json', JSON.stringify(payload, null, 1), 'application/json');

  Logger.log(
    'FILE=' + f.getId() +
    '\nGHL contacts: ' + contacts.length +
    '\nResolved fine: ' + ok.length +
    '\nSuppressed (skipped on purpose): ' + suppressed.length +
    '\nBROKEN — no email in sheet: ' + noEmail.length +
    (noEmail.length ? '\n   ' + noEmail.map(function (x) { return x.name; }).join(', ') : '') +
    '\nBROKEN — no GHL contact: ' + noContact.length +
    (noContact.length ? '\n   ' + noContact.map(function (x) { return x.name; }).join(', ') : '') +
    '\nDuplicate emails in GHL: ' + dupes.length
  );
  return f.getId();
}

// -------------------------------------------------------------------- dump
/** Every contact, as JSON and as a CSV you can drop into a sheet. */
function cwGhl_DumpContacts() {
  var all = cwGhl_fetchAll_();

  var jf = DriveApp.createFile('cw_ghl_contacts.json', JSON.stringify({
    generated: Utilities.formatDate(new Date(), CWG_TZ, 'yyyy-MM-dd HH:mm'),
    total: all.length,
    contacts: all
  }, null, 1), 'application/json');

  function esc(v) { return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"'; }
  var csv = ['id,name,email,phone,tags,dateAdded'];
  for (var i = 0; i < all.length; i++) {
    var c = all[i];
    csv.push([esc(c.id), esc(c.name), esc(c.email), esc(c.phone),
              esc((c.tags || []).join('|')), esc(c.dateAdded)].join(','));
  }
  var cf = DriveApp.createFile('cw_ghl_contacts.csv', csv.join('\n'), 'text/csv');

  var missingEmail = all.filter(function (c) { return !c.email; }).length;
  var missingPhone = all.filter(function (c) { return !c.phone; }).length;

  Logger.log('JSON=' + jf.getId() + '\nCSV=' + cf.getId() +
             '\n' + all.length + ' contacts | no email: ' + missingEmail +
             ' | no phone: ' + missingPhone);
  return cf.getId();
}

// -------------------------------------------------------------------- find
/** Partial match on name, email or phone. Logs the hits. */
function cwGhl_Find(fragment) {
  var frag = String(fragment || '').toLowerCase().trim();
  if (!frag) { Logger.log('Pass something to search for, e.g. cwGhl_Find("broome")'); return ''; }

  var all = cwGhl_fetchAll_();
  var hits = all.filter(function (c) {
    return (c.name + ' ' + c.email + ' ' + c.phone).toLowerCase().indexOf(frag) >= 0;
  });

  if (!hits.length) { Logger.log('No GHL contact matching "' + fragment + '"'); return ''; }

  var lines = hits.map(function (c) {
    return c.name + ' | ' + (c.email || 'no email') + ' | ' + (c.phone || 'no phone') + ' | id=' + c.id;
  });
  Logger.log('MATCHES for "' + fragment + '" (' + hits.length + '):\n' + lines.join('\n'));
  return lines.join('\n');
}

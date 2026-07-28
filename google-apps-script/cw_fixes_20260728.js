/*************************************************************************
 * CW OPS FIXES — 2026-07-28
 *
 * Paste into the Champion's Way Automations project, run the functions
 * you want, then delete the file. Nothing here creates or deletes triggers
 * and nothing sends a message to a client.
 *
 *   cwFix_All()                  -> runs the four data repairs below
 *   cwFix_SherriMonday()         -> repair Jul 27 attendance that never counted
 *   cwFix_BriOutOfWeek()         -> set her blank "Out of (wk)" to 3
 *   cwFix_ChierstenRemoveMonday()-> drop Mon from her primary slot
 *   cwFix_BenRemoveMonday()      -> drop Mon from his primary slot, target 3 -> 2
 *   cwFix_NormalizeDayTokens()   -> Tues/Thur/Thurs -> Tue/Thu across the tab
 *
 *   cwDump_MonthlySummaries()    -> pull July Trainerize summaries to a Drive file
 *   cwDump_NinaGhlJuly()         -> pull Nina's July GHL thread to a Drive file
 *************************************************************************/

var CWFIX_SS_ID = '1dwUGqqYqgAjQYyYMeFkA5UEVsCFTH01S89z93qNd1ys';
var CWFIX_TZ    = 'America/Chicago';

function cwFix_All() {
  var out = [];
  out.push(cwFix_SherriMonday());
  out.push(cwFix_BriOutOfWeek());
  out.push(cwFix_ChierstenRemoveMonday());
  out.push(cwFix_BenRemoveMonday());
  Logger.log('\n===== SUMMARY =====\n' + out.join('\n'));
  return out.join(' | ');
}

// ---------------------------------------------------------------- helpers
function cwfix_accSheet_() {
  return SpreadsheetApp.openById(CWFIX_SS_ID).getSheetByName('Accountability Log');
}
function cwfix_stdSheet_() {
  return SpreadsheetApp.openById(CWFIX_SS_ID).getSheetByName('Standing_Schedules');
}

/** Find a row in Accountability Log by name (col C, data starts row 3). */
function cwfix_findAccRow_(name) {
  var sh = cwfix_accSheet_();
  var v = sh.getRange(3, 3, sh.getLastRow() - 2, 1).getValues();
  var want = String(name).toLowerCase().replace(/\s+/g, ' ').trim();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0] || '').toLowerCase().replace(/\s+/g, ' ').trim() === want) return i + 3;
  }
  return 0;
}

/** Find a row in Standing_Schedules by name (col B, data starts row 2). */
function cwfix_findStdRow_(name) {
  var sh = cwfix_stdSheet_();
  var v = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  var want = String(name).toLowerCase().replace(/\s+/g, ' ').trim();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0] || '').toLowerCase().replace(/\s+/g, ' ').trim() === want) return i + 2;
  }
  return 0;
}

/** Remove every "<Day> <time>" chunk for the given day from a "A / B / C" slot string. */
function cwfix_stripDay_(slotStr, dayPrefix) {
  var parts = String(slotStr || '').split('/');
  var kept = [];
  var re = new RegExp('^' + dayPrefix, 'i');
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i].trim();
    if (!p) continue;
    if (re.test(p)) continue;
    kept.push(p);
  }
  return kept.join(' / ');
}

// ------------------------------------------------- 1. Sherri, Mon Jul 27
/**
 * She attended Monday but Came (wk) still reads 0. The dedup stamp in col N
 * is written AFTER the D/H/K increment, so a run that died in between leaves
 * the stamp present and the counts untouched — and incrementAttendance() will
 * then skip her forever. This repairs the counts directly and reports the
 * stamp state so you can see whether that is what happened.
 */
function cwFix_SherriMonday() {
  var NAME  = 'Sherri McDonald';
  var LABEL = '4:30 PM MTTh';
  var DATE  = new Date(2026, 6, 27);           // Mon Jul 27 2026 (month is 0-based)

  var sh  = cwfix_accSheet_();
  var row = cwfix_findAccRow_(NAME);
  if (!row) return 'Sherri: ROW NOT FOUND — nothing changed';

  var dateKey  = Utilities.formatDate(DATE, CWFIX_TZ, 'yyyy-MM-dd');
  var stampKey = dateKey + '|' + LABEL;
  var stampCell = sh.getRange(row, 14);         // N — Today Export
  var stamp = String(stampCell.getValue() || '');
  var hadStamp = stamp.indexOf(stampKey) !== -1;

  var wk  = sh.getRange(row, 4);                // D — Came (wk)
  var mth = sh.getRange(row, 8);                // H — Came (mth)
  var ovr = sh.getRange(row, 11);               // K — Came (OVR)

  var before = [Number(wk.getValue()) || 0, Number(mth.getValue()) || 0, Number(ovr.getValue()) || 0];

  wk.setValue(before[0] + 1);
  mth.setValue(before[1] + 1);
  ovr.setValue(before[2] + 1);

  if (!hadStamp) stampCell.setValue(stamp ? stamp + ', ' + stampKey : stampKey);

  var msg = 'Sherri row ' + row + ': wk ' + before[0] + '->' + (before[0] + 1) +
            ', mth ' + before[1] + '->' + (before[1] + 1) +
            ', ovr ' + before[2] + '->' + (before[2] + 1) +
            ' | stamp ' + (hadStamp ? 'ALREADY PRESENT (confirms the partial-run bug)' : 'added');
  Logger.log(msg);
  return msg;
}

// ------------------------------------------------------- 2. Bri Trosclair
/** Her Accountability Log row exists but Out of (wk) is blank, so she never
 *  registers as having missed anything. Initialize the week target. */
function cwFix_BriOutOfWeek() {
  var NAME = 'Bri Trosclair';
  var sh = cwfix_accSheet_();
  var row = cwfix_findAccRow_(NAME);
  if (!row) return 'Bri: ROW NOT FOUND — nothing changed';

  var cell = sh.getRange(row, 5);               // E — Out of (wk)
  var before = cell.getValue();
  cell.setValue(3);

  // seed the counters so the row behaves like every other member
  var seeded = [];
  [[4, 'Came (wk)'], [8, 'Came (mth)'], [11, 'Came (OVR)']].forEach(function (c) {
    var cl = sh.getRange(row, c[0]);
    if (cl.getValue() === '' || cl.getValue() === null) { cl.setValue(0); seeded.push(c[1]); }
  });

  var msg = 'Bri row ' + row + ': Out of (wk) ' + JSON.stringify(before) + ' -> 3' +
            (seeded.length ? ' | seeded 0 into ' + seeded.join(', ') : '');
  Logger.log(msg);
  return msg;
}

// -------------------------------------------------- 3. Chiersten Murchison
/** Now TuTh only. Dropping Mon from her primary leaves Tue in primary and
 *  Thu already sitting in secondary, which matches her target of 2. */
function cwFix_ChierstenRemoveMonday() {
  return cwfix_removeMondayFrom_('Chiersten Murchison', null);
}

// -------------------------------------------------------- 4. Ben Young
/** Dropped from 3 days to 2 (Tu/Th). Removes Mon and sets the target to 2.
 *  His secondary slot (Fri 8:45 AM) is left alone — logged so you can decide. */
function cwFix_BenRemoveMonday() {
  return cwfix_removeMondayFrom_('Ben Young', 2);
}

function cwfix_removeMondayFrom_(name, newTarget) {
  var sh = cwfix_stdSheet_();
  var row = cwfix_findStdRow_(name);
  if (!row) return name + ': ROW NOT FOUND — nothing changed';

  var primaryCell = sh.getRange(row, 3);        // C — Primary Slot
  var before = String(primaryCell.getValue() || '');
  var after  = cwfix_stripDay_(before, 'Mon');
  primaryCell.setValue(after);

  var secondary = String(sh.getRange(row, 4).getValue() || '');
  var targetNote = '';
  if (newTarget !== null && newTarget !== undefined) {
    var tCell = sh.getRange(row, 5);            // E — Sessions/Wk Target
    var tBefore = tCell.getValue();
    tCell.setValue(newTarget);
    targetNote = ' | target ' + tBefore + ' -> ' + newTarget;
  }

  var msg = name + ' row ' + row + ': primary "' + before + '" -> "' + after + '"' + targetNote +
            ' | secondary left as "' + secondary + '" (review if this should stay)';
  Logger.log(msg);
  return msg;
}

// --------------------------------------------- 5. Day-token normalization
/**
 * 936 of the ~940 rows already use Mon/Tue/Wed/Thu/Fri. Four rows use Tues,
 * Thur or Thurs. Rather than guess what the Standing_Schedules reader expects,
 * this makes the outliers match the overwhelming majority already in the tab.
 *
 * Run with preview = true first to see exactly what would change.
 */
function cwFix_NormalizeDayTokens(preview) {
  preview = (preview === true);
  var sh = cwfix_stdSheet_();
  var last = sh.getLastRow();
  var rng = sh.getRange(2, 2, last - 1, 3);     // B name, C primary, D secondary
  var vals = rng.getValues();

  // longest-first so "Thurs" is matched before "Thur", and "Thur" before "Thu"
  var MAP = [
    [/\bThurs\b/g, 'Thu'], [/\bThur\b/g, 'Thu'], [/\bThursday\b/g, 'Thu'],
    [/\bTues\b/g,  'Tue'], [/\bTuesday\b/g, 'Tue'],
    [/\bMonday\b/g, 'Mon'], [/\bWednesday\b/g, 'Wed'], [/\bWeds\b/g, 'Wed'],
    [/\bFriday\b/g, 'Fri'], [/\bSaturday\b/g, 'Sat'], [/\bSunday\b/g, 'Sun']
  ];

  var changes = [];
  for (var i = 0; i < vals.length; i++) {
    for (var c = 1; c <= 2; c++) {              // primary, secondary
      var orig = String(vals[i][c] || '');
      if (!orig) continue;
      var fixed = orig;
      for (var m = 0; m < MAP.length; m++) fixed = fixed.replace(MAP[m][0], MAP[m][1]);
      if (fixed !== orig) {
        changes.push(vals[i][0] + ' [' + (c === 1 ? 'primary' : 'secondary') + '] "' + orig + '" -> "' + fixed + '"');
        vals[i][c] = fixed;
      }
    }
  }

  if (!preview && changes.length) rng.setValues(vals);

  var msg = (preview ? '[PREVIEW] ' : '[APPLIED] ') + changes.length + ' slot string(s)\n' + changes.join('\n');
  Logger.log(msg);
  return msg;
}

// ============================================================ DUMP HELPERS
// These only READ. They write findings to a Drive file so you can review the
// real wording before you send anything.

/**
 * Pulls every long trainer-sent Trainerize message from July for each client
 * with a DONE cue — that is where the monthly summary lives. Review the output
 * to see who actually received one and what it committed them to for August.
 */
function cwDump_MonthlySummaries() {
  var ss = SpreadsheetApp.openById(CWFIX_SS_ID);
  var dc = ss.getSheetByName('Daily DONE Check');
  var idx = cwdTzIndex_(ss);                    // name -> Trainerize id
  var since = new Date(2026, 6, 1);             // Jul 1 2026

  var lr = dc.getLastRow();
  var rows = dc.getRange(2, 1, lr - 1, 6).getValues();

  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var name = String(rows[i][1] || '').trim();
    var cue  = String(rows[i][3] || '').trim();
    if (!name || !cue) continue;

    var tzId = idx[doneNameKey_(name)] || '';
    if (!tzId) { out.push({ name: name, error: 'no Trainerize id' }); continue; }

    try {
      var th = mpa_tzPost_('message/getThreads',
        { userID: 3336777, view: 'byClient', clientID: Number(tzId), start: 0, count: 5 });
      if (th.code !== 200 || !th.data.threads || !th.data.threads.length) {
        out.push({ name: name, error: 'no thread' }); continue;
      }
      var threadID = th.data.threads[0].threadID;
      var cnt = mpa_tzPost_('message/getMessages', { threadID: threadID, start: 0, count: 1 });
      var total = (cnt.code === 200 && cnt.data.totalCount) ? cnt.data.totalCount : 0;
      var fetch = 60, startOff = Math.max(0, total - fetch);
      var res = mpa_tzPost_('message/getMessages', { threadID: threadID, start: startOff, count: fetch });
      if (res.code !== 200 || !res.data.messages) { out.push({ name: name, error: 'read failed' }); continue; }

      var found = [];
      for (var m = 0; m < res.data.messages.length; m++) {
        var msg = res.data.messages[m];
        var st = (msg.sender && msg.sender.type) ? msg.sender.type : '';
        if (st !== 'trainer' && st !== 'coach') continue;      // only what WE sent
        var at = new Date(msg.sentTime || msg.createdAt || 0);
        if (at.getTime() < since.getTime()) continue;
        var body = String(msg.body || '');
        if (body.length < 200) continue;                        // summaries are long
        found.push({ at: Utilities.formatDate(at, CWFIX_TZ, 'yyyy-MM-dd HH:mm'), body: body });
      }
      out.push({ name: name, cue: cue, summariesFound: found.length, messages: found });
      Utilities.sleep(250);                                     // be gentle on the API
    } catch (e) {
      out.push({ name: name, error: String(e) });
    }
  }

  var withNone = out.filter(function (o) { return !o.error && !o.summariesFound; })
                    .map(function (o) { return o.name; });

  var payload = {
    generated: Utilities.formatDate(new Date(), CWFIX_TZ, 'yyyy-MM-dd HH:mm'),
    window: 'July 2026',
    totalClients: out.length,
    clientsWithNoJulySummary: withNone,
    detail: out
  };
  var f = DriveApp.createFile('cw_monthly_summaries_july.json', JSON.stringify(payload, null, 1), 'application/json');
  Logger.log('FILE=' + f.getId() + '\nNo July summary found for: ' + withNone.join(', '));
  return f.getId();
}

/** Nina's full July GHL thread, so her freeze message can reference what was
 *  actually agreed rather than a guess. Her next change date is Jul 30. */
function cwDump_NinaGhlJuly() {
  var ss = SpreadsheetApp.openById(CWFIX_SS_ID);
  var emailIdx = doneBuildEmailIndex_(ss);
  var email = emailIdx[doneNameKey_('Nina Foster')] || '';
  if (!email) { Logger.log('Nina: no email on file'); return 'no email'; }

  var contactId = doneFindContactByEmail_(email);
  if (!contactId) { Logger.log('Nina: no GHL contact for ' + email); return 'no contact'; }

  var conv = tpFetchConvo_(doneGhlApiKey_(), contactId, new Date(2026, 6, 1));
  if (!conv) { Logger.log('Nina: GHL read failed'); return 'read failed'; }

  function fmt(list, dir) {
    return list.map(function (r) {
      return {
        dir: dir,
        at: r.at ? Utilities.formatDate(r.at, CWFIX_TZ, 'yyyy-MM-dd HH:mm') : '',
        text: r.text,
        attachments: (r.attachments || []).length
      };
    });
  }
  var all = fmt(conv.inbound, 'FROM NINA').concat(fmt(conv.outbound, 'FROM US'));
  all.sort(function (a, b) { return a.at < b.at ? -1 : 1; });

  var f = DriveApp.createFile('cw_nina_ghl_july.json', JSON.stringify({
    contactId: contactId,
    nextChangeDate: '2026-07-30',
    messageCount: all.length,
    messages: all
  }, null, 1), 'application/json');

  Logger.log('FILE=' + f.getId() + ' — ' + all.length + ' messages');
  return f.getId();
}

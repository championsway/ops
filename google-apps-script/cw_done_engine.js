/*************************************************************************
 * CW DONE ENGINE (2026-07-10)
 * ONE hourly trigger -> CW_doneTick(). Decides by Central hour what to do.
 *
 * Reads the rebuilt Daily DONE Check tab:
 *   col A suppress | col B name | col C tools | col D DONE cue | col E notes
 *   col 7..37 = Day 1..31 | col 38 Nudge Timing (we write the cue class here)
 *   39 Last DONE | 40 Last Checked | 41 Last Nudge | 42 Auto Status
 *
 * Per cue class we REMIND before the window, then CHECK after it:
 *   05 check(night, for YESTERDAY)
 *   07 remind(breakfast)      10 check(breakfast)
 *   11 remind(lunch)          15 check(lunch)
 *   13 remind(snack)          16 check(snack)
 *   17 remind(dinner) + remind(daily)
 *   20 remind(night)
 *   21 check(dinner) + check(daily)
 *
 * Idempotent: one remind per class per day (Script Property), and a check
 * never re-marks a day cell that already reads ✓ / ✗.
 * NEVER creates or deletes triggers.
 *************************************************************************/

var CWD_TZ = 'America/Chicago';

var CWD_SCHEDULE = {
  5:  [['check','night', -1]],
  7:  [['remind','breakfast', 0]],
  10: [['check','breakfast', 0]],
  11: [['remind','lunch', 0]],
  13: [['remind','snack', 0]],
  15: [['check','lunch', 0]],
  16: [['check','snack', 0]],
  17: [['remind','dinner', 0], ['remind','daily', 0], ['remind','nightshift', 0]],
  20: [['remind','night', 0]],
  21: [['check','dinner', 0], ['check','daily', 0]],
  23: [['check','nightshift', 0]]
};

// night-shift six-week clients (SW_Enrollments Rhythm Category = night-shift):
// their day is flipped, so meal cues remind at 5pm ("their morning") and check
// at 11pm instead of the standard day windows.
function cwdNightShiftSet_() {
  var set = {};
  try {
    var ss = SpreadsheetApp.openById('1dwUGqqYqgAjQYyYMeFkA5UEVsCFTH01S89z93qNd1ys');
    var en = (typeof swcTable_ === 'function') ? swcTable_(ss, 'SW_Enrollments') : null;
    if (!en) return set;
    var cN = en.col('Client Name'), cR = en.col('Rhythm Category'), cSt = en.col('Status');
    for (var i = 0; i < en.rows.length; i++) {
      if (String(en.rows[i][cSt] || '').toLowerCase() !== 'active') continue;
      if (/night/i.test(String(en.rows[i][cR] || ''))) set[String(en.rows[i][cN]).toLowerCase().replace(/\s+/g, ' ').trim()] = true;
    }
  } catch (e) {}
  return set;
}

// ---- cue -> class -------------------------------------------------------
function cwdClassify_(cue) {
  var d = String(cue || '').trim();
  if (!d) return '';
  var low = d.toLowerCase();
  if (/scorecard/.test(low) && /week/.test(low))                 return 'weekly'; // not a daily habit
  if (/2 ?pm|2-3pm|2p-3p|afternoon snack|shake/.test(low))       return 'snack';
  if (/breakfast/.test(low))                                     return 'breakfast';
  if (/lunch/.test(low))                                         return 'lunch';
  if (/dinner/.test(low))                                        return 'dinner';
  if (/night|lights.?out|wind.?down|mindset|in bed|bed by/.test(low)) return 'night';
  if (/snack/.test(low))                                         return 'snack';
  if (/morning/.test(low))                                       return 'breakfast';
  return 'daily';
}
function cwdHasPhoto_(cue) { return /photo|picture|upload/i.test(String(cue || '')); }

// ---- copy ---------------------------------------------------------------
function cwdRemindMsg_(first, cls, photo) {
  var p = photo ? ' and send a photo of it with your DONE' : '';
  switch (cls) {
    case 'breakfast': return 'Hey ' + first + '! Reminder to get your breakfast in this morning. Text me DONE right after you eat' + p + '. -Coach Anthony';
    case 'lunch':     return 'Hey ' + first + '! Reminder to get your lunch in today. Text me DONE right after you eat' + p + '. -Coach Anthony';
    case 'dinner':    return 'Hey ' + first + '! Reminder to get your dinner in tonight. Text me DONE right after you eat' + p + '. -Coach Anthony';
    case 'snack':     return 'Hey ' + first + '! Reminder to get your afternoon snack in. Text me DONE right after' + p + '. -Coach Anthony';
    case 'night':     return 'Hey ' + first + '! Reminder to run your nighttime routine tonight. Text me DONE once you get it done' + p + '. -Coach Anthony';
    default:          return 'Hey ' + first + '! Reminder to get your daily habit in today. Text me DONE once you do' + p + '. -Coach Anthony';
  }
}
// short acknowledgment when a DONE is received, rotated by day so it doesn't repeat word-for-word
function cwdAckMsg_(first) {
  var opts = ['Logged it 💪', 'Got it, logged ✅', 'Seen and logged 💪', 'Logged ✅ keep stacking days', 'Got you, its logged ✅'];
  var d = new Date();
  var idx = (d.getDate() + first.length) % opts.length;
  return opts[idx];
}

function cwdNudgeMsg_(first, cls, photo) {
  switch (cls) {
    case 'breakfast': return 'Hey ' + first + '! Checking in on your breakfast, I didnt see your DONE come through today. Did you get it in? Reminder to text me DONE right after breakfast tomorrow' + (photo ? ', and send a photo of it with your DONE' : '') + '. -Coach Anthony';
    case 'lunch':     return 'Hey ' + first + '! Just reminding you about your DONE reply for lunch, I didnt get that earlier. What did you eat? Also a reminder to send DONE right after lunch tomorrow' + (photo ? ', with a photo of what you ate' : '') + '. -Coach Anthony';
    case 'dinner':    return 'Hey ' + first + '! I didnt see your DONE come through for dinner tonight. Did you get it in? Text me DONE right after you eat tomorrow' + (photo ? ', with a photo of the meal' : '') + '. -Coach Anthony';
    case 'snack':     return 'Hey ' + first + '! I didnt see your DONE for your afternoon snack today. Did you get it in? Text me DONE right after it tomorrow' + (photo ? ', with a photo' : '') + '. -Coach Anthony';
    case 'night':     return 'Hey ' + first + '! I didnt see your DONE come through for your nighttime routine last night. Did you get it done? Text me DONE tonight once you do. -Coach Anthony';
    default:          return 'Hey ' + first + '! I didnt see your DONE come through for your daily habit. Did you get it in? Text me DONE once you do. -Coach Anthony';
  }
}

// ---- Trainerize integration (2026-07-10) ---------------------------------
// DONE replies can now arrive in the TZ app thread, and photo cues are
// verified against actual meal-photo uploads (same endpoints hourlyMealScan
// uses). Reuses mpa_tzPost_ from **Meal Photo Tracking.js.

// name -> Trainerize userID from Accountability Log (col A = TZ id, col C = name, data from row 3)
function cwdTzIndex_(ss) {
  var idx = {};
  try {
    var al = ss.getSheetByName('Accountability Log');
    var v = al.getRange(3, 1, al.getLastRow() - 2, 3).getValues();
    for (var i = 0; i < v.length; i++) {
      var nm = String(v[i][2] || '').trim(), id = String(v[i][0] || '').trim();
      if (nm && id) idx[doneNameKey_(nm)] = id;
    }
  } catch (e) { Logger.log('[cwdTzIndex_] ' + e); }
  return idx;
}

// did the client text DONE in their Trainerize thread since sinceDate?
function cwdTzDoneSince_(tzUserId, sinceDate) {
  try {
    var th = mpa_tzPost_('message/getThreads', { userID: 3336777, view: 'byClient', clientID: Number(tzUserId), start: 0, count: 5 });
    if (th.code !== 200 || !th.data.threads || !th.data.threads.length) return false;
    var threadID = th.data.threads[0].threadID;
    var cnt = mpa_tzPost_('message/getMessages', { threadID: threadID, start: 0, count: 1 });
    var total = (cnt.code === 200 && cnt.data.totalCount) ? cnt.data.totalCount : 0;
    var fetch = 25, startOff = Math.max(0, total - fetch);
    var res = mpa_tzPost_('message/getMessages', { threadID: threadID, start: startOff, count: fetch });
    if (res.code !== 200 || !res.data.messages) return false;
    for (var i = 0; i < res.data.messages.length; i++) {
      var m = res.data.messages[i];
      var senderType = (m.sender && m.sender.type) ? m.sender.type : '';
      if (senderType === 'trainer' || senderType === 'coach') continue;
      var at = new Date(m.sentTime || m.createdAt || 0);
      if (sinceDate && at.getTime() < sinceDate.getTime() - 60000) continue;
      if (DONE_CFG.doneRegex.test(String(m.body || ''))) return true;
    }
  } catch (e) { Logger.log('[cwdTzDoneSince_] ' + e); }
  return false;
}

// has the client uploaded at least one meal photo today (TZ meal tracker)?
function cwdTzPhotoToday_(tzUserId, targetDate) {
  try {
    var dstr = Utilities.formatDate(targetDate || new Date(), CWD_TZ, 'yyyy-MM-dd');
    // getList first (returns 200 with empty array when nothing logged; /get alone 404s on empty days)
    var lst = mpa_tzPost_('dailyNutrition/getList', { userID: Number(tzUserId), startDate: dstr, endDate: dstr });
    if (lst.code !== 200 || !lst.data.nutrition || !lst.data.nutrition.length) return false;
    var res = mpa_tzPost_('dailyNutrition/get', { userID: Number(tzUserId), date: dstr });
    if (res.code !== 200 || !res.data.nutrition) return false;
    var day = res.data.nutrition;
    if (day.mealPhoto && day.mealPhoto.id) return true;
    var meals = day.meals || [];
    for (var i = 0; i < meals.length; i++) {
      if (meals[i].mealPhoto && meals[i].mealPhoto.id) return true;
    }
  } catch (e) { Logger.log('[cwdTzPhotoToday_] ' + e); }
  return false;
}

// did the client send a photo in their Trainerize message thread since sinceDate?
// covers members who snap a photo inside the TZ app chat rather than logging it
// in the nutrition tracker.
function cwdTzPhotoInThread_(tzUserId, sinceDate) {
  try {
    var th = mpa_tzPost_('message/getThreads', { userID: 3336777, view: 'byClient', clientID: Number(tzUserId), start: 0, count: 5 });
    if (th.code !== 200 || !th.data.threads || !th.data.threads.length) return false;
    var threadID = th.data.threads[0].threadID;
    var cnt = mpa_tzPost_('message/getMessages', { threadID: threadID, start: 0, count: 1 });
    var total = (cnt.code === 200 && cnt.data.totalCount) ? cnt.data.totalCount : 0;
    var fetch = 25, startOff = Math.max(0, total - fetch);
    var res = mpa_tzPost_('message/getMessages', { threadID: threadID, start: startOff, count: fetch });
    if (res.code !== 200 || !res.data.messages) return false;
    var imgPat = /\.(jpe?g|png|gif|webp|heic)(\?[^\s]*)?$/i;
    for (var i = 0; i < res.data.messages.length; i++) {
      var m = res.data.messages[i];
      var senderType = (m.sender && m.sender.type) ? m.sender.type : '';
      if (senderType === 'trainer' || senderType === 'coach') continue;
      var at = new Date(m.sentTime || m.createdAt || 0);
      if (sinceDate && at.getTime() < sinceDate.getTime() - 60000) continue;
      var atts = m.attachments || m.images || m.media || [];
      if (atts.length) return true;
      if (imgPat.test(String(m.body || ''))) return true;
    }
  } catch (e) { Logger.log('[cwdTzPhotoInThread_] ' + e); }
  return false;
}

// ---- GHL conversation fetch (lives HERE permanently; SW_swapScan + photo
// logging depend on it). Returns {inbound:[{text,at,attachments[]}],
// outbound:[{text,at}]} since sinceDate, or null on read failure.
// Attachment-only messages (e.g. a texted meal photo with no words) are KEPT.
function tpFetchConvo_(apiKey, contactId, sinceDate) {
  var base = doneGhlBaseUrl_(), locId = doneGhlLocationId_();
  var headers = { 'Authorization': 'Bearer ' + apiKey, 'Version': DONE_CFG.ghlReadVersion };
  var sUrl = base + '/conversations/search?locationId=' + encodeURIComponent(locId) + '&contactId=' + encodeURIComponent(contactId);
  var sResp = UrlFetchApp.fetch(sUrl, { headers: headers, muteHttpExceptions: true });
  if (sResp.getResponseCode() >= 400) return null;
  var convs = (JSON.parse(sResp.getContentText()).conversations) || [];
  if (convs.length === 0) return { inbound: [], outbound: [] };
  var mUrl = base + '/conversations/' + convs[0].id + '/messages';
  var mResp = UrlFetchApp.fetch(mUrl, { headers: headers, muteHttpExceptions: true });
  if (mResp.getResponseCode() >= 400) return null;
  var mData = JSON.parse(mResp.getContentText());
  var msgs = (mData.messages && mData.messages.messages) || mData.messages || [];
  var out = { inbound: [], outbound: [] };
  for (var i = 0; i < msgs.length; i++) {
    var m = msgs[i];
    var body = String(m.body || m.message || '').trim();
    var atts = m.attachments || [];
    // GHL MMS: photo URL sometimes arrives in body text instead of attachments[]
    if (!atts.length && /\bhttps?:\/\/\S+\.(?:jpe?g|png|gif|webp|heic)(?:\?[^\s]*)?\b/i.test(body)) {
      atts = [body.match(/\bhttps?:\/\/\S+\.(?:jpe?g|png|gif|webp|heic)(?:\?[^\s]*)?\b/i)[0]];
    }
    if (!body && !atts.length) continue;
    var added = m.dateAdded || m.dateUpdated || m.createdAt;
    var at = added ? new Date(added) : null;
    if (at && sinceDate && at.getTime() < sinceDate.getTime() - 60000) continue;
    var dir = String(m.direction || '').toLowerCase();
    var rec = { text: body, at: at, attachments: atts };
    if (dir === 'inbound') out.inbound.push(rec);
    else if (dir === 'outbound') out.outbound.push(rec);
  }
  return out;
}

// send a message INTO the member's Trainerize app thread (used for acks when
// the member replied in the app — answer them where they talked to us)
function cwdTzSendMsg_(tzUserId, text) {
  try {
    var th = mpa_tzPost_('message/getThreads', { userID: 3336777, view: 'byClient', clientID: Number(tzUserId), start: 0, count: 5 });
    if (th.code !== 200 || !th.data.threads || !th.data.threads.length) return { success: false, code: th.code, reason: 'no thread' };
    var threadID = th.data.threads[0].threadID;
    var res = mpa_tzPost_('message/reply', { userID: 3336777, threadID: threadID, body: text, type: 'text' });
    return { success: res.code >= 200 && res.code < 300, code: res.code, reason: res.code >= 400 ? JSON.stringify(res.data).slice(0, 100) : '' };
  } catch (e) { return { success: false, code: 0, reason: String(e) }; }
}

// texted meal photos via GHL since sinceDate (interim channel until members
// move back to Trainerize). Returns array of attachment URLs (photos only).
function cwdGhlPhotosSince_(apiKey, contactId, sinceDate) {
  try {
    var conv = tpFetchConvo_(apiKey, contactId, sinceDate);
    if (!conv) return [];
    var urls = [];
    for (var i = 0; i < conv.inbound.length; i++) {
      var atts = conv.inbound[i].attachments || [];
      for (var a = 0; a < atts.length; a++) {
        var u = String(atts[a]);
        if (/\.(pdf|mp4|mov|mp3|vcf|txt)(\?|$)/i.test(u)) continue;   // not photos
        urls.push(u);
      }
    }
    return urls;
  } catch (e) { Logger.log('[cwdGhlPhotosSince_] ' + e); return []; }
}

// ---- helpers ------------------------------------------------------------
function cwdHour_(d) { return Number(Utilities.formatDate(d, CWD_TZ, 'H')); }
function cwdDayKey_(d) { return Utilities.formatDate(d, CWD_TZ, 'yyyyMMdd'); }
function cwdShift_(d, days) { var x = new Date(d); x.setDate(x.getDate() + days); return x; }

// ---- entry point (THE hourly trigger target) ----------------------------
function CW_doneTick() {
  // six-week reset engine rides the same hourly trigger (day-locked inside; no-op with no active enrollments)
  try { if (typeof SW_dailyTick === 'function') SW_dailyTick(); } catch (e) { Logger.log('[SW piggyback] ' + e); }
  // silent meal-photo logging (no client feedback) — 2pm + 9pm snapshots
  try { if (typeof CW_photoWatchTick === 'function') CW_photoWatchTick(); } catch (e2) { Logger.log('[photoWatch piggyback] ' + e2); }
  var now = new Date();
  var jobs = CWD_SCHEDULE[cwdHour_(now)];
  if (!jobs) return 'no jobs for hour ' + cwdHour_(now);
  var res = [];
  for (var i = 0; i < jobs.length; i++) {
    try { res.push(CW_doneRun(jobs[i][0], jobs[i][1], jobs[i][2], false)); }
    catch (e) { Logger.log('[doneTick] ' + jobs[i][0] + '/' + jobs[i][1] + ' ERR ' + e); }
  }
  return res.join(' | ');
}

/**
 * action 'remind' | 'check'; cls = cue class; dayOffset 0=today, -1=yesterday
 */
function CW_doneRun(action, cls, dayOffset, dryRun) {
  dryRun = (dryRun === true);
  dayOffset = dayOffset || 0;
  var now = new Date();
  var target = cwdShift_(now, dayOffset);
  var props = PropertiesService.getScriptProperties();
  var runKey = 'cwd:' + action + ':' + cls + ':' + cwdDayKey_(target);
  if (!dryRun && action === 'remind' && props.getProperty(runKey)) return action + '/' + cls + ' already ran';

  var ss = SpreadsheetApp.openById('1dwUGqqYqgAjQYyYMeFkA5UEVsCFTH01S89z93qNd1ys');
  var dc = ss.getSheetByName('Daily DONE Check');
  var lr = dc.getLastRow();
  var W = Math.max(dc.getLastColumn(), 43);
  var vals = dc.getRange(2, 1, lr - 1, W).getValues();

  var emailIdx = doneBuildEmailIndex_(ss);
  var apiKey = doneGhlApiKey_();
  var dayCol = DONE_CFG.col.firstDay + (target.getDate() - 1);
  var since = doneMidnight_(target);
  var cwdNSCache_ = cwdNightShiftSet_();

  var R = { action: action, cls: cls, day: cwdDayKey_(target), sent: [], marked: [], skipped: [] };

  for (var i = 0; i < vals.length; i++) {
    var row = vals[i], sheetRow = i + 2;
    var name = String(row[1] || '').trim();
    if (!name) continue;
    if (String(row[0] || '').trim()) { continue; }                       // col A suppress
    var cue = String(row[3] || '').trim();
    if (!cue) continue;                                                  // no cue -> not in program
    var nameKey = name.toLowerCase().replace(/\s+/g, ' ').trim();
    var isNS = !!cwdNSCache_[nameKey];
    var rowCls = cwdClassify_(cue);
    if (cls === 'nightshift') {
      // flipped-day pass: night-shift clients with a MEAL cue
      if (!isNS || ['breakfast', 'lunch', 'snack', 'dinner'].indexOf(rowCls) < 0) continue;
    } else {
      if (rowCls !== cls) continue;
      // night-shift meal clients are handled by the 5pm/11pm nightshift pass, not day windows
      if (isNS && ['breakfast', 'lunch', 'snack', 'dinner'].indexOf(rowCls) >= 0) continue;
    }
    var notes = String(row[4] || '').toLowerCase();
    if (notes.indexOf('freeze') >= 0 || notes.indexOf('on hold') >= 0) { R.skipped.push(name + ' (freeze)'); continue; }

    var first = name.split(' ')[0];
    var photo = cwdHasPhoto_(cue);
    var email = emailIdx[doneNameKey_(name)] || '';
    if (!email) { R.skipped.push(name + ' (no email)'); continue; }
    var contactId = ''; try { contactId = doneFindContactByEmail_(email); } catch (e) {}
    if (!contactId) { R.skipped.push(name + ' (no GHL contact)'); continue; }

    var msgCls = (cls === 'nightshift') ? rowCls : cls;
    if (action === 'remind') {
      // Skip if already DONE today (e.g. completed habit before reminder time)
      var dayBox = String(row[dayCol - 1] || '').trim();
      if (dayBox === '✓') { R.skipped.push(name + ' (already DONE)'); continue; }
      var riCheck = doneGhlInboundSince_(apiKey, contactId, since);
      if (riCheck !== null && doneFindDone_(riCheck)) {
        if (!dryRun) {
          dc.getRange(sheetRow, dayCol).setValue('✓');
          dc.getRange(sheetRow, DONE_CFG.autoCols.lastDone).setValue(now);
          dc.getRange(sheetRow, DONE_CFG.autoCols.status).setValue('DONE already received (remind skipped)');
        }
        R.skipped.push(name + ' (already DONE)');
        continue;
      }
      if (!R._tzIdx) R._tzIdx = cwdTzIndex_(dc.getParent());
      var tzIdR = R._tzIdx[doneNameKey_(name)] || '';
      if (tzIdR && cwdTzDoneSince_(tzIdR, since)) {
        if (!dryRun) {
          dc.getRange(sheetRow, dayCol).setValue('✓');
          dc.getRange(sheetRow, DONE_CFG.autoCols.lastDone).setValue(now);
          dc.getRange(sheetRow, DONE_CFG.autoCols.status).setValue('DONE already received via TZ (remind skipped)');
        }
        R.skipped.push(name + ' (already DONE via TZ)');
        continue;
      }
      var rmsg = cwdRemindMsg_(first, msgCls, photo);
      if (dryRun) { R.sent.push(name + ' DRY :: ' + rmsg); continue; }
      var rr = doneSendGhlSms_(apiKey, contactId, rmsg);
      R.sent.push(name + ' HTTP ' + rr.code + (rr.success ? '' : ' FAIL ' + rr.reason));
      continue;
    }

    // action === 'check'
    var box = String(row[dayCol - 1] || '').trim();
    if (box === '✓' || box === '✗') { R.skipped.push(name + ' (already marked)'); continue; }

    var inbound = doneGhlInboundSince_(apiKey, contactId, since);
    if (inbound === null) { R.skipped.push(name + ' (GHL read failed)'); continue; }
    var hit = doneFindDone_(inbound);

    // Trainerize signals: DONE typed in the app thread counts too; photo cues
    // verified against actual meal-photo uploads. (tzIdx built once per run below)
    if (!R._tzIdx) R._tzIdx = cwdTzIndex_(dc.getParent());
    var tzId = R._tzIdx[doneNameKey_(name)] || '';
    var via = hit ? 'sms' : '';
    if (!hit && tzId && cwdTzDoneSince_(tzId, since)) { hit = { at: now }; via = 'trainerize'; }
    var photoOk = null;   // null = not a photo cue
    var photoVia = '';
    if (photo) {
      // Check all photo channels: TZ nutrition tracker, TZ chat thread, GHL SMS/MMS
      if (tzId && cwdTzPhotoToday_(tzId, target)) { photoOk = true; photoVia = 'tz'; }
      else if (tzId && cwdTzPhotoInThread_(tzId, since)) { photoOk = true; photoVia = 'tz-chat'; }
      else if (cwdGhlPhotosSince_(apiKey, contactId, since).length) { photoOk = true; photoVia = 'ghl'; }
      else photoOk = false;
      if (!hit && photoOk) { hit = { at: now }; via = 'photo upload'; }   // the photo itself proves the habit
    }

    if (hit) {
      var needPhotoNudge = (photo && photoOk === false);
      // reply in the channel THEY used: app reply -> app ack; text -> SMS ack
      var ackInApp = (via === 'trainerize' || (via === 'photo upload' && photoVia === 'tz')) && tzId;
      if (!dryRun) {
        dc.getRange(sheetRow, dayCol).setValue('✓');
        dc.getRange(sheetRow, DONE_CFG.autoCols.lastDone).setValue(hit.at || now);
        dc.getRange(sheetRow, DONE_CFG.autoCols.lastCheck).setValue(now);
        dc.getRange(sheetRow, DONE_CFG.autoCols.status).setValue('DONE received (' + cls + ' via ' + (via || 'sms') + (needPhotoNudge ? ', no photo' : '') + ')');
        // acknowledge; if their cue includes a photo and none was found, fold the ask into the ack
        var ackMsg = needPhotoNudge
          ? cwdAckMsg_(first) + ' One thing: I didnt see a meal photo come through — send one with your DONE next time.'
          : cwdAckMsg_(first);
        var ak = ackInApp ? cwdTzSendMsg_(tzId, ackMsg) : doneSendGhlSms_(apiKey, contactId, ackMsg);
        if (ackInApp && !ak.success) ak = doneSendGhlSms_(apiKey, contactId, ackMsg);   // app send failed -> fall back to SMS
        R.marked.push(name + ' ✓ via ' + (via || 'sms') + (needPhotoNudge ? ' (photo missing)' : '') + ' ack ' + (ackInApp ? 'in-app' : 'sms') + ' HTTP ' + ak.code);
      } else {
        R.marked.push(name + ' ✓ DRY via ' + (via || 'sms') + ' ack would go ' + (ackInApp ? 'in-app' : 'sms'));
      }
    } else {
      var nmsg = cwdNudgeMsg_(first, msgCls, photo);
      if (dryRun) { R.sent.push(name + ' DRY-NUDGE :: ' + nmsg); continue; }
      var nr = doneSendGhlSms_(apiKey, contactId, nmsg);
      if (nr.success) {
        dc.getRange(sheetRow, dayCol).setValue('✗');
        dc.getRange(sheetRow, DONE_CFG.autoCols.lastNudge).setValue(now);
        dc.getRange(sheetRow, DONE_CFG.autoCols.lastCheck).setValue(now);
        dc.getRange(sheetRow, DONE_CFG.autoCols.status).setValue('nudged (' + cls + ')');
        R.marked.push(name + ' ✗ nudged HTTP ' + nr.code);
      } else {
        R.skipped.push(name + ' send FAIL ' + nr.code + ' ' + nr.reason);
      }
    }
  }

  if (!dryRun && action === 'remind') props.setProperty(runKey, String(now));
  Logger.log('[' + (dryRun ? 'DRY' : 'LIVE') + '] ' + action + '/' + cls + ' day=' + R.day + ' sent=' + R.sent.length + ' marked=' + R.marked.length + ' skipped=' + R.skipped.length);
  return action + '/' + cls + ' sent=' + R.sent.length + ' marked=' + R.marked.length + ' skipped=' + R.skipped.length;
}

/** Write each member's cue class into col 38 (Nudge Timing) so it's visible. Read-only elsewhere. */
function CW_writeCueClasses() {
  var ss = SpreadsheetApp.openById('1dwUGqqYqgAjQYyYMeFkA5UEVsCFTH01S89z93qNd1ys');
  var dc = ss.getSheetByName('Daily DONE Check');
  var lr = dc.getLastRow();
  var vals = dc.getRange(2, 1, lr - 1, 43).getValues();
  var counts = {}, out = [];
  for (var i = 0; i < vals.length; i++) {
    var name = String(vals[i][1] || '').trim();
    var cls = name ? cwdClassify_(String(vals[i][3] || '')) : '';
    out.push([cls]);
    if (name && cls) counts[cls] = (counts[cls] || 0) + 1;
  }
  dc.getRange(2, 38, out.length, 1).setValues(out);
  Logger.log('cue classes: ' + JSON.stringify(counts));
  return JSON.stringify(counts);
}

/** Full dry preview of every window, sends nothing. */
function CW_donePreviewAll() {
  var order = [['remind','breakfast'],['check','breakfast'],['remind','lunch'],['check','lunch'],
               ['remind','snack'],['check','snack'],['remind','dinner'],['check','dinner'],
               ['remind','night'],['check','night'],['remind','daily'],['check','daily']];
  var res = [];
  for (var i = 0; i < order.length; i++) res.push(CW_doneRun(order[i][0], order[i][1], 0, true));
  var f = DriveApp.createFile('cw_done_preview.json', JSON.stringify(res, null, 1), 'application/json');
  Logger.log('PREVIEW FILE=' + f.getId() + '\n' + res.join('\n'));
  return f.getId();
}

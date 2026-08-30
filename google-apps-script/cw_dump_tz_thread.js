/*************************************************************************
 * CW — TRAINERIZE THREAD DUMP
 *
 * Read-only. Pulls a client's Trainerize message thread and writes it to a
 * Drive file so it can be reviewed outside the app. Sends nothing.
 *
 *   cwDump_TzThread('Jeff Broome')        // look him up by name
 *   cwDump_TzThread(12345678)             // or pass the Trainerize ID directly
 *   cwDump_TzThread('Jeff Broome', 90)    // last 90 days instead of the default 60
 *
 * If the name lookup fails, run cwFind_TzClient('Broome') to see every
 * matching client and their ID, then call this with the ID.
 *************************************************************************/

var CWTZ_SS_ID   = '1dwUGqqYqgAjQYyYMeFkA5UEVsCFTH01S89z93qNd1ys';
var CWTZ_TZ      = 'America/Chicago';
var CWTZ_TRAINER = 3336777;

// Tabs that may carry a Trainerize id -> name mapping. Scanned in order.
var CWTZ_LOOKUP_TABS = [
  'Accountability Log', 'TZ_API_Roster', 'Trainerize_Clients',
  'TZ_ClientList_Export', 'Client_Profiles'
];

/**
 * Search every client tab for a name fragment. Returns matches with their ids
 * so you can see who exists before dumping. Case-insensitive, partial match.
 */
function cwFind_TzClient(fragment) {
  var frag = String(fragment || '').toLowerCase().trim();
  if (!frag) { Logger.log('Pass a name fragment, e.g. cwFind_TzClient("Broome")'); return ''; }

  var ss = SpreadsheetApp.openById(CWTZ_SS_ID);
  var found = [];

  for (var t = 0; t < CWTZ_LOOKUP_TABS.length; t++) {
    var sh = ss.getSheetByName(CWTZ_LOOKUP_TABS[t]);
    if (!sh) continue;
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2) continue;

    var vals = sh.getRange(1, 1, Math.min(lastRow, 2000), lastCol).getValues();

    // locate the id column and name column from whichever row looks like a header
    var idCol = -1, nameCol = -1, headerRow = -1;
    for (var h = 0; h < Math.min(3, vals.length); h++) {
      for (var c = 0; c < vals[h].length; c++) {
        var head = String(vals[h][c] || '').toLowerCase();
        if (idCol < 0 && /(trainerize.*id|tz.*id|client.?id|^userid$|^id$)/.test(head)) { idCol = c; headerRow = h; }
        if (nameCol < 0 && /name/.test(head) && !/file|first|last/.test(head))          { nameCol = c; headerRow = h; }
      }
      if (idCol >= 0 && nameCol >= 0) break;
    }
    if (nameCol < 0) continue;

    for (var r = (headerRow < 0 ? 1 : headerRow + 1); r < vals.length; r++) {
      var nm = String(vals[r][nameCol] || '').trim();
      if (!nm || nm.toLowerCase().indexOf(frag) < 0) continue;
      found.push({
        tab: CWTZ_LOOKUP_TABS[t],
        row: r + 1,
        name: nm,
        id: idCol >= 0 ? String(vals[r][idCol] || '').replace(/\.0$/, '') : '(no id column)'
      });
    }
  }

  if (!found.length) {
    Logger.log('No client matching "' + fragment + '" in: ' + CWTZ_LOOKUP_TABS.join(', ') +
               '\nHe may not be in the sheet yet. Grab the id from the Trainerize URL ' +
               '(.../clients/<ID>/...) and pass it straight to cwDump_TzThread(<ID>).');
    return '';
  }

  var lines = found.map(function (f) { return f.tab + ' row ' + f.row + ' | ' + f.name + ' | id=' + f.id; });
  Logger.log('MATCHES for "' + fragment + '":\n' + lines.join('\n'));
  return lines.join('\n');
}

/**
 * Dump one client's Trainerize thread to a Drive file.
 * @param nameOrId  client name (looked up in the tabs) or the numeric Trainerize id
 * @param sinceDays how far back to include; default 60
 */
function cwDump_TzThread(nameOrId, sinceDays) {
  sinceDays = sinceDays || 60;
  var since = new Date(new Date().getTime() - sinceDays * 86400000);

  // ---- resolve the client id ----
  var tzId = '', label = String(nameOrId);
  if (typeof nameOrId === 'number' || /^\d+$/.test(String(nameOrId).trim())) {
    tzId = String(nameOrId).trim();
  } else {
    var ss = SpreadsheetApp.openById(CWTZ_SS_ID);
    try {
      var idx = cwdTzIndex_(ss);                       // Accountability Log index
      tzId = idx[doneNameKey_(String(nameOrId))] || '';
    } catch (e) { Logger.log('[index] ' + e); }

    if (!tzId) {                                       // fall back to a wider scan
      var hits = cwFind_TzClient(String(nameOrId).split(' ').pop());
      if (!hits) return 'Could not resolve "' + nameOrId + '" to a Trainerize id.';
      var m = hits.match(/id=(\d+)/);
      if (!m) return 'Found a name match but no id. See the log and pass the id directly.';
      tzId = m[1];
      Logger.log('Resolved via tab scan -> id ' + tzId);
    }
  }

  // ---- pull the thread ----
  var th = mpa_tzPost_('message/getThreads',
    { userID: CWTZ_TRAINER, view: 'byClient', clientID: Number(tzId), start: 0, count: 5 });
  if (th.code !== 200) return 'getThreads failed, HTTP ' + th.code + ' ' + JSON.stringify(th.data).slice(0, 200);
  if (!th.data.threads || !th.data.threads.length) return 'No Trainerize thread exists for ' + label + ' (id ' + tzId + ')';

  var threadID = th.data.threads[0].threadID;
  var cnt = mpa_tzPost_('message/getMessages', { threadID: threadID, start: 0, count: 1 });
  var total = (cnt.code === 200 && cnt.data.totalCount) ? cnt.data.totalCount : 0;

  // page backwards from the newest so long histories still work
  var out = [], fetch = 50, off = Math.max(0, total - fetch);
  while (off >= 0) {
    var res = mpa_tzPost_('message/getMessages', { threadID: threadID, start: off, count: fetch });
    if (res.code !== 200 || !res.data.messages) break;

    var reachedEnd = false;
    for (var i = 0; i < res.data.messages.length; i++) {
      var m = res.data.messages[i];
      var at = new Date(m.sentTime || m.createdAt || 0);
      if (at.getTime() < since.getTime()) { reachedEnd = true; continue; }
      var st = (m.sender && m.sender.type) ? m.sender.type : '';
      out.push({
        at: Utilities.formatDate(at, CWTZ_TZ, 'yyyy-MM-dd HH:mm'),
        from: (st === 'trainer' || st === 'coach') ? 'COACH' : 'CLIENT',
        body: String(m.body || ''),
        attachments: (m.attachments || m.images || m.media || []).length
      });
    }
    if (reachedEnd || off === 0) break;
    off = Math.max(0, off - fetch);
    Utilities.sleep(200);
  }

  out.sort(function (a, b) { return a.at < b.at ? -1 : 1; });

  var payload = {
    client: label,
    trainerizeId: tzId,
    threadId: threadID,
    windowDays: sinceDays,
    since: Utilities.formatDate(since, CWTZ_TZ, 'yyyy-MM-dd'),
    totalInThread: total,
    returned: out.length,
    lastClientMessage: (function () {
      for (var i = out.length - 1; i >= 0; i--) if (out[i].from === 'CLIENT') return out[i].at;
      return 'none in window';
    })(),
    messages: out
  };

  var safe = label.replace(/[^A-Za-z0-9]+/g, '_').toLowerCase();
  var f = DriveApp.createFile('cw_tz_thread_' + safe + '.json',
                              JSON.stringify(payload, null, 1), 'application/json');

  Logger.log('FILE=' + f.getId() + '\n' + label + ' (id ' + tzId + ') — ' +
             out.length + ' of ' + total + ' messages in the last ' + sinceDays + ' days' +
             '\nLast message from client: ' + payload.lastClientMessage);
  return f.getId();
}

const BADGE_IDS = ['kitchen_clear', 'pantry_clear', 'freezer_clear', 'garden_clear', 'kitchen_stars', 'pantry_stars', 'freezer_stars', 'garden_stars', 'kitchen_bonus', 'pantry_bonus', 'freezer_bonus'];   // v0.16.35
/* Kitchen Heist — play-log collector + playtest leaderboard (Google Apps Script) — v0.16
   Setup: Sheet → Extensions → Apps Script → paste → save → Deploy → Manage deployments → edit → New version → Deploy.
   POST: game events, one row each. GET ?lb=1: leaderboard (stars first, then level; one row per nickname; animal + costume for the icon), cached 60 s.
   GET ?check=name&id=: nickname taken by someone else?  ?restore=name: furthest level.  ?funnel=1: first-session funnel for BUILD (new players → first move → levels 1–5 → next-day return). */
const BUILD = 'v0.16';
const FUNNEL_KEY = 'heist';                                         // ?funnel=1&key=heist — change it here if it ever leaks                                               // one leaderboard / funnel per test build
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(['received', 'player', 'build', 'lang', 'event_time', 'event', 'level', 'ms', 'fails', 'extra']);
    const now = new Date(), rows = [];
    for (const ev of body.events || []) {
      const { t, e: name, lvl, ms, fails, ...rest } = ev;
      if (body.name) rest.name = body.name;                       // every row carries the current nickname
      rows.push([now, body.id, body.v, body.lang || '', new Date(t), name, lvl == null ? '' : lvl, ms == null ? '' : ms, fails == null ? '' : fails, JSON.stringify(rest)]);
      if (name === 'feedback') {                                   // v0.16.18: rating + message also go to their own tab, readable at a glance
        const ss = SpreadsheetApp.getActiveSpreadsheet(); let fb = ss.getSheetByName('feedback');
        if (!fb) { fb = ss.insertSheet('feedback'); fb.appendRow(['time', 'nickname', 'player', 'build', 'rating', 'message', 'level reached', 'animal', 'wants app']); fb.setColumnWidth(6, 420); }
        fb.appendRow([new Date(t), rest.name || '', body.id, body.v, rest.rating || '', rest.text || '', rest.lvl || '', rest.animal || '', rest.wantApp || '']);
      }
    }
    if (sheet.getMaxRows() < sheet.getLastRow() + rows.length + 50) sheet.insertRowsAfter(sheet.getMaxRows(), 2000);
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return ContentService.createTextOutput('ok');
  } catch (err) { return ContentService.createTextOutput('error: ' + err); }
  finally { lock.releaseLock(); }
}
function doGet(e) {
  if (!(e && e.parameter && (e.parameter.lb || e.parameter.restore || e.parameter.check || e.parameter.funnel))) return ContentService.createTextOutput('Kitchen Heist collector is running');
  if (e.parameter.funnel) {
    if (e.parameter.key !== FUNNEL_KEY) return ContentService.createTextOutput('{"error":"key"}').setMimeType(ContentService.MimeType.JSON);                                          // ?funnel=1: how the first session goes, per player, for the current build
    const body = JSON.stringify(funnel()), fcb = e.parameter.callback;
    if (fcb && /^[A-Za-z0-9_]+$/.test(fcb)) return ContentService.createTextOutput(fcb + '(' + body + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
  }
  const cache = CacheService.getScriptCache(), hit = cache.get('lb');
  const out = hit || JSON.stringify(leaderboard());
  if (!hit) cache.put('lb', out, 60);
  if (e.parameter.check) {                                           // ?check=<nickname>&id=<player>: is this name used by someone else?
    const want = String(e.parameter.check).trim().toLowerCase(), me = String(e.parameter.id || '');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0], n = sheet.getLastRow();
    let taken = false; const owner = {};
    if (n >= 2) for (const [id, , , , ev, , , , extra] of sheet.getRange(2, 2, n - 1, 9).getValues()) {
      if (ev === 'reset_progress') { for (const k in owner) if (owner[k] === String(id)) delete owner[k]; continue; }   // v0.16.38: an erased player's name is free again
      if (ev !== 'nickname') continue; let x = {}; try { x = JSON.parse(extra || '{}'); } catch (err) {}
      const nm = String(x.name || '').trim().toLowerCase(); if (nm) owner[nm] = String(id);
    }
    taken = !!owner[want] && owner[want] !== me;
    const body = JSON.stringify({ taken }), ccb = e.parameter.callback;
    if (ccb && /^[A-Za-z0-9_]+$/.test(ccb)) return ContentService.createTextOutput(ccb + '(' + body + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.restore) {                                         // ?restore=<nickname>: furthest level for that name, for a new device
    const want = String(e.parameter.restore).trim().toLowerCase();
    const row = leaderboard().find(r => String(r.name).trim().toLowerCase() === want) || null;
    const body = JSON.stringify(row ? { level: row.level } : { level: 0 }), rcb = e.parameter.callback;
    if (rcb && /^[A-Za-z0-9_]+$/.test(rcb)) return ContentService.createTextOutput(rcb + '(' + body + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
  }
  const cb = e.parameter.callback;                                   // JSONP: the game loads this as a <script>, so no cross-origin block
  if (cb && /^[A-Za-z0-9_]+$/.test(cb)) return ContentService.createTextOutput(cb + '(' + out + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}
function leaderboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0], n = sheet.getLastRow(); if (n < 2) return [];
  const vals = sheet.getRange(2, 2, n - 1, 9).getValues();      // player, build, lang, event_time, event, level, ms, fails, extra
  const P = {}, seen = new Set();
  for (const [id, build, lang, time, ev, lvl, ms, fails, extra] of vals) {
    const sig = [id, String(time), ev, lvl, ms, fails].join('|'); if (seen.has(sig)) continue; seen.add(sig);   // drop duplicate posts
    if (!id || String(build).indexOf(BUILD) < 0) continue;        // one leaderboard per test build
    if (ev === 'reset_progress') { delete P[id]; continue; }      // v0.16.38: Settings → Erase wipes that player from the leaderboard (rows stay in the Sheet)
    const p = P[id] || (P[id] = { id, name: '', level: 0, fails: 0, secs: 0, starsBy: {}, animal: 'raccoon', costume: '' });
    let x = {}; try { x = JSON.parse(extra || '{}'); } catch (err) {}
    if (x.name) p.name = x.name;
    if (['closet_pick', 'closet_wear', 'level_start', 'level_complete', 'nickname', 'daily_solved', 'feedback'].indexOf(ev) >= 0) {   // never from closet_tease (tapping a locked shadow)
      if (x.animal && ['raccoon', 'rabbit'].indexOf(x.animal) >= 0) p.animal = x.animal;
      if (x.costume != null && ['', 'chef', 'sack', 'snow', 'gardener'].indexOf(x.costume) >= 0) p.costume = x.costume;
    }
    if (ev === 'daily_complete') { const dd = new Date(time); if (!isNaN(dd)) { const ds = dd.toISOString().slice(0, 10); if (!p.daily || ds > p.daily) p.daily = ds; } }   // last day they solved the daily
    if (ev === 'badges' && x.list != null) p.badges = String(x.list).split(',').filter(b => BADGE_IDS.indexOf(b) >= 0);   // v0.16.35: latest badge list
    if (x.daily || x.garden) continue;                              // daily puzzles and the garden don't count toward the level race
    if (ev === 'level_complete') { if (!x.bonus) p.level = Math.max(p.level, Number(lvl) || 0); p.secs += (Number(ms) || 0) / 1000; const st = Number(x.stars) || 0; if (st > (p.starsBy[lvl] || 0)) p.starsBy[lvl] = st; }
    if (ev === 'level_fail' || (ev === 'restart' && x.rule === 4)) p.fails += 1;
  }
  const hid = hiddenSet_();                                          // v0.16.40: owner-hidden players (Sheet tab "hidden")
  const rows = Object.values(P).filter(p => p.level > 0 && !hid.has(String(p.id).toLowerCase()) && !hid.has(String(p.name || '').trim().toLowerCase())).map(p => ({ id: p.id, name: p.name || '🦝', level: p.level, stars: Object.values(p.starsBy).reduce((a, b) => a + b, 0), fails: p.fails, secs: Math.round(p.secs), animal: p.animal, costume: p.costume, daily: p.daily || '', badges: p.badges || [] }))
    .sort((a, b) => b.stars - a.stars || b.badges.length - a.badges.length || b.level - a.level || a.fails - b.fails || a.secs - b.secs);   // v0.16: stars first
  const byName = new Map(); for (const r of rows) { const k = String(r.name).trim().toLowerCase(); if (!byName.has(k)) byName.set(k, r); }   // one entry per nickname: the best one
  return [...byName.values()].slice(0, 100);
}

function funnel() {                                                 // first-session funnel for BUILD: counts of distinct players reaching each step
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0], n = sheet.getLastRow(); if (n < 2) return { build: BUILD, players: 0 };
  const vals = sheet.getRange(2, 2, n - 1, 9).getValues(), P = {};
  for (const [id, build, lang, time, ev, lvl, ms, fails, extra] of vals) {
    if (!id || String(build).indexOf(BUILD) < 0) continue;
    const p = P[id] || (P[id] = { days: new Set(), boot: false, newPlayer: false, firstMove: false, done: new Set(), nudged: false, rewound: false, help: false, closet: false, hints: 0, firstDay: '', finishedDay: '', dailyDays: new Set() });
    const day = new Date(time), dstr = isNaN(day) ? '' : day.toISOString().slice(0, 10); if (dstr) { p.days.add(dstr); if (!p.firstDay || dstr < p.firstDay) p.firstDay = dstr; }
    let x = {}; try { x = JSON.parse(extra || '{}'); } catch (err) {}
    if (ev === 'boot_ok') p.boot = true; if (ev === 'new_player') p.newPlayer = true; if (ev === 'first_move' && !x.daily) p.firstMove = true;
    if (ev === 'level_complete' && !x.daily && !x.bonus && !x.garden) { p.done.add(Number(lvl)); if (Number(lvl) === 40 && dstr && (!p.finishedDay || dstr < p.finishedDay)) p.finishedDay = dstr; }
    if ((ev === 'daily_solved' || ev === 'daily_open') && dstr) p.dailyDays.add(dstr); if (ev === 'nudge') p.nudged = true; if (ev === 'auto_rewind') p.rewound = true;
    if (ev === 'help_open') p.help = true; if (ev === 'closet_open') p.closet = true; if (ev === 'hint_used') p.hints += 1;
  }
  const visitors = Object.values(P), tried = visitors.filter(p => p.firstMove), all = visitors.filter(p => p.done.has(1)), c = f => all.filter(f).length, fresh = all.filter(p => p.newPlayer);
  // Gate 1 reading (rule set Oct 15): a PLAYER is someone who COMPLETED LEVEL 1. Anyone else (opened the link, tapped around, was shown the game on a friend's phone) is a visitor and stays out of every denominator.
  // Players who finished all 40 levels on their first day ran out of game, not interest: they count as returned only if they came back for the daily on a later day; otherwise they leave the next-day count entirely.
  const finishedDay1 = all.filter(p => p.finishedDay && p.finishedDay === p.firstDay), backForDaily = p => [...p.dailyDays].some(d => d !== p.firstDay);
  const regular = all.filter(p => !(p.finishedDay && p.finishedDay === p.firstDay)), eligible = regular.concat(finishedDay1.filter(backForDaily));
  const returned = regular.filter(p => p.days.size >= 2).length + finishedDay1.filter(backForDaily).length, finishersBack = finishedDay1.filter(p => p.days.size >= 2).length, finishersDaily = finishedDay1.filter(p => [...p.dailyDays].some(d => d !== p.firstDay)).length;
  const pct = (a, b) => b ? Math.round(100 * a / b) : null;
  return { build: BUILD, visitors: visitors.length, triedButNoLevel1: tried.length - all.length, players: all.length, newPlayers: fresh.length, visitorsWhoNeverPlayed: visitors.length - all.length, level1: c(p => p.done.has(1)), level2: c(p => p.done.has(2)), level3: c(p => p.done.has(3)), level4: c(p => p.done.has(4)), level5: c(p => p.done.has(5)), level10: c(p => p.done.has(10)), level20: c(p => p.done.has(20)), level40: c(p => p.done.has(40)),
    gate_tutorialPct: pct(c(p => p.done.has(2)), all.length), gate_tenLevelsPct: pct(c(p => p.done.has(10)), all.length),
    nextDay_eligible: eligible.length, nextDay_returned: returned, gate_nextDayPct: pct(returned, eligible.length),
    finishedDay1: finishedDay1.length, finishersCameBack: finishersBack, finishersCameBackForDaily: finishersDaily,
    nextDay_allPlayers_oldMethod: c(p => p.days.size >= 2), finishersLeftOut: finishedDay1.filter(p => !backForDaily(p)).length, nextDay_allVisitors_v011Method: visitors.filter(p => p.days.size >= 2).length,
    nudged: c(p => p.nudged), autoRewound: c(p => p.rewound), openedHelp: c(p => p.help), openedCloset: c(p => p.closet), hintsUsed: all.reduce((a, p) => a + p.hints, 0) };
}

/* Nightly backup. Run setupBackup() ONCE from the editor (Run → setupBackup, authorize): it installs a 3 AM trigger that copies the
   log into a 'backup-YYYY-MM-DD' sheet, keeping the last 7. If rows ever vanish or junk appears, copy a backup tab back. */
function setupBackup() { ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'nightlyBackup').forEach(t => ScriptApp.deleteTrigger(t)); ScriptApp.newTrigger('nightlyBackup').timeBased().everyDays(1).atHour(3).create(); }
function nightlyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), src = ss.getSheets()[0], name = 'backup-' + new Date().toISOString().slice(0, 10);
  const old = ss.getSheetByName(name); if (old) ss.deleteSheet(old);
  const copy = src.copyTo(ss).setName(name); ss.setActiveSheet(ss.getSheets()[0]);
  const backups = ss.getSheets().filter(sh => sh.getName().indexOf('backup-') === 0).sort((a, b) => a.getName() < b.getName() ? -1 : 1);
  while (backups.length > 7) ss.deleteSheet(backups.shift());
}


/* ---------- v0.16.40 owner tools: Sheet menu "Kitchen Heist" (only you see it; the public game link cannot call these) ----------
   Hide   = off the leaderboard, nothing deleted (undo with Unhide or by removing the line in the "hidden" tab).
   Delete = removes every row of that player from the log, the feedback tab and the nightly backup tabs; a copy goes to "deleted-backup".
   Their phone keeps its own save: if they play again they come back with only the progress made after the delete. */
const OWNER_EMAIL = 'yasuharu.to@gmail.com';   // optional: put your Google email here, e.g. 'you@gmail.com'; then nobody else can use these tools even if you share the Sheet

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Kitchen Heist')
    .addItem('Hide a player from the leaderboard…', 'ownerHide')
    .addItem('Unhide a player…', 'ownerUnhide')
    .addSeparator()
    .addItem('Delete a player completely…', 'ownerDelete')
    .addToUi();
}
function ownerOk_() {
  const me = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (OWNER_EMAIL && me !== OWNER_EMAIL.toLowerCase()) { SpreadsheetApp.getUi().alert('Only the game owner can use this.'); return false; }
  return true;
}
function hiddenSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let h = ss.getSheetByName('hidden');
  if (!h) { h = ss.insertSheet('hidden', ss.getSheets().length); h.appendRow(['nickname or player id, one per row: hidden from the leaderboard (nothing is deleted)']); h.setColumnWidth(1, 520); }
  return h;
}
function hiddenSet_() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('hidden'), s = new Set();
  if (h && h.getLastRow() >= 2) h.getRange(2, 1, h.getLastRow() - 1, 1).getValues().forEach(([v]) => { const k = String(v || '').trim().toLowerCase(); if (k) s.add(k); });
  return s;
}
function ask_(title, text) {
  const ui = SpreadsheetApp.getUi(), r = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
  return r.getSelectedButton() === ui.Button.OK ? r.getResponseText().trim() : '';
}
function ownerHide() {
  if (!ownerOk_()) return; const v = ask_('Hide a player', 'Nickname (or player id) to hide from the leaderboard:'); if (!v) return;
  hiddenSheet_().appendRow([v]); CacheService.getScriptCache().remove('lb');
  SpreadsheetApp.getUi().alert(`"${v}" is now hidden from the leaderboard. Nothing was deleted.`);
}
function ownerUnhide() {
  if (!ownerOk_()) return; const v = ask_('Unhide a player', 'Nickname (or player id) to show again:'); if (!v) return;
  const h = hiddenSheet_(), want = v.toLowerCase(); let gone = 0;
  for (let r = h.getLastRow(); r >= 2; r--) if (String(h.getRange(r, 1).getValue()).trim().toLowerCase() === want) { h.deleteRow(r); gone++; }
  CacheService.getScriptCache().remove('lb'); SpreadsheetApp.getUi().alert(gone ? `"${v}" is back on the leaderboard.` : `"${v}" was not in the hidden list.`);
}
function playerIds_(want) {   // every player id that used this nickname (or the id itself)
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0], n = sheet.getLastRow(), ids = new Set(); if (n < 2) return ids;
  for (const [id, , , , , , , , extra] of sheet.getRange(2, 2, n - 1, 9).getValues()) {
    if (String(id).toLowerCase() === want) { ids.add(String(id)); continue; }
    let x = {}; try { x = JSON.parse(extra || '{}'); } catch (err) {}
    if (String(x.name || '').trim().toLowerCase() === want) ids.add(String(id));
  }
  return ids;
}
function removeRows_(sh, col, ids, backup) {   // delete rows whose player column is one of ids, copying them to backup first; returns the count
  const n = sh.getLastRow(); if (n < 2) return 0;
  const w = sh.getLastColumn(), vals = sh.getRange(2, 1, n - 1, w).getValues(), hit = [];
  vals.forEach((row, i) => { if (ids.has(String(row[col - 1]))) hit.push(i + 2); });
  if (!hit.length) return 0;
  hit.forEach(r => backup.appendRow([new Date(), sh.getName()].concat(vals[r - 2])));
  for (let k = hit.length - 1; k >= 0;) { let end = hit[k], start = end; while (k > 0 && hit[k - 1] === start - 1) { k--; start--; } sh.deleteRows(start, end - start + 1); k--; }
  return hit.length;
}
function ownerDelete() {
  if (!ownerOk_()) return; const ui = SpreadsheetApp.getUi();
  const v = ask_('Delete a player completely', 'Nickname (or player id) to delete. Every row of theirs is removed (a copy is kept in "deleted-backup"):'); if (!v) return;
  const ids = playerIds_(v.toLowerCase()); if (!ids.size) { ui.alert(`No player called "${v}" was found.`); return; }
  if (ui.alert('Delete player', `Delete everything for "${v}" (${ids.size} device id${ids.size > 1 ? 's' : ''}: ${[...ids].join(', ')})? This cannot be undone from the game.`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet(), lock = LockService.getScriptLock(); lock.waitLock(30000);   // same lock as the game's writes
  try {
    let bk = ss.getSheetByName('deleted-backup'); if (!bk) { bk = ss.insertSheet('deleted-backup', ss.getSheets().length); bk.appendRow(['deleted at', 'from tab', '…the original row']); }
    let total = removeRows_(ss.getSheets()[0], 2, ids, bk);                                  // play log: player = column B
    const fb = ss.getSheetByName('feedback'); if (fb) total += removeRows_(fb, 3, ids, bk);  // feedback tab: player = column C
    ss.getSheets().filter(sh => sh.getName().indexOf('backup-') === 0).forEach(sh => { total += removeRows_(sh, 2, ids, bk); });   // nightly backups too
    CacheService.getScriptCache().remove('lb');
    ui.alert(`Deleted ${total} rows for "${v}". A copy is in the "deleted-backup" tab — delete that tab too if this was a "please delete my data" request.`);
  } finally { lock.releaseLock(); }
}

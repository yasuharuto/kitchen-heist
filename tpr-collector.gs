/* Kitchen Heist — play-log collector + playtest leaderboard (Google Apps Script) — v0.16
   Setup: Sheet → Extensions → Apps Script → paste → save → Deploy → Manage deployments → edit → New version → Deploy.
   POST: game events, one row each. GET ?lb=1: leaderboard (stars first, then level; one row per nickname; animal + costume for the icon), cached 60 s.
   GET ?check=name&id=: nickname taken by someone else?  ?restore=name: furthest level.  ?funnel=1: first-session funnel for BUILD (new players → first move → levels 1–5 → next-day return). */
const BUILD = 'v0.16';                                               // one leaderboard / funnel per test build
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
    }
    if (sheet.getMaxRows() < sheet.getLastRow() + rows.length + 50) sheet.insertRowsAfter(sheet.getMaxRows(), 2000);
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return ContentService.createTextOutput('ok');
  } catch (err) { return ContentService.createTextOutput('error: ' + err); }
  finally { lock.releaseLock(); }
}
function doGet(e) {
  if (!(e && e.parameter && (e.parameter.lb || e.parameter.restore || e.parameter.check || e.parameter.funnel))) return ContentService.createTextOutput('Kitchen Heist collector is running');
  if (e.parameter.funnel) {                                          // ?funnel=1: how the first session goes, per player, for the current build
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
    let taken = false;
    if (n >= 2) for (const [id, , , , ev, , , , extra] of sheet.getRange(2, 2, n - 1, 9).getValues()) {
      if (ev !== 'nickname') continue; let x = {}; try { x = JSON.parse(extra || '{}'); } catch (err) {}
      if (String(x.name || '').trim().toLowerCase() === want && String(id) !== me) { taken = true; break; }
    }
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
    const p = P[id] || (P[id] = { id, name: '', level: 0, fails: 0, secs: 0, starsBy: {}, animal: 'raccoon', costume: '' });
    let x = {}; try { x = JSON.parse(extra || '{}'); } catch (err) {}
    if (x.name) p.name = x.name;
    if (x.animal) p.animal = x.animal; if (x.costume != null) p.costume = x.costume;   // latest animal / costume seen → leaderboard icon
    if (x.daily || x.garden) continue;                              // daily puzzles and the garden don't count toward the level race
    if (ev === 'level_complete') { if (!x.bonus) p.level = Math.max(p.level, Number(lvl) || 0); p.secs += (Number(ms) || 0) / 1000; const st = Number(x.stars) || 0; if (st > (p.starsBy[lvl] || 0)) p.starsBy[lvl] = st; }
    if (ev === 'level_fail' || (ev === 'restart' && x.rule === 4)) p.fails += 1;
  }
  const rows = Object.values(P).filter(p => p.level > 0).map(p => ({ id: p.id, name: p.name || '🦝', level: p.level, stars: Object.values(p.starsBy).reduce((a, b) => a + b, 0), fails: p.fails, secs: Math.round(p.secs), animal: p.animal, costume: p.costume }))
    .sort((a, b) => b.stars - a.stars || b.level - a.level || a.fails - b.fails || a.secs - b.secs);   // v0.16: stars first
  const byName = new Map(); for (const r of rows) { const k = String(r.name).trim().toLowerCase(); if (!byName.has(k)) byName.set(k, r); }   // one entry per nickname: the best one
  return [...byName.values()].slice(0, 100);
}

function funnel() {                                                 // first-session funnel for BUILD: counts of distinct players reaching each step
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0], n = sheet.getLastRow(); if (n < 2) return { build: BUILD, players: 0 };
  const vals = sheet.getRange(2, 2, n - 1, 9).getValues(), P = {};
  for (const [id, build, lang, time, ev, lvl, ms, fails, extra] of vals) {
    if (!id || String(build).indexOf(BUILD) < 0) continue;
    const p = P[id] || (P[id] = { days: new Set(), boot: false, newPlayer: false, firstMove: false, done: new Set(), nudged: false, rewound: false, help: false, closet: false, hints: 0 });
    const day = new Date(time); if (!isNaN(day)) p.days.add(day.toISOString().slice(0, 10));
    let x = {}; try { x = JSON.parse(extra || '{}'); } catch (err) {}
    if (ev === 'boot_ok') p.boot = true; if (ev === 'new_player') p.newPlayer = true; if (ev === 'first_move' && !x.daily) p.firstMove = true;
    if (ev === 'level_complete' && !x.daily && !x.bonus && !x.garden) p.done.add(Number(lvl)); if (ev === 'nudge') p.nudged = true; if (ev === 'auto_rewind') p.rewound = true;
    if (ev === 'help_open') p.help = true; if (ev === 'closet_open') p.closet = true; if (ev === 'hint_used') p.hints += 1;
  }
  const all = Object.values(P), c = f => all.filter(f).length, fresh = all.filter(p => p.newPlayer);
  return { build: BUILD, players: all.length, newPlayers: fresh.length, firstMove: c(p => p.firstMove), level1: c(p => p.done.has(1)), level2: c(p => p.done.has(2)), level3: c(p => p.done.has(3)), level4: c(p => p.done.has(4)), level5: c(p => p.done.has(5)), level10: c(p => p.done.has(10)),
    nextDay: c(p => p.days.size >= 2), nudged: c(p => p.nudged), autoRewound: c(p => p.rewound), openedHelp: c(p => p.help), openedCloset: c(p => p.closet), hintsUsed: all.reduce((a, p) => a + p.hints, 0) };
}

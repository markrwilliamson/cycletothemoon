/**
 * CYCLE TO THE MOON
 * Strava -> Google Sheet ledger, with a self-serving dashboard.
 *
 * Design notes for future-you (or whoever is reading this in 2041):
 *
 *  - The 'rides' tab is the ledger. It is the only thing that matters.
 *    One row per Strava activity, keyed on activity_id. Append-only.
 *    It holds facts, never totals. Every total in this system is derived,
 *    so it can never drift out of step with the underlying data.
 *
 *  - 'adjustments' is for anything that is not a Strava activity: a ride
 *    you did that never got recorded, a correction, a pre-Strava estimate.
 *    Lifetime total = sum(rides) + sum(adjustments). Always auditable.
 *
 *  - 'config' holds every assumption. Nothing is hardcoded in this file
 *    that you might reasonably want to change over 20 years.
 *
 *  - 'sync_log' is the heartbeat. If the sync silently dies the way the
 *    Zapier version did in 2026, HEARTBEAT_DAYS will email you.
 *
 *  - No external libraries. No CDNs. Nothing that can rot or 404.
 *
 * SETUP: run setupMoonTracker() once from the Apps Script editor.
 */

// ---------------------------------------------------------------------------
// CONSTANTS
// ---------------------------------------------------------------------------

/**
 * Strava is migrating the API base URL. From 1 June 2027 the old host stops
 * working. When that happens, change this one line to:
 *   'https://www.api-v3.strava.com'
 * Everything else in this file already uses Bearer auth, which is the other
 * 2027 requirement, so that single edit is the whole migration.
 */
var STRAVA_API_BASE = 'https://www.strava.com/api/v3';
var STRAVA_OAUTH_URL = 'https://www.strava.com/oauth/token';

var SHEET_RIDES = 'rides';
var SHEET_ADJUSTMENTS = 'adjustments';
var SHEET_CONFIG = 'config';
var SHEET_YEARLY = 'yearly';
var SHEET_LOG = 'sync_log';

var RIDE_HEADERS = ['activity_id', 'date', 'year', 'name', 'distance_km',
                    'moving_time_s', 'elevation_m', 'sport_type', 'location'];

// ---------------------------------------------------------------------------
// ONE-TIME SETUP
// ---------------------------------------------------------------------------

function setupMoonTracker() {
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();

  ['STRAVA_CLIENT_ID', 'STRAVA_CLIENT_SECRET', 'STRAVA_REFRESH_TOKEN'].forEach(function (k) {
    if (!props.getProperty(k)) {
      throw new Error('Missing script property: ' + k +
        '. Add it under Project Settings > Script Properties before running setup.');
    }
  });

  ensureSheet_(ss, SHEET_RIDES, RIDE_HEADERS);
  ensureSheet_(ss, SHEET_ADJUSTMENTS, ['date', 'description', 'distance_km', 'moving_time_s']);
  ensureSheet_(ss, SHEET_YEARLY, ['year', 'rides', 'distance_km', 'hours', 'cumulative_km']);
  ensureSheet_(ss, SHEET_LOG, ['timestamp', 'status', 'added', 'updated', 'removed', 'total_km', 'message']);
  seedConfig_(ss);

  // Drop the default empty sheet if it is still hanging around.
  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);

  // Triggers. Removed first so setup is safe to re-run.
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncNewRides').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('reconcileRecent').timeBased().onMonthDay(1).atHour(4).create();
  ScriptApp.newTrigger('heartbeat').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  log_('setup', 0, 0, 0, null, 'Tabs and triggers created. Run fullBackfill() next.');
  try {
    SpreadsheetApp.getUi().alert(
      'Setup complete.\n\nTabs created and triggers installed.\n\n' +
      'Next: run fullBackfill() once to pull your entire Strava history.');
  } catch (e) { /* no UI when run from the script editor; harmless */ }
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#f0f0f0');
    sh.setFrozenRows(1);
  }
  return sh;
}

function seedConfig_(ss) {
  var sh = ss.getSheetByName(SHEET_CONFIG);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_CONFIG);
  var rows = [
    ['key', 'value', 'notes'],
    ['target_km', 384400, 'Mean distance to the moon'],
    ['date_of_birth', '1973-11-18', 'Used for the age projection'],
    ['start_year', 2013, 'First year of recorded riding'],
    ['sport_types', 'Ride,GravelRide,MountainBikeRide,VirtualRide',
      'Which Strava sport types count towards the goal. Comma separated.'],
    ['alert_email', Session.getEffectiveUser().getEmail(),
      'Where the heartbeat warning goes if the sync stops'],
    ['heartbeat_days', 7, 'Warn if no successful sync in this many days'],
    ['dashboard_title', 'Cycle to the Moon', 'Heading on the public dashboard'],
    ['', '', ''],
    ['PROJECTION MODEL', 'km per year', 'Age-tapered. Edit freely as the decades pass.'],
    ['from_age_52', 15000, 'Recent actual pace'],
    ['from_age_65', 10000, 'Easing off'],
    ['from_age_75', 5000, 'Still turning the pedals']
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f0f0f0');
  sh.getRange(10, 1, 1, 3).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 130); sh.setColumnWidth(3, 380);
  return sh;
}

function getConfig_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  var values = sh.getDataRange().getValues();
  var cfg = {}, bands = [];
  values.forEach(function (r) {
    var k = String(r[0]).trim();
    if (!k) return;
    if (k.indexOf('from_age_') === 0) {
      bands.push({ age: Number(k.replace('from_age_', '')), rate: Number(r[1]) });
    } else {
      cfg[k] = r[1];
    }
  });
  bands.sort(function (a, b) { return a.age - b.age; });
  cfg.bands = bands;
  cfg.sportSet = {};
  String(cfg.sport_types || 'Ride').split(',').forEach(function (s) {
    cfg.sportSet[s.trim()] = true;
  });
  return cfg;
}

// ---------------------------------------------------------------------------
// ONE-TIME AUTHORISATION
// These two helpers exist so you never need a terminal. Run getAuthUrl(),
// follow it, paste the code back, run exchangeAuthCode(). Done forever.
// ---------------------------------------------------------------------------

/** STEP A: run this, then open the URL it prints in the execution log. */
function getAuthUrl() {
  var id = PropertiesService.getScriptProperties().getProperty('STRAVA_CLIENT_ID');
  if (!id) throw new Error('Set STRAVA_CLIENT_ID in Project Settings > Script Properties first.');
  var url = 'https://www.strava.com/oauth/authorize' +
    '?client_id=' + id +
    '&response_type=code' +
    '&redirect_uri=http://localhost/exchange_token' +
    '&approval_prompt=force' +
    '&scope=activity:read_all';
  Logger.log('\n\nOpen this URL, click Authorize, then copy the "code=" value\n' +
             'out of the address bar of the page that fails to load:\n\n' + url + '\n\n');
  return url;
}

/**
 * STEP B: paste the code into a script property called STRAVA_AUTH_CODE,
 * then run this. It swaps the one-time code for a permanent refresh token
 * and stores it. The auth code is deleted afterwards; it is single use.
 */
function exchangeAuthCode() {
  var props = PropertiesService.getScriptProperties();
  var code = props.getProperty('STRAVA_AUTH_CODE');
  if (!code) throw new Error('Add STRAVA_AUTH_CODE to Script Properties first (from getAuthUrl).');

  var res = UrlFetchApp.fetch(STRAVA_OAUTH_URL, {
    method: 'post',
    payload: {
      client_id: props.getProperty('STRAVA_CLIENT_ID'),
      client_secret: props.getProperty('STRAVA_CLIENT_SECRET'),
      code: code,
      grant_type: 'authorization_code'
    },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Exchange failed (' + res.getResponseCode() + '): ' + res.getContentText() +
      '\n\nAuth codes are single use and expire quickly. Run getAuthUrl() again for a fresh one.');
  }
  var data = JSON.parse(res.getContentText());
  props.setProperty('STRAVA_REFRESH_TOKEN', data.refresh_token);
  props.deleteProperty('STRAVA_AUTH_CODE');
  Logger.log('Success. Refresh token stored. You can now run setupMoonTracker().');
  return 'ok';
}

// ---------------------------------------------------------------------------
// STRAVA API
// ---------------------------------------------------------------------------

function getAccessToken_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('strava_access_token');
  if (cached) return cached;

  var props = PropertiesService.getScriptProperties();
  var res = UrlFetchApp.fetch(STRAVA_OAUTH_URL, {
    method: 'post',
    payload: {
      client_id: props.getProperty('STRAVA_CLIENT_ID'),
      client_secret: props.getProperty('STRAVA_CLIENT_SECRET'),
      grant_type: 'refresh_token',
      refresh_token: props.getProperty('STRAVA_REFRESH_TOKEN')
    },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Strava token refresh failed (' + res.getResponseCode() + '): ' + res.getContentText());
  }
  var data = JSON.parse(res.getContentText());

  // Strava rotates the refresh token. Persist the new one or the next run fails.
  if (data.refresh_token) props.setProperty('STRAVA_REFRESH_TOKEN', data.refresh_token);
  cache.put('strava_access_token', data.access_token, 18000); // 5 hours
  return data.access_token;
}

function stravaGet_(path, params) {
  var qs = Object.keys(params || {}).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
  var url = STRAVA_API_BASE + path + (qs ? '?' + qs : '');

  for (var attempt = 0; attempt < 4; attempt++) {
    var res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + getAccessToken_() },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code === 200) return JSON.parse(res.getContentText());
    if (code === 429) { Utilities.sleep(60000); continue; }  // rate limited, wait it out
    if (code >= 500) { Utilities.sleep(3000 * (attempt + 1)); continue; }
    throw new Error('Strava API ' + code + ' on ' + path + ': ' + res.getContentText());
  }
  throw new Error('Strava API failed after retries: ' + path);
}

/** Pages through /athlete/activities. after/before are epoch seconds. */
function fetchActivities_(after, before) {
  var out = [], page = 1;
  while (page <= 60) {  // 12,000 activities of headroom
    var params = { per_page: 200, page: page };
    if (after) params.after = after;
    if (before) params.before = before;
    var batch = stravaGet_('/athlete/activities', params);
    if (!batch.length) break;
    out = out.concat(batch);
    if (batch.length < 200) break;
    page++;
  }
  return out;
}

function toRow_(a) {
  var d = new Date(a.start_date_local);
  return [
    String(a.id),
    Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'),
    d.getUTCFullYear(),
    a.name,
    Math.round((a.distance / 1000) * 100) / 100,
    a.moving_time,
    Math.round(a.total_elevation_gain || 0),
    a.sport_type || a.type,
    a.location_city || a.location_country || ''
  ];
}

// ---------------------------------------------------------------------------
// SYNC
// ---------------------------------------------------------------------------

/**
 * Pulls your entire Strava history and rebuilds the ledger from source.
 * Safe to re-run: it replaces the rides tab wholesale.
 */
function fullBackfill() {
  var cfg = getConfig_();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_RIDES);

  var all = fetchActivities_(null, null).filter(function (a) {
    return cfg.sportSet[a.sport_type || a.type];
  });
  all.sort(function (a, b) {
    return new Date(a.start_date_local) - new Date(b.start_date_local);
  });

  var rows = all.map(toRow_);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, RIDE_HEADERS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, RIDE_HEADERS.length).setValues(rows);

  var totals = rebuildDerived_();
  cacheDashboardModel_();
  publishQuietly_(true);
  log_('backfill', rows.length, 0, 0, totals.totalKm,
       'Rebuilt ledger from source: ' + rows.length + ' rides.');
  compareWithStrava_(totals.rideKm);
  return rows.length;
}

/** Daily. Appends anything newer than the last ride in the ledger. */
function syncNewRides() {
  try {
    var cfg = getConfig_();
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_RIDES);
    var existing = existingIds_(sh);

    var lastDate = sh.getLastRow() > 1
      ? new Date(sh.getRange(sh.getLastRow(), 2).getValue())
      : new Date(String(cfg.start_year || 2013) + '-01-01');
    // Overlap by three days so a late upload is never missed.
    var after = Math.floor(lastDate.getTime() / 1000) - 3 * 86400;

    var fresh = fetchActivities_(after, null).filter(function (a) {
      return cfg.sportSet[a.sport_type || a.type] && !existing[String(a.id)];
    });
    fresh.sort(function (a, b) {
      return new Date(a.start_date_local) - new Date(b.start_date_local);
    });

    if (fresh.length) {
      var rows = fresh.map(toRow_);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, RIDE_HEADERS.length).setValues(rows);
    }
    var totals = rebuildDerived_();
    cacheDashboardModel_();
    publishQuietly_(fresh.length > 0);
    log_('ok', fresh.length, 0, 0, totals.totalKm,
         fresh.length ? 'Added ' + fresh.length + ' ride(s).' : 'No new rides.');
    return fresh.length;
  } catch (err) {
    log_('error', 0, 0, 0, null, String(err));
    throw err;
  }
}

/**
 * Monthly. Re-reads the last 90 days and fixes anything you edited or
 * deleted in Strava after the fact, so the ledger stays true to source.
 */
function reconcileRecent() {
  try {
    var cfg = getConfig_();
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_RIDES);
    var windowStart = new Date();
    windowStart.setDate(windowStart.getDate() - 90);
    var afterEpoch = Math.floor(windowStart.getTime() / 1000);

    var live = {};
    fetchActivities_(afterEpoch, null).forEach(function (a) {
      if (cfg.sportSet[a.sport_type || a.type]) live[String(a.id)] = a;
    });

    var last = sh.getLastRow();
    if (last < 2) { log_('ok', 0, 0, 0, null, 'Nothing to reconcile.'); return; }
    var data = sh.getRange(2, 1, last - 1, RIDE_HEADERS.length).getValues();

    var kept = [], updated = 0, removed = 0, seen = {};
    data.forEach(function (row) {
      var id = String(row[0]);
      var rowDate = new Date(row[1]);
      if (rowDate < windowStart) { kept.push(row); return; }   // outside window, untouched
      seen[id] = true;
      if (!live[id]) { removed++; return; }                     // deleted in Strava
      var fresh = toRow_(live[id]);
      if (JSON.stringify(fresh) !== JSON.stringify(row.map(String).map(function (v, i) {
            return i === 2 || i === 4 || i === 5 || i === 6 ? Number(row[i]) : row[i];
          }))) updated++;
      kept.push(fresh);
    });
    Object.keys(live).forEach(function (id) {
      if (!seen[id]) { kept.push(toRow_(live[id])); }           // appeared late
    });
    kept.sort(function (a, b) { return new Date(a[1]) - new Date(b[1]); });

    sh.getRange(2, 1, last - 1, RIDE_HEADERS.length).clearContent();
    if (kept.length) sh.getRange(2, 1, kept.length, RIDE_HEADERS.length).setValues(kept);

    var totals = rebuildDerived_();
    cacheDashboardModel_();
    publishQuietly_(true);
    log_('ok', 0, updated, removed, totals.totalKm,
         'Reconciled last 90 days. ' + updated + ' updated, ' + removed + ' removed.');
    compareWithStrava_(totals.rideKm);
  } catch (err) {
    log_('error', 0, 0, 0, null, String(err));
    throw err;
  }
}

function existingIds_(sh) {
  var map = {};
  if (sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
    if (r[0]) map[String(r[0])] = true;
  });
  return map;
}

/**
 * Cross-checks the ledger against Strava and flags real drift.
 *
 * Note for future-you: the API's all_ride_totals is NOT the same number Strava
 * shows as All-Time Distance on your profile page. There is a constant offset
 * between them at Strava's end (about 686 km when this was built, while the
 * ledger matched the profile page exactly, to the decimal).
 *
 * So comparing against it directly produces a permanent false alarm. Instead we
 * record that offset once, then only shout if it MOVES, which is what genuine
 * drift looks like: a ride edited or deleted in Strava after we logged it.
 *
 * If you deliberately change sport_types in config and re-run fullBackfill, the
 * offset legitimately changes. Delete the STRAVA_STATS_OFFSET script property
 * to re-baseline it.
 */
function compareWithStrava_(rideKm) {
  try {
    var profile = stravaGet_('/athlete', {});
    var stats = stravaGet_('/athletes/' + profile.id + '/stats', {});
    var stravaKm = (stats.all_ride_totals.distance || 0) / 1000;
    var diff = Math.round((rideKm - stravaKm) * 10) / 10;

    var props = PropertiesService.getScriptProperties();
    var baseline = props.getProperty('STRAVA_STATS_OFFSET');

    if (baseline === null) {
      props.setProperty('STRAVA_STATS_OFFSET', String(diff));
      log_('note', 0, 0, 0, null,
        'Baselined the known ledger-to-Strava-API offset at ' + diff +
        ' km. Future runs flag only if this moves.');
      return;
    }

    var drift = Math.round((diff - Number(baseline)) * 10) / 10;
    if (Math.abs(drift) > 5) {
      log_('note', 0, 0, 0, null,
        'Ledger has drifted ' + drift + ' km against Strava since baseline. ' +
        'Most likely a ride edited or deleted in Strava. Run reconcileRecent().');
    }
  } catch (e) { /* non-critical */ }
}

// ---------------------------------------------------------------------------
// DERIVED DATA
// ---------------------------------------------------------------------------

function readLedger_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_RIDES);
  var rides = [];
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, RIDE_HEADERS.length).getValues().forEach(function (r) {
      if (!r[0]) return;
      rides.push({ date: new Date(r[1]), year: Number(r[2]), name: r[3],
                   km: Number(r[4]) || 0, secs: Number(r[5]) || 0 });
    });
  }
  var adj = [];
  var ash = ss.getSheetByName(SHEET_ADJUSTMENTS);
  if (ash && ash.getLastRow() > 1) {
    ash.getRange(2, 1, ash.getLastRow() - 1, 4).getValues().forEach(function (r) {
      if (r[0] === '' && !r[2]) return;
      adj.push({ date: new Date(r[0]), km: Number(r[2]) || 0, secs: Number(r[3]) || 0 });
    });
  }
  return { rides: rides, adjustments: adj };
}

/**
 * Pure aggregation: no sheet reads, no writes. Kept separate so a page view
 * never triggers a spreadsheet write.
 */
function summarise_(led) {
  var byYear = {};
  led.rides.forEach(function (r) {
    if (!byYear[r.year]) byYear[r.year] = { rides: 0, km: 0, secs: 0 };
    byYear[r.year].rides++; byYear[r.year].km += r.km; byYear[r.year].secs += r.secs;
  });
  led.adjustments.forEach(function (a) {
    var y = a.date.getFullYear();
    if (!byYear[y]) byYear[y] = { rides: 0, km: 0, secs: 0 };
    byYear[y].km += a.km; byYear[y].secs += a.secs;
  });

  var years = Object.keys(byYear).map(Number).sort(function (a, b) { return a - b; });
  var rideKm = led.rides.reduce(function (s, r) { return s + r.km; }, 0);
  var adjKm = led.adjustments.reduce(function (s, r) { return s + r.km; }, 0);
  return { rideKm: rideKm, totalKm: rideKm + adjKm, years: years, byYear: byYear };
}

/** Writes the yearly tab. Sync path only, never a page view. */
function rebuildDerived_(led) {
  led = led || readLedger_();
  var d = summarise_(led);

  var cum = 0, rows = [];
  d.years.forEach(function (y) {
    cum += d.byYear[y].km;
    rows.push([y, d.byYear[y].rides, Math.round(d.byYear[y].km),
               Math.round(d.byYear[y].secs / 3600), Math.round(cum)]);
  });

  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_YEARLY);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);

  return d;
}

/** Age-tapered projection. Walks forward at the rate for each age band. */
function projectArrival_(remainingKm, currentAge, bands) {
  if (remainingKm <= 0) return { years: 0, age: currentAge };
  if (!bands.length) return null;
  var age = currentAge, km = remainingKm, years = 0, guard = 0;
  while (km > 0 && guard++ < 5000) {
    var rate = 0;
    for (var i = 0; i < bands.length; i++) if (age >= bands[i].age) rate = bands[i].rate;
    if (rate <= 0) rate = bands[0].rate;          // below the first band, use the first rate
    if (rate <= 0) return null;                    // never arrives
    var step = 0.05;
    if (rate * step >= km) { years += km / rate; age += km / rate; km = 0; }
    else { km -= rate * step; years += step; age += step; }
  }
  return km > 0 ? null : { years: years, age: age };
}

/**
 * Builds the dashboard's data model. Reads the ledger ONCE and writes nothing.
 * Called after a sync, not on a page view. See cacheDashboardModel_.
 */
function buildDashboardModel_() {
  var cfg = getConfig_();
  var led = readLedger_();
  var d = summarise_(led);

  var target = Number(cfg.target_km) || 384400;
  var total = d.totalKm;
  var remaining = Math.max(0, target - total);
  var pct = total / target * 100;

  var dob = new Date(cfg.date_of_birth);
  var now = new Date();
  var age = (now - dob) / (365.2425 * 86400000);

  var startYear = Number(cfg.start_year) || d.years[0] || now.getFullYear();
  var elapsedYears = Math.max(1, (now.getFullYear() + (now.getMonth() / 12)) - startYear);
  var avgPerYear = total / elapsedYears;

  // Trailing 365 days, a fairer read of current form than a calendar year.
  var cutoff = new Date(now.getTime() - 365 * 86400000);
  var last365 = led.rides.reduce(function (s, r) { return r.date >= cutoff ? s + r.km : s; }, 0);

  var proj = projectArrival_(remaining, age, cfg.bands);

  var years = d.years.map(function (y) {
    return { year: y, km: Math.round(d.byYear[y].km), rides: d.byYear[y].rides };
  });

  var lastRide = led.rides.length ? led.rides[led.rides.length - 1] : null;
  var lastSync = lastSyncInfo_();

  // Most recent ten, newest first, for the list at the foot of the page.
  var recent = led.rides.slice(-10).reverse().map(function (r) {
    return {
      date: Utilities.formatDate(r.date, 'Europe/London', 'EEE d MMM'),
      year: r.date.getFullYear(),
      name: r.name,
      km: r.km,
      mins: Math.round(r.secs / 60)
    };
  });

  return {
    title: cfg.dashboard_title || 'Cycle to the Moon',
    target: target, total: total, remaining: remaining, pct: pct,
    age: age, avgPerYear: avgPerYear, last365: last365,
    bands: cfg.bands,
    arrivalYears: proj ? proj.years : null,
    arrivalAge: proj ? proj.age : null,
    arrivalYear: proj ? now.getFullYear() + proj.years : null,
    startYear: startYear,
    years: years,
    rideCount: led.rides.length,
    recent: recent,
    lastRide: lastRide ? { date: Utilities.formatDate(lastRide.date, 'Europe/London', 'EEE d MMM yyyy'),
                           name: lastRide.name, km: lastRide.km } : null,
    lastSync: lastSync,
    generated: Utilities.formatDate(now, 'Europe/London', 'EEE d MMM yyyy, HH:mm')
  };
}

// ---------------------------------------------------------------------------
// LOGGING AND HEARTBEAT
// ---------------------------------------------------------------------------

function log_(status, added, updated, removed, totalKm, message) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  if (!sh) return;
  sh.appendRow([new Date(), status, added, updated, removed,
                totalKm == null ? '' : Math.round(totalKm), message]);
  // Keep the log to a sane length. 500 entries is years of history.
  if (sh.getLastRow() > 520) sh.deleteRows(2, sh.getLastRow() - 500);
}

function lastSyncInfo_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return null;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i][1] === 'ok' || rows[i][1] === 'backfill') {
      return { when: Utilities.formatDate(new Date(rows[i][0]), 'Europe/London', 'EEE d MMM yyyy, HH:mm'),
               daysAgo: Math.floor((new Date() - new Date(rows[i][0])) / 86400000) };
    }
  }
  return null;
}

/**
 * The thing the Zapier setup never had. If the sync quietly dies,
 * this tells you within a week instead of within a season.
 */
function heartbeat() {
  var cfg = getConfig_();
  var props = PropertiesService.getScriptProperties();
  var limit = Number(cfg.heartbeat_days) || 7;
  var email = cfg.alert_email || Session.getEffectiveUser().getEmail();
  var url = SpreadsheetApp.getActive().getUrl();
  var problems = [];

  // 1. Is Strava data still arriving?
  var info = lastSyncInfo_();
  if (!info || info.daysAgo > limit) {
    problems.push('Strava sync has not succeeded ' +
      (info ? 'since ' + info.when + ' (' + info.daysAgo + ' days ago).' : 'at all yet.') +
      ' Most likely the Strava refresh token needs renewing.');
  }

  // 2. Is the public page still being updated? A healthy sync with a dead
  //    publish freezes the dashboard silently, which is exactly how the
  //    Zapier version failed. Catch it separately.
  if (props.getProperty('GITHUB_TOKEN')) {
    var last = Number(props.getProperty('LAST_PUBLISH') || 0);
    var daysAgo = last ? Math.floor((Date.now() - last) / 86400000) : null;
    if (!last || daysAgo > limit) {
      problems.push('The public dashboard has not been published ' +
        (last ? 'since ' + daysAgo + ' days ago.' : 'at all yet.') +
        ' Most likely the GitHub token has expired. The sheet is still correct;' +
        ' only the public page is frozen.');
    }
  }

  if (problems.length) {
    MailApp.sendEmail(email, 'Moon tracker needs attention',
      problems.join('\n\n') + '\n\nSheet: ' + url + '\nCheck the sync_log tab.');
  }
}

// ---------------------------------------------------------------------------
// DASHBOARD WEB APP
// ---------------------------------------------------------------------------

var PROP_MODEL = 'DASHBOARD_MODEL';

/**
 * Recomputes the dashboard model and stores it as JSON. Called at the end of
 * every sync, and from the Moon menu when you change something in config.
 */
function cacheDashboardModel_() {
  var m = buildDashboardModel_();
  PropertiesService.getScriptProperties().setProperty(PROP_MODEL, JSON.stringify(m));
  return m;
}

/** Menu-friendly wrapper. */
function refreshDashboard() {
  var m = cacheDashboardModel_();
  log_('ok', 0, 0, 0, m.total, 'Dashboard snapshot refreshed.');
}

/**
 * Serves the dashboard from the stored snapshot rather than recomputing.
 *
 * The whole point: reading ~6,000 ledger rows and rewriting the yearly tab on
 * every page view made the page take seconds to appear. The data only changes
 * when a sync runs, so it is computed then and simply handed over here.
 * Falls back to computing on the spot if no snapshot exists yet.
 */
function doGet(e) {
  var props = PropertiesService.getScriptProperties();

  // Private one-tap sync. Bookmark the web app URL with ?sync=<SYNC_KEY> on
  // your phone. The key lives in Script Properties and never appears in the
  // published page, so the public dashboard stays a plain read-only file.
  var key = props.getProperty('SYNC_KEY');
  if (e && e.parameter && e.parameter.sync && key && e.parameter.sync === key) {
    var added = 0, err = null;
    try { added = syncNewRides(); } catch (ex) { err = String(ex); }
    return HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<div style="font:17px/1.6 system-ui,-apple-system,sans-serif;padding:40px 24px;' +
      'text-align:center;color:#191b24">' +
      (err
        ? '<h2>Sync failed</h2><p style="color:#a3352a">' + err + '</p>'
        : '<h2>Synced</h2><p>' +
          (added ? added + (added === 1 ? ' new ride added.' : ' new rides added.')
                 : 'No new rides found.') +
          '</p><p style="color:#5f6478;font-size:15px">The page updates within a minute.</p>' +
          '<p><a href="https://cycletothemoon.com/">Open the dashboard</a></p>') +
      '</div>');
  }

  var raw = props.getProperty(PROP_MODEL);
  var model;
  try {
    model = raw ? JSON.parse(raw) : cacheDashboardModel_();
  } catch (e) {
    model = cacheDashboardModel_();
  }
  var t = HtmlService.createTemplateFromFile('Dashboard');
  t.model = model;
  return t.evaluate()
    .setTitle(model.title || 'Cycle to the Moon')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------------------
// STATIC PUBLISHING
//
// The public dashboard is a plain HTML file on GitHub Pages, not this web app.
// Nobody's browser should have to talk to Google to read a page of numbers.
// Serving it statically means:
//   - it works on any device regardless of which Google account you are in,
//     which is what broke the web app version
//   - it loads instantly from a CDN, no cold start
//   - cycletothemoon.com points straight at it, so the short URL stays put
//   - it is a single self-contained file, which will outlive all of this
//
// doGet() above still works as a private preview. This is what people see.
// ---------------------------------------------------------------------------

/** Renders the dashboard to a standalone HTML string. */
function renderStaticHtml_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(PROP_MODEL);
  var model;
  try { model = raw ? JSON.parse(raw) : cacheDashboardModel_(); }
  catch (e) { model = cacheDashboardModel_(); }

  var t = HtmlService.createTemplateFromFile('Dashboard');
  t.model = model;
  return t.evaluate().getContent();
}

/**
 * Pushes the rendered page to GitHub Pages as index.html.
 * Needs two script properties:
 *   GITHUB_REPO   e.g. markwilliamson/cycletothemoon
 *   GITHUB_TOKEN  a fine-grained token with Contents: read and write on that repo
 */
function publishStaticPage() {
  var props = PropertiesService.getScriptProperties();
  var repo  = props.getProperty('GITHUB_REPO');
  var token = props.getProperty('GITHUB_TOKEN');
  if (!repo || !token) {
    throw new Error('Set GITHUB_REPO and GITHUB_TOKEN in Project Settings > Script Properties.');
  }

  var html = renderStaticHtml_();
  var api = 'https://api.github.com/repos/' + repo + '/contents/index.html';
  var headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };

  // Replacing an existing file requires its current sha.
  var sha = null;
  var probe = UrlFetchApp.fetch(api, { headers: headers, muteHttpExceptions: true });
  if (probe.getResponseCode() === 200) sha = JSON.parse(probe.getContentText()).sha;

  var body = {
    message: 'Dashboard update ' +
             Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd HH:mm'),
    content: Utilities.base64Encode(html, Utilities.Charset.UTF_8)
  };
  if (sha) body.sha = sha;

  var res = UrlFetchApp.fetch(api, {
    method: 'put', headers: headers, contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error('GitHub publish failed (' + code + '): ' + res.getContentText());
  }
  props.setProperty('LAST_PUBLISH', String(Date.now()));
  return true;
}

/**
 * Used in the sync path: never let a publishing hiccup lose a ride.
 *
 * The sync runs hourly, but publishing every hour would mean 24 commits a day
 * for the rest of the project, nearly all of them changing nothing but a
 * timestamp. So it publishes when there are new rides, and otherwise only if
 * the page has not been refreshed for half a day, which keeps the "synced at"
 * line honest without filling the repo with noise.
 */
function publishQuietly_(force) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty('GITHUB_TOKEN')) return;
    var last = Number(props.getProperty('LAST_PUBLISH') || 0);
    var stale = !last || (Date.now() - last) > 12 * 3600 * 1000;
    if (force || stale) publishStaticPage();
  } catch (e) {
    log_('error', 0, 0, 0, null, 'Static publish failed: ' + e);
  }
}

/** Menu-friendly: refresh the snapshot and push the page in one go. */
function refreshAndPublish() {
  var m = cacheDashboardModel_();
  publishStaticPage();
  log_('ok', 0, 0, 0, m.total, 'Dashboard refreshed and published.');
}

/** Handy menu inside the sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Moon')
    .addItem('Sync new rides now', 'syncNewRides')
    .addItem('Refresh and publish now', 'refreshAndPublish')
    .addItem('Reconcile last 90 days', 'reconcileRecent')
    .addSeparator()
    .addItem('Full rebuild from Strava', 'fullBackfill')
    .addToUi();
}

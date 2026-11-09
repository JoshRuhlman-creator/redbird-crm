/**
 * ============================================================================
 *  REDBIRD INVENTORY PIPELINE
 *  Product Allocation + Propagation  →  "Inventory Projections Mirror"
 *  (feeds the CRM Projected tab, Share Projected emails, and the retailer portal)
 * ============================================================================
 *
 *  WHAT IT DOES (every hour via syncAll(), automatically)
 *    1. Reads Product Allocation (Chris's sheet) and Propagation (Cultivation
 *       Overview) BY HEADER NAME. Columns can be moved or inserted in either
 *       sheet without breaking anything. If a required column disappears, the
 *       run stops, the last good data stays live, and you get an email.
 *    2. Merges them:
 *         - Product Allocation is authoritative for any harvest week it has
 *           (it carries Chris's 7g / 28g / wholesale / fresh-frozen decisions).
 *         - Every later week comes from Propagation, with units projected using
 *           the same math as Product Allocation and per-tote yields taken from
 *           the most recent Product Allocation harvests (self-calibrating).
 *         - Tote count uses the best number available:
 *           FM2 Actual → FM1 Actual → Flower Ready → Target.
 *    3. Validates before writing. It never replaces good data with an empty or
 *       broken set.
 *    4. Writes the mirror in ONE fixed layout (MIRROR_SCHEMA below). That layout
 *       is the contract the CRM and portal read. It does not change when the
 *       source sheets change.
 *    5. Records every run on a "Pipeline Status" tab (the CRM reads it for its
 *       freshness indicator) and a "Pipeline Log" tab (run history).
 *    6. Emails ALERT_EMAILS on failures or warnings (same issue repeats at most
 *       every 12 hours), and once more when things recover.
 *    7. A watchdog runs every 6 hours. If the hourly sync hasn't succeeded in
 *       3+ hours (trigger deleted, auth expired, quota), it emails you.
 *
 *  INSTALL (one time)
 *    1. Run  installInventoryPipeline()  once and approve the permissions.
 *       It confirms the hourly syncAll trigger, adds the watchdog, and runs
 *       a first sync.
 *    2. Check the "Pipeline Status" tab in the CRM Master sheet. It should say OK.
 *
 *  Everything in this file is self-contained. It does not depend on any other
 *  helper function in the project.
 * ============================================================================
 */

var INV_PIPELINE = {
  CRM_SHEET_ID: '1Nrz05Rr8B1-5-2HXIXByy_pL6TH9U6wz-HUgu9LSkjw',      // CRM Master
  MIRROR_TAB: 'Inventory Projections Mirror',
  STATUS_TAB: 'Pipeline Status',
  LOG_TAB: 'Pipeline Log',
  LOG_KEEP_ROWS: 1000,

  PA_SHEET_ID: '1IuSfAblyYWSO5gYZhdgL_KdaWSfxG7sjvBWv2zWURJ4',       // Product Allocation workbook
  PA_TAB: 'Product Allocation',

  PROP_SHEET_ID: '1UVNIp5DWbTabg21fGdABFP6KYAw4qNgU3M852LE6JhY',     // Cultivation Overview
  PROP_TAB: 'Propagation',

  ALERT_EMAILS: ['Josh.Ruhlman@redbird-cannabis.com'],
  ALERT_REPEAT_HOURS: 12,

  FLOWER_DAYS_AFTER_WEEK: 14,   // Week of → flower delivery (matches Product Allocation)
  JOINT_DAYS_AFTER_WEEK: 21,    // Week of → joint delivery
  MIN_RUNWAY_WEEKS: 6,          // warn if projected deliveries end sooner than this
  STALE_HOURS: 3,               // watchdog alerts if no successful sync in this long

  // Fallbacks used only if Product Allocation history can't provide a rate
  DEFAULT_P35_PER_TOTE: 70,
  DEFAULT_MICROS_PER_TOTE: 12,
  DEFAULT_GRAMS_PER_TOTE: 850,
  DEFAULT_TRIM: 0.35,
  RATE_LOOKBACK_WEEKS: 8
};

/**
 * The mirror layout. Column positions are the contract read by
 * sheets-loader.js (CRM) and index.html (portal). Don't reorder.
 * Add new fields only at the end.
 */
var MIRROR_SCHEMA = [
  'Week Of',            // 0
  'Grow Room',          // 1
  'Totes',              // 2
  'Plant Count',        // 3
  'Total Days',         // 4
  'Strain',             // 5
  'Flower Delivery',    // 6
  'Joint Delivery',     // 7
  'Proj 3.5g',          // 8   ← CRM p35 / portal p35
  'Grams Per Tote',     // 9
  'Actual 3.5g',        // 10  ← CRM a35
  '',                   // 11
  'Target 7g',          // 12
  'Target 28g',         // 13
  'Wholesale Tier 1',   // 14
  'Proj Micros 14g',    // 15  ← pm
  'Actual Micros 14g',  // 16
  'Wholesale Micros',   // 17
  'Post Grind/Sift g',  // 18
  'Trim %',             // 19
  '1g Joints',          // 20  ← j1
  '2pk Joints',         // 21  ← j2
  '10pk Joints',        // 22  ← j10
  'Infused',            // 23
  'Batch Count',        // 24
  'Total Proj 1g',      // 25
  'Total Proj 0.5g',    // 26
  'Fresh Frozen',       // 27  ← live resin flag
  'Cured',              // 28
  'Synced At',          // 29
  'Source'              // 30
];
var MC = {}; MIRROR_SCHEMA.forEach(function (h, i) { if (h) MC[h] = i; });

/**
 * Source headers, matched case- and punctuation-insensitively, with aliases
 * so routine renames ("Proj. 3.5g" → "Projected 3.5g") don't break the sync.
 * required: true → the run stops and alerts if the column can't be found.
 */
var PA_FIELDS = {
  week:      { names: ['Week of Date', 'Week of', 'Week'], required: true },
  room:      { names: ['Grow Room', 'Room'], required: true },
  totes:     { names: ['Totes'], required: true },
  plants:    { names: ['Plant Count', 'Plants'] },
  totalDays: { names: ['Total Days'] },
  strain:    { names: ['Strain', 'Strain Name'], required: true },
  fd:        { names: ['Prj. Flower Delivery Date', 'Proj Flower Delivery Date', 'Flower Delivery Date', 'Flower Delivery'] },
  jd:        { names: ['Prj. Joint Delivery Date', 'Proj Joint Delivery Date', 'Joint Delivery Date', 'Joint Delivery'] },
  perTote:   { names: ['Per Tote', 'Grams Per Tote'] },
  p35:       { names: ['Proj. 3.5g', 'Projected 3.5g', 'Proj 3.5g'], required: true },
  a35:       { names: ['3.5-g', 'Actual 3.5g', '3.5g'] },
  t7:        { names: ['Target 7g'] },
  t28:       { names: ['Target 28g'] },
  wt1:       { names: ['Wholesale Tier 1 Flower', 'Wholesale Tier 1'] },
  micros:    { names: ['Proj. Micros 14g', 'Projected Micros 14g', 'Proj Micros 14g'], required: true },
  aMicros:   { names: ['Micros 14g', 'Actual Micros 14g'] },
  wMicros:   { names: ['Wholesale Micros'] },
  grind:     { names: ['Post Grind/Sift', 'Post Grind Sift'] },
  trim:      { names: ['% Trim', 'Trim %', 'Trim'] },
  j1:        { names: ['1g', '1g Joints', '1g Singles'] },
  j2:        { names: ['2pk', '2pk Joints'] },
  j10:       { names: ['10pks', '10pk', '10pk Joints'], required: true },
  infused:   { names: ['Infused'] },
  batch:     { names: ['Batch Count'] },
  total1g:   { names: ['Total Proj. 1g', 'Total Proj 1g'] },
  totalHalf: { names: ['Total Proj. 0.5g', 'Total Proj 0.5g'] },
  ff:        { names: ['Fresh Frozen', 'Live Resin'] },
  cured:     { names: ['Cured', 'Cured Resin'] }
};

var PROP_FIELDS = {
  strain:       { names: ['Strain Name', 'Strain'], required: true },
  room:         { names: ['Grow Room', 'Room'], required: true },
  weekOf:       { names: ['Week of Date', 'Week of'], required: true },
  totalDays:    { names: ['Total Days'] },
  tgtTotes:     { names: ['Tgt Totes', 'Target Totes'], required: true },
  flowerTotes:  { names: ['Flower Ready Totes'] },
  flowerPlants: { names: ['Flower Ready Plants'] },
  fm1Totes:     { names: ['FM1 Actual Totes'] },
  fm2Totes:     { names: ['FM2 Actual Totes'] },
  kill:         { names: ['Kill'] },
  killDone:     { names: ['Kill Completed Date'] }
};

// ============================================================================
//  ENTRY POINTS
// ============================================================================

/**
 * One-time setup. Safe to re-run.
 *  - Makes sure the hourly syncAll() trigger exists (syncAll calls mirrorInventory)
 *  - Removes any stray standalone mirrorInventory triggers (avoids double runs)
 *  - Installs the 6-hour watchdog
 *  - Runs a first sync and reports the result
 */
function installInventoryPipeline() {
  var hasSyncAll = false;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'mirrorInventory' || fn === 'inventoryPipelineWatchdog') ScriptApp.deleteTrigger(t);
    else if (fn === 'syncAll') {
      if (hasSyncAll) ScriptApp.deleteTrigger(t);   // keep exactly one
      hasSyncAll = true;
    }
  });
  if (!hasSyncAll) ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('inventoryPipelineWatchdog').timeBased().everyHours(6).create();
  var result = mirrorInventory();
  Logger.log('Hourly syncAll trigger ' + (hasSyncAll ? 'already present' : 'created') +
             '; 6-hour watchdog installed. First run: ' + JSON.stringify(result));
}

/** Hourly sync. Called by syncAll() every hour. Same name as before, so anything already calling it keeps working. */
function mirrorInventory() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60 * 1000)) {
    Logger.log('Another inventory sync is already running. Skipping.');
    return { status: 'SKIPPED' };
  }
  var run = { started: new Date(), status: 'OK', warnings: [], errors: [], stats: {} };
  try {
    _invRun(run);
  } catch (e) {
    run.status = 'ERROR';
    run.errors.push('Unexpected error: ' + e.message);
    Logger.log(e.stack);
  } finally {
    try { _invRecord(run); } catch (e2) { Logger.log('Could not write status: ' + e2.message); }
    try { _invAlert(run); } catch (e3) { Logger.log('Could not send alert: ' + e3.message); }
    lock.releaseLock();
  }
  return { status: run.status, warnings: run.warnings, errors: run.errors, stats: run.stats };
}

/** Every 6 hours: alert if the hourly sync has stopped succeeding. */
function inventoryPipelineWatchdog() {
  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('INV_LAST_SUCCESS') || 0);
  var hours = last ? (Date.now() - last) / 36e5 : Infinity;
  var hasHourly = ScriptApp.getProjectTriggers().some(function (t) {
    var fn = t.getHandlerFunction();
    return fn === 'syncAll' || fn === 'mirrorInventory';
  });
  var problems = [];
  if (!hasHourly) problems.push('The hourly syncAll trigger is missing. Run installInventoryPipeline() to restore it.');
  if (hours > INV_PIPELINE.STALE_HOURS) {
    problems.push('No successful inventory sync in ' + (isFinite(hours) ? hours.toFixed(1) + ' hours' : 'ever') +
                  '. The CRM Projected tab and the retailer portal are showing stale projections.');
  }
  if (problems.length) {
    _invSendAlert('WATCHDOG', 'Inventory sync has stopped', problems, []);
  }
  _checkCultiveraOrderSync();
}

/**
 * Cultivera → OrdersLog sync (runs on the redbird-sync VM, 4x/day).
 * crossreference.js records every run on the "Cultivera Sync Status" tab.
 * Alerts when a run was BLOCKED / FAILED (it refused to overwrite history),
 * when a write was interrupted, or when no run has succeeded in 8+ hours
 * (VM stopped, billing lapsed, cron broken, Cultivera login failing).
 * Same problem emails at most every 12 hours.
 */
function _checkCultiveraOrderSync() {
  var sh = SpreadsheetApp.openById(INV_PIPELINE.CRM_SHEET_ID).getSheetByName('Cultivera Sync Status');
  if (!sh) return;   // not installed yet
  var s = {};
  sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 2).getValues().forEach(function (r) { if (r[0]) s[r[0]] = r[1]; });
  var problems = [];
  var status = String(s['Status'] || '');
  if (status && status !== 'OK') problems.push('Last run was ' + status + ': ' + (s['Message'] || '(no message)'));
  var last = s['Last Success'] ? new Date(s['Last Success']) : null;
  var hours = last && !isNaN(last) ? (Date.now() - last.getTime()) / 36e5 : Infinity;
  if (hours > 8) {
    problems.push('No successful Cultivera order sync in ' + (isFinite(hours) ? hours.toFixed(1) + ' hours' : 'ever') +
      '. Check that the redbird-sync VM (us-west1-a) is running and GCP billing is active. New orders are not reaching the CRM or Reports.');
  }
  var props = PropertiesService.getScriptProperties();
  if (!problems.length) {
    if (props.getProperty('CULT_ALERTING') === '1') {
      _invSendAlert('RESOLVED', 'Cultivera order sync is healthy again', ['Last success: ' + s['Last Success'] + '. ' + (s['Message'] || '')], []);
      props.deleteProperty('CULT_ALERTING'); props.deleteProperty('CULT_ALERT_SIG');
    }
    return;
  }
  var sig = problems.join('|').replace(/[\d.]+ hours/, 'N hours');
  if (sig === props.getProperty('CULT_ALERT_SIG') && Date.now() - Number(props.getProperty('CULT_ALERT_AT') || 0) < INV_PIPELINE.ALERT_REPEAT_HOURS * 36e5) return;
  _invSendAlert('ERROR', 'Cultivera order sync needs attention', problems, []);
  props.setProperty('CULT_ALERT_SIG', sig);
  props.setProperty('CULT_ALERT_AT', String(Date.now()));
  props.setProperty('CULT_ALERTING', '1');
}

// ============================================================================
//  CORE
// ============================================================================

function _invRun(run) {
  var tz = Session.getScriptTimeZone();
  var today = _invDay(new Date());
  var runwayDays = INV_PIPELINE.MIN_RUNWAY_WEEKS * 7;

  // ---- Read sources ----
  var pa = _invReadTable(INV_PIPELINE.PA_SHEET_ID, INV_PIPELINE.PA_TAB, PA_FIELDS, 'Product Allocation');
  var prop = _invReadTable(INV_PIPELINE.PROP_SHEET_ID, INV_PIPELINE.PROP_TAB, PROP_FIELDS, 'Propagation');
  [pa, prop].forEach(function (t) {
    if (t.error) run.errors.push(t.error);
    t.missingOptional.forEach(function (m) { run.warnings.push(t.label + ': optional column "' + m + '" not found. Using a default.'); });
  });
  if (run.errors.length) { run.status = 'ERROR'; return; }

  // ---- Product Allocation rows ----
  var paRows = [];
  pa.rows.forEach(function (r) {
    var w = _invDate(pa.get(r, 'week'));
    var strain = String(pa.get(r, 'strain') || '').trim();
    if (!w || !strain || /^no strains/i.test(strain)) return;
    var rec = {
      w: w, strain: strain,
      room: pa.get(r, 'room'), totes: _invNum(pa.get(r, 'totes')),
      plants: pa.get(r, 'plants'), totalDays: pa.get(r, 'totalDays'),
      fd: _invDate(pa.get(r, 'fd')) || _invAdd(w, INV_PIPELINE.FLOWER_DAYS_AFTER_WEEK),
      jd: _invDate(pa.get(r, 'jd')) || _invAdd(w, INV_PIPELINE.JOINT_DAYS_AFTER_WEEK),
      perTote: _invNum(pa.get(r, 'perTote')),
      p35: _invNum(pa.get(r, 'p35')), a35: _invNum(pa.get(r, 'a35')),
      t7: _invNum(pa.get(r, 't7')), t28: _invNum(pa.get(r, 't28')), wt1: _invNum(pa.get(r, 'wt1')),
      micros: _invNum(pa.get(r, 'micros')), aMicros: _invNum(pa.get(r, 'aMicros')), wMicros: _invNum(pa.get(r, 'wMicros')),
      grind: _invNum(pa.get(r, 'grind')), trim: _invPct(pa.get(r, 'trim')),
      j1: _invNum(pa.get(r, 'j1')), j2: _invNum(pa.get(r, 'j2')), j10raw: pa.get(r, 'j10'),
      infused: _invBool(pa.get(r, 'infused')), batch: _invNum(pa.get(r, 'batch')),
      total1g: _invNum(pa.get(r, 'total1g')), totalHalf: _invNum(pa.get(r, 'totalHalf')),
      ff: _invBool(pa.get(r, 'ff')), cured: _invBool(pa.get(r, 'cured')),
      source: 'Product Allocation'
    };
    rec.j10 = _invNum(rec.j10raw);
    paRows.push(rec);
  });

  // Rates calibrated from recent Product Allocation harvests
  var rates = _invRates(paRows);
  run.stats.rates = rates.summary;

  var paWeeks = {};
  paRows.forEach(function (r) { paWeeks[_invKey(r.w, tz)] = true; });

  var keep = [];
  paRows.forEach(function (r) {
    var latest = Math.max(r.w.getTime(), r.fd.getTime(), r.jd.getTime());
    if (latest < today.getTime()) return;
    if (/over\s*alloc/i.test(String(r.j10raw))) {
      run.warnings.push('Product Allocation: ' + r.strain + ' (week of ' + _invFmt(r.w, tz) + ') joints are OVER ALLOCATED. 10pk shown as 0.');
    }
    if (r.totes > 0 && r.p35 <= 0 && r.micros <= 0 && !r.ff) {
      run.warnings.push('Product Allocation: ' + r.strain + ' (week of ' + _invFmt(r.w, tz) + ') has ' + r.totes + ' totes but 0 projected 3.5g and Micros.');
    }
    keep.push(r);
  });
  run.stats.paRows = keep.length;

  // ---- Propagation rows for weeks Product Allocation doesn't cover ----
  var merged = {}, propStrainsByWeek = {};
  prop.rows.forEach(function (r) {
    var w = _invDate(prop.get(r, 'weekOf'));
    var strain = String(prop.get(r, 'strain') || '').trim();
    var room = String(prop.get(r, 'room') || '').trim();
    if (!w || !strain) return;
    if (_invBool(prop.get(r, 'kill')) || String(prop.get(r, 'killDone') || '').trim()) return;
    if (!room || /^n\/?a$/i.test(room)) return;
    var jd = _invAdd(w, INV_PIPELINE.JOINT_DAYS_AFTER_WEEK);
    if (jd < today) return;

    var totes = _invNum(prop.get(r, 'fm2Totes')) || _invNum(prop.get(r, 'fm1Totes')) ||
                _invNum(prop.get(r, 'flowerTotes')) || _invNum(prop.get(r, 'tgtTotes'));
    if (totes <= 0) return;

    var wk = _invKey(w, tz);
    (propStrainsByWeek[wk] = propStrainsByWeek[wk] || {})[strain.toLowerCase()] = strain;
    if (paWeeks[wk]) return;   // Product Allocation owns this week

    var k = wk + '|' + strain.toLowerCase();
    if (!merged[k]) merged[k] = { w: w, strain: strain, room: room, totes: 0, plants: 0, totalDays: prop.get(r, 'totalDays') };
    merged[k].totes += totes;
    merged[k].plants += _invNum(prop.get(r, 'flowerPlants'));
  });

  // Flag weeks where the two sheets disagree on strains (PA wins, but tell someone)
  var paStrainsByWeek = {};
  keep.forEach(function (r) { (paStrainsByWeek[_invKey(r.w, tz)] = paStrainsByWeek[_invKey(r.w, tz)] || {})[r.strain.toLowerCase()] = r.strain; });
  Object.keys(paStrainsByWeek).forEach(function (wk) {
    var ps = propStrainsByWeek[wk];
    if (!ps) return;
    var onlyProp = Object.keys(ps).filter(function (s) { return !paStrainsByWeek[wk][s]; }).map(function (s) { return ps[s]; });
    if (onlyProp.length) {
      run.warnings.push('Week of ' + wk + ': in Propagation but not Product Allocation → ' + onlyProp.join(', ') + '. Product Allocation is used for this week.');
    }
  });

  var propCount = 0;
  Object.keys(merged).forEach(function (k) {
    keep.push(_invProject(merged[k], rates));
    propCount++;
  });
  run.stats.propagationRows = propCount;

  // ---- Validate ----
  keep.sort(function (a, b) { return a.w - b.w || String(a.strain).localeCompare(String(b.strain)); });
  var futureJoint = keep.filter(function (r) { return r.jd >= today; });
  var weeks = {};
  futureJoint.forEach(function (r) { weeks[_invKey(r.w, tz)] = true; });
  run.stats.harvestWeeks = Object.keys(weeks).length;
  run.stats.totalRows = keep.length;
  var furthest = futureJoint.length ? new Date(Math.max.apply(null, futureJoint.map(function (r) { return r.jd.getTime(); }))) : null;
  run.stats.furthestDelivery = furthest ? _invFmt(furthest, tz) : '';

  if (!keep.length) {
    run.status = 'ERROR';
    run.errors.push('Both sources produced zero upcoming harvests. The mirror was NOT overwritten; the CRM and portal still show the last good data.');
    return;
  }
  if (!furthest || (furthest - today) / 864e5 < runwayDays) {
    run.warnings.push('Projected deliveries end ' + (furthest ? _invFmt(furthest, tz) : 'now') +
      ', less than ' + INV_PIPELINE.MIN_RUNWAY_WEEKS + ' weeks out. Add upcoming harvests to Propagation.');
  }

  // ---- Write (overwrite in place; never leaves the tab empty) ----
  var stamp = new Date();
  var values = [MIRROR_SCHEMA.slice()].concat(keep.map(function (r) { return _invToMirrorRow(r, stamp); }));
  var ss = SpreadsheetApp.openById(INV_PIPELINE.CRM_SHEET_ID);
  var sh = ss.getSheetByName(INV_PIPELINE.MIRROR_TAB) || ss.insertSheet(INV_PIPELINE.MIRROR_TAB);
  if (sh.getMaxColumns() < MIRROR_SCHEMA.length) sh.insertColumnsAfter(sh.getMaxColumns(), MIRROR_SCHEMA.length - sh.getMaxColumns());
  if (sh.getMaxRows() < values.length) sh.insertRowsAfter(sh.getMaxRows(), values.length - sh.getMaxRows());
  sh.getRange(1, 1, values.length, MIRROR_SCHEMA.length).setValues(values);
  var extra = sh.getLastRow() - values.length;
  if (extra > 0) sh.getRange(values.length + 1, 1, extra, sh.getMaxColumns()).clearContent();
  if (values.length > 1) {
    [MC['Week Of'], MC['Flower Delivery'], MC['Joint Delivery']].forEach(function (c) {
      sh.getRange(2, c + 1, values.length - 1, 1).setNumberFormat('M/d/yyyy');
    });
    sh.getRange(2, MC['Synced At'] + 1, values.length - 1, 1).setNumberFormat('M/d/yyyy h:mm am/pm');
  }
  sh.getRange(1, 1, 1, MIRROR_SCHEMA.length).setFontWeight('bold').setBackground('#f3f3f3');
  sh.setFrozenRows(1);
  SpreadsheetApp.flush();

  if (run.warnings.length && run.status === 'OK') run.status = 'WARN';
}

/** Project units for a Propagation harvest using Product Allocation's math. */
function _invProject(h, rates) {
  var trim = rates.trimByStrain[h.strain.toLowerCase()] || rates.trim;
  var grind = rates.perTote * h.totes * trim;
  return {
    w: h.w, strain: h.strain, room: h.room, totes: h.totes, plants: h.plants || '', totalDays: h.totalDays || '',
    fd: _invAdd(h.w, INV_PIPELINE.FLOWER_DAYS_AFTER_WEEK),
    jd: _invAdd(h.w, INV_PIPELINE.JOINT_DAYS_AFTER_WEEK),
    perTote: rates.perTote,
    p35: Math.round(h.totes * rates.p35), a35: 0, t7: 0, t28: 0, wt1: 0,
    micros: Math.round(h.totes * rates.micros), aMicros: 0, wMicros: 0,
    grind: Math.round(grind), trim: trim,
    j1: 0, j2: 0, j10: Math.round(grind * 0.75 / 5 / 10) * 10,
    infused: false, batch: 0, total1g: 0, totalHalf: Math.floor(grind * 0.75 / 5) * 10,
    ff: false, cured: false, source: 'Propagation'
  };
}

/** Per-tote yields from the last N weeks of Product Allocation. */
function _invRates(paRows) {
  var C = INV_PIPELINE;
  var rows = paRows.filter(function (r) { return r.totes > 0; }).sort(function (a, b) { return a.w - b.w; });
  var weekKeys = [];
  rows.forEach(function (r) { var k = r.w.getTime(); if (weekKeys.indexOf(k) < 0) weekKeys.push(k); });
  var recentSet = weekKeys.slice(-C.RATE_LOOKBACK_WEEKS);
  var recent = rows.filter(function (r) { return recentSet.indexOf(r.w.getTime()) >= 0 && !r.ff; });

  var p35 = [], micros = [];
  recent.forEach(function (r) {
    if (!r.t7 && !r.t28 && !r.wt1 && r.p35 > 0) p35.push(r.p35 / r.totes);
    if (!r.wMicros && r.micros > 0) micros.push(r.micros / r.totes);
  });
  var perTote = 0, trim = 0, trimByStrain = {};
  rows.forEach(function (r) {
    if (r.perTote > 0) perTote = r.perTote;
    if (r.trim > 0) { trim = r.trim; trimByStrain[r.strain.toLowerCase()] = r.trim; }
  });
  var out = {
    p35: _invMedian(p35) || C.DEFAULT_P35_PER_TOTE,
    micros: _invMedian(micros) || C.DEFAULT_MICROS_PER_TOTE,
    perTote: perTote || C.DEFAULT_GRAMS_PER_TOTE,
    trim: trim || C.DEFAULT_TRIM,
    trimByStrain: trimByStrain
  };
  out.summary = '3.5g/tote ' + out.p35 + ' · Micros/tote ' + out.micros + ' · g/tote ' + out.perTote +
                ' · trim ' + Math.round(out.trim * 100) + '% (from ' + recentSet.length + ' recent weeks)';
  return out;
}

function _invToMirrorRow(r, stamp) {
  var row = MIRROR_SCHEMA.map(function () { return ''; });
  row[MC['Week Of']] = r.w;
  row[MC['Grow Room']] = r.room;
  row[MC['Totes']] = r.totes;
  row[MC['Plant Count']] = r.plants;
  row[MC['Total Days']] = r.totalDays;
  row[MC['Strain']] = r.strain;
  row[MC['Flower Delivery']] = r.fd;
  row[MC['Joint Delivery']] = r.jd;
  row[MC['Proj 3.5g']] = r.p35;
  row[MC['Grams Per Tote']] = r.perTote;
  row[MC['Actual 3.5g']] = r.a35;
  row[MC['Target 7g']] = r.t7;
  row[MC['Target 28g']] = r.t28;
  row[MC['Wholesale Tier 1']] = r.wt1;
  row[MC['Proj Micros 14g']] = r.micros;
  row[MC['Actual Micros 14g']] = r.aMicros;
  row[MC['Wholesale Micros']] = r.wMicros;
  row[MC['Post Grind/Sift g']] = r.grind;
  row[MC['Trim %']] = r.trim;
  row[MC['1g Joints']] = r.j1;
  row[MC['2pk Joints']] = r.j2;
  row[MC['10pk Joints']] = r.j10;
  row[MC['Infused']] = r.infused ? 'TRUE' : 'FALSE';
  row[MC['Batch Count']] = r.batch;
  row[MC['Total Proj 1g']] = r.total1g;
  row[MC['Total Proj 0.5g']] = r.totalHalf;
  row[MC['Fresh Frozen']] = r.ff ? 'TRUE' : 'FALSE';
  row[MC['Cured']] = r.cured ? 'TRUE' : 'FALSE';
  row[MC['Synced At']] = stamp;
  row[MC['Source']] = r.source;
  return row;
}

// ============================================================================
//  STATUS, LOG, ALERTS
// ============================================================================

function _invRecord(run) {
  var tz = Session.getScriptTimeZone();
  var ss = SpreadsheetApp.openById(INV_PIPELINE.CRM_SHEET_ID);
  var props = PropertiesService.getScriptProperties();
  if (run.status === 'OK' || run.status === 'WARN') props.setProperty('INV_LAST_SUCCESS', String(Date.now()));
  var lastSuccess = Number(props.getProperty('INV_LAST_SUCCESS') || 0);

  var st = ss.getSheetByName(INV_PIPELINE.STATUS_TAB) || ss.insertSheet(INV_PIPELINE.STATUS_TAB);
  var s = run.stats;
  var rows = [
    ['Field', 'Value'],
    ['Status', run.status],
    ['Last Run', run.started],
    ['Last Success', lastSuccess ? new Date(lastSuccess) : ''],
    ['Product Allocation Rows', s.paRows || 0],
    ['Propagation Rows', s.propagationRows || 0],
    ['Harvest Weeks', s.harvestWeeks || 0],
    ['Furthest Delivery', s.furthestDelivery || ''],
    ['Yield Rates', s.rates || ''],
    ['Errors', run.errors.join('\n')],
    ['Warnings', run.warnings.join('\n')]
  ];
  st.clearContents();
  st.getRange(1, 1, rows.length, 2).setValues(rows);
  st.getRange('B3:B4').setNumberFormat('M/d/yyyy h:mm am/pm');
  st.getRange('A1:B1').setFontWeight('bold');
  st.getRange('B2').setBackground(run.status === 'OK' ? '#d9ead3' : run.status === 'WARN' ? '#fff2cc' : '#f4cccc');
  st.setColumnWidth(2, 700);
  st.getRange('B10:B11').setWrap(true);

  var lg = ss.getSheetByName(INV_PIPELINE.LOG_TAB) || ss.insertSheet(INV_PIPELINE.LOG_TAB);
  if (lg.getLastRow() === 0) {
    lg.appendRow(['Run At', 'Status', 'PA Rows', 'Propagation Rows', 'Harvest Weeks', 'Furthest Delivery', 'Seconds', 'Errors', 'Warnings']);
    lg.setFrozenRows(1);
  }
  lg.appendRow([run.started, run.status, s.paRows || 0, s.propagationRows || 0, s.harvestWeeks || 0,
                s.furthestDelivery || '', Math.round((Date.now() - run.started) / 1000),
                run.errors.join(' | '), run.warnings.join(' | ')]);
  var excess = lg.getLastRow() - 1 - INV_PIPELINE.LOG_KEEP_ROWS;
  if (excess > 0) lg.deleteRows(2, excess);
}

function _invAlert(run) {
  var props = PropertiesService.getScriptProperties();
  var wasAlerting = props.getProperty('INV_ALERTING') === '1';
  if (run.status === 'OK') {
    if (wasAlerting) {
      _invSendAlert('RESOLVED', 'Inventory sync is healthy again',
        ['Latest run succeeded: ' + (run.stats.harvestWeeks || 0) + ' harvest weeks, deliveries through ' + (run.stats.furthestDelivery || 'n/a') + '.'], []);
      props.deleteProperty('INV_ALERTING');
      props.deleteProperty('INV_ALERT_SIG');
    }
    return;
  }
  if (run.status === 'SKIPPED') return;
  // Same problems → email at most every ALERT_REPEAT_HOURS
  var sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, run.status + run.errors.join('|') + run.warnings.join('|')));
  var lastSig = props.getProperty('INV_ALERT_SIG');
  var lastAt = Number(props.getProperty('INV_ALERT_AT') || 0);
  if (sig === lastSig && Date.now() - lastAt < INV_PIPELINE.ALERT_REPEAT_HOURS * 36e5) return;

  _invSendAlert(run.status, run.status === 'ERROR' ? 'Inventory sync FAILED — portal is showing stale data' : 'Inventory sync needs attention',
                run.errors, run.warnings);
  props.setProperty('INV_ALERT_SIG', sig);
  props.setProperty('INV_ALERT_AT', String(Date.now()));
  props.setProperty('INV_ALERTING', '1');
}

function _invSendAlert(level, subject, errors, warnings) {
  var url = 'https://docs.google.com/spreadsheets/d/' + INV_PIPELINE.CRM_SHEET_ID;
  var body = '<div style="font-family:Arial,sans-serif;font-size:14px">';
  body += '<p><b>' + subject + '</b></p>';
  if (errors.length) body += '<p><b style="color:#C8102E">Errors</b></p><ul>' + errors.map(function (e) { return '<li>' + _invEsc(e) + '</li>'; }).join('') + '</ul>';
  if (warnings.length) body += '<p><b style="color:#B45309">Warnings</b></p><ul>' + warnings.map(function (e) { return '<li>' + _invEsc(e) + '</li>'; }).join('') + '</ul>';
  body += '<p>Status tab: <a href="' + url + '">CRM Master → Pipeline Status</a></p>';
  body += '<p style="color:#888;font-size:12px">Sent by the Redbird inventory pipeline (Apps Script). Repeats of the same issue are limited to once every ' +
          INV_PIPELINE.ALERT_REPEAT_HOURS + ' hours.</p></div>';
  MailApp.sendEmail({ to: INV_PIPELINE.ALERT_EMAILS.join(','), subject: '[Redbird Inventory ' + level + '] ' + subject, htmlBody: body });
}

// ============================================================================
//  SOURCE READING (by header name)
// ============================================================================

function _invReadTable(sheetId, tabName, fields, label) {
  var out = { label: label, rows: [], missingOptional: [], error: '' };
  var sh;
  try {
    sh = SpreadsheetApp.openById(sheetId).getSheetByName(tabName);
  } catch (e) {
    out.error = label + ': cannot open the spreadsheet (' + e.message + '). Check that the script owner still has access.';
    return out;
  }
  if (!sh) { out.error = label + ': tab "' + tabName + '" not found. Was it renamed?'; return out; }
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) { out.error = label + ': tab is empty.'; return out; }
  var data = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // Find the header row: the first of the top 10 rows that matches the most field names
  var best = -1, bestHits = 0, bestMap = null;
  for (var i = 0; i < Math.min(10, data.length); i++) {
    var map = _invMapHeaders(data[i], fields);
    var hits = Object.keys(map).length;
    if (hits > bestHits) { best = i; bestHits = hits; bestMap = map; }
  }
  if (best < 0) { out.error = label + ': could not find a header row.'; return out; }

  var missingReq = [];
  Object.keys(fields).forEach(function (k) {
    if (bestMap[k] === undefined) {
      if (fields[k].required) missingReq.push(fields[k].names[0]);
      else out.missingOptional.push(fields[k].names[0]);
    }
  });
  if (missingReq.length) {
    out.error = label + ': required column(s) not found: ' + missingReq.join(', ') +
                '. A column was likely renamed or deleted. Sync stopped; last good data is still live.';
    return out;
  }
  out.rows = data.slice(best + 1);
  out.get = function (row, key) { var c = bestMap[key]; return c === undefined ? '' : row[c]; };
  return out;
}

function _invMapHeaders(headerRow, fields) {
  var norm = headerRow.map(_invNorm);
  var map = {};
  Object.keys(fields).forEach(function (k) {
    var names = fields[k].names.map(_invNorm);
    for (var n = 0; n < names.length; n++) {
      var idx = norm.indexOf(names[n]);       // first match wins (e.g. duplicate "Per Tote")
      if (idx >= 0) { map[k] = idx; break; }
    }
  });
  return map;
}

// ============================================================================
//  SMALL HELPERS
// ============================================================================

function _invNorm(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9%]/g, ''); }
function _invDay(d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function _invAdd(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); x.setHours(0, 0, 0, 0); return x; }
function _invDate(v) {
  if (v instanceof Date && !isNaN(v)) return _invDay(v);
  if (typeof v === 'string' && /\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4}/.test(v)) { var p = new Date(v); if (!isNaN(p)) return _invDay(p); }
  return null;
}
function _invKey(d, tz) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
function _invFmt(d, tz) { return Utilities.formatDate(d, tz, 'M/d/yyyy'); }
function _invNum(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = parseFloat(String(v || '').replace(/[,$\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
function _invPct(v) {
  var n = _invNum(String(v).replace('%', ''));
  if (typeof v === 'string' && v.indexOf('%') >= 0) n = n / 100;
  else if (n > 1) n = n / 100;
  return n;
}
function _invBool(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function _invMedian(a) {
  if (!a.length) return 0;
  var s = a.slice().sort(function (x, y) { return x - y; }), m = Math.floor(s.length / 2);
  return Math.round((s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2) * 100) / 100;
}
function _invEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

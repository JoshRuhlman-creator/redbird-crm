// =============================================================================
// crossreference.js — merge the latest Cultivera order export into OrdersLog
// (CRM Master sheet). Runs from sync.sh after download.js.
//
// Merge rule (unchanged): every order in the new export replaces its rows in
// OrdersLog; every other order already in OrdersLog is kept. So OrdersLog is
// a growing, de-duplicated history of every order ever exported.
//
// Safety (new, 2026-11-09 — after OrdersLog lost Mar–Jul 2026 history):
//   1. Empty-export guard   — an export with 0 rows is treated as a failed
//                             download; nothing is written.
//   2. No-shrink guard      — if the merged result has fewer rows than
//                             OrdersLog already has (beyond a small tolerance
//                             for edited orders), nothing is written.
//                             Override for a deliberate cleanup only:
//                             ALLOW_SHRINK=1 node crossreference.js
//   3. Local backup         — the current OrdersLog is saved to
//                             backups/orderslog_<timestamp>.json before every
//                             write (last 60 kept) so any bad write can be
//                             restored with restore_orderslog.js.
//   4. Safe write order     — new data is written over the top first; only
//                             leftover rows below it are cleared afterwards.
//                             The tab is never empty mid-write.
//   5. Single runner        — a lock file stops a manual backfill and the
//                             cron run from writing at the same time.
//   6. Status tab           — every run (success or blocked) is recorded on
//                             the "Cultivera Sync Status" tab. The CRM Apps
//                             Script watchdog reads it and emails if the sync
//                             is blocked or hasn't succeeded recently.
// =============================================================================
const ExcelJS = require('exceljs');
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

const SAVE_DIR       = process.env.SAVE_DIR || (fs.existsSync('/opt/redbird-sync/exports') ? '/opt/redbird-sync/exports' : path.join(process.env.HOME, 'Downloads', 'Redbird CRM Exports'));
const SPREADSHEET_ID = '1Nrz05Rr8B1-5-2HXIXByy_pL6TH9U6wz-HUgu9LSkjw';
const SHEET_TAB      = 'OrdersLog';
const STATUS_TAB     = 'Cultivera Sync Status';
const CREDENTIALS    = path.join(__dirname, 'credentials.json');
const BACKUP_DIR     = path.join(__dirname, 'backups');
const LOCK_FILE      = path.join(__dirname, '.crossreference.lock');
const KEEP_BACKUPS   = 60;
const SHRINK_TOLERANCE = (existing) => Math.max(50, Math.round(existing * 0.02)); // edited orders can lose a few lines

const KEEP_HEADERS = [
  'Order #', 'Product', 'Product Line', 'Strain', 'Units', 'Line Total',
  'Sample', 'Submitted Date', 'Client', 'License #', 'Status', 'Estimated delivery date'
];

const ts = () => new Date().toISOString();
const sleep = ms => new Promise(r => setTimeout(r, ms));

function getLatestFile() {
  const files = fs.readdirSync(SAVE_DIR)
    .filter(f => f.startsWith('order_summary_') && f.endsWith('.xlsx'))
    .map(f => ({ name: f, time: fs.statSync(path.join(SAVE_DIR, f)).mtime }))
    .sort((a, b) => b.time - a.time);
  if (!files.length) throw new Error('No order_summary_*.xlsx files found in ' + SAVE_DIR);
  return path.join(SAVE_DIR, files[0].name);
}

async function readExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[0];
  const rows = [];
  let headers = [];
  sheet.eachRow((row, rowNum) => {
    const values = row.values.slice(1);
    if (rowNum === 1) {
      headers = values.map(v => (v || '').toString().trim());
    } else {
      const obj = {};
      headers.forEach((h, i) => obj[h] = values[i] !== undefined ? values[i] : '');
      rows.push(obj);
    }
  });
  console.log('Excel rows read:', rows.length);
  return { rows, headers };
}

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: 'v4', auth: authClient });
}

function rowToValues(row) {
  return KEEP_HEADERS.map(h => {
    const val = row[h];
    if (val instanceof Date) return val.toLocaleDateString('en-US');
    return val !== undefined && val !== null ? val.toString() : '';
  });
}

// ---- lock: one writer at a time (cron vs. manual backfill) ----
async function acquireLock() {
  for (let waited = 0; waited < 15 * 60; waited += 10) {
    try {
      fs.writeFileSync(LOCK_FILE, String(process.pid), { flag: 'wx' });
      return;
    } catch (e) {
      const pid = Number(fs.readFileSync(LOCK_FILE, 'utf8')) || 0;
      let alive = false;
      try { if (pid) { process.kill(pid, 0); alive = true; } } catch (_) {}
      if (!alive) { fs.unlinkSync(LOCK_FILE); continue; }        // stale lock from a crashed run
      if (waited === 0) console.log('Another crossreference run (pid ' + pid + ') is writing. Waiting…');
      await sleep(10000);
    }
  }
  throw new Error('Timed out after 15 min waiting for another crossreference run to finish');
}
function releaseLock() { try { fs.unlinkSync(LOCK_FILE); } catch (_) {} }

// ---- status tab (read by the CRM Apps Script watchdog) ----
async function ensureTab(sheets, title) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields: 'sheets.properties.title' });
  if (!meta.data.sheets.find(s => s.properties.title === title)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title } } }] },
    });
    console.log('Created sheet:', title);
  }
}

async function writeStatus(sheets, s) {
  try {
    await ensureTab(sheets, STATUS_TAB);
    let lastSuccess = s.ok ? s.at : '';
    if (!s.ok) {
      const prev = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: STATUS_TAB + '!A1:B10' }).catch(() => null);
      const row = ((prev && prev.data.values) || []).find(r => r[0] === 'Last Success');
      lastSuccess = row ? row[1] : '';
    }
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: STATUS_TAB + '!A1:B10',
      valueInputOption: 'RAW',
      requestBody: { values: [
        ['Field', 'Value'],
        ['Status', s.state || (s.ok ? 'OK' : 'BLOCKED')],
        ['Last Run', s.at],
        ['Last Success', lastSuccess],
        ['Export File', s.file || ''],
        ['Rows Before', String(s.before ?? '')],
        ['Rows After', String(s.after ?? '')],
        ['Earliest Order Date', s.earliest || ''],
        ['Message', s.message || ''],
        ['Backup', s.backup || ''],
      ] },
    });
  } catch (e) {
    console.error('ERROR: could not write status tab:', e.message);
  }
}

function earliestSubmitted(rows) {
  let min = '';
  rows.forEach(r => {
    const d = new Date(r[7]);
    if (!isNaN(d) && d.getFullYear() > 2000) {
      const iso = d.toISOString().slice(0, 10);
      if (!min || iso < min) min = iso;
    }
  });
  return min;
}

function backupExisting(values) {
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
  const file = path.join(BACKUP_DIR, 'orderslog_' + ts().replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(file, JSON.stringify(values));
  const old = fs.readdirSync(BACKUP_DIR).filter(f => f.startsWith('orderslog_')).sort();
  old.slice(0, Math.max(0, old.length - KEEP_BACKUPS)).forEach(f => fs.unlinkSync(path.join(BACKUP_DIR, f)));
  console.log('Backup saved:', file, '(' + (values.length - 1) + ' rows)');
  return file;
}


// Write a full table to OrdersLog A:L in 1,000-row batches (each retried 3x),
// then clear any leftover rows below it.
async function writeTable(sheets, values, previousLength) {
  const BATCH = 1000;
  for (let i = 0; i < values.length; i += BATCH) {
    const batch = values.slice(i, i + BATCH);
    const startRow = i + 1;
    for (let attempt = 1; ; attempt++) {
      try {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: SHEET_TAB + '!A' + startRow,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: batch },
        });
        break;
      } catch (e) {
        if (attempt >= 3) throw e;
        console.log('Write of rows ' + startRow + '+ failed (' + e.message + '), retry ' + attempt + '/2…');
        await sleep(5000 * attempt);
      }
    }
    console.log('Written rows', startRow, 'to', startRow + batch.length - 1);
    if (i + BATCH < values.length) await sleep(1000);
  }
  if (previousLength > values.length) {
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: SHEET_TAB + '!A' + (values.length + 1) + ':L' + previousLength,
    });
    console.log('Cleared leftover rows', values.length + 1, 'to', previousLength);
  }
}

async function readStatus(sheets) {
  try {
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: STATUS_TAB + '!A1:B12' });
    const o = {}; (r.data.values || []).forEach(x => { if (x[0]) o[x[0]] = x[1] || ''; });
    return o;
  } catch (_) { return {}; }
}

(async () => {
  let sheets;
  const status = { ok: false, at: ts() };
  try {
    await acquireLock();
    const filePath = getLatestFile();
    status.file = path.basename(filePath);
    console.log('Reading file:', filePath);
    const { rows: excelRows, headers } = await readExcel(filePath);
    sheets = await getSheetsClient();

    // Recovery: a previous run was killed mid-write → put its backup back first
    const prev = await readStatus(sheets);
    if (prev['Status'] === 'WRITING' && prev['Backup'] && fs.existsSync(prev['Backup'])) {
      console.log('Previous run stopped mid-write. Restoring', prev['Backup'], 'before merging…');
      const restore = JSON.parse(fs.readFileSync(prev['Backup'], 'utf8'));
      const cur = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: SHEET_TAB + '!A:L' });
      await writeTable(sheets, restore, (cur.data.values || []).length);
      console.log('Restored', restore.length - 1, 'rows.');
    }

    // Guard 1: empty / malformed export = failed download
    const missing = ['Order #', 'Status', 'Line Total'].filter(h => headers.indexOf(h) < 0);
    if (missing.length) throw Object.assign(new Error('Export is missing columns: ' + missing.join(', ') + ' — Cultivera export format changed or download failed'), { blocked: true });
    if (!excelRows.length) throw Object.assign(new Error('Export has 0 rows — treating as a failed download'), { blocked: true });

    await ensureTab(sheets, SHEET_TAB);
    const existingRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: SHEET_TAB + '!A:L',
    });
    const existingAll = existingRes.data.values || [];
    const existingRows = existingAll.slice(1).filter(r => r.some(v => String(v || '').trim()));

    const newOrderNums = new Set(excelRows.map(r => (r['Order #'] || '').toString().trim()).filter(Boolean));
    const keepRows = existingRows.filter(row => {
      const orderNum = (row[0] || '').toString().trim();
      return orderNum && !newOrderNums.has(orderNum);
    });

    const allValues = [KEEP_HEADERS, ...keepRows, ...excelRows.map(rowToValues)];
    const before = existingRows.length, after = allValues.length - 1;
    status.before = before; status.after = after;
    status.earliest = earliestSubmitted(allValues.slice(1));

    console.log('Existing line items:', before);
    console.log('New export line items:', excelRows.length);
    console.log('Orders in new export:', newOrderNums.size);
    console.log('Historical rows to keep:', keepRows.length);
    console.log('Rows after merge:', after);

    // Guard 2: never shrink OrdersLog (beyond edits to orders in this export)
    if (after < before - SHRINK_TOLERANCE(before) && process.env.ALLOW_SHRINK !== '1') {
      throw Object.assign(new Error('Merged result has ' + after + ' rows but OrdersLog has ' + before +
        '. Refusing to shrink history. If this is intentional, re-run with ALLOW_SHRINK=1.'), { blocked: true });
    }

    // Backup → mark WRITING → write over the top → clear leftover tail.
    // On failure, put the backup straight back so the tab is never left half-written.
    status.backup = backupExisting(existingAll);
    await writeStatus(sheets, Object.assign({}, status, { state: 'WRITING', message: 'Write in progress' }));
    try {
      await writeTable(sheets, allValues, existingAll.length);
    } catch (writeErr) {
      console.error('Write failed (' + writeErr.message + '). Restoring backup…');
      try {
        const cur = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: SHEET_TAB + '!A:L' });
        await writeTable(sheets, existingAll, (cur.data.values || []).length);
        throw Object.assign(new Error('Write failed (' + writeErr.message + '); OrdersLog restored from backup, no data lost'), { state: 'FAILED' });
      } catch (restoreErr) {
        if (restoreErr.state) throw restoreErr;
        throw Object.assign(new Error('Write failed AND restore failed (' + restoreErr.message + '). Next run restores ' + status.backup + ' automatically, or run: node restore_orderslog.js ' + path.basename(status.backup)), { state: 'WRITING' });
      }
    }

    const syncTime = ts();
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: SHEET_TAB + '!N1',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['Last Sync', syncTime]] },
    });

    status.ok = true; status.at = syncTime;
    status.message = 'Merged ' + newOrderNums.size + ' orders from export; history ' + before + ' → ' + after + ' rows.';
    console.log('Cross-reference complete:', after, 'total rows.');
    console.log('Sync timestamp:', syncTime);
  } catch (err) {
    status.ok = false;
    status.state = err.blocked ? 'BLOCKED' : (err.state || 'FAILED');
    status.message = (status.state === 'WRITING' ? 'FAILED: ' : status.state + ': ') + err.message;
    console.error('ERROR: ' + status.message);
    process.exitCode = 1;
  } finally {
    if (sheets) await writeStatus(sheets, status);
    else {
      try { await writeStatus(await getSheetsClient(), status); } catch (_) {}
    }
    releaseLock();
  }
})();

// One-time installer for the hardened order sync. Run on the VM:
//   cd /opt/redbird-sync && node /tmp/redbird-vm/patch-sync.js
// Idempotent: safe to run more than once. Backs up every file it changes.
//
//  1. crossreference.js  → replaced with the guarded version (backup kept)
//  2. download.js        → date range can be overridden for backfills:
//                            ORDERS_FROM=03/01/2026 ORDERS_TO=07/25/2026 node download.js
//                          (normal cron runs keep the 60-day window);
//                          backfill exports get their own file name
//  3. download.js + inventory_download.js → Cultivera login read from .env
//                          (moved out of the code; .env is chmod 600)
//  4. restore_orderslog.js + load-env.js installed
'use strict';
const fs = require('fs');
const path = require('path');

const DIR = '/opt/redbird-sync';
const SRC = __dirname;
const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
const backup = f => { const b = f + '.bak-' + stamp; if (!fs.existsSync(b)) fs.copyFileSync(f, b); return b; };
const log = m => console.log('  ' + m);

// ---- 1. support files ----
['load-env.js', 'restore_orderslog.js'].forEach(f => { fs.copyFileSync(path.join(SRC, f), path.join(DIR, f)); log('installed ' + f); });

// ---- 2. crossreference.js ----
const xref = path.join(DIR, 'crossreference.js');
const newXref = fs.readFileSync(path.join(SRC, 'crossreference.js'), 'utf8');
if (fs.readFileSync(xref, 'utf8') !== newXref) { log('backup ' + backup(xref)); fs.writeFileSync(xref, newXref); log('crossreference.js replaced (guards + backups + status tab)'); }
else log('crossreference.js already current');

// ---- 3. credentials → .env ----
const envFile = path.join(DIR, '.env');
let env = fs.existsSync(envFile) ? fs.readFileSync(envFile, 'utf8') : '';
function moveCreds(file) {
  const p = path.join(DIR, file);
  if (!fs.existsSync(p)) { log(file + ' not found, skipped'); return; }
  let src = fs.readFileSync(p, 'utf8');
  const u = src.match(/const\s+USERNAME\s*=\s*(['"])(.*?)\1\s*;/);
  const pw = src.match(/const\s+PASSWORD\s*=\s*(['"])(.*?)\1\s*;/);
  if (!u && !pw) { log(file + ': login already read from .env'); return; }
  backup(p);
  if (u && !/^CULTIVERA_USER=/m.test(env)) env += 'CULTIVERA_USER=' + u[2] + '\n';
  if (pw && !/^CULTIVERA_PASSWORD=/m.test(env)) env += 'CULTIVERA_PASSWORD=' + pw[2] + '\n';
  if (u) src = src.replace(u[0], "const USERNAME  = process.env.CULTIVERA_USER;");
  if (pw) src = src.replace(pw[0], "const PASSWORD  = process.env.CULTIVERA_PASSWORD;");
  if (!/require\(['"]\.\/load-env['"]\)/.test(src)) src = "require('./load-env');\n" + src;
  fs.writeFileSync(p, src);
  log(file + ': login moved to .env');
}
moveCreds('download.js');
moveCreds('inventory_download.js');
fs.writeFileSync(envFile, env, { mode: 0o600 });
fs.chmodSync(envFile, 0o600);
log('.env written (chmod 600)');

// ---- 4. download.js date override ----
const dl = path.join(DIR, 'download.js');
let d = fs.readFileSync(dl, 'utf8');
if (!/ORDERS_FROM/.test(d)) {
  backup(dl);
  const fromRe = /const\s+FROM_DATE\s*=\s*formatDate\((\w+)\);/;
  const toRe = /const\s+TO_DATE\s*=\s*formatDate\((\w+)\);/;
  if (!fromRe.test(d) || !toRe.test(d)) throw new Error('download.js: FROM_DATE/TO_DATE lines not found — not patched');
  d = d.replace(fromRe, "const FROM_DATE = process.env.ORDERS_FROM || formatDate($1);   // MM/DD/YYYY override for backfills");
  d = d.replace(toRe, "const TO_DATE   = process.env.ORDERS_TO   || formatDate($1);");
  d = d.replace(/const\s+SAVE_FILE\s*=\s*path\.join\(SAVE_DIR,\s*'order_summary_'\s*\+\s*new Date\(\)\.toISOString\(\)\.slice\(0,10\)\s*\+\s*'\.xlsx'\);/,
    "const SAVE_FILE = path.join(SAVE_DIR, 'order_summary_' + new Date().toISOString().slice(0,10) + (process.env.ORDERS_FROM ? '_backfill' : '') + '.xlsx');");
  fs.writeFileSync(dl, d);
  log('download.js: ORDERS_FROM / ORDERS_TO override added');
} else log('download.js: date override already present');

console.log('\nPatch complete. Run:  node -c download.js && node -c crossreference.js  to verify syntax.');

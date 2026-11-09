// Restore OrdersLog from a backup written by crossreference.js.
//   node restore_orderslog.js                 → lists available backups
//   node restore_orderslog.js <backup file>   → restores that backup (A:L)
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1Nrz05Rr8B1-5-2HXIXByy_pL6TH9U6wz-HUgu9LSkjw';
const SHEET_TAB = 'OrdersLog';
const BACKUP_DIR = path.join(__dirname, 'backups');

(async () => {
  const arg = process.argv[2];
  if (!arg) {
    const files = fs.existsSync(BACKUP_DIR) ? fs.readdirSync(BACKUP_DIR).filter(f => f.startsWith('orderslog_')).sort() : [];
    files.slice(-15).forEach(f => {
      const rows = JSON.parse(fs.readFileSync(path.join(BACKUP_DIR, f), 'utf8')).length - 1;
      console.log(f, '—', rows, 'rows');
    });
    if (!files.length) console.log('No backups yet.');
    return;
  }
  const file = fs.existsSync(arg) ? arg : path.join(BACKUP_DIR, arg);
  const values = JSON.parse(fs.readFileSync(file, 'utf8'));
  const auth = new google.auth.GoogleAuth({ keyFile: path.join(__dirname, 'credentials.json'), scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  const sheets = google.sheets({ version: 'v4', auth: await auth.getClient() });
  const cur = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: SHEET_TAB + '!A:L' });
  const curLen = (cur.data.values || []).length;
  for (let i = 0; i < values.length; i += 1000) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID, range: SHEET_TAB + '!A' + (i + 1),
      valueInputOption: 'USER_ENTERED', requestBody: { values: values.slice(i, i + 1000) },
    });
    console.log('Restored rows', i + 1, 'to', Math.min(i + 1000, values.length));
  }
  if (curLen > values.length) {
    await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: SHEET_TAB + '!A' + (values.length + 1) + ':L' + curLen });
  }
  console.log('Restore complete:', values.length - 1, 'rows from', path.basename(file));
})().catch(e => { console.error('Restore failed:', e.message); process.exit(1); });

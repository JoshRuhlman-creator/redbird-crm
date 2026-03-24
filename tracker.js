const http = require('http');
const https = require('https');
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
const url = require('url');

// ── Config ──
const PORT = 3100;
const SPREADSHEET_ID = '1Nrz05Rr8B1-5-2HXIXByy_pL6TH9U6wz-HUgu9LSkjw';
const SHEET_TAB = 'EmailTracking';
const CREDENTIALS = path.join(__dirname, 'credentials.json');

// 1x1 transparent PNG pixel
const PIXEL = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==',
  'base64'
);

let sheetsClient = null;

async function getSheetsClient() {
  if (sheetsClient) return sheetsClient;
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const authClient = await auth.getClient();
  sheetsClient = google.sheets({ version: 'v4', auth: authClient });
  return sheetsClient;
}

async function ensureSheet() {
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets.some(s => s.properties.title === SHEET_TAB);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_TAB } } }] }
    });
    // Write headers
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: SHEET_TAB + '!A1',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['Timestamp', 'Campaign ID', 'Event', 'Email', 'Account', 'Link URL', 'User Agent', 'IP']] },
    });
    console.log('[Tracker] Created EmailTracking sheet with headers');
  }
}

async function logEvent(campaignId, event, email, account, linkUrl, userAgent, ip) {
  try {
    const sheets = await getSheetsClient();
    const timestamp = new Date().toISOString();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: SHEET_TAB + '!A:H',
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[timestamp, campaignId, event, email, account, linkUrl || '', userAgent || '', ip || '']]
      },
    });
    console.log(`[Tracker] ${event}: ${email} (campaign: ${campaignId})`);
  } catch (err) {
    console.error('[Tracker] Failed to log event:', err.message);
  }
}

// ── HTTP Server ──
const server = http.createServer((req, res) => {
  const parsed = url.parse(req.url, true);
  const q = parsed.query;
  const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
  const ua = req.headers['user-agent'] || '';

  // CORS headers for CRM dashboard queries
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  // ── /pixel — tracking pixel (email open) ──
  if (parsed.pathname === '/pixel' || parsed.pathname === '/pixel.png') {
    const cid = q.cid || '';   // campaign ID
    const e = q.e || '';       // email (base64)
    const a = q.a || '';       // account (base64)

    const email = e ? Buffer.from(e, 'base64').toString() : 'unknown';
    const account = a ? Buffer.from(a, 'base64').toString() : 'unknown';

    // Log open event (fire-and-forget)
    logEvent(cid, 'open', email, account, '', ua, ip);

    // Return the transparent pixel
    res.writeHead(200, {
      'Content-Type': 'image/png',
      'Content-Length': PIXEL.length,
      'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0',
      'Pragma': 'no-cache',
    });
    res.end(PIXEL);
    return;
  }

  // ── /click — click redirect ──
  if (parsed.pathname === '/click') {
    const cid = q.cid || '';
    const e = q.e || '';
    const a = q.a || '';
    const u = q.u || '';       // destination URL (base64)

    const email = e ? Buffer.from(e, 'base64').toString() : 'unknown';
    const account = a ? Buffer.from(a, 'base64').toString() : 'unknown';
    const destUrl = u ? Buffer.from(u, 'base64').toString() : '/';

    // Log click event
    logEvent(cid, 'click', email, account, destUrl, ua, ip);

    // Redirect to destination
    res.writeHead(302, { 'Location': destUrl });
    res.end();
    return;
  }

  // ── /stats — campaign stats API for CRM dashboard ──
  if (parsed.pathname === '/stats') {
    const cid = q.cid || '';
    const callback = q.callback || '';

    getSheetsClient().then(sheets => {
      return sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: SHEET_TAB + '!A:H',
      });
    }).then(result => {
      const rows = (result.data.values || []).slice(1); // skip header
      let events = rows;
      if (cid) {
        events = rows.filter(r => (r[1] || '') === cid);
      }

      // Aggregate stats
      const opens = {};
      const clicks = {};
      const bounces = {};
      events.forEach(r => {
        const event = r[2] || '';
        const email = r[3] || '';
        const campaignId = r[1] || '';
        if (event === 'open') opens[email] = (opens[email] || 0) + 1;
        if (event === 'click') clicks[email] = (clicks[email] || 0) + 1;
        if (event === 'bounce') bounces[email] = true;
      });

      // Group by campaign
      const byCampaign = {};
      events.forEach(r => {
        const cmpId = r[1] || 'unknown';
        if (!byCampaign[cmpId]) byCampaign[cmpId] = { opens: {}, clicks: {}, bounces: {}, events: [] };
        const bc = byCampaign[cmpId];
        const event = r[2] || '';
        const email = r[3] || '';
        if (event === 'open') bc.opens[email] = (bc.opens[email] || 0) + 1;
        if (event === 'click') bc.clicks[email] = (bc.clicks[email] || 0) + 1;
        if (event === 'bounce') bc.bounces[email] = true;
        bc.events.push({ time: r[0], event: event, email: email, account: r[4] || '', link: r[5] || '' });
      });

      const stats = {};
      Object.keys(byCampaign).forEach(id => {
        const bc = byCampaign[id];
        stats[id] = {
          uniqueOpens: Object.keys(bc.opens).length,
          totalOpens: Object.values(bc.opens).reduce((s, v) => s + v, 0),
          uniqueClicks: Object.keys(bc.clicks).length,
          totalClicks: Object.values(bc.clicks).reduce((s, v) => s + v, 0),
          bounces: Object.keys(bc.bounces).length,
          openEmails: Object.keys(bc.opens),
          clickEmails: Object.keys(bc.clicks),
          bounceEmails: Object.keys(bc.bounces),
          recentEvents: bc.events.slice(-20),
        };
      });

      const json = JSON.stringify({ ok: true, stats: stats });
      if (callback) {
        res.writeHead(200, { 'Content-Type': 'application/javascript' });
        res.end(callback + '(' + json + ')');
      } else {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(json);
      }
    }).catch(err => {
      const errJson = JSON.stringify({ ok: false, error: err.message });
      if (callback) {
        res.writeHead(200, { 'Content-Type': 'application/javascript' });
        res.end(callback + '(' + errJson + ')');
      } else {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(errJson);
      }
    });
    return;
  }

  // ── /health ──
  if (parsed.pathname === '/health') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ status: 'ok', uptime: process.uptime() }));
    return;
  }

  // 404
  res.writeHead(404, { 'Content-Type': 'text/plain' });
  res.end('Not found');
});

// ── Start ──
ensureSheet().then(() => {
  server.listen(PORT, '0.0.0.0', () => {
    console.log(`[Tracker] Email tracking server running on port ${PORT}`);
  });
}).catch(err => {
  console.error('[Tracker] Failed to start:', err.message);
  process.exit(1);
});

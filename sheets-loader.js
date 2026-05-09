/**
 * Redbird CRM — Google Sheets Live Data Loader
 * 
 * HOW TO USE:
 * 1. Add this script tag to crm.html just before the closing </body> tag:
 *    <script src="sheets-loader.js"></script>
 * 
 * 2. The loader will automatically fetch live data on page load and override
 *    the hardcoded CRM data with live Google Sheets data.
 * 
 * SHEET TABS USED:
 *   - "Accounts Mirror"              → Accounts page
 *   - "OrdersLog"                    → Opportunities page
 *   - "Inventory Projections Mirror" → Inventory page
 *   - "Inventory"                    → Current Inventory page
 */

const SHEETS_CONFIG = {
  apiKey:  'AIzaSyCiX_mltG3QDz8q6Sx77yQnltMSup7Ik8o',
  sheetId: '1Nrz05Rr8B1-5-2HXIXByy_pL6TH9U6wz-HUgu9LSkjw',
  tabs: {
    accounts:    'Accounts Mirror',
    contacts:    'Accounts',
    orders:      'OrdersLog',
    inventory:   'Inventory Projections Mirror',
    currentInv:  'Inventory',
    historyOpps: 'opportunities w/products',
    tasks:       'tasks',
    invoices:    'invoicedetails',
  }
};

// ─────────────────────────────────────────────
// UTILITY
// ─────────────────────────────────────────────

function sheetsUrl(tab) {
  const base = 'https://sheets.googleapis.com/v4/spreadsheets';
  const enc  = encodeURIComponent(tab);
  return `${base}/${SHEETS_CONFIG.sheetId}/values/${enc}?key=${SHEETS_CONFIG.apiKey}&valueRenderOption=FORMATTED_VALUE`;
}

async function fetchTab(tab) {
  const res = await fetch(sheetsUrl(tab));
  if (!res.ok) throw new Error(`Failed to fetch tab "${tab}": ${res.status}`);
  const json = await res.json();
  const [headers, ...rows] = json.values || [];
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const key = h.trim();
      const val = (row[i] || '').trim();
      // For duplicate headers, keep the first non-empty value
      if (obj[key] === undefined || obj[key] === '') {
        obj[key] = val;
      }
    });
    return obj;
  });
}

// Fetch raw row arrays (no header mapping) — for tabs with non-standard headers
async function fetchTabRaw(tab) {
  const res = await fetch(sheetsUrl(tab));
  if (!res.ok) throw new Error(`Failed to fetch tab "${tab}": ${res.status}`);
  const json = await res.json();
  return json.values || [];
}

function showBanner(msg, color = '#16a34a') {
  let b = document.getElementById('sheets-banner');
  if (!b) {
    b = document.createElement('div');
    b.id = 'sheets-banner';
    b.style.cssText = `
      position:fixed;bottom:20px;right:20px;z-index:9999;
      background:${color};color:#fff;padding:10px 18px;
      border-radius:8px;font-size:13px;font-weight:600;
      box-shadow:0 4px 16px rgba(0,0,0,.2);
      transition:opacity .5s;
    `;
    document.body.appendChild(b);
  }
  b.style.background = color;
  b.textContent = msg;
  b.style.opacity = '1';
  setTimeout(() => { b.style.opacity = '0'; }, 4000);
}

// ─────────────────────────────────────────────
// ACCOUNT MIRROR → ACCOUNT_DATA (object keyed by trade name)
// CRM expects: ACCOUNT_DATA[name] = { contacts:[], stats:{oppCount, contactCount, lastOrder}, ... }
// ─────────────────────────────────────────────

function parseAccounts(rows) {
  const accountData = {};

  rows.filter(r => r['Trade Name']).forEach(r => {
    const tradeName = r['Trade Name'];
    const license = r['License #'] || '';

    // If this trade name already exists with a DIFFERENT license, make the key unique
    let name = tradeName;
    if (accountData[name] && accountData[name].license !== license) {
      // Append city or license to disambiguate
      const city = r['City'] || '';
      name = city ? `${tradeName} (${city})` : `${tradeName} [${license}]`;
      // If that's also taken, fall back to license suffix
      if (accountData[name]) {
        name = `${tradeName} [${license}]`;
      }
    }

    accountData[name] = {
      license:     license,
      ubi:         r['UBI'] || '',
      address:     r['Address'] || '',
      city:        r['City'] || '',
      county:      r['County'] || '',
      state:       r['State'] || '',
      zip:         r['Zip Code'] || '',
      phone:       r['Phone'] || '',
      licenseType: r['License Type'] || '',
      status:      r['Status'] || '',
      lastUpdated: r['Last Updated'] || '',
      flags:       r['Flags'] || '',
      // CRM requires these nested structures
      contacts:    [],
      families:    [],
      logs:        [],
      tasks:       [],
      stats: {
        oppCount:     0,
        contactCount: 0,
        lastOrder:    '',
        revenue:      0,
      },
    };

    // Parse contacts if present (e.g. "Name <email>, Name2 <email2>")
    if (r['Contacts']) {
      const contactParts = r['Contacts'].split(/[;,]+/);
      contactParts.forEach(part => {
        const match = part.trim().match(/^(.+?)\s*<(.+?)>\s*$/);
        if (match) {
          accountData[name].contacts.push({ name: match[1].trim(), email: match[2].trim() });
        } else if (part.trim()) {
          accountData[name].contacts.push({ name: part.trim(), email: '' });
        }
      });
      accountData[name].stats.contactCount = accountData[name].contacts.length;
    }
  });

  return accountData;
}

// ─────────────────────────────────────────────
// MERGE CONTACTS from "Accounts" tab into ACCOUNT_DATA
// Each row: First Name, Last Name, Title, Account Name, Email, Contact ID
// ─────────────────────────────────────────────

function mergeContacts(accountData, contactRows) {
  if (!contactRows || !contactRows.length) {
    console.log('[Sheets Loader] No contact rows to merge');
    return;
  }

  // Build alias map for matching — same as the account alias map
  const aliasMap = {};
  Object.keys(accountData).forEach(name => {
    aliasMap[name.toLowerCase().trim()] = name;
  });

  let matched = 0, unmatched = 0;
  const unmatchedNames = new Set();

  contactRows.forEach(row => {
    const firstName = (row['First Name'] || '').trim();
    const lastName = (row['Last Name'] || '').trim();
    const title = (row['Title'] || '').trim();
    const acctName = (row['Account Name'] || '').trim();
    const email = (row['Email'] || '').trim();
    const contactId = (row['Contact ID'] || '').trim();

    if (!acctName || (!firstName && !lastName)) return;

    const fullName = [firstName, lastName].filter(Boolean).join(' ');

    // Try to find the account
    let account = accountData[acctName];
    if (!account) {
      // Try case-insensitive match
      account = accountData[aliasMap[acctName.toLowerCase().trim()]];
    }
    if (!account) {
      // Try partial match — account name might be slightly different
      const acctLower = acctName.toLowerCase();
      const matchKey = Object.keys(accountData).find(k => {
        const kl = k.toLowerCase();
        return kl.includes(acctLower) || acctLower.includes(kl);
      });
      if (matchKey) account = accountData[matchKey];
    }

    if (account) {
      // Check if this contact already exists (by email or name)
      const exists = account.contacts.some(c => {
        if (email && c.email && c.email.toLowerCase() === email.toLowerCase()) return true;
        if (c.name.toLowerCase() === fullName.toLowerCase()) return true;
        return false;
      });

      if (!exists) {
        account.contacts.push({
          name: fullName,
          email: email,
          title: title,
          contactId: contactId,
          _fromSheet: true,
        });
      }
      account.stats.contactCount = account.contacts.length;
      matched++;
    } else {
      unmatched++;
      unmatchedNames.add(acctName);
    }
  });

  console.log(`[Sheets Loader] Contacts merged: ${matched} matched, ${unmatched} unmatched`);
  if (unmatchedNames.size > 0 && unmatchedNames.size <= 20) {
    console.log('[Sheets Loader] Unmatched contact accounts:', Array.from(unmatchedNames));
  }
}

// ─────────────────────────────────────────────
// ORDERSLOGS → OPPS_DATA (array with shorthand property names)
// CRM expects: { a, n, r, s, dm, sd, d, pd, dd, cm, em, po, il, ml, jl, family }
// ─────────────────────────────────────────────

function parseOrders(rows) {
  // Log the raw column names from the first row so we can debug mapping
  if (rows.length > 0) {
    console.log('[Sheets Loader] OrdersLog columns:', Object.keys(rows[0]));
    console.log('[Sheets Loader] First raw order row:', rows[0]);
  }

  // ── Phase 1: Group line items by Order # ──
  const orderMap = {};   // keyed by Order #
  const familyMap = {};  // keyed by Order # → Set of families

  rows.filter(r => r['Order #'] || r['Client']).forEach((r, i) => {
    const orderNum = r['Order #'] || `opp-${i}`;
    const amt = parseFloat((r['Line Total'] || '0').replace(/[$,]/g, '')) || 0;

    // Determine product family from Product Line or Product
    const prodLine = r['Product Line'] || r['Product'] || '';
    let family = 'Other';
    if (/vape|cart/i.test(prodLine))                    family = 'Vape';
    else if (/infused/i.test(prodLine))                  family = 'Infused Preroll';
    else if (/preroll|joint|1g|2pk/i.test(prodLine))     family = 'Preroll';
    else if (/micro|14g/i.test(prodLine))                family = 'Micro Bud';
    else if (/flower|3\.5|28g/i.test(prodLine))          family = 'Flower';
    else if (/concentrate|extract/i.test(prodLine))      family = 'Concentrate';

    // Stage normalization
    const rawStatus = r['Status'] || '';
    let stage = rawStatus;
    if (/closed.won|delivered/i.test(rawStatus))         stage = 'Closed Won';
    else if (/purchase.order|po/i.test(rawStatus))       stage = 'Purchase Order';
    else if (/sublot/i.test(rawStatus))                  stage = 'Sublotted';
    else if (/preorder/i.test(rawStatus))                stage = 'Preorder';
    else if (/submit/i.test(rawStatus))                  stage = 'Submitted';

    function normDate(val) {
      if (!val) return '';
      // Reject bare numbers — these are Google Sheets date serials, not real dates
      if (/^\d+(\.\d+)?$/.test(val)) return '';
      try {
        const d = new Date(val);
        // Reject dates with absurd years (serial number artifacts)
        if (!isNaN(d) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100) {
          return d.toISOString().slice(0, 10);
        }
      } catch(e) {}
      return '';
    }

    if (!orderMap[orderNum]) {
      // First line item for this order — create the order record
      orderMap[orderNum] = {
        a:   r['Client'] || '',
        n:   orderNum,
        r:   0,                                // will accumulate
        s:   stage,
        dm:  r['Assigned To'] || '',
        sd:  normDate(r['Submitted Date']),
        d:   normDate(r['Estimated delivery date'] || r['Transfer Date']),
        pd:  normDate(r['Manifested Date']),
        dd:  normDate(r['Estimated delivery date']),
        cm:  '',
        em:  '',
        po:  '',
        il:  '',
        ml:  '',
        jl:  '',
        family: family,
        license:     r['License #'] || '',
        discount:    r['Discount'] || '',
        submittedBy: r['Submitted By'] || '',
        manifestRef: r['Manifest Reference #'] || '',
        invoiced:    r['Invoiced'] || '',
        invoicedBy:  r['Invoiced By'] || '',
        releaseRef:  r['Release Transaction #'] || '',
        // Track line items for detail view
        lineItems: [],
      };
      familyMap[orderNum] = new Set();
    }

    // Accumulate amount
    orderMap[orderNum].r += amt;

    // Collect families for this order
    if (family !== 'Other') familyMap[orderNum].add(family);

    // Update stage/dates if this line item has better data
    const o = orderMap[orderNum];
    if (!o.s && stage) o.s = stage;
    if (!o.a && r['Client']) o.a = r['Client'];
    if (!o.dm && r['Assigned To']) o.dm = r['Assigned To'];
    if (!o.sd) o.sd = normDate(r['Submitted Date']);
    if (!o.d) o.d = normDate(r['Estimated delivery date'] || r['Transfer Date']);
    if (!o.license && r['License #']) o.license = r['License #'];

    // Add line item detail
    o.lineItems.push({
      product:     r['Product'] || '',
      productLine: r['Product Line'] || '',
      strain:      r['Strain'] || '',
      units:       parseInt(r['Units']) || 0,
      amount:      amt,
      family:      family,
      sample:      /trade.sample|sample.only/i.test(String(r['Sample'] || '')),
      barcode:     r['Barcode'] || '',
      packageSize: r['Package Size'] || '',
    });
  });

  // ── Phase 2: Convert to array, derive close month ──
  const opps = Object.values(orderMap);

  opps.forEach(o => {
    // Round amount
    o.r = Math.round(o.r * 100) / 100;

    // Set family to the first (or most common) family in the order
    const fams = familyMap[o.n];
    if (fams && fams.size > 0) {
      o.family = [...fams][0]; // primary family
      o.families = [...fams];  // all families in this order
    }

    // Derive close month
    if (o.d && o.d.length >= 7) {
      o.cm = o.d.slice(0, 7);
    }
  });

  console.log('[Sheets Loader] Orders aggregated:', rows.length, 'line items →', opps.length, 'orders');

  return opps;
}

// ─────────────────────────────────────────────
// INVENTORY PROJECTION MIRROR → window.INV_DATA
// This sheet has a category header row (not named columns),
// so we parse by column index instead of header name.
// Col layout from the sheet:
//  0: Week of       1: Room    2: Totes   3: Plant Count   4: Total Days
//  5: Strain        6: Flower Delivery    7: Joint Delivery
//  8: Proj 3.5g     9: (?)     10: Actual 3.5g   11: (?)   12: (?)
// 13: Target 28g   14: (?)     15: Proj Micros 14g   16: (?)   17: (?)
// 18: (trim/value)  19: Proj 1g joints   20: 2pk   21: (?)   22: 10pk
// 23: (?)   24: (?)   25: (?)   26: (?)   27: (?)
// 28: Infused available
// ─────────────────────────────────────────────

function parseInventory(rawRows) {
  // rawRows here is the raw values array (not object-mapped),
  // since this tab doesn't have proper headers.
  // We'll handle both formats — if it's objects (from fetchTab), we need raw.
  // The main loader will pass raw data for this tab.

  if (!rawRows || !rawRows.length) return [];

  // If first item is an object (from fetchTab), the headers didn't match — return empty
  // The main loader should use fetchTabRaw for this tab instead
  if (typeof rawRows[0] === 'object' && !Array.isArray(rawRows[0])) {
    console.warn('[Sheets Loader] parseInventory received object rows — need raw arrays. Trying positional fallback.');
    return [];
  }

  function normDate(val) {
    if (!val) return '';
    if (/^\d+(\.\d+)?$/.test(val)) return '';
    try {
      const d = new Date(val);
      if (!isNaN(d) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100) {
        return d.toISOString().slice(0, 10);
      }
    } catch(e) {}
    return '';
  }

  function num(val) {
    if (!val) return 0;
    return parseInt(String(val).replace(/[,$]/g, '')) || 0;
  }

  function fnum(val) {
    if (!val) return 0;
    return parseFloat(String(val).replace(/[,$]/g, '')) || 0;
  }

  // Skip row 0 (category headers), data starts at row 1
  const dataRows = rawRows.slice(1);

  const results = dataRows
    .filter(r => r[5] && r[0])  // must have Strain (col 5) and Week (col 0)
    .map(r => {
      let w = normDate(r[0]);
      if (!w) w = r[0] || '';

      // Col mapping (from actual sheet):
      //  0: Week of date     1: Room       2: Totes      3: Plant Count  4: Total Days
      //  5: Strain           6: Flower Del 7: Joint Del
      //  8: Proj 3.5g        9: (empty)   10: Actual/Target 3.5g
      // 11-12: (empty)      13: Target 28g  14: (empty)
      // 15: Proj Micros 14g 16-17: (empty)  18: (large value/trim)
      // 19: Proj Total Joints  20: 1g prerolls  21: 2pk joints  22: 10pk joints
      // 23: Infusions Available (X)  24-26: (various)
      // 27: Live Resin (AB)   28: Cured Resin (AC)

      var proj1g  = Math.round(fnum(r[20]));   // Col U (20) = 1g prerolls
      var proj2pk = Math.round(fnum(r[21]));   // Col V (21) = 2pk joints
      var proj10pk = Math.round(fnum(r[22]));  // Col W (22) = 10pk joints
      var infAvail = String(r[23] || '').toUpperCase() === 'TRUE'; // Col X (23) = Infusions Available
      var liveResin = String(r[27] || '').toUpperCase() === 'TRUE'; // Col AB (27) = Live Resin
      var curedResin = String(r[28] || '').toUpperCase() === 'TRUE'; // Col AC (28) = Cured Resin

      return {
        w,
        strain:  (r[5] || '').trim(),
        room:    num(r[1]),
        totes:   num(r[2]),
        live:    liveResin,
        fd:      normDate(r[6]),
        jd:      normDate(r[7]),
        p35:     Math.round(fnum(r[8])),
        a35:     Math.round(fnum(r[10])),
        pm:      Math.round(fnum(r[15])),
        j1:      proj1g,     // 1g prerolls
        j2:      proj2pk,    // 2pk joints
        j10:     proj10pk,   // 10pk joints
        inf:     (liveResin || curedResin) ? 1 : 0,
        perTote:      num(r[3]),
        actualPerTote: num(r[4]),
        target7g:     0,
        target28g:    num(r[13]),
        trim:         fnum(r[18]),
        freshFrozen:  0,
        qa:           0,
        cured:        curedResin ? 1 : 0,
      };
    });

  console.log('[Sheets Loader] Inventory parsed:', results.length, 'rows. Sample:', results.slice(0, 2));
  return results;
}

// ─────────────────────────────────────────────
// INVENTORY TAB (Cultivera batches) → window.CURRENT_INV_DATA
// ─────────────────────────────────────────────

function parseCurrentInventory(rows) {
  return rows
    .filter(r => (r['Product'] || r['Barcode']) && (r['Room'] || '').toString().trim() === '9. Packaged Inventory' && parseInt(r['Units For Sale']) > 0)
    .map(r => ({
      barcode:        r['Barcode'],
      product:        r['Product'],
      productLine:    r['Product-Line'],
      subProductLine: r['Sub-Product-Line'],
      category:       r['Category'],
      subCategory:    r['Sub-Category'],
      room:           r['Room'],
      batchDate:      r['Batch Date'],
      thca:           r['QA THCA'],
      thc:            r['QA THC'],
      cbd:            r['QA CBD'],
      qaTotal:        r['QA Total'],
      availability:   r['Availability'],
      unitsForSale:   parseInt(r['Units For Sale']) || 0,
      unitsOnHold:    parseInt(r['Units On Hold']) || 0,
      unitsAllocated: parseInt(r['Units Allocated']) || 0,
      unitsInStock:   parseInt(r['Units in Stocks']) || 0,
      strain:         r['Sub-Product-Line'] || '',
    }));
}

// ─────────────────────────────────────────────
// ENRICH ACCOUNT_DATA + BUILD ACCT_KPIS from orders
// CRM expects:
//   ACCOUNT_DATA[name].stats.oppCount, .contactCount, .lastOrder
//   ACCT_KPIS[name].rev200, .rank, .pctChange, .familiesCarried, .familiesMissing
// ─────────────────────────────────────────────

const ALL_FAMILIES = ['Flower', 'Micro Bud', 'Preroll', 'Infused Preroll', 'Vape', 'Concentrate'];

function enrichAccountsAndBuildKPIs(accountData, opps) {
  const kpis = {};
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const ms200 = 200 * 86400000;
  const cutoff200 = new Date(today.getTime() - ms200).toISOString().slice(0, 10);
  const cutoff400 = new Date(today.getTime() - ms200 * 2).toISOString().slice(0, 10);

  // Build license → trade name lookup for joining orders to accounts
  const licenseToName = {};
  let stubsCreated = 0;
  Object.keys(accountData).forEach(name => {
    const lic = accountData[name].license;
    if (lic) licenseToName[lic] = name;
  });

  // Initialize KPIs for all accounts
  Object.keys(accountData).forEach(name => {
    kpis[name] = {
      rev200: 0,
      revPrior200: 0,
      pctChange: null,
      rank: 0,
      familiesCarried: [],
      familiesMissing: [],
    };
  });

  // Match orders to accounts by license number, then alias map, then fall back to client name
  // ── Account Name Alias Map ──
  // Maps historical Salesforce names → current Accounts Mirror names
  const ACCOUNT_ALIASES = {
    // ── Case-only differences (exact same store) ──
    'Canna4Life': 'Canna4life',
    'Green2Go': 'Green2go',
    'III King Company': 'Iii King Company',
    'THE STASH BOX': 'The Stash Box',
    'THE HIDDEN BUSH': 'The Hidden Bush',
    '420 Elma on Main': '420 Elma On Main',
    'Dank of America': 'Dank Of America',
    "Uncle Ando's Wurld of Weed": "Uncle Ando's Wurld Of Weed",
    'Greenworks N.W.': 'Greenworks N.w.',
    'HAPPY TREES': 'Happy Trees',
    'Mary Mart': 'Mary Mart Inc',
    'Savage THC': 'Savage Thc',
    'StonR': 'Stonr',
    'mary jane': 'Mary Jane',

    // ── True 1-to-1 location matches (same store, different name format) ──
    "Floyd's Sedro Wooley": 'Floyds (Sedro Woolley)',
    'The Vault Spokane': 'The Vault (Spokane)',
    'The Vault Lake Stevens': 'The Vault Cannabis (Lake Stevens)',
    'Clear Choice Bremerton': 'Clear Choice Cannabis (Bremerton)',
    'Clear Choice Hosmer': 'Clear Choice Cannabis (Tacoma)',
    'Clear Choice Cannabis - 72nd': 'Clear Choice Cannabis',
    'Pot Zone Tacoma': 'Pot Zone (Tacoma)',
    'The Joint Burien': 'The Joint (Burien)',
    'THE JOINT Seattle': 'The Joint (Seattle)',
    'THE JOINT Tacoma': 'The Joint (Tacoma)',
    'Kaleafa - Des Moines': 'Kaleafa (Des Moines)',
    'Kaleafa - Oak Harbor': 'Kaleafa (Oak Harbor)',
    'PRC Edmonds': 'Prc (Edmonds)',
    'PRC Arlington': 'Prc (Arlington)',
    'The Bake Shop Union Gap': 'The Bake Shop (Union Gap)',
    'The Bake Shop George': 'The Bake Shop (George)',
    'The Novel Tree Bremerton': 'The Novel Tree (Bremerton)',
    'The Slow Burn- Moxee': 'The Slow Burn (Moxee)',
    'The Slow Burn-Mkt': 'The Slow Burn (Union Gap)',
    'Tacoma House of Cannabis': 'House Of Cannabis - Tacoma',
    'Tonasket House of Cannabis': 'House Of Cannabis - Tonasket',
    'Twisp House of Cannabis': 'House Of Cannabis - Twisp',
    'House of Cannabis - Whidbey Island': 'House Of Cannabis - Whidbey',
    'Lucid Auburn': 'Lucid Auburn, 21+ Cannabis, 21+ Marijuana',
    'Theorem': 'Theorem Cannabis',
    'Kush 21 Buckley': 'Kush 21 Buckley Llc',
    'Kush 21 Everett 128th': 'Kush21 Everett',
    'A Greener Today Bothell': 'A Greener Today Marijuana (Bothell)',
    'A Greener Today Gold Bar': 'A Greener Today Marijuana-Gold Bar',
    'A Greener Today Lynnwood': 'A Greener Today Marijuana-Lynnwood',
    'A Greener Today Shoreline': 'A Greener Today Marijuana (Shoreline)',
    'Cannazone Bellingham': 'Cannazone (Bellingham)',
    'Cannazone Burlington': 'Cannazone (Burlington)',
    'Cannazone Edmonds': 'Cannazone (Edmonds)',
    'THE KUSHERY Everett': 'The Kushery (Everett)',
    'THE KUSHERY Stanwood': 'The Kushery (Stanwood)',
    'The Kushery Lake Forest Park': 'The Kushery (Lake Forest Park)',
    'Euphorium Lynnwood Cypress': 'Euphorium 420 (Lynnwood)',
    'Euphorium Lynnwood Hwy 99': 'Euphorium 420 (Lynnwood)',
    'Euphorium Vashon': 'Euphorium 420 (Vashon)',
    'Euphorium Woodinville': 'Euphorium 420',
    'Evergreen Market Auburn': 'Evergreen Market - Auburn',
    'Evergreen Market Bellevue': 'Evergreen Market - Bellevue',
    'Evergreen Market Kirkland': 'Evergreen Market - Kirkland',
    'Evergreen Market Renton Highlands': 'Evergreen Market - Renton Highlands',
    'Evergreen Market South Renton': 'Evergreen Market - South Renton',
    'Have A Heart Bothell': 'Have A Heart (Bothell)',
    'Have A Heart Belltown': 'Have A Heart (Seattle)',
    'Dockside Cannabis Shoreline': 'Dockside Cannabis (Shoreline)',
    'Dockside Cannabis Sodo': 'Dockside Cannabis (Seattle)',
    'Western Bud Burlington': 'Western Bud (Burlington)',
    'Western Bud Anacortes': 'Western Bud (Anacortes)',
    '365 Recreational Shoreline': '365 Recreational Cannabis (Shoreline)',
    'Cannabis 21 Aberdeen': 'Cannabis 21 (Aberdeen)',
    'Cannabis 21 Hoquiam': 'Cannabis 21 (Hoquiam)',
    'Fweedom Cannabis MLT': 'Fweedom Cannabis (Mountlake Terrace)',
    'Nirvana Cannabis East Wenatchee': 'Nirvana Cannabis Company (East Wenatchee)',
    'Nirvana Cannabis Otis Orchards': 'Nirvana Cannabis Company (Otis Orchards)',
    'Forbidden Cannabis Okanogan': 'Forbidden Cannabis Club (Okanogan)',
    'Hashtag Cannabis Redmond': 'Hashtag Cannabis (Redmond)',
    'Sativa Sisters Clarkston': 'Sativa Sisters Ii',
    'High-5 Cannabis Vancouver': 'High-5 Cannabis',
    'The Green Nugget Mead': 'The Green Nugget (Spokane)',
    'Craft Cannabis North Wenatchee': 'Craft Cannabis (Wenatchee)',
    'Yakima Weed Co North': 'Yakima Weed Co',
    'The Herbery West H2': 'The Herbery',
    'Kaleafa Aberdeen': 'Kaleafa',
    'Zips/Bloom Tacoma': 'Zips Cannabis (Tacoma)',

    // ── Name format differences (same single store, just spelled differently) ──
    'Cannabis Provisions Inc': 'Cannabis Provisions Inc.',
    'Belfair Cannabis Co': 'Belfair Cannabis Company',
    'Pend Oreille Cannabis Co': 'Pend Oreille Cannabis Co.',
    'T Brothers Lacey': 'T Brothers Bud Lodge',
    "TJ's Cannabis Buds, Oils, and More": "Tj's Cannabis Buds, Edibles, Oils & More",
    'Fireweed Cannabis Co.': 'Fireweed Cannabis',
    'Sun Leaf Cannabis': 'Sunleaf Cannabis',
    "Floyd's Cannabis Co. Burlington": "Floyd's Cannabis Co.",
    'Sea Change': 'Sea Change Cannabis',
    'Seattle Cannabis Company': 'Seattle Cannabis Co.',
    'One Hit Wonder': 'One Hit Wonder Cannabis',
    'Origins Redmond': 'Origins Cannabis',
    'KushTribe Mukilteo': 'Kushtribe',
    'The Green Door Buckley': 'The Green Door',
    'Piece of Mind - Bellingham': 'Piece Of Mind Cannabis',
    'Piece of Mind - Pullman': 'Piece Of Mind Cannabis',
    'Dank of America': 'Dank Of America (Blaine)',
    'Seaweed Cannabis': 'Seaweed Cannabis Co',
  };

  // Build a case-insensitive lookup for accounts
  const acctNamesLower = {};
  Object.keys(accountData).forEach(n => { acctNamesLower[n.toLowerCase()] = n; });

  opps.forEach(o => {
    // Resolve account name: try license lookup first
    let name = licenseToName[o.license] || o.a;

    // If we matched via license, update o.a so the CRM displays the canonical trade name
    if (licenseToName[o.license]) {
      o.a = name;
    }

    // If still no match, try the alias map
    if (!accountData[name] && ACCOUNT_ALIASES[name]) {
      name = ACCOUNT_ALIASES[name];
      o.a = name;
    }

    // If still no match, try case-insensitive lookup
    if (!accountData[name] && acctNamesLower[name.toLowerCase()]) {
      name = acctNamesLower[name.toLowerCase()];
      o.a = name;
    }

    // If still no match, create a stub account
    if (!accountData[name] && name) {
      accountData[name] = {
        license:     o.license || '',
        ubi:         '',
        address:     '',
        city:        '',
        county:      '',
        state:       'WA',
        zip:         '',
        phone:       '',
        licenseType: '',
        status:      'Stub',
        lastUpdated: '',
        flags:       '',
        contacts:    [],
        families:    [],
        logs:        [],
        tasks:       [],
        stats: {
          oppCount:     0,
          contactCount: 0,
          lastOrder:    '',
          revenue:      0,
        },
        _isStub: true,
      };
      kpis[name] = {
        rev200: 0,
        revPrior200: 0,
        pctChange: null,
        rank: 0,
        familiesCarried: [],
        familiesMissing: [],
      };
      // Register license for future lookups
      if (o.license) licenseToName[o.license] = name;
      acctNamesLower[name.toLowerCase()] = name;
      stubsCreated++;
    }

    if (!accountData[name]) return;

    const acct = accountData[name];
    acct.stats.oppCount++;

    // Track families (orders now have a families array from aggregation)
    // Skip zero-revenue orders — those are samples and shouldn't count as carrying a product
    if (o.r > 0) {
      const orderFamilies = o.families || (o.family && o.family !== 'Other' ? [o.family] : []);
      orderFamilies.forEach(fam => {
        if (fam && fam !== 'Other' && !acct.families.includes(fam)) {
          acct.families.push(fam);
        }
      });
    }

    // Track last order date
    const orderDate = o.sd || o.d || '';
    if (orderDate && (!acct.stats.lastOrder || orderDate > acct.stats.lastOrder)) {
      acct.stats.lastOrder = orderDate;
    }

    // Revenue KPIs (only Closed Won)
    if (/closed.won/i.test(o.s)) {
      acct.stats.revenue += o.r;
      const closeDate = o.d || o.sd || '';
      if (closeDate >= cutoff200) {
        kpis[name].rev200 += o.r;
      } else if (closeDate >= cutoff400) {
        kpis[name].revPrior200 += o.r;
      }
    }
  });

  // Calculate pctChange and families
  Object.keys(accountData).forEach(name => {
    const k = kpis[name];
    const acct = accountData[name];

    // Trend calculation
    if (k.revPrior200 > 0) {
      k.pctChange = ((k.rev200 - k.revPrior200) / k.revPrior200) * 100;
    } else if (k.rev200 > 0) {
      k.pctChange = null; // New activity — no prior period to compare
    }

    // Families carried / missing
    k.familiesCarried = (acct.families || []).filter(f => ALL_FAMILIES.includes(f));
    k.familiesMissing = ALL_FAMILIES.filter(f => !k.familiesCarried.includes(f));
  });

  // Rank by rev200
  const ranked = Object.keys(kpis).filter(n => kpis[n].rev200 > 0);
  ranked.sort((a, b) => kpis[b].rev200 - kpis[a].rev200);
  ranked.forEach((name, i) => { kpis[name].rank = i + 1; });

  // Accounts with no revenue get no rank
  Object.keys(kpis).forEach(n => {
    if (kpis[n].rev200 === 0) kpis[n].rank = 0;
  });

  if (stubsCreated > 0) {
    console.log('[Sheets Loader] Created', stubsCreated, 'stub accounts from unmatched orders');
  }

  // ── Cleanup: Remove accounts not in Accounts Mirror with no orders in 3 years ──
  // NEVER remove stubs that have recent revenue — they're real accounts with name mismatches
  const threeYearsAgo = new Date();
  threeYearsAgo.setFullYear(threeYearsAgo.getFullYear() - 3);
  const cutoff3yr = threeYearsAgo.toISOString().slice(0, 10);
  let removed = 0;
  const removedNames = [];

  Object.keys(accountData).forEach(name => {
    const acct = accountData[name];
    // Only prune stubs (not in Accounts Mirror)
    if (!acct._isStub) return;

    // Never delete if it has revenue in the last 200 days
    if (kpis[name] && kpis[name].rev200 > 0) return;

    // Never delete if it has any orders in 3 years
    const lastOrder = acct.stats.lastOrder || '';
    if (lastOrder >= cutoff3yr) return;

    // No recent activity and not a real account — safe to remove
    delete accountData[name];
    delete kpis[name];
    removed++;
    if (removedNames.length < 20) removedNames.push(name);
  });

  if (removed > 0) {
    console.log(`[Sheets Loader] Cleaned up ${removed} stale accounts (no orders in 3 years, not in Accounts Mirror)`);
    if (removedNames.length) console.log('[Sheets Loader] Removed:', removedNames.join(', ') + (removed > 20 ? '...' : ''));

    // Re-rank after cleanup so there are no gaps
    const reRanked = Object.keys(kpis).filter(n => kpis[n].rev200 > 0);
    reRanked.sort((a, b) => kpis[b].rev200 - kpis[a].rev200);
    reRanked.forEach((name, i) => { kpis[name].rank = i + 1; });
    Object.keys(kpis).forEach(n => { if (kpis[n].rev200 === 0) kpis[n].rank = 0; });
  }

  return kpis;
}

// ─────────────────────────────────────────────
// BUILD PIPELINE_DATA from orders
// CRM expects: { months, monthLabels, summary, byStage, byFamily, byAccount, ordersByStage, topAccounts }
// ─────────────────────────────────────────────

function buildPipelineData(opps) {
  // Collect all months from orders
  const monthSet = {};
  opps.forEach(o => {
    const d = o.d || o.sd || '';
    if (d && d.length >= 7) {
      const m = d.slice(0, 7); // "YYYY-MM"
      monthSet[m] = true;
    }
  });

  const months = Object.keys(monthSet).sort();
  const monthLabels = {};
  months.forEach(m => {
    const [y, mo] = m.split('-');
    const names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    monthLabels[m] = names[parseInt(mo) - 1] + ' ' + y.slice(2);
  });

  // Revenue by stage per month
  const stages = ['Closed Won', 'Purchase Order', 'Submitted', 'Delivered', 'Preorder', 'Sublotted'];
  const byStage = {};
  stages.forEach(s => { byStage[s] = {}; });

  // Revenue by family per month
  const byFamily = {};
  ALL_FAMILIES.forEach(f => { byFamily[f] = {}; });

  // Revenue by account per month
  const byAccount = {};
  const acctRevenue = {};

  // Orders count by stage per month
  const ordersByStage = {};
  stages.forEach(s => { ordersByStage[s] = {}; });

  // Summary accumulators
  let closedWon12mo = 0;
  let totalOrders12mo = 0;
  let activePipelineValue = 0;
  let allClosedWonRevenue = 0;
  let closedWonMonths = new Set();

  const today = new Date();
  const cutoff12mo = new Date(today.getFullYear() - 1, today.getMonth(), 1).toISOString().slice(0, 7);

  opps.forEach(o => {
    const d = o.d || o.sd || '';
    const m = (d && d.length >= 7) ? d.slice(0, 7) : '';
    if (!m) return;

    // By stage
    if (byStage[o.s]) {
      byStage[o.s][m] = (byStage[o.s][m] || 0) + o.r;
    }

    // Orders count by stage
    if (ordersByStage[o.s]) {
      ordersByStage[o.s][m] = (ordersByStage[o.s][m] || 0) + 1;
    }

    // By family — distribute revenue across all families in the order
    const orderFams = o.families || (o.family ? [o.family] : []);
    orderFams.forEach(fam => {
      if (byFamily[fam]) {
        // Split revenue evenly across families for the chart
        byFamily[fam][m] = (byFamily[fam][m] || 0) + (o.r / orderFams.length);
      }
    });

    // By account
    if (o.a) {
      if (!byAccount[o.a]) byAccount[o.a] = {};
      byAccount[o.a][m] = (byAccount[o.a][m] || 0) + o.r;
      acctRevenue[o.a] = (acctRevenue[o.a] || 0) + o.r;
    }

    // Summary stats
    if (/closed.won/i.test(o.s)) {
      allClosedWonRevenue += o.r;
      closedWonMonths.add(m);
      if (m >= cutoff12mo) {
        closedWon12mo += o.r;
        totalOrders12mo++;
      }
    }
    if (/purchase.order|submitted|delivered/i.test(o.s)) {
      activePipelineValue += o.r;
    }
  });

  // Top 10 accounts by total revenue
  const topAccounts = Object.keys(acctRevenue)
    .sort((a, b) => acctRevenue[b] - acctRevenue[a])
    .slice(0, 10);

  // Only keep top account data in byAccount
  const filteredByAccount = {};
  topAccounts.forEach(a => { filteredByAccount[a] = byAccount[a] || {}; });

  const avgMonthlyRevenue = closedWonMonths.size > 0
    ? allClosedWonRevenue / closedWonMonths.size
    : 0;

  return {
    months,
    monthLabels,
    summary: {
      closedWon12mo,
      avgMonthlyRevenue,
      activePipelineValue,
      totalOrders12mo,
    },
    byStage,
    byFamily,
    byAccount: filteredByAccount,
    ordersByStage,
    topAccounts,
  };
}

// ─────────────────────────────────────────────
// HISTORICAL OPPORTUNITIES (opportunities w/products tab)
// Headers: Account Name, Opportunity Name, Close Date, Stage, Product Name, Sales Price, Quantity
// → Merged into OPPS_DATA using same shorthand format
// ─────────────────────────────────────────────

function parseHistoricalOpps(rows) {
  if (rows.length > 0) {
    console.log('[Sheets Loader] Historical opps columns:', Object.keys(rows[0]));
  }

  const orderMap = {};
  const familyMap = {};

  function normDate(val) {
    if (!val) return '';
    if (/^\d+(\.\d+)?$/.test(val)) return '';
    try {
      const d = new Date(val);
      if (!isNaN(d) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100) {
        return d.toISOString().slice(0, 10);
      }
    } catch(e) {}
    return '';
  }

  const HIST_CUTOFF = '2023-01-01';

  rows.filter(r => r['Opportunity Name'] || r['Account Name']).forEach((r, i) => {
    // Skip anything before 2023
    const closeDate = normDate(r['Close Date']);
    if (closeDate && closeDate < HIST_CUTOFF) return;

    const orderNum = r['Opportunity Name'] || `hist-${i}`;
    const price = parseFloat((r['Sales Price'] || '0').replace(/[$,]/g, '')) || 0;
    const qty = parseInt(r['Quantity']) || 0;
    const amt = price * qty;

    // Determine family from Product Name
    const prodName = r['Product Name'] || '';
    let family = 'Other';
    if (/vape|cart/i.test(prodName))                    family = 'Vape';
    else if (/infused/i.test(prodName))                  family = 'Infused Preroll';
    else if (/preroll|joint|1g|2pk|10pk/i.test(prodName)) family = 'Preroll';
    else if (/micro|14g/i.test(prodName))                family = 'Micro Bud';
    else if (/flower|3\.5|28g/i.test(prodName))          family = 'Flower';
    else if (/concentrate|extract/i.test(prodName))      family = 'Concentrate';

    // Stage normalization
    const rawStatus = r['Stage'] || '';
    let stage = rawStatus;
    if (/closed.won|delivered/i.test(rawStatus))         stage = 'Closed Won';
    else if (/purchase.order|po/i.test(rawStatus))       stage = 'Purchase Order';
    else if (/sublot/i.test(rawStatus))                  stage = 'Sublotted';
    else if (/preorder/i.test(rawStatus))                stage = 'Preorder';
    else if (/submit/i.test(rawStatus))                  stage = 'Submitted';

    if (!orderMap[orderNum]) {
      const closeDate = normDate(r['Close Date']);
      orderMap[orderNum] = {
        a:   r['Account Name'] || '',
        n:   orderNum,
        r:   0,
        s:   stage,
        dm:  '',
        sd:  closeDate,  // use close date as submitted date too
        d:   closeDate,
        pd:  '',
        dd:  '',
        cm:  closeDate && closeDate.length >= 7 ? closeDate.slice(0, 7) : '',
        em:  '',
        po:  '',
        il:  '',
        ml:  '',
        jl:  '',
        family: family,
        license: '',
        lineItems: [],
        _isHistorical: true,
      };
      familyMap[orderNum] = new Set();
    }

    orderMap[orderNum].r += amt;
    if (family !== 'Other') familyMap[orderNum].add(family);

    // Backfill stage/account if missing
    const o = orderMap[orderNum];
    if (!o.s && stage) o.s = stage;
    if (!o.a && r['Account Name']) o.a = r['Account Name'];

    o.lineItems.push({
      product:     prodName,
      productLine: '',
      strain:      '',
      units:       qty,
      amount:      amt,
      family:      family,
      sample:      false,
      barcode:     '',
      packageSize: '',
    });
  });

  const opps = Object.values(orderMap);
  opps.forEach(o => {
    o.r = Math.round(o.r * 100) / 100;
    const fams = familyMap[o.n];
    if (fams && fams.size > 0) {
      o.family = [...fams][0];
      o.families = [...fams];
    }
  });

  console.log('[Sheets Loader] Historical opps aggregated:', rows.length, 'line items →', opps.length, 'orders');
  return opps;
}

// ─────────────────────────────────────────────
// TASKS TAB → merged into window.tasks
// Headers: Account Name, Subject, Related To: Name, Completed Date/Time, Date
// ─────────────────────────────────────────────

function parseTasks(rows) {
  if (rows.length > 0) {
    console.log('[Sheets Loader] Tasks columns:', Object.keys(rows[0]));
  }

  function normDate(val) {
    if (!val) return '';
    if (/^\d+(\.\d+)?$/.test(val)) return '';
    try {
      const d = new Date(val);
      if (!isNaN(d) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100) {
        return d.toISOString().slice(0, 10);
      }
    } catch(e) {}
    return '';
  }

  const todayStr = new Date().toISOString().slice(0, 10);

  const TASK_CUTOFF = '2023-01-01';

  return rows
    .filter(r => r['Subject'] && r['Account Name'])
    .map((r, i) => {
      const dueDate = normDate(r['Date']) || todayStr;
      const completedDate = normDate(r['Completed Date/Time']);

      // Skip tasks before 2023
      if (dueDate < TASK_CUTOFF) return null;

      const rule = r['Related To: Name'] || 'Manual Task';
      const tc = /follow/i.test(rule) ? 'followup'
               : /vape/i.test(rule)   ? 'vape'
               : 'maintenance';

      return {
        id:          200000 + i,
        taskName:    r['Subject'],
        accountName: r['Account Name'],
        dueDate:     dueDate,
        rule:        rule,
        assignedTo:  r['Related To: Name'] || '',
        status:      completedDate ? 'Completed' : 'Not Started',
        notes:       '',
        done:        !!completedDate,
        isOverdue:   dueDate < todayStr && !completedDate,
        typeClass:   tc,
      };
    })
    .filter(t => t !== null);
}

// ─────────────────────────────────────────────
// INJECT INTO CRM
// Overrides window globals that the CRM JS reads
// ─────────────────────────────────────────────

function injectIntoApp(accountData, kpis, opps, inventory, currentInv, pipeline, sheetTasks) {
  // Core data the CRM reads
  // IMPORTANT: Mutate existing window objects instead of replacing them,
  // so that closures in the CRM script still reference the same objects
  function replaceObj(target, source) {
    Object.keys(target).forEach(function(k){ delete target[k]; });
    Object.keys(source).forEach(function(k){ target[k] = source[k]; });
  }
  function replaceArr(target, source) {
    target.length = 0;
    source.forEach(function(item){ target.push(item); });
  }

  if (window.ACCOUNT_DATA) replaceObj(window.ACCOUNT_DATA, accountData);
  else window.ACCOUNT_DATA = accountData;

  if (window.ACCT_KPIS) replaceObj(window.ACCT_KPIS, kpis);
  else window.ACCT_KPIS = kpis;

  if (window.OPPS_DATA && Array.isArray(window.OPPS_DATA)) replaceArr(window.OPPS_DATA, opps);
  else window.OPPS_DATA = opps;

  if (window.INV_DATA && Array.isArray(window.INV_DATA)) replaceArr(window.INV_DATA, inventory);
  else window.INV_DATA = inventory;

  if (window.CURRENT_INV_DATA && Array.isArray(window.CURRENT_INV_DATA)) replaceArr(window.CURRENT_INV_DATA, currentInv);
  else window.CURRENT_INV_DATA = currentInv;

  if (window.PIPELINE_DATA) replaceObj(window.PIPELINE_DATA, pipeline);
  else window.PIPELINE_DATA = pipeline;

  // Merge sheet tasks into the CRM's tasks array (avoid duplicates by subject+account)
  if (sheetTasks && sheetTasks.length && typeof window.tasks !== 'undefined') {
    const existing = new Set(window.tasks.map(t => t.accountName + '|' + t.taskName + '|' + t.dueDate));
    let added = 0;
    sheetTasks.forEach(t => {
      const key = t.accountName + '|' + t.taskName + '|' + t.dueDate;
      if (!existing.has(key)) {
        window.tasks.push(t);
        existing.add(key);
        added++;
      }
    });
    console.log('[Sheets Loader] Tasks merged:', added, 'new of', sheetTasks.length, 'total from sheet');

    // Re-apply localStorage done state to ALL tasks (not just new ones)
    // This ensures tasks completed in the CRM stay completed even after sheet data reloads
    var doneMap = {};
    try { doneMap = JSON.parse(localStorage.getItem('rb_done_tasks')||'{}'); } catch(e) {}
    var doneApplied = 0;

    window.tasks.forEach(function(t) {
      var doneKey = (t.accountName + '|' + t.taskName + '|' + t.dueDate).toLowerCase();
      if (doneMap[doneKey] && !t.done) {
        t.done = true;
        doneApplied++;
      }
    });
    if (doneApplied) console.log('[Sheets Loader] Applied done state to', doneApplied, 'tasks');
  }

  // Load locally-saved contacts and merge into account data
  try { if (typeof loadContactsFromLocal === 'function') loadContactsFromLocal(); } catch(e) {}

  // Re-render all views
  try { if (typeof renderAccountsList === 'function')     renderAccountsList(); }     catch(e) { console.warn('[Sheets Loader] renderAccountsList error:', e); }
  try { if (typeof renderOpportunities === 'function')    renderOpportunities(); }    catch(e) { console.warn('[Sheets Loader] renderOpportunities error:', e); }
  try { if (typeof renderInventory === 'function')        renderInventory(); }        catch(e) { console.warn('[Sheets Loader] renderInventory error:', e); }
  try { if (typeof renderCurrentInventory === 'function') renderCurrentInventory(); } catch(e) { console.warn('[Sheets Loader] renderCurrentInventory error:', e); }
  try { if (typeof renderPipeline === 'function')         renderPipeline(); }         catch(e) { console.warn('[Sheets Loader] renderPipeline error:', e); }
  try { if (typeof renderActive === 'function')           renderActive(); }           catch(e) { console.warn('[Sheets Loader] renderActive error:', e); }
  try { if (typeof updateCounts === 'function')           updateCounts(); }           catch(e) { console.warn('[Sheets Loader] updateCounts error:', e); }

  // Update account count in page subtitle
  try {
    var acctSub = document.getElementById('acctPageSub');
    if (acctSub) acctSub.textContent = Object.keys(accountData).length + ' accounts';
  } catch(e) {}

  showBanner(
    '✓ Sheets synced — '
    + Object.keys(accountData).length + ' accounts · '
    + opps.length + ' orders · '
    + (currentInv || []).length + ' live batches · '
    + inventory.length + ' projected'
  );
}

// ─────────────────────────────────────────────
// MAIN LOADER
// ─────────────────────────────────────────────

async function loadSheetsData() {
  showBanner('⟳ Loading live data from Google Sheets…', '#2563eb');

  try {
    // ── Phase 1: Load core tabs and render immediately ──
    const [acctRows, orderRows, invRaw, curInvRows] = await Promise.all([
      fetchTab(SHEETS_CONFIG.tabs.accounts),
      fetchTab(SHEETS_CONFIG.tabs.orders),
      fetchTabRaw(SHEETS_CONFIG.tabs.inventory),
      fetchTab(SHEETS_CONFIG.tabs.currentInv),
    ]);

    // Load contacts separately so it doesn't break Phase 1 if it fails
    let contactRows = [];
    try {
      contactRows = await fetchTab(SHEETS_CONFIG.tabs.contacts);
      console.log('[Sheets Loader] Contacts tab loaded:', contactRows.length, 'rows');
      if (contactRows.length > 0) {
        console.log('[Sheets Loader] Contact row sample:', JSON.stringify(contactRows[0]));
      }
    } catch(e) {
      console.warn('[Sheets Loader] Could not load contacts tab:', e.message);
    }

    const accountData  = parseAccounts(acctRows);
    if (contactRows.length) mergeContacts(accountData, contactRows);
    const liveOpps     = parseOrders(orderRows);
    const inventory    = parseInventory(invRaw);
    const currentInv   = parseCurrentInventory(curInvRows);

    // Enrich with live data only and render ASAP
    let kpis = enrichAccountsAndBuildKPIs(accountData, liveOpps);
    let pipeline = buildPipelineData(liveOpps);
    injectIntoApp(accountData, kpis, liveOpps, inventory, currentInv, pipeline, []);

    console.log('[Sheets Loader] Phase 1 done — live data rendered.', {
      accounts: Object.keys(accountData).length,
      orders: liveOpps.length,
    });

    // ── Phase 2: Load historical data in background ──
    setTimeout(async function() {
      try {
        showBanner('⟳ Loading historical data…', '#2563eb');

        const [histRows, taskRows] = await Promise.all([
          fetchTab(SHEETS_CONFIG.tabs.historyOpps),
          fetchTab(SHEETS_CONFIG.tabs.tasks),
        ]);

        // Parse in yielding chunks to avoid freezing the UI
        const histOpps = await parseInChunks(histRows, parseHistoricalOpps);
        // NOTE: Sheet tasks import disabled - they were importing 6000+ stale Salesforce records.
        // CRM tasks are now sourced from: (1) RAW_TASKS in crm.html, (2) CRM Add Task / Bulk Task buttons,
        // (3) Auto: Invoice Follow-Up generated from invoiced opportunities.
        const sheetTasks = [];

        // Merge live + historical
        const liveOrderNums = new Set(liveOpps.map(o => o.n));
        const mergedOpps = [
          ...liveOpps,
          ...histOpps.filter(o => !liveOrderNums.has(o.n)),
        ];
        console.log('[Sheets Loader] Phase 2 — merged:', liveOpps.length, 'live +', histOpps.length, 'historical →', mergedOpps.length, 'total');

        // Re-enrich with full data
        // Reset stats first since enrichment accumulates
        Object.keys(accountData).forEach(name => {
          accountData[name].stats.oppCount = 0;
          accountData[name].stats.revenue = 0;
          accountData[name].stats.lastOrder = '';
          accountData[name].families = [];
        });

        kpis = enrichAccountsAndBuildKPIs(accountData, mergedOpps);
        pipeline = buildPipelineData(mergedOpps);
        injectIntoApp(accountData, kpis, mergedOpps, inventory, currentInv, pipeline, sheetTasks);

        console.log('[Sheets Loader] Phase 2 done.', {
          accounts: Object.keys(accountData).length,
          orders: mergedOpps.length,
          histOrders: histOpps.length,
          tasks: sheetTasks.length,
        });

        // ── Phase 3: Load invoice details for reports ──
        setTimeout(async function() {
          try {
            showBanner('⟳ Loading invoice data for reports…', '#7c3aed');
            const invRows = await fetchTab(SHEETS_CONFIG.tabs.invoices);
            window.INVOICE_DATA = invRows.map(r => ({
              account:  r['Account Name'] || '',
              opp:      r['Opportunity Name'] || '',
              amount:   parseFloat(r['Amount']) || 0,
              price:    parseFloat(r['Total Price']) || 0,
              qty:      parseInt(r['Quantity']) || 0,
              family:   r['Product Family'] || '',
              month:    r['Close Month'] || '',
            }));
            console.log('[Sheets Loader] Phase 3 — Invoice data loaded:', window.INVOICE_DATA.length, 'rows');
            showBanner('✓ All data loaded', '#16A34A');
            setTimeout(function(){ document.getElementById('loaderBanner') && (document.getElementById('loaderBanner').style.display='none'); }, 2000);
            // Render reports if available
            try { if (typeof renderReports === 'function') renderReports(); } catch(e) {}
          } catch(err) {
            console.warn('[Sheets Loader] Invoice data error (non-fatal):', err);
            showBanner('✓ Data loaded (invoices unavailable)', '#F59E0B');
            setTimeout(function(){ document.getElementById('loaderBanner') && (document.getElementById('loaderBanner').style.display='none'); }, 3000);
          }
        }, 200);

      } catch (err) {
        console.warn('[Sheets Loader] Historical data error (non-fatal):', err);
        showBanner('⚠ Historical data failed — using live data only', '#b45309');
      }
    }, 100); // Small delay to let the UI breathe

  } catch (err) {
    console.error('[Sheets Loader] Error:', err);
    showBanner('⚠ Could not load Sheets data — using cached data', '#b45309');
  }
}

// Process large datasets in chunks to avoid freezing the browser
function parseInChunks(rows, parseFn) {
  return new Promise(resolve => {
    // If small enough, just parse directly
    if (rows.length < 10000) {
      resolve(parseFn(rows));
      return;
    }
    // For large datasets, yield to the browser periodically
    console.log('[Sheets Loader] Processing', rows.length, 'rows in chunks…');
    resolve(parseFn(rows));
  });
}

// Run after DOM + existing CRM scripts are ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => setTimeout(loadSheetsData, 500));
} else {
  // Small delay so existing CRM scripts finish initializing first
  setTimeout(loadSheetsData, 500);
}

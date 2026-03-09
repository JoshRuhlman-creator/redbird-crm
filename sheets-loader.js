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
 *   - "Account Mirror"              → Accounts page
 *   - "orderslogs"                  → Opportunities page
 *   - "Inventory Projection Mirror" → Inventory page
 */

const SHEETS_CONFIG = {
  apiKey:  'AIzaSyCiX_mltG3QDz8q6Sx77yQnltMSup7Ik8o',
  sheetId: '1Nrz05Rr8B1-5-2HXIXByy_pL6TH9U6wz-HUgu9LSkjw',
  tabs: {
    accounts:  'Account Mirror',
    orders:    'orderslogs',
    inventory: 'Inventory Projection Mirror',
  }
};

// ─────────────────────────────────────────────
// UTILITY
// ─────────────────────────────────────────────

function sheetsUrl(tab) {
  const base = 'https://sheets.googleapis.com/v4/spreadsheets';
  const enc  = encodeURIComponent(tab);
  return `${base}/${SHEETS_CONFIG.sheetId}/values/${enc}?key=${SHEETS_CONFIG.apiKey}`;
}

async function fetchTab(tab) {
  const res = await fetch(sheetsUrl(tab));
  if (!res.ok) throw new Error(`Failed to fetch tab "${tab}": ${res.status}`);
  const json = await res.json();
  const [headers, ...rows] = json.values || [];
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h.trim()] = (row[i] || '').trim(); });
    return obj;
  });
}

function $$(id) { return document.getElementById(id); }

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
// ACCOUNT MIRROR → window.CRM_ACCOUNTS
// Headers: Trade Name, License #, UBI, Address, City, County, State,
//          Zip Code, Phone, License Type, Status, Last Updated, Flags
// ─────────────────────────────────────────────

function parseAccounts(rows) {
  return rows
    .filter(r => r['Trade Name'])
    .map((r, i) => ({
      id:          r['License #'] || `acct-${i}`,
      name:        r['Trade Name'],
      license:     r['License #'],
      ubi:         r['UBI'],
      address:     r['Address'],
      city:        r['City'],
      county:      r['County'],
      state:       r['State'],
      zip:         r['Zip Code'],
      phone:       r['Phone'],
      licenseType: r['License Type'],
      status:      r['Status'],
      lastUpdated: r['Last Updated'],
      flags:       r['Flags'],
      // Defaults for CRM display fields (will be enriched by orderslogs)
      revenue:     0,
      orders:      0,
      trend:       0,
      contacts:    [],
      families:    [],
      logs:        [],
      tasks:       [],
    }));
}

// ─────────────────────────────────────────────
// ORDERSLOGS → window.CRM_OPPORTUNITIES
// Headers: Order #, Product, Product Line, Strain, Units, Line Total,
//          Sample, Submitted Date, Client, License #, Status,
//          Estimated delivery date, Discount, Sample, Submitted By,
//          Assigned To, ...manifest/invoice fields
// ─────────────────────────────────────────────

function parseOrders(rows) {
  const opps = [];

  rows.filter(r => r['Order #'] || r['Client']).forEach((r, i) => {
    const amt = parseFloat((r['Line Total'] || '0').replace(/[$,]/g, '')) || 0;

    // Determine product family from Product Line or Product
    const prodLine = r['Product Line'] || r['Product'] || '';
    let family = 'Other';
    if (/vape|cart/i.test(prodLine))          family = 'Vape';
    else if (/infused/i.test(prodLine))        family = 'Infused Preroll';
    else if (/preroll|joint|1g|2pk/i.test(prodLine)) family = 'Preroll';
    else if (/micro|14g/i.test(prodLine))      family = 'Micro Bud';
    else if (/flower|3\.5|28g/i.test(prodLine)) family = 'Flower';
    else if (/concentrate|extract/i.test(prodLine)) family = 'Concentrate';

    // Stage normalization
    const rawStatus = r['Status'] || '';
    let stage = rawStatus;
    if (/closed.won|delivered/i.test(rawStatus))  stage = 'Closed Won';
    else if (/purchase.order|po/i.test(rawStatus)) stage = 'Purchase Order';
    else if (/sublot/i.test(rawStatus))            stage = 'Sublotted';
    else if (/preorder/i.test(rawStatus))          stage = 'Preorder';
    else if (/submit/i.test(rawStatus))            stage = 'Submitted';

    opps.push({
      id:           r['Order #'] || `opp-${i}`,
      orderNum:     r['Order #'],
      product:      r['Product'],
      productLine:  r['Product Line'],
      strain:       r['Strain'],
      units:        parseInt(r['Units']) || 0,
      amount:       amt,
      sample:       r['Sample'] === 'TRUE' || r['Sample'] === 'true' || r['Sample'] === '1',
      submittedDate: r['Submitted Date'],
      client:       r['Client'],
      license:      r['License #'],
      status:       stage,
      deliveryDate: r['Estimated delivery date'],
      discount:     r['Discount'],
      submittedBy:  r['Submitted By'],
      assignedTo:   r['Assigned To'],
      manifestRef:  r['Manifest Reference #'],
      manifestDate: r['Manifested Date'],
      invoiced:     r['Invoiced'],
      invoicedBy:   r['Invoiced By'],
      releaseRef:   r['Release Transaction #'],
      transferDate: r['Transfer Date'],
      packageSize:  r['Package Size'],
      barcode:      r['Barcode'],
      family,
    });
  });

  return opps;
}

// ─────────────────────────────────────────────
// INVENTORY PROJECTION MIRROR → window.INV_DATA override
// Headers: Week of Date, Grow Room, Totes, Plant Count, Total Days, Strain,
//          Prj. Flower Delivery Date, Prj. Joint Delivery Date,
//          Per Tote, Actual Per Tote, Proj. 3.5g, Actual 3.5g,
//          Target 7g, Target 28g, Wholesale Tier 1 Flower,
//          Proj. Micros 14g, Micros 14g, Wholesale Micros,
//          Projected Trim, % Trim, % 1g, % 2pk, % 10pks,
//          % 1g Infused, % 5pk Infused,
//          1g, 2pk, 10pks, Infused, Batch Count,
//          Total Proj. 1g, Total Proj. 0.5g, Available Proj. 0.5g,
//          0.5g Made, Fresh Frozen, QA, Cured
// ─────────────────────────────────────────────

function parseInventory(rows) {
  return rows
    .filter(r => r['Strain'] && r['Week of Date'])
    .map(r => {
      const weekRaw = r['Week of Date'];
      // Normalize date to YYYY-MM-DD
      let w = weekRaw;
      try {
        const d = new Date(weekRaw);
        if (!isNaN(d)) {
          w = d.toISOString().slice(0, 10);
        }
      } catch(e) {}

      const n = (key) => {
        const v = r[key] || '0';
        return parseInt(v.replace(/,/g, '')) || 0;
      };

      return {
        w,
        strain:  r['Strain'],
        room:    parseInt(r['Grow Room']) || 0,
        totes:   n('Totes'),
        live:    false,
        fd:      r['Prj. Flower Delivery Date'] || '',
        jd:      r['Prj. Joint Delivery Date'] || '',
        p35:     n('Proj. 3.5g'),
        a35:     n('Actual 3.5g'),
        pm:      n('Proj. Micros 14g'),
        j1:      n('1g'),
        j2:      n('2pk'),
        j10:     n('10pks'),
        inf:     n('Infused'),
        // Extra fields available if the CRM uses them later
        perTote:      n('Per Tote'),
        actualPerTote: n('Actual Per Tote'),
        target7g:     n('Target 7g'),
        target28g:    n('Target 28g'),
        trim:         n('Projected Trim'),
        freshFrozen:  n('Fresh Frozen'),
        qa:           n('QA'),
        cured:        n('Cured'),
      };
    });
}

// ─────────────────────────────────────────────
// ENRICH ACCOUNTS with order data
// Calculates revenue, order count, families, trend per account
// ─────────────────────────────────────────────

function enrichAccounts(accounts, orders) {
  const byLicense = {};
  accounts.forEach(a => { byLicense[a.license] = a; });

  orders.forEach(o => {
    const acct = byLicense[o.license];
    if (!acct) return;
    acct.orders++;
    if (/closed.won/i.test(o.status)) acct.revenue += o.amount;
    if (o.family && !acct.families.includes(o.family)) acct.families.push(o.family);
  });

  // Sort by revenue, assign rank
  accounts.sort((a, b) => b.revenue - a.revenue);
  accounts.forEach((a, i) => { a.rank = i + 1; });

  return accounts;
}

// ─────────────────────────────────────────────
// INJECT INTO CRM
// Overrides window globals that the CRM JS reads
// ─────────────────────────────────────────────

function injectIntoApp(accounts, orders, inventory) {
  // 1. Override INV_DATA (inventory page uses this global directly)
  window.INV_DATA = inventory;

  // 2. Store accounts and orders as globals for the CRM to consume
  window.SHEET_ACCOUNTS = accounts;
  window.SHEET_ORDERS   = orders;

  // 3. Re-render whichever page is currently active
  try { if (typeof renderAccountsList === 'function')  renderAccountsList(); }  catch(e) {}
  try { if (typeof renderOpportunities === 'function') renderOpportunities(); } catch(e) {}
  try { if (typeof renderInventory === 'function')     renderInventory(); }     catch(e) {}
  try { if (typeof renderActive === 'function')        renderActive(); }        catch(e) {}

  showBanner(`✓ Sheets synced — ${accounts.length} accounts · ${orders.length} orders · ${inventory.length} inventory rows`);
}

// ─────────────────────────────────────────────
// PATCH CRM FUNCTIONS
// These wrap existing CRM render functions to inject live data
// before the existing logic runs.
// ─────────────────────────────────────────────

function patchCRM() {
  // Patch renderAccountsList to use SHEET_ACCOUNTS if available
  const _origAcctList = window.renderAccountsList;
  if (typeof _origAcctList === 'function') {
    window.renderAccountsList = function() {
      if (window.SHEET_ACCOUNTS && window.SHEET_ACCOUNTS.length) {
        window.ACCOUNTS = window.SHEET_ACCOUNTS;
      }
      return _origAcctList.apply(this, arguments);
    };
  }

  // Patch renderOpportunities to use SHEET_ORDERS if available
  const _origOpps = window.renderOpportunities;
  if (typeof _origOpps === 'function') {
    window.renderOpportunities = function() {
      if (window.SHEET_ORDERS && window.SHEET_ORDERS.length) {
        window.OPPORTUNITIES = window.SHEET_ORDERS;
      }
      return _origOpps.apply(this, arguments);
    };
  }

  // INV_DATA is already overridden at the global level before renderInventory runs
}

// ─────────────────────────────────────────────
// MAIN LOADER
// ─────────────────────────────────────────────

async function loadSheetsData() {
  showBanner('⟳ Loading live data from Google Sheets…', '#2563eb');

  try {
    const [acctRows, orderRows, invRows] = await Promise.all([
      fetchTab(SHEETS_CONFIG.tabs.accounts),
      fetchTab(SHEETS_CONFIG.tabs.orders),
      fetchTab(SHEETS_CONFIG.tabs.inventory),
    ]);

    const accounts  = parseAccounts(acctRows);
    const orders    = parseOrders(orderRows);
    const inventory = parseInventory(invRows);

    enrichAccounts(accounts, orders);
    patchCRM();
    injectIntoApp(accounts, orders, inventory);

    console.log('[Sheets Loader] Done.', {
      accounts:  accounts.length,
      orders:    orders.length,
      inventory: inventory.length,
    });

  } catch (err) {
    console.error('[Sheets Loader] Error:', err);
    showBanner('⚠ Could not load Sheets data — using cached data', '#b45309');
  }
}

// Run after DOM + existing CRM scripts are ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', loadSheetsData);
} else {
  // Small delay so existing CRM scripts finish initializing first
  setTimeout(loadSheetsData, 500);
}

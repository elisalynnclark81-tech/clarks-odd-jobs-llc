// Clark's Odd Jobs — Google Apps Script Backend
// ─────────────────────────────────────────────
// SETUP:
// 1. Open your Google Sheet
// 2. Extensions → Apps Script → paste this entire file
// 3. Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the Web App URL into your dashboard Settings → Google Sheets Sync

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'sync') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (data.clients  && data.clients.length)  syncClients(ss, data.clients);
      if (data.jobs     && data.jobs.length)     syncJobs(ss, data.jobs);
      if (data.invoices && data.invoices.length) syncInvoices(ss, data.invoices);
    }

    return buildResponse({ status: 'ok' });
  } catch (err) {
    return buildResponse({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  // Handles ping test and CORS preflight
  return buildResponse({ status: 'pong', message: "Clark's Odd Jobs connected." });
}

// ── CORS-safe response builder ──
function buildResponse(obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── CLIENTS ──
function syncClients(ss, clients) {
  const sh = getOrCreateSheet(ss, 'Clients',
    ['ID','Name','First','Last','Address','City','ZIP','Phone','Email','Rate','Services','Notes','Start Date','Visits/Mo','Active','Last Updated']);

  const existingIds = getColumnValues(sh, 1);

  clients.forEach(c => {
    const row = [
      c.id, c.name, c.first||'', c.last||'',
      c.address||'', c.city||'', c.zip||'',
      c.phone||'', c.email||'', c.rate||'',
      (c.services||[]).join(', '), c.notes||'',
      c.start||'', c.mows||3, c.active ? 'Yes' : 'No',
      new Date().toLocaleString()
    ];

    const existingRow = existingIds.indexOf(c.id);
    if (existingRow < 0) {
      sh.appendRow(row);
    } else {
      // Update all columns on existing record
      sh.getRange(existingRow + 2, 1, 1, row.length).setValues([row]);
    }
  });
}

// ── JOBS ──
function syncJobs(ss, jobs) {
  const sh = getOrCreateSheet(ss, 'Jobs',
    ['ID','Client','Address','Date','Tasks','Completed','Amount','Notes','Invoiced','Last Updated']);

  const existingIds = getColumnValues(sh, 1);

  jobs.forEach(j => {
    const row = [
      j.id, j.clientName||'', j.address||'', j.date||'',
      (j.tasks||[]).join(', '), j.completed ? 'Yes' : 'No',
      j.amount||0, j.notes||'', j.invoiced ? 'Yes' : 'No',
      new Date().toLocaleString()
    ];

    const existingRow = existingIds.indexOf(j.id);
    if (existingRow < 0) {
      sh.appendRow(row);
    } else {
      sh.getRange(existingRow + 2, 1, 1, row.length).setValues([row]);
    }
  });
}

// ── INVOICES ──
function syncInvoices(ss, invoices) {
  const sh = getOrCreateSheet(ss, 'Invoices',
    ['ID','Client','Date','Total','Paid','Job Count','Last Updated']);

  const existingIds = getColumnValues(sh, 1);

  invoices.forEach(inv => {
    const row = [
      inv.id, inv.clientName||'', inv.date||'',
      inv.total||0, inv.paid ? 'Yes' : 'No',
      (inv.jobIds||[]).length,
      new Date().toLocaleString()
    ];

    const existingRow = existingIds.indexOf(inv.id);
    if (existingRow < 0) {
      sh.appendRow(row);
    } else {
      sh.getRange(existingRow + 2, 1, 1, row.length).setValues([row]);
    }
  });
}

// ── HELPERS ──
function getOrCreateSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    // Style the header row
    const hdr = sh.getRange(1, 1, 1, headers.length);
    hdr.setBackground('#E8620A');
    hdr.setFontColor('#FFFFFF');
    hdr.setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getColumnValues(sh, col) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, col, lastRow - 1, 1).getValues().flat();
}

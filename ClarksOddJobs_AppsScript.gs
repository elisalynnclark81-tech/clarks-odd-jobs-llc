// Clark's Odd Jobs — Google Apps Script Backend
// Paste into Extensions > Apps Script in your Google Sheet, then Deploy as Web App

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (data.action === 'sync') {
      syncSheet(ss, 'Clients', ['ID','Name','Address','City','ZIP','Phone','Email','Rate','Services','Notes','Start','Active'],
        data.clients.map(c=>[c.id,c.name,c.address,c.city,c.zip,c.phone,c.email,c.rate,(c.services||[]).join(', '),c.notes,c.start,c.active]));
      syncSheet(ss, 'Jobs', ['ID','Client','Address','Date','Tasks','Complete','Amount','Notes','Invoiced'],
        data.jobs.map(j=>[j.id,j.clientName,j.address,j.date,(j.tasks||[]).join(', '),j.completed,j.amount,j.notes,j.invoiced]),
        {6:j=>j.completed, 9:j=>j.invoiced});
      syncSheet(ss, 'Invoices', ['ID','Client','Date','Total','Paid'],
        data.invoices.map(i=>[i.id,i.clientName,i.date,i.total,i.paid]));
    }
    return ok();
  } catch(e) { return ContentService.createTextOutput(JSON.stringify({status:'error',msg:e.toString()})).setMimeType(ContentService.MimeType.JSON); }
}

function syncSheet(ss, name, headers, rows, updateCols) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() < 1) sh.appendRow(headers);
  const existingIds = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat() : [];
  rows.forEach(row => {
    const idx = existingIds.indexOf(row[0]);
    if (idx < 0) sh.appendRow(row);
    else if (updateCols) {
      Object.keys(updateCols).forEach(col => sh.getRange(idx+2,parseInt(col)).setValue(row[parseInt(col)-1]));
    }
  });
}

function doGet(e) { return ok(); }
function ok() { return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON); }
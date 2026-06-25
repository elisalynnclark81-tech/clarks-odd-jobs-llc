// Clark's Odd Jobs - Google Apps Script Backend
// Paste this in Extensions > Apps Script in your Google Sheet

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (data.action === 'sync') {
      // Clients sheet
      let cs = ss.getSheetByName('Clients') || ss.insertSheet('Clients');
      if (cs.getLastRow() < 2) cs.appendRow(['ID','Name','Address','City','ZIP','Phone','Email','Rate','Services','Notes','Start Date','Active']);
      const existingIds = cs.getRange(2, 1, Math.max(cs.getLastRow()-1,1), 1).getValues().flat();
      data.clients.forEach(c => {
        if (!existingIds.includes(c.id)) {
          cs.appendRow([c.id, c.name, c.address, c.city, c.zip, c.phone, c.email, c.rate, (c.services||[]).join(', '), c.notes, c.startDate, c.active]);
        }
      });
      
      // Jobs sheet
      let js = ss.getSheetByName('Jobs') || ss.insertSheet('Jobs');
      if (js.getLastRow() < 2) js.appendRow(['ID','Client','Address','Date','Tasks','Completed','Amount','Notes','Invoiced']);
      const existingJobIds = js.getRange(2, 1, Math.max(js.getLastRow()-1,1), 1).getValues().flat();
      data.jobs.forEach(j => {
        if (!existingJobIds.includes(j.id)) {
          js.appendRow([j.id, j.clientName, j.address, j.date, (j.tasks||[]).join(', '), j.completed, j.amount, j.notes, j.invoiced]);
        } else {
          const row = existingJobIds.indexOf(j.id) + 2;
          js.getRange(row, 6).setValue(j.completed);
          js.getRange(row, 9).setValue(j.invoiced);
        }
      });
      
      // Invoices sheet
      let invS = ss.getSheetByName('Invoices') || ss.insertSheet('Invoices');
      if (invS.getLastRow() < 2) invS.appendRow(['ID','Client','Date','Amount','Paid']);
      const existingInvIds = invS.getRange(2, 1, Math.max(invS.getLastRow()-1,1), 1).getValues().flat();
      data.invoices.forEach(i => {
        if (!existingInvIds.includes(i.id)) {
          invS.appendRow([i.id, i.clientName, i.date, i.amount||i.total, i.paid]);
        } else {
          const row = existingInvIds.indexOf(i.id) + 2;
          invS.getRange(row, 5).setValue(i.paid);
        }
      });
    }
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status:'pong'})).setMimeType(ContentService.MimeType.JSON);
}

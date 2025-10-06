// macro.gs — Menu, Sidebar launcher, Data providers, and helpers
// ---------------------------------------------------------------

/** Adds the custom menu on open */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quote Tool')
    .addItem('Open Quote Form', 'showQuoteForm')
    .addItem('Save Client Quote', 'prepareAndEmailQuoteCopy') // implemented in QuoteExport.gs
    .addToUi();
}

/** If the menu disappears, run this once from the editor */
function installMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Quote Tool')
    .addItem('Open Quote Form', 'showQuoteForm')
    .addItem('Save Client Quote', 'prepareAndEmailQuoteCopy')
    .addToUi();
}

/** Opens the sidebar (loads QuoteForm.html by name) */
function showQuoteForm() {
  const html = HtmlService.createHtmlOutputFromFile('QuoteForm')
    .setTitle('Quote Tool')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Tiny connectivity ping used by the form on load */
function ping() { return 'ok'; }

/* =======================================================
 * DATA PROVIDERS
 * ======================================================= */

/** TT Towbars → returns [{make,model,body,year,type}] using header names */
function getFullDropdownData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('TT Towbars');
  if (!sheet) throw new Error("Sheet 'TT Towbars' not found.");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const norm = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const hmap = {};
  headers.forEach((h, i) => { const k = norm(h); if (k) hmap[k] = i; });

  const find = (cands) => {
    for (const c of cands) { const k = norm(c); if (k in hmap) return hmap[k]; }
    for (const c of cands) {
      const k = norm(c);
      const hit = Object.keys(hmap).find(x => x.includes(k));
      if (hit) return hmap[hit];
    }
    return -1;
  };

  const idxType  = find(['TOWBAR TYPE', 'Towbar Type', 'Type']);
  const idxMake  = find(['MAKE', 'Make', 'Manufacturer', 'Brand']);
  const idxModel = find(['MODEL', 'Model', 'Vehicle Model']);
  const idxBody  = find(['BODY TYPE', 'Body Type', 'Body', 'Body Style']);
  const idxYear  = find(['YEAR', 'Year', 'Year Range', 'Fitment Year']);

  const missing = [['TOWBAR TYPE', idxType], ['MAKE', idxMake], ['MODEL', idxModel], ['BODY TYPE', idxBody], ['YEAR', idxYear]]
    .filter(([, i]) => i < 0).map(([n]) => n);
  if (missing.length) throw new Error("TT Towbars: missing column(s): " + missing.join(', '));

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const s = v => v == null ? '' : String(v).trim();

  return data.map(row => ({
    type:  s(row[idxType]),
    make:  s(row[idxMake]),
    model: s(row[idxModel]),
    body:  s(row[idxBody]),
    year:  s(row[idxYear])
  }));
}

/** TT VSK → returns [{make,model,body,year,type,part}] using header names */
function getVSKDropdownData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('TT VSK');
  if (!sheet) throw new Error("Sheet 'TT VSK' not found.");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const norm = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const hmap = {};
  headers.forEach((h, i) => { const k = norm(h); if (k) hmap[k] = i; });

  const find = (cands) => {
    for (const c of cands) { const k = norm(c); if (k in hmap) return hmap[k]; }
    for (const c of cands) {
      const k = norm(c);
      const hit = Object.keys(hmap).find(x => x.includes(k));
      if (hit) return hmap[hit];
    }
    return -1;
  };

  const idxType  = find(['VSK TYPE', 'Type', 'Kit Type', 'Towbar Type']);
  const idxPart  = find(['PART No.', 'Part No', 'Part Number', 'SKU', 'Code']);
  const idxMake  = find(['MAKE', 'Make', 'Manufacturer', 'Brand']);
  const idxModel = find(['MODEL', 'Model', 'Vehicle Model']);
  const idxBody  = find(['BODY TYPE', 'Body Type', 'Body', 'Body Style']);
  const idxYear  = find(['YEAR', 'Year', 'Year Range', 'Fitment Year', 'Date']);

  const missing = [['VSK TYPE', idxType], ['PART No.', idxPart], ['MAKE', idxMake], ['MODEL', idxModel], ['BODY TYPE', idxBody], ['YEAR', idxYear]]
    .filter(([, i]) => i < 0).map(([n]) => n);
  if (missing.length) throw new Error("TT VSK: missing column(s): " + missing.join(', '));

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const s = v => v == null ? '' : String(v).trim();
  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  return data.map(row => {
    const y = row[idxYear];
    const yearStr = (y instanceof Date) ? fmt(y) : s(y);
    return {
      type:  s(row[idxType]),
      part:  s(row[idxPart]),
      make:  s(row[idxMake]),
      model: s(row[idxModel]),
      body:  s(row[idxBody]),
      year:  yearStr
    };
  });
}

/* =======================================================
 * VSK REQUIRED LOOKUP (exact match on Make/MODEL/BODY TYPE/YEAR)
 * Returns the "VSK Required?" cell content.
 * =======================================================
 */
function getVSKRequired(payload) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('VSK Required');
  if (!sh) return 'Sheet "VSK Required" not found';

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return 'No data';

  const norm = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const headers = values[0];
  const hmap = {};
  headers.forEach((h, i) => { hmap[norm(h)] = i; });

  const idxMake   = hmap[norm('Make')];
  const idxModel  = hmap[norm('MODEL')];
  const idxBody   = hmap[norm('BODY TYPE')];
  const idxYear   = hmap[norm('YEAR')];
  const idxVSKReq = hmap[norm('VSK Required?')];

  const missing = [];
  if (idxMake==null)   missing.push('Make');
  if (idxModel==null)  missing.push('MODEL');
  if (idxBody==null)   missing.push('BODY TYPE');
  if (idxYear==null)   missing.push('YEAR');
  if (idxVSKReq==null) missing.push('VSK Required?');
  if (missing.length) return 'VSK Required: missing column(s): ' + missing.join(', ');

  const wantMake  = String(payload.make  || '').trim();
  const wantModel = String(payload.model || '').trim();
  const wantBody  = String(payload.body  || '').trim();
  const wantYear  = String(payload.year  || '').trim();

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const mk = String(row[idxMake]  || '').trim();
    const md = String(row[idxModel] || '').trim();
    const bd = String(row[idxBody]  || '').trim();
    const yr = String(row[idxYear]  || '').trim();
    if (mk === wantMake && md === wantModel && bd === wantBody && yr === wantYear) {
      return String(row[idxVSKReq] || '');
    }
  }
  return '';
}

/* =======================================================
 * OPTIONAL UTILITIES
 * ======================================================= */

/** Optional: create an installable onOpen trigger to auto-open the sidebar */
function ensureAutoOpenSidebar() {
  const ssId = SpreadsheetApp.getActive().getId();
  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'showQuoteForm' && t.getEventType() === ScriptApp.EventType.ON_OPEN);
  if (!exists) ScriptApp.newTrigger('showQuoteForm').forSpreadsheet(ssId).onOpen().create();
  SpreadsheetApp.getActive().toast('Auto-open enabled. Reopen this spreadsheet.');
}

/** Optional: Auto-reset 'Add VSK?' cell if Q1–Q4 cleared on 'Quote' sheet */
function onEdit(e) {
  if (!e || !e.source) return;
  const sheet = e.source.getActiveSheet();
  if (!sheet || sheet.getName() !== 'Quote') return;

  const editedA1 = e.range.getA1Notation();
  const qCells = ['B2', 'C2', 'D2', 'E2']; // Q1–Q4 inputs
  const addVskCell = 'F2';                 // 'Add VSK?' cell
  if (qCells.includes(editedA1)) {
    const allEmpty = qCells.every(a1 => {
      const v = sheet.getRange(a1).getValue();
      return v === '' || v === null;
    });
    if (allEmpty) sheet.getRange(addVskCell).setValue('');
  }
}

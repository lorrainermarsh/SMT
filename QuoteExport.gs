/**
 * QuoteExport.gs — Prepare Email Quote
 *
 * File naming: "Name - Make - Model" (from sidebar payload when available; fallback to Quote sheet).
 * Appends: Accessories (blank row before), then Terms & Conditions (blank row before).
 * Keeps only: Quote, Towbar Warranties, Options and Information.
 * Logs to QuoteTracker with a "View quote" hyperlink.
 */

function prepareAndEmailQuoteCopy(payloadOrNameArg, _forceRun) {
  const payload = (payloadOrNameArg && typeof payloadOrNameArg === 'object') ? payloadOrNameArg : null;

  const ss      = SpreadsheetApp.getActive();
  const srcFile = DriveApp.getFileById(ss.getId());
  const parents = srcFile.getParents();

  const quoteSrc = ss.getSheetByName('Quote');
  const tbSrc    = ss.getSheetByName('TT Towbars');
  const vskSrc   = ss.getSheetByName('TT VSK');
  const tncSrc   = ss.getSheetByName('Terms and Conditions');
  const accSrc   = ss.getSheetByName('Accessories');
  if (!quoteSrc || !tbSrc || !vskSrc) throw new Error("Required sheets missing (Quote / TT Towbars / TT VSK).");

  // Build filename: "Name - Make - Model"
  const copyName = buildFileNameFromInputs_(quoteSrc, payload);

  const copyFile = parents.hasNext()
    ? srcFile.makeCopy(copyName, parents.next())
    : srcFile.makeCopy(copyName);

  const copy  = SpreadsheetApp.openById(copyFile.getId());
  const quote = copy.getSheetByName('Quote');
  const tb    = copy.getSheetByName('TT Towbars');
  const vsk   = copy.getSheetByName('TT VSK');
  const tnc   = copy.getSheetByName('Terms and Conditions');
  const acc   = copy.getSheetByName('Accessories');
  if (!quote || !tb || !vsk) throw new Error("Copy missing one of: Quote / TT Towbars / TT VSK.");

  // Strip detail tables from Quote
  stripDetailsFromQuote_Fast_(quote, tb, vsk);

  // Append Accessories then Terms & Conditions (1 blank row before each)
  if (acc) appendSheetToQuote_Fast_(quote, acc, 1);
  if (tnc) appendSheetToQuote_Fast_(quote, tnc, 1);

  // Keep only requested sheets
  keepOnlySheets_(copy, ['Quote', 'Towbar Warranties', 'Options and Information']);
  copy.setActiveSheet(quote);

  // Append to QuoteTracker (Name, Email, Car, "View quote" hyperlink)
  const parsed = readNameEmailCar_(quote, payload);
  appendToQuoteTracker_(srcFile, {
    Name: parsed.name,
    Email: parsed.email,
    Car: parsed.car,
    'Quote Link': copyFile.getUrl()
  });

  showLinkDialog_(copyFile);
  return copyFile.getUrl();
}

/* ===================== Naming helpers ===================== */

/** Build "Name - Make - Model" from payload; fallback to Quote sheet (A5/A6 + header A8/A7). */
function buildFileNameFromInputs_(quoteSheet, payload) {
  let name = '', make = '', model = '';

  if (payload) {
    name  = String(payload.clientName || '').trim();
    make  = String(payload.make || '').trim();
    model = String(payload.model || '').trim();
  }

  if (!name) {
    const a5 = String(quoteSheet.getRange(5, 1).getDisplayValue() || '').trim();
    name = a5.replace(/^Name:\s*/i, '').trim();
  }
  if ((!make || !model)) {
    const header = String(quoteSheet.getRange('A8').getDisplayValue() || '').trim()
                || String(quoteSheet.getRange('A7').getDisplayValue() || '').trim();
    const parts = header.split(',').map(s => String(s || '').trim());
    if (parts.length >= 2) {
      if (!make)  make  = parts[0] || '';
      if (!model) model = parts[1] || '';
    }
  }

  const tokens = [name, make, model].filter(Boolean);
  return tokens.length ? tokens.join(' - ') : 'Quote';
}

/** Read Name, Email, Car for QuoteTracker logging. */
function readNameEmailCar_(quoteSheet, payload) {
  let name  = (payload && String(payload.clientName || '').trim()) || '';
  let email = (payload && String(payload.clientEmail || '').trim()) || '';

  if (!name)  name  = String(quoteSheet.getRange(5, 1).getDisplayValue() || '').replace(/^Name:\s*/i, '').trim();
  if (!email) email = String(quoteSheet.getRange(6, 1).getDisplayValue() || '').replace(/^Email:\s*/i, '').trim();

  let car = '';
  if (payload && payload.make) {
    car = [payload.make, payload.model, payload.body, payload.year]
      .map(v => String(v || '').trim()).filter(Boolean).join(', ');
  } else {
    car = String(quoteSheet.getRange('A8').getDisplayValue() || '').trim()
       || String(quoteSheet.getRange('A7').getDisplayValue() || '').trim()
       || '';
  }
  return { name, email, car };
}

/* ===================== Append / strip helpers ===================== */

function stripDetailsFromQuote_Fast_(quote, tbSheet, vskSheet) {
  const lr = quote.getLastRow();
  if (lr < 1) return;

  const HEADING = 'DETAILED INFO - INTERNAL VIEW ONLY';
  const headingCell = quote.getRange(1, 1, lr, 1)
    .createTextFinder(HEADING)
    .matchCase(true)
    .matchEntireCell(true)
    .findNext();
  if (!headingCell) return;

  const headingRow = headingCell.getRow();
  quote.deleteRows(headingRow, lr - headingRow + 1);
}

function appendSheetToQuote_Fast_(quote, srcSheet, spacerRows) {
  const lastUsed = findLastUsedRowAcrossSheet_(quote);
  const spacer   = Math.max(0, Number(spacerRows || 0));
  const destTop  = Math.max(1, lastUsed + 1 + spacer);

  const src = srcSheet.getDataRange();
  const h   = src.getNumRows();
  const w   = src.getNumColumns();

  if (quote.getMaxRows() < destTop + h - 1) {
    quote.insertRowsAfter(quote.getMaxRows(), destTop + h - 1 - quote.getMaxRows());
  }
  if (quote.getMaxColumns() < w) {
    quote.insertColumnsAfter(quote.getMaxColumns(), w - quote.getMaxColumns());
  }
  src.copyTo(quote.getRange(destTop, 1), { contentsOnly: false });

  const srcTop  = src.getRow(), srcLeft = src.getColumn();
  const merges  = src.getMergedRanges();
  const needMerges = merges.length &&
                     !quote.getRange(destTop, 1, merges[0].getNumRows(), merges[0].getNumColumns()).isPartOfMerge();
  if (needMerges) {
    merges.forEach(m => {
      const dr = destTop + (m.getRow()   - srcTop);
      const dc = 1       + (m.getColumn()- srcLeft);
      try { quote.getRange(dr,dc,m.getNumRows(),m.getNumColumns()).merge(); } catch (_) {}
    });
  }
}

/* ===================== QuoteTracker helpers ===================== */

function getQuoteTracker_(srcFile) {
  let parent = null;
  const parents = srcFile.getParents();
  if (parents.hasNext()) parent = parents.next();

  const candidates = ['QuoteTracker', 'Quote Tracker'];

  if (parent) {
    const it = parent.getFilesByName(candidates[0]);
    if (it.hasNext()) return ensureTrackerSheet_(SpreadsheetApp.openById(it.next().getId()));
    const it2 = parent.getFilesByName(candidates[1]);
    if (it2.hasNext()) return ensureTrackerSheet_(SpreadsheetApp.openById(it2.next().getId()));
  }
  for (const name of candidates) {
    const files = DriveApp.getFilesByName(name);
    if (files.hasNext()) return ensureTrackerSheet_(SpreadsheetApp.openById(files.next().getId()));
  }

  const newSS = SpreadsheetApp.create('QuoteTracker');
  if (parent) DriveApp.getFileById(newSS.getId()).moveTo(parent);
  return ensureTrackerSheet_(newSS);
}

function ensureTrackerSheet_(ss) {
  const sheetName = 'QuoteTracker';
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const required = ['Name', 'Email', 'Car', 'Quote Link'];

  const lc = Math.max(1, sh.getLastColumn());
  let headers = sh.getRange(1,1,1,lc).getValues()[0];
  if (headers.every(v => v === '' || v == null)) headers = [];
  const existing = headers.map(h => String(h||'').trim());

  for (const col of required) if (!existing.includes(col)) existing.push(col);
  sh.getRange(1,1,1,existing.length).setValues([existing]);

  return { ss, sh };
}

function appendToQuoteTracker_(srcFile, record) {
  const { ss, sh } = getQuoteTracker_(srcFile);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
  const map = {}; headers.forEach((h,i)=> map[h]=i+1);

  const nextRow = sh.getLastRow() + 1;
  if (map['Name'])  sh.getRange(nextRow, map['Name']).setValue(record['Name'] || '');
  if (map['Email']) sh.getRange(nextRow, map['Email']).setValue(record['Email'] || '');
  if (map['Car'])   sh.getRange(nextRow, map['Car']).setValue(record['Car'] || '');

  if (map['Quote Link']) {
    const url = String(record['Quote Link'] || '').trim();
    const formula = url ? `=HYPERLINK("${url}","View quote")` : '';
    sh.getRange(nextRow, map['Quote Link']).setValue(formula);
  }
}

/* ===================== Misc ===================== */

function keepOnlySheets_(spreadsheet, keepNames) {
  const keep = new Set(keepNames);
  spreadsheet.getSheets().forEach(sh => { if (!keep.has(sh.getName())) spreadsheet.deleteSheet(sh); });
}
function findLastUsedRowAcrossSheet_(sheet) {
  const lr = sheet.getLastRow(), lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return 0;
  const vals = sheet.getRange(1, 1, lr, lc).getValues();
  for (let r = lr - 1; r >= 0; r--) {
    for (let c = 0; c < lc; c++) {
      if (vals[r][c] !== '' && vals[r][c] !== null) return r + 1;
    }
  }
  return 0;
}
function showLinkDialog_(file) {
  const html = HtmlService
    .createHtmlOutput(
      '<div style="font:14px Arial;padding:12px 8px;">' +
        '<div>File created:</div>' +
        '<div style="margin-top:8px;"><a target="_blank" href="' + file.getUrl() + '">' +
          file.getName() + '</a></div>' +
      '</div>'
    ).setWidth(380).setHeight(130);
  SpreadsheetApp.getUi().showModalDialog(html, 'Prepare Email Quote');
}

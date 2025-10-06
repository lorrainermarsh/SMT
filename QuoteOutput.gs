/** QuoteOutput.gs — unified generator (no VSK-only), rules-driven output */

function form_getInitialData() {
  // Build lists from TT Towbars using mapping headers (Make, Model, Body Type, Vehicle Description)
  const map = SMT_CFG.NAMES.MAP_TOWBARS;
  const mkH  = resolveHeader_(map, SMT_CFG.REQS.MAKE);
  const mdH  = resolveHeader_(map, SMT_CFG.REQS.MODEL);
  const bdH  = resolveHeader_(map, SMT_CFG.REQS.BODY);
  const vdH  = resolveHeader_(map, SMT_CFG.REQS.VEH_DESC);

  const s = sh_(SMT_CFG.SHEETS.TOWBARS);
  const vals = s.getRange(2, 1, s.getLastRow()-1, s.getLastColumn()).getValues();
  const uniq = (i) => [...new Set(vals.map(r => r[i]).filter(Boolean))].sort();
  const makes  = uniq(headerIndex_(s.getName(), mkH));
  const models = uniq(headerIndex_(s.getName(), mdH));
  const bodies = uniq(headerIndex_(s.getName(), bdH));
  const vehDescs = uniq(headerIndex_(s.getName(), vdH));

  // Towbar types from TT Towbars (distinct)
  const typeIdx = headerIndex_(s.getName(), 'Towbar Type') // fallback if present
  let towbarTypes = [];
  try {
    towbarTypes = uniq(typeIdx);
  } catch (e) { towbarTypes = []; }

  // VSK Towbar Type from TT VSK if present
  let vskTowbarTypes = [];
  try {
    const vs = sh_(SMT_CFG.SHEETS.VSK);
    const vTypeIdx = headerIndex_(vs.getName(), 'VSK TYPE');
    const vVals = vs.getRange(2, 1, vs.getLastRow()-1, vs.getLastColumn()).getValues();
    vskTowbarTypes = [...new Set(vVals.map(r=>r[vTypeIdx]).filter(Boolean))].sort();
  } catch (e) { vskTowbarTypes = []; }

  return {
    makes, models, bodies,
    vehicleDescriptions: vehDescs,
    towbarTypes,
    vskTowbarTypes
  };
}

/* Lookups requested by the form */

/** Towbar Notes under “VSK required” row — from TT Towbar Notes (by Make/Model/Body/Vehicle Description) */
function form_lookupTowbarNotes(payload){
  const match = {
    [SMT_CFG.REQS.MAKE]: payload.make,
    [SMT_CFG.REQS.MODEL]: payload.model,
    [SMT_CFG.REQS.BODY]: payload.body,
    [SMT_CFG.REQS.VEH_DESC]: payload.vehDescTow || payload.vehDescVsk
  };
  return lookupByRequirements_(SMT_CFG.SHEETS.TOWBAR_NOTES, SMT_CFG.NAMES.MAP_TOWBARS, match, SMT_CFG.REQS.NOTES);
}

/** VSK Notes under “VSK Towbar Type” — from TT VSK Notes */
function form_lookupVskNotes(payload){
  const match = {
    [SMT_CFG.REQS.MAKE]: payload.make,
    [SMT_CFG.REQS.MODEL]: payload.model,
    [SMT_CFG.REQS.BODY]: payload.body,
    [SMT_CFG.REQS.VEH_DESC]: payload.vehDescVsk || payload.vehDescTow
  };
  return lookupByRequirements_(SMT_CFG.SHEETS.VSK_NOTES, SMT_CFG.NAMES.MAP_VSK, match, SMT_CFG.REQS.NOTES);
}

/** Extension kit required — from TT VSK */
function form_lookupVskExtensionRequired(payload){
  const match = {
    [SMT_CFG.REQS.MAKE]: payload.make,
    [SMT_CFG.REQS.MODEL]: payload.model,
    [SMT_CFG.REQS.BODY]: payload.body,
    [SMT_CFG.REQS.VEH_DESC]: payload.vehDescVsk || payload.vehDescTow
  };
  return lookupByRequirements_(SMT_CFG.SHEETS.VSK, SMT_CFG.NAMES.MAP_VSK, match, SMT_CFG.REQS.EXTENSION_REQ) || '';
}

/** Entrypoint from the form */
function processQuoteFromForm(payload){
  // Write useful fields to the Operations sheet if you store state (optional),
  // then call processQuote() which performs the output write.
  return processQuote(payload);
}

/** Main: build Quote + Price Breakdown using Quote_rules */
function processQuote(payload){
  const quoteSh = sh_(SMT_CFG.SHEETS.QUOTE);

  // 1) Hard reset section under INTERNAL_MARKER (and include the marker row itself)
  const markerText = SMT_CFG.INTERNAL_MARKER;
  const data = quoteSh.getDataRange().getValues();
  let markerRow = data.findIndex(r => String(r[0]).trim() === markerText);
  if (markerRow === -1) {
    // if not found, append at end as a safety
    markerRow = data.length;
    quoteSh.getRange(markerRow+1,1).setValue(markerText);
  }
  // clear from markerRow (1-based indexing for Sheets)
  const startClear = markerRow + 1; // 0-based -> to next line
  if (quoteSh.getLastRow() > startClear) {
    quoteSh.getRange(startClear+1, 1, quoteSh.getLastRow() - startClear, quoteSh.getLastColumn()).clearContent().clearFormat();
  }

  // 2) Load Quote_rules
  const rules = readNamedTable_(SMT_CFG.NAMES.QUOTE_RULES);
  if (!rules.length) throw new Error('Quote_rules is empty or missing headers.');

  // expected headers (normalised):
  // output table | label | calc | static value | include when | order
  // Any missing fields are treated gracefully.
  const QUOTE = 'quote', BREAK = 'price breakdown';

  const sections = {
    [QUOTE]: [],
    [BREAK]: []
  };

  // Optionally compute any needed derived values up-front (kept minimal here)
  const derived = computeDerived_(payload);

  // Filter & compute rows
  for (const r of rules) {
    const outTbl = norm_(r['output table'] || QUOTE);
    const label  = r['label'] ?? '';
    const calc   = String(r['calc'] || '').trim();
    const staticVal = r['static value'];
    const includeExpr = String(r['include when'] || '').trim();
    const order = Number(r['order'] ?? 9999);

    // evaluate include
    let include = true;
    if (includeExpr) {
      try { include = !!evaluate_(includeExpr, derived); } catch (e) { include = true; }
    }

    if (!include) continue;

    let value = '';
    if (calc) {
      value = computeCalc_(calc, derived);
    } else if (staticVal !== undefined) {
      value = staticVal;
    }

    const row = [label, value];
    if (outTbl === BREAK) sections[BREAK].push({order, row});
    else sections[QUOTE].push({order, row});
  }

  // 3) Sort and write sections starting at marker row + 1
  const startRow = markerRow + 2; // line below marker
  let cursor = startRow;

  // QUOTE section
  sections[QUOTE].sort((a,b)=>a.order-b.order);
  if (sections[QUOTE].length) {
    quoteSh.getRange(cursor, 1, sections[QUOTE].length, 2).setValues(sections[QUOTE].map(x=>x.row));
    cursor += sections[QUOTE].length + 1; // spacer row
  }

  // PRICE BREAKDOWN
  if (sections[BREAK].length) {
    sections[BREAK].sort((a,b)=>a.order-b.order);
    quoteSh.getRange(cursor, 1, sections[BREAK].length, 2).setValues(sections[BREAK].map(x=>x.row));
    cursor += sections[BREAK].length;
  }

  return 'Quote generated.';
}

/* ------- calculations & overrides ------- */

function computeDerived_(payload) {
  // Pull commonly-needed values based on mapping names.
  // In your current workbook you likely already compute:
  // - Towbar price, VSK price, Fit times, Labour rate, Cable, etc.
  // This is a placeholder wiring to show where overrides happen.
  const pVars = sh_(SMT_CFG.SHEETS.PRICE_VARIABLES);
  // You can read single named cells here if you have them as named ranges,
  // e.g., rng_('Towbar_fitting_cost_per_hour').getValue()

  const make = payload?.make || '';
  const model = payload?.model || '';
  const body = payload?.body || '';
  const veh = payload?.vehDescVsk || payload?.vehDescTow || '';

  // base computations (replace with your real calcs as needed)
  const towbarPrice = 0;
  const vskPrice = 0;
  const wiringPinCost = 0; // will be overridden if non-standard match

  // override using Non_standard_VSK_fitting_costs
  const pinOverride = getNonStdVskPinOverride_(make, model, body, veh);
  const towbarAndVskWiringFitting = (pinOverride != null) ? pinOverride : wiringPinCost;

  return {
    make, model, body, veh,
    towbarPrice, vskPrice,
    towbarAndVskWiringFitting
  };
}

/** Evaluate simple include expressions like "vsk_required == 'Yes'" against a dict */
function evaluate_(expr, ctx) {
  // VERY simple & safe evaluator: supports ctx.<key>, ==, !=, truthy vars
  const safe = expr.replace(/[^\w\s\.\=\!\'\-]/g,'');
  const parts = safe.split(/\s+/);
  if (parts.length === 1) return !!ctx[parts[0]];
  if (parts.length >= 3) {
    const [lhs, op, ...rest] = parts;
    const rhs = rest.join(' ').replace(/^'|'$/g,'');
    const L = String(ctx[lhs] ?? '');
    if (op === '==') return L === rhs;
    if (op === '!=') return L !== rhs;
  }
  return true;
}

/** Minimal calculator to map rule calc keys to derived values */
function computeCalc_(key, d) {
  const k = norm_(key);
  if (k === 'towbar and vsk wiring fitting £' || k === 'towbar & vsk wiring fitting £') return d.towbarAndVskWiringFitting;
  if (k === 'towbar price') return d.towbarPrice;
  if (k === 'vsk price') return d.vskPrice;
  return '';
}

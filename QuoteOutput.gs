/**
* QuoteOutput — Rule-driven version
* - Both "Quote" and "Price breakdown" are driven by Quote_rules:
*   - Row inclusion/exclusion via pickRule_()
*   - Labels via rule.label (fallback to mapPinLabel_())
*   - Cable add-ons via rule.cable (fallback to Pin Cable price variable)
*
* Assumes these helpers already exist in your codebase:
*   - loadQuoteRules_() -> Array<Object>
*   - pickRule_(rules, quoteTypeUI, kind, pinNK, vskYesNo, towbarTypeText, vskTypeNK) -> Object|null
*   - NK(s) -> string  (normalize key)
*   - numval(x) -> number
*   - mapPinLabel_(pinText, requiredBool) -> string
*   - build indices: towPriceIdx, vskIdx, towUniFitIdx, towVskFitIdx, cableIdx
*   - data gatherer: getQuoteInput_() or similar (returns model/body/year, pins, etc.)
*   - any styling helpers you already used to format the target sheet
*
* If any names differ in your project, rename the calls below to match.
*/




/* -----------------------------------------------------------
* Canonical normalization for Quote Type + pickup cable flag
* -----------------------------------------------------------
*/
function normalizeQuoteType_(data) {
 const quoteTypeRaw = String(data.quoteType || 'Towbars + Fitting +/- VSK').trim();
 const quoteTypeUI  = (quoteTypeRaw.toLowerCase() === 'towbar') ? 'Towbars + Fitting +/- VSK'
                   : (quoteTypeRaw.toLowerCase() === 'vsk')    ? 'VSK only'
                   : quoteTypeRaw;
 const pickupYes = (quoteTypeUI !== 'VSK only') && (String(data.pickupCable || '').toLowerCase() === 'yes');
 return { quoteTypeRaw, quoteTypeUI, pickupYes };
}




/* -----------------------------------------------------------
* Generate the main "Quote" table (Table 1)
* -----------------------------------------------------------
* Expected behavior (unchanged except now fully rule-driven):
* - Uses Quote_rules to decide which lines render.
* - Labels come from rules.label (fallback mapPinLabel_()).
* - Cable add-ons come from rules.cable (numeric) else Pin Cable price.
*/
function processQuote() {
 const ss = SpreadsheetApp.getActive();
 const sh = ss.getSheetByName('Quote'); // adjust if your output sheet is differently named


 // --- Gather all input data you already use in your project
 // (Make/model/body/year, pin families/types, price indices, etc.)
 const data = getQuoteInput_(); // <-- your existing helper that reads form/quote input
 const { quoteTypeUI, pickupYes } = normalizeQuoteType_(data);


 // Price/fit/cable indices you already build elsewhere
 const towPriceIdx   = buildTowPriceIndex_();     // key: towbar type text -> price
 const towUniFitIdx  = buildTowUniFitIndex_();    // key: PIN NK -> hours
 const towVskFitIdx  = buildTowVskFitIndex_();    // key: PIN NK -> hours
 const vskIdx        = buildVskPriceIndex_();     // key: [model,body,year,vskType].map(NK).join('|') -> price
 const cableIdx      = buildCableIndex_();        // key: PIN NK -> cable price
 const rules         = loadQuoteRules_();


 // Unpack user selection / computed candidates
 const pins          = data.pinFamilies || [];             // array of strings (e.g. ["7 PIN", "13 PIN", "TWIN"])
 const towbarTypes   = data.towbarTypes || [];             // array of towbar type strings (contain "Detachable"/"Fixed" etc.)
 const vskTypes      = data.vskTypes || [];                // array of vsk family strings (e.g. ["7 PIN", "13 PIN"])
 const vskBase       = !!data.vskAvailable;                // boolean: whether VSK data exists for the vehicle
 const vskModel      = String(data.model || '');
 const vskBody       = String(data.body  || '');
 const vskYear       = String(data.year  || '');


 // Output rows we’ll write to the Quote table
 const out = [];


 // --- Header or formatting reset (reuse what you had before if you style rows 1–9 etc.)
 resetQuoteTableArea_(sh); // optional: your existing function that clears rows below header, styles, etc.


 // --- Business Rule #1: no Twin with Detachable towbars
 const isDetachable = (t) => String(t).toUpperCase().includes('DETACHABLE');


 // --- Build rows (towbar-only + towbar+VSK when applicable)
 for (const t of towbarTypes) {
   for (const p of pins) {
     const pNK = NK(p);
     const towPriceNum = numval(towPriceIdx.get(String(t)) || 0);


     // 1) Towbar Only
     if (!(isDetachable(t) && pNK === 'TWIN')) {
       const rTowOnly = pickRule_(rules, quoteTypeUI, 'Towbar Only', pNK, 'No', t, '');
       if (rTowOnly && String(rTowOnly.enable).toLowerCase() === 'yes') {
         // label
         const label = rTowOnly.label || mapPinLabel_(p, /*required*/ false);
         // cable (Towbar only branch) — prefer rule.cable if numeric > 0
         const cab = (() => {
           if (!pickupYes) return 0;
           const override = Number(rTowOnly.cable || 0);
           if (override > 0) return override;
           const price = cableIdx.get(pNK) ?? 0;
           return (typeof price === 'number' ? price : numval(price)) || 0;
         })();
         // fit time
         const uniFit = towUniFitIdx.get(pNK) || 0;


         out.push([label, towPriceNum, '', uniFit, '', cab, null]);
       }
     }


     // 2) Towbar + VSK
     if (quoteTypeUI !== 'VSK only' && vskBase && vskTypes.length > 0) {
       for (const vt of vskTypes) {
         if (pNK === 'TWIN') continue; // never combine Twin with VSK
         const vtNK = NK(vt);


         const rTowPlus = pickRule_(rules, quoteTypeUI, 'Towbar Plus VSK', pNK, 'Yes', t, vtNK);
         if (!rTowPlus || String(rTowPlus.enable).toLowerCase() !== 'yes') continue;


         // label
         const label = rTowPlus.label || mapPinLabel_(p, /*required*/ true);
         // prices
         const vRaw = vskIdx.get([vskModel, vskBody, vskYear, vt].map(NK).join('|')) ?? '';
         const vNum = numval(vRaw);
         const vskFit = towVskFitIdx.get(pNK) || 0;


         // cable (Towbar+VSK branch)
         const cab = (() => {
           if (!pickupYes) return 0;
           const override = Number(rTowPlus.cable || 0);
           if (override > 0) return override;
           const price = cableIdx.get(pNK) ?? 0;
           return (typeof price === 'number' ? price : numval(price)) || 0;
         })();


         out.push([label, towPriceNum, vNum, '', vskFit, cab, null]);
       }
     }
   }
 }


 // 3) VSK-only table (when quoteType is VSK only)
 if (quoteTypeUI === 'VSK only' && vskBase && vskTypes.length > 0) {
   // Render only rule-enabled VSK rows
   for (const vt of vskTypes) {
     const vtNK = NK(vt);
     const rVsk = pickRule_(rules, quoteTypeUI, 'VSK Only', '', 'Yes', '', vtNK);
     if (!rVsk || String(rVsk.enable).toLowerCase() !== 'yes') continue;


     const label = rVsk.label || String(vt);
     const vRaw = vskIdx.get([vskModel, vskBody, vskYear, vt].map(NK).join('|')) ?? '';
     const vNum = numval(vRaw);


     out.push([label, '', vNum, '', '', 0, null]);
   }
 }


 // --- Write to sheet (reuse your column targets / formats)
 // Expecting columns: [Label, Towbar, VSK, UniFit, VskFit, Cable, (maybe Total/Notes)]
 writeQuoteRows_(sh, out); // your existing function to write values & apply formats
}




/* -----------------------------------------------------------
* Generate the "Price breakdown" (Part & Price Info)
* -----------------------------------------------------------
* Mirrors the same rule-driven logic as processQuote() so the
* same rows/labels/cable amounts appear here too.
*/
function generatePartPriceInfo() {
 const ss = SpreadsheetApp.getActive();
 const sh = ss.getSheetByName('Quote'); // or whichever sheet hosts the Part & Price Info area


 // Inputs / data
 const data = getQuoteInput_();
 const { quoteTypeUI, pickupYes } = normalizeQuoteType_(data);


 const towPriceIdx   = buildTowPriceIndex_();
 const towUniFitIdx  = buildTowUniFitIndex_();
 const towVskFitIdx  = buildTowVskFitIndex_();
 const vskIdx        = buildVskPriceIndex_();
 const cableIdx      = buildCableIndex_();
 const rules         = loadQuoteRules_();


 const pins          = data.pinFamilies || [];
 const towbarTypes   = data.towbarTypes || [];
 const vskTypes      = data.vskTypes || [];
 const vskBase       = !!data.vskAvailable;
 const vskModel      = String(data.model || '');
 const vskBody       = String(data.body  || '');
 const vskYear       = String(data.year  || '');


 const rows = [];


 // Optional: header/format duplication for the "Price breakdown" blocks
 resetPriceInfoArea_(sh); // your existing reset/styling routine


 const isDetachable = (t) => String(t).toUpperCase().includes('DETACHABLE');


 for (const t of towbarTypes) {
   // (Optional visual group header row – keep if you used one before)
   // rows.push([String(t), '', '', '', '', '', '']);


   for (const p of pins) {
     const pNK = NK(p);
     const towPriceNum = numval(towPriceIdx.get(String(t)) || 0);


     // Towbar only (rule-driven)
     if (!(isDetachable(t) && pNK === 'TWIN')) {
       const rTowOnly = pickRule_(rules, quoteTypeUI, 'Towbar Only', pNK, 'No', t, '');
       if (rTowOnly && String(rTowOnly.enable).toLowerCase() === 'yes') {
         const label = rTowOnly.label || mapPinLabel_(p, /*required*/ false);


         const cab = (() => {
           if (!pickupYes) return 0;
           const override = Number(rTowOnly.cable || 0);
           if (override > 0) return override;
           const price = cableIdx.get(pNK) ?? 0;
           return (typeof price === 'number' ? price : numval(price)) || 0;
         })();


         const uniFit = towUniFitIdx.get(pNK) || 0;


         rows.push([
           label,        // labels from rules
           towPriceNum,  // Towbar price
           '',           // VSK price (none in Towbar-only)
           uniFit,       // Towbar fit time (hrs)
           '',           // VSK fit time (hrs)
           cab,          // Cable cost (override or pin price)
           null          // optional placeholder/notes
         ]);
       }
     }


     // Towbar + VSK (rule-driven)
     if (quoteTypeUI !== 'VSK only' && vskBase && vskTypes.length > 0 && pNK !== 'TWIN') {
       for (const vt of vskTypes) {
         const vtNK = NK(vt);
         const rTowPlus = pickRule_(rules, quoteTypeUI, 'Towbar Plus VSK', pNK, 'Yes', t, vtNK);
         if (!rTowPlus || String(rTowPlus.enable).toLowerCase() !== 'yes') continue;


         const label = rTowPlus.label || mapPinLabel_(p, /*required*/ true);


         const vRaw = vskIdx.get([vskModel, vskBody, vskYear, vt].map(NK).join('|')) ?? '';
         const vNum = numval(vRaw);
         const vskFit = towVskFitIdx.get(pNK) || 0;


         const cab = (() => {
           if (!pickupYes) return 0;
           const override = Number(rTowPlus.cable || 0);
           if (override > 0) return override;
           const price = cableIdx.get(pNK) ?? 0;
           return (typeof price === 'number' ? price : numval(price)) || 0;
         })();


         rows.push([
           label,
           towPriceNum,
           vNum,
           '',
           vskFit,
           cab,
           null
         ]);
       }
     }
   }
 }


 // VSK-only (rule-driven)
 if (quoteTypeUI === 'VSK only' && vskBase && vskTypes.length > 0) {
   for (const vt of vskTypes) {
     const vtNK = NK(vt);
     const rVsk = pickRule_(rules, quoteTypeUI, 'VSK Only', '', 'Yes', '', vtNK);
     if (!rVsk || String(rVsk.enable).toLowerCase() !== 'yes') continue;


     const label = rVsk.label || String(vt);
     const vRaw = vskIdx.get([vskModel, vskBody, vskYear, vt].map(NK).join('|')) ?? '';
     const vNum = numval(vRaw);


     rows.push([label, '', vNum, '', '', 0, null]);
   }
 }


 writePriceInfoRows_(sh, rows); // your existing writer for the Part & Price Info tables
}




/* -----------------------------------------------------------
* Below: helper placeholders that this file expects to exist.
* If your project already defines them elsewhere, remove these
* stubs. They are shown here only to illustrate expected I/O.
* -----------------------------------------------------------
*/


// Gather inputs from your sheet / sidebar form, etc.
function getQuoteInput_() {
 // IMPLEMENTATION NOTE:
 // Return an object with at least:
 // {
 //   quoteType: 'Towbars + Fitting +/- VSK' | 'VSK only' | 'Towbar',
 //   pickupCable: 'Yes' | 'No',
 //   model, body, year,
 //   towbarTypes: [ ... towbar type labels ... ],
 //   pinFamilies: [ '7 PIN', '13 PIN', 'TWIN', ... ],
 //   vskTypes: [ '7 PIN', '13 PIN', ... ],
 //   vskAvailable: true/false
 // }
 // In your live project, you already have this—keep your version.
 throw new Error('getQuoteInput_ not implemented in this stub. Use your project’s version.');
}


// Build indices you already maintain in your project
function buildTowPriceIndex_(){ throw new Error('Stub: use project implementation.'); }
function buildVskPriceIndex_(){ throw new Error('Stub: use project implementation.'); }
function buildTowUniFitIndex_(){ throw new Error('Stub: use project implementation.'); }
function buildTowVskFitIndex_(){ throw new Error('Stub: use project implementation.'); }
function buildCableIndex_(){ throw new Error('Stub: use project implementation.'); }


// Persist output to the Quote sheet (your existing writers)
function resetQuoteTableArea_(_sh){ /* optional: clear rows, formats, etc. */ }
function writeQuoteRows_(_sh, _rows){ throw new Error('Stub: use project implementation.'); }


function resetPriceInfoArea_(_sh){ /* optional: clear rows, formats, etc. */ }
function writePriceInfoRows_(_sh, _rows){ throw new Error('Stub: use project implementation.'); }


// Utilities (expected to exist in your project)
function loadQuoteRules_(){ throw new Error('Stub: use project implementation.'); }
function pickRule_(_rules,_qtUI,_kind,_pinNK,_vskYesNo,_towbarTypeText,_vskTypeNK){ throw new Error('Stub: use project implementation.'); }
function NK(_s){ throw new Error('Stub: use project implementation.'); }
function numval(_x){ throw new Error('Stub: use project implementation.'); }
function mapPinLabel_(_pinText,_required){ throw new Error('Stub: use project implementation.'); }




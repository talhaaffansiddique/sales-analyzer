/*
 * Compile the cleaned multi-sheet workbook (Captain&CO_Sales_2009-2021_Cleaned.xlsx)
 * into src/data.json for the enhanced analyzer.
 *
 * The "Sales Detail" sheet is already one tidy row per transaction, so parsing is
 * straightforward. We additionally:
 *   - carry a real Product Code (SKU) field
 *   - keep Ex-VAT sales as `sales` and Incl-VAT as `salesWithVat` (only 2018+)
 *   - backfill blank formulations from the Formulations catalog when a code/name
 *     maps to exactly one formulation (conservative)
 *   - detect partial years from the "Customer-Year Summary" header (e.g. "2021 (thru May)")
 *   - emit a formulation catalog for the reference view
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, 'Captain&CO_Sales_2009-2021_Cleaned.xlsx');
const OUT = path.join(__dirname, 'src', 'data.json');

const cleanStr = (v) => typeof v === 'string' ? v.replace(/\s+/g, ' ').trim() : v;
const num = (v) => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };

console.log('Loading', SRC);
const wb = XLSX.readFile(SRC);

// ---- Formulation catalog ----
const fm = XLSX.utils.sheet_to_json(wb.Sheets['Formulations']);
const catByCode = {};
const catByName = {};
fm.forEach(r => {
  const code = cleanStr((r['Product Code'] || '').toString());
  const name = cleanStr((r.Product || '').toString());
  const form = cleanStr((r.Formulation || '').toString());
  if (!form) return;
  if (code) (catByCode[code] = catByCode[code] || new Set()).add(form);
  if (name) (catByName[name] = catByName[name] || new Set()).add(form);
});

const formulationCatalog = [];
{
  const seen = new Set();
  fm.forEach(r => {
    const name = cleanStr((r.Product || '').toString());
    const code = cleanStr((r['Product Code'] || '').toString());
    const key = code || name;
    if (!key || seen.has(key)) return;
    seen.add(key);
    const set = (code && catByCode[code]) || catByName[name] || new Set();
    formulationCatalog.push({ product: name, productCode: code || null, formulations: [...set].sort() });
  });
}

// ---- Partial-year detection ----
const cy = XLSX.utils.sheet_to_json(wb.Sheets['Customer-Year Summary'], { header: 1 });
const partialYears = {};
(cy[0] || []).forEach(h => {
  const m = String(h || '').match(/(\d{4})\s*\(thru\s+([A-Za-z]+)\)/i);
  if (m) partialYears[m[1]] = m[2];
});

// ---- Transactions ----
const sd = XLSX.utils.sheet_to_json(wb.Sheets['Sales Detail']);
let backfilled = 0;
const transactions = sd.map(r => {
  const company = cleanStr((r.Customer || '').toString());
  const product = cleanStr((r.Product || '').toString());
  const productCode = cleanStr((r['Product Code'] || '').toString()) || null;
  const year = parseInt(r.Year, 10);
  const qty = num(r['Quantity (KG)']);
  const sales = num(r['Sales (AED, Ex VAT)']);
  const incl = num(r['Sales (AED, Incl VAT)']);
  let formulation = cleanStr((r.Formulation || '').toString());

  if (!formulation) {
    const byCode = productCode && catByCode[productCode];
    const byName = catByName[product];
    if (byCode && byCode.size === 1) { formulation = [...byCode][0]; backfilled++; }
    else if (!byCode && byName && byName.size === 1) { formulation = [...byName][0]; backfilled++; }
  }
  if (!formulation) formulation = 'N/A';

  return {
    company,
    ref: null,               // no customer ref code in the cleaned file
    product,
    productCode,
    formulation,
    year,
    qty,
    sales,                    // AED ex VAT
    salesWithVat: incl > 0 ? incl : sales  // incl VAT (2018+); falls back to ex-VAT
  };
}).filter(t => t.company && t.product && !isNaN(t.year));

const yearsPresent = [...new Set(transactions.map(t => t.year))].sort((a, b) => a - b);

const data = {
  metadata: {
    generatedAt: new Date().toISOString(),
    sourceFile: 'Captain&CO_Sales_2009-2021_Cleaned.xlsx',
    format: 'cleaned-multi-sheet',
    yearMin: yearsPresent[0],
    yearMax: yearsPresent[yearsPresent.length - 1],
    partialYears,               // { "2021": "May" }
    transactionCount: transactions.length,
    formulationsBackfilled: backfilled
  },
  transactions,
  formulationCatalog
};

fs.writeFileSync(OUT, JSON.stringify(data));
console.log('Wrote', OUT);
console.log('  transactions:', transactions.length);
console.log('  formulations backfilled:', backfilled);
console.log('  partial years:', JSON.stringify(partialYears));
console.log('  formulation catalog entries:', formulationCatalog.length);
const totEx = transactions.reduce((s, t) => s + t.sales, 0);
const totKg = transactions.reduce((s, t) => s + t.qty, 0);
console.log('  total ex-VAT AED:', Math.round(totEx).toLocaleString());
console.log('  total KG:', Math.round(totKg).toLocaleString());

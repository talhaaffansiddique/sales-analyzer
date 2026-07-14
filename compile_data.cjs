const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const filePath = path.join(__dirname, 'Historical Report 2009-2021 (1).xls');
const outputPath = path.join(__dirname, 'src', 'data.json');

try {
  console.log('Loading workbook...');
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  
  console.log('Decoding rows...');
  const rawRows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  
  const years = [2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021];
  
  const prodYearCols = {
    2009: { qty: 5, val: 6, vat: null },
    2010: { qty: 10, val: 11, vat: null },
    2011: { qty: 15, val: 16, vat: null },
    2012: { qty: 20, val: 21, vat: null },
    2013: { qty: 25, val: 26, vat: null },
    2014: { qty: 30, val: 31, vat: null },
    2015: { qty: 35, val: 36, vat: null },
    2016: { qty: 40, val: 41, vat: null },
    2017: { qty: 45, val: 46, vat: null },
    2018: { qty: 50, val: 51, vat: 52 },
    2019: { qty: 57, val: 58, vat: 59 },
    2020: { qty: 64, val: 65, vat: 66 },
    2021: { qty: 72, val: 73, vat: 74 }
  };
  
  const totalYearCols = {
    2009: { qty: 7, val: 8, vat: null },
    2010: { qty: 12, val: 13, vat: null },
    2011: { qty: 17, val: 18, vat: null },
    2012: { qty: 22, val: 23, vat: null },
    2013: { qty: 27, val: 28, vat: null },
    2014: { qty: 32, val: 33, vat: null },
    2015: { qty: 37, val: 38, vat: null },
    2016: { qty: 42, val: 43, vat: null },
    2017: { qty: 47, val: 48, vat: null },
    2018: { qty: 53, val: 54, vat: 55 },
    2019: { qty: 60, val: 61, vat: 62 },
    2020: { qty: 67, val: 68, vat: 69 },
    2021: { qty: 75, val: 76, vat: 77 }
  };
  
  const transactions = [];
  const companies = [];
  const products = [];
  const companyMap = {};
  
  let currentCompany = null;
  let currentRef = null;
  
  for (let r = 6; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.length === 0 || (row[0] === null && row[1] === null && row[3] === null)) {
      continue;
    }
    
    const sNo = row[0];
    const companyName = row[1];
    const refNo = row[2];
    const productName = row[3];
    const formulation = row[71]; // Col 71
    
    // Check if it's a total row
    const isTotalRow = (sNo === null || sNo === undefined) && 
                      (companyName === null || companyName === undefined) && 
                      (productName === null || productName === undefined);
                      
    if (isTotalRow) {
      const totalSales = {};
      years.forEach(yr => {
        const colMap = totalYearCols[yr];
        const qty = Number(row[colMap.qty]) || 0;
        const val = Number(row[colMap.val]) || 0;
        const vat = colMap.vat ? (Number(row[colMap.vat]) || 0) : null;
        if (qty > 0 || val > 0) {
          totalSales[yr] = { qty, val, valWithVat: vat !== null ? vat : val };
        }
      });
      
      if (currentCompany && companyMap[currentCompany]) {
        companyMap[currentCompany].totals = totalSales;
      }
      continue;
    }
    
    if (companyName) {
      currentCompany = companyName;
      currentRef = refNo || currentRef;
      if (!companyMap[currentCompany]) {
        companyMap[currentCompany] = {
          name: currentCompany,
          ref: currentRef,
          totals: {},
          products: []
        };
      }
    }
    
    if (productName && String(productName).trim().toUpperCase() !== 'GRAND TOTAL' && currentCompany) {
      if (!companyMap[currentCompany].products.includes(productName)) {
        companyMap[currentCompany].products.push(productName);
      }
      
      if (!products.includes(productName)) {
        products.push(productName);
      }
      
      // Parse yearly sales
      const productSales = {};
      years.forEach(yr => {
        const colMap = prodYearCols[yr];
        const qty = Number(row[colMap.qty]) || 0;
        const val = Number(row[colMap.val]) || 0;
        const vat = colMap.vat ? (Number(row[colMap.vat]) || 0) : null;
        
        if (qty > 0 || val > 0) {
          productSales[yr] = { qty, val, valWithVat: vat !== null ? vat : val };
          transactions.push({
            company: currentCompany,
            ref: currentRef,
            product: productName,
            formulation: formulation || 'N/A',
            year: yr,
            qty,
            sales: val,
            salesWithVat: vat !== null ? vat : val
          });
        }
      });
    }
  }
  
  // Convert company map to array
  Object.keys(companyMap).forEach(key => {
    companies.push(companyMap[key]);
  });
  
  // Calculate summary statistics
  let totalSalesAED = 0;
  let totalQty = 0;
  const salesByYear = {};
  const qtyByYear = {};
  
  years.forEach(yr => {
    salesByYear[yr] = 0;
    qtyByYear[yr] = 0;
  });
  
  transactions.forEach(t => {
    totalSalesAED += t.sales;
    totalQty += t.qty;
    salesByYear[t.year] += t.sales;
    qtyByYear[t.year] += t.qty;
  });
  
  // Sort companies by total sales descending
  companies.forEach(c => {
    c.totalSales = Object.values(c.totals).reduce((sum, t) => sum + t.val, 0);
    c.totalQty = Object.values(c.totals).reduce((sum, t) => sum + t.qty, 0);
  });
  companies.sort((a, b) => b.totalSales - a.totalSales);
  
  // Find top products
  const productSalesMap = {};
  transactions.forEach(t => {
    if (!productSalesMap[t.product]) {
      productSalesMap[t.product] = { product: t.product, sales: 0, qty: 0 };
    }
    productSalesMap[t.product].sales += t.sales;
    productSalesMap[t.product].qty += t.qty;
  });
  const topProducts = Object.values(productSalesMap)
    .sort((a, b) => b.sales - a.sales)
    .slice(0, 20);
  
  const parsedData = {
    metadata: {
      generatedAt: new Date().toISOString(),
      sourceFile: filePath,
      totalRows: rawRows.length
    },
    stats: {
      totalSalesAED,
      totalQty,
      companyCount: companies.length,
      productCount: products.length,
      transactionCount: transactions.length,
      salesByYear,
      qtyByYear,
      topProducts
    },
    companies: companies.slice(0, 100), // top 100 companies by sales to keep JSON size optimized
    allCompanyNames: companies.map(c => c.name), // list of all company names for dropdowns
    topProducts,
    transactions: transactions // all 8,929 transactions (very small JSON size, ~500kb)
  };
  
  // Make sure src/ directory exists
  if (!fs.existsSync(path.join(__dirname, 'src'))) {
    fs.mkdirSync(path.join(__dirname, 'src'));
  }
  
  fs.writeFileSync(outputPath, JSON.stringify(parsedData, null, 2));
  console.log(`Successfully compiled data to ${outputPath}`);
  console.log(`Stats:`);
  console.log(`  Total Sales Value: ${totalSalesAED.toLocaleString()} AED`);
  console.log(`  Total Qty: ${totalQty.toLocaleString()}`);
  console.log(`  Total Companies: ${companies.length}`);
  console.log(`  Total Transactions: ${transactions.length}`);
} catch (e) {
  console.error('Error compiling data:', e);
}

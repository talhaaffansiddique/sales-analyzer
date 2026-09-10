import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  FileSpreadsheet, 
  TrendingUp, 
  Users, 
  Layers, 
  Search, 
  Download, 
  Bot, 
  Send, 
  Key, 
  BarChart3, 
  SlidersHorizontal,
  RefreshCw,
  HelpCircle,
  Check,
  ChevronLeft,
  ChevronRight,
  Sparkles,
  Database,
  Menu,
  X,
  ChevronsLeft,
  ChevronsRight
} from 'lucide-react';
import { Line, Bar, Pie } from 'react-chartjs-2';
import * as XLSX from 'xlsx';
import Anthropic from '@anthropic-ai/sdk';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ArcElement
} from 'chart.js';

// Import our pre-compiled database
import compiledData from './data.json';

// Register ChartJS modules
ChartJS.register(
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ArcElement
);

// Format helper
const formatCurrency = (val) => {
  const num = Number(val) || 0;
  if (num >= 1e6) return `${(num / 1e6).toFixed(2)}M AED`;
  if (num >= 1e3) return `${(num / 1e3).toFixed(1)}K AED`;
  return `${num.toFixed(0)} AED`;
};

const formatNumber = (val) => {
  const num = Number(val) || 0;
  if (num >= 1e6) return `${(num / 1e6).toFixed(2)}M`;
  if (num >= 1e3) return `${(num / 1e3).toFixed(1)}K`;
  return num.toLocaleString(undefined, { maximumFractionDigits: 0 });
};

const cleanStr = (v) => typeof v === 'string' ? v.replace(/\s+/g, ' ').trim() : v;

// Wide-format parser: mirrors compile_data.cjs's layout for the original
// "Historical Report 2009-2021" spreadsheet shape (one qty/value column block
// per year, data starting row 6). Only re-parses files with that exact shape
// (e.g. a corrected or re-exported copy of the original report) — it does not
// know how to locate columns for years outside 2009-2021.
const WIDE_FORMAT_YEARS = [2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021];
const WIDE_FORMAT_PROD_YEAR_COLS = {
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

const looksLikeWideFormatReport = (rawRows) => {
  if (!rawRows || rawRows.length < 8) return false;
  // Scan the first ~30 data rows for at least one with a company or product name
  // in the expected columns, confirming this matches the original report layout.
  for (let r = 6; r < Math.min(rawRows.length, 36); r++) {
    const row = rawRows[r];
    if (row && (row[1] || row[3])) return true;
  }
  return false;
};

const parseWideFormatRows = (rawRows) => {
  const transactions = [];
  let currentCompany = null;
  let currentRef = null;

  for (let r = 6; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.length === 0 || (row[0] === null && row[1] === null && row[3] === null)) {
      continue;
    }

    const companyName = cleanStr(row[1]);
    const refNo = row[2];
    const productName = cleanStr(row[3]);
    const formulation = cleanStr(row[71]);

    const isTotalRow = !row[0] && !companyName && !productName;
    if (isTotalRow) continue;

    if (companyName) {
      currentCompany = companyName;
      currentRef = refNo || currentRef;
    }

    if (productName && String(productName).trim().toUpperCase() !== 'GRAND TOTAL' && currentCompany) {
      WIDE_FORMAT_YEARS.forEach(yr => {
        const colMap = WIDE_FORMAT_PROD_YEAR_COLS[yr];
        const qty = Number(row[colMap.qty]) || 0;
        const val = Number(row[colMap.val]) || 0;
        const vat = colMap.vat ? (Number(row[colMap.vat]) || 0) : null;

        if (qty > 0 || val > 0) {
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

  return transactions;
};

// Tidy CSV parser: expects one row per transaction with the same column
// headers the app's own "Export Excel" produces, so an exported-then-edited
// file round-trips cleanly. Header matching is case-insensitive and tolerant
// of a few common aliases.
const TIDY_HEADER_ALIASES = {
  company: ['company name', 'company'],
  ref: ['ref code', 'ref', 'reference'],
  product: ['product'],
  formulation: ['formulation'],
  year: ['year'],
  qty: ['quantity', 'qty'],
  sales: ['sales value (aed)', 'sales', 'sales value'],
  salesWithVat: ['sales with vat (aed)', 'sales with vat', 'salesv with vat']
};

const looksLikeTidyCsv = (rawRows) => {
  if (!rawRows || rawRows.length < 1) return false;
  const header = (rawRows[0] || []).map(h => cleanStr(String(h || '')).toLowerCase());
  const hasCompany = TIDY_HEADER_ALIASES.company.some(a => header.includes(a));
  const hasProduct = TIDY_HEADER_ALIASES.product.some(a => header.includes(a));
  const hasYear = TIDY_HEADER_ALIASES.year.some(a => header.includes(a));
  return hasCompany && hasProduct && hasYear;
};

const parseTidyCsvRows = (rawRows) => {
  const header = (rawRows[0] || []).map(h => cleanStr(String(h || '')).toLowerCase());
  const colIndex = {};
  Object.keys(TIDY_HEADER_ALIASES).forEach(field => {
    const idx = header.findIndex(h => TIDY_HEADER_ALIASES[field].includes(h));
    colIndex[field] = idx;
  });

  const transactions = [];
  for (let r = 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.length === 0) continue;
    const company = colIndex.company >= 0 ? cleanStr(row[colIndex.company]) : null;
    const product = colIndex.product >= 0 ? cleanStr(row[colIndex.product]) : null;
    const year = colIndex.year >= 0 ? Number(row[colIndex.year]) : null;
    if (!company || !product || !year) continue;

    const qty = colIndex.qty >= 0 ? Number(row[colIndex.qty]) || 0 : 0;
    const sales = colIndex.sales >= 0 ? Number(row[colIndex.sales]) || 0 : 0;
    const salesWithVatRaw = colIndex.salesWithVat >= 0 ? Number(row[colIndex.salesWithVat]) : NaN;

    transactions.push({
      company,
      ref: colIndex.ref >= 0 ? cleanStr(row[colIndex.ref]) : null,
      product,
      formulation: colIndex.formulation >= 0 ? (cleanStr(row[colIndex.formulation]) || 'N/A') : 'N/A',
      year,
      qty,
      sales,
      salesWithVat: !isNaN(salesWithVatRaw) ? salesWithVatRaw : sales
    });
  }

  return transactions;
};

export default function App() {
  // Authentication states
  const [isAuthenticated, setIsAuthenticated] = useState(sessionStorage.getItem('cornell_authenticated') === 'true');
  const [usernameInput, setUsernameInput] = useState('');
  const [passwordInput, setPasswordInput] = useState('');
  const [loginError, setLoginError] = useState('');

  // App states
  const [activeTab, setActiveTab] = useState('overview');
  const [transactions, setTransactions] = useState(compiledData.transactions);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isChartBuilderEnabled, setIsChartBuilderEnabled] = useState(false);

  // Overview dashboard year filter
  const [overviewFromYear, setOverviewFromYear] = useState(2009);
  const [overviewToYear, setOverviewToYear] = useState(2021);
  const [overviewTrendChartType, setOverviewTrendChartType] = useState('line');

  // Table filters
  const [searchTerm, setSearchTerm] = useState('');
  const [fromYear, setFromYear] = useState(2009);
  const [toYear, setToYear] = useState(2021);
  const [filterFormulation, setFilterFormulation] = useState('All');
  const [selectedProducts, setSelectedProducts] = useState([]);
  const [productSearchInput, setProductSearchInput] = useState('');
  const [sortKey, setSortKey] = useState('sales');
  const [sortDirection, setSortDirection] = useState('desc');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);

  // Chart builder states
  const [chartType, setChartType] = useState('bar');
  const [chartXAxis, setChartXAxis] = useState('year');
  const [chartYAxis, setChartYAxis] = useState('sales');
  const [chartLimit, setChartLimit] = useState('10');

  // AI Chat states
  const getWelcomeMessage = () => ({
    sender: 'ai',
    text: "Hello! I am your AI Data Assistant. I have analyzed **Historical Report 2009-2021 (1).xls**.\n\nYou can ask me questions about this dataset, such as which companies have the highest sales, sales trends by year, or product performance. Type \"help\" to see everything I can do, or enter a Gemini API Key in the settings above to unlock advanced natural language questions!",
    timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
  });
  const [messages, setMessages] = useState(() => {
    try {
      const saved = localStorage.getItem('cornell_chat_history');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch (e) {
      // Ignore corrupt saved history and fall back to the welcome message
    }
    return [getWelcomeMessage()];
  });

  // Persist chat history across reloads
  useEffect(() => {
    localStorage.setItem('cornell_chat_history', JSON.stringify(messages));
  }, [messages]);

  const handleClearChat = () => {
    const fresh = [getWelcomeMessage()];
    setMessages(fresh);
  };
  const [chatInput, setChatInput] = useState('');
  const [aiProvider, setAiProvider] = useState(localStorage.getItem('ai_provider') || 'gemini');
  const [apiKey, setApiKey] = useState(localStorage.getItem('gemini_api_key') || '');
  const [claudeApiKey, setClaudeApiKey] = useState(localStorage.getItem('claude_api_key') || '');
  const [showApiSettings, setShowApiSettings] = useState(false);
  const [isAiTyping, setIsAiTyping] = useState(false);
  const chatEndRef = useRef(null);

  // Scroll to bottom of chat
  useEffect(() => {
    if (chatEndRef.current) {
      chatEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [messages, isAiTyping]);

  // Chart.js canvases in this layout can get measured against a stale
  // container size (e.g. right after a tab mounts, or when the sidebar
  // opens/closes and changes available width) and never self-correct —
  // even an explicit chart.resize() call doesn't reliably recompute against
  // the canvas's *current* container. The reliable fix is to fully remount
  // the chart (fresh mounts always measure correctly) whenever something
  // that changes chart container size happens, by changing its React key.
  const [chartLayoutKey, setChartLayoutKey] = useState(0);

  useEffect(() => {
    // Bump once immediately (covers tab switches), then again after the
    // sidebar's CSS transition (--transition-smooth, 0.3s) settles.
    setChartLayoutKey(k => k + 1);
    const timeoutId = setTimeout(() => setChartLayoutKey(k => k + 1), 350);
    return () => clearTimeout(timeoutId);
  }, [activeTab, isSidebarOpen, overviewFromYear, overviewToYear, chartType]);

  useEffect(() => {
    let debounceId;
    const handleResize = () => {
      clearTimeout(debounceId);
      debounceId = setTimeout(() => setChartLayoutKey(k => k + 1), 150);
    };
    window.addEventListener('resize', handleResize);
    return () => {
      window.removeEventListener('resize', handleResize);
      clearTimeout(debounceId);
    };
  }, []);

  // Handle Login authentication
  const handleLogin = (e) => {
    e.preventDefault();
    if (usernameInput === 'captain' && passwordInput === 'talha123') {
      setIsAuthenticated(true);
      setLoginError('');
      sessionStorage.setItem('cornell_authenticated', 'true');
    } else {
      setLoginError('Invalid Username or Password');
    }
  };

  // Handle Logout
  const handleLogout = () => {
    setIsAuthenticated(false);
    setUsernameInput('');
    setPasswordInput('');
    sessionStorage.removeItem('cornell_authenticated');
  };

  // Save API Key (provider-aware: Gemini or Claude)
  const handleSaveApiKey = (e) => {
    e.preventDefault();
    localStorage.setItem('ai_provider', aiProvider);
    const activeKey = aiProvider === 'claude' ? claudeApiKey : apiKey;
    if (aiProvider === 'claude') {
      localStorage.setItem('claude_api_key', claudeApiKey);
    } else {
      localStorage.setItem('gemini_api_key', apiKey);
    }
    setShowApiSettings(false);
    const providerName = aiProvider === 'claude' ? 'Claude' : 'Gemini';
    setMessages(prev => [
      ...prev,
      {
        sender: 'ai',
        text: activeKey
          ? `✅ **${providerName} API Key saved successfully!** You can now ask complex questions and request deep business analysis on your dataset.`
          : `⚠️ **${providerName} API Key removed.** Chat will now run in offline local query mode.`,
        timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
      }
    ]);
  };

  // Get filter list options (reactive to imported/live transactions)
  const formulationOptions = useMemo(() => {
    const formulations = new Set();
    transactions.forEach(t => {
      if (t.formulation && t.formulation !== 'N/A') {
        formulations.add(t.formulation);
      }
    });
    return ['All', 'N/A', ...Array.from(formulations).sort()];
  }, [transactions]);

  const uniqueProductsList = useMemo(() => {
    const productsSet = new Set();
    transactions.forEach(t => {
      productsSet.add(t.product);
    });
    return Array.from(productsSet).sort();
  }, [transactions]);

  const yearsList = useMemo(() => {
    const yearsSet = new Set(transactions.map(t => t.year));
    return Array.from(yearsSet).sort((a, b) => a - b);
  }, [transactions]);

  // Full dataset stats, recomputed whenever transactions change (e.g. after an import)
  const globalStats = useMemo(() => {
    let totalSalesAED = 0;
    let totalQty = 0;
    const salesByYear = {};
    const qtyByYear = {};
    const companyMap = {};
    const productMap = {};

    transactions.forEach(t => {
      totalSalesAED += t.sales;
      totalQty += t.qty;
      salesByYear[t.year] = (salesByYear[t.year] || 0) + t.sales;
      qtyByYear[t.year] = (qtyByYear[t.year] || 0) + t.qty;

      if (!companyMap[t.company]) {
        companyMap[t.company] = { name: t.company, ref: t.ref, totalSales: 0, totalQty: 0 };
      }
      companyMap[t.company].totalSales += t.sales;
      companyMap[t.company].totalQty += t.qty;

      if (!productMap[t.product]) {
        productMap[t.product] = { product: t.product, sales: 0, qty: 0 };
      }
      productMap[t.product].sales += t.sales;
      productMap[t.product].qty += t.qty;
    });

    const companies = Object.values(companyMap).sort((a, b) => b.totalSales - a.totalSales);
    const topProducts = Object.values(productMap).sort((a, b) => b.sales - a.sales).slice(0, 20);

    return {
      totalSalesAED,
      totalQty,
      companyCount: companies.length,
      productCount: Object.keys(productMap).length,
      salesByYear,
      qtyByYear,
      topProducts,
      companies,
      allCompanyNames: companies.map(c => c.name),
      allProductNames: Object.keys(productMap)
    };
  }, [transactions]);

  // Process table data (Search + Filters + Sort)
  const filteredAndSortedTransactions = useMemo(() => {
    let result = [...transactions];

    // Global Search
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      result = result.filter(t => 
        t.company.toLowerCase().includes(term) || 
        t.product.toLowerCase().includes(term) ||
        (t.ref && t.ref.toLowerCase().includes(term))
      );
    }

    // Filter by Year Range
    result = result.filter(t => t.year >= fromYear && t.year <= toYear);

    // Filter by Formulation
    if (filterFormulation !== 'All') {
      result = result.filter(t => t.formulation === filterFormulation);
    }

    // Filter by Selected Products (Multi-select)
    if (selectedProducts.length > 0) {
      result = result.filter(t => selectedProducts.includes(t.product));
    }

    // Sorting
    result.sort((a, b) => {
      let aVal = a[sortKey];
      let bVal = b[sortKey];

      // Handle nulls/undefined
      if (aVal === undefined || aVal === null) aVal = '';
      if (bVal === undefined || bVal === null) bVal = '';

      // Check if both are numeric (either numbers or numeric strings)
      const aNum = Number(aVal);
      const bNum = Number(bVal);
      const isNumeric = !isNaN(aNum) && !isNaN(bNum) && aVal !== '' && bVal !== '';

      if (isNumeric) {
        return sortDirection === 'asc' ? aNum - bNum : bNum - aNum;
      }

      // Fallback to string comparison
      const aStr = String(aVal);
      const bStr = String(bVal);

      return sortDirection === 'asc'
        ? aStr.localeCompare(bStr)
        : bStr.localeCompare(aStr);
    });

    return result;
  }, [transactions, searchTerm, fromYear, toYear, filterFormulation, selectedProducts, sortKey, sortDirection]);

  // Aggregated totals for currently filtered rows
  const filteredTotals = useMemo(() => {
    let qty = 0;
    let sales = 0;
    let salesWithVat = 0;
    filteredAndSortedTransactions.forEach(t => {
      qty += t.qty;
      sales += t.sales;
      salesWithVat += t.salesWithVat;
    });
    return { qty, sales, salesWithVat };
  }, [filteredAndSortedTransactions]);

  // Paginated Data
  const paginatedTransactions = useMemo(() => {
    const startIdx = (currentPage - 1) * pageSize;
    return filteredAndSortedTransactions.slice(startIdx, startIdx + pageSize);
  }, [filteredAndSortedTransactions, currentPage, pageSize]);

  const totalPages = Math.ceil(filteredAndSortedTransactions.length / pageSize);

  // Sorting Handler
  const handleSort = (key) => {
    if (sortKey === key) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setSortKey(key);
      setSortDirection('desc');
    }
    setCurrentPage(1);
  };

  // Export current filtered rows to Excel
  const handleExportExcel = () => {
    const dataToExport = filteredAndSortedTransactions.map(t => ({
      'Company Name': t.company,
      'Ref Code': t.ref || 'N/A',
      'Product': t.product,
      'Formulation': t.formulation,
      'Year': t.year,
      'Quantity': t.qty,
      'Sales Value (AED)': t.sales,
      'Sales With VAT (AED)': t.salesWithVat
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Filtered Sales');
    XLSX.writeFile(workbook, 'Cornell_Sales_Filtered_Export.xlsx');
  };

  // Import a new Excel/CSV file — detects whether it's the original wide report
  // shape or the app's own tidy export-column shape, parses it client-side, and
  // upserts the resulting rows into the live dataset (keyed on company+product+year).
  const importFileInputRef = useRef(null);
  const [importStatus, setImportStatus] = useState(null); // { type: 'success' | 'error', text }

  const resetYearFilters = (minYear, maxYear) => {
    setOverviewFromYear(minYear);
    setOverviewToYear(maxYear);
    setFromYear(minYear);
    setToYear(maxYear);
  };

  const handleImportFile = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = ''; // allow re-selecting the same file later
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        let imported;
        let formatLabel;
        if (looksLikeTidyCsv(rawRows)) {
          imported = parseTidyCsvRows(rawRows);
          formatLabel = 'tidy CSV template';
        } else if (looksLikeWideFormatReport(rawRows)) {
          imported = parseWideFormatRows(rawRows);
          formatLabel = 'original wide report format';
        } else {
          throw new Error('Unrecognized file format. Use either the tidy CSV template (Company, Product, Year, Quantity, Sales... columns) or a file shaped like the original Historical Report spreadsheet.');
        }

        if (imported.length === 0) {
          throw new Error('No transaction rows were found in this file.');
        }

        const keyOf = (t) => `${t.company}|${t.product}|${t.year}`;
        const merged = [...transactions];
        const indexByKey = new Map(merged.map((t, i) => [keyOf(t), i]));
        let addedCount = 0;
        let updatedCount = 0;

        imported.forEach(t => {
          const k = keyOf(t);
          if (indexByKey.has(k)) {
            merged[indexByKey.get(k)] = t;
            updatedCount++;
          } else {
            merged.push(t);
            indexByKey.set(k, merged.length - 1);
            addedCount++;
          }
        });

        setTransactions(merged);

        const mergedYears = merged.map(t => t.year);
        resetYearFilters(Math.min(...mergedYears), Math.max(...mergedYears));

        const summary = `✅ **Import complete** (${formatLabel}, from "${file.name}"): ${addedCount.toLocaleString()} new rows added, ${updatedCount.toLocaleString()} existing rows updated.`;
        setImportStatus({ type: 'success', text: summary });
        setMessages(prev => [
          ...prev,
          { sender: 'ai', text: summary, timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) }
        ]);
      } catch (err) {
        console.error(err);
        const errorText = `❌ **Import failed:** ${err.message || 'Could not parse this file.'}`;
        setImportStatus({ type: 'error', text: errorText });
        setMessages(prev => [
          ...prev,
          { sender: 'ai', text: errorText, timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) }
        ]);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleReloadOriginalData = () => {
    setTransactions(compiledData.transactions);
    resetYearFilters(2009, 2021);
    setImportStatus(null);
    setMessages(prev => [
      ...prev,
      { sender: 'ai', text: '↺ **Dataset reset** to the original bundled Historical Report (2009-2021).', timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) }
    ]);
  };

  // Aggregate stats for the Overview dashboard, recomputed for the selected year range.
  // Falls back to the already-computed globalStats when the full range is selected (no extra work needed).
  const overviewStats = useMemo(() => {
    const minYear = yearsList[0];
    const maxYear = yearsList[yearsList.length - 1];
    if (overviewFromYear <= minYear && overviewToYear >= maxYear) {
      return globalStats;
    }

    const filteredTx = transactions.filter(
      t => t.year >= overviewFromYear && t.year <= overviewToYear
    );

    let totalSalesAED = 0;
    let totalQty = 0;
    const salesByYear = {};
    const qtyByYear = {};
    yearsList.forEach(yr => { salesByYear[yr] = 0; qtyByYear[yr] = 0; });

    const companyMap = {};
    const productMap = {};

    filteredTx.forEach(t => {
      totalSalesAED += t.sales;
      totalQty += t.qty;
      salesByYear[t.year] += t.sales;
      qtyByYear[t.year] += t.qty;

      if (!companyMap[t.company]) {
        companyMap[t.company] = { name: t.company, ref: t.ref, totalSales: 0, totalQty: 0 };
      }
      companyMap[t.company].totalSales += t.sales;
      companyMap[t.company].totalQty += t.qty;

      if (!productMap[t.product]) {
        productMap[t.product] = { product: t.product, sales: 0, qty: 0 };
      }
      productMap[t.product].sales += t.sales;
      productMap[t.product].qty += t.qty;
    });

    const companies = Object.values(companyMap).sort((a, b) => b.totalSales - a.totalSales);
    const topProducts = Object.values(productMap).sort((a, b) => b.sales - a.sales).slice(0, 20);

    return {
      totalSalesAED,
      totalQty,
      companyCount: companies.length,
      productCount: Object.keys(productMap).length,
      salesByYear,
      qtyByYear,
      topProducts,
      companies
    };
  }, [overviewFromYear, overviewToYear, transactions, globalStats, yearsList]);

  // Top Companies list for Overview Tab
  const topCompanies = useMemo(() => {
    return overviewStats.companies.slice(0, 5);
  }, [overviewStats]);

  // Overview Main Chart Data (Sales & Qty over years)
  const overviewChartData = useMemo(() => {
    const sortedYears = [...yearsList].filter(yr => yr >= overviewFromYear && yr <= overviewToYear).sort();
    const salesData = sortedYears.map(yr => overviewStats.salesByYear[yr] || 0);
    const qtyData = sortedYears.map(yr => overviewStats.qtyByYear[yr] || 0);
    const isBar = overviewTrendChartType === 'bar';

    return {
      labels: sortedYears,
      datasets: [
        {
          label: 'Total Revenue (AED)',
          data: salesData,
          borderColor: '#6366f1', // Indigo
          backgroundColor: isBar ? 'rgba(99, 102, 241, 0.85)' : 'rgba(99, 102, 241, 0.2)',
          borderWidth: isBar ? 0 : 3,
          borderRadius: isBar ? 4 : 0,
          tension: 0.3,
          fill: true,
          yAxisID: 'y'
        },
        {
          label: 'Units Sold (Qty)',
          data: qtyData,
          borderColor: '#06b6d4', // Cyan
          backgroundColor: isBar ? 'rgba(6, 182, 212, 0.85)' : 'rgba(6, 182, 212, 0.1)',
          borderWidth: isBar ? 0 : 2,
          borderRadius: isBar ? 4 : 0,
          tension: 0.3,
          borderDash: isBar ? undefined : [5, 5],
          yAxisID: 'y1'
        }
      ]
    };
  }, [overviewStats, overviewFromYear, overviewToYear, overviewTrendChartType]);

  const overviewChartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: 'top',
        labels: { color: '#fafafa', font: { family: 'Outfit', size: 12 } }
      },
      tooltip: {
        padding: 12,
        backgroundColor: '#18181b',
        titleFont: { family: 'Outfit', size: 13, weight: 'bold' },
        bodyFont: { family: 'Outfit', size: 12 },
        borderColor: '#3f3f46',
        borderWidth: 1
      }
    },
    scales: {
      x: {
        grid: { color: 'rgba(63, 63, 70, 0.2)' },
        ticks: { color: '#a1a1aa', font: { family: 'Outfit' } }
      },
      y: {
        type: 'linear',
        display: true,
        position: 'left',
        grid: { color: 'rgba(63, 63, 70, 0.2)' },
        ticks: { 
          color: '#a1a1aa', 
          font: { family: 'Outfit' },
          callback: (value) => value >= 1e6 ? `${(value / 1e6).toFixed(1)}M` : value.toLocaleString()
        }
      },
      y1: {
        type: 'linear',
        display: true,
        position: 'right',
        grid: { drawOnChartArea: false },
        ticks: { 
          color: '#a1a1aa', 
          font: { family: 'Outfit' },
          callback: (value) => value >= 1e6 ? `${(value / 1e6).toFixed(1)}M` : value.toLocaleString()
        }
      }
    }
  };

  // Top Products Chart Data
  const topProductsChartData = useMemo(() => {
    const list = overviewStats.topProducts.slice(0, 7);
    return {
      labels: list.map(p => p.product.substring(0, 15) + (p.product.length > 15 ? '...' : '')),
      datasets: [
        {
          label: 'Total Sales (AED)',
          data: list.map(p => p.sales),
          backgroundColor: [
            'rgba(99, 102, 241, 0.85)',
            'rgba(168, 85, 247, 0.85)',
            'rgba(6, 182, 212, 0.85)',
            'rgba(16, 185, 129, 0.85)',
            'rgba(245, 158, 11, 0.85)',
            'rgba(239, 68, 68, 0.85)',
            'rgba(236, 72, 153, 0.85)'
          ],
          borderWidth: 0,
          borderRadius: 4
        }
      ]
    };
  }, [overviewStats]);

  // Process data for Chart Builder
  const builderChartData = useMemo(() => {
    const aggregationMap = {};

    filteredAndSortedTransactions.forEach(t => {
      let keyVal = t[chartXAxis];
      if (chartXAxis === 'year') keyVal = String(t.year);
      if (!keyVal) keyVal = 'N/A';

      const numericVal = chartYAxis === 'sales' ? t.sales : t.qty;

      if (!aggregationMap[keyVal]) {
        aggregationMap[keyVal] = 0;
      }
      aggregationMap[keyVal] += numericVal;
    });

    let items = Object.keys(aggregationMap).map(k => ({
      label: k,
      value: aggregationMap[k]
    }));

    // Sort appropriately
    if (chartXAxis === 'year') {
      items.sort((a, b) => parseInt(a.label) - parseInt(b.label));
    } else {
      items.sort((a, b) => b.value - a.value);
    }

    // Apply Limit
    if (chartLimit !== 'all') {
      const limitNum = parseInt(chartLimit);
      items = items.slice(0, limitNum);
    }

    const labels = items.map(item => {
      if (item.label.length > 25) return item.label.substring(0, 22) + '...';
      return item.label;
    });
    const values = items.map(item => item.value);

    return {
      labels,
      datasets: [
        {
          label: chartYAxis === 'sales' ? 'Sales Revenue (AED)' : 'Quantity Sold',
          data: values,
          backgroundColor: chartType === 'line' ? 'rgba(99, 102, 241, 0.2)' : 'rgba(99, 102, 241, 0.8)',
          borderColor: '#6366f1',
          borderWidth: chartType === 'line' ? 3 : 1,
          tension: 0.3,
          fill: chartType === 'line',
          borderRadius: 6
        }
      ]
    };
  }, [filteredAndSortedTransactions, chartXAxis, chartYAxis, chartType, chartLimit]);

  // AI local engine query parser
  const runLocalAIQuery = (query) => {
    const q = query.toLowerCase().trim();

    // Help / capability listing
    if (q === 'help' || q.includes('what can you do') || q.includes('what can i ask') || q.includes('what questions can')) {
      return `🤖 **Here's what I can answer offline (no API key needed):**\n- "Top products" / "Top companies"\n- "Total sales" / "How many companies/products"\n- "Show sales in 2018" (any year 2009-2021)\n- "Compare 2020 and 2021" (or compare two companies/products)\n- "Sales trend" / "Growth by year"\n- Ask about a specific company or product by name and I'll pull its full sales history\n\nFor open-ended analysis, add a **Gemini API Key** in settings.`;
    }

    // Comparison queries: "compare X and Y" / "X vs Y"
    if (q.includes('compare') || q.includes(' vs ') || q.includes(' versus ')) {
      const mentionedYears = yearsList.filter(yr => q.includes(String(yr)));
      if (mentionedYears.length >= 2) {
        const [yrA, yrB] = mentionedYears;
        const salesA = globalStats.salesByYear[yrA] || 0;
        const salesB = globalStats.salesByYear[yrB] || 0;
        const qtyA = globalStats.qtyByYear[yrA] || 0;
        const qtyB = globalStats.qtyByYear[yrB] || 0;
        const pctChange = salesA !== 0 ? (((salesB - salesA) / salesA) * 100).toFixed(1) : 'N/A';
        return `📊 **${yrA} vs ${yrB}:**\n\n| Metric | ${yrA} | ${yrB} |\n|---|---|---|\n| Revenue | ${salesA.toLocaleString()} AED | ${salesB.toLocaleString()} AED |\n| Units Sold | ${qtyA.toLocaleString()} | ${qtyB.toLocaleString()} |\n\nRevenue changed by **${pctChange}%** from ${yrA} to ${yrB}.`;
      }

      const mentionedCompanies = findAllMentionedCompanies(q);
      if (mentionedCompanies.length >= 2) {
        const [nameA, nameB] = mentionedCompanies;
        const compA = globalStats.companies.find(c => c.name === nameA);
        const compB = globalStats.companies.find(c => c.name === nameB);
        if (compA && compB) {
          return `🏢 **${compA.name} vs ${compB.name}:**\n\n| Metric | ${compA.name} | ${compB.name} |\n|---|---|---|\n| Total Sales | ${compA.totalSales.toLocaleString()} AED | ${compB.totalSales.toLocaleString()} AED |\n| Total Quantity | ${compA.totalQty.toLocaleString()} | ${compB.totalQty.toLocaleString()} |`;
        }
      }

      const mentionedProducts = findAllMentionedProducts(q);
      if (mentionedProducts.length >= 2) {
        const [prodA, prodB] = mentionedProducts;
        const agg = (name) => {
          const rows = transactions.filter(t => t.product === name);
          return rows.reduce((acc, t) => ({ sales: acc.sales + t.sales, qty: acc.qty + t.qty }), { sales: 0, qty: 0 });
        };
        const aggA = agg(prodA);
        const aggB = agg(prodB);
        return `📦 **${prodA} vs ${prodB}:**\n\n| Metric | ${prodA} | ${prodB} |\n|---|---|---|\n| Total Sales | ${aggA.sales.toLocaleString()} AED | ${aggB.sales.toLocaleString()} AED |\n| Total Quantity | ${aggA.qty.toLocaleString()} | ${aggB.qty.toLocaleString()} |`;
      }

      return `🤔 I can compare two **years** (e.g. "compare 2019 and 2021"), two **companies**, or two **products** — but I need to recognize both names/years in your question. Try being more specific, or use full company/product names.`;
    }

    // Trend / growth queries
    if (q.includes('trend') || q.includes('growth') || q.includes('yoy') || q.includes('year over year') || q.includes('year-over-year')) {
      const sortedYears = [...yearsList].sort();
      let bestYear = null, bestGrowth = -Infinity, worstYear = null, worstGrowth = Infinity;
      const lines = sortedYears.map((yr, idx) => {
        const sales = globalStats.salesByYear[yr] || 0;
        if (idx === 0) return `- **${yr}:** ${sales.toLocaleString()} AED`;
        const prevSales = globalStats.salesByYear[sortedYears[idx - 1]] || 0;
        const growth = prevSales !== 0 ? ((sales - prevSales) / prevSales) * 100 : 0;
        if (growth > bestGrowth) { bestGrowth = growth; bestYear = yr; }
        if (growth < worstGrowth) { worstGrowth = growth; worstYear = yr; }
        const arrow = growth >= 0 ? '▲' : '▼';
        return `- **${yr}:** ${sales.toLocaleString()} AED (${arrow} ${growth.toFixed(1)}%)`;
      });
      return `📈 **Revenue Trend (2009 - 2021):**\n\n${lines.join('\n')}\n\n**Strongest growth:** ${bestYear} (${bestGrowth.toFixed(1)}%)\n**Weakest growth:** ${worstYear} (${worstGrowth.toFixed(1)}%)`;
    }

    // Check for "how many companies" or similar
    if (q.includes('how many companies') || q.includes('how many customers') || q.includes('total companies')) {
      return `📊 There are **${globalStats.companyCount} unique companies** present in this sales report. The top customer by sales is **${globalStats.companies[0].name}**, with a total sales volume of **${globalStats.companies[0].totalSales.toLocaleString()} AED** over the active years.`;
    }

    // Check for "how many products" or similar
    if (q.includes('how many products') || q.includes('total products')) {
      return `📦 There are **${globalStats.productCount.toLocaleString()} unique products** listed in this report. The highest-selling product by revenue is **${globalStats.topProducts[0].product}**, generating **${globalStats.topProducts[0].sales.toLocaleString()} AED** in revenue.`;
    }

    // Check for "total sales" or "total revenue"
    if (q.includes('total sales') || q.includes('total revenue') || q.includes('sales volume')) {
      return `💰 The total sales revenue aggregated from **2009 to 2021** is **${globalStats.totalSalesAED.toLocaleString(undefined, {maximumFractionDigits:2})} AED**, representing a total quantity of **${globalStats.totalQty.toLocaleString()} units** sold.`;
    }

    // Check for top products
    if (q.includes('top product') || q.includes('highest selling product') || q.includes('best product')) {
      const topList = globalStats.topProducts.slice(0, 5).map((p, idx) => `${idx + 1}. **${p.product}** - Sales: *${p.sales.toLocaleString()} AED* (Qty: ${p.qty.toLocaleString()})`).join('\n');
      return `🏆 Here are the **Top 5 Products** by sales revenue:\n\n${topList}`;
    }

    // Check for top companies
    if (q.includes('top company') || q.includes('highest sales company') || q.includes('best customer')) {
      const topList = globalStats.companies.slice(0, 5).map((c, idx) => `${idx + 1}. **${c.name}** - Total Sales: *${c.totalSales.toLocaleString()} AED* (Ref: ${c.ref || 'N/A'})`).join('\n');
      return `🏢 Here are the **Top 5 Companies** by sales revenue:\n\n${topList}`;
    }

    // Check for a specific year
    for (const yr of yearsList) {
      if (q.includes(String(yr))) {
        const sales = globalStats.salesByYear[yr] || 0;
        const qty = globalStats.qtyByYear[yr] || 0;
        return `📅 **Sales Summary for Year ${yr}:**\n- **Total Revenue:** ${sales.toLocaleString(undefined, {maximumFractionDigits: 2})} AED\n- **Volume Sold:** ${qty.toLocaleString()} units\n- **Contribution:** ${((sales / globalStats.totalSalesAED) * 100).toFixed(2)}% of historical total sales.`;
      }
    }

    // Check if a specific company is mentioned in the query
    const compName = findMentionedCompany(q);
    if (compName) {
      // Aggregate data for this company
      const compData = globalStats.companies.find(c => c.name === compName);
      if (compData) {
        const historyText = Object.keys(compData.totals).map(yr => {
          const t = compData.totals[yr];
          return `- **Year ${yr}:** Qty ${t.qty.toLocaleString()}, Sales ${t.val.toLocaleString()} AED`;
        }).join('\n');

        return `🏢 **Sales report for ${compData.name}** (Ref: ${compData.ref || 'N/A'}):\n- **Total Sales:** ${compData.totalSales.toLocaleString()} AED\n- **Total Quantity:** ${compData.totalQty.toLocaleString()} units\n\n**Annual Performance Breakdown:**\n${historyText || 'No sales records found'}`;
      }
    }

    // Check if a specific product is mentioned in the query
    const prodName = findMentionedProduct(q);
    if (prodName) {
      const prodTx = transactions.filter(t => t.product === prodName);
      const byYear = {};
      let totalSales = 0, totalQty = 0;
      const byCompany = {};
      prodTx.forEach(t => {
        totalSales += t.sales;
        totalQty += t.qty;
        byYear[t.year] = byYear[t.year] || { qty: 0, sales: 0 };
        byYear[t.year].qty += t.qty;
        byYear[t.year].sales += t.sales;
        byCompany[t.company] = (byCompany[t.company] || 0) + t.sales;
      });
      const historyText = Object.keys(byYear).sort().map(yr =>
        `- **Year ${yr}:** Qty ${byYear[yr].qty.toLocaleString()}, Sales ${byYear[yr].sales.toLocaleString()} AED`
      ).join('\n');
      const topBuyer = Object.entries(byCompany).sort((a, b) => b[1] - a[1])[0];
      return `📦 **Sales report for ${prodName}:**\n- **Total Sales:** ${totalSales.toLocaleString()} AED\n- **Total Quantity:** ${totalQty.toLocaleString()} units\n- **Top Buyer:** ${topBuyer ? `${topBuyer[0]} (${topBuyer[1].toLocaleString()} AED)` : 'N/A'}\n\n**Annual Performance Breakdown:**\n${historyText || 'No sales records found'}`;
    }

    return `🤖 **Local Assistant Note:** I couldn't run a precise local query for your question. To get deep analytical insights, ask custom questions, or find patterns in your sales data, please enter a **Gemini API Key** in the settings panel above!\n\n*Suggestions for offline mode:*\n- "Show top products"\n- "What are the total sales?"\n- "How many companies are there?"\n- "Show sales in 2018"\n- "Compare 2019 and 2021"\n- "Sales trend"\n- Type "help" to see everything I can do`;
  };

  const findMentionedCompany = (queryText) => {
    const lowerQuery = queryText.toLowerCase();
    // Scan all companies
    for (const comp of globalStats.allCompanyNames) {
      if (lowerQuery.includes(comp.toLowerCase())) {
        return comp;
      }
    }
    return null;
  };

  // Find every company mentioned (used for "compare X and Y" queries).
  // Names under 5 chars are skipped to avoid matching common words.
  const findAllMentionedCompanies = (queryText) => {
    const lowerQuery = queryText.toLowerCase();
    return globalStats.allCompanyNames.filter(
      comp => comp.length >= 5 && lowerQuery.includes(comp.toLowerCase())
    );
  };

  const findMentionedProduct = (queryText) => {
    const lowerQuery = queryText.toLowerCase();
    for (const prod of globalStats.allProductNames) {
      if (prod.length >= 5 && lowerQuery.includes(prod.toLowerCase())) {
        return prod;
      }
    }
    return null;
  };

  const findAllMentionedProducts = (queryText) => {
    const lowerQuery = queryText.toLowerCase();
    return globalStats.allProductNames.filter(
      prod => prod.length >= 5 && lowerQuery.includes(prod.toLowerCase())
    );
  };

  // Run Gemini API Assistant
  const handleSendMessage = async (e) => {
    e.preventDefault();
    if (!chatInput.trim()) return;

    const userText = chatInput.trim();
    setChatInput('');

    // Append user message
    const userMsg = {
      sender: 'user',
      text: userText,
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    };
    setMessages(prev => [...prev, userMsg]);
    setIsAiTyping(true);

    const activeKey = aiProvider === 'claude' ? claudeApiKey : apiKey;

    try {
      if (activeKey) {
        // Build RAG context
        // Scan for mentioned companies/products
        const mentionedCompany = findMentionedCompany(userText);
        const mentionedProduct = findMentionedProduct(userText);
        let customContext = '';

        if (mentionedCompany) {
          // Get transactions of this company
          const compTx = transactions
            .filter(t => t.company.toLowerCase() === mentionedCompany.toLowerCase())
            .slice(0, 150); // limit to 150 records to fit comfortably

          customContext += `
Specific query context for mentioned company "${mentionedCompany}":
The company has ${compTx.length} detailed transaction points in this file. Here they are:
${JSON.stringify(compTx.map(t => ({ product: t.product, year: t.year, qty: t.qty, sales: t.sales })))}
`;
        }

        if (mentionedProduct) {
          const prodTx = transactions
            .filter(t => t.product === mentionedProduct)
            .slice(0, 150);

          customContext += `
Specific query context for mentioned product "${mentionedProduct}":
The product has ${prodTx.length} detailed transaction points in this file. Here they are:
${JSON.stringify(prodTx.map(t => ({ company: t.company, year: t.year, qty: t.qty, sales: t.sales })))}
`;
        }

        const systemInstruction = `You are an expert sales analyst for CORNELL INTERNATIONAL DMCC.
You are helping the management analyze a sales dataset spanning 2009 to 2021.
The source file is 'Historical Report 2009-2021 (1).xls'.
Here is the high-level summary of the dataset:
- Total Sales Revenue: ${globalStats.totalSalesAED.toLocaleString()} AED
- Total Units Sold: ${globalStats.totalQty.toLocaleString()}
- Total Unique Customers/Companies: ${globalStats.companyCount}
- Total Unique Products: ${globalStats.productCount}

Sales by Year:
${JSON.stringify(globalStats.salesByYear)}

Units Sold by Year:
${JSON.stringify(globalStats.qtyByYear)}

Top 10 Products by Sales:
${JSON.stringify(globalStats.topProducts.slice(0, 10))}

Top 10 Companies by Sales:
${JSON.stringify(globalStats.companies.slice(0, 10).map(c => ({ name: c.name, totalSales: c.totalSales, totalQty: c.totalQty })))}

${customContext}

Answer the user's question accurately using the data above. Be direct, professional, and insightful. Format your response with markdown tables or bullets where helpful. If the question requires data not present above (e.g. a company/product not listed), say so rather than guessing.`;

        // Send API call to the selected provider
        const aiText = aiProvider === 'claude'
          ? await callClaudeAPI(claudeApiKey, userText, systemInstruction)
          : await callGeminiAPI(apiKey, userText, systemInstruction);


        setMessages(prev => [
          ...prev,
          {
            sender: 'ai',
            text: aiText,
            timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
          }
        ]);
      } else {
        // Simulate local response with delay
        setTimeout(() => {
          const reply = runLocalAIQuery(userText);
          setMessages(prev => [
            ...prev,
            {
              sender: 'ai',
              text: reply,
              timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
            }
          ]);
          setIsAiTyping(false);
        }, 600);
        return;
      }
    } catch (err) {
      console.error(err);
      setMessages(prev => [
        ...prev,
        {
          sender: 'ai',
          text: `❌ **Error calling ${aiProvider === 'claude' ? 'Claude' : 'Gemini'} API:** ${err.message || 'Check your internet connection and API Key.'}`,
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        }
      ]);
    } finally {
      setIsAiTyping(false);
    }
  };

  const callGeminiAPI = async (key, prompt, systemInstruction) => {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`;
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        contents: [
          {
            role: 'user',
            parts: [{ text: `${systemInstruction}\n\nUser Question: ${prompt}` }]
          }
        ]
      })
    });

    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.error?.message || 'Failed to call Gemini API');
    }

    const resData = await response.json();
    return resData.candidates[0].content.parts[0].text;
  };

  const callClaudeAPI = async (key, prompt, systemInstruction) => {
    const anthropic = new Anthropic({ apiKey: key, dangerouslyAllowBrowser: true });

    const response = await anthropic.messages.create({
      model: 'claude-opus-5',
      max_tokens: 2048,
      system: systemInstruction,
      output_config: { effort: 'medium' },
      messages: [{ role: 'user', content: prompt }]
    });

    if (response.stop_reason === 'refusal') {
      throw new Error('Claude declined to answer this request.');
    }

    const textBlock = response.content.find(block => block.type === 'text');
    return textBlock ? textBlock.text : '';
  };

  // Convert markdown-like syntax to HTML strings safely
  const formatMessageText = (text) => {
    // Replace markdown bold **
    let formatted = text.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    // Replace markdown italic *
    formatted = formatted.replace(/\*(.*?)\*/g, '<em>$1</em>');
    // Replace bullet points
    formatted = formatted.replace(/^\s*-\s+(.*)$/gm, '• $1');
    // Replace newlines with breaks
    formatted = formatted.split('\n').join('<br />');
    return formatted;
  };

  if (!isAuthenticated) {
    return (
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: '100vw',
        height: '100vh',
        backgroundColor: 'var(--bg-primary)',
        fontFamily: 'var(--font-family)'
      }}>
        <form onSubmit={handleLogin} className="login-card">
          <div style={{ textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '10px' }}>
            <div className="logo-icon" style={{ width: '48px', height: '48px' }}>
              <FileSpreadsheet size={26} color="#fff" />
            </div>
            <h2 style={{ fontSize: '1.5rem', fontWeight: '700', letterSpacing: '-0.025em', marginTop: '10px' }}>
              CORNELL SALES ANALYZER
            </h2>
            <p style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
              Sign in to access corporate data dashboard
            </p>
          </div>

          {loginError && (
            <div style={{
              backgroundColor: 'rgba(239, 68, 68, 0.1)',
              border: '1px solid rgba(239, 68, 68, 0.2)',
              color: '#f87171',
              padding: '12px',
              borderRadius: 'var(--border-radius-sm)',
              fontSize: '0.85rem',
              textAlign: 'center'
            }}>
              {loginError}
            </div>
          )}

          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <label style={{ fontSize: '0.8rem', fontWeight: '600', color: 'var(--text-secondary)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              Username
            </label>
            <input 
              type="text" 
              placeholder="Enter username" 
              className="input-glow"
              value={usernameInput}
              onChange={(e) => setUsernameInput(e.target.value)}
              required
            />
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <label style={{ fontSize: '0.8rem', fontWeight: '600', color: 'var(--text-secondary)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              Password
            </label>
            <input 
              type="password" 
              placeholder="Enter password" 
              className="input-glow"
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              required
            />
          </div>

          <button type="submit" className="btn-primary" style={{ padding: '14px', fontSize: '0.95rem', fontWeight: '700', marginTop: '8px' }}>
            Sign In
          </button>
        </form>
      </div>
    );
  }

  return (
    <div className="app-container">
      {/* Sidebar: Chat, API Settings, Upload info */}
      <aside className={`sidebar ${isSidebarOpen ? '' : 'collapsed'}`}>
        <div className="sidebar-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div className="logo-icon">
              <FileSpreadsheet size={22} color="#fff" />
            </div>
            <div className="logo-text">
              <h1>CORNELL ANALYZER</h1>
              <p>Sales Intelligence Hub</p>
            </div>
          </div>
          <button 
            onClick={() => setIsSidebarOpen(false)}
            style={{ background: 'none', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', display: 'flex', alignItems: 'center', padding: '4px' }}
            title="Collapse Sidebar"
          >
            <X size={18} />
          </button>
        </div>

        <div className="sidebar-sections">
          {/* File Card (Active File) */}
          <div className="panel">
            <div className="panel-title">
              <Database size={14} color="#6366f1" /> Active Dataset
            </div>
            <div className="panel-content">
              <div className="file-info-card">
                <div className="file-icon-box">
                  <FileSpreadsheet size={18} />
                </div>
                <div className="file-details">
                  <div className="file-name" title="Historical Report 2009-2021 (1).xls">
                    Historical Report 2009-2021 (1).xls
                  </div>
                  <div className="file-size">8.25 MB • 49,442 Rows loaded</div>
                </div>
              </div>
              <p className="help-text">
                Dataset parsed and cached locally. Direct access is enabled for tabular, dashboard views, and the chatbot.
              </p>
            </div>
          </div>

          {/* AI Settings */}
          <div className="panel">
            <div className="panel-title" style={{ justifyContent: 'space-between', display: 'flex' }}>
              <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Key size={14} color="#a855f7" /> Built-in AI Settings
              </span>
              <button 
                onClick={() => setShowApiSettings(!showApiSettings)} 
                className="btn-secondary"
                style={{ padding: '3px 8px', fontSize: '0.75rem', borderRadius: '4px' }}
              >
                {showApiSettings ? 'Close' : 'Setup'}
              </button>
            </div>
            {showApiSettings ? (
              <form onSubmit={handleSaveApiKey} className="panel-content" style={{ marginTop: '10px' }}>
                <div style={{ display: 'flex', gap: '6px', marginBottom: '4px' }}>
                  <button
                    type="button"
                    onClick={() => setAiProvider('gemini')}
                    className={aiProvider === 'gemini' ? 'btn-primary' : 'btn-secondary'}
                    style={{ flex: 1, padding: '8px', fontSize: '0.8rem' }}
                  >
                    Gemini
                  </button>
                  <button
                    type="button"
                    onClick={() => setAiProvider('claude')}
                    className={aiProvider === 'claude' ? 'btn-primary' : 'btn-secondary'}
                    style={{ flex: 1, padding: '8px', fontSize: '0.8rem' }}
                  >
                    Claude
                  </button>
                </div>
                <label className="help-text" style={{ fontSize: '0.75rem' }}>
                  {aiProvider === 'claude' ? 'Claude API Key:' : 'Gemini API Key:'}
                </label>
                <div className="api-input-group">
                  {aiProvider === 'claude' ? (
                    <input
                      type="password"
                      placeholder="sk-ant-..."
                      className="input-glow"
                      value={claudeApiKey}
                      onChange={(e) => setClaudeApiKey(e.target.value)}
                    />
                  ) : (
                    <input
                      type="password"
                      placeholder="AIzaSy..."
                      className="input-glow"
                      value={apiKey}
                      onChange={(e) => setApiKey(e.target.value)}
                    />
                  )}
                  <button type="submit" className="btn-primary" style={{ padding: '8px 12px' }}>
                    Save
                  </button>
                </div>
                <span className="help-text">
                  Your key is stored only in this browser's secure `localStorage` and called directly.
                </span>
              </form>
            ) : (
              <div style={{ fontSize: '0.75rem', color: (aiProvider === 'claude' ? claudeApiKey : apiKey) ? '#10b981' : '#71717a', display: 'flex', alignItems: 'center', gap: '6px', marginTop: '6px' }}>
                <div style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: (aiProvider === 'claude' ? claudeApiKey : apiKey) ? '#10b981' : '#71717a' }}></div>
                {(aiProvider === 'claude' ? claudeApiKey : apiKey)
                  ? `${aiProvider === 'claude' ? 'Claude' : 'Gemini'} Advanced AI Mode Active`
                  : 'Offline Local Mode Active (limited questions)'}
              </div>
            )}
          </div>

          {/* AI Insights Chat */}
          <div className="chat-container">
            <div className="panel-title" style={{ padding: '12px', borderBottom: '1px solid var(--border-light)', margin: 0, backgroundColor: 'rgba(39, 39, 42, 0.2)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ display: 'flex', alignItems: 'center' }}>
                <Bot size={16} color="#6366f1" style={{ marginRight: '6px' }} /> AI Business Analyst
              </span>
              <button
                onClick={handleClearChat}
                className="btn-secondary"
                style={{ padding: '3px 8px', fontSize: '0.7rem', borderRadius: '4px', textTransform: 'none' }}
                title="Clear chat history"
              >
                Clear
              </button>
            </div>
            
            <div className="chat-messages">
              {messages.map((msg, i) => (
                <div key={i} className={`chat-message ${msg.sender}`}>
                  <div 
                    className="chat-message-content"
                    dangerouslySetInnerHTML={{ __html: formatMessageText(msg.text) }}
                  />
                  <span className="chat-message-time">{msg.timestamp}</span>
                </div>
              ))}
              {isAiTyping && (
                <div className="chat-message ai">
                  <div className="typing-indicator">
                    <div className="typing-dot"></div>
                    <div className="typing-dot"></div>
                    <div className="typing-dot"></div>
                  </div>
                </div>
              )}
              <div ref={chatEndRef} />
            </div>

            {/* Quick Suggestions */}
            <div style={{ padding: '0 12px 8px 12px', backgroundColor: 'rgba(24, 24, 27, 0.6)' }}>
              <div className="ai-prompts-suggestions">
                <button 
                  onClick={() => { setChatInput('Which company has the highest sales?'); }}
                  className="suggestion-pill"
                >
                  Top Customer
                </button>
                <button 
                  onClick={() => { setChatInput('What is the total revenue and units sold?'); }}
                  className="suggestion-pill"
                >
                  Key Stats
                </button>
                <button 
                  onClick={() => { setChatInput('Show the best product'); }}
                  className="suggestion-pill"
                >
                  Best Product
                </button>
              </div>
            </div>

            <form onSubmit={handleSendMessage} className="chat-input-area">
              <input 
                type="text" 
                placeholder="Ask about companies, sales, trends..." 
                className="input-glow"
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
              />
              <button type="submit" className="btn-primary" style={{ padding: '10px' }}>
                <Send size={16} />
              </button>
            </form>
          </div>
        </div>
      </aside>

      {/* Main Panel */}
      <main className="main-content">
        <header className="main-header">
          <div className="header-top" style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
            {!isSidebarOpen && (
              <button 
                onClick={() => setIsSidebarOpen(true)} 
                className="btn-secondary"
                style={{ padding: '8px', borderRadius: '8px', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                title="Expand Sidebar"
              >
                <Menu size={20} />
              </button>
            )}
            <div className="header-title" style={{ flex: 1 }}>
              <h2>Historical Report (2009 - 2021)</h2>
              <p>Corporate Sales Summary & Data Intelligence Dashboard</p>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
              <span className="badge-year" style={{ padding: '8px 14px', fontSize: '0.85rem' }}>
                Cornell International DMCC
              </span>
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '4px' }}>
                <button 
                  onClick={handleLogout}
                  className="btn-secondary"
                  style={{ padding: '8px 14px', fontSize: '0.85rem', cursor: 'pointer' }}
                >
                  Logout (Captain)
                </button>
                <span style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>v 1.1</span>
              </div>
            </div>
          </div>

          <nav className="tabs">
            <button 
              className={`tab-btn ${activeTab === 'overview' ? 'active' : ''}`}
              onClick={() => setActiveTab('overview')}
            >
              Overview Dashboard
            </button>
            <button 
              className={`tab-btn ${activeTab === 'table' ? 'active' : ''}`}
              onClick={() => setActiveTab('table')}
            >
              Data Table View
            </button>
            <button 
              className={`tab-btn ${activeTab === 'charts' ? 'active' : ''}`}
              onClick={() => isChartBuilderEnabled && setActiveTab('charts')}
              disabled={!isChartBuilderEnabled}
              style={{
                opacity: isChartBuilderEnabled ? 1 : 0.4,
                cursor: isChartBuilderEnabled ? 'pointer' : 'not-allowed'
              }}
              title={isChartBuilderEnabled ? 'Dynamic Chart Builder' : 'Click "Show on Dynamic Chart" in Table View to enable'}
            >
              Dynamic Chart Builder
            </button>
          </nav>
        </header>

        {/* Viewport content */}
        <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
          
          {/* OVERVIEW DASHBOARD */}
          {activeTab === 'overview' && (
            <div className="tab-panel">
              {/* Year Range Filter */}
              <div className="table-controls" style={{ flexWrap: 'wrap' }}>
                <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', fontWeight: 600 }}>
                  Year Range:
                </span>
                <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
                  <select
                    className="select-custom"
                    value={overviewFromYear}
                    onChange={(e) => setOverviewFromYear(Number(e.target.value))}
                  >
                    {yearsList.map(yr => (
                      <option key={yr} value={yr} disabled={yr > overviewToYear}>{yr}</option>
                    ))}
                  </select>
                  <span style={{ color: 'var(--text-muted)' }}>to</span>
                  <select
                    className="select-custom"
                    value={overviewToYear}
                    onChange={(e) => setOverviewToYear(Number(e.target.value))}
                  >
                    {yearsList.map(yr => (
                      <option key={yr} value={yr} disabled={yr < overviewFromYear}>{yr}</option>
                    ))}
                  </select>
                  <button
                    onClick={() => { setOverviewFromYear(yearsList[0]); setOverviewToYear(yearsList[yearsList.length - 1]); }}
                    className="btn-secondary"
                    style={{ padding: '10px 16px', fontSize: '0.85rem' }}
                  >
                    All Time
                  </button>
                  <button
                    onClick={() => { setOverviewFromYear(Math.max(yearsList[0], yearsList[yearsList.length - 1] - 2)); setOverviewToYear(yearsList[yearsList.length - 1]); }}
                    className="btn-secondary"
                    style={{ padding: '10px 16px', fontSize: '0.85rem' }}
                  >
                    Last 3 Years
                  </button>
                </div>

                <div style={{ marginLeft: 'auto', display: 'flex', gap: '10px', alignItems: 'center' }}>
                  <input
                    ref={importFileInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    style={{ display: 'none' }}
                    onChange={handleImportFile}
                  />
                  <button
                    onClick={() => importFileInputRef.current && importFileInputRef.current.click()}
                    className="btn-secondary"
                    style={{ padding: '10px 16px', fontSize: '0.85rem', gap: '8px', display: 'flex', alignItems: 'center' }}
                    title="Import a new report — either the tidy CSV template or a file shaped like the original spreadsheet"
                  >
                    <FileSpreadsheet size={16} color="#6366f1" /> Import Excel / CSV
                  </button>
                  <button
                    onClick={handleReloadOriginalData}
                    className="btn-secondary"
                    style={{ padding: '10px 16px', fontSize: '0.85rem', gap: '8px', display: 'flex', alignItems: 'center' }}
                    title="Discard any imported data and restore the original bundled report"
                  >
                    <RefreshCw size={16} /> Reload Original Data
                  </button>
                </div>
              </div>

              {importStatus && (
                <div style={{
                  fontSize: '0.85rem',
                  padding: '10px 16px',
                  borderRadius: 'var(--border-radius-sm)',
                  backgroundColor: importStatus.type === 'error' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(16, 185, 129, 0.1)',
                  color: importStatus.type === 'error' ? '#f87171' : '#10b981',
                  border: `1px solid ${importStatus.type === 'error' ? 'rgba(239, 68, 68, 0.2)' : 'rgba(16, 185, 129, 0.2)'}`
                }}>
                  {importStatus.text.replace(/\*\*/g, '')}
                </div>
              )}

              {/* Stats Cards */}
              <div className="stats-grid">
                <div className="stats-card">
                  <div className="stats-icon-box">
                    <TrendingUp size={24} />
                  </div>
                  <div className="stats-info">
                    <span className="stats-label">Total Revenue</span>
                    <span className="stats-value">{formatCurrency(overviewStats.totalSalesAED)}</span>
                  </div>
                </div>

                <div className="stats-card">
                  <div className="stats-icon-box cyan">
                    <Layers size={24} />
                  </div>
                  <div className="stats-info">
                    <span className="stats-label">Units Sold</span>
                    <span className="stats-value">{formatNumber(overviewStats.totalQty)}</span>
                  </div>
                </div>

                <div className="stats-card">
                  <div className="stats-icon-box purple">
                    <Users size={24} />
                  </div>
                  <div className="stats-info">
                    <span className="stats-label">Active Customers</span>
                    <span className="stats-value">{overviewStats.companyCount}</span>
                  </div>
                </div>

                <div className="stats-card">
                  <div className="stats-icon-box emerald">
                    <FileSpreadsheet size={24} />
                  </div>
                  <div className="stats-info">
                    <span className="stats-label">Total Products</span>
                    <span className="stats-value">{overviewStats.productCount}</span>
                  </div>
                </div>
              </div>

              {/* Main Charts Row */}
              <div className="dashboard-row">
                <div className="visual-card">
                  <div className="visual-card-header">
                    <span className="visual-card-title">
                      <Sparkles size={16} color="#6366f1" style={{ marginRight: '8px' }} />
                      Revenue & Quantity Sales Trends ({overviewFromYear} - {overviewToYear})
                    </span>
                    <div style={{ display: 'flex', gap: '4px' }}>
                      <button
                        onClick={() => setOverviewTrendChartType('line')}
                        className={overviewTrendChartType === 'line' ? 'btn-primary' : 'btn-secondary'}
                        style={{ padding: '6px 12px', fontSize: '0.75rem' }}
                      >
                        Line
                      </button>
                      <button
                        onClick={() => setOverviewTrendChartType('bar')}
                        className={overviewTrendChartType === 'bar' ? 'btn-primary' : 'btn-secondary'}
                        style={{ padding: '6px 12px', fontSize: '0.75rem' }}
                      >
                        Bar
                      </button>
                    </div>
                  </div>
                  <div style={{ flex: 1, position: 'relative', height: '320px' }}>
                    {overviewTrendChartType === 'line' ? (
                      <Line key={`overview-line-${chartLayoutKey}`} data={overviewChartData} options={overviewChartOptions} />
                    ) : (
                      <Bar key={`overview-bar-${chartLayoutKey}`} data={overviewChartData} options={overviewChartOptions} />
                    )}
                  </div>
                </div>

                <div className="visual-card">
                  <div className="visual-card-header">
                    <span className="visual-card-title">Top 7 Products (AED)</span>
                  </div>
                  <div style={{ flex: 1, position: 'relative', height: '320px' }}>
                    <Bar
                      key={`top-products-${chartLayoutKey}`}
                      data={topProductsChartData}
                      options={{
                        responsive: true,
                        maintainAspectRatio: false,
                        indexAxis: 'y',
                        plugins: {
                          legend: { display: false },
                          tooltip: {
                            padding: 10,
                            backgroundColor: '#18181b',
                            titleFont: { family: 'Outfit', size: 12 },
                            bodyFont: { family: 'Outfit', size: 11 }
                          }
                        },
                        scales: {
                          x: {
                            grid: { color: 'rgba(63, 63, 70, 0.1)' },
                            ticks: { 
                              color: '#a1a1aa', 
                              font: { family: 'Outfit', size: 10 },
                              callback: (val) => val >= 1e6 ? `${(val / 1e6).toFixed(1)}M` : val.toLocaleString()
                            }
                          },
                          y: {
                            grid: { display: false },
                            ticks: { color: '#fafafa', font: { family: 'Outfit', size: 10 } }
                          }
                        }
                      }} 
                    />
                  </div>
                </div>
              </div>

              {/* Top Companies Summary Card */}
              <div className="visual-card" style={{ minHeight: '500px' }}>
                <div className="visual-card-header">
                  <span className="visual-card-title">Leading Corporate Accounts (Top 5 Sales)</span>
                </div>
                <div className="table-wrapper">
                  <table style={{ minWidth: '100%' }}>
                    <thead>
                      <tr>
                        <th style={{ cursor: 'default' }}>Rank</th>
                        <th style={{ cursor: 'default' }}>Company Name</th>
                        <th style={{ cursor: 'default' }}>Ref Code</th>
                        <th style={{ cursor: 'default', textTransform: 'none' }}>Total Volume (Qty)</th>
                        <th style={{ cursor: 'default', textTransform: 'none' }}>Total Revenue (AED)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {topCompanies.map((c, i) => (
                        <tr key={i}>
                          <td><strong>#{i + 1}</strong></td>
                          <td className="td-company">{c.name}</td>
                          <td>{c.ref || 'N/A'}</td>
                          <td>{c.totalQty.toLocaleString(undefined, {maximumFractionDigits: 0})}</td>
                          <td><span className="badge-year">{c.totalSales.toLocaleString()} AED</span></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* FILTERED DATA TABLE VIEW */}
          {activeTab === 'table' && (
            <div className="tab-panel">
              {/* Filter controls */}
              <div className="table-controls" style={{ display: 'flex', flexDirection: 'column', gap: '16px', alignItems: 'stretch' }}>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '16px', alignItems: 'center' }}>
                  <div className="search-box-wrapper" style={{ flex: 1, minWidth: '220px' }}>
                    <Search size={18} className="search-icon" />
                    <input 
                      type="text" 
                      placeholder="Search company, product, ref..." 
                      className="search-input"
                      value={searchTerm}
                      onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
                    />
                  </div>

                  <div className="filter-group" style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center' }}>
                    {/* From Year */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>From Year</span>
                      <select 
                        className="select-custom"
                        value={fromYear}
                        onChange={(e) => { setFromYear(Number(e.target.value)); setCurrentPage(1); }}
                      >
                        {yearsList.map(yr => (
                          <option key={yr} value={yr}>{yr}</option>
                        ))}
                      </select>
                    </div>

                    {/* To Year */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>To Year</span>
                      <select 
                        className="select-custom"
                        value={toYear}
                        onChange={(e) => { setToYear(Number(e.target.value)); setCurrentPage(1); }}
                      >
                        {yearsList.map(yr => (
                          <option key={yr} value={yr}>{yr}</option>
                        ))}
                      </select>
                    </div>

                    {/* Formulation */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>Formulation</span>
                      <select 
                        className="select-custom"
                        value={filterFormulation}
                        onChange={(e) => { setFilterFormulation(e.target.value); setCurrentPage(1); }}
                      >
                        {formulationOptions.map(f => (
                          <option key={f} value={f}>{f}</option>
                        ))}
                      </select>
                    </div>

                    {/* Product Multiselect Search */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', position: 'relative', minWidth: '220px' }}>
                      <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>Products (Multi-select)</span>
                      <div style={{ position: 'relative' }}>
                        <input 
                          type="text" 
                          placeholder="Search product to add..." 
                          className="input-glow"
                          style={{ padding: '10px 12px', fontSize: '0.9rem' }}
                          value={productSearchInput}
                          onChange={(e) => setProductSearchInput(e.target.value)}
                        />
                        {productSearchInput.trim() !== '' && (
                          <div style={{ 
                            position: 'absolute', 
                            top: '100%', 
                            left: 0, 
                            right: 0, 
                            backgroundColor: 'var(--bg-secondary)', 
                            border: '1px solid var(--border-color)', 
                            borderRadius: 'var(--border-radius-sm)', 
                            zIndex: 20, 
                            maxHeight: '180px', 
                            overflowY: 'auto',
                            boxShadow: '0 4px 12px rgba(0,0,0,0.5)'
                          }}>
                            {uniqueProductsList
                              .filter(p => p.toLowerCase().includes(productSearchInput.toLowerCase()) && !selectedProducts.includes(p))
                              .slice(0, 8)
                              .map((p, idx, arr) => (
                                <div 
                                  key={idx}
                                  onClick={() => {
                                    setSelectedProducts([...selectedProducts, p]);
                                    setProductSearchInput('');
                                  }}
                                  style={{ 
                                    padding: '8px 12px', 
                                    fontSize: '0.85rem', 
                                    cursor: 'pointer',
                                    borderBottom: idx < arr.length - 1 ? '1px solid var(--border-light)' : 'none',
                                    color: 'var(--text-primary)'
                                  }}
                                  onMouseEnter={(e) => e.target.style.backgroundColor = 'rgba(99, 102, 241, 0.15)'}
                                  onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                  {p}
                                </div>
                              ))
                            }
                            {uniqueProductsList.filter(p => p.toLowerCase().includes(productSearchInput.toLowerCase()) && !selectedProducts.includes(p)).length === 0 && (
                              <div style={{ padding: '8px 12px', fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                                No matching products found
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  <div style={{ marginLeft: 'auto', display: 'flex', gap: '10px', alignItems: 'center' }}>
                    <button 
                      onClick={handleExportExcel}
                      className="btn-secondary"
                      style={{ padding: '12px 18px', gap: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                    >
                      <FileSpreadsheet size={16} color="#10b981" /> Export Excel
                    </button>

                    <button 
                      onClick={() => {
                        setIsChartBuilderEnabled(true);
                        setActiveTab('charts');
                      }}
                      className="btn-primary"
                      style={{ padding: '12px 18px', gap: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                    >
                      <BarChart3 size={16} /> Show on Dynamic Chart
                    </button>
                  </div>
                </div>

                {/* Selected Products Chips */}
                {selectedProducts.length > 0 && (
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', padding: '8px 12px', backgroundColor: 'var(--bg-primary)', borderRadius: 'var(--border-radius-sm)', border: '1px solid var(--border-light)' }}>
                    <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', display: 'flex', alignItems: 'center', marginRight: '6px' }}>
                      Filtering Products ({selectedProducts.length}):
                    </span>
                    {selectedProducts.map((p, idx) => (
                      <span 
                        key={idx} 
                        className="badge-formulation" 
                        style={{ 
                          display: 'inline-flex', 
                          alignItems: 'center', 
                          gap: '6px', 
                          padding: '4px 10px', 
                          fontSize: '0.75rem',
                          borderRadius: '16px',
                          backgroundColor: 'rgba(168, 85, 247, 0.15)',
                          color: 'var(--accent-secondary)',
                          border: '1px solid rgba(168, 85, 247, 0.3)'
                        }}
                      >
                        {p}
                        <button 
                          onClick={() => setSelectedProducts(selectedProducts.filter(item => item !== p))}
                          style={{ 
                            background: 'none', 
                            border: 'none', 
                            color: 'var(--accent-secondary)', 
                            cursor: 'pointer', 
                            fontSize: '0.85rem',
                            fontWeight: 'bold',
                            display: 'inline-flex',
                            alignItems: 'center'
                          }}
                        >
                          ×
                        </button>
                      </span>
                    ))}
                    <button 
                      onClick={() => setSelectedProducts([])}
                      style={{ 
                        background: 'none', 
                        border: 'none', 
                        color: 'var(--text-muted)', 
                        cursor: 'pointer', 
                        fontSize: '0.75rem',
                        textDecoration: 'underline',
                        marginLeft: 'auto'
                      }}
                    >
                      Clear All
                    </button>
                  </div>
                )}
              </div>

              {/* Data Table */}
              <div className="table-card">
                <div className="table-wrapper">
                  <table>
                    <thead>
                      <tr>
                        <th onClick={() => handleSort('company')}>
                          Company Name {sortKey === 'company' && (sortDirection === 'asc' ? '▲' : '▼')}
                        </th>
                        <th onClick={() => handleSort('product')}>
                          Product {sortKey === 'product' && (sortDirection === 'asc' ? '▲' : '▼')}
                        </th>
                        <th onClick={() => handleSort('formulation')}>
                          Formulation {sortKey === 'formulation' && (sortDirection === 'asc' ? '▲' : '▼')}
                        </th>
                        <th onClick={() => handleSort('year')}>
                          Year {sortKey === 'year' && (sortDirection === 'asc' ? '▲' : '▼')}
                        </th>
                        <th onClick={() => handleSort('qty')}>
                          Quantity {sortKey === 'qty' && (sortDirection === 'asc' ? '▲' : '▼')}
                        </th>
                        <th onClick={() => handleSort('sales')}>
                          Sales (AED) {sortKey === 'sales' && (sortDirection === 'asc' ? '▲' : '▼')}
                        </th>
                        <th onClick={() => handleSort('salesWithVat')}>
                          Sales w/ VAT (AED) {sortKey === 'salesWithVat' && (sortDirection === 'asc' ? '▲' : '▼')}
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {paginatedTransactions.length > 0 ? (
                        paginatedTransactions.map((t, idx) => (
                          <tr key={idx}>
                            <td className="td-company" title={t.company}>{t.company}</td>
                            <td className="td-product" title={t.product}>{t.product}</td>
                            <td>
                              <span className={t.formulation === 'N/A' ? 'help-text' : 'badge-formulation'}>
                                {t.formulation}
                              </span>
                            </td>
                            <td><span className="badge-year">{t.year}</span></td>
                            <td>{t.qty.toLocaleString()}</td>
                            <td><strong>{t.sales.toLocaleString()}</strong></td>
                            <td>{t.salesWithVat.toLocaleString()}</td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan="7" style={{ textAlign: 'center', padding: '30px', color: 'var(--text-secondary)' }}>
                            No transaction points matched your filters.
                          </td>
                        </tr>
                      )}
                    </tbody>
                    {filteredAndSortedTransactions.length > 0 && (
                      <tfoot style={{ position: 'sticky', bottom: 0, zIndex: 5, backgroundColor: '#18181b' }}>
                        <tr style={{ borderTop: '2px solid var(--border-color)', fontWeight: 'bold' }}>
                          <td colSpan="4" style={{ padding: '14px 20px', color: 'var(--text-primary)' }}>Total (Filtered)</td>
                          <td style={{ padding: '14px 20px', color: 'var(--accent-cyan)' }}>{filteredTotals.qty.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
                          <td style={{ padding: '14px 20px', color: 'var(--accent-primary)' }}>{filteredTotals.sales.toLocaleString()} AED</td>
                          <td style={{ padding: '14px 20px', color: 'var(--accent-secondary)' }}>{filteredTotals.salesWithVat.toLocaleString()} AED</td>
                        </tr>
                      </tfoot>
                    )}
                  </table>
                </div>

                {/* Table Pagination */}
                {filteredAndSortedTransactions.length > 0 && (
                  <div className="pagination" style={{ display: 'flex', flexWrap: 'wrap', gap: '16px', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '16px', flexWrap: 'wrap' }}>
                      <span className="pagination-info">
                        Showing {(currentPage - 1) * pageSize + 1} - {Math.min(currentPage * pageSize, filteredAndSortedTransactions.length)} of {filteredAndSortedTransactions.length} records
                      </span>
                      
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>Rows per page:</span>
                        <select 
                          className="select-custom" 
                          style={{ padding: '4px 8px', fontSize: '0.8rem', borderRadius: 'var(--border-radius-sm)' }}
                          value={pageSize}
                          onChange={(e) => {
                            setPageSize(Number(e.target.value));
                            setCurrentPage(1);
                          }}
                        >
                          <option value={50}>50</option>
                          <option value={100}>100</option>
                          <option value={300}>300</option>
                          <option value={500}>500</option>
                          <option value={1000}>1000</option>
                        </select>
                      </div>
                    </div>
                    
                    {totalPages > 1 && (
                      <div className="pagination-controls" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <button 
                          className="pagination-btn"
                          onClick={() => setCurrentPage(1)}
                          disabled={currentPage === 1}
                          title="First Page"
                        >
                          <ChevronsLeft size={16} />
                        </button>

                        <button 
                          className="pagination-btn"
                          onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                          disabled={currentPage === 1}
                          title="Previous Page"
                        >
                          <ChevronLeft size={16} />
                        </button>
                        
                        <span style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '0.85rem', margin: '0 8px' }}>
                          Page <strong>{currentPage}</strong> of {totalPages}
                        </span>

                        <button 
                          className="pagination-btn"
                          onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                          disabled={currentPage === totalPages}
                          title="Next Page"
                        >
                          <ChevronRight size={16} />
                        </button>

                        <button 
                          className="pagination-btn"
                          onClick={() => setCurrentPage(totalPages)}
                          disabled={currentPage === totalPages}
                          title="Last Page"
                        >
                          <ChevronsRight size={16} />
                        </button>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* DYNAMIC CHART BUILDER */}
          {activeTab === 'charts' && (
            <div className="tab-panel">
              <div className="chart-builder-controls">
                <div className="control-item">
                  <label>Chart Type</label>
                  <select 
                    className="select-custom" 
                    value={chartType}
                    onChange={(e) => setChartType(e.target.value)}
                  >
                    <option value="bar">Bar Chart</option>
                    <option value="line">Line Chart</option>
                  </select>
                </div>

                <div className="control-item">
                  <label>X-Axis Dimension</label>
                  <select 
                    className="select-custom" 
                    value={chartXAxis}
                    onChange={(e) => setChartXAxis(e.target.value)}
                  >
                    <option value="year">Sales Year</option>
                    <option value="company">Company Account</option>
                    <option value="product">Product Name</option>
                    <option value="formulation">Formulation type</option>
                  </select>
                </div>

                <div className="control-item">
                  <label>Y-Axis Metric</label>
                  <select 
                    className="select-custom" 
                    value={chartYAxis}
                    onChange={(e) => setChartYAxis(e.target.value)}
                  >
                    <option value="sales">Sales Revenue (AED)</option>
                    <option value="qty">Quantity Sold</option>
                  </select>
                </div>

                <div className="control-item">
                  <label>Limit Dimension Items</label>
                  <select 
                    className="select-custom" 
                    value={chartLimit}
                    onChange={(e) => setChartLimit(e.target.value)}
                  >
                    <option value="5">Top 5 Items</option>
                    <option value="10">Top 10 Items</option>
                    <option value="25">Top 25 Items</option>
                    <option value="50">Top 50 Items</option>
                    <option value="all">All Items</option>
                  </select>
                </div>
              </div>

              {/* Chart Panel Output */}
              <div className="chart-container-wrapper">
                <div style={{ position: 'relative', width: '95%', height: '420px' }}>
                  {chartType === 'line' ? (
                    <Line
                      key={`builder-line-${chartLayoutKey}`}
                      data={builderChartData}
                      options={{
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                          legend: { display: false }
                        },
                        scales: {
                          x: {
                            grid: { color: 'rgba(63, 63, 70, 0.15)' },
                            ticks: { color: '#fafafa', font: { family: 'Outfit', size: 10 } }
                          },
                          y: {
                            grid: { color: 'rgba(63, 63, 70, 0.15)' },
                            ticks: {
                              color: '#a1a1aa',
                              font: { family: 'Outfit', size: 10 },
                              callback: (val) => val >= 1e6 ? `${(val / 1e6).toFixed(1)}M` : val.toLocaleString()
                            }
                          }
                        }
                      }}
                    />
                  ) : (
                    <Bar
                      key={`builder-bar-${chartLayoutKey}`}
                      data={builderChartData}
                      options={{
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                          legend: { display: false }
                        },
                        scales: {
                          x: {
                            grid: { display: false },
                            ticks: { color: '#fafafa', font: { family: 'Outfit', size: 10 } }
                          },
                          y: {
                            grid: { color: 'rgba(63, 63, 70, 0.15)' },
                            ticks: { 
                              color: '#a1a1aa',
                              font: { family: 'Outfit', size: 10 },
                              callback: (val) => val >= 1e6 ? `${(val / 1e6).toFixed(1)}M` : val.toLocaleString()
                            }
                          }
                        }
                      }} 
                    />
                  )}
                </div>
              </div>
            </div>
          )}

        </div>
      </main>
    </div>
  );
}

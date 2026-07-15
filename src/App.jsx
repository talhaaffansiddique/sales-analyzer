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

export default function App() {
  // Authentication states
  const [isAuthenticated, setIsAuthenticated] = useState(sessionStorage.getItem('cornell_authenticated') === 'true');
  const [usernameInput, setUsernameInput] = useState('');
  const [passwordInput, setPasswordInput] = useState('');
  const [loginError, setLoginError] = useState('');

  // App states
  const [activeTab, setActiveTab] = useState('overview');
  const [transactions, setTransactions] = useState(compiledData.transactions);
  const [stats, setStats] = useState(compiledData.stats);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isChartBuilderEnabled, setIsChartBuilderEnabled] = useState(false);
  
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
  const [messages, setMessages] = useState([
    {
      sender: 'ai',
      text: "Hello! I am your AI Data Assistant. I have analyzed **Historical Report 2009-2021 (1).xls**.\n\nYou can ask me questions about this dataset, such as which companies have the highest sales, sales trends by year, or product performance. Enter a Gemini API Key in the settings above to unlock advanced natural language questions!",
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    }
  ]);
  const [chatInput, setChatInput] = useState('');
  const [apiKey, setApiKey] = useState(localStorage.getItem('gemini_api_key') || '');
  const [showApiSettings, setShowApiSettings] = useState(false);
  const [isAiTyping, setIsAiTyping] = useState(false);
  const chatEndRef = useRef(null);

  // Scroll to bottom of chat
  useEffect(() => {
    if (chatEndRef.current) {
      chatEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [messages, isAiTyping]);

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

  // Save API Key
  const handleSaveApiKey = (e) => {
    e.preventDefault();
    localStorage.setItem('gemini_api_key', apiKey);
    setShowApiSettings(false);
    setMessages(prev => [
      ...prev,
      {
        sender: 'ai',
        text: apiKey 
          ? "✅ **Gemini API Key saved successfully!** You can now ask complex questions and request deep business analysis on your dataset."
          : "⚠️ **Gemini API Key removed.** Chat will now run in offline local query mode.",
        timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
      }
    ]);
  };

  // Get filter list options
  const formulationOptions = useMemo(() => {
    const formulations = new Set();
    compiledData.transactions.forEach(t => {
      if (t.formulation && t.formulation !== 'N/A') {
        formulations.add(t.formulation);
      }
    });
    return ['All', 'N/A', ...Array.from(formulations).sort()];
  }, []);

  const uniqueProductsList = useMemo(() => {
    const productsSet = new Set();
    compiledData.transactions.forEach(t => {
      productsSet.add(t.product);
    });
    return Array.from(productsSet).sort();
  }, []);

  const yearsList = [2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021];

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

  // Top Companies list for Overview Tab
  const topCompanies = useMemo(() => {
    return compiledData.companies.slice(0, 5);
  }, []);

  // Overview Main Chart Data (Sales & Qty over years)
  const overviewChartData = useMemo(() => {
    const sortedYears = [...yearsList].sort();
    const salesData = sortedYears.map(yr => stats.salesByYear[yr] || 0);
    const qtyData = sortedYears.map(yr => stats.qtyByYear[yr] || 0);

    return {
      labels: sortedYears,
      datasets: [
        {
          label: 'Total Revenue (AED)',
          data: salesData,
          borderColor: '#6366f1', // Indigo
          backgroundColor: 'rgba(99, 102, 241, 0.2)',
          borderWidth: 3,
          tension: 0.3,
          fill: true,
          yAxisID: 'y'
        },
        {
          label: 'Units Sold (Qty)',
          data: qtyData,
          borderColor: '#06b6d4', // Cyan
          backgroundColor: 'rgba(6, 182, 212, 0.1)',
          borderWidth: 2,
          tension: 0.3,
          borderDash: [5, 5],
          yAxisID: 'y1'
        }
      ]
    };
  }, [stats]);

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
    const list = stats.topProducts.slice(0, 7);
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
  }, [stats]);

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
    
    // Check for "how many companies" or similar
    if (q.includes('how many companies') || q.includes('how many customers') || q.includes('total companies')) {
      return `📊 There are **582 unique companies** present in this sales report. The top customer by sales is **${compiledData.companies[0].name}**, with a total sales volume of **${compiledData.companies[0].totalSales.toLocaleString()} AED** over the active years.`;
    }
    
    // Check for "how many products" or similar
    if (q.includes('how many products') || q.includes('total products')) {
      return `📦 There are **5,687 unique products** listed in this report. The highest-selling product by revenue is **${stats.topProducts[0].product}**, generating **${stats.topProducts[0].sales.toLocaleString()} AED** in revenue.`;
    }
    
    // Check for "total sales" or "total revenue"
    if (q.includes('total sales') || q.includes('total revenue') || q.includes('sales volume')) {
      return `💰 The total sales revenue aggregated from **2009 to 2021** is **${stats.totalSalesAED.toLocaleString(undefined, {maximumFractionDigits:2})} AED**, representing a total quantity of **${stats.totalQty.toLocaleString()} units** sold.`;
    }

    // Check for top products
    if (q.includes('top product') || q.includes('highest selling product') || q.includes('best product')) {
      const topList = stats.topProducts.slice(0, 5).map((p, idx) => `${idx + 1}. **${p.product}** - Sales: *${p.sales.toLocaleString()} AED* (Qty: ${p.qty.toLocaleString()})`).join('\n');
      return `🏆 Here are the **Top 5 Products** by sales revenue:\n\n${topList}`;
    }

    // Check for top companies
    if (q.includes('top company') || q.includes('highest sales company') || q.includes('best customer')) {
      const topList = compiledData.companies.slice(0, 5).map((c, idx) => `${idx + 1}. **${c.name}** - Total Sales: *${c.totalSales.toLocaleString()} AED* (Ref: ${c.ref || 'N/A'})`).join('\n');
      return `🏢 Here are the **Top 5 Companies** by sales revenue:\n\n${topList}`;
    }

    // Check for a specific year
    for (const yr of yearsList) {
      if (q.includes(String(yr))) {
        const sales = stats.salesByYear[yr] || 0;
        const qty = stats.qtyByYear[yr] || 0;
        return `📅 **Sales Summary for Year ${yr}:**\n- **Total Revenue:** ${sales.toLocaleString(undefined, {maximumFractionDigits: 2})} AED\n- **Volume Sold:** ${qty.toLocaleString()} units\n- **Contribution:** ${((sales / stats.totalSalesAED) * 100).toFixed(2)}% of historical total sales.`;
      }
    }

    // Check if a specific company is mentioned in the query
    const compName = findMentionedCompany(q);
    if (compName) {
      // Aggregate data for this company
      const compData = compiledData.companies.find(c => c.name === compName);
      if (compData) {
        const historyText = Object.keys(compData.totals).map(yr => {
          const t = compData.totals[yr];
          return `- **Year ${yr}:** Qty ${t.qty.toLocaleString()}, Sales ${t.val.toLocaleString()} AED`;
        }).join('\n');
        
        return `🏢 **Sales report for ${compData.name}** (Ref: ${compData.ref || 'N/A'}):\n- **Total Sales:** ${compData.totalSales.toLocaleString()} AED\n- **Total Quantity:** ${compData.totalQty.toLocaleString()} units\n\n**Annual Performance Breakdown:**\n${historyText || 'No sales records found'}`;
      }
    }

    return `🤖 **Local Assistant Note:** I couldn't run a precise local query for your question. To get deep analytical insights, ask custom questions, or find patterns in your sales data, please enter a **Gemini API Key** in the settings panel above!\n\n*Suggestions for offline mode:*\n- "Show top products"\n- "What are the total sales?"\n- "How many companies are there?"\n- "Show sales in 2018"`;
  };

  const findMentionedCompany = (queryText) => {
    const lowerQuery = queryText.toLowerCase();
    // Scan all companies
    for (const comp of compiledData.allCompanyNames) {
      if (lowerQuery.includes(comp.toLowerCase())) {
        return comp;
      }
    }
    return null;
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

    try {
      if (apiKey) {
        // Build RAG context
        // Scan for mentioned companies/products
        const mentionedCompany = findMentionedCompany(userText);
        let customContext = '';

        if (mentionedCompany) {
          // Get transactions of this company
          const compTx = compiledData.transactions
            .filter(t => t.company.toLowerCase() === mentionedCompany.toLowerCase())
            .slice(0, 150); // limit to 150 records to fit comfortably

          customContext = `
Specific query context for mentioned company "${mentionedCompany}":
The company has ${compTx.length} detailed transaction points in this file. Here they are:
${JSON.stringify(compTx.map(t => ({ product: t.product, year: t.year, qty: t.qty, sales: t.sales })))}
`;
        }

        const systemInstruction = `You are an expert sales analyst for CORNELL INTERNATIONAL DMCC.
You are helping the management analyze a sales dataset spanning 2009 to 2021.
The source file is 'Historical Report 2009-2021 (1).xls'.
Here is the high-level summary of the dataset:
- Total Sales Revenue: ${stats.totalSalesAED.toLocaleString()} AED
- Total Units Sold: ${stats.totalQty.toLocaleString()}
- Total Unique Customers/Companies: ${stats.companyCount}
- Total Unique Products: ${stats.productCount}

Sales by Year:
${JSON.stringify(stats.salesByYear)}

Top 10 Products by Sales:
${JSON.stringify(stats.topProducts.slice(0, 10))}

${customContext}

Answer the user's question accurately using the data above. Be direct, professional, and insightful. Format your response with markdown tables or bullets where helpful.`;

        // Send API call to Gemini
        const aiText = await callGeminiAPI(apiKey, userText, systemInstruction);
        
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
          text: `❌ **Error calling Gemini API:** ${err.message || 'Check your internet connection and API Key.'}`,
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
        <form onSubmit={handleLogin} style={{
          backgroundColor: 'var(--bg-secondary)',
          border: '1px solid var(--border-light)',
          borderRadius: 'var(--border-radius-lg)',
          padding: '40px',
          width: '420px',
          boxShadow: '0 10px 40px rgba(0, 0, 0, 0.5)',
          display: 'flex',
          flexDirection: 'column',
          gap: '24px',
          backdropFilter: 'blur(10px)'
        }}>
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
                <label className="help-text" style={{ fontSize: '0.75rem' }}>Gemini API Key:</label>
                <div className="api-input-group">
                  <input 
                    type="password" 
                    placeholder="AIzaSy..." 
                    className="input-glow"
                    value={apiKey}
                    onChange={(e) => setApiKey(e.target.value)}
                  />
                  <button type="submit" className="btn-primary" style={{ padding: '8px 12px' }}>
                    Save
                  </button>
                </div>
                <span className="help-text">
                  Your key is stored only in this browser's secure `localStorage` and called directly.
                </span>
              </form>
            ) : (
              <div style={{ fontSize: '0.75rem', color: apiKey ? '#10b981' : '#71717a', display: 'flex', alignItems: 'center', gap: '6px', marginTop: '6px' }}>
                <div style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: apiKey ? '#10b981' : '#71717a' }}></div>
                {apiKey ? 'Gemini Advanced AI Mode Active' : 'Offline Local Mode Active (limited questions)'}
              </div>
            )}
          </div>

          {/* AI Insights Chat */}
          <div className="chat-container">
            <div className="panel-title" style={{ padding: '12px', borderBottom: '1px solid var(--border-light)', margin: 0, backgroundColor: 'rgba(39, 39, 42, 0.2)' }}>
              <Bot size={16} color="#6366f1" style={{ marginRight: '6px' }} /> AI Business Analyst
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
              {/* Stats Cards */}
              <div className="stats-grid">
                <div className="stats-card">
                  <div className="stats-icon-box">
                    <TrendingUp size={24} />
                  </div>
                  <div className="stats-info">
                    <span className="stats-label">Total Revenue</span>
                    <span className="stats-value">{formatCurrency(stats.totalSalesAED)}</span>
                  </div>
                </div>

                <div className="stats-card">
                  <div className="stats-icon-box cyan">
                    <Layers size={24} />
                  </div>
                  <div className="stats-info">
                    <span className="stats-label">Units Sold</span>
                    <span className="stats-value">{formatNumber(stats.totalQty)}</span>
                  </div>
                </div>

                <div className="stats-card">
                  <div className="stats-icon-box purple">
                    <Users size={24} />
                  </div>
                  <div className="stats-info">
                    <span className="stats-label">Active Customers</span>
                    <span className="stats-value">{stats.companyCount}</span>
                  </div>
                </div>

                <div className="stats-card">
                  <div className="stats-icon-box emerald">
                    <FileSpreadsheet size={24} />
                  </div>
                  <div className="stats-info">
                    <span className="stats-label">Total Products</span>
                    <span className="stats-value">{stats.productCount}</span>
                  </div>
                </div>
              </div>

              {/* Main Charts Row */}
              <div className="dashboard-row">
                <div className="visual-card">
                  <div className="visual-card-header">
                    <span className="visual-card-title">
                      <Sparkles size={16} color="#6366f1" style={{ marginRight: '8px' }} />
                      Revenue & Quantity Sales Trends (2009 - 2021)
                    </span>
                  </div>
                  <div style={{ flex: 1, position: 'relative', height: '320px' }}>
                    <Line data={overviewChartData} options={overviewChartOptions} />
                  </div>
                </div>

                <div className="visual-card">
                  <div className="visual-card-header">
                    <span className="visual-card-title">Top 7 Products (AED)</span>
                  </div>
                  <div style={{ flex: 1, position: 'relative', height: '320px' }}>
                    <Bar 
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
              <div className="visual-card" style={{ minHeight: '200px' }}>
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

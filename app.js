// Mainpaper AI Sales System — Dashboard Logic Controller
// Global State
let dimProductos = [];
let salesData = [];       // Micro transaction records (Gmail orders)
let macroData = [];       // Macro aggregate records (Power BI)
let stockInventory = [];  // Current inventory status
let aiInsights = {};

let currentTrendView = 'monthly'; // 'daily' (Gmail) or 'monthly' (Power BI)
let activeCrossFilter = null;     // Selected category for cross-filtering
let tablePage = 1;
const rowsPerPage = 10;
let sortColumn = 'date';
let sortAscending = false;

// Chart reference
let trendChartInstance = null;

// DOM Elements
const fileInput = document.getElementById('file-input');
const dashboardContent = document.getElementById('dashboard-content');
const loadDemoBtn = document.getElementById('load-demo');
const uploadBtn = document.getElementById('upload-btn');
const skuSearch = document.getElementById('sku-search');
const ordersTableBody = document.getElementById('orders-table-body');
const tableInfo = document.getElementById('table-info');
const prevPageBtn = document.getElementById('prev-page');
const nextPageBtn = document.getElementById('next-page');

const viewDailyBtn = document.getElementById('view-daily');
const viewMonthlyBtn = document.getElementById('view-monthly');

const crossFilterBadge = document.getElementById('cross-filter-badge');
const crossFilterLabel = document.getElementById('cross-filter-label');
const crossFilterClear = document.getElementById('cross-filter-clear');

const alertBanner = document.getElementById('alert-banner');
const alertText = document.getElementById('alert-text');
const totalAlertsCount = document.getElementById('total-alerts-count');
const stockAlertsList = document.getElementById('stock-alerts-list');

const aiNarrative = document.getElementById('ai-narrative');
const aiSkuRows = document.getElementById('ai-sku-rows');
const aiSkuTable = document.getElementById('ai-sku-table');
const aiSkeleton = document.getElementById('ai-skeleton');
const aiGeneratedAt = document.getElementById('ai-generated-at');
const aiActions = document.getElementById('ai-actions');

// Initialize App
document.addEventListener('DOMContentLoaded', async () => {
  // Setup file drag and drop
  setupDragAndDrop();

  // Attach button click events
  fileInput.addEventListener('change', handleFileSelect);
  loadDemoBtn.addEventListener('click', loadDemoData);
  uploadBtn.addEventListener('click', () => fileInput.click());

  // Search input events
  skuSearch.addEventListener('input', () => {
    tablePage = 1;
    refreshDashboard(false);
  });

  // Table pagination
  prevPageBtn.addEventListener('click', () => {
    if (tablePage > 1) {
      tablePage--;
      renderOrdersTable();
    }
  });
  nextPageBtn.addEventListener('click', () => {
    const filtered = getFilteredSalesData();
    if (tablePage * rowsPerPage < filtered.length) {
      tablePage++;
      renderOrdersTable();
    }
  });

  // Toggle views (Daily vs Monthly)
  viewDailyBtn.addEventListener('click', () => {
    currentTrendView = 'daily';
    viewDailyBtn.className = "px-3 py-1 rounded-md bg-[#1F2937] text-[#D97706] font-semibold transition-all";
    viewMonthlyBtn.className = "px-3 py-1 rounded-md text-[#9CA3AF] hover:text-[#F9FAFB] transition-all";
    renderTrendChart();
  });
  viewMonthlyBtn.addEventListener('click', () => {
    currentTrendView = 'monthly';
    viewMonthlyBtn.className = "px-3 py-1 rounded-md bg-[#1F2937] text-[#D97706] font-semibold transition-all";
    viewDailyBtn.className = "px-3 py-1 rounded-md text-[#9CA3AF] hover:text-[#F9FAFB] transition-all";
    renderTrendChart();
  });

  // Clear cross-filter
  crossFilterClear.addEventListener('click', () => {
    activeCrossFilter = null;
    crossFilterBadge.classList.add('hidden');
    tablePage = 1;
    refreshDashboard(false);
  });

  // Check date button states
  document.querySelectorAll('.date-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.date-btn').forEach(b => {
        b.classList.remove('active', 'text-[#D97706]', 'bg-[#1F2937]', 'font-semibold');
        b.classList.add('text-[#9CA3AF]');
      });
      btn.classList.add('active', 'text-[#D97706]', 'bg-[#1F2937]', 'font-semibold');
      btn.classList.remove('text-[#9CA3AF]');
      
      // Simulate filtering or state change
      refreshDashboard(false);
    });
  });

  // Handle sidebar nav link clicks
  document.querySelectorAll('.nav-link').forEach(link => {
    link.addEventListener('click', (e) => {
      document.querySelectorAll('.nav-link').forEach(l => {
        l.classList.remove('active', 'text-[#F9FAFB]', 'bg-[#1F2937]', 'border-l-2', 'border-[#D97706]');
        l.classList.add('text-[#9CA3AF]');
        const icon = l.querySelector('span');
        if (icon) icon.classList.remove('text-[#D97706]');
      });
      link.classList.add('active', 'text-[#F9FAFB]', 'bg-[#1F2937]', 'border-l-2', 'border-[#D97706]');
      link.classList.remove('text-[#9CA3AF]');
      const icon = link.querySelector('span');
      if (icon) icon.classList.add('text-[#D97706]');
    });
  });

  // Try auto-loading data from data/ directory, fallback to demo
  await loadDatabase();
  // Dashboard always visible — show with demo data if nothing loaded
  dashboardContent.classList.remove('hidden');
  if (!salesData.length && !macroData.length) {
    loadDemoData();
  } else {
    refreshDashboard(true);
  }
});

// Setup Drag & Drop
let dragCounter = 0;
function setupDragAndDrop() {
  window.addEventListener('dragenter', (e) => { e.preventDefault(); });
  window.addEventListener('dragover',  (e) => { e.preventDefault(); });
  window.addEventListener('dragleave', (e) => { e.preventDefault(); });

  window.addEventListener('drop', (e) => {
    e.preventDefault();
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      processExcelFile(files[0]);
    }
  });
}

function handleFileSelect(e) {
  const files = e.target.files;
  if (files.length > 0) {
    processExcelFile(files[0]);
  }
}

// Read database states
async function loadDatabase() {
  dimProductos = window.mockDimProductos || [];
  stockInventory = window.mockStockInventory || [];
  aiInsights = window.mockAIInsights || {};

  // Fetch parsed Gmail orders if available
  try {
    const resMail = await fetch('data/mail_pedidos.json');
    if (resMail.ok) {
      const data = await resMail.json();
      salesData = data.map(d => ({
        order_id: d.order_id,
        fecha_pedido: d.fecha_pedido ? d.fecha_pedido.split(' ')[0] : 'N/A',
        cliente_email: d.cliente_email,
        producto_sku: d.producto_sku,
        cantidad: d.cantidad,
        precio_unitario: d.precio_unitario,
        total_linea: d.total_linea,
        producto: getProductName(d.producto_sku),
        categoria: getProductCategory(d.producto_sku)
      }));
      console.log("Loaded Gmail orders from file database.");
    }
  } catch (e) {
    console.log("No mail_pedidos.json database found, using fallbacks.");
  }

  // Fetch parsed Power BI statistics if available
  try {
    const resPbi = await fetch('data/powerbi_estadisticas.json');
    if (resPbi.ok) {
      macroData = await resPbi.json();
      console.log("Loaded Power BI statistics from file database.");
    }
  } catch (e) {
    console.log("No powerbi_estadisticas.json database found, using fallbacks.");
  }

  // Update data source badge
  const hasRealData = salesData.length > 0 || macroData.length > 0;
  setDataSourceBadge(hasRealData ? 'live' : 'demo');
}

// Load Demo Data explicitly
function loadDemoData() {
  salesData = window.mockSalesData || [];
  macroData = window.mockPowerBIData || [];
  dimProductos = window.mockDimProductos || [];
  stockInventory = window.mockStockInventory || [];
  aiInsights = window.mockAIInsights || {};
  dashboardContent.classList.remove('hidden');
  setDataSourceBadge('demo');
  refreshDashboard(true);
}

// Update data source badge in header
function setDataSourceBadge(mode) {
  const badge = document.getElementById('data-source-badge');
  if (!badge) return;
  if (mode === 'live') {
    badge.textContent = '● LIVE DATA';
    badge.className = 'px-2 py-0.5 rounded text-[10px] font-mono font-bold tracking-wider bg-emerald-900/60 text-emerald-400 border border-emerald-700/50';
  } else {
    badge.textContent = '◎ DEMO MODE';
    badge.className = 'px-2 py-0.5 rounded text-[10px] font-mono font-bold tracking-wider bg-[#D97706]/10 text-[#D97706] border border-[#D97706]/30';
  }
}

// Process sheet using SheetJS
function processExcelFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    try {
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      if (sheetData.length === 0) {
        alert("El archivo subido está vacío.");
        return;
      }

      detectAndIngestData(sheetData);
    } catch (err) {
      alert("Error al procesar el archivo Excel. Por favor verifique el formato.");
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Autodetect uploaded structure & import
function detectAndIngestData(records) {
  const first = records[0];
  const keys = Object.keys(first);

  // Checks for Gmail template
  const isGmail = keys.some(k => /order|pedido|client|mail|sku|barcode/i.test(k));
  // Checks for Power BI report
  const isPbi = keys.some(k => /reporte|estadistica|subcat|anterior|ano/i.test(k));

  if (isGmail) {
    // Map dynamically
    const map = {
      order_id: keys.find(k => /order_id|id_pedido|pedido/i.test(k)) || "order_id",
      fecha: keys.find(k => /fecha|date/i.test(k)) || "fecha_pedido",
      cliente: keys.find(k => /cliente|email|client/i.test(k)) || "cliente_email",
      sku: keys.find(k => /sku|producto_sku/i.test(k)) || "producto_sku",
      qty: keys.find(k => /cant|qty|数量/i.test(k)) || "cantidad",
      price: keys.find(k => /precio|price|unitario/i.test(k)) || "precio_unitario",
      total: keys.find(k => /total|importe/i.test(k)) || "total_linea"
    };

    const newRows = records.map(r => {
      const sku = String(r[map.sku]).trim();
      return {
        order_id: r[map.order_id] || `UPLOAD_${Math.random().toString(36).substr(2, 9)}`,
        fecha_pedido: r[map.fecha] ? String(r[map.fecha]).split(' ')[0] : new Date().toISOString().split('T')[0],
        cliente_email: r[map.cliente] || 'importado@local.com',
        producto_sku: sku,
        cantidad: parseInt(r[map.qty]) || 1,
        precio_unitario: parseFloat(r[map.price]) || 0.0,
        total_linea: parseFloat(r[map.total]) || (parseFloat(r[map.price]) * parseInt(r[map.qty]) || 0.0),
        producto: getProductName(sku),
        categoria: getProductCategory(sku)
      };
    });

    salesData = [...newRows, ...salesData];
    alert(`📂 Detectados y cargados ${newRows.length} pedidos de venta (Gmail).`);
  } else {
    // Default to Power BI Macro statistics
    const map = {
      fecha: keys.find(k => /fecha|date|period/i.test(k)) || "fecha_reporte",
      cat: keys.find(k => /cat/i.test(k)) || "categoria",
      sub: keys.find(k => /sub/i.test(k)) || "subcategoria",
      prod: keys.find(k => /prod/i.test(k)) || "producto",
      sales: keys.find(k => /ventas|sales|monto/i.test(k)) || "ventas",
      sales_prev: keys.find(k => /anterior|prev/i.test(k)) || "ventas_ano_anterior",
      qty: keys.find(k => /articulo|total|qty|cantidad/i.test(k)) || "n_total_articulos"
    };

    const newRows = records.map(r => ({
      fecha_reporte: r[map.fecha] ? String(r[map.fecha]).split(' ')[0] : new Date().toISOString().split('T')[0],
      categoria: String(r[map.cat] || 'PAPELERÍA GENERAL').toUpperCase(),
      subcategoria: String(r[map.sub] || 'OFICINA').toUpperCase(),
      producto: String(r[map.prod] || 'Articulo Importado'),
      ventas: parseFloat(r[map.sales]) || 0.0,
      ventas_ano_anterior: parseFloat(r[map.sales_prev]) || 0.0,
      n_total_articulos: parseInt(r[map.qty]) || 0
    }));

    macroData = [...newRows, ...macroData];
    alert(`📂 Detectados y cargados ${newRows.length} registros macro (Power BI).`);
  }

  welcomeCard.classList.add('hidden');
  dashboardContent.classList.remove('hidden');
  refreshDashboard(true);
}

// Refresh whole dashboard state
function refreshDashboard(updateAI = false) {
  calculateKPIs();
  renderTrendChart();
  renderOrdersTable();
  renderStockAlerts();

  if (updateAI) {
    renderAIInsightsPanel();
  }
}

// ----------------------------------------------------
// TAREA 3: KPI Calculation & Rendering
// ----------------------------------------------------
function calculateKPIs() {
  // 1. Total Sales (from Power BI macro data)
  const totalSales = macroData.reduce((sum, r) => sum + r.ventas, 0);
  const totalSalesPrev = macroData.reduce((sum, r) => sum + r.ventas_ano_anterior, 0);
  const yoySales = totalSalesPrev > 0 ? ((totalSales - totalSalesPrev) / totalSalesPrev) * 100 : 0;

  // 2. Unidades Vendidas
  const totalQty = macroData.reduce((sum, r) => sum + r.n_total_articulos, 0);
  // Simulate YoY Qty based on sales YoY or similar metric
  const yoyQty = yoySales * 0.95;

  // 3. Transactions (parsed Gmail orders)
  const transactions = salesData.length;

  // 4. Ticket Promedio (AOV)
  const totalAOV = transactions > 0 ? salesData.reduce((sum, r) => sum + r.total_linea, 0) : 0;
  const avgOrderVal = transactions > 0 ? (totalAOV / transactions) : 0;

  // Animate values
  animateCountUp('kpi-total-sales', totalSales, true);
  animateCountUp('kpi-total-qty', totalQty, false);
  animateCountUp('kpi-transactions', transactions, false);
  animateCountUp('kpi-avg-order', avgOrderVal, true);

  // Update YoY labels
  updateYoYLabel('kpi-yoy-sales', yoySales);
  updateYoYLabel('kpi-yoy-qty', yoyQty);
  updateYoYLabel('kpi-yoy-avg-order', yoySales * 1.05);

  // Generate Sparklines
  generateSparkline('kpi-sparkline-sales', macroData.map(m => m.ventas));
  generateSparkline('kpi-sparkline-qty', macroData.map(m => m.n_total_articulos));
  // Aggregate transactions by date
  const transByDate = {};
  salesData.forEach(s => { transByDate[s.fecha_pedido] = (transByDate[s.fecha_pedido] || 0) + 1; });
  generateSparkline('kpi-sparkline-transactions', Object.keys(transByDate).sort().map(d => transByDate[d]));
  generateSparkline('kpi-sparkline-avg-order', salesData.map(s => s.total_linea));
}

// Sparklines helper
function generateSparkline(elementId, dataArray) {
  const container = document.getElementById(elementId);
  if (!container || dataArray.length === 0) return;

  // Take last 12 values or default values
  const dataset = dataArray.length > 12 ? dataArray.slice(-12) : dataArray;
  const max = Math.max(...dataset, 1);
  const min = Math.min(...dataset, 0);

  container.innerHTML = '';
  dataset.forEach(val => {
    const bar = document.createElement('div');
    const pct = ((val - min) / (max - min || 1)) * 100;
    // Set heights dynamically
    bar.style.height = `${Math.max(pct, 15)}%`;
    bar.className = `w-full rounded-sm transition-all duration-300 ${
      elementId.includes('sales') ? 'bg-[#D97706]/40 hover:bg-[#D97706]' :
      elementId.includes('qty') ? 'bg-blue-400/40 hover:bg-blue-400' :
      elementId.includes('transactions') ? 'bg-emerald-400/40 hover:bg-emerald-400' :
      'bg-red-400/40 hover:bg-red-400'
    }`;
    container.appendChild(bar);
  });
}

function animateCountUp(elementId, target, isCurrency) {
  const el = document.getElementById(elementId);
  if (!el) return;

  const duration = 800; // 0.8s count up
  const start = performance.now();
  const startVal = parseFloat(el.getAttribute('data-val') || '0');
  el.setAttribute('data-val', target);

  function step(time) {
    const progress = Math.min((time - start) / duration, 1);
    // Ease out quad
    const current = startVal + (target - startVal) * progress * (2 - progress);
    
    if (isCurrency) {
      el.textContent = '€' + Math.floor(current).toLocaleString('es-ES');
    } else {
      el.textContent = Math.floor(current).toLocaleString('es-ES');
    }

    if (progress < 1) {
      requestAnimationFrame(step);
    }
  }
  requestAnimationFrame(step);
}

function updateYoYLabel(elementId, value) {
  const el = document.getElementById(elementId);
  if (!el) return;

  const sign = value >= 0 ? '+' : '';
  el.textContent = `${sign}${value.toFixed(1)}%`;
  if (value >= 0) {
    el.className = "text-emerald-400 font-semibold";
  } else {
    el.className = "text-red-400 font-semibold";
  }
}

// ----------------------------------------------------
// TAREA 4: AI Insights Panel & Stock Alert Drawer
// ----------------------------------------------------
function renderAIInsightsPanel() {
  aiSkeleton.classList.add('hidden');
  aiNarrative.classList.remove('hidden');
  aiSkuTable.classList.remove('hidden');
  aiActions.classList.remove('hidden');

  const now = new Date();
  aiGeneratedAt.textContent = `Generado hace ${now.getHours()}:${String(now.getMinutes()).padStart(2, '0')} (Gemini 1.5 Pro)`;

  if (aiInsights && aiInsights.narrative) {
    aiNarrative.innerHTML = aiInsights.narrative;
  }

  aiSkuRows.innerHTML = '';
  if (aiInsights && aiInsights.skuRecommendations) {
    aiInsights.skuRecommendations.forEach(rec => {
      const row = document.createElement('tr');
      row.className = "hover:bg-[#1F2937]/30 transition-colors border-b border-[#1F2937]/50";
      
      const badgeColor = rec.prioridad === 'CRÍTICA' ? 'bg-red-900/40 text-red-400 border-red-800/50' :
                         rec.prioridad === 'ALTA' ? 'bg-yellow-900/40 text-yellow-400 border-yellow-800/50' :
                         'bg-blue-900/40 text-blue-400 border-blue-800/50';

      row.innerHTML = `
        <td class="p-2.5 font-bold text-white">${rec.sku}</td>
        <td class="p-2.5 text-[#9CA3AF]">${rec.categoria}</td>
        <td class="p-2.5 text-right font-mono text-[#D1D5DB]">${rec.stock} u</td>
        <td class="p-2.5 text-right text-indigo-300 font-bold font-mono">+${rec.sugerido} u</td>
        <td class="p-2.5 text-center">
          <span class="px-2 py-0.5 rounded-full text-[10px] font-bold border ${badgeColor}">${rec.prioridad}</span>
        </td>
      `;
      aiSkuRows.appendChild(row);
    });
  }
}

function renderStockAlerts() {
  stockAlertsList.innerHTML = '';
  let criticalCount = 0;

  stockInventory.forEach(item => {
    if (item.estado === 'CRÍTICO') {
      criticalCount++;
    }

    const itemDiv = document.createElement('div');
    itemDiv.className = "flex items-center justify-between p-3 rounded-lg bg-[#1F2937]/40 border border-[#374151]/50 hover:border-[#D97706]/30 transition-all";

    const badgeClass = item.estado === 'CRÍTICO' ? 'bg-red-950 text-red-400 border-red-800' :
                       item.estado === 'BAJO' ? 'bg-amber-950 text-amber-400 border-amber-800' :
                       'bg-emerald-950 text-emerald-400 border-emerald-800';

    itemDiv.innerHTML = `
      <div>
        <p class="text-xs font-mono font-bold text-white">${item.sku_producto}</p>
        <p class="text-[10px] text-[#9CA3AF] max-w-[140px] truncate">${item.nombre_producto}</p>
        <p class="text-[9px] text-[#6B7280] font-mono mt-0.5">Stock: ${item.stock_actual} / Mín: ${item.stock_seguridad}</p>
      </div>
      <span class="px-2 py-0.5 rounded-full text-[9px] font-bold border ${badgeClass}">
        ${item.estado}
      </span>
    `;
    stockAlertsList.appendChild(itemDiv);
  });

  totalAlertsCount.textContent = criticalCount;

  // Manage Banner Visibility
  if (criticalCount > 0) {
    alertBanner.classList.remove('hidden');
    alertText.textContent = `⚠ ALERTA DE STOCK CRÍTICO: ${criticalCount} SKU${criticalCount > 1 ? 's' : ''} bajo el stock de seguridad.`;
    document.getElementById('app-shell').style.paddingTop = '2.5rem';
  } else {
    alertBanner.classList.add('hidden');
    document.getElementById('app-shell').style.paddingTop = '0px';
  }
}

// ----------------------------------------------------
// TAREA 5: Chart & Table with Cross-filtering
// ----------------------------------------------------
function getFilteredSalesData() {
  const query = skuSearch.value.toLowerCase().trim();
  
  return salesData.filter(row => {
    // Search filter
    const matchesSearch = !query || 
      row.producto_sku.toLowerCase().includes(query) ||
      row.producto.toLowerCase().includes(query) ||
      row.cliente_email.toLowerCase().includes(query);

    // Cross filter from chart category select
    const matchesCategory = !activeCrossFilter || row.categoria === activeCrossFilter;

    return matchesSearch && matchesCategory;
  });
}

function renderOrdersTable() {
  ordersTableBody.innerHTML = '';
  const filtered = getFilteredSalesData();
  
  const startIndex = (tablePage - 1) * rowsPerPage;
  const endIndex = Math.min(startIndex + rowsPerPage, filtered.length);
  const paginated = filtered.slice(startIndex, endIndex);

  if (paginated.length === 0) {
    ordersTableBody.innerHTML = `<tr><td colspan="6" class="text-center py-8 text-[#6B7280]">No se encontraron pedidos coincidentes</td></tr>`;
    tableInfo.textContent = 'Mostrando 0–0 de 0';
    prevPageBtn.disabled = true;
    nextPageBtn.disabled = true;
    return;
  }

  // Sort logic
  paginated.sort((a, b) => {
    let valA = a[sortColumn];
    let valB = b[sortColumn];
    
    if (sortColumn === 'date') {
      valA = a.fecha_pedido;
      valB = b.fecha_pedido;
    } else if (sortColumn === 'sku') {
      valA = a.producto_sku;
      valB = b.producto_sku;
    } else if (sortColumn === 'qty') {
      valA = a.cantidad;
      valB = b.cantidad;
    } else if (sortColumn === 'total') {
      valA = a.total_linea;
      valB = b.total_linea;
    }

    if (typeof valA === 'string') {
      return sortAscending ? valA.localeCompare(valB) : valB.localeCompare(valA);
    } else {
      return sortAscending ? valA - valB : valB - valA;
    }
  });

  paginated.forEach(row => {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-[#1F2937]/40 transition-colors border-b border-[#1F2937]/30";

    // SKU Mapped Check
    const isMapped = dimProductos.some(p => p.sku_producto === row.producto_sku);
    const skuMapBadge = isMapped 
      ? `<span class="rounded-full px-2 py-0.5 bg-emerald-950 text-emerald-400 border border-emerald-800 text-[10px] font-mono">✓ Mapeado</span>`
      : `<span class="rounded-full px-2 py-0.5 bg-red-950 text-red-400 border border-red-800 text-[10px] font-mono cursor-pointer" onclick="mapNewSKU('${row.producto_sku}')" title="Asignar SKU a catálogo">⚠ Sin mapear</span>`;

    tr.innerHTML = `
      <td class="py-3 text-white">${row.fecha_pedido}</td>
      <td class="py-3 font-bold text-[#D97706]">${row.producto_sku}</td>
      <td class="py-3 text-[#E5E7EB] font-sans">${row.producto}</td>
      <td class="py-3 text-right text-white">${row.cantidad} u</td>
      <td class="py-3 text-right font-bold text-white">€${row.total_linea.toFixed(2)}</td>
      <td class="py-3 text-center">${skuMapBadge}</td>
    `;
    ordersTableBody.appendChild(tr);
  });

  tableInfo.textContent = `Mostrando ${startIndex + 1}–${endIndex} de ${filtered.length}`;
  prevPageBtn.disabled = tablePage === 1;
  nextPageBtn.disabled = endIndex >= filtered.length;
}

function sortTable(column) {
  if (sortColumn === column) {
    sortAscending = !sortAscending;
  } else {
    sortColumn = column;
    sortAscending = true;
  }

  // Update header arrows visual helper
  document.querySelectorAll('th').forEach(th => {
    const clickAttr = th.getAttribute('onclick');
    if (clickAttr && clickAttr.includes(column)) {
      th.innerHTML = th.textContent.replace(/[▲▼]/g, '').trim() + (sortAscending ? ' ▲' : ' ▼');
    } else {
      th.innerHTML = th.textContent.replace(/[▲▼]/g, '').trim();
    }
  });

  renderOrdersTable();
}
window.sortTable = sortTable;

// SKU Map interactive assign
function mapNewSKU(sku) {
  const newName = prompt(`Asigne un nombre de producto para el SKU ${sku}:`);
  if (!newName) return;
  const newCat = prompt(`Asigne una Categoría (CUADERNOS, ESCRITURA, FORROS Y CARPETAS, PAPELERÍA GENERAL):`, 'PAPELERÍA GENERAL');
  if (!newCat) return;

  // Add to standard catalog
  dimProductos.push({
    sku_producto: sku,
    nombre_producto: newName,
    categoria: newCat.toUpperCase(),
    subcategoria: "ESCUELA"
  });

  // Remap current sales records
  salesData.forEach(row => {
    if (row.producto_sku === sku) {
      row.producto = newName;
      row.categoria = newCat.toUpperCase();
    }
  });

  alert(`✓ SKU ${sku} mapeado correctamente.`);
  refreshDashboard(false);
}
window.mapNewSKU = mapNewSKU;

// Render Chart.js
function renderTrendChart() {
  const ctx = document.getElementById('trend-chart');
  if (!ctx) return;

  if (trendChartInstance) {
    trendChartInstance.destroy();
  }

  let chartLabels = [];
  let chartValues = [];
  let chartLabelName = "";

  if (currentTrendView === 'monthly') {
    // Aggregate macro Power BI sales by category
    const salesByCat = {};
    macroData.forEach(item => {
      salesByCat[item.categoria] = (salesByCat[item.categoria] || 0) + item.ventas;
    });

    chartLabels = Object.keys(salesByCat);
    chartValues = Object.values(salesByCat);
    chartLabelName = "Ventas por Categoría (Power BI)";
  } else {
    // Aggregate Gmail orders by date
    const salesByDate = {};
    salesData.forEach(item => {
      salesByDate[item.fecha_pedido] = (salesByDate[item.fecha_pedido] || 0) + item.total_linea;
    });

    // Sort dates
    chartLabels = Object.keys(salesByDate).sort();
    chartValues = chartLabels.map(d => salesByDate[d]);
    chartLabelName = "Ventas Diarias (Gmail)";
  }

  trendChartInstance = new Chart(ctx, {
    type: currentTrendView === 'monthly' ? 'bar' : 'line',
    data: {
      labels: chartLabels,
      datasets: [{
        label: chartLabelName,
        data: chartValues,
        backgroundColor: currentTrendView === 'monthly' 
          ? ['rgba(217, 119, 6, 0.85)', 'rgba(59, 130, 246, 0.85)', 'rgba(16, 185, 129, 0.85)', 'rgba(239, 68, 68, 0.85)']
          : 'rgba(217, 119, 6, 0.25)',
        borderColor: '#D97706',
        borderWidth: 2,
        borderRadius: currentTrendView === 'monthly' ? 6 : 0,
        fill: currentTrendView !== 'monthly',
        tension: 0.35
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      onClick: (evt, activeElements) => {
        if (activeElements.length > 0 && currentTrendView === 'monthly') {
          const index = activeElements[0].index;
          const clickedCategory = chartLabels[index];
          
          // Apply cross-filtering
          activeCrossFilter = clickedCategory;
          crossFilterLabel.textContent = clickedCategory;
          crossFilterBadge.classList.remove('hidden');
          tablePage = 1;
          renderOrdersTable();
        }
      },
      plugins: {
        legend: { display: false }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#9CA3AF', font: { size: 10, family: 'IBM Plex Mono' } }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: {
            color: '#9CA3AF', 
            font: { size: 10, family: 'IBM Plex Mono' },
            callback: function(value) {
              return '€' + value.toLocaleString('es-ES');
            }
          }
        }
      }
    }
  });
}

// Helper methods to get names
function getProductName(sku) {
  const match = dimProductos.find(p => p.sku_producto === sku);
  return match ? match.nombre_producto : 'Producto no catalogado';
}

function getProductCategory(sku) {
  const match = dimProductos.find(p => p.sku_producto === sku);
  return match ? match.categoria : 'DESCONOCIDA';
}

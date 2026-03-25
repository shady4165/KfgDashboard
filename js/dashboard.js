/* ===================================================================
 *  KFG Executive Dashboard – Dashboard Rendering Module
 *  Builds all charts, KPI cards, and tables for the 8 department pages.
 *  Uses data from demo mode or live SharePoint (via KFGGraph).
 *  Reads chart/table configuration from window.KFGConfig.
 *
 *  Namespace: window.KFGDashboard
 * =================================================================== */
(function () {
  'use strict';

  // --------------------------------------------------------------------------
  //  Demo Data
  // --------------------------------------------------------------------------

  const DEMO = {
    maintenance: {
      kpis: { openJobs: 10, completedMTD: 24, overdueJobs: 2, avgResponseHrs: 6.5, rag: 'Green' },
      jobs: [
        { id: 'MNT-001', site: 'Site A', desc: 'Boiler service', cat: 'Planned', priority: 'High', status: 'In Progress', cost: 350 },
        { id: 'MNT-002', site: 'Site B', desc: 'Roof leak repair', cat: 'Reactive', priority: 'Critical', status: 'Open', cost: 1200 },
        { id: 'MNT-003', site: 'Site C', desc: 'Door lock replacement', cat: 'Reactive', priority: 'Medium', status: 'Completed', cost: 85 },
        { id: 'MNT-004', site: 'Site D', desc: 'Fire alarm test', cat: 'Regulatory', priority: 'High', status: 'Scheduled', cost: 450 },
        { id: 'MNT-005', site: 'Site E', desc: 'Painting hallway', cat: 'Aesthetic', priority: 'Low', status: 'On Hold', cost: 200 },
      ],
      byCat: { Planned: 1, Reactive: 2, Regulatory: 1, Aesthetic: 1 },
      bySite: { LBN: 3, 'RWN-KA': 1, 'ODN-KA': 2, 'ODN-MU': 1, WCN: 2, 'RWN-JP': 1, 'ODN-BU': 0, 'RWN-SA': 0 },
    },
    capex: {
      kpis: { totalBudget: 10049443, utilised: 4665024, remaining: 5384419, utilisedPct: 46.4, negativeSites: 6, rag: 'Amber' },
      nurseries: [
        { name: 'LBN', ebitda: 751508, budget: 67636, utilised: 175358, balance: -107722, balPct: -159, request: 0, balAfter: -107722 },
        { name: 'RWN-KA', ebitda: 5346627, budget: 481196, utilised: 286004, balance: 195192, balPct: 41, request: 276000, balAfter: -80808 },
        { name: 'RWN-SA', ebitda: 5779407, budget: 520147, utilised: 131069, balance: 389078, balPct: 75, request: 0, balAfter: 389078 },
        { name: 'RWN-MU', ebitda: 3949963, budget: 355497, utilised: 72970, balance: 282527, balPct: 79, request: 0, balAfter: 282527 },
        { isTotal: true, name: 'Total RWN AUH', ebitda: 31603101, budget: 2844280, utilised: 764109, balance: 2080171, balPct: 73, request: 276000, balAfter: 1804171 },
        { name: 'ODN-KA', ebitda: 5097890, budget: 458810, utilised: 68011, balance: 390799, balPct: 85, request: 0, balAfter: 390799 },
        { name: 'ODN-MU', ebitda: 3318590, budget: 298673, utilised: 230263, balance: 68410, balPct: 23, request: 0, balAfter: 68410 },
        { isTotal: true, name: 'Total ODN AUH', ebitda: 14912704, budget: 1342144, utilised: 579670, balance: 762474, balPct: 57, request: 0, balAfter: 762474 },
        { isGrand: true, name: 'Total AUH', ebitda: 47267313, budget: 4254060, utilised: 1519137, balance: 2734923, balPct: 64, request: 276000, balAfter: 2458923 },
        { name: 'WCN', ebitda: 4169036, budget: 375213, utilised: 139070, balance: 236143, balPct: 63, request: 0, balAfter: 236143 },
        { name: 'WCN-OC', ebitda: 1819705, budget: 163773, utilised: 317748, balance: -153975, balPct: -94, request: 0, balAfter: -153975 },
        { isTotal: true, name: 'Total Willow', ebitda: 10489715, budget: 944074, utilised: 989046, balance: -44972, balPct: -5, request: 0, balAfter: -44972 },
        { name: 'ODN-BU', ebitda: 3293268, budget: 296394, utilised: 19981, balance: 276413, balPct: 93, request: 0, balAfter: 276413 },
        { name: 'ODN-LL', ebitda: 1751925, budget: 157673, utilised: 97190, balance: 60483, balPct: 38, request: 200000, balAfter: -139517 },
        { name: 'ODN-SZR', ebitda: 2791753, budget: 251258, utilised: 59617, balance: 191642, balPct: 76, request: 338126, balAfter: -146485 },
        { isTotal: true, name: 'Total ODN DXB', ebitda: 15133833, budget: 1362044, utilised: 417751, balance: 944293, balPct: 69, request: 1606126, balAfter: -661833 },
        { name: 'RWN-JP', ebitda: 8961723, budget: 806555, utilised: 133758, balance: 672797, balPct: 83, request: 0, balAfter: 672797 },
        { name: 'RWN-FUR', ebitda: 2459663, budget: 221370, utilised: 14321, balance: 207049, balPct: 94, request: 0, balAfter: 207049 },
        { isTotal: true, name: 'Total RWN DXB', ebitda: 16760478, budget: 1508443, utilised: 304018, balance: 1204425, balPct: 80, request: 0, balAfter: 1204425 },
        { isGrand: true, name: 'KFG (29+HO)', ebitda: 87853007, budget: 10049443, utilised: 4665024, balance: 5384419, balPct: 54, request: 3425126, balAfter: 1959293 },
      ],
      regions: {
        labels: ['AUH', 'Willow', 'ODN DXB', 'RWN DXB', 'Doha'],
        budget: [4254060, 944074, 1362044, 1508443, 339984],
        utilised: [1519137, 989046, 417751, 304018, 116586],
      },
    },
    projects: {
      kpis: { active: 5, onTrack: 3, atRisk: 1, delayed: 1, rag: 'Amber' },
      list: [
        { id: 'PM-001', name: 'Nursery Refurbishment Site A', owner: 'J. Smith', end: '31/03/2026', pct: 65, budget: 50000, spent: 32000, priority: 'High', rag: 'Green' },
        { id: 'PM-002', name: 'IT Infrastructure Upgrade', owner: 'T. Brown', end: '30/04/2026', pct: 30, budget: 80000, spent: 24000, priority: 'Medium', rag: 'Amber' },
        { id: 'PM-003', name: 'Regulatory Compliance Audit', owner: 'R. Jones', end: '28/02/2026', pct: 90, budget: 15000, spent: 13500, priority: 'Critical', rag: 'Green' },
        { id: 'PM-004', name: 'New Nursery Fit-Out Site B', owner: 'S. Davis', end: '31/05/2026', pct: 10, budget: 120000, spent: 12000, priority: 'High', rag: 'Red' },
        { id: 'PM-005', name: 'Supplier Renegotiation', owner: 'L. Wilson', end: '31/03/2026', pct: 50, budget: 5000, spent: 2500, priority: 'Medium', rag: 'Green' },
      ],
    },
    procurement: {
      kpis: { activePOs: 5, pending: 1, overdue: 1, spendYTD: 30950, rag: 'Amber' },
      pos: [
        { po: 'PO-2026-001', supplier: 'ABC Supplies', desc: 'Cleaning materials Q1', dept: 'Maintenance', value: 3500, status: 'Delivered', delivery: '14/01/2026' },
        { po: 'PO-2026-002', supplier: 'XYZ Contractors', desc: 'Electrical works', dept: 'Maintenance', value: 8200, status: 'Approved', delivery: '28/03/2026' },
        { po: 'PO-2026-003', supplier: 'Tech Solutions', desc: 'IT hardware', dept: 'IT', value: 12000, status: 'Pending Approval', delivery: '30/03/2026' },
        { po: 'PO-2026-004', supplier: 'Office Depot', desc: 'Stationery', dept: 'Admin', value: 450, status: 'Delivered', delivery: '24/01/2026' },
        { po: 'PO-2026-005', supplier: 'BuildRight Ltd', desc: 'Carpentry Site D', dept: 'Maintenance', value: 6800, status: 'In Progress', delivery: '15/04/2026' },
      ],
      bySt: { Delivered: 2, Approved: 1, 'Pending Approval': 1, 'In Progress': 1 },
      byDept: { Maintenance: 18500, IT: 12000, Admin: 450 },
    },
    it: {
      kpis: { openTickets: 3, resolved: 12, critical: 1, slaPct: 94, rag: 'Amber' },
      tickets: [
        { id: 'IT-001', site: 'Head Office', cat: 'Hardware', desc: 'Laptop screen broken', priority: 'High', status: 'In Progress', raised: '01/03/2026' },
        { id: 'IT-002', site: 'Site B', cat: 'Network', desc: 'WiFi down', priority: 'Critical', status: 'Resolved', raised: '02/03/2026' },
        { id: 'IT-003', site: 'Site C', cat: 'Software', desc: 'Excel not opening', priority: 'Low', status: 'Open', raised: '05/03/2026' },
        { id: 'IT-004', site: 'Head Office', cat: 'Security', desc: 'Password reset', priority: 'Medium', status: 'Resolved', raised: '08/03/2026' },
        { id: 'IT-005', site: 'Head Office', cat: 'Infrastructure', desc: 'Server backup failed', priority: 'Critical', status: 'In Progress', raised: '10/03/2026' },
      ],
      byCat: { Hardware: 1, Network: 1, Software: 1, Security: 1, Infrastructure: 1 },
      trend: { labels: ['Jan', 'Feb', 'Mar'], open: [8, 5, 3], resolved: [6, 9, 12] },
    },
    ma: {
      kpis: { pipeline: 3, dueDiligence: 1, completed: 1, weightedValue: 3412500, rag: 'Green' },
      deals: [
        { id: 'MA-001', target: '[Confidential]', type: 'Acquisition', value: 2500000, stage: 'Due Diligence', lead: 'CEO', close: '30/06/2026', prob: '65%' },
        { id: 'MA-002', target: '[Confidential]', type: 'Asset Purchase', value: 850000, stage: 'Heads of Terms', lead: 'CFO', close: '31/05/2026', prob: '80%' },
        { id: 'MA-003', target: '[Confidential]', type: 'Acquisition', value: 5000000, stage: 'Initial Review', lead: 'CEO', close: '31/12/2026', prob: '30%' },
      ],
      bySt: { 'Initial Review': 1, 'Heads of Terms': 1, 'Due Diligence': 1 },
      byVal: { 'Initial Review': 5000000, 'Heads of Terms': 850000, 'Due Diligence': 2500000 },
    },
    greenfield: {
      kpis: { inPipeline: 4, underConstruction: 1, completed: 1, totalCapex: 11750000, totalEBITDA: 3653946, rag: 'Green' },
      sites: [
        { id: 'GF-001', name: 'Madinat Al Riyadh - Villa', city: 'Abu Dhabi', indoor: 7840, outdoor: 3660, enrol: 158, lease: '10 Years', ebitda: 1872085, rent: 350000, capex: 4500000, brand: 'Odyssey', opening: 'Jul-26', status: 'Under Construction' },
        { id: 'GF-002', name: 'Damac Hills - Retail', city: 'Dubai', indoor: 6564, outdoor: 2902, enrol: 146, lease: '5 Years', ebitda: 1781861, rent: 1307404, capex: 4000000, brand: 'Redwood', opening: 'Sep-26', status: 'Fit-Out' },
        { id: 'GF-003', name: 'Site C - TBC', city: 'Riyadh', indoor: '', outdoor: '', enrol: '', lease: '', ebitda: '', rent: '', capex: '', brand: 'TBC', opening: '2027', status: 'Feasibility' },
        { id: 'GF-004', name: 'Site D - TBC', city: 'Abu Dhabi', indoor: '', outdoor: '', enrol: '', lease: '', ebitda: '', rent: '', capex: '', brand: 'TBC', opening: '2027', status: 'Planning' },
      ],
    },
    other: {
      kpis: { active: 5, onTrack: 2, atRisk: 2, overdue: 1, rag: 'Amber' },
      projects: [
        { id: 'OP-001', name: 'Staff Wellbeing Programme', cat: 'People & HR', owner: 'HR Manager', pct: 25, budget: 20000, priority: 'Medium', rag: 'Red' },
        { id: 'OP-002', name: 'Operations Manual Update', cat: 'Process Improvement', owner: 'Ops Director', pct: 60, budget: 5000, priority: 'High', rag: 'Amber' },
        { id: 'OP-003', name: 'Supplier ESG Review', cat: 'Research', owner: 'Procurement Lead', pct: 10, budget: 8000, priority: 'Medium', rag: 'Red' },
        { id: 'OP-004', name: 'Emergency Procedures Refresh', cat: 'Health & Safety', owner: 'H&S Manager', pct: 95, budget: 3500, priority: 'Critical', rag: 'Green' },
        { id: 'OP-005', name: 'Brand Refresh Collateral', cat: 'Marketing', owner: 'Marketing Lead', pct: 5, budget: 12000, priority: 'Low', rag: 'Red' },
      ],
      byCat: { 'People & HR': 1, 'Process Improvement': 1, Research: 1, 'Health & Safety': 1, Marketing: 1 },
      bySt: { Green: 1, Amber: 1, Red: 3 },
    },
  };

  // --------------------------------------------------------------------------
  //  Internal State
  // --------------------------------------------------------------------------

  let DATA = null;       // current dataset (demo or live)
  let _isDemo = true;    // true when using demo data
  const CHARTS = {};     // { canvasId: Chart instance }

  // --------------------------------------------------------------------------
  //  Format Helpers
  // --------------------------------------------------------------------------

  /** Format number with locale (comma separator). */
  function fn(v) {
    if (v == null || v === '') return '-';
    return Number(v).toLocaleString('en-AE');
  }

  /** Format as AED currency. */
  function faed(v) {
    if (v == null || v === '') return '-';
    return 'AED ' + Number(v).toLocaleString('en-AE');
  }

  /** Format as percentage. */
  function fpct(v) {
    if (v == null || v === '') return '-';
    return Number(v).toFixed(1) + '%';
  }

  /** Return HTML for a status pill with appropriate colour. */
  function pill(s) {
    if (!s) return '';
    const colours = {
      Green: '#10B981', 'On Track': '#10B981', Completed: '#10B981', Delivered: '#10B981', Resolved: '#10B981',
      Amber: '#F59E0B', 'At Risk': '#F59E0B', 'In Progress': '#F59E0B', Approved: '#F59E0B', 'Heads of Terms': '#F59E0B',
      'Under Construction': '#F59E0B', 'Fit-Out': '#F59E0B', Scheduled: '#3B82F6',
      Red: '#EF4444', Delayed: '#EF4444', Critical: '#EF4444', Open: '#EF4444', Overdue: '#EF4444',
      'Pending Approval': '#94A3B8', 'On Hold': '#94A3B8', Feasibility: '#94A3B8', Planning: '#94A3B8',
      'Initial Review': '#94A3B8', 'Due Diligence': '#2563EB',
    };
    const bg = colours[s] || '#94A3B8';
    return '<span style="display:inline-block;padding:2px 10px;border-radius:12px;font-size:0.78rem;font-weight:600;color:#fff;background:' + bg + '">' + s + '</span>';
  }

  /** Return HTML for a KPI card. */
  function kpiCard(label, value, sub, color, pillText) {
    // color can be a class name (green, amber, red, blue, grey) or a hex string
    var cls = color || 'blue';
    var colorMap = { '#2563EB': 'blue', '#10B981': 'green', '#F59E0B': 'amber', '#EF4444': 'red', '#94A3B8': 'grey', '#8B5CF6': 'purple' };
    if (colorMap[cls]) cls = colorMap[cls];
    var html = '<div class="kpi-card ' + cls + '" role="listitem">';
    html += '<div class="kpi-label">' + label + '</div>';
    html += '<div class="kpi-value">' + value + '</div>';
    if (sub) html += '<div class="kpi-sub">' + sub + '</div>';
    if (pillText) html += '<div class="kpi-badge ' + cls + '"><span class="kpi-dot"></span>' + pillText + '</div>';
    html += '</div>';
    return html;
  }

  // --------------------------------------------------------------------------
  //  Config Helper — read from KFGConfig
  // --------------------------------------------------------------------------

  /** Get chart config from admin panel, with safe fallback. */
  function getCfg(chartId) {
    if (window.KFGConfig && typeof window.KFGConfig.getChartConfig === 'function') {
      return window.KFGConfig.getChartConfig(chartId);
    }
    return null;
  }

  /** Update a card's title and subtitle from config, if the card element exists. */
  function applyCardTitle(chartId) {
    const cfg = getCfg(chartId);
    if (!cfg) return;
    // Look for a card wrapper that contains this canvas
    const canvas = document.getElementById(chartId);
    if (!canvas) return;
    const card = canvas.closest('.card, .chart-card, [data-chart-id]');
    if (!card) return;
    if (cfg.title) {
      const titleEl = card.querySelector('.card-title, h3, h4');
      if (titleEl) titleEl.textContent = cfg.title;
    }
    if (cfg.subtitle) {
      const subEl = card.querySelector('.card-sub, .card-subtitle, .text-muted, small');
      if (subEl) subEl.textContent = cfg.subtitle;
    }
  }

  // --------------------------------------------------------------------------
  //  Chart Helper
  // --------------------------------------------------------------------------

  /**
   * Create or replace a Chart.js chart, merging admin config overrides.
   * @param {string} id - Canvas element ID
   * @param {object} cfg - Chart.js configuration object
   */
  function mkChart(id, cfg) {
    const el = document.getElementById(id);
    if (!el) return null;
    if (CHARTS[id]) { CHARTS[id].destroy(); delete CHARTS[id]; }

    // Apply admin config overrides from KFGConfig
    const adminCfg = getCfg(id);
    if (adminCfg) {
      if (adminCfg.type) cfg.type = adminCfg.type;
      if (adminCfg.indexAxis && cfg.options) {
        cfg.options.indexAxis = adminCfg.indexAxis;
      }
      if (adminCfg.colors && cfg.data && cfg.data.datasets) {
        cfg.data.datasets.forEach((ds, i) => {
          if (adminCfg.colors[i] !== undefined) {
            if (cfg.type === 'doughnut' || cfg.type === 'pie') {
              ds.backgroundColor = adminCfg.colors;
            } else {
              ds.backgroundColor = adminCfg.colors[i];
              ds.borderColor = adminCfg.colors[i];
            }
          }
        });
      }
      if (adminCfg.showLegend != null && cfg.options && cfg.options.plugins) {
        cfg.options.plugins.legend = cfg.options.plugins.legend || {};
        cfg.options.plugins.legend.display = adminCfg.showLegend;
      }
      if (adminCfg.legendPosition && cfg.options && cfg.options.plugins && cfg.options.plugins.legend) {
        cfg.options.plugins.legend.position = adminCfg.legendPosition;
      }
      if (adminCfg.cutout != null && (cfg.type === 'doughnut' || cfg.type === 'pie')) {
        cfg.options = cfg.options || {};
        cfg.options.cutout = adminCfg.cutout;
      }
      if (adminCfg.borderRadius != null && cfg.data && cfg.data.datasets) {
        cfg.data.datasets.forEach(ds => { ds.borderRadius = adminCfg.borderRadius; });
      }
      if (adminCfg.tension != null && cfg.data && cfg.data.datasets) {
        cfg.data.datasets.forEach(ds => { ds.tension = adminCfg.tension; });
      }
      if (adminCfg.fill != null && cfg.data && cfg.data.datasets) {
        cfg.data.datasets.forEach(ds => { ds.fill = adminCfg.fill; });
      }
    }

    // Apply card title/subtitle from config
    applyCardTitle(id);

    CHARTS[id] = new Chart(el, cfg);
    return CHARTS[id];
  }

  // --------------------------------------------------------------------------
  //  Shared chart option presets
  // --------------------------------------------------------------------------

  function noGridOpts(extra) {
    const base = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#F1F5F9' }, ticks: { font: { size: 11 } } },
      },
    };
    return Object.assign(base, extra || {});
  }

  function doughnutOpts(pos) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: true, position: pos || 'right', labels: { font: { size: 12 }, padding: 14 } } },
    };
  }

  // --------------------------------------------------------------------------
  //  Build: Home / Executive Summary
  // --------------------------------------------------------------------------

  function buildHome() {
    const D = DATA;
    if (!D) return;

    // RAG strip
    const ragEl = document.getElementById('rag-strip');
    if (ragEl) {
      const depts = [
        { key: 'projects', icon: '\u{1F3D7}', label: 'Projects', page: 'projects' },
        { key: 'maintenance', icon: '\u{1F527}', label: 'Maintenance', page: 'maintenance' },
        { key: 'capex', icon: '\u{1F4B0}', label: 'Capex', page: 'capex' },
        { key: 'procurement', icon: '\u{1F6D2}', label: 'Procurement', page: 'procurement' },
        { key: 'it', icon: '\u{1F4BB}', label: 'IT', page: 'it' },
        { key: 'ma', icon: '\u{1F91D}', label: 'M&A', page: 'ma' },
        { key: 'greenfield', icon: '\u{1F331}', label: 'Greenfield', page: 'greenfield' },
        { key: 'other', icon: '\u{1F4CB}', label: 'Other', page: 'other' },
      ];
      let html = '';
      depts.forEach(function (d) {
        var rag = D[d.key] && D[d.key].kpis ? D[d.key].kpis.rag : '';
        var ragLower = (rag || '').toLowerCase() || 'grey';
        html += '<div class="rag-dept" onclick="goByKey(\'' + d.page + '\')" role="listitem">';
        html += '<div class="rag-icon">' + d.icon + '</div>';
        html += '<div class="rag-name">' + d.label + '</div>';
        html += '<div class="rag-badge ' + ragLower + '">' + (rag || 'No Data') + '</div>';
        html += '</div>';
      });
      ragEl.innerHTML = html;
    }

    // Portfolio health doughnut
    const ragCounts = { Green: 0, Amber: 0, Red: 0 };
    ['maintenance', 'capex', 'projects', 'procurement', 'it', 'ma', 'greenfield', 'other'].forEach(k => {
      const rag = D[k] && D[k].kpis ? D[k].kpis.rag : 'Green';
      if (ragCounts[rag] !== undefined) ragCounts[rag]++;
    });

    const ragCfg = getCfg('c-rag');
    mkChart('c-rag', {
      type: 'doughnut',
      data: {
        labels: ['Green', 'Amber', 'Red'],
        datasets: [{
          data: [ragCounts.Green, ragCounts.Amber, ragCounts.Red],
          backgroundColor: ragCfg ? ragCfg.colors : ['#10B981', '#F59E0B', '#EF4444'],
        }],
      },
      options: Object.assign(doughnutOpts(ragCfg ? ragCfg.legendPosition : 'right'), {
        cutout: ragCfg ? ragCfg.cutout : '65%',
      }),
    });

    // Capex by region bar
    const capReg = D.capex ? D.capex.regions : null;
    if (capReg) {
      const crCfg = getCfg('c-cap-reg');
      mkChart('c-cap-reg', {
        type: 'bar',
        data: {
          labels: capReg.labels,
          datasets: [
            { label: 'Budget', data: capReg.budget, backgroundColor: crCfg ? crCfg.colors[0] : '#2563EB', borderRadius: 4 },
            { label: 'Utilised', data: capReg.utilised, backgroundColor: crCfg ? crCfg.colors[1] : '#F59E0B', borderRadius: 4 },
          ],
        },
        options: noGridOpts({ plugins: { legend: { display: true, position: crCfg ? crCfg.legendPosition : 'top' } } }),
      });
    }

    // Department summary table
    const tbl = document.getElementById('t-home');
    if (tbl) {
      const rows = [
        { dept: 'Maintenance', metric: 'Open Jobs', value: D.maintenance ? D.maintenance.kpis.openJobs : '-', rag: D.maintenance ? D.maintenance.kpis.rag : '' },
        { dept: 'Capex', metric: 'Utilised %', value: D.capex ? fpct(D.capex.kpis.utilisedPct) : '-', rag: D.capex ? D.capex.kpis.rag : '' },
        { dept: 'Projects', metric: 'Active', value: D.projects ? D.projects.kpis.active : '-', rag: D.projects ? D.projects.kpis.rag : '' },
        { dept: 'Procurement', metric: 'Active POs', value: D.procurement ? D.procurement.kpis.activePOs : '-', rag: D.procurement ? D.procurement.kpis.rag : '' },
        { dept: 'IT', metric: 'Open Tickets', value: D.it ? D.it.kpis.openTickets : '-', rag: D.it ? D.it.kpis.rag : '' },
        { dept: 'M&A', metric: 'Pipeline', value: D.ma ? D.ma.kpis.pipeline : '-', rag: D.ma ? D.ma.kpis.rag : '' },
        { dept: 'Greenfield', metric: 'In Pipeline', value: D.greenfield ? D.greenfield.kpis.inPipeline : '-', rag: D.greenfield ? D.greenfield.kpis.rag : '' },
        { dept: 'Other Projects', metric: 'Active', value: D.other ? D.other.kpis.active : '-', rag: D.other ? D.other.kpis.rag : '' },
      ];
      let html = '';
      rows.forEach(r => {
        html += '<tr><td>' + r.dept + '</td><td>' + r.metric + '</td><td>' + r.value + '</td><td>' + pill(r.rag) + '</td></tr>';
      });
      tbl.innerHTML = html;
    }
  }

  // --------------------------------------------------------------------------
  //  Build: Maintenance
  // --------------------------------------------------------------------------

  function buildMaintenance() {
    const D = DATA;
    if (!D || !D.maintenance) return;
    const m = D.maintenance;

    // KPIs
    const kpiEl = document.getElementById('k-maint');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('Open Jobs', m.kpis.openJobs, '', '#2563EB'),
        kpiCard('Completed MTD', m.kpis.completedMTD, '', '#10B981'),
        kpiCard('Overdue', m.kpis.overdueJobs, '', '#EF4444'),
        kpiCard('Avg Response', m.kpis.avgResponseHrs + 'h', '', '#F59E0B'),
      ].join('');
    }

    // Jobs by site bar
    const siteCfg = getCfg('c-maint-site');
    const siteLabels = Object.keys(m.bySite);
    const siteData = Object.values(m.bySite);
    mkChart('c-maint-site', {
      type: 'bar',
      data: {
        labels: siteLabels,
        datasets: [{ label: 'Open Jobs', data: siteData, backgroundColor: siteCfg ? siteCfg.colors[0] : '#2563EB', borderRadius: 4 }],
      },
      options: noGridOpts({ indexAxis: siteCfg ? siteCfg.indexAxis : 'y' }),
    });

    // Jobs by category doughnut
    const catCfg = getCfg('c-maint-cat');
    mkChart('c-maint-cat', {
      type: 'doughnut',
      data: {
        labels: Object.keys(m.byCat),
        datasets: [{ data: Object.values(m.byCat), backgroundColor: catCfg ? catCfg.colors : ['#2563EB', '#EF4444', '#F59E0B', '#10B981'] }],
      },
      options: Object.assign(doughnutOpts(catCfg ? catCfg.legendPosition : 'right'), { cutout: catCfg ? catCfg.cutout : '60%' }),
    });

    // Jobs table
    const tbl = document.getElementById('t-maint-jobs');
    if (tbl) {
      let html = '';
      m.jobs.forEach(j => {
        html += '<tr><td>' + j.id + '</td><td>' + j.site + '</td><td>' + j.desc + '</td><td>' + j.cat + '</td><td>' + pill(j.priority) + '</td><td>' + pill(j.status) + '</td><td>' + fn(j.cost) + '</td></tr>';
      });
      tbl.innerHTML = html;
    }
  }

  // --------------------------------------------------------------------------
  //  Build: Capex
  // --------------------------------------------------------------------------

  function buildCapex() {
    const D = DATA;
    if (!D || !D.capex) return;
    const c = D.capex;

    // KPIs
    const kpiEl = document.getElementById('k-capex');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('Total Budget', faed(c.kpis.totalBudget), '', '#2563EB'),
        kpiCard('Utilised', faed(c.kpis.utilised), fpct(c.kpis.utilisedPct), '#F59E0B'),
        kpiCard('Remaining', faed(c.kpis.remaining), '', '#10B981'),
        kpiCard('Over-Budget Sites', c.kpis.negativeSites, '', '#EF4444'),
      ].join('');
    }

    // Budget vs utilised by region bar
    const regCfg = getCfg('c-capex-reg');
    mkChart('c-capex-reg', {
      type: 'bar',
      data: {
        labels: c.regions.labels,
        datasets: [
          { label: 'Budget', data: c.regions.budget, backgroundColor: regCfg ? regCfg.colors[0] : '#2563EB', borderRadius: 4 },
          { label: 'Utilised', data: c.regions.utilised, backgroundColor: regCfg ? regCfg.colors[1] : '#F59E0B', borderRadius: 4 },
        ],
      },
      options: noGridOpts({ plugins: { legend: { display: true, position: regCfg ? regCfg.legendPosition : 'top' } } }),
    });

    // Over-budget sites bar
    const negCfg = getCfg('c-capex-neg');
    const negSites = c.nurseries.filter(n => !n.isTotal && !n.isGrand && n.balance < 0);
    mkChart('c-capex-neg', {
      type: 'bar',
      data: {
        labels: negSites.map(n => n.name),
        datasets: [{ label: 'Negative Balance', data: negSites.map(n => Math.abs(n.balance)), backgroundColor: negCfg ? negCfg.colors[0] : '#EF4444', borderRadius: 4 }],
      },
      options: noGridOpts({ indexAxis: negCfg ? negCfg.indexAxis : 'y' }),
    });

    // Nursery budget table
    const tbl = document.getElementById('t-capex');
    if (tbl) {
      let html = '';
      c.nurseries.forEach(n => {
        const cls = n.isGrand ? 'font-weight:800;background:#E2E8F0' : n.isTotal ? 'font-weight:700;background:#F1F5F9' : '';
        const balColor = n.balance < 0 ? 'color:#EF4444' : '';
        html += '<tr style="' + cls + '">';
        html += '<td>' + n.name + '</td>';
        html += '<td>' + faed(n.ebitda) + '</td>';
        html += '<td>' + faed(n.budget) + '</td>';
        html += '<td>' + faed(n.utilised) + '</td>';
        html += '<td style="' + balColor + '">' + faed(n.balance) + '</td>';
        html += '<td style="' + balColor + '">' + fpct(n.balPct) + '</td>';
        html += '<td>' + faed(n.request) + '</td>';
        html += '<td style="' + (n.balAfter < 0 ? 'color:#EF4444' : '') + '">' + faed(n.balAfter) + '</td>';
        html += '</tr>';
      });
      tbl.innerHTML = html;
    }
  }

  // --------------------------------------------------------------------------
  //  Build: Projects
  // --------------------------------------------------------------------------

  function buildProjects() {
    const D = DATA;
    if (!D || !D.projects) return;
    const p = D.projects;

    // KPIs
    const kpiEl = document.getElementById('k-pm');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('Active Projects', p.kpis.active, '', '#2563EB'),
        kpiCard('On Track', p.kpis.onTrack, '', '#10B981'),
        kpiCard('At Risk', p.kpis.atRisk, '', '#F59E0B'),
        kpiCard('Delayed', p.kpis.delayed, '', '#EF4444'),
      ].join('');
    }

    // Status doughnut
    const stCfg = getCfg('c-pm-status');
    mkChart('c-pm-status', {
      type: 'doughnut',
      data: {
        labels: ['On Track', 'At Risk', 'Delayed'],
        datasets: [{ data: [p.kpis.onTrack, p.kpis.atRisk, p.kpis.delayed], backgroundColor: stCfg ? stCfg.colors : ['#10B981', '#F59E0B', '#EF4444'] }],
      },
      options: Object.assign(doughnutOpts(stCfg ? stCfg.legendPosition : 'right'), { cutout: stCfg ? stCfg.cutout : '60%' }),
    });

    // Budget vs spend bar
    const budCfg = getCfg('c-pm-budget');
    mkChart('c-pm-budget', {
      type: 'bar',
      data: {
        labels: p.list.map(x => x.name.length > 20 ? x.name.substring(0, 20) + '...' : x.name),
        datasets: [
          { label: 'Budget', data: p.list.map(x => x.budget), backgroundColor: budCfg ? budCfg.colors[0] : '#2563EB', borderRadius: 4 },
          { label: 'Spent', data: p.list.map(x => x.spent), backgroundColor: budCfg ? budCfg.colors[1] : '#F59E0B', borderRadius: 4 },
        ],
      },
      options: noGridOpts({ plugins: { legend: { display: true, position: budCfg ? budCfg.legendPosition : 'top' } } }),
    });

    // Projects table
    const tbl = document.getElementById('t-pm');
    if (tbl) {
      let html = '';
      p.list.forEach(r => {
        html += '<tr><td>' + r.id + '</td><td>' + r.name + '</td><td>' + r.owner + '</td><td>' + r.end + '</td>';
        html += '<td><div style="background:#E2E8F0;border-radius:6px;height:8px;width:100px;display:inline-block;vertical-align:middle"><div style="background:#2563EB;height:100%;border-radius:6px;width:' + r.pct + '%"></div></div> ' + r.pct + '%</td>';
        html += '<td>' + faed(r.budget) + '</td><td>' + faed(r.spent) + '</td>';
        html += '<td>' + pill(r.priority) + '</td><td>' + pill(r.rag) + '</td></tr>';
      });
      tbl.innerHTML = html;
    }
  }

  // --------------------------------------------------------------------------
  //  Build: Procurement
  // --------------------------------------------------------------------------

  function buildProcurement() {
    const D = DATA;
    if (!D || !D.procurement) return;
    const pr = D.procurement;

    // KPIs
    const kpiEl = document.getElementById('k-proc');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('Active POs', pr.kpis.activePOs, '', '#2563EB'),
        kpiCard('Pending Approval', pr.kpis.pending, '', '#F59E0B'),
        kpiCard('Overdue', pr.kpis.overdue, '', '#EF4444'),
        kpiCard('Spend YTD', faed(pr.kpis.spendYTD), '', '#10B981'),
      ].join('');
    }

    // PO status doughnut
    const stCfg = getCfg('c-proc-status');
    mkChart('c-proc-status', {
      type: 'doughnut',
      data: {
        labels: Object.keys(pr.bySt),
        datasets: [{ data: Object.values(pr.bySt), backgroundColor: stCfg ? stCfg.colors : ['#10B981', '#2563EB', '#F59E0B', '#3B82F6'] }],
      },
      options: Object.assign(doughnutOpts(stCfg ? stCfg.legendPosition : 'right'), { cutout: stCfg ? stCfg.cutout : '60%' }),
    });

    // Spend by department bar
    const spCfg = getCfg('c-proc-spend');
    mkChart('c-proc-spend', {
      type: 'bar',
      data: {
        labels: Object.keys(pr.byDept),
        datasets: [{ label: 'Spend (AED)', data: Object.values(pr.byDept), backgroundColor: spCfg ? spCfg.colors[0] : '#2563EB', borderRadius: 4 }],
      },
      options: noGridOpts(),
    });

    // PO table
    const tbl = document.getElementById('t-proc');
    if (tbl) {
      let html = '';
      pr.pos.forEach(r => {
        html += '<tr><td>' + r.po + '</td><td>' + r.supplier + '</td><td>' + r.desc + '</td><td>' + r.dept + '</td><td>' + fn(r.value) + '</td><td>' + pill(r.status) + '</td><td>' + r.delivery + '</td></tr>';
      });
      tbl.innerHTML = html;
    }
  }

  // --------------------------------------------------------------------------
  //  Build: IT
  // --------------------------------------------------------------------------

  function buildIT() {
    const D = DATA;
    if (!D || !D.it) return;
    const it = D.it;

    // KPIs
    const kpiEl = document.getElementById('k-it');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('Open Tickets', it.kpis.openTickets, '', '#2563EB'),
        kpiCard('Resolved MTD', it.kpis.resolved, '', '#10B981'),
        kpiCard('Critical', it.kpis.critical, '', '#EF4444'),
        kpiCard('SLA Met', fpct(it.kpis.slaPct), '', '#F59E0B'),
      ].join('');
    }

    // Tickets by category bar
    const catCfg = getCfg('c-it-cat');
    mkChart('c-it-cat', {
      type: 'bar',
      data: {
        labels: Object.keys(it.byCat),
        datasets: [{ label: 'Tickets', data: Object.values(it.byCat), backgroundColor: catCfg ? catCfg.colors[0] : '#2563EB', borderRadius: 4 }],
      },
      options: noGridOpts(),
    });

    // Open vs resolved trend line
    const trCfg = getCfg('c-it-trend');
    mkChart('c-it-trend', {
      type: 'line',
      data: {
        labels: it.trend.labels,
        datasets: [
          { label: 'Open', data: it.trend.open, borderColor: trCfg ? trCfg.colors[0] : '#EF4444', backgroundColor: trCfg ? trCfg.colors[0] + '20' : '#EF444420', tension: 0.4, fill: true },
          { label: 'Resolved', data: it.trend.resolved, borderColor: trCfg ? trCfg.colors[1] : '#10B981', backgroundColor: trCfg ? trCfg.colors[1] + '20' : '#10B98120', tension: 0.4, fill: true },
        ],
      },
      options: noGridOpts({ plugins: { legend: { display: true, position: trCfg ? trCfg.legendPosition : 'top' } } }),
    });

    // Tickets table
    const tbl = document.getElementById('t-it');
    if (tbl) {
      let html = '';
      it.tickets.forEach(r => {
        html += '<tr><td>' + r.id + '</td><td>' + r.site + '</td><td>' + r.cat + '</td><td>' + r.desc + '</td><td>' + pill(r.priority) + '</td><td>' + pill(r.status) + '</td><td>' + r.raised + '</td></tr>';
      });
      tbl.innerHTML = html;
    }
  }

  // --------------------------------------------------------------------------
  //  Build: M&A
  // --------------------------------------------------------------------------

  function buildMA() {
    const D = DATA;
    if (!D || !D.ma) return;
    const ma = D.ma;

    // KPIs
    const kpiEl = document.getElementById('k-ma');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('Pipeline', ma.kpis.pipeline, '', '#2563EB'),
        kpiCard('Due Diligence', ma.kpis.dueDiligence, '', '#F59E0B'),
        kpiCard('Completed', ma.kpis.completed, '', '#10B981'),
        kpiCard('Weighted Value', faed(ma.kpis.weightedValue), '', '#8B5CF6'),
      ].join('');
    }

    // Deals by stage bar
    const stCfg = getCfg('c-ma-stage');
    const stLabels = Object.keys(ma.bySt);
    mkChart('c-ma-stage', {
      type: 'bar',
      data: {
        labels: stLabels,
        datasets: [{ label: 'Deals', data: Object.values(ma.bySt), backgroundColor: stCfg ? stCfg.colors : ['#94A3B8', '#F59E0B', '#2563EB'], borderRadius: 4 }],
      },
      options: noGridOpts(),
    });

    // Value by stage bar
    const valCfg = getCfg('c-ma-val');
    const valLabels = Object.keys(ma.byVal);
    mkChart('c-ma-val', {
      type: 'bar',
      data: {
        labels: valLabels,
        datasets: [{ label: 'Value (AED)', data: Object.values(ma.byVal), backgroundColor: valCfg ? valCfg.colors : ['#94A3B8', '#F59E0B', '#2563EB'], borderRadius: 4 }],
      },
      options: noGridOpts(),
    });

    // Deals table
    const tbl = document.getElementById('t-ma');
    if (tbl) {
      let html = '';
      ma.deals.forEach(r => {
        html += '<tr><td>' + r.id + '</td><td>' + r.target + '</td><td>' + r.type + '</td><td>' + faed(r.value) + '</td><td>' + pill(r.stage) + '</td><td>' + r.lead + '</td><td>' + r.close + '</td><td>' + r.prob + '</td></tr>';
      });
      tbl.innerHTML = html;
    }
  }

  // --------------------------------------------------------------------------
  //  Build: Greenfield
  // --------------------------------------------------------------------------

  function buildGreenfield() {
    const D = DATA;
    if (!D || !D.greenfield) return;
    const g = D.greenfield;

    // KPIs
    const kpiEl = document.getElementById('k-gf');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('In Pipeline', g.kpis.inPipeline, '', '#2563EB'),
        kpiCard('Under Construction', g.kpis.underConstruction, '', '#F59E0B'),
        kpiCard('Completed', g.kpis.completed, '', '#10B981'),
        kpiCard('Total Capex', faed(g.kpis.totalCapex), '', '#8B5CF6'),
      ].join('');
    }

    // Site cards
    const cardsEl = document.getElementById('gf-site-cards');
    if (cardsEl) {
      let html = '';
      g.sites.forEach(s => {
        html += '<div style="background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,0.08);min-width:260px">';
        html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px">';
        html += '<strong>' + s.name + '</strong>' + pill(s.status);
        html += '</div>';
        html += '<div style="font-size:0.82rem;color:#64748B;margin-bottom:6px">' + s.city + ' &bull; ' + s.brand + ' &bull; ' + s.opening + '</div>';
        if (s.indoor) html += '<div style="font-size:0.8rem">Indoor: ' + fn(s.indoor) + ' sqft &bull; Outdoor: ' + fn(s.outdoor) + ' sqft</div>';
        if (s.enrol) html += '<div style="font-size:0.8rem">Enrolment: ' + fn(s.enrol) + ' &bull; Lease: ' + s.lease + '</div>';
        if (s.ebitda) html += '<div style="font-size:0.8rem;margin-top:4px">EBITDA: ' + faed(s.ebitda) + ' &bull; Rent: ' + faed(s.rent) + '</div>';
        if (s.capex) html += '<div style="font-size:0.8rem">Capex: ' + faed(s.capex) + '</div>';
        html += '</div>';
      });
      cardsEl.innerHTML = html;
    }

    // EBITDA:Rent ratio bar
    const ratioCfg = getCfg('c-gf-ratio');
    const withData = g.sites.filter(s => s.ebitda && s.rent);
    mkChart('c-gf-ratio', {
      type: 'bar',
      data: {
        labels: withData.map(s => s.name.length > 20 ? s.name.substring(0, 20) + '...' : s.name),
        datasets: [{
          label: 'EBITDA:Rent Ratio',
          data: withData.map(s => (s.ebitda / s.rent).toFixed(1)),
          backgroundColor: withData.map(s => (s.ebitda / s.rent) >= 5
            ? (ratioCfg ? ratioCfg.colors[0] : '#10B981')
            : (ratioCfg ? ratioCfg.colors[1] : '#F59E0B')),
          borderRadius: 6,
        }],
      },
      options: noGridOpts(),
    });

    // Capex by site bar
    const capCfg = getCfg('c-gf-capex');
    const withCapex = g.sites.filter(s => s.capex);
    mkChart('c-gf-capex', {
      type: 'bar',
      data: {
        labels: withCapex.map(s => s.name.length > 20 ? s.name.substring(0, 20) + '...' : s.name),
        datasets: [{ label: 'Capex (AED)', data: withCapex.map(s => s.capex), backgroundColor: capCfg ? capCfg.colors[0] : '#2563EB', borderRadius: 6 }],
      },
      options: noGridOpts(),
    });
  }

  // --------------------------------------------------------------------------
  //  Build: Other Projects
  // --------------------------------------------------------------------------

  function buildOther() {
    const D = DATA;
    if (!D || !D.other) return;
    const o = D.other;

    // KPIs
    const kpiEl = document.getElementById('k-op');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('Active', o.kpis.active, '', '#2563EB'),
        kpiCard('On Track', o.kpis.onTrack, '', '#10B981'),
        kpiCard('At Risk', o.kpis.atRisk, '', '#F59E0B'),
        kpiCard('Overdue', o.kpis.overdue, '', '#EF4444'),
      ].join('');
    }

    // By category doughnut
    const catCfg = getCfg('c-op-cat');
    mkChart('c-op-cat', {
      type: 'doughnut',
      data: {
        labels: Object.keys(o.byCat),
        datasets: [{ data: Object.values(o.byCat), backgroundColor: catCfg ? catCfg.colors : ['#2563EB', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6'] }],
      },
      options: Object.assign(doughnutOpts(catCfg ? catCfg.legendPosition : 'right'), { cutout: catCfg ? catCfg.cutout : '60%' }),
    });

    // By status doughnut
    const stCfg = getCfg('c-op-status');
    mkChart('c-op-status', {
      type: 'doughnut',
      data: {
        labels: Object.keys(o.bySt),
        datasets: [{ data: Object.values(o.bySt), backgroundColor: stCfg ? stCfg.colors : ['#10B981', '#F59E0B', '#EF4444'] }],
      },
      options: Object.assign(doughnutOpts(stCfg ? stCfg.legendPosition : 'right'), { cutout: stCfg ? stCfg.cutout : '60%' }),
    });

    // Projects table
    const tbl = document.getElementById('t-op');
    if (tbl) {
      let html = '';
      o.projects.forEach(r => {
        html += '<tr><td>' + r.id + '</td><td>' + r.name + '</td><td>' + r.cat + '</td><td>' + r.owner + '</td>';
        html += '<td><div style="background:#E2E8F0;border-radius:6px;height:8px;width:100px;display:inline-block;vertical-align:middle"><div style="background:#2563EB;height:100%;border-radius:6px;width:' + r.pct + '%"></div></div> ' + r.pct + '%</td>';
        html += '<td>' + faed(r.budget) + '</td><td>' + pill(r.priority) + '</td><td>' + pill(r.rag) + '</td></tr>';
      });
      tbl.innerHTML = html;
    }
  }

  // --------------------------------------------------------------------------
  //  Navigation
  // --------------------------------------------------------------------------

  /** Page key to section-ID mapping (must match id="pg-*" in index.html). */
  const PAGES = {
    home: 'pg-home',
    maintenance: 'pg-maintenance',
    capex: 'pg-capex',
    projects: 'pg-projects',
    procurement: 'pg-procurement',
    it: 'pg-it',
    ma: 'pg-ma',
    greenfield: 'pg-greenfield',
    other: 'pg-other',
  };

  /** Build function mapping per page. */
  const BUILDERS = {
    home: buildHome,
    maintenance: buildMaintenance,
    capex: buildCapex,
    projects: buildProjects,
    procurement: buildProcurement,
    it: buildIT,
    ma: buildMA,
    greenfield: buildGreenfield,
    other: buildOther,
  };

  /**
   * Navigate to a page, optionally highlighting the clicked nav element.
   * @param {string} name - Page key (e.g. 'maintenance')
   * @param {HTMLElement} [el] - The clicked nav element
   */
  function go(name, el) {
    // Remove .active from all pages (CSS: .page { display:none } / .page.active { display:block })
    document.querySelectorAll('.page').forEach(s => s.classList.remove('active'));

    // Add .active to target page
    const targetId = PAGES[name];
    if (targetId) {
      const target = document.getElementById(targetId);
      if (target) target.classList.add('active');
    }

    // Update active nav button
    document.querySelectorAll('.nav-item, [data-page]').forEach(n => n.classList.remove('active'));
    if (el) el.classList.add('active');
    else {
      const navEl = document.querySelector('[data-page="' + name + '"]');
      if (navEl) navEl.classList.add('active');
    }

    // Build / refresh the page content
    if (BUILDERS[name]) BUILDERS[name]();

    // Scroll to top of content area
    const main = document.getElementById('main-content');
    if (main) main.scrollTop = 0;
    window.scrollTo({ top: 60, behavior: 'smooth' });
  }

  /**
   * Navigate by key without a reference element.
   * @param {string} name - Page key
   */
  function goByKey(name) {
    // Try to find and activate the matching nav element
    const navEl = document.querySelector('[data-page="' + name + '"]');
    go(name, navEl);
  }

  // --------------------------------------------------------------------------
  //  Build All
  // --------------------------------------------------------------------------

  function buildAll() {
    buildHome();
    buildMaintenance();
    buildCapex();
    buildProjects();
    buildProcurement();
    buildIT();
    buildMA();
    buildGreenfield();
    buildOther();
  }

  // --------------------------------------------------------------------------
  //  Refresh All — destroy existing charts and rebuild
  // --------------------------------------------------------------------------

  function refreshAll() {
    // Destroy all chart instances
    Object.keys(CHARTS).forEach(id => {
      if (CHARTS[id]) { CHARTS[id].destroy(); delete CHARTS[id]; }
    });
    buildAll();
  }

  // --------------------------------------------------------------------------
  //  Public API
  // --------------------------------------------------------------------------

  window.KFGDashboard = {

    /** Initialize with demo data and build all pages. */
    init() {
      DATA = JSON.parse(JSON.stringify(DEMO));
      _isDemo = true;
      console.log('[KFGDashboard] Initialised with demo data.');
      buildAll();
    },

    /** Build all pages using current data. */
    buildAll,

    /** Destroy all charts and rebuild from current data. */
    refreshAll,

    /**
     * Replace internal data with live SharePoint data.
     * Call this after KFGGraph fetches all departments.
     * @param {object} data - Structured data matching DEMO shape
     */
    setData(data) {
      DATA = data;
      _isDemo = false;
      console.log('[KFGDashboard] Switched to live data.');
      refreshAll();
    },

    /** Get the current data object. */
    getData() {
      return DATA;
    },

    /** Check whether the dashboard is currently using demo data. */
    isDemo() {
      return _isDemo;
    },

    // Expose navigation
    go,
    goByKey,

    // Expose individual build functions for targeted refresh
    buildHome,
    buildMaintenance,
    buildCapex,
    buildProjects,
    buildProcurement,
    buildIT,
    buildMA,
    buildGreenfield,
    buildOther,

    // Expose format helpers for use by other modules
    fn,
    faed,
    fpct,
    pill,
    kpiCard,
  };

  console.log('[KFGDashboard] Module loaded.');
})();

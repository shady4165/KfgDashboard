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
        { id: 'MNT-001', site: 'Site A', desc: 'Boiler service', cat: 'Planned', priority: 'High', status: 'In Progress', cost: 350, raisedDate: '01/03/2026' },
        { id: 'MNT-002', site: 'Site B', desc: 'Roof leak repair', cat: 'Reactive', priority: 'Critical', status: 'Open', cost: 1200, raisedDate: '05/03/2026' },
        { id: 'MNT-003', site: 'Site C', desc: 'Door lock replacement', cat: 'Reactive', priority: 'Medium', status: 'Completed', cost: 85, raisedDate: '08/03/2026' },
        { id: 'MNT-004', site: 'Site D', desc: 'Fire alarm test', cat: 'Regulatory', priority: 'High', status: 'Scheduled', cost: 450, raisedDate: '12/03/2026' },
        { id: 'MNT-005', site: 'Site E', desc: 'Painting hallway', cat: 'Aesthetic', priority: 'Low', status: 'On Hold', cost: 200, raisedDate: '18/03/2026' },
      ],
      byCat: { Planned: 1, Reactive: 2, Regulatory: 1, Aesthetic: 1 },
      bySite: { LBN: 3, 'RWN-KA': 1, 'ODN-KA': 2, 'ODN-MU': 1, WCN: 2, 'RWN-JP': 1, 'ODN-BU': 0, 'RWN-SA': 0 },
      poCosts: [
        { id: 'PO-001', site: 'Site A', category: 'HVAC', vendor: 'CoolTech', desc: 'AC servicing', amount: 2400, date: '03/03/2026' },
        { id: 'PO-002', site: 'Site B', category: 'Civil', vendor: 'BuildRight', desc: 'Leak rectification', amount: 1800, date: '10/03/2026' },
        { id: 'PO-003', site: 'Site A', category: 'Electrical', vendor: 'VoltFix', desc: 'Lighting replacement', amount: 950, date: '15/03/2026' },
        { id: 'PO-004', site: 'Site C', category: 'HVAC', vendor: 'CoolTech', desc: 'Compressor repair', amount: 3100, date: '20/03/2026' },
        { id: 'PO-005', site: 'Site B', category: 'Fire', vendor: 'SafeAlarm', desc: 'Fire panel testing', amount: 1200, date: '28/03/2026' },
      ],
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
      milestones: [
        { projectId: 'PM-001', project: 'Nursery Refurbishment Site A', milestone: 'Design Sign-off', dueDate: '15/01/2026', owner: 'J. Smith', status: 'Completed', priority: 'High', notes: '' },
        { projectId: 'PM-001', project: 'Nursery Refurbishment Site A', milestone: 'Works Start', dueDate: '01/02/2026', owner: 'Contractor', status: 'Completed', priority: 'High', notes: '' },
        { projectId: 'PM-001', project: 'Nursery Refurbishment Site A', milestone: 'Handover', dueDate: '31/03/2026', owner: 'J. Smith', status: 'In Progress', priority: 'Critical', notes: 'On track' },
        { projectId: 'PM-002', project: 'IT Infrastructure Upgrade', milestone: 'Server Procurement', dueDate: '28/02/2026', owner: 'T. Brown', status: 'Completed', priority: 'High', notes: '' },
        { projectId: 'PM-002', project: 'IT Infrastructure Upgrade', milestone: 'Installation', dueDate: '30/04/2026', owner: 'IT Team', status: 'In Progress', priority: 'High', notes: '' },
        { projectId: 'PM-004', project: 'New Nursery Fit-Out Site B', milestone: 'Planning Permission', dueDate: '31/01/2026', owner: 'S. Davis', status: 'Completed', priority: 'Critical', notes: '' },
        { projectId: 'PM-004', project: 'New Nursery Fit-Out Site B', milestone: 'Contractor Appointed', dueDate: '15/02/2026', owner: 'S. Davis', status: 'Open', priority: 'High', notes: 'Delayed — retendering' },
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
      byNursery: { 'Nursery A': 12000, 'Nursery B': 9500, 'Nursery C': 7200, 'HQ': 2250 },
      suppliers: [
        { supplier: 'ABC Supplies', category: 'Cleaning', contractExp: '31/12/2026', contact: 'info@abc.com', annualSpend: 42000, status: 'Active', insuranceValid: 'Yes', notes: 'Annual contract' },
        { supplier: 'XYZ Contractors', category: 'Maintenance', contractExp: '30/06/2026', contact: 'contracts@xyz.com', annualSpend: 98400, status: 'Active', insuranceValid: 'Yes', notes: 'Preferred vendor' },
        { supplier: 'Tech Solutions', category: 'IT', contractExp: '28/02/2027', contact: 'sales@tech.com', annualSpend: 144000, status: 'Active', insuranceValid: 'Valid', notes: 'Hardware + support' },
        { supplier: 'Office Depot', category: 'Stationery', contractExp: '31/03/2026', contact: 'corp@od.com', annualSpend: 5400, status: 'Active', insuranceValid: 'Yes', notes: '' },
        { supplier: 'BuildRight Ltd', category: 'Carpentry', contractExp: '30/09/2026', contact: 'projects@brl.com', annualSpend: 81600, status: 'Active', insuranceValid: 'Yes', notes: 'Approved contractor' },
      ],
      rawSuppliers: [
        { 'Supplier Name': 'ABC Supplies', 'Category': 'Cleaning', 'Contract Exp.': '31/12/2026', 'Contact': 'info@abc.com', 'Annual Spend (AED)': 42000, 'Status': 'Active', 'Insurance Valid': 'Yes' },
        { 'Supplier Name': 'XYZ Contractors', 'Category': 'Maintenance', 'Contract Exp.': '30/06/2026', 'Contact': 'contracts@xyz.com', 'Annual Spend (AED)': 98400, 'Status': 'Active', 'Insurance Valid': 'Yes' },
        { 'Supplier Name': 'Tech Solutions', 'Category': 'IT', 'Contract Exp.': '28/02/2027', 'Contact': 'sales@tech.com', 'Annual Spend (AED)': 144000, 'Status': 'Active', 'Insurance Valid': 'Valid' },
        { 'Supplier Name': 'Office Depot', 'Category': 'Stationery', 'Contract Exp.': '31/03/2026', 'Contact': 'corp@od.com', 'Annual Spend (AED)': 5400, 'Status': 'Active', 'Insurance Valid': 'Yes' },
        { 'Supplier Name': 'BuildRight Ltd', 'Category': 'Carpentry', 'Contract Exp.': '30/09/2026', 'Contact': 'projects@brl.com', 'Annual Spend (AED)': 81600, 'Status': 'Active', 'Insurance Valid': 'Yes' },
      ],
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
      milestones: [
        { dealId: 'MA-001', target: '[Confidential]', milestone: 'Financial Due Diligence', dueDate: '30/04/2026', owner: 'CFO', status: 'In Progress', priority: 'Critical', notes: 'Awaiting audited accounts' },
        { dealId: 'MA-001', target: '[Confidential]', milestone: 'Legal Review', dueDate: '15/05/2026', owner: 'Legal Counsel', status: 'Open', priority: 'High', notes: '' },
        { dealId: 'MA-002', target: '[Confidential]', milestone: 'Heads of Terms Signed', dueDate: '30/04/2026', owner: 'CEO', status: 'Completed', priority: 'High', notes: 'Signed 10 Apr 2026' },
        { dealId: 'MA-002', target: '[Confidential]', milestone: 'Board Approval', dueDate: '15/05/2026', owner: 'Board', status: 'Open', priority: 'Critical', notes: 'Next board meeting May' },
        { dealId: 'MA-003', target: '[Confidential]', milestone: 'Initial NDA', dueDate: '30/04/2026', owner: 'CEO', status: 'Completed', priority: 'Medium', notes: '' },
      ],
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

  // Inline Chart.js plugin: draw value labels on top of bar columns
  const _barLabelPlugin = {
    id: 'kfgBarLabels',
    afterDatasetsDraw: function (chart, args, opts) {
      if (!opts || !opts.enabled) return;
      var ctx = chart.ctx;
      chart.data.datasets.forEach(function (dataset, i) {
        var meta = chart.getDatasetMeta(i);
        if (meta.hidden) return;
        meta.data.forEach(function (el, idx) {
          var val = dataset.data[idx];
          if (!val && val !== 0) return;
          ctx.save();
          ctx.font = 'bold 10px sans-serif';
          ctx.fillStyle = '#374151';
          ctx.textAlign = 'center';
          ctx.textBaseline = 'bottom';
          var label = Number(val).toLocaleString('en-AE', { maximumFractionDigits: 0 });
          ctx.fillText(label, el.x, el.y - 3);
          ctx.restore();
        });
      });
    },
  };
  if (window.Chart) window.Chart.register(_barLabelPlugin);
  const MAINT_ANALYTICS_STATE = {
    jobs: { sites: [], cats: [], from: '', to: '' },
    po: { sites: [], cats: [], from: '', to: '' }
  };
  const PROC_ANALYTICS_STATE = {
    sites: [], cats: [], from: '', to: '' }
;
  const CAPEX_CATEGORY_STATE = { sites: [], cats: [] };
  const PROC_SUPPLIER_STATE = { suppliers: [], cats: [], from: '', to: '' };
  const MA_MILESTONE_STATE = { projects: [], from: '', to: '' };
  const PM_MILESTONE_STATE = { projects: [], from: '', to: '' };
  const OTHER_ACTION_STATE = { projects: [], from: '', to: '' };

  // --------------------------------------------------------------------------
  //  Format Helpers
  // --------------------------------------------------------------------------

  /** Format number with locale (comma separator). */
  function fn(v) {
    if (v == null || v === '') return '-';
    return Number(v).toLocaleString('en-AE');
  }

  /** Format as AED currency (no decimals). */
  function faed(v) {
    if (v == null || v === '') return '-';
    return 'AED ' + Number(v).toLocaleString('en-AE', { maximumFractionDigits: 0 });
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


  function parseDashboardDate(v) {
    if (!v) return null;
    if (v instanceof Date && !isNaN(v.getTime())) return v;
    if (typeof v === 'number' && v > 1000) return new Date((v - 25569) * 86400 * 1000);
    var s = String(v).trim();
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
      var y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      return new Date(y, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    }
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  function monthKey(d) {
    if (!d) return '';
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  }

  function monthLabel(key) {
    if (!key) return 'No Date';
    var parts = key.split('-');
    var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, 1);
    return d.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' });
  }

  function safeUnique(list) {
    return Array.from(new Set((list || []).filter(Boolean).map(function (v) { return String(v).trim(); }))).sort();
  }

  function canonicalSiteValue(v) {
    var s = String(v || '').trim().toUpperCase();
    if (!s) return '';
    if (s === 'CON' || s === 'CON-UM') return 'CON-UM';
    if (s === 'KFG' || s === 'KFG-LLC') return 'KFG-LLC';
    return s;
  }

  function siteFilterOptions(values) {
    var canonValues = safeUnique((values || []).map(canonicalSiteValue));
    var optionSet = new Set();
    canonValues.forEach(function (site) {
      if (!site) return;
      optionSet.add(site);
      var prefix = site.split('-')[0];
      var groupMembers = canonValues.filter(function (v) { return v === prefix || v.indexOf(prefix + '-') === 0; });
      if (groupMembers.length > 1) optionSet.add(prefix + ' (All)');
    });
    return Array.from(optionSet).sort();
  }

  function matchesSiteSelection(siteValue, selectedSet) {
    if (!selectedSet || !selectedSet.size) return true;
    var site = canonicalSiteValue(siteValue);
    if (!site) return false;
    if (selectedSet.has(site)) return true;
    for (var val of selectedSet.values()) {
      if (String(val).endsWith(' (All)')) {
        var prefix = String(val).slice(0, -6);
        if (site === prefix || site.indexOf(prefix + '-') === 0) return true;
      }
    }
    return false;
  }

  function ensureStateSelection(stateArr, allOptions) {
    if (!stateArr || !stateArr.length) return allOptions.slice();
    var allowed = new Set(allOptions);
    var filtered = stateArr.filter(function (v) { return allowed.has(v); });
    return filtered.length ? filtered : allOptions.slice();
  }

  function populateMultiSelect(id, options, selectedValues) {
    var el = document.getElementById(id);
    if (!el) return;
    var selected = new Set(selectedValues || []);
    el.innerHTML = options.map(function (opt) {
      return '<option value="' + String(opt).replace(/"/g, '&quot;') + '"' + (selected.has(opt) ? ' selected' : '') + '>' + opt + '</option>';
    }).join('');
  }

  function readMultiSelect(id) {
    var el = document.getElementById(id);
    if (!el) return [];
    return Array.from(el.selectedOptions).map(function (o) { return o.value; });
  }

  function inMonthRange(key, from, to) {
    if (!key) return false;
    if (from && key < from) return false;
    if (to && key > to) return false;
    return true;
  }

  function inDateRange(dateObj, from, to) {
    if (!dateObj) return false;
    var d = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()).getTime();
    if (from) {
      var f = parseDashboardDate(from);
      if (f && d < new Date(f.getFullYear(), f.getMonth(), f.getDate()).getTime()) return false;
    }
    if (to) {
      var t = parseDashboardDate(to);
      if (t && d > new Date(t.getFullYear(), t.getMonth(), t.getDate()).getTime()) return false;
    }
    return true;
  }

  function maintenancePalette() {
    return ['#2563EB', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#0EA5C9', '#1D4ED8', '#64748B'];
  }

  function renderAnalyticsSummary(id, stats) {
    var el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = stats.map(function (s) {
      return '<div class="analytics-stat"><div class="analytics-stat-label">' + s.label + '</div><div class="analytics-stat-value">' + s.value + '</div></div>';
    }).join('');
  }

  function renderEmptyChartState(chartId, summaryId, message) {
    if (CHARTS[chartId]) { CHARTS[chartId].destroy(); delete CHARTS[chartId]; }
    var summary = document.getElementById(summaryId);
    if (summary) summary.innerHTML = '<div class="empty-state">' + message + '</div>';
  }

  function bindMaintenanceAnalyticsControls(kind, rerender) {
    var prefix = kind === 'po' ? 'maint-po' : 'maint-jobs';
    var sitesEl = document.getElementById(prefix + '-sites');
    var catsEl = document.getElementById(prefix + '-cats');
    var fromEl = document.getElementById(prefix + '-from');
    var toEl = document.getElementById(prefix + '-to');
    var resetEl = document.getElementById(prefix + '-reset');
    if (sitesEl) sitesEl.onchange = rerender;
    if (catsEl) catsEl.onchange = rerender;
    if (fromEl) fromEl.onchange = rerender;
    if (toEl) toEl.onchange = rerender;
    if (resetEl) {
      resetEl.onclick = function () {
        MAINT_ANALYTICS_STATE[kind] = { sites: [], cats: [], from: '', to: '' };
        rerender();
      };
    }
  }

  function renderMaintenanceJobsAnalytics(m) {
    var jobs = (m.jobs || []).map(function (j) {
      var d = parseDashboardDate(j.raisedDate || j.raisedDateObj || j.date);
      return Object.assign({}, j, { _month: monthKey(d) });
    }).filter(function (j) { return j.site && j._month; });

    if (!jobs.length) {
      renderEmptyChartState('c-maint-jobs-analytics', 'maint-jobs-summary', 'No jobs register rows available for analytics.');
      return;
    }

    var allSites = siteFilterOptions(jobs.map(function (j) { return j.site; }));
    var allCats = safeUnique(jobs.map(function (j) { return j.cat || 'Uncategorised'; }));
    MAINT_ANALYTICS_STATE.jobs.sites = ensureStateSelection(MAINT_ANALYTICS_STATE.jobs.sites, allSites);
    MAINT_ANALYTICS_STATE.jobs.cats = ensureStateSelection(MAINT_ANALYTICS_STATE.jobs.cats, allCats);

    populateMultiSelect('maint-jobs-sites', allSites, MAINT_ANALYTICS_STATE.jobs.sites);
    populateMultiSelect('maint-jobs-cats', allCats, MAINT_ANALYTICS_STATE.jobs.cats);

    var fromEl = document.getElementById('maint-jobs-from');
    var toEl = document.getElementById('maint-jobs-to');
    if (fromEl) fromEl.value = MAINT_ANALYTICS_STATE.jobs.from || '';
    if (toEl) toEl.value = MAINT_ANALYTICS_STATE.jobs.to || '';

    var selectedSites = new Set(MAINT_ANALYTICS_STATE.jobs.sites);
    var selectedCats = new Set(MAINT_ANALYTICS_STATE.jobs.cats);
    var filtered = jobs.filter(function (j) {
      var cat = j.cat || 'Uncategorised';
      return matchesSiteSelection(j.site, selectedSites) && selectedCats.has(cat) && inMonthRange(j._month, MAINT_ANALYTICS_STATE.jobs.from, MAINT_ANALYTICS_STATE.jobs.to);
    });

    if (!filtered.length) {
      renderEmptyChartState('c-maint-jobs-analytics', 'maint-jobs-summary', 'No jobs matched the current filter selection.');
      bindMaintenanceAnalyticsControls('jobs', function () {
        MAINT_ANALYTICS_STATE.jobs.sites = readMultiSelect('maint-jobs-sites');
        MAINT_ANALYTICS_STATE.jobs.cats = readMultiSelect('maint-jobs-cats');
        MAINT_ANALYTICS_STATE.jobs.from = (document.getElementById('maint-jobs-from') || {}).value || '';
        MAINT_ANALYTICS_STATE.jobs.to = (document.getElementById('maint-jobs-to') || {}).value || '';
        renderMaintenanceJobsAnalytics(m);
      });
      return;
    }

    var months = safeUnique(filtered.map(function (j) { return j._month; }));
    var cats = Array.from(selectedCats);
    var palette = maintenancePalette();
    var datasets = cats.map(function (cat, idx) {
      return {
        label: cat,
        data: months.map(function (month) {
          return filtered.filter(function (j) { return j._month === month && (j.cat || 'Uncategorised') === cat; }).length;
        }),
        backgroundColor: palette[idx % palette.length],
        borderRadius: 4,
      };
    });

    renderAnalyticsSummary('maint-jobs-summary', [
      { label: 'Filtered Jobs', value: filtered.length },
      { label: 'Open Jobs', value: filtered.filter(function (j) { return ['open','in progress','scheduled','assigned'].indexOf(String(j.status || '').toLowerCase()) !== -1; }).length },
      { label: 'Completed', value: filtered.filter(function (j) { return String(j.status || '').toLowerCase() === 'completed'; }).length },
      { label: 'Est. Cost', value: faed(filtered.reduce(function (s, j) { return s + (Number(j.cost) || 0); }, 0)) },
    ]);

    mkChart('c-maint-jobs-analytics', {
      type: 'bar',
      data: { labels: months.map(monthLabel), datasets: datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: true, position: 'top' } },
        scales: {
          x: { stacked: false, grid: { display: false } },
          y: { stacked: false, beginAtZero: true, ticks: { precision: 0 }, grid: { color: '#F1F5F9' } }
        }
      }
    });

    bindMaintenanceAnalyticsControls('jobs', function () {
      MAINT_ANALYTICS_STATE.jobs.sites = readMultiSelect('maint-jobs-sites');
      MAINT_ANALYTICS_STATE.jobs.cats = readMultiSelect('maint-jobs-cats');
      MAINT_ANALYTICS_STATE.jobs.from = (document.getElementById('maint-jobs-from') || {}).value || '';
      MAINT_ANALYTICS_STATE.jobs.to = (document.getElementById('maint-jobs-to') || {}).value || '';
      renderMaintenanceJobsAnalytics(m);
    });
  }

  function renderMaintenancePOAnalytics(m) {
    var rows = (m.poCosts || []).map(function (r) {
      var d = parseDashboardDate(r.date || r.dateObj);
      return Object.assign({}, r, { _month: monthKey(d), _cat: r.category || 'Uncategorised' });
    }).filter(function (r) { return r.site && r._month; });

    if (!rows.length) {
      renderEmptyChartState('c-maint-po-analytics', 'maint-po-summary', 'No MaintenancePOcost rows were found. Confirm the sheet name is exactly MaintenancePOcost and that it contains site, category, amount and date columns.');
      return;
    }

    var allSites = siteFilterOptions(rows.map(function (r) { return r.site; }));
    var allCats = safeUnique(rows.map(function (r) { return r._cat; }));
    MAINT_ANALYTICS_STATE.po.sites = ensureStateSelection(MAINT_ANALYTICS_STATE.po.sites, allSites);
    MAINT_ANALYTICS_STATE.po.cats = ensureStateSelection(MAINT_ANALYTICS_STATE.po.cats, allCats);

    populateMultiSelect('maint-po-sites', allSites, MAINT_ANALYTICS_STATE.po.sites);
    populateMultiSelect('maint-po-cats', allCats, MAINT_ANALYTICS_STATE.po.cats);

    var fromEl = document.getElementById('maint-po-from');
    var toEl = document.getElementById('maint-po-to');
    if (fromEl) fromEl.value = MAINT_ANALYTICS_STATE.po.from || '';
    if (toEl) toEl.value = MAINT_ANALYTICS_STATE.po.to || '';

    var selectedSites = new Set(MAINT_ANALYTICS_STATE.po.sites);
    var selectedCats = new Set(MAINT_ANALYTICS_STATE.po.cats);
    var filtered = rows.filter(function (r) {
      return matchesSiteSelection(r.site, selectedSites) && selectedCats.has(r._cat) && inMonthRange(r._month, MAINT_ANALYTICS_STATE.po.from, MAINT_ANALYTICS_STATE.po.to);
    });

    if (!filtered.length) {
      renderEmptyChartState('c-maint-po-analytics', 'maint-po-summary', 'No maintenance expense rows matched the current filter selection.');
      bindMaintenanceAnalyticsControls('po', function () {
        MAINT_ANALYTICS_STATE.po.sites = readMultiSelect('maint-po-sites');
        MAINT_ANALYTICS_STATE.po.cats = readMultiSelect('maint-po-cats');
        MAINT_ANALYTICS_STATE.po.from = (document.getElementById('maint-po-from') || {}).value || '';
        MAINT_ANALYTICS_STATE.po.to = (document.getElementById('maint-po-to') || {}).value || '';
        renderMaintenancePOAnalytics(m);
      });
      return;
    }

    var months = safeUnique(filtered.map(function (r) { return r._month; }));
    var cats = Array.from(selectedCats);
    var palette = maintenancePalette();
    var datasets = cats.map(function (cat, idx) {
      return {
        label: cat,
        data: months.map(function (month) {
          return filtered.filter(function (r) { return r._month === month && r._cat === cat; }).reduce(function (s, r) { return s + (Number(r.amount) || 0); }, 0);
        }),
        backgroundColor: palette[idx % palette.length],
        borderRadius: 4,
      };
    });

    renderAnalyticsSummary('maint-po-summary', [
      { label: 'Filtered POs', value: filtered.length },
      { label: 'Total Spend', value: faed(filtered.reduce(function (s, r) { return s + (Number(r.amount) || 0); }, 0)) },
      { label: 'Avg. PO', value: faed(filtered.length ? filtered.reduce(function (s, r) { return s + (Number(r.amount) || 0); }, 0) / filtered.length : 0) },
      { label: 'Vendors', value: safeUnique(filtered.map(function (r) { return r.vendor; })).length },
    ]);

    mkChart('c-maint-po-analytics', {
      type: 'bar',
      data: { labels: months.map(monthLabel), datasets: datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: true, position: 'top' } },
        scales: {
          x: { stacked: false, grid: { display: false } },
          y: { stacked: false, beginAtZero: true, grid: { color: '#F1F5F9' } }
        }
      }
    });

    bindMaintenanceAnalyticsControls('po', function () {
      MAINT_ANALYTICS_STATE.po.sites = readMultiSelect('maint-po-sites');
      MAINT_ANALYTICS_STATE.po.cats = readMultiSelect('maint-po-cats');
      MAINT_ANALYTICS_STATE.po.from = (document.getElementById('maint-po-from') || {}).value || '';
      MAINT_ANALYTICS_STATE.po.to = (document.getElementById('maint-po-to') || {}).value || '';
      renderMaintenancePOAnalytics(m);
    });
  }

  // --------------------------------------------------------------------------
  //  Build: Maintenance
  // --------------------------------------------------------------------------

  function buildMaintenance() {
    const D = DATA;
    const m = (D && D.maintenance) ? D.maintenance : DEMO.maintenance;

    const kpiEl = document.getElementById('k-maint');
    if (kpiEl) {
      kpiEl.innerHTML = [
        kpiCard('Open Jobs', m.kpis.openJobs, '', '#2563EB'),
        kpiCard('Completed MTD', m.kpis.completedMTD, '', '#10B981'),
        kpiCard('Overdue', m.kpis.overdueJobs, '', '#EF4444'),
        kpiCard('Avg Response', m.kpis.avgResponseHrs + 'h', '', '#F59E0B'),
      ].join('');
    }

    const siteCfg = getCfg('c-maint-site');
    const siteLabels = Object.keys(m.bySite || {});
    const siteData = Object.values(m.bySite || {});
    mkChart('c-maint-site', {
      type: 'bar',
      data: {
        labels: siteLabels,
        datasets: [{ label: 'Open Jobs', data: siteData, backgroundColor: siteCfg ? siteCfg.colors[0] : '#2563EB', borderRadius: 4 }],
      },
      options: noGridOpts({ indexAxis: siteCfg ? siteCfg.indexAxis : 'y' }),
    });

    const catCfg = getCfg('c-maint-cat');
    mkChart('c-maint-cat', {
      type: 'doughnut',
      data: {
        labels: Object.keys(m.byCat || {}),
        datasets: [{ data: Object.values(m.byCat || {}), backgroundColor: catCfg ? catCfg.colors : ['#2563EB', '#EF4444', '#F59E0B', '#10B981'] }],
      },
      options: Object.assign(doughnutOpts(catCfg ? catCfg.legendPosition : 'right'), { cutout: catCfg ? catCfg.cutout : '60%' }),
    });

    const tbl = document.getElementById('t-maint-jobs');
    if (tbl) {
      let html = '';
      (m.jobs || []).forEach(j => {
        html += '<tr><td>' + j.id + '</td><td>' + j.site + '</td><td>' + j.desc + '</td><td>' + j.cat + '</td><td>' + pill(j.priority) + '</td><td>' + pill(j.status) + '</td><td>' + fn(j.cost) + '</td></tr>';
      });
      tbl.innerHTML = html;
    }

    renderMaintenanceJobsAnalytics(m);
    renderMaintenancePOAnalytics(m);
  }

  function bindCapexCategoryControls(rerender) {
    var sitesEl = document.getElementById('capex-cat-sites');
    var catsEl = document.getElementById('capex-cat-cats');
    var resetEl = document.getElementById('capex-cat-reset');
    if (sitesEl) sitesEl.onchange = rerender;
    if (catsEl) catsEl.onchange = rerender;
    if (resetEl) resetEl.onclick = function () {
      CAPEX_CATEGORY_STATE.sites = [];
      CAPEX_CATEGORY_STATE.cats = [];
      rerender();
    };
  }

  function renderCapexCategoryBudget(c) {
    var projectRows = c.projects || [];
    var categoryRows = c.categories || [];
    // Pull nurseries from both the Capex Register rows (column C) and the Projects Register
    var allSites = siteFilterOptions(
      categoryRows.map(function (r) { return r.nursery; }).concat(
        projectRows.map(function (r) { return r.nursery; })
      )
    );
    var allCats = safeUnique(categoryRows.map(function (r) { return r.category; }).concat(projectRows.map(function (r) { return r.category; })));
    CAPEX_CATEGORY_STATE.sites = ensureStateSelection(CAPEX_CATEGORY_STATE.sites, allSites);
    CAPEX_CATEGORY_STATE.cats = ensureStateSelection(CAPEX_CATEGORY_STATE.cats, allCats);
    populateMultiSelect('capex-cat-sites', allSites, CAPEX_CATEGORY_STATE.sites);
    populateMultiSelect('capex-cat-cats', allCats, CAPEX_CATEGORY_STATE.cats);

    var selectedSites = new Set(CAPEX_CATEGORY_STATE.sites || []);
    var selectedCats = new Set(CAPEX_CATEGORY_STATE.cats || []);
    var rows = [];
    var usingNurseryFilter = selectedSites.size && selectedSites.size !== allSites.length;
    if (usingNurseryFilter) {
      var grouped = {};
      projectRows.filter(function (r) {
        return matchesSiteSelection(r.nursery, selectedSites) && (!selectedCats.size || selectedCats.has(r.category));
      }).forEach(function (r) {
        var cat = r.category || 'Uncategorised';
        if (!grouped[cat]) grouped[cat] = { category: cat, budget: 0, spent: 0, committed: 0, remaining: 0, pct: 0, rag: '' };
        grouped[cat].budget += Number(r.budget) || 0;
        grouped[cat].spent += Number(r.spent) || 0;
        grouped[cat].committed += Number(r.committed) || 0;
        grouped[cat].remaining += Number(r.remaining) || 0;
      });
      rows = Object.keys(grouped).sort().map(function (k) {
        var row = grouped[k];
        row.pct = row.budget ? ((row.spent + row.committed) / row.budget * 100) : 0;
        row.rag = row.remaining < 0 ? 'Red' : row.pct > 80 ? 'Amber' : 'Green';
        return row;
      });
    } else {
      rows = categoryRows.filter(function (r) { return !selectedCats.size || selectedCats.has(r.category); });
    }

    renderAnalyticsSummary('capex-cat-summary', [
      { label: 'Tasks Completed', value: rows.length },
      { label: 'Spent', value: faed(rows.reduce(function (s, r) { return s + (Number(r.spent) || 0); }, 0)) },
    ]);

    // Chart: Total Spent per Category — with data labels on top of bars
    const spentByCat = {};
    rows.forEach(function (r) { if (r.category) spentByCat[r.category] = (spentByCat[r.category] || 0) + (Number(r.spent) || 0); });
    const catEntries = Object.entries(spentByCat).sort(function (a, b) { return b[1] - a[1]; });
    mkChart('c-capex-cat-spent', {
      type: 'bar',
      data: {
        labels: catEntries.map(function (e) { return e[0]; }),
        datasets: [{ label: 'Spent (AED)', data: catEntries.map(function (e) { return e[1]; }), backgroundColor: '#2563EB', borderRadius: 4 }],
      },
      options: noGridOpts({ plugins: { legend: { display: false }, kfgBarLabels: { enabled: true } } }),
    });

    // Chart: Total Spent per Nursery — always from projectRows so nursery filter works reliably
    const filteredForNursery = projectRows.filter(function (r) {
      var siteMatch = !usingNurseryFilter || matchesSiteSelection(r.nursery, selectedSites);
      var catMatch  = !selectedCats.size   || selectedCats.has(r.category);
      return siteMatch && catMatch;
    });
    const spentByNursery = {};
    filteredForNursery.forEach(function (r) {
      if (r.nursery) spentByNursery[r.nursery] = (spentByNursery[r.nursery] || 0) + (Number(r.spent) || 0);
    });
    const nurseryEntries = Object.entries(spentByNursery).sort(function (a, b) { return b[1] - a[1]; });
    mkChart('c-capex-cat-nursery', {
      type: 'bar',
      data: {
        labels: nurseryEntries.map(function (e) { return e[0]; }),
        datasets: [{ label: 'Spent (AED)', data: nurseryEntries.map(function (e) { return e[1]; }), backgroundColor: '#8B5CF6', borderRadius: 4 }],
      },
      options: noGridOpts({ plugins: { legend: { display: false }, kfgBarLabels: { enabled: true } } }),
    });

    bindCapexCategoryControls(function () {
      CAPEX_CATEGORY_STATE.sites = readMultiSelect('capex-cat-sites');
      CAPEX_CATEGORY_STATE.cats = readMultiSelect('capex-cat-cats');
      renderCapexCategoryBudget(c);
    });
  }

  function bindProcurementSupplierControls(rerender) {
    ['proc-supplier-names','proc-supplier-cats','proc-supplier-from','proc-supplier-to'].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.onchange = rerender;
    });
    var resetEl = document.getElementById('proc-supplier-reset');
    if (resetEl) resetEl.onclick = function () {
      PROC_SUPPLIER_STATE.suppliers = [];
      PROC_SUPPLIER_STATE.cats = [];
      PROC_SUPPLIER_STATE.from = '';
      PROC_SUPPLIER_STATE.to = '';
      rerender();
    };
  }

  function renderProcurementSuppliers(pr) {
    // Use structured rows for filtering, raw rows for generic table display
    var rows = (pr.suppliers || []).slice();
    var rawRows = (pr.rawSuppliers || []).filter(function (r) {
      return Object.values(r).some(function (v) { return v !== null && v !== undefined && v !== ''; });
    });

    var allSuppliers = safeUnique(rows.map(function (r) { return r.supplier; }).filter(Boolean));
    var allCats = safeUnique(rows.map(function (r) { return r.category; }).filter(Boolean));
    // If structured mapping yielded no supplier names, skip filter dropdowns
    PROC_SUPPLIER_STATE.suppliers = ensureStateSelection(PROC_SUPPLIER_STATE.suppliers, allSuppliers);
    PROC_SUPPLIER_STATE.cats = ensureStateSelection(PROC_SUPPLIER_STATE.cats, allCats);
    populateMultiSelect('proc-supplier-names', allSuppliers, PROC_SUPPLIER_STATE.suppliers);
    populateMultiSelect('proc-supplier-cats', allCats, PROC_SUPPLIER_STATE.cats);
    var fromEl = document.getElementById('proc-supplier-from');
    var toEl = document.getElementById('proc-supplier-to');
    if (fromEl) fromEl.value = PROC_SUPPLIER_STATE.from || '';
    if (toEl) toEl.value = PROC_SUPPLIER_STATE.to || '';

    renderAnalyticsSummary('proc-supplier-summary', [
      { label: 'Records', value: rawRows.length },
      { label: 'Annual Spend', value: faed(rows.reduce(function (s, r) { return s + (Number(r.annualSpend) || 0); }, 0)) },
      { label: 'Active', value: rows.filter(function (r) { return String(r.status || '').toLowerCase() === 'active'; }).length },
    ]);

    // Generic table — render all columns from raw Excel data
    var tbl  = document.getElementById('t-proc-suppliers');
    var thead = document.getElementById('t-proc-suppliers-head');
    if (tbl && rawRows.length) {
      var cols = Object.keys(rawRows[0]);
      if (thead) {
        thead.innerHTML = '<tr>' + cols.map(function (c) { return '<th scope="col">' + c + '</th>'; }).join('') + '</tr>';
      }
      tbl.innerHTML = rawRows.map(function (r) {
        return '<tr>' + cols.map(function (c) {
          var v = r[c];
          if (v === null || v === undefined) return '<td></td>';
          if (typeof v === 'number' && !isNaN(v)) return '<td>' + Math.round(v).toLocaleString('en-AE') + '</td>';
          return '<td>' + v + '</td>';
        }).join('') + '</tr>';
      }).join('');
    }
    bindProcurementSupplierControls(function () {
      PROC_SUPPLIER_STATE.suppliers = readMultiSelect('proc-supplier-names');
      PROC_SUPPLIER_STATE.cats = readMultiSelect('proc-supplier-cats');
      PROC_SUPPLIER_STATE.from = (document.getElementById('proc-supplier-from') || {}).value || '';
      PROC_SUPPLIER_STATE.to = (document.getElementById('proc-supplier-to') || {}).value || '';
      renderProcurementSuppliers(pr);
    });
  }

  function bindMAMilestoneControls(rerender) {
    ['ma-projects','ma-from','ma-to'].forEach(function (id) { var el = document.getElementById(id); if (el) el.onchange = rerender; });
    var resetEl = document.getElementById('ma-reset');
    if (resetEl) resetEl.onclick = function () { MA_MILESTONE_STATE.projects=[]; MA_MILESTONE_STATE.from=''; MA_MILESTONE_STATE.to=''; rerender(); };
  }

  function renderMAMilestones(ma) {
    var rows = (ma.milestones || []).map(function (r) { return Object.assign({}, r, { _projectLabel: (r.dealId || '') + (r.target ? ' — ' + r.target : '') }); });
    var allProjects = safeUnique(rows.map(function (r) { return r._projectLabel; }));
    MA_MILESTONE_STATE.projects = ensureStateSelection(MA_MILESTONE_STATE.projects, allProjects);
    populateMultiSelect('ma-projects', allProjects, MA_MILESTONE_STATE.projects);
    var fromEl = document.getElementById('ma-from'); var toEl = document.getElementById('ma-to');
    if (fromEl) fromEl.value = MA_MILESTONE_STATE.from || ''; if (toEl) toEl.value = MA_MILESTONE_STATE.to || '';
    var selectedProjects = new Set(MA_MILESTONE_STATE.projects || []);
    var filtered = rows.filter(function (r) {
      return (!selectedProjects.size || selectedProjects.has(r._projectLabel)) && (!r.dueDateObj || inDateRange(r.dueDateObj, MA_MILESTONE_STATE.from, MA_MILESTONE_STATE.to));
    });
    renderAnalyticsSummary('ma-summary', [
      { label: 'Milestones', value: filtered.length },
      { label: 'Completed', value: filtered.filter(function (r) { return String(r.status || '').toLowerCase() === 'completed'; }).length },
      { label: 'Open', value: filtered.filter(function (r) { return ['open','in progress','pending'].indexOf(String(r.status || '').toLowerCase()) !== -1; }).length },
      { label: 'Critical', value: filtered.filter(function (r) { return String(r.priority || '').toLowerCase() === 'critical'; }).length }
    ]);
    var tbl = document.getElementById('t-ma-milestones');
    if (tbl) tbl.innerHTML = filtered.map(function (r) { return '<tr><td>' + (r.dealId||'') + '</td><td>' + (r.target||'') + '</td><td>' + (r.milestone||'') + '</td><td>' + (r.dueDate||'') + '</td><td>' + (r.owner||'') + '</td><td>' + pill(r.status) + '</td><td>' + pill(r.priority) + '</td><td>' + (r.notes||'') + '</td></tr>'; }).join('');
    bindMAMilestoneControls(function () {
      MA_MILESTONE_STATE.projects = readMultiSelect('ma-projects');
      MA_MILESTONE_STATE.from = (document.getElementById('ma-from') || {}).value || '';
      MA_MILESTONE_STATE.to = (document.getElementById('ma-to') || {}).value || '';
      renderMAMilestones(ma);
    });
  }

  function bindOtherActionControls(rerender) {
    ['other-projects','other-from','other-to'].forEach(function (id) { var el = document.getElementById(id); if (el) el.onchange = rerender; });
    var resetEl = document.getElementById('other-reset');
    if (resetEl) resetEl.onclick = function () { OTHER_ACTION_STATE.projects=[]; OTHER_ACTION_STATE.from=''; OTHER_ACTION_STATE.to=''; rerender(); };
  }

  function renderOtherActions(o) {
    var rows = (o.actions || []).map(function (r) { return Object.assign({}, r, { _projectLabel: (r.projectId || '') + (r.project ? ' — ' + r.project : '') }); });
    var allProjects = safeUnique(rows.map(function (r) { return r._projectLabel; }));
    OTHER_ACTION_STATE.projects = ensureStateSelection(OTHER_ACTION_STATE.projects, allProjects);
    populateMultiSelect('other-projects', allProjects, OTHER_ACTION_STATE.projects);
    var fromEl = document.getElementById('other-from'); var toEl = document.getElementById('other-to');
    if (fromEl) fromEl.value = OTHER_ACTION_STATE.from || ''; if (toEl) toEl.value = OTHER_ACTION_STATE.to || '';
    var selectedProjects = new Set(OTHER_ACTION_STATE.projects || []);
    var filtered = rows.filter(function (r) {
      return (!selectedProjects.size || selectedProjects.has(r._projectLabel)) && (!r.dueDateObj || inDateRange(r.dueDateObj, OTHER_ACTION_STATE.from, OTHER_ACTION_STATE.to));
    });
    renderAnalyticsSummary('other-summary', [
      { label: 'Actions', value: filtered.length },
      { label: 'Open', value: filtered.filter(function (r) { return String(r.status || '').toLowerCase() === 'open'; }).length },
      { label: 'In Progress', value: filtered.filter(function (r) { return String(r.status || '').toLowerCase() === 'in progress'; }).length },
      { label: 'Completed', value: filtered.filter(function (r) { return String(r.status || '').toLowerCase() === 'completed'; }).length }
    ]);
    var tbl = document.getElementById('t-other-actions');
    if (tbl) tbl.innerHTML = filtered.map(function (r) { return '<tr><td>' + (r.projectId||'') + '</td><td>' + (r.project||'') + '</td><td>' + (r.action||'') + '</td><td>' + (r.owner||'') + '</td><td>' + (r.raisedDate||'') + '</td><td>' + (r.dueDate||'') + '</td><td>' + pill(r.status) + '</td><td>' + pill(r.priority) + '</td><td>' + (r.notes||'') + '</td></tr>'; }).join('');
    bindOtherActionControls(function () {
      OTHER_ACTION_STATE.projects = readMultiSelect('other-projects');
      OTHER_ACTION_STATE.from = (document.getElementById('other-from') || {}).value || '';
      OTHER_ACTION_STATE.to = (document.getElementById('other-to') || {}).value || '';
      renderOtherActions(o);
    });
  }

  // --------------------------------------------------------------------------
  //  Build: Capex
  // --------------------------------------------------------------------------

  function buildCapex() {
    const D = DATA;
    const c = (D && D.capex) ? D.capex : DEMO.capex;

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

    renderCapexCategoryBudget(c);
  }

  // --------------------------------------------------------------------------
  //  Build: Projects
  // --------------------------------------------------------------------------

  function buildProjects() {
    const D = DATA;
    const p = (D && D.projects) ? D.projects : DEMO.projects;

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

    renderPMMilestones(p);
  }

  function renderPMMilestones(p) {
    var rows = (p.milestones || []).map(function (r) {
      return Object.assign({}, r, { _projectLabel: (r.projectId || '') + (r.project ? ' — ' + r.project : '') });
    });
    var allProjects = safeUnique(rows.map(function (r) { return r._projectLabel; }));
    PM_MILESTONE_STATE.projects = ensureStateSelection(PM_MILESTONE_STATE.projects, allProjects);
    populateMultiSelect('pm-milestone-projects', allProjects, PM_MILESTONE_STATE.projects);
    var fromEl = document.getElementById('pm-milestone-from');
    var toEl   = document.getElementById('pm-milestone-to');
    if (fromEl) fromEl.value = PM_MILESTONE_STATE.from || '';
    if (toEl)   toEl.value   = PM_MILESTONE_STATE.to   || '';
    var selectedProjects = new Set(PM_MILESTONE_STATE.projects || []);
    var filtered = rows.filter(function (r) {
      return (!selectedProjects.size || selectedProjects.has(r._projectLabel)) &&
             (!r.dueDateObj || inDateRange(r.dueDateObj, PM_MILESTONE_STATE.from, PM_MILESTONE_STATE.to));
    });
    renderAnalyticsSummary('pm-milestone-summary', [
      { label: 'Milestones', value: filtered.length },
      { label: 'Completed', value: filtered.filter(function (r) { return String(r.status || '').toLowerCase() === 'completed'; }).length },
      { label: 'In Progress', value: filtered.filter(function (r) { return String(r.status || '').toLowerCase() === 'in progress'; }).length },
      { label: 'Open', value: filtered.filter(function (r) { return String(r.status || '').toLowerCase() === 'open'; }).length },
    ]);
    var tbl = document.getElementById('t-pm-milestones');
    if (tbl) {
      tbl.innerHTML = filtered.map(function (r) {
        return '<tr><td>' + (r.projectId||'') + '</td><td>' + (r.project||'') + '</td><td>' + (r.milestone||'') + '</td><td>' + (r.dueDate||'') + '</td><td>' + (r.owner||'') + '</td><td>' + pill(r.status) + '</td><td>' + pill(r.priority) + '</td><td>' + (r.notes||'') + '</td></tr>';
      }).join('');
    }
    // bind controls
    ['pm-milestone-projects','pm-milestone-from','pm-milestone-to'].forEach(function (id) {
      var el = document.getElementById(id); if (el) el.onchange = function () {
        PM_MILESTONE_STATE.projects = readMultiSelect('pm-milestone-projects');
        PM_MILESTONE_STATE.from = (document.getElementById('pm-milestone-from') || {}).value || '';
        PM_MILESTONE_STATE.to   = (document.getElementById('pm-milestone-to')   || {}).value || '';
        renderPMMilestones(p);
      };
    });
    var resetEl = document.getElementById('pm-milestone-reset');
    if (resetEl) resetEl.onclick = function () { PM_MILESTONE_STATE.projects=[]; PM_MILESTONE_STATE.from=''; PM_MILESTONE_STATE.to=''; renderPMMilestones(p); };
  }


  function renderCategorySpendBreakdown(id, rows, labelKey, valueKey) {
    var el = document.getElementById(id);
    if (!el) return;
    var map = {};
    (rows || []).forEach(function (r) {
      var label = String(r[labelKey] || 'Uncategorised').trim() || 'Uncategorised';
      if (!map[label]) map[label] = { count: 0, spend: 0 };
      map[label].count += 1;
      map[label].spend += Number(r[valueKey]) || 0;
    });
    var entries = Object.keys(map).sort().map(function (k) {
      return { label: k, count: map[k].count, spend: map[k].spend };
    });
    if (!entries.length) {
      el.innerHTML = '';
      return;
    }
    el.innerHTML = '<div class="analytics-breakdown-title">Category Spend Breakdown</div>' +
      '<div class="analytics-breakdown-grid">' +
      entries.map(function (e) {
        return '<div class="analytics-breakdown-item">' +
          '<div class="analytics-breakdown-label">' + e.label + '</div>' +
          '<div class="analytics-breakdown-meta">' + e.count + ' PO' + (e.count === 1 ? '' : 's') + '</div>' +
          '<div class="analytics-breakdown-value">' + faed(e.spend) + '</div>' +
        '</div>';
      }).join('') +
      '</div>';
  }

  function bindProcurementAnalyticsControls(rerender) {
    var sitesEl = document.getElementById('proc-sites');
    var catsEl = document.getElementById('proc-cats');
    var fromEl = document.getElementById('proc-from');
    var toEl = document.getElementById('proc-to');
    var resetEl = document.getElementById('proc-reset');
    if (sitesEl) sitesEl.onchange = rerender;
    if (catsEl) catsEl.onchange = rerender;
    if (fromEl) fromEl.onchange = rerender;
    if (toEl) toEl.onchange = rerender;
    if (resetEl) {
      resetEl.onclick = function () {
        PROC_ANALYTICS_STATE.sites = [];
        PROC_ANALYTICS_STATE.cats = [];
        PROC_ANALYTICS_STATE.from = '';
        PROC_ANALYTICS_STATE.to = '';
        rerender();
      };
    }
  }

  function renderProcurementAnalytics(pr) {
    var rows = (pr.pos || []).map(function (r) {
      var d = parseDashboardDate(r.raisedDate || r.deliveryDate || r.date);
      return Object.assign({}, r, {
        _month: monthKey(d),
        _site: canonicalSiteValue(r.nursery || r.site || r.ref || r.po),
        _cat: String(r.category || r.dept || 'Uncategorised').trim() || 'Uncategorised'
      });
    }).filter(function (r) { return r._site && r._month; });

    if (!rows.length) {
      renderEmptyChartState('c-proc-analytics', 'proc-summary', 'No procurement rows available for analytics.');
      renderCategorySpendBreakdown('proc-category-breakdown', [], 'category', 'value');
      return;
    }

    var allSites = siteFilterOptions(rows.map(function (r) { return r._site; }));
    var allCats = safeUnique(rows.map(function (r) { return r._cat; }));
    PROC_ANALYTICS_STATE.sites = ensureStateSelection(PROC_ANALYTICS_STATE.sites, allSites);
    PROC_ANALYTICS_STATE.cats = ensureStateSelection(PROC_ANALYTICS_STATE.cats, allCats);

    populateMultiSelect('proc-sites', allSites, PROC_ANALYTICS_STATE.sites);
    populateMultiSelect('proc-cats', allCats, PROC_ANALYTICS_STATE.cats);

    var fromEl = document.getElementById('proc-from');
    var toEl = document.getElementById('proc-to');
    if (fromEl) fromEl.value = PROC_ANALYTICS_STATE.from || '';
    if (toEl) toEl.value = PROC_ANALYTICS_STATE.to || '';

    var selectedSites = new Set(PROC_ANALYTICS_STATE.sites);
    var selectedCats = new Set(PROC_ANALYTICS_STATE.cats);
    var filtered = rows.filter(function (r) {
      return matchesSiteSelection(r._site, selectedSites) && selectedCats.has(r._cat) && inMonthRange(r._month, PROC_ANALYTICS_STATE.from, PROC_ANALYTICS_STATE.to);
    });

    if (!filtered.length) {
      renderEmptyChartState('c-proc-analytics', 'proc-summary', 'No procurement rows matched the current filter selection.');
      renderCategorySpendBreakdown('proc-category-breakdown', [], 'category', 'value');
      bindProcurementAnalyticsControls(function () {
        PROC_ANALYTICS_STATE.sites = readMultiSelect('proc-sites');
        PROC_ANALYTICS_STATE.cats = readMultiSelect('proc-cats');
        PROC_ANALYTICS_STATE.from = (document.getElementById('proc-from') || {}).value || '';
        PROC_ANALYTICS_STATE.to = (document.getElementById('proc-to') || {}).value || '';
        renderProcurementAnalytics(pr);
      });
      return;
    }

    var months = safeUnique(filtered.map(function (r) { return r._month; }));
    var cats = Array.from(selectedCats);
    var palette = maintenancePalette();
    var datasets = cats.map(function (cat, idx) {
      return {
        label: cat,
        data: months.map(function (month) {
          return filtered.filter(function (r) { return r._month === month && r._cat === cat; }).reduce(function (s, r) { return s + (Number(r.value) || 0); }, 0);
        }),
        backgroundColor: palette[idx % palette.length],
        borderRadius: 4
      };
    });

    renderAnalyticsSummary('proc-summary', [
      { label: 'PO Created', value: filtered.length },
      { label: 'Total Spend', value: faed(filtered.reduce(function (s, r) { return s + (Number(r.value) || 0); }, 0)) },
      { label: 'Nurseries', value: safeUnique(filtered.map(function (r) { return r._site; })).length },
      { label: 'Suppliers', value: safeUnique(filtered.map(function (r) { return r.supplier; })).length }
    ]);

    renderCategorySpendBreakdown('proc-category-breakdown', filtered.map(function (r) {
      return { category: r._cat, value: r.value };
    }), 'category', 'value');

    mkChart('c-proc-analytics', {
      type: 'bar',
      data: { labels: months.map(monthLabel), datasets: datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: true, position: 'top' } },
        scales: {
          x: { stacked: false, grid: { display: false } },
          y: { stacked: false, beginAtZero: true, grid: { color: '#F1F5F9' } }
        }
      }
    });

    bindProcurementAnalyticsControls(function () {
      PROC_ANALYTICS_STATE.sites = readMultiSelect('proc-sites');
      PROC_ANALYTICS_STATE.cats = readMultiSelect('proc-cats');
      PROC_ANALYTICS_STATE.from = (document.getElementById('proc-from') || {}).value || '';
      PROC_ANALYTICS_STATE.to = (document.getElementById('proc-to') || {}).value || '';
      renderProcurementAnalytics(pr);
    });
  }

  // --------------------------------------------------------------------------
  //  Build: Procurement
  // --------------------------------------------------------------------------

  function buildProcurement() {
    const D = DATA;
    const pr = (D && D.procurement) ? D.procurement : DEMO.procurement;

    // KPIs — only Spend YTD
    const kpiEl = document.getElementById('k-proc');
    if (kpiEl) {
      kpiEl.innerHTML = kpiCard('Spend YTD', faed(pr.kpis.spendYTD), '', '#10B981');
    }

    // Spend by department bar — sorted highest to lowest, full width
    const spCfg = getCfg('c-proc-spend');
    const deptEntries = Object.entries(pr.byDept || {}).sort(function (a, b) { return b[1] - a[1]; });
    mkChart('c-proc-spend', {
      type: 'bar',
      data: {
        labels: deptEntries.map(function (e) { return e[0]; }),
        datasets: [{ label: 'Spend (AED)', data: deptEntries.map(function (e) { return e[1]; }), backgroundColor: spCfg ? spCfg.colors[0] : '#2563EB', borderRadius: 4 }],
      },
      options: noGridOpts(),
    });

    // Spend by Nursery bar — sorted highest to lowest
    const nurCfg = getCfg('c-proc-nursery');
    const nurseryEntries = Object.entries(pr.byNursery || {}).sort(function (a, b) { return b[1] - a[1]; });
    mkChart('c-proc-nursery', {
      type: 'bar',
      data: {
        labels: nurseryEntries.map(function (e) { return e[0]; }),
        datasets: [{ label: 'Spend (AED)', data: nurseryEntries.map(function (e) { return e[1]; }), backgroundColor: nurCfg ? nurCfg.colors[0] : '#8B5CF6', borderRadius: 4 }],
      },
      options: noGridOpts(),
    });

    const refHead = document.getElementById('t-proc-ref-label');
    if (refHead) refHead.textContent = pr.refLabel || 'PO Number';

    // PO table
    const tbl = document.getElementById('t-proc');
    if (tbl) {
      let html = '';
      pr.pos.forEach(r => {
        var refValue = r.ref || r.po || '';
        html += '<tr><td>' + refValue + '</td><td>' + r.supplier + '</td><td>' + r.desc + '</td><td>' + (r.category || r.dept || '') + '</td><td>' + fn(r.value) + '</td><td>' + pill(r.status) + '</td><td>' + (r.raisedDate || r.delivery || '') + '</td></tr>';
      });
      tbl.innerHTML = html;
    }

    renderProcurementAnalytics(pr);
    renderProcurementSuppliers(pr);
  }

  // --------------------------------------------------------------------------
  //  Build: IT
  // --------------------------------------------------------------------------

  function buildIT() {
    const D = DATA;
    const it = (D && D.it) ? D.it : DEMO.it;

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

    // Asset Register table — generic rendering based on available columns
    const assetTbl = document.getElementById('t-it-assets');
    const assetHead = document.getElementById('t-it-assets-head');
    if (assetTbl && it.assetRegister && it.assetRegister.length) {
      const cols = Object.keys(it.assetRegister[0]);
      if (assetHead) {
        assetHead.innerHTML = '<tr>' + cols.map(function (c) { return '<th scope="col">' + c + '</th>'; }).join('') + '</tr>';
      }
      assetTbl.innerHTML = it.assetRegister.map(function (r) {
        return '<tr>' + cols.map(function (c) { return '<td>' + (r[c] !== null && r[c] !== undefined ? r[c] : '') + '</td>'; }).join('') + '</tr>';
      }).join('');
    }
  }

  // --------------------------------------------------------------------------
  //  Build: M&A
  // --------------------------------------------------------------------------

  function buildMA() {
    const D = DATA;
    const ma = (D && D.ma) ? D.ma : DEMO.ma;

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

    renderMAMilestones(ma);
  }

  // --------------------------------------------------------------------------
  //  Build: Greenfield
  // --------------------------------------------------------------------------

  function buildGreenfield() {
    const D = DATA;
    const g = (D && D.greenfield) ? D.greenfield : DEMO.greenfield;

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

    // Pipeline Overview raw table — format numbers + rename/shorten column headers
    const pipeTbl  = document.getElementById('t-gf-pipeline');
    const pipeHead = document.getElementById('t-gf-pipeline-head');
    const pipeRows = g.rawPipeline || [];
    if (pipeTbl && pipeRows.length) {
      const cols = Object.keys(pipeRows[0]);

      // Column label overrides (match by lowercased substring)
      function pipeColLabel(c) {
        var lc = c.toLowerCase();
        if (lc.includes('site id') || lc === 'id' || lc === 'site no')         return 'ID';
        if (lc.includes('indoor') && lc.includes('sqft'))                        return 'Indoor (SQFT)';
        if (lc.includes('outdoor') && lc.includes('sqft'))                       return 'Outdoor (SQFT)';
        if (lc.includes('location') && lc.includes('city'))                      return 'City';
        if (lc.includes('opening'))                                               return 'Opening';
        return c; // keep original
      }

      // Cell value transform (strip "Location / " prefix from city cells)
      var locationColIdx = cols.findIndex(function (c) { return c.toLowerCase().includes('location') && c.toLowerCase().includes('city'); });

      if (pipeHead) {
        pipeHead.innerHTML = '<tr>' + cols.map(function (c, i) {
          var label = pipeColLabel(c);
          var isId = label === 'ID';
          return '<th scope="col" style="white-space:nowrap' + (isId ? ';width:60px;max-width:60px' : '') + '">' + label + '</th>';
        }).join('') + '</tr>';
      }
      pipeTbl.innerHTML = pipeRows.map(function (r) {
        return '<tr>' + cols.map(function (c, i) {
          var v = r[c];
          if (v === null || v === undefined) return '<td>-</td>';
          // Strip "Location / " or "Location/" prefix from city column
          if (i === locationColIdx && typeof v === 'string') {
            v = v.replace(/^Location\s*\/\s*/i, '').trim();
          }
          if (typeof v === 'number' && !isNaN(v)) {
            return '<td style="text-align:right;white-space:nowrap">' + Math.round(v).toLocaleString('en-AE') + '</td>';
          }
          return '<td style="white-space:nowrap">' + v + '</td>';
        }).join('') + '</tr>';
      }).join('');
    }
  }

  // --------------------------------------------------------------------------
  //  Build: Other Projects
  // --------------------------------------------------------------------------

  function buildOther() {
    const D = DATA;
    const o = (D && D.other) ? D.other : DEMO.other;

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

    renderOtherActions(o);
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
      // Accept transformed live data; fall back to demo for any null department
      var normalized = {};
      ['maintenance','capex','projects','procurement','it','ma','greenfield','other'].forEach(function(k) {
        var dept = data && data[k];
        // Use live data if it has a kpis object with at least one non-null key
        if (dept && dept.kpis && Object.keys(dept.kpis).length > 0) {
          normalized[k] = dept;
        } else {
          normalized[k] = null; // falls back to DEMO in build functions
        }
      });
      DATA = normalized;
      _isDemo = false;
      console.log('[KFGDashboard] Switched to live data:', normalized);
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

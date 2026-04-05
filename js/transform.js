/**
 * KFG Executive Dashboard — Data Transformation Module
 *
 * Converts raw SheetJS rows (fetched from SharePoint Excel files) into the
 * structured format expected by the Dashboard rendering module.
 *
 * Namespace: window.KFGTransform
 */
(function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Utility helpers
  // ---------------------------------------------------------------------------

  /** Safe number parsing — returns 0 for null / empty / NaN. */
  function num(v) {
    if (v === null || v === undefined || v === '') return 0;
    var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
    return isNaN(n) ? 0 : n;
  }

  /**
   * Get a cell value from a row by trying multiple column name variants.
   * Exact match first, then case-insensitive substring fallback.
   */
  function col(row) {
    var names = Array.prototype.slice.call(arguments, 1);
    var i, j, n, lname, keys, v;
    for (i = 0; i < names.length; i++) {
      n = names[i];
      if (row[n] !== undefined && row[n] !== null && row[n] !== '') return row[n];
    }
    keys = Object.keys(row);
    for (i = 0; i < names.length; i++) {
      lname = names[i].toLowerCase();
      for (j = 0; j < keys.length; j++) {
        if (keys[j].toLowerCase().includes(lname)) {
          v = row[keys[j]];
          if (v !== null && v !== undefined && v !== '') return v;
        }
      }
    }
    return null;
  }

  /**
   * Search KPI rows (two-column key→value format) for a metric name.
   * Column A contains the label, Column B the value.
   */
  function kpiVal(rows, name) {
    if (!rows || !rows.length) return null;
    var lname = name.toLowerCase();
    for (var i = 0; i < rows.length; i++) {
      var vals = Object.values(rows[i]).filter(function (v) {
        return v !== null && v !== undefined && v !== '';
      });
      if (vals.length < 1) continue;
      var key = String(vals[0]).toLowerCase();
      if (key.includes(lname)) {
        return vals.length > 1 ? vals[1] : null;
      }
    }
    return null;
  }

  /**
   * Format an Excel date serial or date string as DD/MM/YYYY.
   */
  function fmtDate(v) {
    if (!v) return '';
    if (typeof v === 'number' && v > 1000) {
      // Excel date serial (days since 1900-01-01)
      var d = new Date((v - 25569) * 86400 * 1000);
      return d.toLocaleDateString('en-GB');
    }
    return String(v);
  }

  function parseDateValue(v) {
    if (!v) return null;
    if (v instanceof Date && !isNaN(v.getTime())) return v;
    if (typeof v === 'number' && v > 1000) {
      return new Date((v - 25569) * 86400 * 1000);
    }
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

  // ---------------------------------------------------------------------------
  // Maintenance
  // ---------------------------------------------------------------------------

  function transformMaintenance(raw) {
    if (!raw) return null;
    var kpiRows   = raw.kpi     || [];
    var dataRows  = raw.data    || [];
    var siteRows  = raw.sites   || [];
    var poCostRows = raw.poCosts || [];

    var jobs = dataRows.map(function (r, idx) {
      var raisedRaw = col(r, 'Raised Date', 'Date Raised', 'Created Date', 'Date');
      var completedRaw = col(r, 'Completed Date', 'Closed Date', 'Resolved Date');
      return {
        id:           col(r, 'Job ID', 'ID', 'Job No', 'Job Number', 'Ref') || ('MNT-' + String(idx + 1).padStart(3, '0')),
        site:         col(r, 'Site / Nursery', 'Site', 'Location', 'Branch', 'Nursery') || '',
        desc:         col(r, 'Description', 'Job Description', 'Desc', 'Details', 'Work') || '',
        cat:          col(r, 'Category', 'Cat', 'Type', 'Job Type', 'Work Type') || '',
        priority:     col(r, 'Priority') || 'Medium',
        status:       col(r, 'Status', 'Job Status') || '',
        cost:         num(col(r, 'Est. Cost (£)', 'Est. Cost (AED)', 'Cost', 'Cost (AED)', 'Cost (£)', 'Amount', 'Value', 'Estimate')),
        raisedDate:   fmtDate(raisedRaw),
        completedDate: fmtDate(completedRaw),
        raisedDateObj: parseDateValue(raisedRaw),
        completedDateObj: parseDateValue(completedRaw),
        assignedTo:   col(r, 'Assigned To', 'Assigned', 'Technician', 'Owner') || '',
        notes:        col(r, 'Notes', 'Remarks') || '',
      };
    }).filter(function (j) { return j.site || j.desc; });

    var openJobs     = jobs.filter(function (j) {
      var s = (j.status || '').toLowerCase();
      return s === 'open' || s === 'in progress' || s === 'scheduled' || s === 'assigned';
    }).length;
    var completedMTD = jobs.filter(function (j) {
      return (j.status || '').toLowerCase() === 'completed';
    }).length;
    var overdueJobs  = jobs.filter(function (j) {
      var s = (j.status || '').toLowerCase();
      return s === 'overdue' || s === 'on hold';
    }).length;

    var avgResp = num(kpiVal(kpiRows, 'avg') || kpiVal(kpiRows, 'response') || kpiVal(kpiRows, 'hours'));
    var ragKpi  = kpiVal(kpiRows, 'rag') || kpiVal(kpiRows, 'overall');
    var rag     = ragKpi ? String(ragKpi) : (overdueJobs > 3 ? 'Red' : overdueJobs > 0 ? 'Amber' : 'Green');

    var byCat = {};
    jobs.forEach(function (j) {
      if (j.cat) byCat[j.cat] = (byCat[j.cat] || 0) + 1;
    });
    if (!Object.keys(byCat).length) byCat = { 'General': openJobs || jobs.length };

    var bySite = {};
    if (siteRows.length) {
      siteRows.forEach(function (r) {
        var siteName = col(r, 'Site', 'Site Name', 'Location', 'Branch') || '';
        var cnt = num(col(r, 'Open Jobs', 'Open', 'Count', 'Jobs', 'Total Jobs') || 0);
        if (siteName) bySite[siteName] = cnt;
      });
    }
    if (!Object.keys(bySite).length) {
      jobs.filter(function (j) {
        var s = (j.status || '').toLowerCase();
        return s === 'open' || s === 'in progress';
      }).forEach(function (j) {
        if (j.site) bySite[j.site] = (bySite[j.site] || 0) + 1;
      });
    }
    if (!Object.keys(bySite).length) bySite = { 'All Sites': openJobs };

    var poCosts = poCostRows.map(function (r, idx) {
      var rawDate = col(r, 'Date', 'PO Date', 'Invoice Date', 'Posting Date', 'Document Date', 'Created Date');
      return {
        id:          col(r, 'PO No', 'PO Number', 'PO', 'LPO', 'LPO No', 'Document No', 'Reference') || ('PO-' + String(idx + 1).padStart(3, '0')),
        site:        col(r, 'Site / Nursery', 'Site', 'Nursery', 'Location', 'Branch', 'School', 'Campus') || '',
        category:    col(r, 'Category', 'Expense Category', 'Sub Category', 'Cost Category', 'Type', 'Expense Type', 'GL Category') || '',
        vendor:      col(r, 'Vendor', 'Supplier', 'Vendor Name', 'Supplier Name') || '',
        desc:        col(r, 'Description', 'Item Description', 'Narration', 'Details', 'Remarks') || '',
        amount:      num(col(r, 'Amount', 'Net Amount', 'Gross Amount', 'Total', 'Value', 'Cost', 'PO Value', 'Invoice Amount', 'Spend')),
        date:        fmtDate(rawDate),
        dateObj:     parseDateValue(rawDate),
        deptTag:     col(r, 'Department', 'Expense Type', 'Type', 'Cost Center', 'Function') || '',
      };
    }).filter(function (p) { return p.site || p.category || p.amount || p.vendor || p.desc; });

    var maintenanceTagged = poCosts.filter(function (p) {
      var hay = [p.deptTag, p.category, p.desc].join(' ').toLowerCase();
      return hay.indexOf('maint') !== -1;
    });
    if (maintenanceTagged.length) {
      poCosts = maintenanceTagged;
    }

    return {
      kpis: {
        openJobs:       openJobs,
        completedMTD:   completedMTD,
        overdueJobs:    overdueJobs,
        avgResponseHrs: avgResp || 0,
        rag:            rag,
      },
      jobs:    jobs,
      byCat:   byCat,
      bySite:  bySite,
      poCosts: poCosts,
    };
  }

  // ---------------------------------------------------------------------------
  // Capex
  // ---------------------------------------------------------------------------

  function transformCapex(raw) {
    if (!raw) return null;
    var kpiRows  = raw.kpi  || [];
    var dataRows = raw.data || [];
    var categoryRows = raw.categories || [];
    var projectRows = raw.projects || [];

    var nurseries = dataRows.map(function (r) {
      var name = col(r, 'Nursery', 'Site', 'Branch', 'Name', 'Location') || '';
      if (!name) return null;
      var nameStr = String(name);
      var nm = nameStr.toLowerCase().trim();
      var isTotal = nm.includes('total') && !nm.includes('grand');
      var isGrand = nm.includes('29+ho') || nm === 'kfg (29+ho)' || nm.includes('grand total');
      return {
        name:     nameStr,
        ebitda:   num(col(r, 'EBITDA', 'Ebitda (AED)', 'Annual EBITDA')),
        budget:   num(col(r, 'Budget', 'Capex Budget', 'Annual Budget', 'Allocated Budget')),
        utilised: num(col(r, 'Utilised', 'Spent', 'Actual', 'Used', 'Actual Spend', 'Utilization')),
        balance:  num(col(r, 'Balance', 'Remaining', 'Balance (AED)', 'Net Balance', 'Balance Available Value')),
        balPct:   num(col(r, 'Bal%', 'Balance %', '% Remaining', 'Bal %', '% Balance', 'Balance Available %')),
        request:  num(col(r, 'Request', 'Cap Request', 'Additional', 'Pending Request', 'Request for Approval')),
        balAfter: num(col(r, 'Bal After', 'Balance After', 'Net After Request')),
        isTotal:  isTotal,
        isGrand:  isGrand,
      };
    }).filter(function (n) { return n !== null && n.name; });

    var categories = categoryRows.map(function (r) {
      return {
        category: String(col(r, 'Category', 'Capex Category') || '').trim(),
        budget: num(col(r, 'Annual Budget (AED)', 'Budget', 'Budget (AED)')),
        spent: num(col(r, 'Spent YTD (AED)', 'Spent', 'Spent (AED)')),
        committed: num(col(r, 'Committed (AED)', 'Committed')),
        remaining: num(col(r, 'Remaining (AED)', 'Remaining')),
        pct: num(col(r, '% Utilised', 'Utilised %', 'Pct Utilised')),
        rag: col(r, 'Final RAG', 'RAG', 'Auto RAG') || ''
      };
    }).filter(function (r) { return r.category; });

    var projects = projectRows.map(function (r, idx) {
      return {
        id: col(r, 'Project ID', 'ID', 'Ref') || ('CPX-' + String(idx + 1).padStart(3, '0')),
        name: col(r, 'Project Name', 'Name', 'Title') || '',
        nursery: col(r, 'Nursery', 'Site', 'Branch') || '',
        category: col(r, 'Category', 'Capex Category') || '',
        budget: num(col(r, 'Budget (AED)', 'Budget')),
        spent: num(col(r, 'Spent (AED)', 'Spent')),
        committed: num(col(r, 'Committed (AED)', 'Committed')),
        remaining: num(col(r, 'Remaining (AED)', 'Remaining')),
        status: col(r, 'Status') || '',
        approvalDate: fmtDate(col(r, 'Approval Date', 'Approved Date', 'Date'))
      };
    }).filter(function (r) { return r.id || r.nursery || r.category; });

    var nonTotals = nurseries.filter(function (n) { return !n.isTotal && !n.isGrand; });
    var grandRow = nurseries.find(function (n) {
      var nm = String(n.name || '').toLowerCase().trim();
      return nm.includes('29+ho') || nm === 'kfg (29+ho)' || nm.includes('grand total');
    });
    if (!grandRow) {
      grandRow = nurseries.find(function (n) {
        var nm = String(n.name || '').toLowerCase().trim();
        return nm.includes('kfg') && !nm.includes('llc');
      });
    }

    var totalBudget   = num(kpiVal(kpiRows, 'total budget') || kpiVal(kpiRows, 'budget'));
    var utilised      = num(kpiVal(kpiRows, 'utilised')     || kpiVal(kpiRows, 'spent'));
    var remaining     = num(kpiVal(kpiRows, 'remaining')    || kpiVal(kpiRows, 'balance'));
    var utilisedPct   = num(kpiVal(kpiRows, 'utilised%')    || kpiVal(kpiRows, '% utilised') || kpiVal(kpiRows, 'pct'));
    var negativeSites = num(kpiVal(kpiRows, 'negative')     || kpiVal(kpiRows, 'over budget') || kpiVal(kpiRows, 'over-budget'));
    var ragKpi        = kpiVal(kpiRows, 'rag') || kpiVal(kpiRows, 'overall');

    if (!totalBudget) totalBudget = grandRow ? grandRow.budget : nonTotals.reduce(function (s, n) { return s + n.budget; }, 0);
    if (!utilised)    utilised    = grandRow ? grandRow.utilised : nonTotals.reduce(function (s, n) { return s + n.utilised; }, 0);
    if (!remaining)   remaining   = totalBudget - utilised;
    if (!utilisedPct && totalBudget) utilisedPct = parseFloat((utilised / totalBudget * 100).toFixed(1));
    if (!negativeSites) negativeSites = nonTotals.filter(function (n) { return n.balance < 0; }).length;

    var rag = ragKpi ? String(ragKpi) : (negativeSites > 3 ? 'Red' : negativeSites > 0 ? 'Amber' : 'Green');

    var regionMap = {};
    nonTotals.forEach(function (n) {
      var prefix = n.name.split('-')[0].split(' ')[0] || 'Other';
      if (!regionMap[prefix]) regionMap[prefix] = { budget: 0, utilised: 0 };
      regionMap[prefix].budget   += n.budget;
      regionMap[prefix].utilised += n.utilised;
    });
    var regions = {
      labels:   Object.keys(regionMap),
      budget:   Object.values(regionMap).map(function (r) { return r.budget; }),
      utilised: Object.values(regionMap).map(function (r) { return r.utilised; }),
    };
    if (!regions.labels.length) {
      regions = { labels: ['Total'], budget: [totalBudget], utilised: [utilised] };
    }

    return {
      kpis: {
        totalBudget:   totalBudget,
        utilised:      utilised,
        remaining:     remaining,
        utilisedPct:   utilisedPct,
        negativeSites: negativeSites,
        rag:           rag,
      },
      nurseries: nurseries.length ? nurseries : [],
      regions:   regions,
      categories: categories,
      projects: projects,
    };
  }

  // ---------------------------------------------------------------------------
  // Projects
  // ---------------------------------------------------------------------------

  function transformProjects(raw) {
    if (!raw) return null;
    var kpiRows  = raw.kpi  || [];
    var dataRows = raw.data || [];

    var list = dataRows.map(function (r, idx) {
      var rag = col(r, 'RAG', 'RAG Status', 'Overall RAG', 'Status') || 'Amber';
      return {
        id:       col(r, 'Project ID', 'ID', 'Ref', 'Project No', 'Project Ref') || ('PM-' + String(idx + 1).padStart(3, '0')),
        name:     col(r, 'Project Name', 'Name', 'Title', 'Project Title') || '',
        owner:    col(r, 'Owner', 'Project Lead', 'Lead', 'Manager', 'Project Manager', 'Responsible') || '',
        end:      fmtDate(col(r, 'End Date', 'Due Date', 'Target Date', 'Completion Date', 'Target End')),
        pct:      num(col(r, '% Complete', 'Progress', 'Completion %', '% Completion', 'Complete %')),
        budget:   num(col(r, 'Budget', 'Budget (AED)', 'Total Budget', 'Approved Budget')),
        spent:    num(col(r, 'Spent', 'Actual', 'Spend', 'Spent (AED)', 'Actual Cost', 'Cost to Date')),
        priority: col(r, 'Priority') || 'Medium',
        rag:      rag,
      };
    }).filter(function (p) { return p.name; });

    var active  = list.length;
    var onTrack = list.filter(function (p) {
      var r = (p.rag || '').toLowerCase();
      return r === 'green' || r.includes('on track') || r.includes('track');
    }).length;
    var atRisk  = list.filter(function (p) {
      var r = (p.rag || '').toLowerCase();
      return r === 'amber' || r.includes('risk') || r.includes('amber');
    }).length;
    var delayed = list.filter(function (p) {
      var r = (p.rag || '').toLowerCase();
      return r === 'red' || r.includes('delay') || r.includes('red');
    }).length;

    var ragKpi = kpiVal(kpiRows, 'rag') || kpiVal(kpiRows, 'overall');
    var rag    = ragKpi ? String(ragKpi) : (delayed > 0 ? 'Amber' : 'Green');

    return {
      kpis: { active: active, onTrack: onTrack, atRisk: atRisk, delayed: delayed, rag: rag },
      list: list,
    };
  }

  // ---------------------------------------------------------------------------
  // Procurement
  // ---------------------------------------------------------------------------

  function transformProcurement(raw) {
    if (!raw) return null;
    var kpiRows  = raw.kpi  || [];
    var dataRows = raw.data || [];
    var supplierRows = raw.suppliers || [];

    var refLabel = 'PO Number';
    if (dataRows.length) {
      var firstKeys = Object.keys(dataRows[0]);
      if (firstKeys.some(function (k) { return String(k).toLowerCase().trim() === 'nursery'; })) {
        refLabel = 'Nursery';
      }
    }

    var pos = dataRows.map(function (r, idx) {
      var refValue = col(r, 'Nursery', 'PO Number', 'PO No', 'PO #', 'Purchase Order No', 'Ref') || '';
      var nurseryValue = col(r, 'Nursery', 'Site / Nursery', 'Site', 'Location', 'Branch') || refValue || '';
      var categoryValue = col(r, 'Category', 'Dept', 'Department', 'Cost Centre', 'Division') || '';
      var raisedRaw = col(r, 'Raised Date', 'Date Raised', 'Created Date', 'Date');
      return {
        ref:       refValue,
        po:        refValue,
        nursery:   nurseryValue,
        site:      nurseryValue,
        supplier:  col(r, 'Supplier', 'Vendor', 'Supplier Name') || '',
        desc:      col(r, 'Description', 'Desc', 'Item Description', 'Details') || '',
        category:  categoryValue,
        dept:      categoryValue,
        value:     num(col(r, 'Value (£)', 'Value (AED)', 'Value', 'Amount', 'Total Value', 'PO Value')),
        status:    col(r, 'Status', 'PO Status', 'Approval Status') || '',
        raisedDate: fmtDate(raisedRaw),
        raisedDateObj: parseDateValue(raisedRaw),
        date:      fmtDate(raisedRaw),
        delivery:  fmtDate(col(r, 'Actual Delivery', 'Expected Delivery', 'Delivery Date', 'Due Date')),
      };
    }).filter(function (p) { return p.ref || p.supplier || p.nursery; });

    var suppliers = supplierRows.map(function (r) {
      var expRaw = col(r, 'Contract Exp.', 'Contract Expiry', 'Expiry Date', 'Expiry');
      return {
        supplier: col(r, 'Supplier Name', 'Supplier', 'Vendor') || '',
        category: col(r, 'Category', 'Type') || '',
        contractExp: fmtDate(expRaw),
        contractExpObj: parseDateValue(expRaw),
        contact: col(r, 'Contact', 'Contact Person', 'Email', 'Phone') || '',
        annualSpend: num(col(r, 'Annual Spend (£)', 'Annual Spend (AED)', 'Annual Spend', 'Spend')),
        status: col(r, 'Status') || '',
        insuranceValid: col(r, 'Insurance Valid', 'Insurance', 'Insurance Status') || '',
        notes: col(r, 'Notes', 'Remarks') || ''
      };
    }).filter(function (s) { return s.supplier || s.category; });

    var activePOs = pos.filter(function (p) {
      var s = (p.status || '').toLowerCase();
      return s !== 'delivered' && s !== 'cancelled' && s !== 'closed' && s !== 'rejected';
    }).length;
    var pending   = pos.filter(function (p) { return (p.status || '').toLowerCase().includes('pending'); }).length;
    var overdue   = pos.filter(function (p) { return (p.status || '').toLowerCase() === 'overdue'; }).length;
    var spendYTD  = pos.filter(function (p) { return (p.status || '').toLowerCase() === 'delivered'; })
                       .reduce(function (s, p) { return s + p.value; }, 0);

    var ragKpi = kpiVal(kpiRows, 'rag') || kpiVal(kpiRows, 'overall');
    var rag    = ragKpi ? String(ragKpi) : (overdue > 2 ? 'Red' : overdue > 0 ? 'Amber' : 'Green');

    var activePOsK = kpiVal(kpiRows, 'active');
    var pendingK   = kpiVal(kpiRows, 'pending');
    var overdueK   = kpiVal(kpiRows, 'overdue');
    var spendK     = kpiVal(kpiRows, 'spend');

    var bySt = {}, byDept = {};
    pos.forEach(function (p) {
      if (p.status) bySt[p.status] = (bySt[p.status] || 0) + 1;
      if (p.category) byDept[p.category] = (byDept[p.category] || 0) + p.value;
    });

    return {
      kpis: {
        activePOs: activePOsK !== null ? num(activePOsK) : activePOs,
        pending:   pendingK   !== null ? num(pendingK)   : pending,
        overdue:   overdueK   !== null ? num(overdueK)   : overdue,
        spendYTD:  spendK     !== null ? num(spendK)     : spendYTD,
        rag:       rag,
      },
      refLabel: refLabel,
      pos: pos,
      suppliers: suppliers,
      bySt:   Object.keys(bySt).length   ? bySt   : { 'N/A': 1 },
      byDept: Object.keys(byDept).length ? byDept : { 'General': spendYTD },
    };
  }

  // ---------------------------------------------------------------------------
  // IT
  // ---------------------------------------------------------------------------

  function transformIT(raw) {
    if (!raw) return null;
    var kpiRows  = raw.kpi  || [];
    var dataRows = raw.data || [];

    var tickets = dataRows.map(function (r, idx) {
      return {
        id:       col(r, 'Ticket ID', 'Ticket No', 'ID', 'Ref', 'Ticket #') || ('IT-' + String(idx + 1).padStart(3, '0')),
        site:     col(r, 'Site', 'Location', 'Branch', 'Office', 'Nursery') || '',
        cat:      col(r, 'Category', 'Cat', 'Type', 'Issue Type', 'Department') || '',
        desc:     col(r, 'Description', 'Desc', 'Issue', 'Summary', 'Details') || '',
        priority: col(r, 'Priority') || 'Medium',
        status:   col(r, 'Status', 'Ticket Status', 'Resolution Status') || '',
        raised:   fmtDate(col(r, 'Raised Date', 'Date Raised', 'Created', 'Open Date', 'Date')),
      };
    }).filter(function (t) { return t.site || t.desc; });

    var openTickets = tickets.filter(function (t) {
      var s = (t.status || '').toLowerCase();
      return s === 'open' || s === 'in progress' || s === 'assigned' || s === 'pending';
    }).length;
    var resolved = tickets.filter(function (t) {
      var s = (t.status || '').toLowerCase();
      return s === 'resolved' || s === 'closed' || s === 'completed';
    }).length;
    var critical = tickets.filter(function (t) {
      return (t.priority || '').toLowerCase() === 'critical';
    }).length;

    var slaPct = num(kpiVal(kpiRows, 'sla'));
    if (!slaPct && tickets.length) {
      slaPct = parseFloat((resolved / tickets.length * 100).toFixed(1));
    }

    var ragKpi = kpiVal(kpiRows, 'rag') || kpiVal(kpiRows, 'overall');
    var rag    = ragKpi ? String(ragKpi) : (critical > 0 ? 'Amber' : openTickets > 10 ? 'Amber' : 'Green');

    var byCat = {};
    tickets.forEach(function (t) { if (t.cat) byCat[t.cat] = (byCat[t.cat] || 0) + 1; });

    // Build monthly trend (last 6 months from raised dates)
    var monthMap = {};
    tickets.forEach(function (t) {
      var m = 'Current';
      if (t.raised) {
        var parts = t.raised.split('/');
        if (parts.length >= 2) m = parts[1] + '/' + (parts[2] || '');
      }
      if (!monthMap[m]) monthMap[m] = { open: 0, resolved: 0 };
      var s = (t.status || '').toLowerCase();
      if (s === 'resolved' || s === 'closed') monthMap[m].resolved++;
      else monthMap[m].open++;
    });
    var trendLabels = Object.keys(monthMap).slice(-6);
    var trend = {
      labels:   trendLabels,
      open:     trendLabels.map(function (m) { return monthMap[m].open; }),
      resolved: trendLabels.map(function (m) { return monthMap[m].resolved; }),
    };
    if (!trendLabels.length) {
      trend = { labels: ['Current'], open: [openTickets], resolved: [resolved] };
    }

    return {
      kpis: {
        openTickets: openTickets,
        resolved:    resolved,
        critical:    critical,
        slaPct:      slaPct,
        rag:         rag,
      },
      tickets: tickets,
      byCat:   Object.keys(byCat).length ? byCat : { 'General': tickets.length },
      trend:   trend,
    };
  }

  // ---------------------------------------------------------------------------
  // M&A
  // ---------------------------------------------------------------------------

  function transformMA(raw) {
    if (!raw) return null;
    var kpiRows  = raw.kpi  || [];
    var dataRows = raw.data || [];
    var milestoneRows = raw.milestones || [];

    var deals = dataRows.map(function (r, idx) {
      var prob    = col(r, 'Probability %', 'Probability', 'Prob %', 'Prob') || 0;
      var probNum = num(String(prob).replace('%', ''));
      var value   = num(col(r, 'Est. Value (AED)', 'Est. Value', 'Value (AED)', 'Value', 'Estimated Value'));
      return {
        id:      col(r, 'Deal ID', 'ID', 'Ref', 'Deal No') || ('MA-' + String(idx + 1).padStart(3, '0')),
        target:  col(r, 'Target Name', 'Target', 'Company', 'Business Name') || '[Confidential]',
        type:    col(r, 'Deal Type', 'Type', 'Transaction Type') || '',
        value:   value,
        stage:   col(r, 'Stage', 'Deal Stage', 'Status', 'Current Stage') || '',
        lead:    col(r, 'Lead', 'Deal Lead', 'Owner', 'Responsible') || '',
        close:   fmtDate(col(r, 'Target Close', 'Close Date', 'Expected Close', 'Target Close Date')),
        prob:    probNum + '%',
        probNum: probNum,
      };
    }).filter(function (d) { return d.stage; });

    var dealMap = {};
    deals.forEach(function (d) { dealMap[d.id] = d.target; });

    var milestones = milestoneRows.map(function (r) {
      var dueRaw = col(r, 'Due Date', 'Target Date', 'Date');
      var dealId = col(r, 'Deal ID', 'Project ID', 'ID', 'Ref') || '';
      return {
        dealId: dealId,
        target: dealMap[dealId] || '',
        milestone: col(r, 'Milestone', 'Task', 'Action') || '',
        dueDate: fmtDate(dueRaw),
        dueDateObj: parseDateValue(dueRaw),
        owner: col(r, 'Owner', 'Lead', 'Responsible') || '',
        status: col(r, 'Status') || '',
        priority: col(r, 'Priority') || '',
        notes: col(r, 'Notes', 'Remarks') || ''
      };
    }).filter(function (m) { return m.dealId || m.milestone; });

    var pipeline     = deals.length;
    var dueDiligence = deals.filter(function (d) {
      return (d.stage || '').toLowerCase().includes('due diligence');
    }).length;
    var completed = deals.filter(function (d) {
      var s = (d.stage || '').toLowerCase();
      return s === 'completed' || s === 'closed' || s === 'done';
    }).length;
    var weightedValue = deals.reduce(function (s, d) {
      return s + (d.value * d.probNum / 100);
    }, 0);

    var ragKpi = kpiVal(kpiRows, 'rag') || kpiVal(kpiRows, 'overall');
    var rag    = ragKpi ? String(ragKpi) : (pipeline > 0 ? 'Green' : 'Amber');

    var bySt = {}, byVal = {};
    deals.forEach(function (d) {
      if (d.stage) {
        bySt[d.stage]  = (bySt[d.stage] || 0) + 1;
        byVal[d.stage] = (byVal[d.stage] || 0) + d.value;
      }
    });

    return {
      kpis: {
        pipeline:      pipeline,
        dueDiligence:  dueDiligence,
        completed:     completed,
        weightedValue: Math.round(weightedValue),
        rag:           rag,
      },
      deals: deals,
      milestones: milestones,
      bySt:  Object.keys(bySt).length  ? bySt  : { 'N/A': 0 },
      byVal: Object.keys(byVal).length ? byVal : { 'N/A': 0 },
    };
  }

  // ---------------------------------------------------------------------------
  // Greenfield
  // ---------------------------------------------------------------------------

  function transformGreenfield(raw) {
    if (!raw) return null;
    var kpiRows  = raw.kpi  || [];
    var dataRows = raw.data || [];

    // Pipeline Overview headers (row 3): Site ID, Site Name, Location / City,
    // Space — Indoor (Sqft), Space — Outdoor (Sqft), Enrol Capacity, Lease Tenure,
    // EBITDA @ 85% (AED), Rent Yr 1 (AED), Capex (AED), Brand, Opening Date,
    // Build Status, RAG Override, Auto RAG, Final RAG
    var sites = dataRows.map(function (r) {
      var status = col(r, 'Build Status', 'Status', 'Development Status') || '';
      var rag    = col(r, 'Final RAG', 'RAG Override', 'Auto RAG', 'RAG') || '';
      return {
        id:      col(r, 'Site ID', 'ID', 'Ref', 'Site No') || '',
        name:    col(r, 'Site Name', 'Name', 'Site', 'Development Name') || '',
        city:    col(r, 'Location / City', 'Location', 'City', 'Area') || '',
        indoor:  num(col(r, 'Space — Indoor (Sqft)', 'Indoor (Sqft)', 'Indoor', 'GLA Indoor', 'Indoor Area')),
        outdoor: num(col(r, 'Space — Outdoor (Sqft)', 'Outdoor (Sqft)', 'Outdoor', 'Outdoor Area')),
        enrol:   num(col(r, 'Enrol Capacity', 'Enrolment Capacity', 'Capacity', 'Max Enrolment', 'Enrollment')),
        lease:   col(r, 'Lease Tenure', 'Lease', 'Tenure', 'Lease Term') || '',
        ebitda:  num(col(r, 'EBITDA @ 85% (AED)', 'EBITDA (AED)', 'EBITDA', 'Projected EBITDA')),
        rent:    num(col(r, 'Rent Yr 1 (AED)', 'Rent (AED)', 'Rent', 'Annual Rent', 'Rent Year 1')),
        capex:   num(col(r, 'Capex (AED)', 'Capex', 'Capital', 'Investment', 'Total Capex')),
        brand:   col(r, 'Brand', 'Brand Name', 'Nursery Brand') || '',
        opening: fmtDate(col(r, 'Opening Date', 'Opening', 'Target Opening', 'Expected Opening')),
        status:  status,
        rag:     rag,
      };
    }).filter(function (s) { return s.name; });

    var inPipeline = sites.length;
    var underConstruction = sites.filter(function (s) {
      var st = (s.status || '').toLowerCase();
      return st.includes('construction') || st.includes('fit-out') || st.includes('fit out') || st.includes('build');
    }).length;
    var completed = sites.filter(function (s) {
      var st = (s.status || '').toLowerCase();
      return st === 'completed' || st === 'open' || st === 'operational' || st === 'trading';
    }).length;
    var totalCapex   = sites.reduce(function (s, site) { return s + (site.capex   || 0); }, 0);
    var totalEBITDA  = sites.reduce(function (s, site) { return s + (site.ebitda  || 0); }, 0);

    var ragKpi = kpiVal(kpiRows, 'rag') || kpiVal(kpiRows, 'overall');
    var rag    = ragKpi ? String(ragKpi) : 'Green';

    return {
      kpis: {
        inPipeline:        inPipeline,
        underConstruction: underConstruction,
        completed:         completed,
        totalCapex:        totalCapex,
        totalEBITDA:       totalEBITDA,
        rag:               rag,
      },
      sites: sites,
    };
  }

  // ---------------------------------------------------------------------------
  // Other Projects
  // ---------------------------------------------------------------------------

  function transformOther(raw) {
    if (!raw) return null;
    var kpiRows  = raw.kpi  || [];
    var dataRows = raw.data || [];
    var actionRows = raw.actions || [];

    var projects = dataRows.map(function (r, idx) {
      var rag = col(r, 'RAG', 'RAG Status', 'Overall RAG', 'Status') || 'Amber';
      return {
        id:       col(r, 'Project ID', 'ID', 'Ref', 'Project No', 'Project Ref') || ('OP-' + String(idx + 1).padStart(3, '0')),
        name:     col(r, 'Project Name', 'Name', 'Title', 'Project Title') || '',
        cat:      col(r, 'Category', 'Cat', 'Type', 'Project Type', 'Project Category') || '',
        owner:    col(r, 'Owner', 'Project Lead', 'Lead', 'Manager', 'Responsible') || '',
        pct:      num(col(r, '% Complete', 'Progress', 'Completion %', '% Completion', 'Complete %')),
        budget:   num(col(r, 'Budget', 'Budget (AED)', 'Total Budget', 'Approved Budget')),
        priority: col(r, 'Priority') || 'Medium',
        rag:      rag,
      };
    }).filter(function (p) { return p.name; });

    var projectMap = {};
    projects.forEach(function (p) { projectMap[p.id] = p.name; });

    var actions = actionRows.map(function (r) {
      var raisedRaw = col(r, 'Raised Date', 'Date Raised', 'Raised');
      var dueRaw = col(r, 'Due Date', 'Target Date', 'Date');
      var pid = col(r, 'Project ID', 'ID', 'Ref') || '';
      return {
        projectId: pid,
        project: projectMap[pid] || '',
        action: col(r, 'Action / Decision', 'Action', 'Decision', 'Task') || '',
        owner: col(r, 'Owner', 'Responsible', 'Lead') || '',
        raisedDate: fmtDate(raisedRaw),
        raisedDateObj: parseDateValue(raisedRaw),
        dueDate: fmtDate(dueRaw),
        dueDateObj: parseDateValue(dueRaw),
        status: col(r, 'Status') || '',
        priority: col(r, 'Priority') || '',
        notes: col(r, 'Notes', 'Remarks') || ''
      };
    }).filter(function (a) { return a.projectId || a.action; });

    var active  = projects.length;
    var onTrack = projects.filter(function (p) {
      var r = (p.rag || '').toLowerCase();
      return r === 'green' || r.includes('on track') || r.includes('track');
    }).length;
    var atRisk  = projects.filter(function (p) {
      var r = (p.rag || '').toLowerCase();
      return r === 'amber' || r.includes('risk') || r.includes('amber');
    }).length;
    var overdue = projects.filter(function (p) {
      var r = (p.rag || '').toLowerCase();
      return r === 'red' || r.includes('overdue') || r.includes('delay') || r.includes('red');
    }).length;

    var ragKpi = kpiVal(kpiRows, 'rag') || kpiVal(kpiRows, 'overall');
    var rag    = ragKpi ? String(ragKpi) : (overdue > 0 ? 'Amber' : 'Green');

    var byCat = {}, bySt = {};
    projects.forEach(function (p) {
      if (p.cat) byCat[p.cat] = (byCat[p.cat] || 0) + 1;
      if (p.rag) bySt[p.rag]  = (bySt[p.rag] || 0) + 1;
    });

    return {
      kpis: { active: active, onTrack: onTrack, atRisk: atRisk, overdue: overdue, rag: rag },
      projects: projects,
      actions: actions,
      byCat: Object.keys(byCat).length ? byCat : { 'General': active },
      bySt:  Object.keys(bySt).length  ? bySt  : { 'Amber': active },
    };
  }

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------

  window.KFGTransform = {
    /**
     * Convert raw SheetJS department data into the dashboard's expected format.
     * @param {Object} raw  Output of KFGGraph.fetchAllDepartments()
     * @returns {Object}    Structured data matching the DEMO shape in dashboard.js
     */
    processLiveData: function (raw) {
      if (!raw) return null;
      var result = {
        maintenance: transformMaintenance(raw.maintenance),
        capex:       transformCapex(raw.capex),
        projects:    transformProjects(raw.projects),
        procurement: transformProcurement(raw.procurement),
        it:          transformIT(raw.it),
        ma:          transformMA(raw.ma),
        greenfield:  transformGreenfield(raw.greenfield),
        other:       transformOther(raw.other),
      };
      console.log('[KFGTransform] Transformed live data:', result);
      return result;
    },
  };

  console.log('[KFGTransform] Module loaded.');
})();

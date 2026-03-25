/* ===================================================================
 *  KFG Executive Dashboard – Configuration Management
 *  Persists admin-editable settings (charts, tables, refresh, users)
 *  in localStorage and exposes them via window.KFGConfig.
 * =================================================================== */
(function () {
  'use strict';

  const STORAGE_KEY = 'kfg-dashboard-config';

  /* ------------------------------------------------------------------ */
  /*  Default configuration                                              */
  /* ------------------------------------------------------------------ */
  const DEFAULT_CONFIG = {
    version: '1.0.0',
    refreshInterval: 5, // minutes
    adminEmails: ['shady.sabbagh@kidsfirstgroup.com'],
    siteUrl: 'https://kfguae.sharepoint.com/sites/ExecutiveDashboard',

    charts: {
      'c-rag': {
        title: 'Portfolio Health',
        subtitle: 'RAG distribution across all departments',
        type: 'doughnut',
        showLegend: true,
        legendPosition: 'right',
        cutout: '65%',
        colors: ['#10B981', '#F59E0B', '#EF4444'],
      },
      'c-cap-reg': {
        title: 'Capex by Region',
        subtitle: 'Budget vs utilised YTD (AED)',
        type: 'bar',
        showLegend: true,
        legendPosition: 'top',
        colors: ['#2563EB', '#F59E0B'],
        borderRadius: 4,
      },
      'c-maint-site': {
        title: 'Open Jobs by Site',
        subtitle: 'Top 10 sites by open job count',
        type: 'bar',
        indexAxis: 'y',
        showLegend: false,
        colors: ['#2563EB'],
        borderRadius: 4,
      },
      'c-maint-cat': {
        title: 'Jobs by Category',
        subtitle: 'Planned vs reactive vs regulatory',
        type: 'doughnut',
        showLegend: true,
        legendPosition: 'right',
        cutout: '60%',
        colors: ['#2563EB', '#EF4444', '#F59E0B', '#10B981'],
      },
      'c-capex-reg': {
        title: 'Budget vs Utilised by Region',
        subtitle: 'AED - allocated vs spent YTD',
        type: 'bar',
        showLegend: true,
        legendPosition: 'top',
        colors: ['#2563EB', '#F59E0B'],
        borderRadius: 4,
      },
      'c-capex-neg': {
        title: 'Over-Budget Sites',
        subtitle: 'Nurseries with negative balance',
        type: 'bar',
        indexAxis: 'y',
        showLegend: false,
        colors: ['#EF4444'],
        borderRadius: 4,
      },
      'c-pm-status': {
        title: 'Projects by Status',
        subtitle: 'On track / at risk / delayed',
        type: 'doughnut',
        showLegend: true,
        legendPosition: 'right',
        cutout: '60%',
        colors: ['#10B981', '#F59E0B', '#EF4444'],
      },
      'c-pm-budget': {
        title: 'Budget vs Spend',
        subtitle: 'Per project (AED)',
        type: 'bar',
        showLegend: true,
        legendPosition: 'top',
        colors: ['#2563EB', '#F59E0B'],
        borderRadius: 4,
      },
      'c-proc-status': {
        title: 'PO Status Breakdown',
        subtitle: 'Current status of all purchase orders',
        type: 'doughnut',
        showLegend: true,
        legendPosition: 'right',
        cutout: '60%',
        colors: ['#10B981', '#2563EB', '#F59E0B', '#3B82F6'],
      },
      'c-proc-spend': {
        title: 'Spend by Department',
        subtitle: 'AED - YTD spend distribution',
        type: 'bar',
        showLegend: false,
        colors: ['#2563EB'],
        borderRadius: 4,
      },
      'c-it-cat': {
        title: 'Tickets by Category',
        subtitle: 'Volume by issue type',
        type: 'bar',
        showLegend: false,
        colors: ['#2563EB'],
        borderRadius: 4,
      },
      'c-it-trend': {
        title: 'Open vs Resolved Trend',
        subtitle: 'Monthly comparison',
        type: 'line',
        showLegend: true,
        legendPosition: 'top',
        colors: ['#EF4444', '#10B981'],
        tension: 0.4,
        fill: true,
      },
      'c-ma-stage': {
        title: 'Deals by Stage',
        subtitle: 'Pipeline progression',
        type: 'bar',
        showLegend: false,
        colors: ['#94A3B8', '#F59E0B', '#2563EB'],
        borderRadius: 4,
      },
      'c-ma-val': {
        title: 'Value by Stage',
        subtitle: 'AED - estimated deal value',
        type: 'bar',
        showLegend: false,
        colors: ['#94A3B8', '#F59E0B', '#2563EB'],
        borderRadius: 4,
      },
      'c-gf-ratio': {
        title: 'EBITDA : Rent Ratio',
        subtitle: 'Target minimum 5x per site',
        type: 'bar',
        showLegend: false,
        colors: ['#10B981', '#F59E0B'],
        borderRadius: 6,
      },
      'c-gf-capex': {
        title: 'Capex by Site',
        subtitle: 'Development investment (AED)',
        type: 'bar',
        showLegend: false,
        colors: ['#2563EB'],
        borderRadius: 6,
      },
      'c-op-cat': {
        title: 'Projects by Category',
        subtitle: 'Distribution across project types',
        type: 'doughnut',
        showLegend: true,
        legendPosition: 'right',
        cutout: '60%',
        colors: ['#2563EB', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6'],
      },
      'c-op-status': {
        title: 'Projects by Status',
        subtitle: 'On track / at risk / delayed',
        type: 'doughnut',
        showLegend: true,
        legendPosition: 'right',
        cutout: '60%',
        colors: ['#10B981', '#F59E0B', '#EF4444'],
      },
    },

    tables: {
      't-home':       { sortable: true, filterable: false, pageSize: 0,  visibleColumns: null },
      't-maint-jobs': { sortable: true, filterable: true,  pageSize: 20, visibleColumns: null },
      't-capex':      { sortable: true, filterable: true,  pageSize: 0,  visibleColumns: null },
      't-pm':         { sortable: true, filterable: true,  pageSize: 20, visibleColumns: null },
      't-proc':       { sortable: true, filterable: true,  pageSize: 20, visibleColumns: null },
      't-it':         { sortable: true, filterable: true,  pageSize: 20, visibleColumns: null },
      't-ma':         { sortable: true, filterable: true,  pageSize: 20, visibleColumns: null },
      't-op':         { sortable: true, filterable: true,  pageSize: 20, visibleColumns: null },
    },
  };

  /* ------------------------------------------------------------------ */
  /*  Internal state                                                     */
  /* ------------------------------------------------------------------ */
  let _config = null;
  const _listeners = [];

  /* ------------------------------------------------------------------ */
  /*  Helpers                                                            */
  /* ------------------------------------------------------------------ */

  /** Deep-clone a plain object / array. */
  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  /** Recursively merge `src` into `target`. Arrays are replaced, not merged. */
  function deepMerge(target, src) {
    const out = Object.assign({}, target);
    for (const key of Object.keys(src)) {
      if (
        src[key] !== null &&
        typeof src[key] === 'object' &&
        !Array.isArray(src[key]) &&
        typeof target[key] === 'object' &&
        target[key] !== null &&
        !Array.isArray(target[key])
      ) {
        out[key] = deepMerge(target[key], src[key]);
      } else {
        out[key] = deepClone(src[key]);
      }
    }
    // Preserve keys from target that are not in src (new defaults)
    for (const key of Object.keys(target)) {
      if (!(key in src)) {
        out[key] = deepClone(target[key]);
      }
    }
    return out;
  }

  /** Resolve a dot-notation path to the value inside an object. */
  function getByPath(obj, path) {
    const parts = path.split('.');
    let current = obj;
    for (const part of parts) {
      if (current == null || typeof current !== 'object') return undefined;
      current = current[part];
    }
    return current;
  }

  /** Set a value at a dot-notation path, creating intermediate objects. */
  function setByPath(obj, path, value) {
    const parts = path.split('.');
    let current = obj;
    for (let i = 0; i < parts.length - 1; i++) {
      const part = parts[i];
      if (current[part] == null || typeof current[part] !== 'object') {
        current[part] = {};
      }
      current = current[part];
    }
    current[parts[parts.length - 1]] = value;
  }

  /** Persist current config to localStorage. */
  function save() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(_config));
    } catch (e) {
      console.error('[KFGConfig] Failed to save config:', e);
    }
  }

  /** Notify registered listeners. */
  function notify(key, value) {
    for (const cb of _listeners) {
      try {
        cb(key, value);
      } catch (e) {
        console.error('[KFGConfig] onChange callback error:', e);
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Public API                                                         */
  /* ------------------------------------------------------------------ */

  window.KFGConfig = {

    /** Load config from localStorage, deep-merged with defaults. */
    init() {
      let stored = null;
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) stored = JSON.parse(raw);
      } catch (e) {
        console.warn('[KFGConfig] Could not parse stored config, using defaults.', e);
      }

      if (stored) {
        // Deep merge: defaults as base, stored values on top
        _config = deepMerge(deepClone(DEFAULT_CONFIG), stored);
      } else {
        _config = deepClone(DEFAULT_CONFIG);
      }

      save(); // persist the merged result so new defaults are stored
    },

    /** Get a config value by dot-notation key. */
    get(key) {
      if (!_config) this.init();
      const val = getByPath(_config, key);
      // Return a clone for objects to prevent external mutation
      if (val !== null && typeof val === 'object') return deepClone(val);
      return val;
    },

    /** Set a config value by dot-notation key. Triggers onChange callbacks. */
    set(key, value) {
      if (!_config) this.init();
      setByPath(_config, key, typeof value === 'object' && value !== null ? deepClone(value) : value);
      save();
      notify(key, value);
    },

    /** Get chart-specific config. */
    getChartConfig(chartId) {
      return this.get('charts.' + chartId);
    },

    /** Save chart config (merges with existing). */
    setChartConfig(chartId, config) {
      if (!_config) this.init();
      const existing = _config.charts[chartId] || {};
      _config.charts[chartId] = Object.assign({}, existing, deepClone(config));
      save();
      notify('charts.' + chartId, _config.charts[chartId]);
    },

    /** Get table-specific config. */
    getTableConfig(tableId) {
      return this.get('tables.' + tableId);
    },

    /** Save table config (merges with existing). */
    setTableConfig(tableId, config) {
      if (!_config) this.init();
      const existing = _config.tables[tableId] || {};
      _config.tables[tableId] = Object.assign({}, existing, deepClone(config));
      save();
      notify('tables.' + tableId, _config.tables[tableId]);
    },

    /** Get auto-refresh interval in minutes. */
    getRefreshInterval() {
      return this.get('refreshInterval');
    },

    /** Set auto-refresh interval in minutes. */
    setRefreshInterval(mins) {
      const val = Math.max(1, Math.round(Number(mins) || 5));
      this.set('refreshInterval', val);
    },

    /** Get list of admin emails. */
    getAdminEmails() {
      return this.get('adminEmails') || [];
    },

    /** Add an admin email (no duplicates, case-insensitive check). */
    addAdminEmail(email) {
      if (!_config) this.init();
      const normalised = (email || '').trim().toLowerCase();
      if (!normalised) return;
      const list = _config.adminEmails || [];
      if (list.some((e) => e.toLowerCase() === normalised)) return; // already exists
      list.push(normalised);
      _config.adminEmails = list;
      save();
      notify('adminEmails', deepClone(list));
    },

    /** Remove an admin email (case-insensitive match). */
    removeAdminEmail(email) {
      if (!_config) this.init();
      const normalised = (email || '').trim().toLowerCase();
      if (!normalised) return;
      const before = (_config.adminEmails || []).length;
      _config.adminEmails = (_config.adminEmails || []).filter(
        (e) => e.toLowerCase() !== normalised
      );
      if (_config.adminEmails.length !== before) {
        save();
        notify('adminEmails', deepClone(_config.adminEmails));
      }
    },

    /** Export full config as a formatted JSON string. */
    exportConfig() {
      if (!_config) this.init();
      return JSON.stringify(_config, null, 2);
    },

    /** Import config from a JSON string. Validates and merges with defaults. */
    importConfig(jsonString) {
      let incoming;
      try {
        incoming = JSON.parse(jsonString);
      } catch (e) {
        console.error('[KFGConfig] importConfig: invalid JSON.', e);
        return false;
      }
      if (typeof incoming !== 'object' || incoming === null || Array.isArray(incoming)) {
        console.error('[KFGConfig] importConfig: JSON root must be an object.');
        return false;
      }
      _config = deepMerge(deepClone(DEFAULT_CONFIG), incoming);
      save();
      notify('*', deepClone(_config));
      return true;
    },

    /** Reset all config back to defaults. */
    resetToDefaults() {
      _config = deepClone(DEFAULT_CONFIG);
      save();
      notify('*', deepClone(_config));
    },

    /** Register a change listener. Callback receives (key, newValue). */
    onChange(callback) {
      if (typeof callback === 'function') {
        _listeners.push(callback);
      }
    },
  };
})();

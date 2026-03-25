/**
 * KFG Executive Dashboard — Microsoft Graph API Module
 * Fetches Excel data from SharePoint using Microsoft Graph API.
 *
 * Namespace: window.KFGGraph
 */
(function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Configuration
  // ---------------------------------------------------------------------------

  const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

  /** SharePoint Excel files mapped by department key. */
  const SHEETS = {
    maintenance: { file: '02_Maintenance_Dashboard.xlsx',         kpi: 'KPI Summary', data: 'Jobs Register',      sites: 'Sites Summary' },
    capex:       { file: '03_Capex_Dashboard_CEO.xlsx',           kpi: 'KPI Summary', data: 'Nursery Budget View' },
    projects:    { file: '01_Project_Management_Dashboard.xlsx',  kpi: 'KPI Summary', data: 'Active Projects' },
    procurement: { file: '04_Procurement_Dashboard.xlsx',         kpi: 'KPI Summary', data: 'PO Register' },
    it:          { file: '05_IT_Dashboard.xlsx',                  kpi: 'KPI Summary', data: 'Ticket Register' },
    ma:          { file: '06_MA_Dashboard.xlsx',                  kpi: 'KPI Summary', data: 'Deal Pipeline' },
    greenfield:  { file: '07_Greenfield_Dashboard_CEO.xlsx',      kpi: 'KPI Summary', data: 'Pipeline Overview' },
    other:       { file: '08_Other_Projects_Dashboard.xlsx',      kpi: 'KPI Summary', data: 'Projects Register' },
  };

  /** Default auto-refresh interval in milliseconds (5 minutes). */
  const DEFAULT_REFRESH_MS = 5 * 60 * 1000;

  // ---------------------------------------------------------------------------
  // Internal state
  // ---------------------------------------------------------------------------

  let _siteUrl   = 'https://kfguae.sharepoint.com/sites/ExecutiveDashboard';
  let _siteId    = null;
  let _driveId   = null;
  let _fileIdCache   = {};          // { fileName: fileId }
  let _dataCache     = {};          // { cacheKey: { data, timestamp } }
  let _lastFetchTime = null;
  let _refreshTimer  = null;
  let _refreshMs     = DEFAULT_REFRESH_MS;
  let _callbacks     = [];          // onDataUpdate listeners

  // ---------------------------------------------------------------------------
  // Logging helpers
  // ---------------------------------------------------------------------------

  function _log(msg)  { console.log('[KFGGraph] ' + msg); }
  function _warn(msg) { console.warn('[KFGGraph] ' + msg); }

  // ---------------------------------------------------------------------------
  // graphFetch — authenticated fetch wrapper
  // ---------------------------------------------------------------------------

  /**
   * Perform an authenticated GET against Microsoft Graph.
   * Handles 401 (re-auth) and 429 (rate-limit with retry-after backoff).
   *
   * @param {string} url  Full Graph API URL.
   * @param {number} _retries  Internal retry counter.
   * @returns {Promise<Object|null>}  Parsed JSON or null on failure.
   */
  async function graphFetch(url, _retries) {
    if (_retries === undefined) { _retries = 0; }
    var MAX_RETRIES = 3;

    try {
      // Obtain access token from auth module
      if (!window.KFGAuth || typeof window.KFGAuth.getToken !== 'function') {
        _warn('KFGAuth module not loaded — cannot obtain access token.');
        return null;
      }

      var token = await window.KFGAuth.getToken();
      if (!token) {
        _warn('No access token available.');
        return null;
      }

      var response = await fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
      });

      // --- 401 Unauthorized — token may have expired -------------------------
      if (response.status === 401) {
        _warn('401 received — attempting token refresh.');
        if (_retries < MAX_RETRIES && typeof window.KFGAuth.login === 'function') {
          await window.KFGAuth.login();
          return graphFetch(url, _retries + 1);
        }
        _warn('Token refresh failed after ' + _retries + ' retries.');
        return null;
      }

      // --- 429 Too Many Requests — respect Retry-After ----------------------
      if (response.status === 429) {
        var retryAfter = parseInt(response.headers.get('Retry-After') || '5', 10);
        _warn('429 rate-limited — retrying after ' + retryAfter + 's (attempt ' + (_retries + 1) + ').');
        if (_retries < MAX_RETRIES) {
          await _sleep(retryAfter * 1000);
          return graphFetch(url, _retries + 1);
        }
        _warn('Rate-limit retries exhausted.');
        return null;
      }

      // --- Other non-OK statuses --------------------------------------------
      if (!response.ok) {
        _warn('Graph API error: ' + response.status + ' ' + response.statusText + ' — ' + url);
        return null;
      }

      return await response.json();
    } catch (err) {
      _warn('Network error: ' + (err.message || err) + ' — ' + url);
      return null;
    }
  }

  /** Simple promise-based sleep. */
  function _sleep(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  // ---------------------------------------------------------------------------
  // resolveSiteAndDrive — one-time site/drive discovery
  // ---------------------------------------------------------------------------

  /**
   * Resolve the SharePoint site ID and default document-library drive ID.
   * Results are cached so subsequent calls are free.
   *
   * @returns {Promise<{siteId: string, driveId: string}|null>}
   */
  async function resolveSiteAndDrive() {
    // Return cached values if available
    if (_siteId && _driveId) {
      return { siteId: _siteId, driveId: _driveId };
    }

    try {
      // Parse host and relative path from the configured site URL
      var parsed   = new URL(_siteUrl);
      var host     = parsed.hostname;                       // e.g. kfguae.sharepoint.com
      var sitePath = parsed.pathname.replace(/\/+$/, '');   // e.g. /sites/ExecutiveDashboard

      // 1. Resolve site ID
      var siteUrl  = GRAPH_BASE + '/sites/' + host + ':' + sitePath;
      var siteJson = await graphFetch(siteUrl);
      if (!siteJson || !siteJson.id) {
        _warn('Could not resolve SharePoint site ID.');
        return null;
      }
      _siteId = siteJson.id;
      _log('Resolved site ID: ' + _siteId);

      // 2. Resolve drives — pick the default "Documents" library
      var drivesUrl  = GRAPH_BASE + '/sites/' + _siteId + '/drives';
      var drivesJson = await graphFetch(drivesUrl);
      if (!drivesJson || !drivesJson.value || drivesJson.value.length === 0) {
        _warn('No drives found for site.');
        return null;
      }

      // Prefer the drive named "Documents"; fall back to the first drive
      var docDrive = drivesJson.value.find(function (d) {
        return d.name === 'Documents' || d.name === 'Shared Documents';
      });
      _driveId = (docDrive || drivesJson.value[0]).id;
      _log('Resolved drive ID: ' + _driveId);

      return { siteId: _siteId, driveId: _driveId };
    } catch (err) {
      _warn('resolveSiteAndDrive error: ' + (err.message || err));
      return null;
    }
  }

  // ---------------------------------------------------------------------------
  // File-ID resolution (cached)
  // ---------------------------------------------------------------------------

  /**
   * Look up (and cache) the Graph item ID for a given Excel file name.
   *
   * @param {string} fileName  e.g. '02_Maintenance_Dashboard.xlsx'
   * @returns {Promise<string|null>}  The item ID, or null on failure.
   */
  async function resolveFileId(fileName) {
    if (_fileIdCache[fileName]) {
      return _fileIdCache[fileName];
    }

    var ids = await resolveSiteAndDrive();
    if (!ids) return null;

    var searchUrl = GRAPH_BASE + '/drives/' + _driveId +
                    "/root/search(q='" + encodeURIComponent(fileName) + "')";
    var result = await graphFetch(searchUrl);
    if (!result || !result.value || result.value.length === 0) {
      _warn('File not found on SharePoint: ' + fileName);
      return null;
    }

    // Find exact match (search can return partial matches)
    var match = result.value.find(function (item) {
      return item.name === fileName;
    });
    if (!match) {
      // Fall back to first result if exact name match fails
      match = result.value[0];
    }

    _fileIdCache[fileName] = match.id;
    _log('Resolved file ID for ' + fileName + ': ' + match.id);
    return match.id;
  }

  // ---------------------------------------------------------------------------
  // parseRange — convert Excel usedRange values into array of objects
  // ---------------------------------------------------------------------------

  /**
   * Convert a 2-D values array (first row = headers) into an array of plain
   * objects.  Empty rows are discarded and numeric-looking strings are cast.
   *
   * @param {Array<Array>} values  The `values` property from a usedRange response.
   * @returns {Array<Object>}
   */
  function parseRange(values) {
    if (!values || values.length < 2) return [];

    var headers = values[0].map(function (h) {
      return (h != null ? String(h).trim() : '');
    });

    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];

      // Skip completely empty rows
      var isEmpty = row.every(function (cell) {
        return cell === null || cell === undefined || cell === '';
      });
      if (isEmpty) continue;

      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        if (!key) continue; // skip columns with blank headers

        var val = (j < row.length) ? row[j] : null;

        // Auto-detect and convert numeric values
        if (val !== null && val !== '' && !isNaN(val) && typeof val === 'string') {
          val = Number(val);
        }

        obj[key] = val;
      }
      rows.push(obj);
    }

    return rows;
  }

  // ---------------------------------------------------------------------------
  // fetchSheet — fetch a single worksheet from a department's Excel file
  // ---------------------------------------------------------------------------

  /**
   * Fetch and parse a specific worksheet from a department's Excel file.
   *
   * @param {string} deptKey   Department key from SHEETS config (e.g. 'maintenance').
   * @param {string} sheetKey  Sheet key within the department config (e.g. 'kpi', 'data', 'sites').
   * @returns {Promise<Array<Object>|null>}  Parsed rows, or null on error.
   */
  async function fetchSheet(deptKey, sheetKey) {
    var dept = SHEETS[deptKey];
    if (!dept) {
      _warn('Unknown department key: ' + deptKey);
      return null;
    }

    var sheetName = dept[sheetKey];
    if (!sheetName) {
      _warn('Unknown sheet key "' + sheetKey + '" for department "' + deptKey + '".');
      return null;
    }

    // Check cache
    var cacheKey = deptKey + '::' + sheetKey;
    if (_dataCache[cacheKey]) {
      _log('Returning cached data for ' + cacheKey);
      return _dataCache[cacheKey].data;
    }

    try {
      var fileId = await resolveFileId(dept.file);
      if (!fileId) return null;

      // URL-encode the sheet name to handle spaces and special characters
      var encodedSheet = encodeURIComponent(sheetName);
      var rangeUrl = GRAPH_BASE + '/drives/' + _driveId +
                     '/items/' + fileId +
                     '/workbook/worksheets/' + encodedSheet + '/usedRange';

      var rangeJson = await graphFetch(rangeUrl);
      if (!rangeJson || !rangeJson.values) {
        _warn('No data returned for sheet "' + sheetName + '" in ' + dept.file);
        return null;
      }

      var parsed = parseRange(rangeJson.values);

      // Cache the result
      _dataCache[cacheKey] = { data: parsed, timestamp: Date.now() };
      _log('Fetched ' + parsed.length + ' rows from ' + deptKey + '/' + sheetKey);

      return parsed;
    } catch (err) {
      _warn('fetchSheet error (' + deptKey + '/' + sheetKey + '): ' + (err.message || err));
      return null;
    }
  }

  // ---------------------------------------------------------------------------
  // fetchAllDepartments — parallel fetch of every department
  // ---------------------------------------------------------------------------

  /**
   * Fetch all departments' data in parallel.  For each department the 'kpi'
   * and 'data' sheets are fetched; the 'sites' sheet is also fetched for
   * maintenance.
   *
   * @returns {Promise<Object>}  Keyed by department, each value is
   *   { kpi: [...], data: [...], sites?: [...] } or null on failure.
   */
  async function fetchAllDepartments() {
    // Ensure site and drive are resolved before fanning out
    var ids = await resolveSiteAndDrive();
    if (!ids) {
      _warn('Cannot fetch departments — site/drive resolution failed.');
      return null;
    }

    var deptKeys = Object.keys(SHEETS);

    // Build an array of { deptKey, sheetKey } tasks
    var tasks = [];
    deptKeys.forEach(function (dk) {
      var dept = SHEETS[dk];
      tasks.push({ deptKey: dk, sheetKey: 'kpi' });
      tasks.push({ deptKey: dk, sheetKey: 'data' });
      if (dept.sites) {
        tasks.push({ deptKey: dk, sheetKey: 'sites' });
      }
    });

    // Execute all tasks in parallel
    var results = await Promise.allSettled(
      tasks.map(function (t) {
        return fetchSheet(t.deptKey, t.sheetKey).then(function (data) {
          return { deptKey: t.deptKey, sheetKey: t.sheetKey, data: data };
        });
      })
    );

    // Assemble structured output
    var output = {};
    deptKeys.forEach(function (dk) { output[dk] = null; });

    results.forEach(function (r) {
      if (r.status !== 'fulfilled' || !r.value) return;
      var v = r.value;
      if (!output[v.deptKey]) {
        output[v.deptKey] = {};
      }
      output[v.deptKey][v.sheetKey] = v.data;
    });

    _lastFetchTime = new Date();
    _log('All departments fetched at ' + _lastFetchTime.toISOString());

    // Notify registered listeners
    _notifyCallbacks(output);

    return output;
  }

  // ---------------------------------------------------------------------------
  // Refresh & auto-refresh
  // ---------------------------------------------------------------------------

  /**
   * Re-fetch all data, invalidating the sheet-data cache first.
   * @returns {Promise<Object>}
   */
  async function refreshData() {
    _log('Refreshing data — clearing cache.');
    _dataCache = {};
    return fetchAllDepartments();
  }

  /**
   * Set the auto-refresh interval (in minutes).  Pass 0 to disable.
   * @param {number} mins  Interval in minutes.
   */
  function setRefreshInterval(mins) {
    // Clear any existing timer
    if (_refreshTimer) {
      clearInterval(_refreshTimer);
      _refreshTimer = null;
    }

    if (!mins || mins <= 0) {
      _log('Auto-refresh disabled.');
      return;
    }

    _refreshMs = mins * 60 * 1000;
    _refreshTimer = setInterval(function () {
      _log('Auto-refresh triggered.');
      refreshData();
    }, _refreshMs);
    _log('Auto-refresh set to every ' + mins + ' minute(s).');
  }

  /** Return the timestamp of the last successful fetchAllDepartments call. */
  function getLastFetchTime() {
    return _lastFetchTime;
  }

  // ---------------------------------------------------------------------------
  // Callback management
  // ---------------------------------------------------------------------------

  /**
   * Register a callback that fires whenever fetchAllDepartments completes.
   * @param {Function} callback  Receives the full data object.
   */
  function onDataUpdate(callback) {
    if (typeof callback === 'function') {
      _callbacks.push(callback);
    }
  }

  /** Notify all registered callbacks. */
  function _notifyCallbacks(data) {
    _callbacks.forEach(function (cb) {
      try {
        cb(data);
      } catch (err) {
        _warn('onDataUpdate callback error: ' + (err.message || err));
      }
    });
  }

  // ---------------------------------------------------------------------------
  // Initialization
  // ---------------------------------------------------------------------------

  /**
   * Initialize the Graph module with a SharePoint site URL.
   * Resolves site and drive IDs eagerly so subsequent calls are fast.
   *
   * @param {string} [siteUrl]  SharePoint site URL. Defaults to the KFG
   *   Executive Dashboard site.
   */
  async function init(siteUrl) {
    if (siteUrl) {
      _siteUrl = siteUrl;
    }
    _log('Initializing with site URL: ' + _siteUrl);

    // Resolve site and drive eagerly
    var ids = await resolveSiteAndDrive();
    if (ids) {
      _log('Initialization complete — site and drive resolved.');
    } else {
      _warn('Initialization complete but site/drive resolution failed. ' +
            'Calls will retry resolution on demand.');
    }

    // Start default auto-refresh
    setRefreshInterval(_refreshMs / 60000);

    return ids;
  }

  // ---------------------------------------------------------------------------
  // Cleanup on page unload
  // ---------------------------------------------------------------------------

  window.addEventListener('beforeunload', function () {
    if (_refreshTimer) {
      clearInterval(_refreshTimer);
      _refreshTimer = null;
    }
  });

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------

  window.KFGGraph = {
    init:               init,
    fetchSheet:         fetchSheet,
    fetchAllDepartments: fetchAllDepartments,
    refreshData:        refreshData,
    getLastFetchTime:   getLastFetchTime,
    setRefreshInterval: setRefreshInterval,
    onDataUpdate:       onDataUpdate,
  };

  _log('Module loaded.');
})();

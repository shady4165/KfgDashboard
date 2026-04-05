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

  /**
   * SharePoint Excel files mapped by department key.
   * Each file lives in its own named document library (library field).
   */
  const SHEETS = {
    maintenance: { file: '02_Maintenance_Dashboard.xlsx', library: 'Maintenance', kpi: '📊 KPI Summary', data: '🔧 Jobs Register', sites: '🏢 Sites Summary', poCosts: 'MaintenancePOCost', poCostsFile: '02_Maintenance_Dashboard.xlsx', poCostsLibrary: 'Maintenance', dataHeaderRow: 1, sitesHeaderRow: 1, poCostsHeaderRow: 'auto', poCostsHeaderKeywords: ['site', 'nursery', 'category', 'amount', 'date', 'po', 'vendor'] },
    capex:       { file: '03_Capex_Dashboard_CEO.xlsx',     library: 'Capex',             kpi: '📊 KPI Summary', data: '📊 Nursery Budget View', dataHeaderRow: 1 },
    projects:    { file: '01_Project_Management_Dashboard.xlsx', library: 'Project Management', kpi: '📊 KPI Summary', data: '📁 Active Projects', dataHeaderRow: 1 },
    procurement: { file: '04_Procurement_Dashboard.xlsx',   library: 'Procurement',       kpi: '📊 KPI Summary', data: '📋 PO Register',          dataHeaderRow: 2 },
    it:          { file: '05_IT_Dashboard.xlsx',            library: 'IT',                kpi: '📊 KPI Summary', data: '🎫 Ticket Register',      dataHeaderRow: 1 },
    ma:          { file: '06_MA_Dashboard.xlsx',            library: 'M&A',               kpi: '📊 KPI Summary', data: '🤝 Deal Pipeline',        dataHeaderRow: 2 },
    greenfield:  { file: '07_Greenfield_Dashboard_CEO.xlsx', library: 'Greenfield',       kpi: '📊 KPI Summary', data: '🏗 Pipeline Overview',    dataHeaderRow: 2 },
    other:       { file: '08_Other_Projects_Dashboard.xlsx', library: 'Other Projects',   kpi: '📊 KPI Summary', data: '📁 Projects Register',    dataHeaderRow: 1 },
  };

  /** Default auto-refresh interval in milliseconds (5 minutes). */
  const DEFAULT_REFRESH_MS = 5 * 60 * 1000;

  // ---------------------------------------------------------------------------
  // Internal state
  // ---------------------------------------------------------------------------

  let _siteUrl   = 'https://kfguae.sharepoint.com/sites/ExecutiveDashboard';
  let _siteId    = null;
  let _driveId   = null;            // default (Documents) drive
  let _driveMap  = {};              // { libraryName: driveId }
  let _fileIdCache   = {};          // { fileName: itemId }
  let _workbookCache = {};          // { deptKey: XLSX workbook object }
  let _workbookPromiseCache = {};   // { wbCacheKey: Promise<XLSX workbook> }
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
  // graphFetchBinary — authenticated binary (ArrayBuffer) fetch
  // ---------------------------------------------------------------------------

  async function graphFetchBinary(url, _retries) {
    if (_retries === undefined) { _retries = 0; }
    var MAX_RETRIES = 3;

    try {
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
        headers: { Authorization: 'Bearer ' + token }
      });

      if (response.status === 401) {
        _warn('401 received on binary fetch — attempting token refresh.');
        if (_retries < MAX_RETRIES && typeof window.KFGAuth.login === 'function') {
          await window.KFGAuth.login();
          return graphFetchBinary(url, _retries + 1);
        }
        return null;
      }

      if (response.status === 429) {
        var retryAfter = parseInt(response.headers.get('Retry-After') || '5', 10);
        _warn('429 rate-limited on binary fetch — retrying after ' + retryAfter + 's.');
        if (_retries < MAX_RETRIES) {
          await _sleep(retryAfter * 1000);
          return graphFetchBinary(url, _retries + 1);
        }
        return null;
      }

      if (!response.ok) {
        _warn('Binary Graph API error: ' + response.status + ' ' + response.statusText + ' — ' + url);
        return null;
      }

      return await response.arrayBuffer();
    } catch (err) {
      _warn('Binary network error: ' + (err.message || err) + ' — ' + url);
      return null;
    }
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
    // Return cached values if already resolved
    if (_siteId && _driveId) {
      return { siteId: _siteId, driveId: _driveId };
    }

    try {
      var parsed   = new URL(_siteUrl);
      var host     = parsed.hostname;
      var sitePath = parsed.pathname.replace(/\/+$/, '');

      // 1. Resolve site ID
      var siteJson = await graphFetch(GRAPH_BASE + '/sites/' + host + ':' + sitePath);
      if (!siteJson || !siteJson.id) {
        _warn('Could not resolve SharePoint site ID.');
        return null;
      }
      _siteId = siteJson.id;
      _log('Resolved site ID: ' + _siteId);

      // 2. Enumerate ALL drives and build a name→id map
      var drivesJson = await graphFetch(GRAPH_BASE + '/sites/' + _siteId + '/drives');
      if (!drivesJson || !drivesJson.value || drivesJson.value.length === 0) {
        _warn('No drives found for site.');
        return null;
      }

      _driveMap = {};
      drivesJson.value.forEach(function (d) { _driveMap[d.name] = d.id; });
      _log('Site libraries: ' + Object.keys(_driveMap).join(', '));

      // Default drive = Documents (fallback to first)
      var docDrive = drivesJson.value.find(function (d) {
        return d.name === 'Documents' || d.name === 'Shared Documents';
      });
      _driveId = (docDrive || drivesJson.value[0]).id;

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

    var encodedName = encodeURIComponent(fileName);

    // 1. Try direct root path
    var directUrl = GRAPH_BASE + '/drives/' + _driveId + '/root:/' + encodedName;
    var directResult = await graphFetch(directUrl);
    if (directResult && directResult.id) {
      _fileIdCache[fileName] = directResult.id;
      _log('Resolved ' + fileName + ' at root: ' + directResult.id);
      return directResult.id;
    }

    // 2. List root children — log them for diagnostics, then search inside folders
    var childrenUrl = GRAPH_BASE + '/drives/' + _driveId + '/root/children';
    var childrenResult = await graphFetch(childrenUrl);
    if (childrenResult && childrenResult.value && childrenResult.value.length > 0) {
      var rootItems = childrenResult.value.map(function (f) { return f.name + (f.folder ? '/' : ''); });
      _log('Drive root contents: ' + rootItems.join(', '));

      // Check root list directly (some drives skip path-based access)
      var rootFile = childrenResult.value.find(function (f) { return f.name === fileName; });
      if (rootFile) {
        _fileIdCache[fileName] = rootFile.id;
        _log('Resolved ' + fileName + ' from root children: ' + rootFile.id);
        return rootFile.id;
      }

      // Search one level deep inside each subfolder
      var folders = childrenResult.value.filter(function (f) { return !!f.folder; });
      for (var i = 0; i < folders.length; i++) {
        var folderUrl = GRAPH_BASE + '/drives/' + _driveId + '/root:/' +
                        encodeURIComponent(folders[i].name) + '/' + encodedName;
        var folderResult = await graphFetch(folderUrl);
        if (folderResult && folderResult.id) {
          _fileIdCache[fileName] = folderResult.id;
          _log('Resolved ' + fileName + ' in folder "' + folders[i].name + '": ' + folderResult.id);
          return folderResult.id;
        }
      }
    }

    // 3. Drive search as last resort
    var searchUrl = GRAPH_BASE + '/drives/' + _driveId +
                    "/root/search(q='" + encodedName + "')";
    var result = await graphFetch(searchUrl);
    if (!result || !result.value || result.value.length === 0) {
      _warn('File not found on SharePoint: ' + fileName);
      return null;
    }

    var match = result.value.find(function (item) {
      return item.name === fileName;
    });
    if (!match) {
      match = result.value[0];
    }

    _fileIdCache[fileName] = match.id;
    _log('Resolved file ID for ' + fileName + ': ' + match.id + ' (search fallback)');
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
  // detectHeaderRow — auto detect header row for flexible sheets
  // ---------------------------------------------------------------------------

  function detectHeaderRow(ws, keywords) {
    if (!window.XLSX || !ws) return 0;
    var rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
    if (!rows || !rows.length) return 0;
    var bestRow = 0;
    var bestScore = -1;
    var terms = (keywords || []).map(function (k) { return String(k).toLowerCase(); });
    for (var i = 0; i < Math.min(rows.length, 8); i++) {
      var joined = (rows[i] || []).map(function (v) { return String(v || '').toLowerCase(); }).join(' | ');
      var score = 0;
      terms.forEach(function (t) { if (joined.indexOf(t) !== -1) score++; });
      if (score > bestScore) {
        bestScore = score;
        bestRow = i;
      }
    }
    return bestScore >= 2 ? bestRow : 0;
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
      var ids = await resolveSiteAndDrive();
      if (!ids) return null;

      // Allow a specific sheet to come from a different workbook or library
      var effectiveLibrary = dept[sheetKey + 'Library'] || dept.library;
      var effectiveFile = dept[sheetKey + 'File'] || dept.file;

      // Use the library-specific drive for this sheet source
      var driveId = (effectiveLibrary && _driveMap[effectiveLibrary]) ? _driveMap[effectiveLibrary] : _driveId;
      if (!driveId) {
        _warn('Drive not found for library "' + effectiveLibrary + '"');
        return null;
      }

      // Get (or reuse cached) parsed workbook for this sheet source workbook
      var wbCacheKey = deptKey + '::wb::' + effectiveLibrary + '::' + effectiveFile;
      var wb = _workbookCache[wbCacheKey];
      if (!wb) {
        if (!_workbookPromiseCache[wbCacheKey]) {
          _workbookPromiseCache[wbCacheKey] = (async function () {
            var itemId = null;
            var encodedFile = encodeURIComponent(effectiveFile);

            var meta = await graphFetch(GRAPH_BASE + '/drives/' + driveId + '/root:/' + encodedFile);
            if (meta && meta.id) {
              itemId = meta.id;
            }

            if (!itemId) {
              meta = await graphFetch(GRAPH_BASE + '/drives/' + driveId + '/root:/General/' + encodedFile);
              if (meta && meta.id) {
                itemId = meta.id;
              }
            }

            if (!itemId) {
              _log('Searching drive for ' + effectiveFile);
              var sr = await graphFetch(GRAPH_BASE + '/drives/' + driveId +
                "/root/search(q='" + effectiveFile.replace(/'/g, "''") + "')");
              if (sr && sr.value && sr.value.length > 0) {
                var hit = sr.value.find(function (i) { return i.name === effectiveFile; }) || sr.value[0];
                if (hit && hit.id) {
                  itemId = hit.id;
                  _log('Found ' + effectiveFile + ' via search at: ' +
                    (hit.parentReference ? hit.parentReference.path : 'unknown path'));
                }
              }
            }

            if (!itemId) {
              throw new Error('Could not resolve file ID for ' + effectiveFile);
            }

            var arrayBuffer = await graphFetchBinary(
              GRAPH_BASE + '/drives/' + driveId + '/items/' + itemId + '/content'
            );
            if (!arrayBuffer) {
              throw new Error('Download failed for ' + effectiveFile);
            }

            if (!window.XLSX) {
              throw new Error('SheetJS library not loaded');
            }

            var parsed = window.XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
            _workbookCache[wbCacheKey] = parsed;
            _log('Downloaded and parsed ' + effectiveFile + ' — sheets: ' + parsed.SheetNames.join(', '));
            return parsed;
          })();
        }

        try {
          wb = await _workbookPromiseCache[wbCacheKey];
        } catch (wbErr) {
          delete _workbookPromiseCache[wbCacheKey];
          _warn(wbErr.message || wbErr);
          return null;
        }
      }

      // Extract the requested worksheet — match by normalised name so emoji
      // prefixes and case differences do not cause a miss.
      function normalizeSheetName(name) {
        return String(name || '').replace(/^[^\p{L}\p{N}]+/gu, '').trim().toLowerCase();
      }
      var actualSheetName = wb.SheetNames.find(function (n) {
        return normalizeSheetName(n) === normalizeSheetName(sheetName);
      });
      var ws = actualSheetName ? wb.Sheets[actualSheetName] : null;
      if (!ws) {
        _warn('Sheet "' + sheetName + '" not found in ' + effectiveFile +
              '. Available: ' + wb.SheetNames.join(', '));
        return null;
      }

      var sheetJsonOpts = { defval: null };
      var headerRowProp = dept[sheetKey + 'HeaderRow'];
      if (headerRowProp === 'auto') {
        sheetJsonOpts.range = detectHeaderRow(ws, dept[sheetKey + 'HeaderKeywords'] || []);
      } else if (headerRowProp !== undefined && headerRowProp !== null) {
        sheetJsonOpts.range = headerRowProp;
      }
      var rows = window.XLSX.utils.sheet_to_json(ws, sheetJsonOpts);
      _dataCache[cacheKey] = { data: rows, timestamp: Date.now() };
      _log('Parsed ' + rows.length + ' rows from ' + deptKey + '/' + sheetKey);
      return rows;
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
      if (dept.poCosts) {
        tasks.push({ deptKey: dk, sheetKey: 'poCosts' });
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
    _workbookCache = {};
    _workbookPromiseCache = {};
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

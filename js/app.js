/**
 * KFG Executive Dashboard — Main Application Entry Point
 * Orchestrates auth, data loading, admin panel, and dashboard rendering.
 */
'use strict';

(function () {
  /* ══════════════════════════════════════════════
     CONFIGURATION
  ══════════════════════════════════════════════ */
  const APP_CONFIG = {
    tenantId: 'organizations', // Will be overridden if user enters specific tenant
    siteUrl: 'https://kfguae.sharepoint.com/sites/ExecutiveDashboard',
    clientId: '', // Set via login form or environment
  };

  let IS_DEMO = false;
  let refreshTimer = null;

  /* ══════════════════════════════════════════════
     DOM REFERENCES
  ══════════════════════════════════════════════ */
  const $ = (id) => document.getElementById(id);

  /* ══════════════════════════════════════════════
     LOADER / PROGRESS
  ══════════════════════════════════════════════ */
  function setProgress(pct, msg) {
    const bar = $('loader-bar');
    const txt = $('loader-txt');
    if (bar) bar.style.width = pct + '%';
    if (txt) txt.textContent = msg;
  }

  function hideLoader() {
    setTimeout(() => {
      const loader = $('loader');
      if (loader) loader.classList.add('gone');
    }, 400);
  }

  /* ══════════════════════════════════════════════
     TIMESTAMP DISPLAY
  ══════════════════════════════════════════════ */
  function updateTimestamp() {
    const el = $('ts-label');
    if (el) {
      el.textContent = new Date().toLocaleString('en-GB', {
        day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'
      });
    }
  }

  /* ══════════════════════════════════════════════
     DEMO MODE
  ══════════════════════════════════════════════ */
  function startDemo() {
    IS_DEMO = true;
    showDashboard();

    const pulseDot = $('pulse-dot');
    const connLabel = $('conn-label');
    const userChip = $('user-chip');
    const envLabel = $('env-label');

    if (pulseDot) pulseDot.className = 'pulse amber';
    if (connLabel) connLabel.textContent = 'Demo Data';
    if (userChip) userChip.textContent = 'Demo User';
    if (envLabel) envLabel.textContent = 'Sample Data Mode';

    updateTimestamp();
    setProgress(30, 'Loading demo data...');

    // Initialize dashboard with demo data
    if (window.KFGDashboard) {
      window.KFGDashboard.init();
      window.KFGDashboard.buildAll();
    }

    setProgress(100, 'Ready');
    hideLoader();

    // Set up auto-refresh for demo mode
    const interval = window.KFGConfig ? window.KFGConfig.getRefreshInterval() : 5;
    startAutoRefresh(interval);
  }

  /* ══════════════════════════════════════════════
     CREDENTIAL PERSISTENCE
  ══════════════════════════════════════════════ */
  const CREDS_KEY = 'kfg-saved-credentials';

  function saveCredentials(clientId, siteUrl) {
    try {
      localStorage.setItem(CREDS_KEY, JSON.stringify({ clientId, siteUrl }));
    } catch (e) { /* ignore */ }
  }

  function loadCredentials() {
    try {
      const raw = localStorage.getItem(CREDS_KEY);
      const creds = raw ? JSON.parse(raw) : null;
      if (!creds) return null;

      // Detect and fix malformed siteUrl (e.g. URL duplicated or concatenated)
      if (creds.siteUrl) {
        const url = String(creds.siteUrl).trim();
        // If URL contains itself twice or has spaces in unusual places, it's corrupted
        if ((url.match(/sharepoint\.com.*sharepoint\.com/i) || url.includes(' ExecutiveDashboard ExecutiveDashboard'))) {
          console.warn('[App] Detected malformed siteUrl in localStorage, clearing credentials');
          localStorage.removeItem(CREDS_KEY);
          return null;
        }
      }
      return creds;
    } catch (e) { return null; }
  }

  /* ══════════════════════════════════════════════
     MICROSOFT AUTH CONNECTION
  ══════════════════════════════════════════════ */
  async function startConnection() {
    const clientId = $('in-client') ? $('in-client').value.trim() : '';
    const siteUrl = $('in-site') ? $('in-site').value.trim() : APP_CONFIG.siteUrl;

    if (!clientId) {
      alert(
        'Please enter your Azure App Client ID.\n\n' +
        'If you don\'t have one yet, click the setup guide link below the form.\n\n' +
        'Or use Demo Mode to preview the dashboard now.'
      );
      return;
    }

    // Save credentials so they're pre-filled next time
    saveCredentials(clientId, siteUrl);

    APP_CONFIG.clientId = clientId;
    APP_CONFIG.siteUrl = siteUrl;

    // Initialize MSAL auth
    if (!window.KFGAuth) {
      alert('Authentication module not loaded. Please refresh the page.');
      return;
    }

    try {
      await window.KFGAuth.init({
        clientId: clientId,
        tenantId: APP_CONFIG.tenantId,
        redirectUri: window.location.origin + window.location.pathname,
      });

      await window.KFGAuth.login();

      if (window.KFGAuth.isAuthenticated()) {
        await startLiveDashboard();
      }
    } catch (err) {
      console.error('Authentication failed:', err);
      alert('Authentication failed. Please check your Client ID and try again.\n\nError: ' + err.message);
    }
  }

  /* ══════════════════════════════════════════════
     LIVE DASHBOARD (WITH SHAREPOINT DATA)
  ══════════════════════════════════════════════ */
  async function startLiveDashboard() {
    IS_DEMO = false;
    showDashboard();

    const user = window.KFGAuth.getUser();
    const pulseDot = $('pulse-dot');
    const connLabel = $('conn-label');
    const userChip = $('user-chip');
    const envLabel = $('env-label');

    if (pulseDot) pulseDot.className = 'pulse green';
    if (connLabel) connLabel.textContent = 'Live — SharePoint';
    if (userChip) userChip.textContent = user ? user.name : 'Authenticated';
    if (envLabel) envLabel.textContent = 'Operations Hub';

    updateTimestamp();
    setProgress(20, 'Authenticating...');

    // Initialize admin panel if user is admin
    if (window.KFGAuth.isAdmin() && window.KFGAdmin) {
      window.KFGAdmin.init();
    }

    setProgress(40, 'Fetching SharePoint data...');

    // Initialize Graph API and fetch data
    let liveLoaded = false;

    if (window.KFGGraph) {
      window.KFGGraph.init(APP_CONFIG.siteUrl);

      try {
        const liveData = await window.KFGGraph.fetchAllDepartments();
        if (liveData && window.KFGDashboard) {
          const transformed = window.KFGTransform
            ? window.KFGTransform.processLiveData(liveData)
            : liveData;
          window.KFGDashboard.setData(transformed);
          liveLoaded = true;
        }
      } catch (err) {
        console.warn('Could not fetch live data, falling back to demo:', err);
        if (window.KFGAdmin) {
          window.KFGAdmin.showToast('Could not load live data. Showing demo data.', 'error');
        }
      }
    }

    setProgress(80, 'Building dashboard...');

    if (window.KFGDashboard && !liveLoaded) {
      window.KFGDashboard.init();
      window.KFGDashboard.buildAll();
    }

    setProgress(100, 'Ready');
    hideLoader();

    // Auto-refresh
    const interval = window.KFGConfig ? window.KFGConfig.getRefreshInterval() : 5;
    startAutoRefresh(interval);
  }

  /* ══════════════════════════════════════════════
     SHOW / HIDE SCREENS
  ══════════════════════════════════════════════ */
  function showDashboard() {
    const login = $('login-screen');
    const app   = $('app');
    if (login) { login.style.display = 'none'; }
    if (app)   { app.style.display   = 'block'; }
    // Ensure the home page is active on first show
    const homePage = $('pg-home');
    if (homePage && !document.querySelector('.page.active')) {
      homePage.classList.add('active');
      const homeNav = document.querySelector('[data-page="home"]');
      if (homeNav) homeNav.classList.add('active');
    }
  }

  function showLogin() {
    const login = $('login-screen');
    const app   = $('app');
    if (login) { login.style.display = 'flex'; }
    if (app)   { app.style.display   = 'none'; }
  }

  /* ══════════════════════════════════════════════
     REFRESH
  ══════════════════════════════════════════════ */
  async function refreshAll() {
    updateTimestamp();

    if (!IS_DEMO && window.KFGGraph) {
      try {
        const liveData = await window.KFGGraph.refreshData();
        if (liveData && window.KFGDashboard) {
          const transformed = window.KFGTransform
            ? window.KFGTransform.processLiveData(liveData)
            : liveData;
          window.KFGDashboard.setData(transformed);
        }
      } catch (err) {
        console.warn('Refresh failed:', err);
      }
    }

    if (window.KFGDashboard) {
      window.KFGDashboard.refreshAll();
    }

    if (window.KFGAdmin && window.KFGAuth && window.KFGAuth.isAdmin()) {
      window.KFGAdmin.showToast('Dashboard refreshed', 'info');
    }
  }

  function startAutoRefresh(intervalMins) {
    if (refreshTimer) clearInterval(refreshTimer);
    if (intervalMins > 0) {
      refreshTimer = setInterval(refreshAll, intervalMins * 60 * 1000);
    }
  }

  /* ══════════════════════════════════════════════
     SETUP GUIDE
  ══════════════════════════════════════════════ */
  function showSetupGuide() {
    alert(
      'HOW TO GET YOUR AZURE CLIENT ID (5 minutes):\n\n' +
      '1. Open your browser and go to:\n   https://portal.azure.com\n\n' +
      '2. Sign in with your company Microsoft 365 account\n\n' +
      '3. In the search bar at the top, type "App registrations" and click it\n\n' +
      '4. Click "+ New registration" (top left button)\n\n' +
      '5. Fill in:\n' +
      '   - Name: KFG Executive Dashboard\n' +
      '   - Account types: choose "Accounts in this organizational directory only"\n' +
      '   - Redirect URI: choose "Single-page application (SPA)" and paste:\n' +
      '     ' + window.location.origin + window.location.pathname + '\n\n' +
      '6. Click "Register"\n\n' +
      '7. You will see "Application (client) ID" — copy that number\n\n' +
      '8. Now click "API permissions" on the left\n' +
      '9. Click "+ Add a permission" → Microsoft Graph → Delegated permissions\n' +
      '10. Search for and add: Sites.Read.All, Files.Read.All, User.Read\n' +
      '11. Click "Grant admin consent for Kids First Group"\n\n' +
      '12. Paste the Client ID back into the dashboard login screen\n\n' +
      'That\'s it! You only do this setup once.'
    );
  }

  /* ══════════════════════════════════════════════
     PAGE TAB SWITCHING
  ══════════════════════════════════════════════ */
  function switchPageTab(btn, pageId, tabId) {
    var page = document.getElementById(pageId);
    if (!page) return;
    page.querySelectorAll('.page-tab-btn').forEach(function (b) { b.classList.remove('active'); });
    btn.classList.add('active');
    page.querySelectorAll('.tab-pane').forEach(function (p) { p.classList.remove('active'); });
    var pane = document.getElementById(pageId + '-' + tabId);
    if (pane) pane.classList.add('active');
  }

  /* ══════════════════════════════════════════════
     NAVIGATION
  ══════════════════════════════════════════════ */
  function go(name, el) {
    // Switch visible page using .active class (CSS: .page{display:none} .page.active{display:block})
    document.querySelectorAll('.page').forEach(function (p) { p.classList.remove('active'); });
    document.querySelectorAll('.nav-item').forEach(function (n) { n.classList.remove('active'); });
    var page = $('pg-' + name);
    if (page) page.classList.add('active');
    if (el) el.classList.add('active');

    // Update topbar breadcrumb label
    var labelEl = $('topbar-page-label');
    if (labelEl && el) labelEl.textContent = el.textContent.trim();

    window.scrollTo({ top: 0, behavior: 'smooth' });

    // Delegate to dashboard module to build/refresh the page content
    if (window.KFGDashboard && typeof window.KFGDashboard.goByKey === 'function') {
      window.KFGDashboard.goByKey(name);
    }
  }

  function goByKey(name) {
    var navEl = document.querySelector('[data-page="' + name + '"]');
    go(name, navEl);
  }

  /* ══════════════════════════════════════════════
     LOGOUT
  ══════════════════════════════════════════════ */
  function logout() {
    if (refreshTimer) clearInterval(refreshTimer);
    IS_DEMO = false;
    // Clear all MSAL cached tokens so next login gets fresh consent
    try { localStorage.clear(); } catch (e) { /* ignore */ }
    if (window.KFGAuth) {
      window.KFGAuth.logout();
    } else {
      showLogin();
    }
  }

  /* ══════════════════════════════════════════════
     INITIALIZATION
  ══════════════════════════════════════════════ */
  async function initApp() {
    // Initialize config manager
    if (window.KFGConfig) {
      window.KFGConfig.init();

      // Listen for refresh interval changes
      window.KFGConfig.onChange(function (key, value) {
        if (key === 'refreshInterval') {
          startAutoRefresh(value);
        }
      });
    }

    // Check for auth redirect callback
    if (window.KFGAuth && window.location.hash && window.location.hash.includes('code=')) {
      try {
        await window.KFGAuth.init({
          clientId: APP_CONFIG.clientId,
          tenantId: APP_CONFIG.tenantId,
          redirectUri: window.location.origin + window.location.pathname,
        });
        // handleRedirectPromise is called inside init
        if (window.KFGAuth.isAuthenticated()) {
          await startLiveDashboard();
          return;
        }
      } catch (err) {
        console.warn('Redirect callback handling failed:', err);
      }
    }

    // Pre-fill saved credentials
    var saved = loadCredentials();
    if (saved) {
      if (saved.clientId && $('in-client')) $('in-client').value = saved.clientId;
      if (saved.siteUrl && $('in-site')) $('in-site').value = saved.siteUrl;
    }

    // Show login screen
    var loader = $('loader');
    if (loader) loader.classList.add('gone');
    showLogin();
  }

  /* ══════════════════════════════════════════════
     EXPOSE GLOBAL FUNCTIONS
  ══════════════════════════════════════════════ */
  // These are called from onclick handlers in HTML
  window.doDemo = startDemo;
  window.doConnect = startConnection;
  window.refreshAll = refreshAll;
  window.showSetupGuide = showSetupGuide;
  window.go = go;
  window.goByKey = goByKey;
  window.logout = logout;
  window.switchPageTab = switchPageTab;

  /* ══════════════════════════════════════════════
     BOOT
  ══════════════════════════════════════════════ */
  window.addEventListener('load', initApp);
})();

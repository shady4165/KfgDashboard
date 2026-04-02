/**
 * KFG Executive Dashboard - Authentication Module
 *
 * Uses MSAL.js 2.x (Browser) for Azure AD / Microsoft Entra authentication
 * with OAuth 2.0 Authorization Code Flow + PKCE.
 *
 * Assumes the global `msal` object is available via CDN script tag:
 *   <script src="https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js"></script>
 *
 * Usage:
 *   KFGAuth.init({ clientId: '...', tenantId: '...', redirectUri: '...' });
 *   await KFGAuth.login();
 *   const token = await KFGAuth.getToken();
 */
(function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Private state
  // ---------------------------------------------------------------------------

  /** @type {msal.PublicClientApplication|null} */
  let msalInstance = null;

  /** @type {{ name: string, email: string, isAdmin: boolean }|null} */
  let currentUser = null;

  /** @type {boolean} */
  let authenticated = false;

  /** @type {Function[]} */
  const authStateCallbacks = [];

  /** @type {string[]} */
  let adminEmails = ['shady.sabbagh@kidsfirstgroup.com'];

  /** Allowed tenant domain */
  const ALLOWED_DOMAIN = 'kidsfirstgroup.com';

  /** Microsoft Graph scopes requested during login / token acquisition */
  const LOGIN_SCOPES = ['User.Read'];
  const API_SCOPES = ['Sites.Read.All', 'Files.Read.All', 'User.Read'];

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  /**
   * Notify all registered auth-state listeners.
   * @param {boolean} isAuthenticated
   * @param {{ name: string, email: string, isAdmin: boolean }|null} user
   */
  function notifyAuthStateChange(isAuthenticated, user) {
    authStateCallbacks.forEach(function (cb) {
      try {
        cb(isAuthenticated, user);
      } catch (err) {
        console.error('[KFGAuth] Error in auth state callback:', err);
      }
    });
  }

  /**
   * Set the internal authenticated state and fire listeners.
   */
  function setAuthState(isAuth, user) {
    authenticated = isAuth;
    currentUser = user;
    notifyAuthStateChange(authenticated, currentUser);
  }

  /**
   * Validate that the email belongs to the allowed domain.
   * @param {string} email
   * @returns {boolean}
   */
  function isAllowedDomain(email) {
    if (!email) return false;
    var parts = email.split('@');
    return parts.length === 2 && parts[1].toLowerCase() === ALLOWED_DOMAIN;
  }

  /**
   * Check whether an email is in the admin list.
   * @param {string} email
   * @returns {boolean}
   */
  function checkAdmin(email) {
    if (!email) return false;
    return adminEmails.some(function (admin) {
      return admin.toLowerCase() === email.toLowerCase();
    });
  }

  /**
   * Fetch the current user's profile from Microsoft Graph /me endpoint.
   * @param {string} accessToken
   * @returns {Promise<{ name: string, email: string, isAdmin: boolean }|null>}
   */
  async function fetchUserProfile(accessToken) {
    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: 'Bearer ' + accessToken },
      });

      if (!response.ok) {
        console.error('[KFGAuth] Graph /me request failed:', response.status);
        return null;
      }

      const profile = await response.json();
      const email = (profile.mail || profile.userPrincipalName || '').toLowerCase();

      if (!isAllowedDomain(email)) {
        console.error('[KFGAuth] User domain not allowed:', email);
        return null;
      }

      return {
        name: profile.displayName || '',
        email: email,
        isAdmin: checkAdmin(email),
      };
    } catch (err) {
      console.error('[KFGAuth] Failed to fetch user profile:', err);
      return null;
    }
  }

  /**
   * Return the currently active MSAL account (first one), or null.
   * @returns {msal.AccountInfo|null}
   */
  function getActiveAccount() {
    if (!msalInstance) return null;

    // Prefer the active account if one is set
    var active = msalInstance.getActiveAccount();
    if (active) return active;

    // Otherwise pick the first account in the cache
    var accounts = msalInstance.getAllAccounts();
    if (accounts && accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
      return accounts[0];
    }

    return null;
  }

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------

  /**
   * Initialize MSAL with the provided configuration.
   *
   * @param {Object} config
   * @param {string} config.clientId   - Azure AD app (client) ID.
   * @param {string} config.tenantId   - Azure AD tenant ID.
   * @param {string} [config.redirectUri] - Redirect URI (defaults to window.location.origin).
   * @param {string[]} [config.adminEmails] - Additional admin email addresses.
   * @returns {Promise<void>}
   */
  async function init(config) {
    if (!config || !config.clientId || !config.tenantId) {
      throw new Error('[KFGAuth] init() requires config with clientId and tenantId.');
    }

    // Merge any extra admin emails from config
    if (Array.isArray(config.adminEmails)) {
      adminEmails = adminEmails.concat(config.adminEmails);
    }

    var msalConfig = {
      auth: {
        clientId: config.clientId,
        authority: 'https://login.microsoftonline.com/' + config.tenantId,
        redirectUri: config.redirectUri || window.location.origin,
        postLogoutRedirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: true,
      },
    };

    msalInstance = new msal.PublicClientApplication(msalConfig);

    // Handle redirect promise (for redirect-flow returns)
    try {
      var redirectResult = await msalInstance.handleRedirectPromise();

      if (redirectResult) {
        // User just returned from a redirect login
        msalInstance.setActiveAccount(redirectResult.account);
        var profile = await fetchUserProfile(redirectResult.accessToken);
        if (profile) {
          setAuthState(true, profile);
        } else {
          // Domain validation failed — log the user out
          console.warn('[KFGAuth] Domain validation failed after redirect login.');
          await logout();
        }
      } else {
        // Check if there is already a cached account
        var account = getActiveAccount();
        if (account) {
          // Silently re-acquire User.Read token to restore session without consent prompts
          try {
            var silentResult = await msalInstance.acquireTokenSilent({
              scopes: LOGIN_SCOPES,
              account: account,
            });
            var userProfile = await fetchUserProfile(silentResult.accessToken);
            if (userProfile) {
              setAuthState(true, userProfile);
            }
          } catch (silentErr) {
            console.warn('[KFGAuth] Silent session restore failed:', silentErr);
          }
        }
      }
    } catch (err) {
      console.error('[KFGAuth] Error handling redirect promise:', err);
    }
  }

  /**
   * Trigger an interactive login. Tries popup first, falls back to redirect.
   * @returns {Promise<{ name: string, email: string, isAdmin: boolean }|null>}
   */
  async function login() {
    if (!msalInstance) {
      throw new Error('[KFGAuth] Must call init() before login().');
    }

    // Request all scopes at login so user consents once — avoids separate popups later
    var loginRequest = {
      scopes: API_SCOPES,
    };

    try {
      // Attempt popup login with full scopes
      var response = await msalInstance.loginPopup(loginRequest);

      msalInstance.setActiveAccount(response.account);

      // Use the login token to fetch profile
      var profile = await fetchUserProfile(response.accessToken);

      if (!profile) {
        console.error('[KFGAuth] Login succeeded but domain validation failed.');
        await logout();
        return null;
      }

      setAuthState(true, profile);
      return profile;
    } catch (popupErr) {
      // If popup was blocked or failed, fall back to redirect
      if (
        popupErr instanceof msal.BrowserAuthError ||
        (popupErr.message && popupErr.message.indexOf('popup') !== -1)
      ) {
        console.warn('[KFGAuth] Popup blocked, falling back to redirect login.');
        try {
          await msalInstance.loginRedirect(loginRequest);
          // Redirect will navigate away — this code won't continue
          return null;
        } catch (redirectErr) {
          console.error('[KFGAuth] Redirect login failed:', redirectErr);
          return null;
        }
      }

      console.error('[KFGAuth] Login failed:', popupErr);
      return null;
    }
  }

  /**
   * Log out the current user, clear session, and redirect to origin.
   * @returns {Promise<void>}
   */
  async function logout() {
    if (!msalInstance) return;

    var account = getActiveAccount();

    // Clear internal state immediately
    setAuthState(false, null);

    try {
      await msalInstance.logoutPopup({
        account: account,
        postLogoutRedirectUri: window.location.origin,
      });
    } catch (err) {
      // If popup logout fails, try redirect
      console.warn('[KFGAuth] Popup logout failed, trying redirect:', err);
      try {
        await msalInstance.logoutRedirect({
          account: account,
        });
      } catch (redirectErr) {
        console.error('[KFGAuth] Redirect logout failed:', redirectErr);
        // As a last resort, just clear sessionStorage
        sessionStorage.clear();
      }
    }
  }

  /**
   * Get a valid access token for Microsoft Graph API calls.
   *
   * Attempts silent acquisition first (using cached refresh token).
   * Falls back to an interactive popup if silent fails.
   *
   * @returns {Promise<string|null>} The access token string, or null on failure.
   */
  async function getToken() {
    if (!msalInstance) {
      console.error('[KFGAuth] Must call init() before getToken().');
      return null;
    }

    var account = getActiveAccount();
    if (!account) {
      console.warn('[KFGAuth] No active account — user must log in first.');
      return null;
    }

    var tokenRequest = {
      scopes: API_SCOPES,
      account: account,
    };

    // 1. Try silent token acquisition
    try {
      var response = await msalInstance.acquireTokenSilent(tokenRequest);
      return response.accessToken;
    } catch (silentErr) {
      // 2. If interaction is required, try popup
      if (silentErr instanceof msal.InteractionRequiredAuthError) {
        console.warn('[KFGAuth] Silent token acquisition failed, trying popup.');
        try {
          var popupResponse = await msalInstance.acquireTokenPopup(tokenRequest);
          return popupResponse.accessToken;
        } catch (popupErr) {
          console.error('[KFGAuth] Interactive token acquisition failed:', popupErr);
          return null;
        }
      }

      console.error('[KFGAuth] Token acquisition failed:', silentErr);
      return null;
    }
  }

  /**
   * Get the current user's profile.
   * @returns {{ name: string, email: string, isAdmin: boolean }|null}
   */
  function getUser() {
    return currentUser;
  }

  /**
   * Check whether the user is currently authenticated.
   * @returns {boolean}
   */
  function isAuthenticated() {
    return authenticated;
  }

  /**
   * Check whether the current user is an admin.
   * @returns {boolean}
   */
  function isAdmin() {
    return currentUser !== null && currentUser.isAdmin === true;
  }

  /**
   * Register a callback that fires whenever the auth state changes.
   *
   * @param {function(boolean, object|null): void} cb
   *   Called with (isAuthenticated, userProfile).
   */
  function onAuthStateChange(cb) {
    if (typeof cb === 'function') {
      authStateCallbacks.push(cb);
    }
  }

  // ---------------------------------------------------------------------------
  // Expose public API on window.KFGAuth
  // ---------------------------------------------------------------------------

  window.KFGAuth = {
    init: init,
    login: login,
    logout: logout,
    getToken: getToken,
    getUser: getUser,
    isAuthenticated: isAuthenticated,
    isAdmin: isAdmin,
    onAuthStateChange: onAuthStateChange,
  };
})();

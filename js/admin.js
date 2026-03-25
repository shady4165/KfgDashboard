/**
 * KFG Executive Dashboard - Admin Panel Module
 *
 * Provides an in-dashboard admin interface for authorized admins to edit
 * chart configurations, table settings, refresh intervals, and manage
 * admin users.
 *
 * Dependencies:
 *   - window.KFGAuth  (authentication & admin check)
 *   - window.KFGConfig (configuration read/write)
 *
 * Usage:
 *   KFGAdmin.init();           // Initialize admin panel
 *   KFGAdmin.toggle();         // Toggle admin toolbar
 *   KFGAdmin.showToast('Saved', 'success');
 */
(function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Private state
  // ---------------------------------------------------------------------------

  /** Whether the admin panel has been initialized */
  let initialized = false;

  /** Whether admin mode is currently active (edit buttons visible) */
  let adminModeActive = false;

  /** Reference to the admin toolbar element */
  let toolbar = null;

  /** Reference to the toast container */
  let toastContainer = null;

  /** Reference to the settings panel */
  let settingsPanel = null;

  /** Reference to the settings backdrop */
  let settingsBackdrop = null;

  /** Reference to the currently open modal */
  let activeModal = null;

  /** Previously focused element before modal opened (for focus restore) */
  let previousFocus = null;

  /** Admin email for the current session */
  const ADMIN_EMAIL = 'shady.sabbagh@kidsfirstgroup.com';

  /** Audit log session storage key */
  const AUDIT_LOG_KEY = 'kfg_admin_audit_log';

  /** Maximum audit log entries to keep */
  const AUDIT_LOG_MAX = 50;

  // ---------------------------------------------------------------------------
  // CSS Injection
  // ---------------------------------------------------------------------------

  /**
   * Inject all admin panel styles into the document head.
   */
  function injectStyles() {
    if (document.getElementById('kfg-admin-styles')) return;

    const style = document.createElement('style');
    style.id = 'kfg-admin-styles';
    style.textContent = `
      /* ---- Admin Toolbar ---- */
      .kfg-admin-toolbar {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        height: 52px;
        background: #1e293b;
        color: #f1f5f9;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 16px;
        padding: 0 24px;
        z-index: 9000;
        box-shadow: 0 -2px 12px rgba(0,0,0,0.25);
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        font-size: 14px;
        transform: translateY(100%);
        transition: transform 0.3s ease;
      }
      .kfg-admin-toolbar.visible {
        transform: translateY(0);
      }
      .kfg-admin-toolbar-label {
        font-weight: 600;
        margin-right: 8px;
        user-select: none;
      }

      /* Toggle switch */
      .kfg-toggle {
        position: relative;
        display: inline-block;
        width: 40px;
        height: 22px;
        flex-shrink: 0;
      }
      .kfg-toggle input {
        opacity: 0;
        width: 0;
        height: 0;
      }
      .kfg-toggle-slider {
        position: absolute;
        inset: 0;
        background: #475569;
        border-radius: 22px;
        cursor: pointer;
        transition: background 0.2s;
      }
      .kfg-toggle-slider::before {
        content: '';
        position: absolute;
        left: 3px;
        top: 3px;
        width: 16px;
        height: 16px;
        background: #fff;
        border-radius: 50%;
        transition: transform 0.2s;
      }
      .kfg-toggle input:checked + .kfg-toggle-slider {
        background: #3b82f6;
      }
      .kfg-toggle input:checked + .kfg-toggle-slider::before {
        transform: translateX(18px);
      }
      .kfg-toggle input:focus-visible + .kfg-toggle-slider {
        outline: 2px solid #60a5fa;
        outline-offset: 2px;
      }

      /* Toolbar buttons */
      .kfg-admin-btn {
        padding: 6px 14px;
        border: 1px solid #475569;
        border-radius: 6px;
        background: transparent;
        color: #f1f5f9;
        cursor: pointer;
        font-size: 13px;
        font-family: inherit;
        transition: background 0.15s, border-color 0.15s;
        white-space: nowrap;
      }
      .kfg-admin-btn:hover {
        background: #334155;
        border-color: #64748b;
      }
      .kfg-admin-btn:focus-visible {
        outline: 2px solid #60a5fa;
        outline-offset: 2px;
      }
      .kfg-admin-btn--primary {
        background: #3b82f6;
        border-color: #3b82f6;
      }
      .kfg-admin-btn--primary:hover {
        background: #2563eb;
        border-color: #2563eb;
      }
      .kfg-admin-btn--danger {
        background: #ef4444;
        border-color: #ef4444;
      }
      .kfg-admin-btn--danger:hover {
        background: #dc2626;
        border-color: #dc2626;
      }
      .kfg-admin-btn--small {
        padding: 4px 10px;
        font-size: 12px;
      }

      /* ---- Edit buttons on chart cards ---- */
      .kfg-edit-btn {
        position: absolute;
        top: 8px;
        right: 8px;
        width: 30px;
        height: 30px;
        border: none;
        border-radius: 6px;
        background: rgba(59,130,246,0.9);
        color: #fff;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        z-index: 10;
        opacity: 0;
        transition: opacity 0.2s;
        box-shadow: 0 1px 4px rgba(0,0,0,0.2);
      }
      body.is-admin .kfg-edit-btn {
        opacity: 1;
      }
      .kfg-edit-btn:hover {
        background: rgba(37,99,235,1);
      }
      .kfg-edit-btn:focus-visible {
        outline: 2px solid #60a5fa;
        outline-offset: 2px;
      }

      /* Ensure chart cards are positioned for absolute children */
      body.is-admin .card {
        position: relative;
      }

      /* ---- Modal ---- */
      .kfg-modal-backdrop {
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.5);
        z-index: 9500;
        display: flex;
        align-items: center;
        justify-content: center;
        animation: kfgFadeIn 0.15s ease;
      }
      .kfg-modal {
        background: #fff;
        border-radius: 12px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.25);
        width: 90%;
        max-width: 560px;
        max-height: 85vh;
        overflow-y: auto;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        animation: kfgSlideUp 0.2s ease;
      }
      .kfg-modal-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 16px 20px;
        border-bottom: 1px solid #e2e8f0;
      }
      .kfg-modal-header h2 {
        margin: 0;
        font-size: 18px;
        color: #1e293b;
      }
      .kfg-modal-close {
        background: none;
        border: none;
        font-size: 22px;
        cursor: pointer;
        color: #64748b;
        padding: 4px;
        line-height: 1;
        border-radius: 4px;
      }
      .kfg-modal-close:hover {
        color: #1e293b;
        background: #f1f5f9;
      }
      .kfg-modal-close:focus-visible {
        outline: 2px solid #60a5fa;
        outline-offset: 2px;
      }
      .kfg-modal-body {
        padding: 20px;
      }
      .kfg-modal-footer {
        display: flex;
        align-items: center;
        justify-content: flex-end;
        gap: 8px;
        padding: 12px 20px;
        border-top: 1px solid #e2e8f0;
      }

      /* Form fields */
      .kfg-field {
        margin-bottom: 16px;
      }
      .kfg-field label {
        display: block;
        font-size: 13px;
        font-weight: 600;
        color: #374151;
        margin-bottom: 4px;
      }
      .kfg-field input[type="text"],
      .kfg-field input[type="number"],
      .kfg-field input[type="email"],
      .kfg-field input[type="url"],
      .kfg-field select {
        width: 100%;
        padding: 8px 10px;
        border: 1px solid #d1d5db;
        border-radius: 6px;
        font-size: 14px;
        font-family: inherit;
        color: #1e293b;
        background: #fff;
        box-sizing: border-box;
      }
      .kfg-field input:focus,
      .kfg-field select:focus {
        outline: none;
        border-color: #3b82f6;
        box-shadow: 0 0 0 3px rgba(59,130,246,0.15);
      }
      .kfg-field-row {
        display: flex;
        align-items: center;
        gap: 12px;
      }
      .kfg-field-hint {
        font-size: 12px;
        color: #6b7280;
        margin-top: 2px;
      }
      .kfg-color-row {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 4px;
      }
      .kfg-color-row input[type="color"] {
        width: 36px;
        height: 36px;
        border: 2px solid #d1d5db;
        border-radius: 6px;
        cursor: pointer;
        padding: 2px;
      }
      .kfg-color-row input[type="color"]:focus-visible {
        outline: 2px solid #60a5fa;
        outline-offset: 2px;
      }
      .kfg-slider-row {
        display: flex;
        align-items: center;
        gap: 12px;
      }
      .kfg-slider-row input[type="range"] {
        flex: 1;
      }
      .kfg-slider-value {
        min-width: 36px;
        text-align: right;
        font-size: 14px;
        color: #374151;
        font-variant-numeric: tabular-nums;
      }

      /* Checkbox list */
      .kfg-checkbox-list {
        max-height: 180px;
        overflow-y: auto;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 8px;
      }
      .kfg-checkbox-list label {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 4px 0;
        font-weight: 400;
        cursor: pointer;
      }

      /* ---- Settings Panel ---- */
      .kfg-settings-backdrop {
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.35);
        z-index: 9400;
        opacity: 0;
        transition: opacity 0.25s ease;
        pointer-events: none;
      }
      .kfg-settings-backdrop.visible {
        opacity: 1;
        pointer-events: auto;
      }
      .kfg-settings-panel {
        position: fixed;
        top: 0;
        right: 0;
        bottom: 0;
        width: 400px;
        max-width: 90vw;
        background: #fff;
        z-index: 9450;
        box-shadow: -4px 0 24px rgba(0,0,0,0.15);
        transform: translateX(100%);
        transition: transform 0.3s ease;
        display: flex;
        flex-direction: column;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      }
      .kfg-settings-panel.visible {
        transform: translateX(0);
      }
      .kfg-settings-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 16px 20px;
        border-bottom: 1px solid #e2e8f0;
        flex-shrink: 0;
      }
      .kfg-settings-header h2 {
        margin: 0;
        font-size: 18px;
        color: #1e293b;
      }
      .kfg-settings-body {
        flex: 1;
        overflow-y: auto;
        padding: 20px;
      }
      .kfg-settings-section {
        margin-bottom: 24px;
      }
      .kfg-settings-section h3 {
        font-size: 14px;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin: 0 0 12px 0;
      }

      /* Admin user list */
      .kfg-admin-list {
        list-style: none;
        margin: 0;
        padding: 0;
      }
      .kfg-admin-list li {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 8px 0;
        border-bottom: 1px solid #f1f5f9;
        font-size: 14px;
        color: #374151;
      }
      .kfg-admin-list li:last-child {
        border-bottom: none;
      }
      .kfg-admin-remove-btn {
        background: none;
        border: none;
        color: #ef4444;
        cursor: pointer;
        font-size: 18px;
        padding: 2px 6px;
        border-radius: 4px;
        line-height: 1;
      }
      .kfg-admin-remove-btn:hover {
        background: #fef2f2;
      }
      .kfg-admin-remove-btn:focus-visible {
        outline: 2px solid #60a5fa;
        outline-offset: 2px;
      }
      .kfg-add-admin-row {
        display: flex;
        gap: 8px;
        margin-top: 8px;
      }
      .kfg-add-admin-row input {
        flex: 1;
        padding: 6px 10px;
        border: 1px solid #d1d5db;
        border-radius: 6px;
        font-size: 14px;
        font-family: inherit;
      }
      .kfg-add-admin-row input:focus {
        outline: none;
        border-color: #3b82f6;
        box-shadow: 0 0 0 3px rgba(59,130,246,0.15);
      }

      /* Audit log */
      .kfg-audit-log {
        max-height: 240px;
        overflow-y: auto;
        font-size: 12px;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
      }
      .kfg-audit-log table {
        width: 100%;
        border-collapse: collapse;
      }
      .kfg-audit-log th {
        position: sticky;
        top: 0;
        background: #f8fafc;
        text-align: left;
        padding: 6px 8px;
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        border-bottom: 1px solid #e2e8f0;
      }
      .kfg-audit-log td {
        padding: 5px 8px;
        border-bottom: 1px solid #f1f5f9;
        color: #374151;
        vertical-align: top;
      }
      .kfg-audit-log tr:last-child td {
        border-bottom: none;
      }

      /* ---- Toast ---- */
      .kfg-toast-container {
        position: fixed;
        bottom: 64px;
        right: 20px;
        z-index: 9999;
        display: flex;
        flex-direction: column-reverse;
        gap: 8px;
        pointer-events: none;
      }
      .kfg-toast {
        padding: 10px 18px;
        border-radius: 8px;
        color: #fff;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        font-size: 14px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        pointer-events: auto;
        animation: kfgSlideIn 0.25s ease, kfgFadeOut 0.3s ease 2.7s forwards;
        max-width: 340px;
        word-wrap: break-word;
      }
      .kfg-toast--success { background: #16a34a; }
      .kfg-toast--error   { background: #dc2626; }
      .kfg-toast--info    { background: #2563eb; }

      /* Preview section */
      .kfg-preview {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px;
        margin-top: 12px;
        font-size: 13px;
        color: #475569;
      }
      .kfg-preview-title {
        font-weight: 600;
        color: #374151;
        margin-bottom: 6px;
      }
      .kfg-preview-item {
        margin: 2px 0;
      }

      /* ---- Animations ---- */
      @keyframes kfgFadeIn {
        from { opacity: 0; }
        to   { opacity: 1; }
      }
      @keyframes kfgSlideUp {
        from { transform: translateY(20px); opacity: 0; }
        to   { transform: translateY(0); opacity: 1; }
      }
      @keyframes kfgSlideIn {
        from { transform: translateX(100%); opacity: 0; }
        to   { transform: translateX(0); opacity: 1; }
      }
      @keyframes kfgFadeOut {
        from { opacity: 1; }
        to   { opacity: 0; }
      }

      /* File input styling */
      .kfg-file-input-label {
        display: inline-block;
        padding: 6px 14px;
        border: 1px solid #475569;
        border-radius: 6px;
        background: transparent;
        color: #f1f5f9;
        cursor: pointer;
        font-size: 13px;
        font-family: inherit;
        transition: background 0.15s;
      }
      .kfg-file-input-label:hover {
        background: #334155;
      }
      .kfg-file-hidden {
        display: none;
      }
    `;
    document.head.appendChild(style);
  }

  // ---------------------------------------------------------------------------
  // DOM Helpers
  // ---------------------------------------------------------------------------

  /**
   * Create an element with attributes and children.
   * @param {string} tag
   * @param {Object} [attrs]
   * @param {...(Node|string)} children
   * @returns {HTMLElement}
   */
  function el(tag, attrs, ...children) {
    const element = document.createElement(tag);
    if (attrs) {
      for (const [key, value] of Object.entries(attrs)) {
        if (key === 'className') {
          element.className = value;
        } else if (key === 'textContent') {
          element.textContent = value;
        } else if (key.startsWith('on') && typeof value === 'function') {
          element.addEventListener(key.slice(2).toLowerCase(), value);
        } else if (key === 'style' && typeof value === 'object') {
          Object.assign(element.style, value);
        } else {
          element.setAttribute(key, value);
        }
      }
    }
    for (const child of children) {
      if (typeof child === 'string') {
        element.appendChild(document.createTextNode(child));
      } else if (child instanceof Node) {
        element.appendChild(child);
      }
    }
    return element;
  }

  /**
   * Check if the current user is an admin.
   * @returns {boolean}
   */
  function isAdmin() {
    if (typeof KFGAuth !== 'undefined' && KFGAuth.isAdmin) {
      return KFGAuth.isAdmin();
    }
    return false;
  }

  /**
   * Get the current user email.
   * @returns {string}
   */
  function getCurrentUserEmail() {
    if (typeof KFGAuth !== 'undefined' && KFGAuth.getUser) {
      const user = KFGAuth.getUser();
      return user ? user.email : ADMIN_EMAIL;
    }
    return ADMIN_EMAIL;
  }

  // ---------------------------------------------------------------------------
  // Audit Log
  // ---------------------------------------------------------------------------

  /**
   * Log an admin action to sessionStorage.
   * @param {string} action - Description of the action.
   */
  function logAuditAction(action) {
    try {
      let log = JSON.parse(sessionStorage.getItem(AUDIT_LOG_KEY) || '[]');
      log.unshift({
        timestamp: new Date().toISOString(),
        user: getCurrentUserEmail(),
        action: action,
      });
      if (log.length > AUDIT_LOG_MAX) {
        log = log.slice(0, AUDIT_LOG_MAX);
      }
      sessionStorage.setItem(AUDIT_LOG_KEY, JSON.stringify(log));
    } catch (err) {
      console.error('[KFGAdmin] Failed to write audit log:', err);
    }
  }

  /**
   * Get audit log entries.
   * @param {number} [limit=20]
   * @returns {Array<{ timestamp: string, user: string, action: string }>}
   */
  function getAuditLog(limit) {
    try {
      const log = JSON.parse(sessionStorage.getItem(AUDIT_LOG_KEY) || '[]');
      return log.slice(0, limit || 20);
    } catch (err) {
      return [];
    }
  }

  // ---------------------------------------------------------------------------
  // Toast Notifications
  // ---------------------------------------------------------------------------

  /**
   * Create the toast container if it does not exist.
   */
  function ensureToastContainer() {
    if (toastContainer) return;
    toastContainer = el('div', {
      className: 'kfg-toast-container',
      'aria-live': 'polite',
      'aria-label': 'Notifications',
    });
    document.body.appendChild(toastContainer);
  }

  /**
   * Show a toast notification.
   * @param {string} message
   * @param {'success'|'error'|'info'} [type='info']
   */
  function showToast(message, type) {
    ensureToastContainer();
    const validTypes = ['success', 'error', 'info'];
    const toastType = validTypes.includes(type) ? type : 'info';

    const toast = el('div', {
      className: 'kfg-toast kfg-toast--' + toastType,
      role: 'status',
      'aria-label': toastType + ' notification',
      textContent: message,
    });

    toastContainer.appendChild(toast);

    setTimeout(function () {
      if (toast.parentNode) {
        toast.parentNode.removeChild(toast);
      }
    }, 3000);
  }

  // ---------------------------------------------------------------------------
  // Modal System
  // ---------------------------------------------------------------------------

  /**
   * Get all focusable elements within a container.
   * @param {HTMLElement} container
   * @returns {HTMLElement[]}
   */
  function getFocusableElements(container) {
    const selectors = [
      'a[href]', 'button:not([disabled])', 'input:not([disabled])',
      'select:not([disabled])', 'textarea:not([disabled])',
      '[tabindex]:not([tabindex="-1"])',
    ];
    return Array.from(container.querySelectorAll(selectors.join(',')));
  }

  /**
   * Handle tab trapping within a modal.
   * @param {KeyboardEvent} e
   * @param {HTMLElement} container
   */
  function trapTabFocus(e, container) {
    if (e.key !== 'Tab') return;

    const focusable = getFocusableElements(container);
    if (focusable.length === 0) {
      e.preventDefault();
      return;
    }

    const first = focusable[0];
    const last = focusable[focusable.length - 1];

    if (e.shiftKey) {
      if (document.activeElement === first) {
        e.preventDefault();
        last.focus();
      }
    } else {
      if (document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    }
  }

  /**
   * Open a modal with the given title and body content.
   * @param {string} title
   * @param {HTMLElement} bodyContent
   * @param {HTMLElement} [footerContent]
   * @returns {{ backdrop: HTMLElement, modal: HTMLElement, close: Function }}
   */
  function openModal(title, bodyContent, footerContent) {
    closeModal();
    previousFocus = document.activeElement;

    const closeBtn = el('button', {
      className: 'kfg-modal-close',
      'aria-label': 'Close dialog',
      textContent: '\u00D7',
    });

    const header = el('div', { className: 'kfg-modal-header' },
      el('h2', { textContent: title }),
      closeBtn
    );

    const body = el('div', { className: 'kfg-modal-body' }, bodyContent);

    const modalEl = el('div', {
      className: 'kfg-modal',
      role: 'dialog',
      'aria-modal': 'true',
      'aria-label': title,
    }, header, body);

    if (footerContent) {
      const footer = el('div', { className: 'kfg-modal-footer' }, footerContent);
      modalEl.appendChild(footer);
    }

    const backdrop = el('div', { className: 'kfg-modal-backdrop' }, modalEl);

    function close() {
      if (backdrop.parentNode) {
        backdrop.parentNode.removeChild(backdrop);
      }
      activeModal = null;
      document.removeEventListener('keydown', keyHandler);
      if (previousFocus && previousFocus.focus) {
        previousFocus.focus();
      }
    }

    function keyHandler(e) {
      if (e.key === 'Escape') {
        close();
      }
      trapTabFocus(e, modalEl);
    }

    closeBtn.addEventListener('click', close);
    backdrop.addEventListener('click', function (e) {
      if (e.target === backdrop) {
        close();
      }
    });

    document.addEventListener('keydown', keyHandler);
    document.body.appendChild(backdrop);
    activeModal = { backdrop: backdrop, modal: modalEl, close: close };

    // Focus the close button
    closeBtn.focus();

    return activeModal;
  }

  /**
   * Close the currently active modal.
   */
  function closeModal() {
    if (activeModal) {
      activeModal.close();
      activeModal = null;
    }
  }

  // ---------------------------------------------------------------------------
  // Admin Toolbar
  // ---------------------------------------------------------------------------

  /**
   * Create the floating admin toolbar.
   */
  function createToolbar() {
    if (toolbar) return;

    // Toggle switch
    const toggleInput = el('input', {
      type: 'checkbox',
      id: 'kfg-admin-mode-toggle',
      'aria-label': 'Toggle admin mode',
    });
    toggleInput.addEventListener('change', function () {
      setAdminMode(toggleInput.checked);
    });

    const toggleSlider = el('span', { className: 'kfg-toggle-slider' });
    const toggleLabel = el('label', { className: 'kfg-toggle', 'for': 'kfg-admin-mode-toggle' },
      toggleInput,
      toggleSlider
    );

    // Buttons
    const settingsBtn = el('button', {
      className: 'kfg-admin-btn',
      'aria-label': 'Open settings panel',
      textContent: 'Settings',
      onClick: function () { openSettings(); },
    });

    const exportBtn = el('button', {
      className: 'kfg-admin-btn',
      'aria-label': 'Export configuration',
      textContent: 'Export Config',
      onClick: function () { exportConfig(); },
    });

    // Import button with hidden file input
    const importInput = el('input', {
      type: 'file',
      accept: '.json',
      className: 'kfg-file-hidden',
      'aria-label': 'Import configuration file',
      onChange: function (e) { handleImportConfig(e); },
    });

    const importBtn = el('button', {
      className: 'kfg-admin-btn',
      'aria-label': 'Import configuration',
      textContent: 'Import Config',
      onClick: function () { importInput.click(); },
    });

    toolbar = el('div', {
      className: 'kfg-admin-toolbar',
      role: 'toolbar',
      'aria-label': 'Admin toolbar',
    },
      el('span', { className: 'kfg-admin-toolbar-label', textContent: 'Admin Mode' }),
      toggleLabel,
      el('span', { style: { width: '1px', height: '24px', background: '#475569' } }),
      settingsBtn,
      exportBtn,
      importBtn,
      importInput
    );

    document.body.appendChild(toolbar);
  }

  /**
   * Toggle toolbar visibility.
   */
  function toggleToolbar() {
    if (!toolbar) return;
    toolbar.classList.toggle('visible');
  }

  /**
   * Show the toolbar.
   */
  function showToolbar() {
    if (!toolbar) return;
    toolbar.classList.add('visible');
  }

  /**
   * Hide the toolbar.
   */
  function hideToolbar() {
    if (!toolbar) return;
    toolbar.classList.remove('visible');
  }

  // ---------------------------------------------------------------------------
  // Admin Mode (edit buttons on charts/tables)
  // ---------------------------------------------------------------------------

  /**
   * Enable or disable admin mode.
   * @param {boolean} active
   */
  function setAdminMode(active) {
    adminModeActive = active;
    if (active) {
      document.body.classList.add('is-admin');
      injectEditButtons();
      logAuditAction('Admin mode enabled');
    } else {
      document.body.classList.remove('is-admin');
      removeEditButtons();
      logAuditAction('Admin mode disabled');
    }
  }

  /**
   * Inject edit (pencil) buttons onto each chart card.
   */
  function injectEditButtons() {
    removeEditButtons();

    // Find all canvas elements (charts) and add edit buttons to their card parents
    const canvases = document.querySelectorAll('canvas');
    canvases.forEach(function (canvas) {
      const card = canvas.closest('.card');
      if (!card) return;

      const chartId = canvas.id;
      if (!chartId) return;

      const editBtn = el('button', {
        className: 'kfg-edit-btn',
        'aria-label': 'Edit chart configuration for ' + chartId,
        'data-kfg-edit': 'chart',
        'data-kfg-target': chartId,
        textContent: '\u270E',
        onClick: function () { openChartEditor(chartId); },
      });

      card.appendChild(editBtn);
    });

    // Find all tables and add edit buttons to their card parents
    const tables = document.querySelectorAll('table');
    tables.forEach(function (table) {
      const card = table.closest('.card');
      if (!card) return;
      // Skip if already has an edit button
      if (card.querySelector('.kfg-edit-btn')) return;

      const tableId = table.id || card.id || ('table-' + Math.random().toString(36).slice(2, 8));
      if (!table.id) table.id = tableId;

      const editBtn = el('button', {
        className: 'kfg-edit-btn',
        'aria-label': 'Edit table configuration for ' + tableId,
        'data-kfg-edit': 'table',
        'data-kfg-target': tableId,
        textContent: '\u270E',
        onClick: function () { openTableEditor(tableId); },
      });

      card.appendChild(editBtn);
    });
  }

  /**
   * Remove all injected edit buttons.
   */
  function removeEditButtons() {
    const buttons = document.querySelectorAll('.kfg-edit-btn');
    buttons.forEach(function (btn) {
      btn.parentNode.removeChild(btn);
    });
  }

  // ---------------------------------------------------------------------------
  // Chart Editor
  // ---------------------------------------------------------------------------

  /**
   * Open the chart configuration editor for a specific chart.
   * @param {string} chartId - The canvas element ID.
   */
  function openChartEditor(chartId) {
    // Get current config from KFGConfig
    let currentConfig = {};
    if (typeof KFGConfig !== 'undefined' && KFGConfig.getChartConfig) {
      currentConfig = KFGConfig.getChartConfig(chartId) || {};
    }

    const defaults = {
      title: currentConfig.title || '',
      subtitle: currentConfig.subtitle || '',
      type: currentConfig.type || 'bar',
      legendDisplay: currentConfig.legendDisplay !== false,
      legendPosition: currentConfig.legendPosition || 'top',
      colors: currentConfig.colors || ['#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899'],
      borderRadius: currentConfig.borderRadius || 0,
      cutout: currentConfig.cutout || 50,
      indexAxis: currentConfig.indexAxis || 'x',
    };

    // Title
    const titleInput = el('input', {
      type: 'text',
      id: 'kfg-chart-title',
      value: defaults.title,
      'aria-label': 'Chart title',
    });
    const titleField = el('div', { className: 'kfg-field' },
      el('label', { 'for': 'kfg-chart-title', textContent: 'Title' }),
      titleInput
    );

    // Subtitle
    const subtitleInput = el('input', {
      type: 'text',
      id: 'kfg-chart-subtitle',
      value: defaults.subtitle,
      'aria-label': 'Chart subtitle',
    });
    const subtitleField = el('div', { className: 'kfg-field' },
      el('label', { 'for': 'kfg-chart-subtitle', textContent: 'Subtitle' }),
      subtitleInput
    );

    // Chart Type
    const typeSelect = el('select', {
      id: 'kfg-chart-type',
      'aria-label': 'Chart type',
    });
    ['bar', 'doughnut', 'line', 'pie'].forEach(function (t) {
      const opt = el('option', { value: t, textContent: t.charAt(0).toUpperCase() + t.slice(1) });
      if (t === defaults.type) opt.selected = true;
      typeSelect.appendChild(opt);
    });
    const typeField = el('div', { className: 'kfg-field' },
      el('label', { 'for': 'kfg-chart-type', textContent: 'Chart Type' }),
      typeSelect
    );

    // Legend toggle
    const legendToggleInput = el('input', {
      type: 'checkbox',
      id: 'kfg-chart-legend-display',
      'aria-label': 'Show legend',
    });
    legendToggleInput.checked = defaults.legendDisplay;

    const legendToggle = el('label', { className: 'kfg-toggle', 'for': 'kfg-chart-legend-display' },
      legendToggleInput,
      el('span', { className: 'kfg-toggle-slider' })
    );

    // Legend position
    const legendPosSelect = el('select', {
      id: 'kfg-chart-legend-pos',
      'aria-label': 'Legend position',
    });
    ['top', 'right', 'bottom', 'left'].forEach(function (pos) {
      const opt = el('option', { value: pos, textContent: pos.charAt(0).toUpperCase() + pos.slice(1) });
      if (pos === defaults.legendPosition) opt.selected = true;
      legendPosSelect.appendChild(opt);
    });

    const legendField = el('div', { className: 'kfg-field' },
      el('label', { textContent: 'Legend' }),
      el('div', { className: 'kfg-field-row' },
        el('span', { textContent: 'Show:' }),
        legendToggle,
        el('span', { textContent: 'Position:' }),
        legendPosSelect
      )
    );

    // Colors
    const colorRow = el('div', { className: 'kfg-color-row' });
    defaults.colors.forEach(function (color, index) {
      const colorInput = el('input', {
        type: 'color',
        value: color,
        'aria-label': 'Dataset color ' + (index + 1),
        'data-color-index': String(index),
      });
      colorRow.appendChild(colorInput);
    });

    // Add color button
    const addColorBtn = el('button', {
      className: 'kfg-admin-btn kfg-admin-btn--small',
      'aria-label': 'Add another color',
      textContent: '+ Color',
      onClick: function () {
        const newColor = el('input', {
          type: 'color',
          value: '#6366f1',
          'aria-label': 'Dataset color ' + (colorRow.children.length + 1),
          'data-color-index': String(colorRow.children.length),
        });
        colorRow.appendChild(newColor);
      },
    });

    const colorsField = el('div', { className: 'kfg-field' },
      el('label', { textContent: 'Colors' }),
      colorRow,
      el('div', { style: { marginTop: '6px' } }, addColorBtn)
    );

    // Border Radius (bar charts)
    const borderRadiusValue = el('span', {
      className: 'kfg-slider-value',
      textContent: String(defaults.borderRadius),
    });
    const borderRadiusInput = el('input', {
      type: 'range',
      min: '0',
      max: '20',
      value: String(defaults.borderRadius),
      'aria-label': 'Border radius',
      onInput: function () {
        borderRadiusValue.textContent = borderRadiusInput.value;
      },
    });
    const borderRadiusField = el('div', { className: 'kfg-field' },
      el('label', { textContent: 'Border Radius (bar charts)' }),
      el('div', { className: 'kfg-slider-row' },
        borderRadiusInput,
        borderRadiusValue
      )
    );

    // Cutout % (doughnut charts)
    const cutoutValue = el('span', {
      className: 'kfg-slider-value',
      textContent: defaults.cutout + '%',
    });
    const cutoutInput = el('input', {
      type: 'range',
      min: '0',
      max: '90',
      value: String(defaults.cutout),
      'aria-label': 'Cutout percentage',
      onInput: function () {
        cutoutValue.textContent = cutoutInput.value + '%';
      },
    });
    const cutoutField = el('div', { className: 'kfg-field' },
      el('label', { textContent: 'Cutout % (doughnut charts)' }),
      el('div', { className: 'kfg-slider-row' },
        cutoutInput,
        cutoutValue
      )
    );

    // Index Axis (bar charts)
    const indexAxisSelect = el('select', {
      id: 'kfg-chart-index-axis',
      'aria-label': 'Index axis direction',
    });
    [
      { value: 'x', label: 'Vertical' },
      { value: 'y', label: 'Horizontal' },
    ].forEach(function (opt) {
      const option = el('option', { value: opt.value, textContent: opt.label });
      if (opt.value === defaults.indexAxis) option.selected = true;
      indexAxisSelect.appendChild(option);
    });
    const indexAxisField = el('div', { className: 'kfg-field' },
      el('label', { 'for': 'kfg-chart-index-axis', textContent: 'Bar Direction' }),
      indexAxisSelect
    );

    // Preview
    const previewEl = el('div', { className: 'kfg-preview' },
      el('div', { className: 'kfg-preview-title', textContent: 'Current Values' }),
      el('div', { className: 'kfg-preview-item', textContent: 'Chart ID: ' + chartId }),
      el('div', { className: 'kfg-preview-item', textContent: 'Type: ' + defaults.type }),
      el('div', { className: 'kfg-preview-item', textContent: 'Legend: ' + (defaults.legendDisplay ? 'visible' : 'hidden') + ' (' + defaults.legendPosition + ')' }),
      el('div', { className: 'kfg-preview-item', textContent: 'Colors: ' + defaults.colors.length + ' defined' })
    );

    // Assemble body
    const bodyContent = el('div', {},
      titleField,
      subtitleField,
      typeField,
      legendField,
      colorsField,
      borderRadiusField,
      cutoutField,
      indexAxisField,
      previewEl
    );

    // Footer buttons
    const saveBtn = el('button', {
      className: 'kfg-admin-btn kfg-admin-btn--primary',
      'aria-label': 'Save chart configuration',
      textContent: 'Save',
      onClick: function () {
        saveChartConfig(chartId);
      },
    });

    const cancelBtn = el('button', {
      className: 'kfg-admin-btn',
      'aria-label': 'Cancel editing',
      textContent: 'Cancel',
      onClick: function () { closeModal(); },
    });

    const resetBtn = el('button', {
      className: 'kfg-admin-btn kfg-admin-btn--danger',
      'aria-label': 'Reset chart to default configuration',
      textContent: 'Reset to Default',
      onClick: function () {
        if (typeof KFGConfig !== 'undefined' && KFGConfig.resetChartConfig) {
          KFGConfig.resetChartConfig(chartId);
          logAuditAction('Reset chart "' + chartId + '" to defaults');
          showToast('Chart reset to defaults', 'success');
          closeModal();
          rebuildChart(chartId);
        }
      },
    });

    const footer = el('div', {}, resetBtn, cancelBtn, saveBtn);
    // Override footer to use flex layout
    footer.style.display = 'contents';

    openModal('Edit Chart: ' + chartId, bodyContent, footer);
  }

  /**
   * Gather values from the chart editor form and save.
   * @param {string} chartId
   */
  function saveChartConfig(chartId) {
    if (!activeModal) return;
    const modal = activeModal.modal;

    // Collect colors from color inputs
    const colorInputs = modal.querySelectorAll('.kfg-color-row input[type="color"]');
    const colors = Array.from(colorInputs).map(function (input) {
      return input.value;
    });

    const newConfig = {
      title: modal.querySelector('#kfg-chart-title').value,
      subtitle: modal.querySelector('#kfg-chart-subtitle').value,
      type: modal.querySelector('#kfg-chart-type').value,
      legendDisplay: modal.querySelector('#kfg-chart-legend-display').checked,
      legendPosition: modal.querySelector('#kfg-chart-legend-pos').value,
      colors: colors,
      borderRadius: parseInt(modal.querySelector('input[type="range"][aria-label="Border radius"]').value, 10),
      cutout: parseInt(modal.querySelector('input[type="range"][aria-label="Cutout percentage"]').value, 10),
      indexAxis: modal.querySelector('#kfg-chart-index-axis').value,
    };

    // Save via KFGConfig
    if (typeof KFGConfig !== 'undefined' && KFGConfig.setChartConfig) {
      KFGConfig.setChartConfig(chartId, newConfig);
    }

    logAuditAction('Updated chart configuration for "' + chartId + '"');
    showToast('Chart configuration saved', 'success');
    closeModal();
    rebuildChart(chartId);
  }

  /**
   * Trigger a rebuild of the chart with the given ID.
   * Looks for common dashboard rebuild patterns.
   * @param {string} chartId
   */
  function rebuildChart(chartId) {
    // Try various common dashboard rebuild patterns
    if (typeof KFGDashboard !== 'undefined') {
      if (KFGDashboard.rebuildChart) {
        KFGDashboard.rebuildChart(chartId);
        return;
      }
      if (KFGDashboard.refresh) {
        KFGDashboard.refresh();
        return;
      }
    }
    // Dispatch a custom event for the dashboard to listen to
    window.dispatchEvent(new CustomEvent('kfg:chart-config-changed', {
      detail: { chartId: chartId },
    }));
  }

  // ---------------------------------------------------------------------------
  // Table Editor
  // ---------------------------------------------------------------------------

  /**
   * Open the table configuration editor for a specific table.
   * @param {string} tableId - The table element ID.
   */
  function openTableEditor(tableId) {
    // Get current config from KFGConfig
    let currentConfig = {};
    if (typeof KFGConfig !== 'undefined' && KFGConfig.getTableConfig) {
      currentConfig = KFGConfig.getTableConfig(tableId) || {};
    }

    const defaults = {
      sortable: currentConfig.sortable !== false,
      filterable: currentConfig.filterable !== false,
      pageSize: currentConfig.pageSize || 0,
      visibleColumns: currentConfig.visibleColumns || [],
    };

    // Sortable toggle
    const sortableInput = el('input', {
      type: 'checkbox',
      id: 'kfg-table-sortable',
      'aria-label': 'Enable sorting',
    });
    sortableInput.checked = defaults.sortable;

    const sortableField = el('div', { className: 'kfg-field' },
      el('label', { textContent: 'Sortable' }),
      el('div', { className: 'kfg-field-row' },
        el('label', { className: 'kfg-toggle', 'for': 'kfg-table-sortable' },
          sortableInput,
          el('span', { className: 'kfg-toggle-slider' })
        ),
        el('span', { className: 'kfg-field-hint', textContent: 'Allow users to sort by column headers' })
      )
    );

    // Filterable toggle
    const filterableInput = el('input', {
      type: 'checkbox',
      id: 'kfg-table-filterable',
      'aria-label': 'Enable filtering',
    });
    filterableInput.checked = defaults.filterable;

    const filterableField = el('div', { className: 'kfg-field' },
      el('label', { textContent: 'Filterable' }),
      el('div', { className: 'kfg-field-row' },
        el('label', { className: 'kfg-toggle', 'for': 'kfg-table-filterable' },
          filterableInput,
          el('span', { className: 'kfg-toggle-slider' })
        ),
        el('span', { className: 'kfg-field-hint', textContent: 'Show search/filter inputs' })
      )
    );

    // Page Size
    const pageSizeInput = el('input', {
      type: 'number',
      id: 'kfg-table-pagesize',
      min: '0',
      value: String(defaults.pageSize),
      'aria-label': 'Page size',
    });
    const pageSizeField = el('div', { className: 'kfg-field' },
      el('label', { 'for': 'kfg-table-pagesize', textContent: 'Page Size' }),
      pageSizeInput,
      el('div', { className: 'kfg-field-hint', textContent: 'Set to 0 to show all rows' })
    );

    // Visible Columns - detect from actual table
    const table = document.getElementById(tableId);
    const columnsContainer = el('div', { className: 'kfg-checkbox-list' });

    if (table) {
      const headerCells = table.querySelectorAll('thead th, thead td');
      headerCells.forEach(function (th, index) {
        const colName = th.textContent.trim() || 'Column ' + (index + 1);
        const colId = 'kfg-col-' + index;
        const isVisible = defaults.visibleColumns.length === 0 ||
          defaults.visibleColumns.includes(colName) ||
          defaults.visibleColumns.includes(index);

        const checkbox = el('input', {
          type: 'checkbox',
          id: colId,
          value: colName,
          'aria-label': 'Show column: ' + colName,
        });
        checkbox.checked = isVisible;

        columnsContainer.appendChild(
          el('label', { 'for': colId },
            checkbox,
            document.createTextNode(colName)
          )
        );
      });
    }

    const columnsField = el('div', { className: 'kfg-field' },
      el('label', { textContent: 'Visible Columns' }),
      columnsContainer
    );

    // Assemble body
    const bodyContent = el('div', {},
      sortableField,
      filterableField,
      pageSizeField,
      columnsField
    );

    // Footer
    const saveBtn = el('button', {
      className: 'kfg-admin-btn kfg-admin-btn--primary',
      'aria-label': 'Save table configuration',
      textContent: 'Save',
      onClick: function () {
        saveTableConfig(tableId);
      },
    });

    const cancelBtn = el('button', {
      className: 'kfg-admin-btn',
      'aria-label': 'Cancel editing',
      textContent: 'Cancel',
      onClick: function () { closeModal(); },
    });

    const footer = el('div', { style: { display: 'contents' } }, cancelBtn, saveBtn);

    openModal('Edit Table: ' + tableId, bodyContent, footer);
  }

  /**
   * Gather values from the table editor form and save.
   * @param {string} tableId
   */
  function saveTableConfig(tableId) {
    if (!activeModal) return;
    const modal = activeModal.modal;

    // Collect visible columns
    const checkboxes = modal.querySelectorAll('.kfg-checkbox-list input[type="checkbox"]');
    const visibleColumns = [];
    checkboxes.forEach(function (cb) {
      if (cb.checked) {
        visibleColumns.push(cb.value);
      }
    });

    const newConfig = {
      sortable: modal.querySelector('#kfg-table-sortable').checked,
      filterable: modal.querySelector('#kfg-table-filterable').checked,
      pageSize: parseInt(modal.querySelector('#kfg-table-pagesize').value, 10) || 0,
      visibleColumns: visibleColumns,
    };

    // Save via KFGConfig
    if (typeof KFGConfig !== 'undefined' && KFGConfig.setTableConfig) {
      KFGConfig.setTableConfig(tableId, newConfig);
    }

    logAuditAction('Updated table configuration for "' + tableId + '"');
    showToast('Table configuration saved', 'success');
    closeModal();

    // Dispatch custom event for dashboard to rebuild the table
    window.dispatchEvent(new CustomEvent('kfg:table-config-changed', {
      detail: { tableId: tableId },
    }));
  }

  // ---------------------------------------------------------------------------
  // Settings Panel
  // ---------------------------------------------------------------------------

  /**
   * Create the settings panel DOM (slide-in from right).
   */
  function createSettingsPanel() {
    if (settingsPanel) return;

    // Backdrop
    settingsBackdrop = el('div', {
      className: 'kfg-settings-backdrop',
      'aria-hidden': 'true',
      onClick: function () { closeSettings(); },
    });

    // Close button
    const closeBtn = el('button', {
      className: 'kfg-modal-close',
      'aria-label': 'Close settings panel',
      textContent: '\u00D7',
      onClick: function () { closeSettings(); },
    });

    // Header
    const header = el('div', { className: 'kfg-settings-header' },
      el('h2', { textContent: 'Settings' }),
      closeBtn
    );

    // Body
    const body = el('div', { className: 'kfg-settings-body' });

    settingsPanel = el('div', {
      className: 'kfg-settings-panel',
      role: 'dialog',
      'aria-modal': 'true',
      'aria-label': 'Settings panel',
    }, header, body);

    document.body.appendChild(settingsBackdrop);
    document.body.appendChild(settingsPanel);
  }

  /**
   * Open the settings panel and populate its content.
   */
  function openSettings() {
    createSettingsPanel();
    const body = settingsPanel.querySelector('.kfg-settings-body');

    // Clear previous content
    while (body.firstChild) {
      body.removeChild(body.firstChild);
    }

    // --- General Settings ---
    const generalSection = el('div', { className: 'kfg-settings-section' },
      el('h3', { textContent: 'General Settings' })
    );

    // Refresh Interval
    let currentInterval = 5;
    if (typeof KFGConfig !== 'undefined' && KFGConfig.getRefreshInterval) {
      currentInterval = KFGConfig.getRefreshInterval() || 5;
    }

    const refreshInput = el('input', {
      type: 'number',
      id: 'kfg-settings-refresh',
      min: '1',
      max: '60',
      value: String(currentInterval),
      'aria-label': 'Refresh interval in minutes',
    });
    refreshInput.addEventListener('change', function () {
      const val = parseInt(refreshInput.value, 10);
      if (val > 0 && typeof KFGConfig !== 'undefined' && KFGConfig.setRefreshInterval) {
        KFGConfig.setRefreshInterval(val);
        logAuditAction('Changed refresh interval to ' + val + ' minutes');
        showToast('Refresh interval updated to ' + val + ' minutes', 'success');
      }
    });

    generalSection.appendChild(el('div', { className: 'kfg-field' },
      el('label', { 'for': 'kfg-settings-refresh', textContent: 'Refresh Interval (minutes)' }),
      refreshInput
    ));

    // SharePoint Site URL (read-only)
    let spUrl = '';
    if (typeof KFGConfig !== 'undefined' && KFGConfig.getSharePointUrl) {
      spUrl = KFGConfig.getSharePointUrl() || '';
    }

    const spInput = el('input', {
      type: 'url',
      id: 'kfg-settings-sp-url',
      value: spUrl,
      readonly: 'readonly',
      'aria-label': 'SharePoint site URL',
      style: { background: '#f1f5f9', color: '#64748b' },
    });

    generalSection.appendChild(el('div', { className: 'kfg-field' },
      el('label', { 'for': 'kfg-settings-sp-url', textContent: 'SharePoint Site URL' }),
      spInput,
      el('div', { className: 'kfg-field-hint', textContent: 'Read-only. Change in configuration file.' })
    ));

    body.appendChild(generalSection);

    // --- Admin Users ---
    const adminSection = el('div', { className: 'kfg-settings-section' },
      el('h3', { textContent: 'Admin Users' })
    );

    const adminList = el('ul', { className: 'kfg-admin-list' });

    function renderAdminList() {
      while (adminList.firstChild) {
        adminList.removeChild(adminList.firstChild);
      }

      let admins = [ADMIN_EMAIL];
      if (typeof KFGConfig !== 'undefined' && KFGConfig.getAdminEmails) {
        admins = KFGConfig.getAdminEmails() || [ADMIN_EMAIL];
      }

      admins.forEach(function (email) {
        const removeBtn = el('button', {
          className: 'kfg-admin-remove-btn',
          'aria-label': 'Remove admin: ' + email,
          textContent: '\u00D7',
          onClick: function () {
            if (email.toLowerCase() === ADMIN_EMAIL.toLowerCase()) {
              showToast('Cannot remove the primary admin', 'error');
              return;
            }
            if (typeof KFGConfig !== 'undefined' && KFGConfig.removeAdminEmail) {
              KFGConfig.removeAdminEmail(email);
              logAuditAction('Removed admin user: ' + email);
              showToast('Removed admin: ' + email, 'success');
              renderAdminList();
            }
          },
        });

        adminList.appendChild(el('li', {},
          el('span', { textContent: email }),
          removeBtn
        ));
      });
    }

    renderAdminList();
    adminSection.appendChild(adminList);

    // Add admin input
    const addAdminInput = el('input', {
      type: 'email',
      placeholder: 'admin@kidsfirstgroup.com',
      'aria-label': 'New admin email address',
    });

    const addAdminBtn = el('button', {
      className: 'kfg-admin-btn kfg-admin-btn--primary kfg-admin-btn--small',
      'aria-label': 'Add admin user',
      textContent: 'Add',
      onClick: function () {
        const email = addAdminInput.value.trim().toLowerCase();
        if (!email || !email.includes('@')) {
          showToast('Please enter a valid email address', 'error');
          return;
        }
        if (typeof KFGConfig !== 'undefined' && KFGConfig.addAdminEmail) {
          KFGConfig.addAdminEmail(email);
          logAuditAction('Added admin user: ' + email);
          showToast('Added admin: ' + email, 'success');
          addAdminInput.value = '';
          renderAdminList();
        }
      },
    });

    adminSection.appendChild(el('div', { className: 'kfg-add-admin-row' },
      addAdminInput,
      addAdminBtn
    ));

    body.appendChild(adminSection);

    // --- Data & Config ---
    const dataSection = el('div', { className: 'kfg-settings-section' },
      el('h3', { textContent: 'Data & Config' })
    );

    const exportBtn = el('button', {
      className: 'kfg-admin-btn',
      'aria-label': 'Export configuration as JSON file',
      textContent: 'Export Config',
      onClick: function () { exportConfig(); },
      style: { color: '#1e293b' },
    });

    const importInput = el('input', {
      type: 'file',
      accept: '.json',
      className: 'kfg-file-hidden',
      'aria-label': 'Import configuration file',
      onChange: function (e) { handleImportConfig(e); },
    });

    const importBtn = el('button', {
      className: 'kfg-admin-btn',
      'aria-label': 'Import configuration from JSON file',
      textContent: 'Import Config',
      onClick: function () { importInput.click(); },
      style: { color: '#1e293b' },
    });

    const resetBtn = el('button', {
      className: 'kfg-admin-btn kfg-admin-btn--danger',
      'aria-label': 'Reset all settings to defaults',
      textContent: 'Reset to Defaults',
      onClick: function () {
        openResetConfirmation();
      },
    });

    dataSection.appendChild(el('div', { style: { display: 'flex', gap: '8px', flexWrap: 'wrap' } },
      exportBtn, importBtn, importInput, resetBtn
    ));

    body.appendChild(dataSection);

    // --- Audit Log ---
    const auditSection = el('div', { className: 'kfg-settings-section' },
      el('h3', { textContent: 'Audit Log' })
    );

    const auditLog = getAuditLog(20);
    const auditContainer = el('div', { className: 'kfg-audit-log' });

    if (auditLog.length === 0) {
      auditContainer.appendChild(
        el('div', { style: { padding: '12px', color: '#94a3b8', textAlign: 'center' }, textContent: 'No admin actions recorded in this session.' })
      );
    } else {
      const table = el('table');
      const thead = el('thead');
      const headerRow = el('tr',{},
        el('th', { textContent: 'Time' }),
        el('th', { textContent: 'User' }),
        el('th', { textContent: 'Action' })
      );
      thead.appendChild(headerRow);
      table.appendChild(thead);

      const tbody = el('tbody');
      auditLog.forEach(function (entry) {
        const time = new Date(entry.timestamp);
        const timeStr = time.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        const userStr = entry.user ? entry.user.split('@')[0] : 'unknown';

        tbody.appendChild(el('tr', {},
          el('td', { textContent: timeStr }),
          el('td', { textContent: userStr }),
          el('td', { textContent: entry.action })
        ));
      });
      table.appendChild(tbody);
      auditContainer.appendChild(table);
    }

    auditSection.appendChild(auditContainer);
    body.appendChild(auditSection);

    // Show panel
    settingsBackdrop.classList.add('visible');
    settingsBackdrop.setAttribute('aria-hidden', 'false');
    settingsPanel.classList.add('visible');

    // Escape key handler
    function settingsKeyHandler(e) {
      if (e.key === 'Escape') {
        closeSettings();
        document.removeEventListener('keydown', settingsKeyHandler);
      }
    }
    document.addEventListener('keydown', settingsKeyHandler);

    // Focus the close button
    const closeBtn = settingsPanel.querySelector('.kfg-modal-close');
    if (closeBtn) closeBtn.focus();
  }

  /**
   * Close the settings panel.
   */
  function closeSettings() {
    if (settingsBackdrop) {
      settingsBackdrop.classList.remove('visible');
      settingsBackdrop.setAttribute('aria-hidden', 'true');
    }
    if (settingsPanel) {
      settingsPanel.classList.remove('visible');
    }
  }

  /**
   * Open a confirmation dialog for resetting to defaults.
   */
  function openResetConfirmation() {
    const message = el('p', {
      textContent: 'Are you sure you want to reset all configuration to defaults? This action cannot be undone.',
      style: { margin: '0', fontSize: '14px', color: '#374151' },
    });

    const confirmBtn = el('button', {
      className: 'kfg-admin-btn kfg-admin-btn--danger',
      'aria-label': 'Confirm reset to defaults',
      textContent: 'Yes, Reset',
      onClick: function () {
        if (typeof KFGConfig !== 'undefined' && KFGConfig.resetAll) {
          KFGConfig.resetAll();
          logAuditAction('Reset all configuration to defaults');
          showToast('All settings reset to defaults', 'success');
          closeModal();
          // Refresh the page to apply defaults
          setTimeout(function () {
            window.location.reload();
          }, 1000);
        } else {
          showToast('Reset function not available', 'error');
          closeModal();
        }
      },
    });

    const cancelBtn = el('button', {
      className: 'kfg-admin-btn',
      'aria-label': 'Cancel reset',
      textContent: 'Cancel',
      onClick: function () { closeModal(); },
    });

    const footer = el('div', { style: { display: 'contents' } }, cancelBtn, confirmBtn);

    openModal('Confirm Reset', message, footer);
  }

  // ---------------------------------------------------------------------------
  // Config Import / Export
  // ---------------------------------------------------------------------------

  /**
   * Export the current configuration as a downloadable JSON file.
   */
  function exportConfig() {
    let config = {};
    if (typeof KFGConfig !== 'undefined' && KFGConfig.exportAll) {
      config = KFGConfig.exportAll();
    } else {
      showToast('Config module not available', 'error');
      return;
    }

    const json = JSON.stringify(config, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'kfg-dashboard-config-' + new Date().toISOString().slice(0, 10) + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    logAuditAction('Exported configuration');
    showToast('Configuration exported', 'success');
  }

  /**
   * Handle importing a configuration JSON file.
   * @param {Event} e - The file input change event.
   */
  function handleImportConfig(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (event) {
      try {
        const config = JSON.parse(event.target.result);

        if (typeof KFGConfig !== 'undefined' && KFGConfig.importAll) {
          KFGConfig.importAll(config);
          logAuditAction('Imported configuration from file: ' + file.name);
          showToast('Configuration imported successfully', 'success');

          // Reload to apply
          setTimeout(function () {
            window.location.reload();
          }, 1000);
        } else {
          showToast('Config module not available', 'error');
        }
      } catch (err) {
        console.error('[KFGAdmin] Failed to parse config file:', err);
        showToast('Invalid configuration file', 'error');
      }
    };

    reader.onerror = function () {
      showToast('Failed to read file', 'error');
    };

    reader.readAsText(file);

    // Reset file input so the same file can be selected again
    e.target.value = '';
  }

  // ---------------------------------------------------------------------------
  // Initialization
  // ---------------------------------------------------------------------------

  /**
   * Initialize the admin panel. Checks admin status and creates UI elements.
   */
  function init() {
    if (initialized) return;
    initialized = true;

    injectStyles();
    ensureToastContainer();

    // Check admin status
    if (!isAdmin()) {
      console.log('[KFGAdmin] Current user is not an admin. Admin panel disabled.');
      return;
    }

    console.log('[KFGAdmin] Admin user detected. Initializing admin panel.');

    createToolbar();
    showToolbar();

    logAuditAction('Admin panel initialized');
  }

  /**
   * Toggle the admin panel (toolbar) visibility.
   */
  function toggle() {
    if (!isAdmin()) {
      showToast('Admin access required', 'error');
      return;
    }
    toggleToolbar();
  }

  // ---------------------------------------------------------------------------
  // Expose public API on window.KFGAdmin
  // ---------------------------------------------------------------------------

  window.KFGAdmin = {
    init: init,
    toggle: toggle,
    openChartEditor: openChartEditor,
    openTableEditor: openTableEditor,
    openSettings: openSettings,
    showToast: showToast,
  };
})();

# KFG Executive Dashboard

> Senior Management Operations Hub — Kids First Group L.L.C

A fully client-side executive dashboard that reads live data from SharePoint Excel files
via Microsoft Graph API, secured with Azure AD (Microsoft Entra ID) authentication.

---

## Quick Start (Demo Mode — no setup needed)

1. Open `index.html` in any modern browser **or** serve it from GitHub Pages.
2. Click **"View with Sample Data (Demo Mode)"**.
3. Explore all 9 department pages with realistic demo data.

---

## Full Setup Guide — Go Live in 5 Steps

### STEP 1 — Register an Azure AD Application (10 minutes)

1. Go to **[https://portal.azure.com](https://portal.azure.com)** and sign in with your
   `@kidsfirstgroup.com` Microsoft 365 account.

2. In the search bar type **"App registrations"** → click it → click **"+ New registration"**.

3. Fill in the form:

   | Field | Value |
   |-------|-------|
   | Name | `KFG Executive Dashboard` |
   | Supported account types | **Accounts in this organizational directory only** |
   | Redirect URI — platform | **Single-page application (SPA)** |
   | Redirect URI — value | `https://<your-github-username>.github.io/KfgDashboard/` |

   > Replace `<your-github-username>` with your actual GitHub username.

4. Click **"Register"**.

5. On the overview page, copy the **Application (client) ID**.
   It looks like: `12345678-abcd-1234-efgh-123456789abc`
   **Save this — you will paste it into the dashboard login screen.**

6. In the left sidebar click **"API permissions"** → **"+ Add a permission"**
   → **"Microsoft Graph"** → **"Delegated permissions"** → search and add:
   - `User.Read`
   - `Sites.Read.All`
   - `Files.Read.All`

7. Click **"Grant admin consent for Kids First Group"** → confirm **Yes**.

8. In the left sidebar click **"Authentication"** and confirm:
   - The SPA redirect URI is listed correctly.
   - Under **"Implicit grant and hybrid flows"** both checkboxes are **unchecked**
     (the dashboard uses the secure PKCE flow, not implicit grant).

---

### STEP 2 — Create a GitHub Repository and Push the Files

**First, create the repository on GitHub:**

1. Go to [https://github.com/new](https://github.com/new)
2. Repository name: `KfgDashboard`
3. Visibility: **Private** (recommended — protected by Azure AD anyway)
4. Do **NOT** tick "Add a README" (you already have one)
5. Click **"Create repository"**

**Then push your files from the project folder:**

Open a terminal (Git Bash or PowerShell) and run these commands one by one:

```bash
# 1. Go to your project folder
cd "C:\Users\Shady\OneDrive - Kids First Group L.L.C\Documents\KfgDashboard"

# 2. Initialise git
git init
git branch -M main

# 3. Create a .gitignore file
echo *.env > .gitignore
echo .DS_Store >> .gitignore
echo Thumbs.db >> .gitignore

# 4. Stage all project files
git add index.html css/ js/ config/ README.md .gitignore

# 5. Create your first commit
git commit -m "Initial commit: KFG Executive Dashboard"

# 6. Connect to your GitHub repository
#    Replace YOUR_USERNAME with your actual GitHub username
git remote add origin https://github.com/YOUR_USERNAME/KfgDashboard.git

# 7. Push to GitHub
git push -u origin main
```

> If prompted for credentials, sign in with your GitHub account or use a
> [Personal Access Token](https://github.com/settings/tokens).

---

### STEP 3 — Enable GitHub Pages

1. In your GitHub repository page, click **"Settings"** (top navigation tab).
2. In the left sidebar, scroll down to click **"Pages"**.
3. Under **"Build and deployment" → "Source"** choose:
   - Branch: `main`
   - Folder: `/ (root)`
4. Click **"Save"**.
5. Wait ~2 minutes. Your live URL will appear at the top of the Pages settings:
   ```
   https://YOUR_USERNAME.github.io/KfgDashboard/
   ```
6. Click the link to verify the dashboard loads.

---

### STEP 4 — Add Your Live URL to Azure AD

Now that your GitHub Pages URL is confirmed:

1. Go back to **[portal.azure.com](https://portal.azure.com)**
2. App registrations → **KFG Executive Dashboard** → **Authentication**
3. Under "Single-page application" → click **"Add URI"**
4. Paste your GitHub Pages URL (with trailing slash):
   ```
   https://YOUR_USERNAME.github.io/KfgDashboard/
   ```
5. Also add a localhost URI for future local development:
   ```
   http://localhost:5500/
   ```
6. Click **"Save"**

---

### STEP 5 — Open the Dashboard and Connect

1. Visit your live GitHub Pages URL.
2. You will see the login screen with two fields:

   **SharePoint Site URL** — paste your SharePoint site, for example:
   ```
   https://kfguae.sharepoint.com/sites/ExecutiveDashboard
   ```

   **Azure App Client ID** — paste the Client ID you saved in Step 1:
   ```
   12345678-abcd-1234-efgh-123456789abc
   ```

3. Click **"Connect & Sign In with Microsoft"**.
4. A Microsoft login popup appears — sign in with `shady.sabbagh@kidsfirstgroup.com`.
5. Approve the permissions when prompted (only happens once).
6. The dashboard loads your live SharePoint Excel data automatically.

---

## Updating the Dashboard After Changes

Whenever you make changes to the project files, push them to GitHub:

```bash
cd "C:\Users\Shady\OneDrive - Kids First Group L.L.C\Documents\KfgDashboard"

git add -A
git commit -m "Update: describe your change here"
git push
```

GitHub Pages automatically re-deploys within ~1 minute.

---

## Admin Panel

Once signed in as `shady.sabbagh@kidsfirstgroup.com`, an **Admin** button appears in the top bar.

| Action | How |
|--------|-----|
| Edit a chart (title, type, colours, legend) | Click the **✏️ pencil** icon on any chart card |
| Edit a table (columns, sorting, page size) | Click the **✏️ pencil** icon on any table card |
| Change auto-refresh interval | Admin → Settings → Refresh Interval |
| Manually refresh all data | Click **↻ Refresh** in the top bar |
| Export all settings to a file | Admin → Settings → Export Config |
| Import settings from a file | Admin → Settings → Import Config |
| Reset everything to defaults | Admin → Settings → Reset to Defaults |
| View audit log | Admin → Settings → Audit Log (last 50 actions) |

Admin settings are saved in your browser's localStorage and persist across sessions.

---

## Connecting to Your SharePoint Excel Files

The dashboard is pre-configured with a demo SharePoint URL. To connect your real files:

1. In `js/graph.js` find the `FILE_MAP` constant and update the file paths to match your
   SharePoint document library:

   ```javascript
   const FILE_MAP = {
     maintenance:  { path: '/Shared Documents/Dashboard/Maintenance.xlsx',  kpi: 'KPIs', data: 'Jobs' },
     capex:        { path: '/Shared Documents/Dashboard/Capex.xlsx',        kpi: 'KPIs', data: 'Nurseries' },
     projects:     { path: '/Shared Documents/Dashboard/Projects.xlsx',     kpi: 'KPIs', data: 'Projects' },
     procurement:  { path: '/Shared Documents/Dashboard/Procurement.xlsx',  kpi: 'KPIs', data: 'POs' },
     it:           { path: '/Shared Documents/Dashboard/IT.xlsx',           kpi: 'KPIs', data: 'Tickets' },
     ma:           { path: '/Shared Documents/Dashboard/MA.xlsx',           kpi: 'KPIs', data: 'Deals' },
     greenfield:   { path: '/Shared Documents/Dashboard/Greenfield.xlsx',   kpi: 'KPIs', data: 'Sites' },
     other:        { path: '/Shared Documents/Dashboard/OtherProjects.xlsx',kpi: 'KPIs', data: 'Projects' },
   };
   ```

2. Each Excel file must have:
   - A sheet named `KPIs` with columns: `Metric | Value`
   - A data sheet (e.g. `Jobs`) with column headers in row 1 and data from row 2.

3. Save the file, commit and push:
   ```bash
   git add js/graph.js
   git commit -m "Connect live SharePoint Excel files"
   git push
   ```

---

## File Structure

```
KfgDashboard/
├── index.html          # All 9 page sections, login screen, navigation
├── css/
│   └── styles.css      # Master stylesheet (WCAG 2.1 AA, responsive, print)
├── js/
│   ├── config.js       # Admin config manager (localStorage)
│   ├── auth.js         # MSAL.js PKCE auth — Azure AD / Microsoft Entra ID
│   ├── graph.js        # Microsoft Graph API — SharePoint Excel data
│   ├── dashboard.js    # Chart + table rendering for all 8 departments
│   ├── admin.js        # Admin panel — chart/table editors, settings
│   └── app.js          # Orchestrator — auth flow, navigation, refresh
├── config/
│   └── dashboard.json  # Shared default chart/table configuration
└── README.md           # This guide
```

---

## Authentication Flow (Summary)

```
Browser opens dashboard URL
        │
        ▼
  MSAL.js checks sessionStorage for cached token
        │
   No token cached
        │
        ▼
  Login screen shown
  User enters SharePoint URL + Client ID
  Clicks "Connect & Sign In"
        │
        ▼
  MSAL loginPopup() → Microsoft login window opens
  User signs in with @kidsfirstgroup.com account
        │
        ▼
  Domain check: must be @kidsfirstgroup.com
  Admin check:  is email in admin list?
        │
        ▼
  Graph API: fetch SharePoint site + drive IDs
        │
        ▼
  Graph API: fetch all 8 Excel files in parallel
        │
        ▼
  Dashboard renders with live data
        │
        ▼
  Auto-refresh every 5 minutes (configurable)
  Token silently renewed before expiry
```

---

## Security

- No secrets or credentials are ever stored in the code or repository.
- The Azure Client ID is not a secret — it is public by design for SPA applications.
- Authentication uses OAuth 2.0 PKCE (no client secret needed).
- Only `@kidsfirstgroup.com` accounts can access the dashboard.
- Tokens are stored in `sessionStorage` only (cleared when browser tab closes).
- GitHub Pages forces HTTPS — all connections are encrypted.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| "AADSTS50011: The reply URL does not match" | Add your exact GitHub Pages URL to Azure AD → Authentication → Redirect URIs |
| Microsoft login popup is blocked | Allow popups for `github.io` in your browser settings |
| "Need admin approval" message | Ask your Microsoft 365 admin to grant consent in Azure AD → API Permissions |
| Charts appear blank | Open DevTools (F12) → Console and check for JavaScript errors |
| "Demo Data" shown even after signing in | SharePoint URL or Client ID may be incorrect — check and retry |
| Data not refreshing | Check that your account has read access to the SharePoint Excel files |

---

## Admin Contact

**Shady Sabbagh** — shady.sabbagh@kidsfirstgroup.com

For Azure AD / Microsoft 365 issues, visit: [https://portal.azure.com](https://portal.azure.com)

# Zoho Outlook Add-in — Setup Guide

## 1. Register a Zoho OAuth App

1. Go to https://api-console.zoho.com/ and sign in.
2. Click **Add Client** → choose **JavaScript Client**.
3. Fill in:
   - **Client Name**: Outlook Case Creator (or any name)
   - **Homepage URL**: `https://localhost:3000`
   - **Authorized Redirect URIs**: `https://localhost:3000/oauth-callback.html`
4. Click **Create**.
5. Copy the **Client ID** shown on the next screen.

## 2. Configure the Add-in

Open `src/config.js` and replace `YOUR_ZOHO_CLIENT_ID` with your Client ID:

```js
clientId: "1000.XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
```

If your Zoho org is on a non-US data center, also update the URLs:

| Region | zohoBaseUrl | zohoAccountsUrl |
|--------|-------------|-----------------|
| EU | `https://www.zohoapis.eu` | `https://accounts.zoho.eu` |
| India | `https://www.zohoapis.in` | `https://accounts.zoho.in` |
| Australia | `https://www.zohoapis.com.au` | `https://accounts.zoho.com.au` |
| Japan | `https://www.zohoapis.jp` | `https://accounts.zoho.jp` |

## 3. Start the Dev Server

```powershell
npm install
npm start
```

This installs a trusted localhost HTTPS certificate and serves the add-in on `https://localhost:3000`.

## 4. Sideload the Manifest in Outlook

### New Outlook (Windows) / Outlook on the Web
1. Open Outlook → click the **Apps** icon (puzzle piece) in the toolbar.
2. Click **Add apps** → **Upload a custom app** (bottom of the panel).
3. Select `manifest.xml` from this folder.
4. Open any email → a **Zoho CRM** tab appears in the ribbon → click **Create Case**.

### Classic Outlook (Windows) — alternate method
1. Go to **File → Manage Add-ins** (opens OWA).
2. Click **+** → **Add from file** → select `manifest.xml`.

## 5. Usage

1. Open an email in Outlook.
2. Click **Create Case** in the Zoho CRM ribbon group.
3. Sign in with Zoho when prompted (first time only — token is cached).
4. Review the pre-filled subject, contact, and description.
5. Adjust Status / Priority if needed.
6. Click **Create Case**.

## Files

```
├── manifest.xml          # Office Add-in manifest
├── taskpane.html         # Task pane UI
├── oauth-callback.html   # OAuth redirect target (popup)
├── function-file.html    # Required by manifest (no-op)
├── styles.css            # Task pane styles
├── src/
│   ├── config.js         # ← Put your Client ID here
│   ├── auth.js           # Zoho OAuth 2.0 + PKCE flow
│   ├── zoho.js           # Zoho CRM API calls
│   └── taskpane.js       # Office.js task pane logic
└── package.json          # Dev server
```

## How Auth Works (Client-Side PKCE)

- On first sign-in a popup opens the Zoho consent screen.
- After the user approves, Zoho redirects to `oauth-callback.html` with an auth code.
- The callback page posts the code back to the task pane via `window.postMessage`.
- The task pane exchanges the code (+ PKCE verifier) for tokens **directly against Zoho's token endpoint** — no backend needed.
- Tokens are stored in `localStorage`. The access token auto-refreshes using the stored refresh token.

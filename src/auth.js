// Zoho OAuth 2.0 with PKCE — client-side only, no backend required

const Auth = (() => {
  // --- PKCE helpers ---
  function randomBytes(length) {
    const arr = new Uint8Array(length);
    crypto.getRandomValues(arr);
    return arr;
  }

  function base64UrlEncode(buffer) {
    return btoa(String.fromCharCode(...new Uint8Array(buffer)))
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=/g, "");
  }

  async function generateCodeVerifier() {
    return base64UrlEncode(randomBytes(32));
  }

  async function generateCodeChallenge(verifier) {
    const encoder = new TextEncoder();
    const data = encoder.encode(verifier);
    const digest = await crypto.subtle.digest("SHA-256", data);
    return base64UrlEncode(digest);
  }

  // --- Token storage ---
  function saveTokens({ access_token, refresh_token, expires_in }) {
    const expiresAt = Date.now() + expires_in * 1000 - 60_000; // 1 min buffer
    localStorage.setItem(CONFIG.storageKeys.accessToken, access_token);
    if (refresh_token) localStorage.setItem(CONFIG.storageKeys.refreshToken, refresh_token);
    localStorage.setItem(CONFIG.storageKeys.expiresAt, String(expiresAt));
  }

  function getAccessToken() {
    const token = localStorage.getItem(CONFIG.storageKeys.accessToken);
    const expiresAt = Number(localStorage.getItem(CONFIG.storageKeys.expiresAt));
    if (!token || Date.now() >= expiresAt) return null;
    return token;
  }

  function clearTokens() {
    Object.values(CONFIG.storageKeys).forEach((k) => localStorage.removeItem(k));
  }

  // --- Auth flow ---
  async function login() {
    const verifier = await generateCodeVerifier();
    const challenge = await generateCodeChallenge(verifier);
    localStorage.setItem(CONFIG.storageKeys.codeVerifier, verifier);

    const params = new URLSearchParams({
      response_type: "code",
      client_id: CONFIG.clientId,
      redirect_uri: CONFIG.redirectUri,
      scope: CONFIG.scope,
      code_challenge: challenge,
      code_challenge_method: "S256",
    });

    const authUrl = `${CONFIG.zohoAccountsUrl}/oauth/v2/auth?${params}`;

    // Open a popup — Zoho redirects back to oauth-callback.html
    return new Promise((resolve, reject) => {
      const popup = window.open(authUrl, "zohoAuth", "width=600,height=700,left=200,top=100");
      if (!popup) {
        reject(new Error("Popup blocked. Please allow popups for this add-in."));
        return;
      }

      const handler = async (event) => {
        if (event.origin !== window.location.origin) return;
        if (event.data?.type !== "zoho_oauth_code") return;
        window.removeEventListener("message", handler);

        try {
          const tokens = await exchangeCode(event.data.code);
          saveTokens(tokens);
          resolve(tokens.access_token);
        } catch (err) {
          reject(err);
        }
      };

      window.addEventListener("message", handler);

      // Detect if popup was closed without completing auth
      const poll = setInterval(() => {
        if (popup.closed) {
          clearInterval(poll);
          window.removeEventListener("message", handler);
          reject(new Error("Auth popup closed before completing login."));
        }
      }, 500);
    });
  }

  // Zoho's token endpoint requires a POST — we call it directly from the browser.
  // Note: Zoho allows this for JavaScript clients registered with PKCE.
  async function exchangeCode(code) {
    const verifier = localStorage.getItem(CONFIG.storageKeys.codeVerifier);
    const body = new URLSearchParams({
      grant_type: "authorization_code",
      client_id: CONFIG.clientId,
      redirect_uri: CONFIG.redirectUri,
      code,
      code_verifier: verifier,
    });

    const res = await fetch(`${CONFIG.zohoAccountsUrl}/oauth/v2/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });

    if (!res.ok) throw new Error(`Token exchange failed: ${res.status} ${await res.text()}`);
    return res.json();
  }

  async function ensureToken() {
    let token = getAccessToken();
    if (token) return token;

    // Try refresh
    const refreshToken = localStorage.getItem(CONFIG.storageKeys.refreshToken);
    if (refreshToken) {
      try {
        token = await refreshAccessToken(refreshToken);
        return token;
      } catch {
        clearTokens();
      }
    }

    // Full login required
    return login();
  }

  async function refreshAccessToken(refreshToken) {
    const body = new URLSearchParams({
      grant_type: "refresh_token",
      client_id: CONFIG.clientId,
      refresh_token: refreshToken,
    });

    const res = await fetch(`${CONFIG.zohoAccountsUrl}/oauth/v2/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });

    if (!res.ok) throw new Error(`Refresh failed: ${res.status}`);
    const data = await res.json();
    saveTokens({ ...data, refresh_token: refreshToken });
    return data.access_token;
  }

  function logout() {
    clearTokens();
  }

  function isLoggedIn() {
    return !!getAccessToken();
  }

  return { ensureToken, logout, isLoggedIn };
})();

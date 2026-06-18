// Zoho OAuth 2.0 implicit flow for JavaScript clients

const Auth = (() => {
  const SCOPE_KEY = "zoho_scope_hash";

  function scopeHash() {
    // Simple fingerprint so a scope change forces re-auth
    let h = 0;
    for (let i = 0; i < CONFIG.scope.length; i++) {
      h = (Math.imul(31, h) + CONFIG.scope.charCodeAt(i)) | 0;
    }
    return String(h);
  }

  function saveToken({ access_token, expires_in }) {
    const expiresAt = Date.now() + Number(expires_in) * 1000 - 60_000;
    localStorage.setItem(CONFIG.storageKeys.accessToken, access_token);
    localStorage.setItem(CONFIG.storageKeys.expiresAt, String(expiresAt));
    localStorage.setItem(SCOPE_KEY, scopeHash());
  }

  function getAccessToken() {
    // Invalidate token if scope has changed since it was issued
    if (localStorage.getItem(SCOPE_KEY) !== scopeHash()) {
      clearTokens();
      return null;
    }
    const token = localStorage.getItem(CONFIG.storageKeys.accessToken);
    const expiresAt = Number(localStorage.getItem(CONFIG.storageKeys.expiresAt));
    if (!token || Date.now() >= expiresAt) return null;
    return token;
  }

  function clearTokens() {
    Object.values(CONFIG.storageKeys).forEach((k) => localStorage.removeItem(k));
    localStorage.removeItem(SCOPE_KEY);
  }

  function login() {
    const params = new URLSearchParams({
      response_type: "token",
      client_id: CONFIG.clientId,
      redirect_uri: CONFIG.redirectUri,
      scope: CONFIG.scope,
      prompt: "consent",
    });

    const authUrl = `${CONFIG.zohoAccountsUrl}/oauth/v2/auth?${params}`;

    return new Promise((resolve, reject) => {
      const popup = window.open(authUrl, "zohoAuth", "width=600,height=700,left=200,top=100");
      if (!popup) {
        reject(new Error("Popup blocked. Please allow popups for this add-in."));
        return;
      }

      const handler = (event) => {
        if (event.origin !== window.location.origin) return;
        if (event.data?.type !== "zoho_oauth_token") return;
        window.removeEventListener("message", handler);
        clearInterval(poll);

        if (event.data.error) {
          reject(new Error(event.data.error));
          return;
        }

        saveToken(event.data);
        resolve(event.data.access_token);
      };

      window.addEventListener("message", handler);

      const poll = setInterval(() => {
        if (popup.closed) {
          clearInterval(poll);
          window.removeEventListener("message", handler);
          reject(new Error("Auth popup closed before completing login."));
        }
      }, 500);
    });
  }

  async function ensureToken() {
    const token = getAccessToken();
    if (token) return token;
    return login();
  }

  function logout() {
    clearTokens();
  }

  function isLoggedIn() {
    return !!getAccessToken();
  }

  return { ensureToken, logout, isLoggedIn };
})();
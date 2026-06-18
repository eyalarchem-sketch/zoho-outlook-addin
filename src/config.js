// Zoho OAuth & API configuration
// Register your app at https://api-console.zoho.com/ as a "JavaScript Client"
const CONFIG = {
  // From your Zoho OAuth app registration
  clientId: "1000.Y7VSJO7W0VVHHWTR1GJ5OWXRB48LMM",

  // Must match exactly what you registered in the Zoho Developer Console
  redirectUri: "https://eyalarchem-sketch.github.io/zoho-outlook-addin/oauth-callback.html",

  // Zoho data center - change if your org is on .eu / .in / .com.au / .jp
  zohoBaseUrl: "https://www.zohoapis.com",
  zohoAccountsUrl: "https://accounts.zoho.com",

  // OAuth scopes needed
  scope: "ZohoCRM.modules.Cases.CREATE,ZohoCRM.modules.Contacts.READ,ZohoCRM.modules.Attachments.CREATE",

  // localStorage keys
  storageKeys: {
    accessToken: "zoho_access_token",
    expiresAt: "zoho_expires_at",
  },
};
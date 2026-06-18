// Zoho CRM API wrapper

const Zoho = (() => {
  async function apiFetch(path, options = {}) {
    const token = await Auth.ensureToken();
    const url = `${CONFIG.zohoBaseUrl}/crm/v6/${path}`;

    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Zoho-oauthtoken ${token}`,
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Zoho API ${res.status}: ${text}`);
    }

    if (res.status === 204) return null;
    return res.json();
  }

  // Get org ID for building case URLs (cached in localStorage)
  async function getOrgId() {
    const cached = localStorage.getItem("zoho_org_id");
    if (cached) return cached;
    try {
      const data = await apiFetch("org");
      const id = String(data?.org?.[0]?.zgid || "");
      if (id) localStorage.setItem("zoho_org_id", id);
      return id;
    } catch {
      return "";
    }
  }

  function caseUrl(caseId, orgId) {
    const base = orgId
      ? `https://crm.zoho.com/crm/org${orgId}/tab/Cases/${caseId}`
      : `https://crm.zoho.com/crm/tab/Cases/${caseId}`;
    return base;
  }

  // Look up a Contact by exact email. Returns first match or null.
  async function findContactByEmail(email) {
    if (!email) return null;
    const criteria = `((Email:equals:${email}))`;
    const data = await apiFetch(
      `Contacts/search?criteria=${encodeURIComponent(criteria)}&fields=id,Full_Name,Email,Account_Name`
    );
    return data?.data?.[0] ?? null;
  }

  // Search contacts by name (starts_with) or email (equals). Returns up to 5 matches.
  async function searchContacts(query) {
    if (!query || query.length < 2) return [];
    try {
      const isEmail = query.includes("@");
      const criteria = isEmail
        ? `((Email:equals:${query}))`
        : `((Full_Name:starts_with:${query}))`;
      const data = await apiFetch(
        `Contacts/search?criteria=${encodeURIComponent(criteria)}&fields=id,Full_Name,Email,Account_Name&per_page=5`
      );
      return data?.data ?? [];
    } catch {
      return [];
    }
  }

  // Create a Case. Returns { id, url }.
  async function createCase({ subject, description, contactId, accountId, status, priority }) {
    const record = {
      Subject: subject,
      Description: description,
      Status: status || "New",
      Priority: priority || "Normal",
      Case_Origin: "Email",
    };

    if (contactId) record.Related_To  = { id: contactId };
    if (accountId) record.Account_Name = { id: accountId };

    const body = { data: [record] };
    const result = await apiFetch("Cases", {
      method: "POST",
      body: JSON.stringify(body),
    });

    const item = result?.data?.[0];
    if (item?.status !== "success") {
      throw new Error(item?.message || "Failed to create case");
    }

    const caseId = item.details.id;
    const orgId  = await getOrgId();
    return { id: caseId, url: caseUrl(caseId, orgId) };
  }

  // Attach a file (as a Blob) to a Case record.
  async function attachFileToCaseRaw(caseId, fileName, blob) {
    const token = await Auth.ensureToken();
    const url = `${CONFIG.zohoBaseUrl}/crm/v6/Cases/${caseId}/Attachments`;

    const form = new FormData();
    form.append("file", blob, fileName);

    const res = await fetch(url, {
      method: "POST",
      headers: { Authorization: `Zoho-oauthtoken ${token}` },
      body: form,
    });

    if (!res.ok) {
      const text = await res.text();
      console.warn(`Attachment upload failed for "${fileName}": ${res.status} ${text}`);
    }
  }

  return { findContactByEmail, searchContacts, createCase, attachFileToCaseRaw };
})();
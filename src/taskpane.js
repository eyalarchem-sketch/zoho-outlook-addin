// Task pane controller

Office.onReady(async () => {
  const btnLogin    = document.getElementById("btnLogin");
  const btnLogout   = document.getElementById("btnLogout");
  const btnSubmit   = document.getElementById("btnSubmit");
  const authSection = document.getElementById("authSection");
  const formSection = document.getElementById("formSection");
  const statusEl    = document.getElementById("status");

  // Related To (Contact)
  const contactSearch       = document.getElementById("contactSearch");
  const contactDropdown     = document.getElementById("contactDropdown");
  const contactSelected     = document.getElementById("contactSelected");
  const contactSelectedName = document.getElementById("contactSelectedName");
  const btnClearContact     = document.getElementById("btnClearContact");
  const contactIdInput      = document.getElementById("contactId");
  const accountIdInput      = document.getElementById("accountId");
  const accountNameRow      = document.getElementById("accountNameRow");
  const accountNameDisplay  = document.getElementById("accountNameDisplay");

  // Supplier
  const supplierSearch       = document.getElementById("supplierSearch");
  const supplierDropdown     = document.getElementById("supplierDropdown");
  const supplierSelected     = document.getElementById("supplierSelected");
  const supplierSelectedName = document.getElementById("supplierSelectedName");
  const btnClearSupplier     = document.getElementById("btnClearSupplier");
  const supplierIdInput      = document.getElementById("supplierId");
  const supplierContactRow   = document.getElementById("supplierContactRow");
  const supplierContactSelect = document.getElementById("supplierContactSelect");
  const supplierContactIdInput = document.getElementById("supplierContactId");

  function setStatus(msg, isError = false, link = null) {
    statusEl.innerHTML = "";
    if (link) {
      statusEl.innerHTML = `${msg} <a href="${link}" target="_blank" rel="noopener">Open in Zoho</a>`;
    } else {
      statusEl.textContent = msg;
    }
    statusEl.className = isError ? "status error" : "status info";
    statusEl.hidden = !msg;
  }

  function showLoading(show) {
    btnSubmit.disabled = show;
    btnSubmit.textContent = show ? "Creating…" : "Create Case";
  }

  // --- Contact helpers ---

  function selectContact(id, name, email, accountId, accountName) {
    contactIdInput.value  = id || "";
    accountIdInput.value  = accountId || "";
    contactSelectedName.textContent = id ? `${name} — ${email}` : "(none)";
    contactSelected.hidden = false;
    contactSearch.hidden   = true;
    contactDropdown.hidden = true;
    accountNameDisplay.value = accountName || "";
    accountNameRow.hidden    = !accountName;
  }

  function clearContact() {
    contactIdInput.value  = "";
    accountIdInput.value  = "";
    accountNameDisplay.value = "";
    accountNameRow.hidden    = true;
    contactSelected.hidden = true;
    contactSearch.hidden   = false;
    contactSearch.value    = "";
    contactDropdown.hidden = true;
    contactSearch.focus();
  }

  function showContactDropdown(contacts) {
    contactDropdown.innerHTML = "";
    if (!contacts.length) {
      const item = document.createElement("div");
      item.className = "dropdown-item dropdown-empty";
      item.textContent = "No contacts found";
      contactDropdown.appendChild(item);
    } else {
      contacts.forEach((c) => {
        const item = document.createElement("div");
        item.className = "dropdown-item";
        item.innerHTML = `<strong>${c.Full_Name || ""}</strong><span>${c.Email || ""}${c.Account_Name ? " · " + c.Account_Name.name : ""}</span>`;
        item.addEventListener("mousedown", (e) => {
          e.preventDefault();
          selectContact(c.id, c.Full_Name || "", c.Email || "", c.Account_Name?.id || "", c.Account_Name?.name || "");
        });
        contactDropdown.appendChild(item);
      });
    }
    contactDropdown.hidden = false;
  }

  let contactSearchTimer = null;
  contactSearch.addEventListener("input", () => {
    clearTimeout(contactSearchTimer);
    const q = contactSearch.value.trim();
    if (!q || q.length < 2) { contactDropdown.hidden = true; return; }
    contactSearchTimer = setTimeout(async () => {
      const results = await Zoho.searchContacts(q);
      showContactDropdown(results);
    }, 350);
  });

  contactSearch.addEventListener("blur",  () => { setTimeout(() => { contactDropdown.hidden = true; }, 150); });
  contactSearch.addEventListener("focus", () => { if (contactSearch.value.trim().length >= 2) contactDropdown.hidden = false; });
  btnClearContact.addEventListener("click", clearContact);

  // --- Supplier helpers ---

  async function selectSupplier(id, name) {
    supplierIdInput.value = id || "";
    supplierSelectedName.textContent = name || "";
    supplierSelected.hidden = false;
    supplierSearch.hidden   = true;
    supplierDropdown.hidden = true;

    // Load contacts for this supplier account
    supplierContactSelect.innerHTML = '<option value="">— Loading… —</option>';
    supplierContactRow.hidden = false;
    supplierContactIdInput.value = "";

    const contacts = await Zoho.getContactsByAccount(id);
    supplierContactSelect.innerHTML = '<option value="">— None —</option>';
    contacts.forEach((c) => {
      const opt = document.createElement("option");
      opt.value = c.id;
      opt.textContent = c.Title ? `${c.Full_Name || c.Email} — ${c.Title}` : (c.Full_Name || c.Email || c.id);
      supplierContactSelect.appendChild(opt);
    });
  }

  function clearSupplier() {
    supplierIdInput.value = "";
    supplierContactIdInput.value = "";
    supplierSelected.hidden = true;
    supplierSearch.hidden   = false;
    supplierSearch.value    = "";
    supplierDropdown.hidden = true;
    supplierContactRow.hidden = true;
    supplierContactSelect.innerHTML = '<option value="">— None —</option>';
    supplierSearch.focus();
  }

  function showSupplierDropdown(accounts) {
    supplierDropdown.innerHTML = "";
    if (!accounts.length) {
      const item = document.createElement("div");
      item.className = "dropdown-item dropdown-empty";
      item.textContent = "No accounts found";
      supplierDropdown.appendChild(item);
    } else {
      accounts.forEach((a) => {
        const item = document.createElement("div");
        item.className = "dropdown-item";
        item.innerHTML = `<strong>${a.Account_Name || ""}</strong>`;
        item.addEventListener("mousedown", (e) => {
          e.preventDefault();
          selectSupplier(a.id, a.Account_Name || "");
        });
        supplierDropdown.appendChild(item);
      });
    }
    supplierDropdown.hidden = false;
  }

  let supplierSearchTimer = null;
  supplierSearch.addEventListener("input", () => {
    clearTimeout(supplierSearchTimer);
    const q = supplierSearch.value.trim();
    if (!q || q.length < 2) { supplierDropdown.hidden = true; return; }
    supplierSearchTimer = setTimeout(async () => {
      try {
        const results = await Zoho.searchAccounts(q);
        showSupplierDropdown(results);
      } catch (err) {
        setStatus(`Supplier search: ${err.message}`, true);
      }
    }, 350);
  });

  supplierSearch.addEventListener("blur",  () => { setTimeout(() => { supplierDropdown.hidden = true; }, 150); });
  supplierSearch.addEventListener("focus", () => { if (supplierSearch.value.trim().length >= 2) supplierDropdown.hidden = false; });
  btnClearSupplier.addEventListener("click", clearSupplier);

  supplierContactSelect.addEventListener("change", () => {
    supplierContactIdInput.value = supplierContactSelect.value;
  });

  // --- Form population ---

  async function populateForm() {
    const item = Office.context.mailbox.item;
    document.getElementById("subject").value = item.subject || "";

    const senderEmail = item.from?.emailAddress || "";
    if (senderEmail) {
      contactSearch.value = senderEmail;
      setStatus("Looking up contact in Zoho…");
      try {
        const contact = await Zoho.findContactByEmail(senderEmail);
        if (contact) {
          selectContact(contact.id, contact.Full_Name || senderEmail, contact.Email || senderEmail, contact.Account_Name?.id || "", contact.Account_Name?.name || "");
        } else {
          contactSearch.value = senderEmail;
        }
        setStatus("");
      } catch (err) {
        setStatus(`Contact lookup: ${err.message}`, true);
        contactSearch.value = senderEmail;
      }
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("description").value = result.value?.trim() || "";
      }
    });
  }

  // --- Submit ---

  async function handleSubmit() {
    const subject = document.getElementById("subject").value.trim();
    if (!subject) { setStatus("Subject is required.", true); return; }

    const type = document.getElementById("caseType").value;
    if (!type || type === "-None-") { setStatus("Case Type is required.", true); return; }

    const description         = document.getElementById("description").value.trim();
    const contactId           = contactIdInput.value.trim() || null;
    const accountId           = accountIdInput.value.trim() || null;
    const supplierId          = supplierIdInput.value.trim() || null;
    const supplierContactId   = supplierContactIdInput.value.trim() || null;
    const status              = document.getElementById("caseStatus").value;
    const priority            = document.getElementById("casePriority").value;

    showLoading(true);
    setStatus("");

    try {
      const { id: caseId, url: caseLink } = await Zoho.createCase({
        subject, description, contactId, accountId,
        supplierId, supplierContactId, type, status, priority,
      });
      setStatus(`Case ${caseId} created!`, false, caseLink);

      if (document.getElementById("includeAttachments").checked) {
        await uploadAttachments(caseId);
      }
    } catch (err) {
      setStatus(`Error: ${err.message}`, true);
    } finally {
      showLoading(false);
    }
  }

  async function uploadAttachments(caseId) {
    const item = Office.context.mailbox.item;
    const attachments = (item.attachments || []).filter(a => !a.isInline);
    if (!attachments.length) return;

    return new Promise((resolve) => {
      let remaining = attachments.length;
      let failed = 0;

      attachments.forEach((att) => {
        item.getAttachmentContentAsync(att.id, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const { content, format } = result.value;
            try {
              let blob;
              if (format === Office.AttachmentContentFormat.Base64) {
                const binary = atob(content);
                const bytes = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                blob = new Blob([bytes]);
              } else {
                blob = new Blob([content]);
              }
              await Zoho.attachFileToCaseRaw(caseId, att.name, blob);
            } catch { failed++; }
          } else { failed++; }

          remaining--;
          if (remaining === 0) {
            if (failed > 0) setStatus(`Case created — ${failed} attachment(s) failed.`, true);
            resolve();
          }
        });
      });
    });
  }

  function renderAuth() {
    const loggedIn = Auth.isLoggedIn();
    authSection.hidden = loggedIn;
    formSection.hidden = !loggedIn;
    btnLogout.hidden   = !loggedIn;
    if (loggedIn) populateForm();
  }

  btnLogin.addEventListener("click", async () => {
    setStatus("Opening Zoho login…");
    try { await Auth.ensureToken(); setStatus(""); renderAuth(); }
    catch (err) { setStatus(`Login failed: ${err.message}`, true); }
  });

  btnLogout.addEventListener("click", () => { Auth.logout(); setStatus(""); renderAuth(); });
  btnSubmit.addEventListener("click", handleSubmit);

  renderAuth();
});
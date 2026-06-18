// Task pane controller

Office.onReady(async () => {
  const btnLogin    = document.getElementById("btnLogin");
  const btnLogout   = document.getElementById("btnLogout");
  const authSection = document.getElementById("authSection");
  const appSection  = document.getElementById("appSection");

  // ── Tab switching ──────────────────────────────────────────────
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(btn.dataset.tab).classList.add("active");
    });
  });

  // ── Tab 1: Create Case ────────────────────────────────────────

  const btnSubmit      = document.getElementById("btnSubmit");
  const createStatusEl = document.getElementById("createStatus");

  const contactSearch        = document.getElementById("contactSearch");
  const contactDropdown      = document.getElementById("contactDropdown");
  const contactSelected      = document.getElementById("contactSelected");
  const contactSelectedName  = document.getElementById("contactSelectedName");
  const btnClearContact      = document.getElementById("btnClearContact");
  const contactIdInput       = document.getElementById("contactId");
  const accountIdInput       = document.getElementById("accountId");
  const accountNameRow       = document.getElementById("accountNameRow");
  const accountNameDisplay   = document.getElementById("accountNameDisplay");

  const supplierSearch         = document.getElementById("supplierSearch");
  const supplierDropdown       = document.getElementById("supplierDropdown");
  const supplierSelected       = document.getElementById("supplierSelected");
  const supplierSelectedName   = document.getElementById("supplierSelectedName");
  const btnClearSupplier       = document.getElementById("btnClearSupplier");
  const supplierIdInput        = document.getElementById("supplierId");
  const supplierContactRow     = document.getElementById("supplierContactRow");
  const supplierContactSelect  = document.getElementById("supplierContactSelect");
  const supplierContactIdInput = document.getElementById("supplierContactId");

  function setCreateStatus(msg, isError = false, link = null) {
    createStatusEl.innerHTML = "";
    if (link) {
      createStatusEl.innerHTML = `${msg} <a href="${link}" target="_blank" rel="noopener">Open in Zoho</a>`;
    } else {
      createStatusEl.textContent = msg;
    }
    createStatusEl.className = isError ? "status error" : "status info";
    createStatusEl.hidden = !msg;
  }

  function showLoading(show) {
    btnSubmit.disabled = show;
    btnSubmit.textContent = show ? "Creating…" : "Create Case";
  }

  // Contact
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
      const el = document.createElement("div");
      el.className = "dropdown-item dropdown-empty";
      el.textContent = "No contacts found";
      contactDropdown.appendChild(el);
    } else {
      contacts.forEach((c) => {
        const el = document.createElement("div");
        el.className = "dropdown-item";
        el.innerHTML = `<strong>${c.Full_Name || ""}</strong><span>${c.Email || ""}${c.Account_Name ? " · " + c.Account_Name.name : ""}</span>`;
        el.addEventListener("mousedown", (e) => {
          e.preventDefault();
          selectContact(c.id, c.Full_Name || "", c.Email || "", c.Account_Name?.id || "", c.Account_Name?.name || "");
        });
        contactDropdown.appendChild(el);
      });
    }
    contactDropdown.hidden = false;
  }

  let contactTimer = null;
  contactSearch.addEventListener("input", () => {
    clearTimeout(contactTimer);
    const q = contactSearch.value.trim();
    if (!q || q.length < 2) { contactDropdown.hidden = true; return; }
    contactTimer = setTimeout(async () => {
      const results = await Zoho.searchContacts(q);
      showContactDropdown(results);
    }, 350);
  });
  contactSearch.addEventListener("blur",  () => { setTimeout(() => { contactDropdown.hidden = true; }, 150); });
  contactSearch.addEventListener("focus", () => { if (contactSearch.value.trim().length >= 2) contactDropdown.hidden = false; });
  btnClearContact.addEventListener("click", clearContact);

  // Supplier
  async function selectSupplier(id, name) {
    supplierIdInput.value = id || "";
    supplierSelectedName.textContent = name || "";
    supplierSelected.hidden = false;
    supplierSearch.hidden   = true;
    supplierDropdown.hidden = true;
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
      const el = document.createElement("div");
      el.className = "dropdown-item dropdown-empty";
      el.textContent = "No accounts found";
      supplierDropdown.appendChild(el);
    } else {
      accounts.forEach((a) => {
        const el = document.createElement("div");
        el.className = "dropdown-item";
        el.innerHTML = `<strong>${a.Account_Name || ""}</strong>`;
        el.addEventListener("mousedown", (e) => {
          e.preventDefault();
          selectSupplier(a.id, a.Account_Name || "");
        });
        supplierDropdown.appendChild(el);
      });
    }
    supplierDropdown.hidden = false;
  }

  let supplierTimer = null;
  supplierSearch.addEventListener("input", () => {
    clearTimeout(supplierTimer);
    const q = supplierSearch.value.trim();
    if (!q || q.length < 2) { supplierDropdown.hidden = true; return; }
    supplierTimer = setTimeout(async () => {
      try {
        const results = await Zoho.searchAccounts(q);
        showSupplierDropdown(results);
      } catch (err) {
        setCreateStatus(`Supplier search: ${err.message}`, true);
      }
    }, 350);
  });
  supplierSearch.addEventListener("blur",  () => { setTimeout(() => { supplierDropdown.hidden = true; }, 150); });
  supplierSearch.addEventListener("focus", () => { if (supplierSearch.value.trim().length >= 2) supplierDropdown.hidden = false; });
  btnClearSupplier.addEventListener("click", clearSupplier);
  supplierContactSelect.addEventListener("change", () => { supplierContactIdInput.value = supplierContactSelect.value; });

  // Populate Create Case form
  async function populateCreateForm() {
    const item = Office.context.mailbox.item;
    document.getElementById("subject").value = item.subject || "";

    const senderEmail = item.from?.emailAddress || "";
    if (senderEmail) {
      contactSearch.value = senderEmail;
      setCreateStatus("Looking up contact in Zoho…");
      try {
        const contact = await Zoho.findContactByEmail(senderEmail);
        if (contact) {
          selectContact(contact.id, contact.Full_Name || senderEmail, contact.Email || senderEmail, contact.Account_Name?.id || "", contact.Account_Name?.name || "");
        } else {
          contactSearch.value = senderEmail;
        }
        setCreateStatus("");
      } catch (err) {
        setCreateStatus(`Contact lookup: ${err.message}`, true);
        contactSearch.value = senderEmail;
      }
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("description").value = result.value?.trim() || "";
      }
    });
  }

  async function handleSubmit() {
    const subject = document.getElementById("subject").value.trim();
    if (!subject) { setCreateStatus("Subject is required.", true); return; }

    const type = document.getElementById("caseType").value;
    if (!type || type === "-None-") { setCreateStatus("Case Type is required.", true); return; }

    const description       = document.getElementById("description").value.trim();
    const contactId         = contactIdInput.value.trim() || null;
    const accountId         = accountIdInput.value.trim() || null;
    const supplierId        = supplierIdInput.value.trim() || null;
    const supplierContactId = supplierContactIdInput.value.trim() || null;
    const status            = document.getElementById("caseStatus").value;
    const priority          = document.getElementById("casePriority").value;

    showLoading(true);
    setCreateStatus("");

    try {
      const { id: caseId, url: caseLink } = await Zoho.createCase({
        subject, description, contactId, accountId,
        supplierId, supplierContactId, type, status, priority,
      });
      setCreateStatus(`Case ${caseId} created!`, false, caseLink);

      if (document.getElementById("includeAttachments").checked) {
        await uploadAttachmentsToCase(caseId);
      }
    } catch (err) {
      setCreateStatus(`Error: ${err.message}`, true);
    } finally {
      showLoading(false);
    }
  }

  btnSubmit.addEventListener("click", handleSubmit);

  // ── Tab 2: Associate Email ────────────────────────────────────

  const caseSearch        = document.getElementById("caseSearch");
  const caseDropdown      = document.getElementById("caseDropdown");
  const caseSelectedEl    = document.getElementById("caseSelected");
  const caseSelectedName  = document.getElementById("caseSelectedName");
  const btnClearCase      = document.getElementById("btnClearCase");
  const caseIdInput       = document.getElementById("caseId");
  const btnAssociate      = document.getElementById("btnAssociate");
  const associateStatusEl = document.getElementById("associateStatus");
  const btnAddNote        = document.getElementById("btnAddNote");
  const noteStatusEl      = document.getElementById("noteStatus");

  function setAssociateStatus(msg, isError = false, link = null) {
    associateStatusEl.innerHTML = "";
    if (link) {
      associateStatusEl.innerHTML = `${msg} <a href="${link}" target="_blank" rel="noopener">Open in Zoho</a>`;
    } else {
      associateStatusEl.textContent = msg;
    }
    associateStatusEl.className = isError ? "status error" : "status info";
    associateStatusEl.hidden = !msg;
  }

  function selectCase(id, caseNumber, subject) {
    caseIdInput.value = id || "";
    caseSelectedName.textContent = `#${caseNumber} — ${subject}`;
    caseSelectedEl.hidden = false;
    caseSearch.hidden     = true;
    caseDropdown.hidden   = true;
  }

  function clearCase() {
    caseIdInput.value = "";
    caseSelectedEl.hidden = true;
    caseSearch.hidden     = false;
    caseSearch.value      = "";
    caseDropdown.hidden   = true;
    caseSearch.focus();
  }

  function showCaseDropdown(cases) {
    caseDropdown.innerHTML = "";
    if (!cases.length) {
      const el = document.createElement("div");
      el.className = "dropdown-item dropdown-empty";
      el.textContent = "No cases found";
      caseDropdown.appendChild(el);
    } else {
      cases.forEach((c) => {
        const el = document.createElement("div");
        el.className = "dropdown-item";
        el.innerHTML = `<strong>#${c.Case_Number}</strong><span>${c.Subject || ""}</span>`;
        el.addEventListener("mousedown", (e) => {
          e.preventDefault();
          selectCase(c.id, c.Case_Number, c.Subject || "");
        });
        caseDropdown.appendChild(el);
      });
    }
    caseDropdown.hidden = false;
  }

  let caseTimer = null;
  caseSearch.addEventListener("input", () => {
    clearTimeout(caseTimer);
    const q = caseSearch.value.trim();
    if (!q || q.length < 2) { caseDropdown.hidden = true; return; }
    caseTimer = setTimeout(async () => {
      try {
        const results = await Zoho.searchCases(q);
        showCaseDropdown(results);
      } catch (err) {
        setAssociateStatus(`Case search: ${err.message}`, true);
      }
    }, 350);
  });
  caseSearch.addEventListener("blur",  () => { setTimeout(() => { caseDropdown.hidden = true; }, 150); });
  caseSearch.addEventListener("focus", () => { if (caseSearch.value.trim().length >= 2) caseDropdown.hidden = false; });
  btnClearCase.addEventListener("click", clearCase);

  // Build RFC 2822 .eml content from the current email
  function buildEml(subject, from, toList, ccList, dateStr, bodyText) {
    const formatAddr = (a) => a.displayName ? `"${a.displayName}" <${a.emailAddress}>` : a.emailAddress;
    const toHeader  = toList.map(formatAddr).join(", ");
    const ccHeader  = ccList.length ? `Cc: ${ccList.map(formatAddr).join(", ")}\r\n` : "";
    const safeBody  = bodyText.replace(/\r?\n/g, "\r\n");

    return [
      `From: ${formatAddr(from)}`,
      `To: ${toHeader}`,
      ccHeader.trimEnd(),
      `Subject: ${subject}`,
      `Date: ${dateStr}`,
      `MIME-Version: 1.0`,
      `Content-Type: text/plain; charset=UTF-8`,
      `Content-Transfer-Encoding: 8bit`,
      ``,
      safeBody,
    ].filter((l) => l !== undefined).join("\r\n");
  }

  async function getEmailAsEml() {
    const item = Office.context.mailbox.item;

    const subject  = item.subject || "(no subject)";
    const from     = item.from || { displayName: "", emailAddress: "" };
    const toList   = item.to  || [];
    const ccList   = item.cc  || [];
    const dateStr  = (item.dateTimeCreated instanceof Date
      ? item.dateTimeCreated
      : new Date()
    ).toUTCString().replace("GMT", "+0000");

    const bodyText = await new Promise((resolve) => {
      item.body.getAsync(Office.CoercionType.Text, (r) => {
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value || "" : "");
      });
    });

    const eml = buildEml(subject, from, toList, ccList, dateStr, bodyText);
    return { eml, subject };
  }

  async function handleAssociate() {
    const caseId = caseIdInput.value.trim();
    if (!caseId) { setAssociateStatus("Please select a case.", true); return; }

    btnAssociate.disabled = true;
    btnAssociate.textContent = "Associating…";
    setAssociateStatus("");

    try {
      // Upload the email itself as an .eml file
      const { eml, subject } = await getEmailAsEml();
      const emlBlob = new Blob([eml], { type: "application/octet-stream" });
      const safeSubject = (subject || "email").replace(/[\\/:*?"<>|]/g, "_");
      await Zoho.attachFileToCaseRaw(caseId, `${safeSubject}.eml`, emlBlob);

      // Upload original attachments if checked
      if (document.getElementById("associateAttachments").checked) {
        await uploadAttachmentsToCase(caseId);
      }

      const caseLink = await Zoho.getCaseUrl(caseId);
      setAssociateStatus("Email associated to case!", false, caseLink);
    } catch (err) {
      setAssociateStatus(`Error: ${err.message}`, true);
    } finally {
      btnAssociate.disabled = false;
      btnAssociate.textContent = "Associate Email";
    }
  }

  btnAssociate.addEventListener("click", handleAssociate);

  async function handleAddNote() {
    const caseId = caseIdInput.value.trim();
    if (!caseId) { setNoteStatus("Please select a case first.", true); return; }

    const noteSubject = document.getElementById("noteSubject").value.trim();
    const noteContent = document.getElementById("noteContent").value.trim();
    if (!noteSubject && !noteContent) { setNoteStatus("Enter a subject or note content.", true); return; }

    btnAddNote.disabled = true;
    btnAddNote.textContent = "Adding…";
    setNoteStatus("");

    try {
      await Zoho.addNoteToCase(caseId, noteSubject || "(no subject)", noteContent);
      const caseLink = await Zoho.getCaseUrl(caseId);
      setNoteStatus("Note added!", false, caseLink);
      document.getElementById("noteSubject").value = "";
      document.getElementById("noteContent").value = "";
    } catch (err) {
      setNoteStatus(`Error: ${err.message}`, true);
    } finally {
      btnAddNote.disabled = false;
      btnAddNote.textContent = "Add Note";
    }
  }

  function setNoteStatus(msg, isError = false, link = null) {
    noteStatusEl.innerHTML = "";
    if (link) {
      noteStatusEl.innerHTML = `${msg} <a href="${link}" target="_blank" rel="noopener">Open in Zoho</a>`;
    } else {
      noteStatusEl.textContent = msg;
    }
    noteStatusEl.className = isError ? "status error" : "status info";
    noteStatusEl.hidden = !msg;
  }

  btnAddNote.addEventListener("click", handleAddNote);

  // ── Shared attachment uploader ────────────────────────────────

  async function uploadAttachmentsToCase(caseId) {
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
            if (failed > 0) setAssociateStatus(`Done — ${failed} attachment(s) failed to upload.`, true);
            resolve();
          }
        });
      });
    });
  }

  // ── Auth ──────────────────────────────────────────────────────

  function renderAuth() {
    const loggedIn = Auth.isLoggedIn();
    authSection.hidden = loggedIn;
    appSection.hidden  = !loggedIn;
    btnLogout.hidden   = !loggedIn;
    if (loggedIn) populateCreateForm();
  }

  btnLogin.addEventListener("click", async () => {
    setCreateStatus("Opening Zoho login…");
    try { await Auth.ensureToken(); setCreateStatus(""); renderAuth(); }
    catch (err) { setCreateStatus(`Login failed: ${err.message}`, true); }
  });

  btnLogout.addEventListener("click", () => { Auth.logout(); renderAuth(); });

  renderAuth();
});
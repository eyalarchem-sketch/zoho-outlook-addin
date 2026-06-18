// Task pane controller

Office.onReady(async () => {
  const btnLogin    = document.getElementById("btnLogin");
  const btnLogout   = document.getElementById("btnLogout");
  const btnSubmit   = document.getElementById("btnSubmit");
  const authSection = document.getElementById("authSection");
  const formSection = document.getElementById("formSection");
  const statusEl    = document.getElementById("status");

  // Contact search elements
  const contactSearch   = document.getElementById("contactSearch");
  const contactDropdown = document.getElementById("contactDropdown");
  const contactSelected = document.getElementById("contactSelected");
  const contactSelectedName = document.getElementById("contactSelectedName");
  const btnClearContact = document.getElementById("btnClearContact");
  const contactIdInput  = document.getElementById("contactId");

  function setStatus(msg, isError = false) {
    statusEl.textContent = msg;
    statusEl.className = isError ? "status error" : "status info";
    statusEl.hidden = !msg;
  }

  function showLoading(show) {
    btnSubmit.disabled = show;
    btnSubmit.textContent = show ? "Creating…" : "Create Case";
  }

  // --- Contact search ---
  function selectContact(id, name, email) {
    contactIdInput.value = id || "";
    contactSelectedName.textContent = id ? `${name} — ${email}` : "(none)";
    contactSelected.hidden = false;
    contactSearch.hidden = true;
    contactDropdown.hidden = true;
  }

  function clearContact() {
    contactIdInput.value = "";
    contactSelected.hidden = true;
    contactSearch.hidden = false;
    contactSearch.value = "";
    contactDropdown.hidden = true;
    contactSearch.focus();
  }

  function showDropdown(contacts) {
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
        item.innerHTML = `<strong>${c.Full_Name || ""}</strong><span>${c.Email || ""}</span>`;
        item.addEventListener("mousedown", (e) => {
          e.preventDefault();
          selectContact(c.id, c.Full_Name || "", c.Email || "");
        });
        contactDropdown.appendChild(item);
      });
    }
    contactDropdown.hidden = false;
  }

  let searchTimer = null;
  contactSearch.addEventListener("input", () => {
    clearTimeout(searchTimer);
    const q = contactSearch.value.trim();
    if (!q || q.length < 2) { contactDropdown.hidden = true; return; }
    searchTimer = setTimeout(async () => {
      const results = await Zoho.searchContacts(q);
      showDropdown(results);
    }, 350);
  });

  contactSearch.addEventListener("blur", () => {
    setTimeout(() => { contactDropdown.hidden = true; }, 150);
  });

  contactSearch.addEventListener("focus", () => {
    if (contactSearch.value.trim().length >= 2) contactDropdown.hidden = false;
  });

  btnClearContact.addEventListener("click", clearContact);

  // --- Populate form from email ---
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
          selectContact(contact.id, contact.Full_Name || senderEmail, contact.Email || senderEmail);
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

    const description = document.getElementById("description").value.trim();
    const contactId   = contactIdInput.value.trim() || null;
    const status      = document.getElementById("caseStatus").value;
    const priority    = document.getElementById("casePriority").value;

    showLoading(true);
    setStatus("");

    try {
      const caseId = await Zoho.createCase({ subject, description, contactId, status, priority });
      setStatus(`Case created! ID: ${caseId}`);

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
    const attachments = item.attachments || [];
    if (!attachments.length) return;

    return new Promise((resolve) => {
      let remaining = attachments.filter(a => !a.isInline).length;
      if (!remaining) return resolve();
      let failed = 0;

      attachments.filter(a => !a.isInline).forEach((att) => {
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
            if (failed > 0) setStatus(`Case created — ${failed} attachment(s) failed to upload.`, true);
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
    btnLogout.hidden = !loggedIn;
    if (loggedIn) populateForm();
  }

  btnLogin.addEventListener("click", async () => {
    setStatus("Opening Zoho login…");
    try {
      await Auth.ensureToken();
      setStatus("");
      renderAuth();
    } catch (err) {
      setStatus(`Login failed: ${err.message}`, true);
    }
  });

  btnLogout.addEventListener("click", () => { Auth.logout(); setStatus(""); renderAuth(); });
  btnSubmit.addEventListener("click", handleSubmit);

  renderAuth();
});
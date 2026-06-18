// Task pane controller — runs after Office.js is ready

Office.onReady(async () => {
  const btnLogin = document.getElementById("btnLogin");
  const btnLogout = document.getElementById("btnLogout");
  const btnSubmit = document.getElementById("btnSubmit");
  const authSection = document.getElementById("authSection");
  const formSection = document.getElementById("formSection");
  const statusEl = document.getElementById("status");

  function setStatus(msg, isError = false) {
    statusEl.textContent = msg;
    statusEl.className = isError ? "status error" : "status info";
    statusEl.hidden = !msg;
  }

  function showLoading(show) {
    btnSubmit.disabled = show;
    btnSubmit.textContent = show ? "Creating…" : "Create Case";
  }

  function renderAuth() {
    const loggedIn = Auth.isLoggedIn();
    authSection.hidden = loggedIn;
    formSection.hidden = !loggedIn;
    btnLogout.hidden = !loggedIn;
    if (loggedIn) populateForm();
  }

  // --- Populate form from the active email ---
  async function populateForm() {
    const item = Office.context.mailbox.item;

    // Subject
    document.getElementById("subject").value = item.subject || "";

    // Sender email → contact lookup
    const senderEmail = item.from?.emailAddress || "";
    document.getElementById("senderEmail").value = senderEmail;

    if (senderEmail) {
      setStatus("Looking up contact in Zoho…");
      try {
        const contact = await Zoho.findContactByEmail(senderEmail);
        if (contact) {
          document.getElementById("contactId").value = contact.id;
          document.getElementById("contactName").value = contact.Full_Name || senderEmail;
          setStatus("");
        } else {
          document.getElementById("contactId").value = "";
          document.getElementById("contactName").value = "(no matching contact found)";
          setStatus("");
        }
      } catch (err) {
        setStatus(`Contact lookup failed: ${err.message}`, true);
      }
    }

    // Body — getBodyAsync strips HTML to plain text
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("description").value = result.value?.trim() || "";
      }
    });
  }

  // --- Submit handler ---
  async function handleSubmit() {
    const subject = document.getElementById("subject").value.trim();
    if (!subject) {
      setStatus("Subject is required.", true);
      return;
    }

    const description = document.getElementById("description").value.trim();
    const contactId = document.getElementById("contactId").value.trim() || null;
    const status = document.getElementById("caseStatus").value;
    const priority = document.getElementById("casePriority").value;

    showLoading(true);
    setStatus("");

    try {
      const caseId = await Zoho.createCase({ subject, description, contactId, status, priority });
      setStatus(`Case created! ID: ${caseId}`);

      // Upload attachments if any
      const attachCheckbox = document.getElementById("includeAttachments");
      if (attachCheckbox.checked) {
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

    return new Promise((resolve) => {
      item.attachments; // The attachments array is synchronous
      const attachments = item.attachments || [];
      if (!attachments.length) return resolve();

      let remaining = attachments.length;
      let failed = 0;

      attachments.forEach((att) => {
        // Skip inline images
        if (att.isInline) {
          remaining--;
          if (remaining === 0) resolve();
          return;
        }

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
            } catch {
              failed++;
            }
          } else {
            failed++;
          }

          remaining--;
          if (remaining === 0) {
            if (failed > 0) setStatus(`Case created — ${failed} attachment(s) failed to upload.`, true);
            resolve();
          }
        });
      });
    });
  }

  // --- Wire up events ---
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

  btnLogout.addEventListener("click", () => {
    Auth.logout();
    setStatus("");
    renderAuth();
  });

  btnSubmit.addEventListener("click", handleSubmit);

  // Initial render
  renderAuth();
});

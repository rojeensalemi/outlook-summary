// ============================================================
// CONFIG — No API key here. It's loaded from Outlook roaming settings.
// ============================================================
const ANTHROPIC_MODEL = "claude-sonnet-4-5";
const HOURS_BACK = 24;
const MAX_EMAILS = 25;
const BODY_CHAR_LIMIT = 1500;

const SETTING_KEY = "anthropicApiKey";

// ============================================================
// Office.js initialization
// ============================================================
Office.onReady(() => {
    document.getElementById("refreshBtn").addEventListener("click", runSummary);
    document.getElementById("settingsBtn").addEventListener("click", openSettings);
    document.getElementById("cancelSettings").addEventListener("click", closeSettings);
    document.getElementById("saveSettings").addEventListener("click", saveSettings);

    document.querySelectorAll(".tab").forEach(tab => {
        tab.addEventListener("click", () => switchTab(tab.dataset.pane));
    });

    // If no key is saved yet, pop the settings dialog automatically
    if (!getApiKey()) {
        openSettings();
    }
});

function switchTab(paneName) {
    document.querySelectorAll(".tab").forEach(t => {
        t.classList.toggle("active", t.dataset.pane === paneName);
    });
    document.querySelectorAll(".pane").forEach(p => {
        p.classList.toggle("active", p.id === `pane-${paneName}`);
    });
}

// ============================================================
// API key storage (Outlook roaming settings)
// ============================================================
function getApiKey() {
    return Office.context.roamingSettings.get(SETTING_KEY) || "";
}

function openSettings() {
    document.getElementById("apiKeyInput").value = getApiKey();
    document.getElementById("settingsOverlay").classList.add("show");
    document.getElementById("apiKeyInput").focus();
}

function closeSettings() {
    document.getElementById("settingsOverlay").classList.remove("show");
}

function saveSettings() {
    const key = document.getElementById("apiKeyInput").value.trim();
    if (!key) {
        alert("Please enter an API key.");
        return;
    }
    Office.context.roamingSettings.set(SETTING_KEY, key);
    Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            closeSettings();
        } else {
            alert("Failed to save key: " + (result.error && result.error.message));
        }
    });
}

// ============================================================
// Main flow
// ============================================================
async function runSummary() {
    const btn = document.getElementById("refreshBtn");
    const summaryEl = document.getElementById("summary-content");
    const todoEl = document.getElementById("todo-content");

    const apiKey = getApiKey();
    if (!apiKey) {
        summaryEl.innerHTML = '<div class="error">No API key set. Click the ⚙ button to add one.</div>';
        openSettings();
        return;
    }

    btn.disabled = true;
    btn.textContent = "Working...";
    summaryEl.innerHTML = '<div class="spinner">Fetching unread emails...</div>';
    todoEl.textContent = "";

    try {
        const emails = await fetchUnreadEmails();
        if (!emails || emails.length === 0) {
            summaryEl.textContent = `No unread emails in the last ${HOURS_BACK} hours.`;
            todoEl.textContent = "";
            stampTime();
            return;
        }

        summaryEl.innerHTML = `<div class="spinner">Analyzing ${emails.length} emails...</div>`;

        const emailDump = formatEmailsForPrompt(emails);
        const response = await callClaude(apiKey, emailDump);
        const { summary, todo } = splitResponse(response);

        summaryEl.innerHTML = renderText(summary);
        todoEl.innerHTML = renderText(todo);
        stampTime();
    } catch (err) {
        summaryEl.innerHTML = `<div class="error">Error: ${escapeHtml(err.message)}</div>`;
        todoEl.textContent = "";
        console.error(err);
    } finally {
        btn.disabled = false;
        btn.textContent = "Summarize";
    }
}

function stampTime() {
    const now = new Date();
    document.getElementById("timestamp").textContent =
        "Generated: " + now.toLocaleString();
}

// ============================================================
// Email fetching via EWS (works on Outlook Web + Desktop)
// ============================================================
function fetchUnreadEmails() {
    return new Promise((resolve, reject) => {
        const cutoff = new Date(Date.now() - HOURS_BACK * 3600 * 1000).toISOString();

        const soap = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:FindItem xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
          <t:FieldURI FieldURI="message:From"/>
          <t:FieldURI FieldURI="message:IsRead"/>
          <t:FieldURI FieldURI="item:Body"/>
        </t:AdditionalProperties>
        <t:BodyType>Text</t:BodyType>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="${MAX_EMAILS}" Offset="0" BasePoint="Beginning"/>
      <m:Restriction>
        <t:And>
          <t:IsEqualTo>
            <t:FieldURI FieldURI="message:IsRead"/>
            <t:FieldURIOrConstant>
              <t:Constant Value="false"/>
            </t:FieldURIOrConstant>
          </t:IsEqualTo>
          <t:IsGreaterThanOrEqualTo>
            <t:FieldURI FieldURI="item:DateTimeReceived"/>
            <t:FieldURIOrConstant>
              <t:Constant Value="${cutoff}"/>
            </t:FieldURIOrConstant>
          </t:IsGreaterThanOrEqualTo>
        </t:And>
      </m:Restriction>
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;

        Office.context.mailbox.makeEwsRequestAsync(soap, (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error("EWS request failed: " + (result.error && result.error.message)));
                return;
            }
            try {
                resolve(parseEwsResponse(result.value));
            } catch (e) {
                reject(e);
            }
        });
    });
}

function parseEwsResponse(xmlString) {
    const doc = new DOMParser().parseFromString(xmlString, "text/xml");
    const nsT = "http://schemas.microsoft.com/exchange/services/2006/types";
    const messages = doc.getElementsByTagNameNS(nsT, "Message");

    const results = [];
    for (let i = 0; i < messages.length && i < MAX_EMAILS; i++) {
        const m = messages[i];
        const getText = (tag) => {
            const el = m.getElementsByTagNameNS(nsT, tag)[0];
            return el ? el.textContent : "";
        };
        const fromMailbox = m.getElementsByTagNameNS(nsT, "From")[0];
        let senderName = "", senderEmail = "";
        if (fromMailbox) {
            const nameEl = fromMailbox.getElementsByTagNameNS(nsT, "Name")[0];
            const emailEl = fromMailbox.getElementsByTagNameNS(nsT, "EmailAddress")[0];
            senderName = nameEl ? nameEl.textContent : "";
            senderEmail = emailEl ? emailEl.textContent : "";
        }

        results.push({
            subject: getText("Subject"),
            received: getText("DateTimeReceived"),
            senderName,
            senderEmail,
            body: getText("Body")
        });
    }
    return results;
}

function formatEmailsForPrompt(emails) {
    return emails.map((e, i) => {
        let body = (e.body || "").replace(/\s+/g, " ").trim();
        if (body.length > BODY_CHAR_LIMIT) {
            body = body.slice(0, BODY_CHAR_LIMIT) + "...[truncated]";
        }
        return `---EMAIL ${i + 1}---
From: ${e.senderName} <${e.senderEmail}>
Subject: ${e.subject}
Received: ${e.received}
Body: ${body}
`;
    }).join("\n");
}

// ============================================================
// Anthropic API call
// ============================================================
async function callClaude(apiKey, emailDump) {
    const systemPrompt =
        "You are an executive assistant reviewing a user's unread emails. " +
        "Output EXACTLY two sections separated by the literal line '===TODO==='. " +
        "Section 1 (before ===TODO===): A concise bullet-point summary grouped by theme or sender importance. " +
        "One bullet per line starting with '- '. Keep it scannable. " +
        "Section 2 (after ===TODO===): A prioritized to-do list of concrete actions the user should take. " +
        "Each item on its own line starting with '- '. Mark urgent items with [URGENT] at the start. " +
        "Skip newsletters, marketing, and automated notifications unless they contain action items. " +
        "Do not include any preamble, headers, closing remarks, or markdown like ** or ## or backticks. " +
        "Use plain text only.";

    const userPrompt =
        `Here are my unread emails from the last ${HOURS_BACK} hours:\n\n${emailDump}`;

    const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "anthropic-dangerous-direct-browser-access": "true"
        },
        body: JSON.stringify({
            model: ANTHROPIC_MODEL,
            max_tokens: 2048,
            system: systemPrompt,
            messages: [{ role: "user", content: userPrompt }]
        })
    });

    if (!res.ok) {
        const errText = await res.text();
        throw new Error(`API ${res.status}: ${errText}`);
    }

    const data = await res.json();
    const textBlock = (data.content || []).find(b => b.type === "text");
    return textBlock ? textBlock.text : "";
}

// ============================================================
// Helpers
// ============================================================
function splitResponse(resp) {
    const marker = "===TODO===";
    const idx = resp.indexOf(marker);
    if (idx === -1) {
        return { summary: resp, todo: "(no to-do list returned)" };
    }
    return {
        summary: resp.slice(0, idx).trim(),
        todo: resp.slice(idx + marker.length).trim()
    };
}

function stripMarkdown(s) {
    return s
        .replace(/\*\*/g, "")
        .replace(/__/g, "")
        .replace(/`/g, "")
        .replace(/^#{1,3}\s*/gm, "");
}

function renderText(s) {
    const cleaned = stripMarkdown(s);
    return escapeHtml(cleaned).replace(
        /\[URGENT\]/g,
        '<span class="urgent">[URGENT]</span>'
    );
}

function escapeHtml(s) {
    return s
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}

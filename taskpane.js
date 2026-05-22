// ── Entry point ──────────────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("generate-btn").addEventListener("click", onGenerate);
  }
});

const WORKER_URL = "https://your-worker.your-domain.workers.dev/get-replies";

// ── Main handler ─────────────────────────────────────────────────────────────
async function onGenerate() {
  setStatus("Reading email...", "info");
  clearOptions();

  try {
    const emailBody = await getEmailBody();
    if (!emailBody) {
      setStatus("No email body found.", "warn");
      return;
    }

    setStatus("Generating replies...", "info");
    const options = await fetchReplies(emailBody);
    renderOptions(options);
    setStatus("Select a reply to insert it.", "success");
  } catch (err) {
    setStatus(`Error: ${err.message}`, "error");
    console.error(err);
  }
}

// ── Office.js helpers ────────────────────────────────────────────────────────

/** Reads the current email body as plain text */
function getEmailBody() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      { asyncContext: "body" },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value.trim());
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}

/** Inserts chosen text into the compose body */
function insertReply(text) {
  Office.context.mailbox.item.body.setAsync(
    text,
    { coercionType: Office.CoercionType.Text },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Failed to insert reply.", "error");
      } else {
        setStatus("Reply inserted ✓", "success");
      }
    }
  );
}

// ── API call ─────────────────────────────────────────────────────────────────

async function fetchReplies(emailBody) {
  const response = await fetch(WORKER_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ emailBody }),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error ?? `HTTP ${response.status}`);
  }

  const data = await response.json();

  if (!Array.isArray(data.options)) {
    throw new Error("Unexpected response from server.");
  }

  return data.options;
}

// ── UI helpers ───────────────────────────────────────────────────────────────

const TONE_LABELS = [
  "📋 Formal & Detailed",
  "😊 Friendly & Concise",
  "🚀 Action-Oriented",
  "🤝 Empathetic",
];

function renderOptions(options) {
  const container = document.getElementById("options-container");
  clearOptions();

  options.forEach((text, i) => {
    const card = document.createElement("div");
    card.className = "option-card";

    const label = document.createElement("p");
    label.className = "option-label";
    label.textContent = TONE_LABELS[i] ?? `Option ${i + 1}`;

    const preview = document.createElement("p");
    preview.className = "option-preview";
    preview.textContent = text.length > 120 ? text.slice(0, 120) + "…" : text;

    const btn = document.createElement("button");
    btn.className = "ms-Button ms-Button--primary";
    btn.textContent = "Use this reply";
    btn.addEventListener("click", () => insertReply(text));

    card.append(label, preview, btn);
    container.appendChild(card);
  });
}

function clearOptions() {
  document.getElementById("options-container").innerHTML = "";
}

function setStatus(msg, type = "info") {
  const el = document.getElementById("status-message");
  el.textContent = msg;
  el.className = `status status--${type}`;   // style via CSS
}
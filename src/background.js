const DEFAULT_SETTINGS = {
  enabled: false,
  botToken: "",
  chatId: "",
  accessUrl: "https://teams.microsoft.com/v2/",
  includePageUrl: false,
  dedupeTtlMinutes: 5,
  minTextLength: 8
};

const recentMessages = new Map();
const pendingMessages = new Map();
const SENT_LOGS_KEY = "sentLogs";
const MAX_SENT_LOGS = 100;

chrome.runtime.onInstalled.addListener(async () => {
  const current = await chrome.storage.local.get(Object.keys(DEFAULT_SETTINGS));
  await chrome.storage.local.set(normalizeSettings({ ...DEFAULT_SETTINGS, ...compact(current) }));
});

if (chrome.action && chrome.action.onClicked) {
  chrome.action.onClicked.addListener(() => {
    chrome.runtime.openOptionsPage();
  });
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message.type !== "string") {
    return false;
  }

  if (message.type === "teams-notification-candidate") {
    handleNotificationCandidate(message.payload, sender)
      .then((result) => sendResponse(result))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "test-telegram") {
    sendTestTelegramMessage()
      .then((result) => sendResponse(result))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "clear-sent-logs") {
    chrome.storage.local.set({ [SENT_LOGS_KEY]: [] })
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  return false;
});

async function handleNotificationCandidate(payload, sender) {
  const settings = await loadSettings();
  if (!settings.enabled) {
    return { ok: true, skipped: "disabled" };
  }

  const normalized = normalizePayload(payload, sender, settings);
  if (!normalized) {
    return { ok: true, skipped: "empty" };
  }

  if (normalized.preview.length < Number(settings.minTextLength || 0)) {
    return { ok: true, skipped: "too-short" };
  }

  const hash = hashMessage(normalized);
  const ttlMs = Math.max(1, Number(settings.dedupeTtlMinutes || 5)) * 60 * 1000;
  pruneRecent(Date.now(), ttlMs);

  if (recentMessages.has(hash)) {
    return { ok: true, skipped: "duplicate" };
  }

  queueTelegramMessage(hash, normalized, settings);
  return { ok: true, queued: true };
}

async function sendTestTelegramMessage() {
  const settings = await loadSettings();
  const message = {
    source: "Extension test",
    sender: "Teams Telegram Notifier",
    preview: "Test message from the Chrome extension.",
    pageUrl: "",
    observedAt: new Date().toISOString()
  };
  await sendTelegramMessage(settings, formatTelegramMessage(message, settings));
  await appendSentLog(message);
  return { ok: true, sent: true };
}

async function loadSettings() {
  const stored = await chrome.storage.local.get(Object.keys(DEFAULT_SETTINGS));
  return normalizeSettings({ ...DEFAULT_SETTINGS, ...compact(stored) });
}

function compact(value) {
  return Object.fromEntries(
    Object.entries(value || {}).filter(([, item]) => item !== undefined && item !== null)
  );
}

function normalizeSettings(settings) {
  const nextSettings = { ...settings };
  if (isDeprecatedCallbackUrl(nextSettings.accessUrl)) {
    nextSettings.accessUrl = DEFAULT_SETTINGS.accessUrl;
  }

  return nextSettings;
}

function isDeprecatedCallbackUrl(value) {
  return /^intent:\/\//i.test(String(value || "")) ||
    /^firefox:\/\//i.test(String(value || "")) ||
    /org\.mozilla\.firefox/i.test(String(value || ""));
}

function normalizePayload(payload, sender, settings) {
  const rawText = extractSingleNotificationText(cleanMultilineText(payload && payload.text));
  const title = cleanText(payload && payload.title);
  const body = extractSingleNotificationText(cleanMultilineText(payload && payload.body));
  const pageUrl = cleanText(payload && payload.pageUrl) || (sender.tab && sender.tab.url) || "";
  const observedAt = cleanText(payload && payload.observedAt) || new Date().toISOString();
  const source = cleanText(payload && payload.source) || "Teams";

  let senderName = title;
  let preview = body || rawText;

  if (isGenericTeamsSender(senderName) && body) {
    const parts = body.split("\n").map((part) => part.trim()).filter(Boolean);
    if (parts.length > 1) {
      senderName = parts[0];
      preview = parts.slice(1).join(" ");
    } else {
      const parsed = parseSenderPrefix(body);
      if (parsed) {
        senderName = parsed.sender;
        preview = parsed.preview;
      }
    }
  }

  if (!senderName && rawText) {
    const parts = rawText.split("\n").map((part) => part.trim()).filter(Boolean);
    senderName = parts[0] || "Teams";
    preview = parts.slice(1).join(" ");
  }

  if (isGenericTeamsSender(senderName) && rawText) {
    const parts = rawText.split("\n").map((part) => part.trim()).filter(Boolean);
    if (parts.length > 2 && isGenericTeamsSender(parts[0])) {
      senderName = parts[1];
      preview = parts.slice(2).join(" ");
    } else {
      const parsed = parseSenderPrefix(rawText);
      if (parsed) {
        senderName = parsed.sender;
        preview = parsed.preview;
      }
    }
  }

  preview = stripLikelyUiNoise(preview || rawText);

  if (!preview) {
    return null;
  }

  return {
    source,
    sender: senderName || "Teams",
    preview,
    pageUrl: settings.includePageUrl ? pageUrl : "",
    observedAt
  };
}

function extractSingleNotificationText(value) {
  const text = cleanMultilineText(value);
  if (!text) {
    return "";
  }

  const lines = text.split("\n").map((part) => part.trim()).filter(Boolean);
  const teamsIndexes = lines
    .map((line, index) => (isGenericTeamsSender(line) ? index : -1))
    .filter((index) => index >= 0);

  if (teamsIndexes.length > 1) {
    return lines.slice(teamsIndexes[teamsIndexes.length - 1]).join("\n");
  }

  return text;
}

function isGenericTeamsSender(value) {
  return /^microsoft teams$/i.test(cleanText(value)) || /^teams$/i.test(cleanText(value));
}

function parseSenderPrefix(value) {
  const text = cleanText(value).replace(/^Microsoft Teams\s+/i, "");
  const commaName = text.match(/^([^,\s][^,]{1,48},\s+[^:,\s]{2,48})\s+(.+)$/u);
  if (commaName) {
    return {
      sender: commaName[1],
      preview: commaName[2]
    };
  }

  const colonName = text.match(/^(.{2,80}?):\s+(.+)$/u);
  if (colonName) {
    return {
      sender: colonName[1],
      preview: colonName[2]
    };
  }

  return null;
}

function cleanText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .replace(/\u00a0/g, " ")
    .trim();
}

function cleanMultilineText(value) {
  return String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function stripLikelyUiNoise(value) {
  return cleanText(value)
    .replace(/^Microsoft Teams\s*/i, "")
    .replace(/\bPress enter to open\b/gi, "")
    .trim();
}

function hashMessage(message) {
  const bucket = Math.floor(new Date(message.observedAt).getTime() / 60000);
  return [
    canonicalPreview(message.preview).toLowerCase(),
    bucket
  ].join("|");
}

function canonicalPreview(preview) {
  return cleanText(preview)
    .replace(/^.+?\s+([A-Z][^:]{1,80}:\s+.+)$/u, "$1");
}

function pruneRecent(now, ttlMs) {
  for (const [hash, timestamp] of recentMessages) {
    if (now - timestamp > ttlMs) {
      recentMessages.delete(hash);
    }
  }
}

function queueTelegramMessage(hash, message, settings) {
  const existing = pendingMessages.get(hash);
  if (existing) {
    if (scoreMessage(message) > scoreMessage(existing.message)) {
      existing.message = message;
      existing.settings = settings;
    }
    return;
  }

  const entry = {
    message,
    settings,
    timer: setTimeout(() => {
      pendingMessages.delete(hash);
      recentMessages.set(hash, Date.now());
      sendTelegramMessage(entry.settings, formatTelegramMessage(entry.message, entry.settings))
        .then(() => appendSentLog(entry.message))
        .catch((error) => console.error("Teams Telegram Notifier:", error));
    }, 900)
  };

  pendingMessages.set(hash, entry);
}

function scoreMessage(message) {
  let score = 0;

  if (!isGenericTeamsSender(message.sender)) {
    score += 10;
  }

  if (message.source === "Teams DOM") {
    score += 3;
  }

  if (!message.preview.toLowerCase().startsWith(message.sender.toLowerCase())) {
    score += 1;
  }

  return score;
}

function formatTelegramMessage(message, settings) {
  const lines = [
    "<b>Teams notification</b>",
    `<b>From:</b> ${escapeHtml(message.sender)}`,
    `<b>Preview:</b> ${escapeHtml(limit(message.preview, 700))}`
  ];

  if (settings.accessUrl) {
    lines.push(`<b>Open:</b> <a href="${escapeAttribute(settings.accessUrl)}">Open Teams</a>`);
  }

  if (message.pageUrl) {
    lines.push(`<b>Page:</b> ${escapeHtml(message.pageUrl)}`);
  }

  return lines.join("\n");
}

function limit(value, maxLength) {
  if (value.length <= maxLength) {
    return value;
  }
  return `${value.slice(0, maxLength - 1)}...`;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeAttribute(value) {
  return escapeHtml(value)
    .replace(/"/g, "&quot;");
}

async function sendTelegramMessage(settings, text) {
  if (!settings.botToken || !settings.chatId) {
    throw new Error("Telegram bot token and chat ID are required.");
  }

  const response = await fetch(`https://api.telegram.org/bot${settings.botToken}/sendMessage`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      chat_id: settings.chatId,
      text,
      parse_mode: "HTML",
      disable_web_page_preview: true
    })
  });

  const result = await response.json().catch(() => ({}));
  if (!response.ok || result.ok === false) {
    throw new Error(result.description || `Telegram request failed with HTTP ${response.status}.`);
  }
}

async function appendSentLog(message) {
  const current = await chrome.storage.local.get(SENT_LOGS_KEY);
  const sentLogs = Array.isArray(current[SENT_LOGS_KEY]) ? current[SENT_LOGS_KEY] : [];
  const nextLogs = [{
    id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    sentAt: new Date().toISOString(),
    source: message.source,
    sender: message.sender,
    preview: message.preview,
    pageUrl: message.pageUrl || ""
  }, ...sentLogs].slice(0, MAX_SENT_LOGS);

  await chrome.storage.local.set({ [SENT_LOGS_KEY]: nextLogs });
}

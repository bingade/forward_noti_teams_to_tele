const DEFAULT_SETTINGS = {
  enabled: false,
  botToken: "",
  chatId: "",
  teamsEnabled: true,
  teamsAccessUrl: "https://teams.microsoft.com/v2/",
  includeTeamsPageUrl: false,
  outlookEnabled: true,
  outlookAccessUrl: "https://outlook.office.com/mail/",
  includeOutlookPageUrl: false,
  dedupeTtlMinutes: 5,
  minTextLength: 8
};

const recentMessages = new Map();
const pendingMessages = new Map();
const SENT_LOGS_KEY = "sentLogs";
const MAX_SENT_LOGS = 100;
const LEGACY_SETTINGS_KEYS = ["accessUrl", "includePageUrl"];
const SETTINGS_KEYS = [...Object.keys(DEFAULT_SETTINGS), ...LEGACY_SETTINGS_KEYS];

chrome.runtime.onInstalled.addListener(async () => {
  const current = await chrome.storage.local.get(SETTINGS_KEYS);
  await chrome.storage.local.set(mergeSettings(current));
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

  if (message.type === "web-notification-candidate" || message.type === "teams-notification-candidate") {
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

  if (!isPlatformEnabled(normalized.platform, settings)) {
    return { ok: true, skipped: "source-disabled" };
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
    platform: "extension",
    source: "Extension test",
    sender: "Microsoft Telegram Notifier",
    preview: "Test message from the Chrome extension.",
    pageUrl: "",
    observedAt: new Date().toISOString()
  };
  await sendTelegramMessage(settings, formatTelegramMessage(message, settings));
  await appendSentLog(message);
  return { ok: true, sent: true };
}

async function loadSettings() {
  const stored = await chrome.storage.local.get(SETTINGS_KEYS);
  return mergeSettings(stored);
}

function compact(value) {
  return Object.fromEntries(
    Object.entries(value || {}).filter(([, item]) => item !== undefined && item !== null)
  );
}

function mergeSettings(stored) {
  const storedSettings = compact(stored);
  const nextSettings = { ...DEFAULT_SETTINGS, ...storedSettings };

  if (storedSettings.accessUrl && !storedSettings.teamsAccessUrl) {
    nextSettings.teamsAccessUrl = storedSettings.accessUrl;
  }

  if (storedSettings.includePageUrl !== undefined && storedSettings.includeTeamsPageUrl === undefined) {
    nextSettings.includeTeamsPageUrl = storedSettings.includePageUrl;
  }

  return normalizeSettings(nextSettings);
}

function normalizeSettings(settings) {
  const nextSettings = { ...settings };

  if (isDeprecatedCallbackUrl(nextSettings.teamsAccessUrl)) {
    nextSettings.teamsAccessUrl = DEFAULT_SETTINGS.teamsAccessUrl;
  }

  if (isDeprecatedCallbackUrl(nextSettings.outlookAccessUrl)) {
    nextSettings.outlookAccessUrl = DEFAULT_SETTINGS.outlookAccessUrl;
  }

  return nextSettings;
}

function isDeprecatedCallbackUrl(value) {
  return /^intent:\/\//i.test(String(value || "")) ||
    /^firefox:\/\//i.test(String(value || "")) ||
    /org\.mozilla\.firefox/i.test(String(value || ""));
}

function normalizePayload(payload, sender, settings) {
  const platform = normalizePlatform(payload && payload.platform, sender);
  if (platform === "outlook") {
    return normalizeOutlookPayload(payload, sender, settings);
  }

  return normalizeTeamsPayload(payload, sender, settings);
}

function normalizeTeamsPayload(payload, sender, settings) {
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
    platform: "teams",
    source,
    sender: senderName || "Teams",
    preview,
    pageUrl: settings.includeTeamsPageUrl ? pageUrl : "",
    observedAt
  };
}

function normalizeOutlookPayload(payload, sender, settings) {
  const rawText = extractSingleOutlookText(cleanMultilineText(payload && payload.text));
  const title = cleanText(payload && payload.title);
  const body = extractSingleOutlookText(cleanMultilineText(payload && payload.body));
  const pageUrl = cleanText(payload && payload.pageUrl) || (sender.tab && sender.tab.url) || "";
  const observedAt = cleanText(payload && payload.observedAt) || new Date().toISOString();
  const source = cleanText(payload && payload.source) || "Outlook";

  let senderName = isGenericOutlookSender(title) ? "" : title;
  let preview = body || rawText;

  if (!senderName && body) {
    const parts = body.split("\n").map((part) => part.trim()).filter(Boolean);
    senderName = stripOutlookNoise(parts[0] || "");
    preview = stripOutlookNoise(parts.slice(1).join(" ") || body);
  }

  if (!senderName && rawText) {
    const parts = rawText.split("\n").map((part) => part.trim()).filter(Boolean);
    const firstMeaningfulLine = parts.find((part) => !isGenericOutlookSender(part));
    senderName = stripOutlookNoise(firstMeaningfulLine || "");
    preview = stripOutlookNoise(parts.filter((part) => part !== firstMeaningfulLine).join(" "));
  }

  const parsed = parseSenderPrefix(preview);
  if ((!senderName || isGenericOutlookSender(senderName)) && parsed) {
    senderName = parsed.sender;
    preview = parsed.preview;
  }

  preview = stripOutlookNoise(preview || rawText);

  if (!preview) {
    return null;
  }

  return {
    platform: "outlook",
    source,
    sender: senderName || "Outlook",
    preview,
    pageUrl: settings.includeOutlookPageUrl ? pageUrl : "",
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

function extractSingleOutlookText(value) {
  const text = cleanMultilineText(value);
  if (!text) {
    return "";
  }

  const lines = text.split("\n").map((part) => part.trim()).filter(Boolean);
  const outlookIndexes = lines
    .map((line, index) => (isGenericOutlookSender(line) ? index : -1))
    .filter((index) => index >= 0);

  if (outlookIndexes.length > 1) {
    return lines.slice(outlookIndexes[outlookIndexes.length - 1]).join("\n");
  }

  return text;
}

function isGenericTeamsSender(value) {
  return /^microsoft teams$/i.test(cleanText(value)) || /^teams$/i.test(cleanText(value));
}

function isGenericOutlookSender(value) {
  const text = cleanText(value);
  return /^microsoft outlook$/i.test(text) ||
    /^outlook$/i.test(text) ||
    /^outlook mail$/i.test(text) ||
    /^new mail$/i.test(text) ||
    /^new email$/i.test(text);
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

function stripOutlookNoise(value) {
  return cleanText(value)
    .replace(/^Microsoft Outlook\s*/i, "")
    .replace(/^Outlook\s*/i, "")
    .replace(/\bPress enter to open\b/gi, "")
    .replace(/\bOpen in Outlook\b/gi, "")
    .trim();
}

function normalizePlatform(value, sender) {
  const platform = cleanText(value).toLowerCase();
  if (platform === "outlook" || platform === "teams") {
    return platform;
  }

  const pageUrl = sender && sender.tab && sender.tab.url ? sender.tab.url : "";
  try {
    const hostname = new URL(pageUrl).hostname.toLowerCase();
    if (hostname.includes("outlook.") || hostname === "mail.live.com" || hostname.endsWith(".mail.live.com")) {
      return "outlook";
    }
  } catch (error) {
    // Ignore malformed sender URLs and fall back to Teams for legacy messages.
  }

  return "teams";
}

function isPlatformEnabled(platform, settings) {
  return platform === "outlook" ? settings.outlookEnabled : settings.teamsEnabled;
}

function hashMessage(message) {
  const bucket = Math.floor(new Date(message.observedAt).getTime() / 60000);
  return [
    message.platform || "teams",
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
        .catch((error) => console.error("Microsoft Telegram Notifier:", error));
    }, 900)
  };

  pendingMessages.set(hash, entry);
}

function scoreMessage(message) {
  let score = 0;

  if (!isGenericTeamsSender(message.sender) && !isGenericOutlookSender(message.sender)) {
    score += 10;
  }

  if (message.source === "Teams DOM" || message.source === "Outlook DOM") {
    score += 3;
  }

  if (!message.preview.toLowerCase().startsWith(message.sender.toLowerCase())) {
    score += 1;
  }

  return score;
}

function formatTelegramMessage(message, settings) {
  const platform = message.platform || "teams";
  const platformName = getPlatformName(platform);
  const lines = [
    `<b>${escapeHtml(platformName)} notification</b>`,
    `<b>From:</b> ${escapeHtml(message.sender)}`,
    `<b>Preview:</b> ${escapeHtml(limit(message.preview, 700))}`
  ];

  const accessUrl = getPlatformAccessUrl(platform, settings);
  if (accessUrl) {
    lines.push(`<b>Open:</b> <a href="${escapeAttribute(accessUrl)}">Open ${escapeHtml(platformName)}</a>`);
  }

  if (message.pageUrl) {
    lines.push(`<b>Page:</b> ${escapeHtml(message.pageUrl)}`);
  }

  return lines.join("\n");
}

function getPlatformName(platform) {
  if (platform === "outlook") {
    return "Outlook";
  }

  if (platform === "extension") {
    return "Extension test";
  }

  return "Teams";
}

function getPlatformAccessUrl(platform, settings) {
  if (platform === "outlook") {
    return settings.outlookAccessUrl;
  }

  if (platform === "teams") {
    return settings.teamsAccessUrl;
  }

  return "";
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
    platform: message.platform || "teams",
    source: message.source,
    sender: message.sender,
    preview: message.preview,
    pageUrl: message.pageUrl || ""
  }, ...sentLogs].slice(0, MAX_SENT_LOGS);

  await chrome.storage.local.set({ [SENT_LOGS_KEY]: nextLogs });
}

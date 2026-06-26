const DEFAULT_SETTINGS = {
  enabled: false,
  botToken: "",
  chatId: "",
  accessUrl: "https://teams.microsoft.com/v2/",
  includePageUrl: false,
  dedupeTtlMinutes: 5,
  minTextLength: 8
};

const form = document.querySelector("#settings-form");
const statusElement = document.querySelector("#status");
const tabButtons = Array.from(document.querySelectorAll("[data-tab-target]"));
const panels = Array.from(document.querySelectorAll("[data-tab-panel]"));
const logSummary = document.querySelector("#log-summary");
const sentLogList = document.querySelector("#sent-log-list");
const refreshLogsButton = document.querySelector("#refresh-logs-button");
const clearLogsButton = document.querySelector("#clear-logs-button");
const fields = Object.fromEntries(
  Object.keys(DEFAULT_SETTINGS).map((key) => [key, document.querySelector(`#${key}`)])
);

document.addEventListener("DOMContentLoaded", restoreSettings);
form.addEventListener("submit", saveSettings);
document.querySelector("#test-button").addEventListener("click", sendTest);
refreshLogsButton.addEventListener("click", loadSentLogs);
clearLogsButton.addEventListener("click", clearSentLogs);
tabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    showTab(button.dataset.tabTarget);
    if (button.dataset.tabTarget === "logs") {
      loadSentLogs();
    }
  });
});

async function restoreSettings() {
  const stored = await chrome.storage.local.get(Object.keys(DEFAULT_SETTINGS));
  const settings = normalizeSettings({ ...DEFAULT_SETTINGS, ...stored });

  for (const [key, field] of Object.entries(fields)) {
    if (!field) {
      continue;
    }

    if (field.type === "checkbox") {
      field.checked = Boolean(settings[key]);
    } else {
      field.value = settings[key] ?? "";
    }
  }

  await loadSentLogs();
}

async function saveSettings(event) {
  event.preventDefault();
  await chrome.storage.local.set(readSettingsFromForm());
  showStatus("Saved.");
}

async function sendTest() {
  await chrome.storage.local.set(readSettingsFromForm());
  showStatus("Sending test...");

  chrome.runtime.sendMessage({ type: "test-telegram" }, (response) => {
    const error = chrome.runtime.lastError || (response && response.error);
    if (error) {
      showStatus(typeof error === "string" ? error : error.message, true);
      return;
    }

    showStatus("Test message sent.");
    loadSentLogs();
  });
}

function readSettingsFromForm() {
  return normalizeSettings({
    enabled: fields.enabled.checked,
    botToken: fields.botToken.value.trim(),
    chatId: fields.chatId.value.trim(),
    accessUrl: fields.accessUrl.value.trim(),
    includePageUrl: fields.includePageUrl.checked,
    dedupeTtlMinutes: clampNumber(fields.dedupeTtlMinutes.value, 1, 120, 5),
    minTextLength: clampNumber(fields.minTextLength.value, 1, 200, 8)
  });
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

function clampNumber(value, min, max, defaultValue) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return defaultValue;
  }

  return Math.min(max, Math.max(min, number));
}

function showStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.classList.toggle("error", isError);
}

function showTab(target) {
  tabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.tabTarget === target);
  });

  panels.forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.tabPanel === target);
  });
}

async function loadSentLogs() {
  const stored = await chrome.storage.local.get("sentLogs");
  const logs = Array.isArray(stored.sentLogs) ? stored.sentLogs : [];
  logSummary.textContent = logs.length ? `${logs.length} sent message${logs.length === 1 ? "" : "s"}.` : "No sent messages.";
  sentLogList.replaceChildren(...logs.map(renderLogItem));
}

function renderLogItem(log) {
  const item = document.createElement("li");
  item.className = "log-item";

  const meta = document.createElement("div");
  meta.className = "log-meta";
  meta.textContent = `${formatDate(log.sentAt)} - ${log.source || "Teams"}`;

  const sender = document.createElement("strong");
  sender.textContent = log.sender || "Teams";

  const preview = document.createElement("p");
  preview.textContent = log.preview || "";

  item.append(meta, sender, preview);

  if (log.pageUrl) {
    const link = document.createElement("a");
    link.href = log.pageUrl;
    link.textContent = log.pageUrl;
    link.target = "_blank";
    link.rel = "noreferrer";
    item.append(link);
  }

  return item;
}

function formatDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "Unknown time";
  }

  return date.toLocaleString();
}

function clearSentLogs() {
  chrome.runtime.sendMessage({ type: "clear-sent-logs" }, (response) => {
    const error = chrome.runtime.lastError || (response && response.error);
    if (error) {
      showStatus(typeof error === "string" ? error : error.message, true);
      return;
    }

    loadSentLogs();
  });
}

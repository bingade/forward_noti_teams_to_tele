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

const LEGACY_SETTINGS_KEYS = ["accessUrl", "includePageUrl"];
const SETTINGS_KEYS = [...Object.keys(DEFAULT_SETTINGS), ...LEGACY_SETTINGS_KEYS];

const forms = Array.from(document.querySelectorAll("[data-settings-form]"));
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
forms.forEach((form) => form.addEventListener("submit", saveSettings));
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
  const stored = await chrome.storage.local.get(SETTINGS_KEYS);
  const settings = mergeSettings(stored);

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
  showStatus("Saved.", false, event.currentTarget);
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
    teamsEnabled: fields.teamsEnabled.checked,
    teamsAccessUrl: fields.teamsAccessUrl.value.trim(),
    includeTeamsPageUrl: fields.includeTeamsPageUrl.checked,
    outlookEnabled: fields.outlookEnabled.checked,
    outlookAccessUrl: fields.outlookAccessUrl.value.trim(),
    includeOutlookPageUrl: fields.includeOutlookPageUrl.checked,
    dedupeTtlMinutes: clampNumber(fields.dedupeTtlMinutes.value, 1, 120, 5),
    minTextLength: clampNumber(fields.minTextLength.value, 1, 200, 8)
  });
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

function compact(value) {
  return Object.fromEntries(
    Object.entries(value || {}).filter(([, item]) => item !== undefined && item !== null)
  );
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

function clampNumber(value, min, max, defaultValue) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return defaultValue;
  }

  return Math.min(max, Math.max(min, number));
}

function showStatus(message, isError = false, scope = document.querySelector(".panel.active")) {
  const statusElement = scope.querySelector("[data-status]");
  if (!statusElement) {
    return;
  }

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

  showStatus("");
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
  meta.textContent = `${formatDate(log.sentAt)} - ${log.source || getPlatformName(log.platform)}`;

  const sender = document.createElement("strong");
  sender.textContent = log.sender || getPlatformName(log.platform);

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
    showStatus("Logs cleared.");
  });
}

function getPlatformName(platform) {
  return platform === "outlook" ? "Outlook" : "Teams";
}

const SCAN_DELAY_MS = 250;
const SEEN_TTL_MS = 5 * 60 * 1000;
const MAX_SEEN = 300;
const INJECTED_SCRIPT = "src/injected-notification-hook.js";
const PLATFORM = detectPlatform();

const seen = new Map();
let scanTimer = null;
let observer = null;
let stopped = false;

installRuntimeErrorHandlers();
installNotificationApiListener();
injectNotificationHook();

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", startDomObserver, { once: true });
} else {
  startDomObserver();
}

function installNotificationApiListener() {
  window.addEventListener("message", (event) => {
    if (stopped || event.source !== window || !event.data || event.data.type !== "web-notification-api") {
      return;
    }

    sendCandidate({
      source: "Notification API",
      platform: PLATFORM,
      title: event.data.title,
      body: event.data.body,
      text: [event.data.title, event.data.body].filter(Boolean).join("\n"),
      pageUrl: location.href,
      observedAt: new Date().toISOString()
    });
  });
}

function injectNotificationHook() {
  if (!isExtensionContextAvailable() || !chrome.runtime.getURL) {
    return;
  }

  try {
    const script = document.createElement("script");
    script.src = chrome.runtime.getURL(INJECTED_SCRIPT);
    script.async = false;
    script.onload = () => script.remove();
    (document.documentElement || document.head || document).appendChild(script);
  } catch (error) {
    if (handleRuntimeError(error)) {
      return;
    }

    throw error;
  }
}

function startDomObserver() {
  if (stopped) {
    return;
  }

  try {
    queueScan();

    observer = new MutationObserver((mutations) => {
      if (stopped) {
        return;
      }

      try {
        if (mutations.some((mutation) => mutation.addedNodes.length || mutation.type === "attributes")) {
          queueScan();
        }
      } catch (error) {
        handleRuntimeError(error);
      }
    });

    observer.observe(document.documentElement, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["class", "style", "aria-hidden", "hidden"]
    });
  } catch (error) {
    if (handleRuntimeError(error)) {
      return;
    }

    throw error;
  }
}

function queueScan() {
  if (stopped || scanTimer) {
    return;
  }

  scanTimer = setTimeout(() => {
    scanTimer = null;
    try {
      scanVisibleNotificationCandidates();
    } catch (error) {
      if (handleRuntimeError(error)) {
        return;
      }

      throw error;
    }
  }, SCAN_DELAY_MS);
}

function scanVisibleNotificationCandidates() {
  if (stopped) {
    return;
  }

  pruneSeen();

  for (const element of findCandidateElements()) {
    if (!isVisible(element)) {
      continue;
    }

    const text = normalizeText(element.innerText || element.textContent || "");
    if (!looksLikePlatformNotification(element, text)) {
      continue;
    }

    const key = makeSeenKey(text);
    if (seen.has(key)) {
      continue;
    }

    seen.set(key, Date.now());
    sendCandidate({
      source: `${getPlatformLabel()} DOM`,
      platform: PLATFORM,
      text,
      pageUrl: location.href,
      observedAt: new Date().toISOString()
    });
  }
}

function findCandidateElements() {
  const selectors = [
    "[role='alert']",
    "[role='status']",
    "[aria-live='assertive']",
    "[aria-live='polite']",
    "[data-tid*='toast' i]",
    "[data-tid*='notification' i]",
    "[class*='toast' i]",
    "[class*='notification' i]",
    "[class*='snackbar' i]"
  ];

  const selector = selectors.join(",");
  try {
    return Array.from(document.querySelectorAll(selector));
  } catch (error) {
    if (handleRuntimeError(error)) {
      return [];
    }

    throw error;
  }
}

function looksLikePlatformNotification(element, text) {
  if (!text || text.length < 8 || text.length > 1200) {
    return false;
  }

  const lowerText = text.toLowerCase();
  const teamsMarker = [
    "microsoft teams",
    "sent",
    "message",
    "chat",
    "mentioned",
    "reply",
    "missed",
    "call"
  ].some((word) => lowerText.includes(word));

  const outlookMarker = [
    "microsoft outlook",
    "outlook",
    "new mail",
    "new email",
    "email",
    "inbox",
    "calendar",
    "meeting",
    "reminder"
  ].some((word) => lowerText.includes(word));

  const elementHint = [
    element.getAttribute("role"),
    element.getAttribute("aria-live"),
    element.getAttribute("data-tid"),
    element.className
  ].join(" ").toLowerCase();

  const hasContainerHint = /toast|notification|alert|status|snackbar/.test(elementHint);
  const marker = PLATFORM === "outlook" ? outlookMarker : teamsMarker;
  return marker || hasContainerHint;
}

function isVisible(element) {
  try {
    if (!element || element.closest("[aria-hidden='true'],[hidden]")) {
      return false;
    }

    const style = window.getComputedStyle(element);
    if (style.display === "none" || style.visibility === "hidden" || Number(style.opacity) === 0) {
      return false;
    }

    const rect = element.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  } catch (error) {
    if (handleRuntimeError(error)) {
      return false;
    }

    throw error;
  }
}

function sendCandidate(payload) {
  if (!isExtensionContextAvailable()) {
    stopContentScript();
    return;
  }

  try {
    chrome.runtime.sendMessage({
      type: "web-notification-candidate",
      payload
    });
  } catch (error) {
    if (handleRuntimeError(error)) {
      return;
    }

    throw error;
  }
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function makeSeenKey(text) {
  return normalizeText(text).toLowerCase();
}

function pruneSeen() {
  const now = Date.now();
  for (const [key, timestamp] of seen) {
    if (now - timestamp > SEEN_TTL_MS) {
      seen.delete(key);
    }
  }

  while (seen.size > MAX_SEEN) {
    const firstKey = seen.keys().next().value;
    seen.delete(firstKey);
  }
}

function detectPlatform() {
  const hostname = location.hostname.toLowerCase();
  if (hostname.includes("outlook.") || hostname === "mail.live.com" || hostname.endsWith(".mail.live.com")) {
    return "outlook";
  }

  return "teams";
}

function getPlatformLabel() {
  return PLATFORM === "outlook" ? "Outlook" : "Teams";
}

function isExtensionContextAvailable() {
  try {
    return typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.id;
  } catch (error) {
    handleRuntimeError(error);
    return false;
  }
}

function isExtensionContextInvalidated(error) {
  return error && /extension context invalidated/i.test(String(error.message || error));
}

function handleRuntimeError(error) {
  if (!isExtensionContextInvalidated(error)) {
    return false;
  }

  stopContentScript();
  return true;
}

function stopContentScript() {
  stopped = true;

  if (scanTimer) {
    clearTimeout(scanTimer);
    scanTimer = null;
  }

  if (observer) {
    observer.disconnect();
    observer = null;
  }
}

function installRuntimeErrorHandlers() {
  window.addEventListener("error", (event) => {
    if (handleRuntimeError(event.error || event.message)) {
      event.preventDefault();
    }
  }, true);

  window.addEventListener("unhandledrejection", (event) => {
    if (handleRuntimeError(event.reason)) {
      event.preventDefault();
    }
  }, true);
}

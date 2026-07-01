(function installMicrosoftTelegramNotificationHook() {
  if (window.__microsoftTelegramNotificationHookInstalled) {
    return;
  }

  window.__microsoftTelegramNotificationHookInstalled = true;

  const NativeNotification = window.Notification;
  if (typeof NativeNotification !== "function") {
    return;
  }

  function HookedNotification(title, options) {
    const body = options && typeof options.body === "string" ? options.body : "";
    window.postMessage({
      type: "web-notification-api",
      title: String(title || ""),
      body
    }, window.location.origin);

    return new NativeNotification(title, options);
  }

  HookedNotification.prototype = NativeNotification.prototype;
  Object.setPrototypeOf(HookedNotification, NativeNotification);

  Object.defineProperty(window, "Notification", {
    configurable: true,
    enumerable: true,
    writable: true,
    value: HookedNotification
  });
})();

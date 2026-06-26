# Teams Telegram Notifier

Chrome extension for forwarding Microsoft Teams web notification previews to Telegram.

## Install

1. Open `chrome://extensions`.
2. Enable Developer mode.
3. Select Load unpacked.
4. Choose this folder: `extension/teams-telegram-notifier`.
5. Open the extension options page and enter:
   - Telegram bot token
   - Telegram chat ID
   - open callback URL, such as `intent://teams.microsoft.com/#Intent;scheme=https;package=org.mozilla.firefox;end`
   - fallback Teams URL, such as `https://teams.microsoft.com/`
6. Click Save, then Send test.

The callback URL is sent as a Telegram HTML link labeled `Open in Firefox`.
The fallback URL is sent as plain text because some Telegram clients do not open custom schemes or Android intent links consistently.
The options page also has a Sent logs tab with the latest 100 successfully sent Telegram messages.

## Test

```sh
node --check src/content.js
node --check src/background.js
node --check options.js
```

## How It Works

- `src/content.js` runs on Teams web pages.
- `src/injected-notification-hook.js` watches Teams web `Notification` calls.
- The content script also broadly scans visible DOM nodes that look like Teams toast or alert notifications, then the background parser deduplicates noisy parent and child candidates.
- `src/background.js` deduplicates candidates and sends a Telegram `sendMessage` request.

## Security Note

This extension stores the Telegram bot token in the Chrome profile via `chrome.storage.local`.
That is convenient for copying the extension to another browser machine, but it is not as strong as keeping the token in a VM-side service.
Treat the browser profile and VM backups as sensitive.

## Current Limitations

- Teams DOM classes and notification layout can change.
- The DOM scanner is intentionally heuristic and may need selector tuning after observing real Teams notifications.
- Browser Notification API capture can emit generic `Microsoft Teams` titles, so the background parser tries to recover sender names from the body.
- If Teams creates notifications before the extension is installed or before the tab reloads, those old notifications will not be captured.

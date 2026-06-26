# Teams Telegram Notifier

Chrome extension for forwarding Microsoft Teams web notification previews to Telegram.

## Prerequisites

1. Open Chrome and sign in to [Teams Web](https://teams.microsoft.com/v2/).
2. Keep one Teams Web tab logged in.
3. Allow Teams browser notifications when Chrome asks.

The extension only watches Teams Web pages that are open in Chrome.

Default Teams callback URL:

```text
https://teams.microsoft.com/v2/
```

## Create a Telegram Bot

1. Open Telegram and search for `@BotFather`.
2. Send `/newbot`.
3. Follow the prompts to choose a bot name and username.
4. Copy the bot token that BotFather returns.
5. Open Telegram and search for `@userinfobot`.
6. Start `@userinfobot`.
7. Copy the ID returned by `@userinfobot`. That number is your Telegram chat ID.

For a private chat, the chat ID usually looks like a normal number, for example `6091378862`.
For a group or channel, it may be negative, for example `-1001234567890`.

## Install

1. Open `chrome://extensions`.
2. Enable Developer mode.
3. Select Load unpacked.
4. Choose this repository folder, the folder that contains `manifest.json`.
5. Open the extension options page and enter:
   - Telegram bot token
   - Telegram chat ID
   - Teams callback URL, default `https://teams.microsoft.com/v2/`
6. Click Save, then Send test.
7. Reload the Teams Web tab after installing or updating the extension.

The Teams callback URL is sent as a Telegram HTML link labeled `Open Teams`.
The options page also has a Sent logs tab with the latest 100 successfully sent Telegram messages.

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

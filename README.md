# Microsoft Telegram Notifier

Chrome extension for forwarding Microsoft Teams and Outlook web notification previews to Telegram.

## Prerequisites

1. Open Chrome and sign in to [Teams Web](https://teams.microsoft.com/v2/).
2. Open Chrome and sign in to [Outlook Web](https://outlook.office.com/mail/) if you want Outlook alerts.
3. Keep the Teams or Outlook Web tab logged in.
4. Allow browser notifications when Chrome asks.

The extension only watches Teams and Outlook Web pages that are open in Chrome.

Default Teams callback URL:

```text
https://teams.microsoft.com/v2/
```

Default Outlook callback URL:

```text
https://outlook.office.com/mail/
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
   - Teams callback URL in the Teams tab, default `https://teams.microsoft.com/v2/`
   - Outlook callback URL in the Outlook tab, default `https://outlook.office.com/mail/`
6. Click Save, then Send test.
7. Reload the Teams and Outlook Web tabs after installing or updating the extension.

The Teams and Outlook callback URLs are sent as Telegram HTML links labeled `Open Teams` or `Open Outlook`.
The options page also has a Sent logs tab with the latest 100 successfully sent Telegram messages.

## How It Works

- `src/content.js` runs on Teams and Outlook web pages.
- `src/injected-notification-hook.js` watches web `Notification` calls from those pages.
- The content script also broadly scans visible DOM nodes that look like Teams or Outlook toast or alert notifications, then the background parser deduplicates noisy parent and child candidates.
- `src/background.js` deduplicates candidates and sends a Telegram `sendMessage` request.

## Security Note

This extension stores the Telegram bot token in the Chrome profile via `chrome.storage.local`.
That is convenient for copying the extension to another browser machine, but it is not as strong as keeping the token in a VM-side service.
Treat the browser profile and VM backups as sensitive.

## Current Limitations

- Teams and Outlook DOM classes and notification layouts can change.
- The DOM scanner is intentionally heuristic and may need selector tuning after observing real Teams or Outlook notifications.
- Browser Notification API capture can emit generic app titles, so the background parser tries to recover sender names from the body.
- If Teams or Outlook creates notifications before the extension is installed or before the tab reloads, those old notifications will not be captured.

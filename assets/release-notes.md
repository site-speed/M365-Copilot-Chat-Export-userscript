# M365 Copilot Chat Conversation Exporter v1.0.43

Supports the redirected Microsoft 365 Copilot Chat web address, `copilot.cloud.microsoft`.

## What it does

Exports the current Microsoft 365 work/school Copilot Chat conversation as readable Markdown and a raw JSON companion Markdown file.

## Improvements

- Runs on `copilot.cloud.microsoft` after the Microsoft 365 Copilot web URL change.
- Keeps running on `m365.cloud.microsoft`, `microsoft365.com`, and `www.microsoft365.com` for tenants that have not been redirected.
- Treats only URLs that carry a real conversation identifier as an exportable chat, so unrelated query parameters no longer select the wrong conversation.
- Detects the conversation download call when it is served from the app's own `*.cloud.microsoft` host.
- Personal Copilot (`copilot.microsoft.com`) remains outside this tool's scope.

## Install

Install from GreasyFork:

```text
https://greasyfork.org/en/scripts/577806-m365-copilot-chat-conversation-exporter
```

## Source and support

Source and support are available from the public userscript repository:

```text
https://github.com/site-speed/M365-Copilot-Chat-Export-userscript
```

## Notes

- Supports the Microsoft 365 Copilot web URL redirect from the August–November 2026 rollout.
- Confirm the update in Tampermonkey or Violentmonkey, then refresh the Copilot Chat tab.
- Keeps opaque sync, telemetry, request, and service-envelope fields JSON-only.
- Exported files may contain sensitive work data and should be handled carefully.

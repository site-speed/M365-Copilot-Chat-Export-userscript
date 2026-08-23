# M365 Copilot Chat Conversation Exporter v1.0.42

Repairs the observed Turn 38 Markdown emphasis, heading, list, and code-fence corruption.

## What it does

Exports the current Microsoft 365 work/school Copilot Chat conversation as readable Markdown and a raw JSON companion Markdown file.

## Improvements

- Preserves original blank lines during duplicate-link cleanup.
- Prevents cross-paragraph bold-marker matching.
- Leaves ordinary `text` code blocks independent and balanced.
- Limits split-heading continuation repair to numbered/lettered `1)` / `A)` headings.
- Adds a focused regression test reproducing the exact issue #7 Turn 38 structure.
- Retains the v1.0.41 curated created/updated time, raw record count, model tone, plugin, and unusual turn-state metadata.

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

- Keeps opaque sync, telemetry, request, and service-envelope fields JSON-only.
- Exported files may contain sensitive work data and should be handled carefully.

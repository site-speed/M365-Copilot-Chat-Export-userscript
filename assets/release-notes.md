# M365 Copilot Chat Conversation Exporter v1.0.27

Documentation and release-metadata sync update for the userscript.

## What it does

Exports the current Microsoft 365 work/school Copilot Chat conversation from:

```text
https://m365.cloud.microsoft/chat
```

as:

- readable Markdown
- raw JSON companion Markdown

The readable Markdown is designed for human review, project handoff, and future conversation rehydration. The raw JSON companion remains the complete local backup.

## Install

Install from GreasyFork:

```text
https://greasyfork.org/en/scripts/577806-m365-copilot-chat-conversation-exporter
```

## Source and support

Source:

```text
https://github.com/site-speed/M365-Copilot-Chat-Export-userscript
```

Issues:

```text
https://github.com/site-speed/M365-Copilot-Chat-Export-userscript/issues
```

## Notes

- Fixes public README version metadata so release checks recognise the current userscript version.
- Keeps the v1.0.26 Substrate URL hardening: passive network hooks use URL parsing and exact `substrate.office.com` hostname checks.
- No export-format or runtime behaviour changes beyond version/metadata alignment.
- Exported files may contain sensitive work data and should be handled carefully.

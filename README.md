# M365 Copilot Chat Conversation Exporter — Userscript

![Greasy Fork Downloads](https://img.shields.io/greasyfork/dt/577806)
![Greasy Fork Rating](https://img.shields.io/greasyfork/rating-count/577806)

[![GitHub Release](https://img.shields.io/github/v/release/site-speed/M365-Copilot-Chat-Export-userscript?style=flat&color=blue)](https://github.com/site-speed/M365-Copilot-Chat-Export-userscript/releases/latest)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
![GitHub Repo stars](https://img.shields.io/github/stars/site-speed/M365-Copilot-Chat-Export-userscript)
![GitHub Downloads (all assets, all releases)](https://img.shields.io/github/downloads/site-speed/M365-Copilot-Chat-Export-userscript/total)
<!-- start user badges -->
![PRs opened in last 30 days](https://img.shields.io/badge/PRs%20opened%20in%20last%2030%20days-0-green?labelColor=555) ![PRs closed in last 30 days](https://img.shields.io/badge/PRs%20closed%20in%20last%2030%20days-0-red?labelColor=555) ![Open PRs](https://img.shields.io/badge/Open%20PRs-0-blue?labelColor=555)

![Issues opened in last 30 days](https://img.shields.io/badge/Issues%20opened%20in%20last%2030%20days-0-green?labelColor=555) ![Issues closed in last 30 days](https://img.shields.io/badge/Issues%20closed%20in%20last%2030%20days-0-red?labelColor=555) ![Open issues](https://img.shields.io/badge/Open%20issues-0-blue?labelColor=555)

![Lines added (last 30 days)](https://img.shields.io/badge/Lines%20added%20(last%2030%20days)-121-green?labelColor=555) ![Lines deleted (last 30 days)](https://img.shields.io/badge/Lines%20deleted%20(last%2030%20days)-40-red?labelColor=555) ![Commits in last 30 days](https://img.shields.io/badge/Commits%20in%20last%2030%20days-33-blue?labelColor=555)

![Contributors (unique)](https://img.shields.io/badge/Contributors%20(unique)-2-blue?labelColor=555) ![Active contributors (last 30d)](https://img.shields.io/badge/Active%20contributors%20(last%2030d)-2-blue?labelColor=555)
<!-- end user badges -->

Version: **v1.0.40**

Export the current Microsoft 365 Copilot Chat conversation to readable Markdown and raw JSON Markdown files.

## What it does

This userscript adds a small exporter panel to Microsoft 365 Copilot Chat pages. From an open conversation, it can export:

1. a readable Markdown file (`.md`) for review, search, and handoff;
2. a raw JSON Markdown companion (`.json.md`) as the complete local backup.

The readable Markdown is designed to be compact and useful to humans. The raw JSON companion is the most complete record and should be kept with the Markdown export.

## Screenshot

![M365 Copilot Chat Conversation Exporter userscript panel](assets/screenshot-ui.png)

## Install

Install from GreasyFork:

```text
https://greasyfork.org/en/scripts/577806-m365-copilot-chat-conversation-exporter
```

Manual installation is also possible from the public GitHub userscript file:

```text
m365-copilot-export.js
```

Use a userscript manager such as Tampermonkey or Violentmonkey and import the raw userscript file.

## Supported pages

The userscript targets Microsoft 365 Copilot Chat web conversations, including common Microsoft 365 chat URL variants such as:

```text
https://m365.cloud.microsoft/chat*
https://m365.cloud.microsoft/*/chat*
https://microsoft365.com/chat*
https://www.microsoft365.com/chat*
```

It is intended for Microsoft 365 work or school Copilot Chat sessions, not personal Copilot chats.

## Usage

1. Open a Microsoft 365 Copilot Chat conversation in the browser.
2. Wait for the exporter panel to detect the current chat.
3. Choose whether to include unclassified records.
4. Click the export button.
5. Keep the generated `.md` and `.json.md` files together.

## Export contents

Readable Markdown may include:

- user prompts and Copilot responses;
- uploaded filenames when detected;
- reasoning/process summaries where available;
- tool-run/code execution details;
- search/source provenance;
- citations and links;
- selected plugin/tool provenance when useful for understanding the conversation.

Raw JSON Markdown includes the full conversation JSON payload wrapped in a Markdown file for easier storage and opening.

## Source and support

Source:

```text
https://github.com/site-speed/M365-Copilot-Chat-Export-userscript
```

Issues:

```text
https://github.com/site-speed/M365-Copilot-Chat-Export-userscript/issues
```

## Privacy and data handling

Exports are generated from your authenticated browser session and may contain sensitive organisation data, prompts, responses, citations, file names, and tool traces.

Treat exported `.md` and `.json.md` files as private unless reviewed and deliberately shared.

## Limitations

- Microsoft 365 Copilot Chat APIs and page structure can change.
- The readable Markdown export is curated for usefulness, not a byte-for-byte mirror of every JSON field.
- The raw JSON companion is the best fallback if a future renderer misses a detail.

## Release notes

Current release notes are available at:

```text
assets/release-notes.md
```

## Security

See `SECURITY.md` for supported-version and reporting guidance.

## Acknowledgements

Thanks to the following MIT-licensed userscript projects and authors that helped inform this work:

- [ingo/m365-copilot-chat-exporter](https://github.com/ingo/m365-copilot-chat-exporter) by Ingo Muschenetz.
- [ganyuke/copilot-exporter](https://github.com/ganyuke/copilot-exporter) by ganyuke.
- [NoahTheGinger/Userscripts](https://github.com/NoahTheGinger/Userscripts) by NoahTheGinger.

This project is MIT licensed.

## Licence

MIT License. Copyright 2026 Tim Moss.

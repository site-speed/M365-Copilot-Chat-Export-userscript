# M365 Copilot Chat Conversation Exporter

A Tampermonkey/Violentmonkey userscript for exporting **M365 Copilot Chat** conversations from the Microsoft 365 work/school chat experience.

Version: **1.0.5**  
Public repo: https://github.com/site-speed/M365-Copilot-Chat-Export-userscript

This tool targets the Microsoft 365 business/work/school version of Copilot Chat at:

```text
https://m365.cloud.microsoft/chat
```

It is **not** intended primarily for the personal Microsoft Copilot Chat experience. Personal Copilot Chat uses a different product surface and URL pattern, and this exporter may or may not work there.

## UI preview

![M365 Copilot Chat Conversation Exporter UI](assets/screenshot-ui.png)

The screenshot is redacted and shows the floating exporter panel in the M365 Copilot Chat web UI.

## What it exports

The exporter downloads the current M365 Copilot Chat conversation as readable Markdown (`.md`) and raw JSON companion Markdown (`.json.md`).

## Requirements

- A Microsoft 365 work or school account with access to M365 Copilot Chat.
- An active browser session at `https://m365.cloud.microsoft/chat`.
- A userscript manager such as Tampermonkey or Violentmonkey.

## Install

### GreasyFork

Install from GreasyFork:

```text
https://greasyfork.org/en/scripts/577806-m365-copilot-chat-conversation-exporter
```


### GreasyFork

GreasyFork is the intended primary install surface for the first public release. Once the GreasyFork listing is live, install from that listing.

### GitHub raw URL

The source userscript URL for GreasyFork import/sync is:

```text
https://raw.githubusercontent.com/site-speed/M365-Copilot-Chat-Export-userscript/main/m365-copilot-export.js
```

## Usage

1. Navigate to `https://m365.cloud.microsoft/chat`.
2. Open the M365 Copilot Chat conversation you want to preserve.
3. Click the exporter panel button.
4. Download the readable Markdown and raw JSON companion files.

## Privacy and security

The script runs locally in your browser and uses your existing authenticated M365 Copilot Chat browser session. Exported files may contain sensitive work data, prompts, generated responses, file names, sources, tool traces, and other conversation metadata. Store exported files carefully and follow your organisation's data-handling rules.

## Limitations

- This exporter targets M365 Copilot Chat, not personal Copilot Chat.
- The M365 Copilot Chat web UI and underlying APIs may change without notice.
- The raw JSON companion is the most complete backup; readable Markdown is intentionally curated.

## Acknowledgements

Thanks to the following MIT-licensed userscript projects and authors that helped inform this work:

- [ingo/m365-copilot-chat-exporter](https://github.com/ingo/m365-copilot-chat-exporter) by Ingo Muschenetz.
- [ganyuke/copilot-exporter](https://github.com/ganyuke/copilot-exporter) by ganyuke.
- [NoahTheGinger/Userscripts](https://github.com/NoahTheGinger/Userscripts) by NoahTheGinger.

This project is MIT licensed.

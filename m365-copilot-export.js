// ==UserScript==
// @name         M365 Copilot Chat Conversation Exporter
// @namespace    https://github.com/site-speed/M365-Copilot-Chat-Export-userscript
// @version 1.0.37
// @description  Export the current Microsoft 365 Copilot Chat conversation to readable Markdown and raw JSON Markdown files.
// @author       Tim Moss
// @license      MIT
// @homepageURL  https://github.com/site-speed/M365-Copilot-Chat-Export-userscript
// @supportURL   https://github.com/site-speed/M365-Copilot-Chat-Export-userscript/issues
// @match        https://m365.cloud.microsoft/chat*
// @match        https://m365.cloud.microsoft/*/chat*
// @match        https://microsoft365.com/chat*
// @match        https://www.microsoft365.com/chat*
// @run-at       document-end
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const SCRIPT_VERSION = "1.0.37";
  const SETTINGS_KEY = "m365ce_export_settings_v9";

  // --------------------
  // Renderer invariants (v1)
  // --------------------
  const ADD_RULE_BETWEEN_BLOCKS = true;
  const COLLAPSE_DUP_STATUS = true;

  // Unclassified-record export evidence is exposed as one user-facing master switch.
  const DEFAULTS = {
    includeUnclassifiedRecords: true,
  };

  function loadSettings() {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      if (!raw) return { ...DEFAULTS };
      const parsed = JSON.parse(raw);
      return { ...DEFAULTS, ...(parsed || {}) };
    } catch {
      return { ...DEFAULTS };
    }
  }

  function saveSettings(s) {
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
    } catch {
      // ignore
    }
  }

  let settings = loadSettings();

  function currentUnclassifiedRecordSetting() {
    return settings.includeUnclassifiedRecords !== false;
  }

  function applyUnclassifiedRecordSetting(enabled) {
    settings.includeUnclassifiedRecords = !!enabled;
    saveSettings(settings);
  }

  function setExportOptions(options = {}) {
    if (Object.prototype.hasOwnProperty.call(options, "includeUnclassifiedRecords")) {
      applyUnclassifiedRecordSetting(options.includeUnclassifiedRecords);
    }
    return { includeUnclassifiedRecords: settings.includeUnclassifiedRecords };
  }

  // --------------------
  // Config
  // --------------------
  const SUBSTRATE_BASE = "https://substrate.office.com/m365Copilot";

  let lastConversation = null;
  let lastConversationId = null;
  const chatTitleCache = new Map();
  let currentChatResolveTimer = null;
  let currentChatResolveInFlightId = null;
  let currentStatusConversationId = null;
  const exportedConversationIds = new Set();

  // --------------------
  // Current chat tracking / UI
  // --------------------

  function cacheConversationSummary(conv) {
    if (!conv?.conversationId) {
      return;
    }
    if (conv.chatName) chatTitleCache.set(conv.conversationId, conv.chatName);
  }

  function setCurrentChatInfo(name, conversationId, state = "") {
    const el = document.getElementById("m365ce-current-chat");
    if (!el) {
      return;
    }
    el.textContent = "";
    const titleLine = document.createElement("div");
    titleLine.textContent = name ? `Chat: ${name}` : "Chat: (resolving...)";
    titleLine.style.cssText =
      "color:#8be9fd;font-size:11px;line-height:1.35;white-space:normal;overflow-wrap:anywhere;";
    el.appendChild(titleLine);
    if (conversationId) {
      const idLine = document.createElement("div");
      idLine.textContent = `ID: ${conversationId}${state ? ` · ${state}` : ""}`;
      idLine.style.cssText =
        "color:#a8b3c7;font-size:10px;line-height:1.35;font-family:ui-monospace,SFMono-Regular,Consolas,monospace;white-space:normal;overflow-wrap:anywhere;";
      el.appendChild(idLine);
    } else if (state) {
      const stateLine = document.createElement("div");
      stateLine.textContent = state;
      stateLine.style.cssText =
        "color:#a8b3c7;font-size:10px;line-height:1.35;";
      el.appendChild(stateLine);
    }
  }

  function updateCurrentChatInfo() {
    const currentId = inferConversationIdFromUrl();
    syncExportStatusForConversation(currentId);
    if (
      lastConversation?.chatName &&
      (!currentId || lastConversation.conversationId === currentId)
    ) {
      setCurrentChatInfo(
        lastConversation.chatName,
        lastConversation.conversationId,
      );
      return;
    }
    if (currentId && chatTitleCache.has(currentId)) {
      setCurrentChatInfo(chatTitleCache.get(currentId), currentId, "cached");
      return;
    }
    if (currentId) {
      const state =
        currentChatResolveInFlightId === currentId
          ? "resolving title..."
          : "selected in page";
      setCurrentChatInfo("", currentId, state);
      return;
    }
    setCurrentChatInfo("", null, "no conversation detected");
  }

  async function resolveCurrentChatTitleIfNeeded() {
    const currentId = inferConversationIdFromUrl();
    if (!currentId) {
      updateCurrentChatInfo();
      return;
    }
    if (
      lastConversation?.conversationId === currentId &&
      lastConversation?.chatName
    ) {
      cacheConversationSummary(lastConversation);
      updateCurrentChatInfo();
      return;
    }
    if (chatTitleCache.has(currentId)) {
      updateCurrentChatInfo();
      return;
    }
    if (currentChatResolveInFlightId === currentId) {
      updateCurrentChatInfo();
      return;
    }
    currentChatResolveInFlightId = currentId;
    updateCurrentChatInfo();
    try {
      const auth = await getTokenAndIds();
      const conv = await substrateGetConversation(auth, currentId);
      cacheConversationSummary(conv);
    } catch (e) {
      console.warn("[M365 Export] title prefetch failed", e);
    } finally {
      if (currentChatResolveInFlightId === currentId) {
        currentChatResolveInFlightId = null;
      }
      updateCurrentChatInfo();
    }
  }

  function scheduleCurrentChatResolution(delay = 450) {
    const currentId = inferConversationIdFromUrl();
    if (!currentId) {
      updateCurrentChatInfo();
      return;
    }
    if (
      (lastConversation?.conversationId === currentId &&
        lastConversation?.chatName) ||
      chatTitleCache.has(currentId)
    ) {
      updateCurrentChatInfo();
      return;
    }
    if (currentChatResolveTimer) clearTimeout(currentChatResolveTimer);
    currentChatResolveTimer = setTimeout(() => {
      currentChatResolveTimer = null;
      resolveCurrentChatTitleIfNeeded();
    }, delay);
    updateCurrentChatInfo();
  }

  function invalidateConversationCacheIfNeeded() {
    const currentId = inferConversationIdFromUrl();
    if (!currentId) {
      updateCurrentChatInfo();
      return;
    }
    if (
      lastConversation?.conversationId &&
      lastConversation.conversationId !== currentId
    ) {
      lastConversation = null;
    }
    lastConversationId = currentId;
    updateCurrentChatInfo();
  }

  // --------------------
  // Helpers
  // --------------------
  function sanitizeFilename(name) {
    const safe = (name || "copilot-chat")
      .replace(/[\x00-\x1F\x7F]/g, "")
      .replace(/[\\/:*?"<>|]+/g, "-")
      .replace(/\s+/g, " ")
      .trim();
    return safe.slice(0, 120) || "copilot-chat";
  }

  function formatReadableTimestamp(ts) {
    if (!ts) return "";
    const d = new Date(ts);
    if (Number.isNaN(d.getTime())) return String(ts);
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    const ss = String(d.getSeconds()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
  }

  function buildTurnMeta(turn) {
    const bits = [];
    if (turn?.turnCount != null) bits.push(`Turn ${turn.turnCount}`);
    const ts = formatReadableTimestamp(turn?.createdAt);
    if (ts) bits.push(ts);
    if (Number.isFinite(turn?.sourceCount) && turn.sourceCount > 0) {
      bits.push(`Sources ${turn.sourceCount}`);
    }
    return bits.length ? `_${bits.join(" · ")}_` : "";
  }

  function downloadText(filename, content) {
    const blob = new Blob([content], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function setStatus(text, isError = false) {
    const el = document.getElementById("m365ce-status");
    if (!el) {
      return;
    }
    el.textContent = text;
    el.style.color = isError ? "#ff6b6b" : "#50fa7b";
  }

  function syncExportStatusForConversation(conversationId) {
    if (!conversationId) {
      return;
    }
    if (currentStatusConversationId === conversationId) {
      return;
    }
    currentStatusConversationId = conversationId;
    if (exportedConversationIds.has(conversationId)) {
      setStatus("Exported ✔ (2 files: .md + .json.md)");
    } else {
      setStatus("");
    }
  }

  function isHttpUrl(s) {
    return typeof s === "string" && /^https?:\/\//i.test(s);
  }

  function domainFromUrl(url) {
    try {
      const host = new URL(url).hostname.replace(/^www\./i, "");
      return host || "";
    } catch {
      return "";
    }
  }

  function stripSearchPayloadDecorators(text) {
    return String(text || "")
      .trim()
      .replace(/^\*+/, "")
      .replace(/\*+$/, "")
      .trim();
  }

  function hiddenTextLooksLikeSearchInvocation(text) {
    return /^search_web\s*\(/i.test(stripSearchPayloadDecorators(text));
  }

  function looksLikeSearchResultsObject(value) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return false;
    }
    const searchResultKeys = [
      "Images",
      "AppResults",
      "WebPages",
      "web_search_results",
    ];
    for (const key of searchResultKeys) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        return true;
      }
    }
    return false;
  }

  function looksLikeRawSearchResultPayload(text) {
    const t = stripSearchPayloadDecorators(text);
    if (!t) return false;
    if (/^(images|web_pages|app_results|web_search_results)\s*\(/i.test(t)) {
      return true;
    }
    if (/^\{\s*"(Images|AppResults|WebPages|web_search_results)"\s*:/i.test(t)) {
      return true;
    }
    if (
      /^\[\s*\{\s*"query"\s*:\s*"/i.test(t) &&
      /"result"\s*:\s*"\{\\?"?(Images|AppResults|WebPages|web_search_results)/i.test(
        t,
      )
    ) {
      return true;
    }
    try {
      const parsed = JSON.parse(t);
      if (looksLikeSearchResultsObject(parsed)) return true;
      if (Array.isArray(parsed)) {
        for (const item of parsed) {
          if (!item || typeof item !== "object") {
            continue;
          }
          if (looksLikeSearchResultsObject(item)) {
            return true;
          }
          const result =
            typeof item.result === "string" ? item.result.trim() : "";
          if (!result) {
            continue;
          }
          try {
            if (looksLikeSearchResultsObject(JSON.parse(result))) {
              return true;
            }
          } catch {
            if (
              /^\{\s*"(Images|AppResults|WebPages|web_search_results)"\s*:/i.test(
                result,
              )
            ) {
              return true;
            }
          }
        }
      }
    } catch {
      // ignore non-JSON text
    }
    return false;
  }

  function renderCitationsDetails(citations) {
    const items = uniqBy(citations || [], (c) => c.url).filter((c) =>
      isHttpUrl(c.url),
    );
    if (!items.length) return "";
    const lines = [
      "<details>",
      `<summary>Sources / citations (${items.length})</summary>`,
      "",
    ];
    for (const c of items) {
      const title = c.title || c.url;
      const domain = domainFromUrl(c.url);
      lines.push(`- [${title}](${c.url})${domain ? ` — ${domain}` : ""}`);
    }
    lines.push("", "</details>");
    return lines.join("\n");
  }

  function improveDisplayMathSpacing(text) {
    if (!text) return text;
    const lines = String(text).split(/\r?\n/);
    const out = [];
    for (const line of lines) {
      const trimmed = line.trim();
      const isStandaloneMath =
        /^\$[^$]+\$$/.test(trimmed) &&
        /(?:=|\\frac|\\sum|\\int|\\sqrt|\^|_|\\begin|\\end)/.test(trimmed);
      if (isStandaloneMath && out.length && out[out.length - 1] !== "") {
        out.push("");
      }
      out.push(line);
      if (isStandaloneMath) out.push("");
    }
    return out
      .join("\n")
      .replace(/\n{4,}/g, "\n\n\n")
      .trim();
  }

  function uniqBy(arr, keyFn) {
    const out = [];
    const seen = new Set();
    for (const item of arr || []) {
      const k = keyFn(item);
      if (!k) {
        continue;
      }
      if (seen.has(k)) {
        continue;
      }
      seen.add(k);
      out.push(item);
    }
    return out;
  }

  function escapeHtml(s) {
    return (s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function escapeMarkdown(s) {
    return String(s ?? "").replace(/([\`*_{}\[\]()#+\-.!|>])/g, "\$1");
  }

  function collapseRepeatedTrailingLinks(text) {
    if (!text) return text;
    const lines = String(text).split(/\r?\n/);
    const mdLinkRe = /(\[[^\]]+\]\((https?:\/\/[^\)]+)\))\s*$/;
    let prevUrl = null;
    for (let i = 0; i < lines.length; i++) {
      const m = lines[i].match(mdLinkRe);
      if (!m) {
        prevUrl = null;
        continue;
      }
      const full = m[1];
      const url = m[2];
      if (prevUrl && url === prevUrl) {
        lines[i] = lines[i].replace(full, "").replace(/\s+$/, "");
      } else {
        prevUrl = url;
      }
    }
    return lines.filter((l) => l !== "").join("\n");
  }

  // --------------------
  // Status line handling
  // --------------------
  const STATUS_TOKENS = new Set([
    "Reviewing the data...",
    "Thinking...",
    "No content returned",
    "Success",
    "Failure",
  ]);

  function normalizeStatusToken(line) {
    const t = (line || "").trim();
    return t.replace(/^_+/, "").replace(/_+$/, "").trim();
  }

  function collapseDuplicateStatusLines(text) {
    if (!text || !COLLAPSE_DUP_STATUS) return text;
    const lines = String(text).split(/\r?\n/);
    const out = [];
    let prev = null;
    for (const line of lines) {
      const tok = normalizeStatusToken(line);
      if (STATUS_TOKENS.has(tok) && prev === tok) {
        continue;
      }
      out.push(line);
      prev = STATUS_TOKENS.has(tok) ? tok : null;
    }
    return out.join("\n");
  }

  // Always drop "No content returned" (noise).
  function dropNoContentReturned(text) {
    if (!text) return text;
    const lines = String(text).split(/\r?\n/);
    const out = [];
    for (const line of lines) {
      if (normalizeStatusToken(line) === "No content returned") {
        continue;
      }
      out.push(line);
    }
    return out.join("\n");
  }

  // --------------------
  // Reasoning step formatting (optional)
  // --------------------
  // Converts sequences like:
  //   *Step title* Some body text...
  // into:
  //   - ✅ **Step title**
  //     - Some body text...
  // Only outside code fences.
  function formatReasoningSteps(text) {
    return text;
  }

  // --------------------
  // System-ish italics (outside code fences)
  // --------------------
  const SYSTEM_LINE_EXACT = new Set([
    "Reviewing the data...",
    "Thinking...",
    "No content returned",
    "Success",
    "Failure",
  ]);

  function italicizeSystemishOutsideFences(text) {
    if (!text) return text;

    // Convert patterns like **Piecing together repo files**OK, ... into italics with spacing
    let out = String(text).replace(
      /\*\*([^*]{3,80})\*\*\s*(?=[A-Z0-9])/g,
      "*$1* ",
    );

    const lines = out.split(/\r?\n/);
    const newLines = [];
    let inFence = false;

    for (const line of lines) {
      const trimmed = line.trim();
      if (trimmed.startsWith("```")) {
        inFence = !inFence;
        newLines.push(line);
        continue;
      }
      if (inFence) {
        newLines.push(line);
        continue;
      }

      if (!trimmed) {
        newLines.push(line);
        continue;
      }

      // Don't italicize bullet reasoning steps
      if (trimmed.startsWith("- ✅ **")) {
        newLines.push(line);
        continue;
      }

      const norm = normalizeStatusToken(trimmed);

      if (SYSTEM_LINE_EXACT.has(norm)) {
        newLines.push(`_${norm}_`);
        continue;
      }

      if (
        /^Traceback\b/.test(norm) ||
        /^FileNotFoundError\b/.test(norm) ||
        /^SyntaxWarning\b/.test(norm)
      ) {
        newLines.push(`_${norm}_`);
        continue;
      }

      newLines.push(line);
    }

    return newLines.join("\n");
  }

  // Normalize LaTeX display math: \[ ... \] to $ ... $
  // Implemented without regex for safety and portability

  // Normalize inline LaTeX math: \( ... \) to $ ... $ (outside code fences)
  function normalizeInlineLatexMath(text) {
    if (!text) return text;
    let out = "";
    let i = 0;
    let inFence = false;
    while (i < text.length) {
      if (text.startsWith("```", i)) {
        inFence = !inFence;
        out += "```";
        i += 3;
        continue;
      }
      if (!inFence && text[i] === "\\" && text[i + 1] === "(") {
        i += 2;
        const start = i;
        while (i < text.length && !(text[i] === "\\" && text[i + 1] === ")")) {
          i++;
        }
        const inner = text.slice(start, i).trim();
        out += "$" + inner + "$";
        i += 2;
        continue;
      }
      out += text[i];
      i++;
    }
    return out;
  }

  function normalizeLatexDisplayMath(text) {
    if (!text) return text;
    let out = "";
    let i = 0;
    while (i < text.length) {
      if (text[i] === "\\" && text[i + 1] === "[") {
        i += 2;
        const start = i;
        while (i < text.length && !(text[i] === "\\" && text[i + 1] === "]")) {
          i++;
        }
        const inner = text.slice(start, i).trim();
        out += "$" + inner + "$";
        i += 2;
      } else {
        out += text[i];
        i++;
      }
    }
    return out;
  }

  // --------------------
  // Active rendering helpers
  // v0.1.25: F02 fix — do not collapse repeated 'Coding and executing' lines; distinct tool runs are preserved per turn.
  // Active rendering helpers (v0.1.23)
  // --------------------

  function repairSplitMarkdownHeadingContinuations(text) {
    if (!text) {
      return text;
    }
    const lines = String(text).replace(/\r\n/g, "\n").split("\n");
    const out = [];
    let inFence = false;
    let fenceMarker = "";
    let fenceLength = 0;

    function isFenceLine(trimmed) {
      return trimmed.match(/^(`{3,}|~{3,})(.*)$/);
    }

    function isContinuationCandidate(value) {
      const s = String(value || "").trim();
      if (!s || s.length > 90) {
        return false;
      }
      if (/^(?:#{1,6}\s+|[-*+]\s+|<\/?details\b|<summary\b|```+|~~~+)/.test(s)) {
        return false;
      }
      return /^[A-Z0-9`"'“‘(]/.test(s) && /[?:)]$/.test(s);
    }

    for (let i = 0; i < lines.length; i += 1) {
      const line = lines[i];
      const trimmed = line.trim();
      const fence = isFenceLine(trimmed);
      if (fence) {
        const marker = fence[1][0];
        const length = fence[1].length;
        if (!inFence) {
          inFence = true;
          fenceMarker = marker;
          fenceLength = length;
        } else if (marker === fenceMarker && length >= fenceLength) {
          inFence = false;
          fenceMarker = "";
          fenceLength = 0;
        }
        out.push(line);
        continue;
      }
      if (!inFence && /^#{1,6}\s+/.test(trimmed) && i + 1 < lines.length && isContinuationCandidate(lines[i + 1])) {
        const currentBody = trimmed.replace(/^#{1,6}\s+/, "");
        const nextTrimmed = lines[i + 1].trim();
        if (currentBody.length >= 18 && currentBody.length <= 100 && !/[.!?:]$/.test(currentBody)) {
          out.push(`${trimmed} ${nextTrimmed}`);
          i += 1;
          continue;
        }
      }
      out.push(line);
    }
    return out.join("\n");
  }

  function normalizeMarkdownBlockBoundaries(text) {
    if (!text) return text;
    const lines = String(text).replace(/\r\n/g, "\n").split("\n");
    const out = [];
    let inFence = false;

    function pushBlankLine() {
      if (out.length === 0) {
        return;
      }
      if (out[out.length - 1] !== "") out.push("");
    }

    function isEnumeratedListLine(trimmed) {
      return /^([A-Z]|\d+)\.\s+/.test(trimmed);
    }

    function isStructuralLine(trimmed) {
      return (
        !trimmed ||
        /^#{1,6}\s+/.test(trimmed) ||
        /^---+$/.test(trimmed) ||
        /^```+/.test(trimmed) ||
        trimmed.startsWith("<details>") ||
        trimmed.startsWith("</details>") ||
        trimmed.startsWith("<summary>") ||
        trimmed.startsWith("**Sources:**") ||
        trimmed.startsWith("<summary>Sources / citations") ||
        trimmed.startsWith("**Images:**") ||
        trimmed.startsWith("**Links (from cards / metadata):**") ||
        trimmed.startsWith("**Code:**") ||
        trimmed.startsWith("**Output:**") ||
        trimmed.startsWith("**Error:**") ||
        trimmed.startsWith("**Status:**") ||
        /^[-*+]\s+/.test(trimmed) ||
        isEnumeratedListLine(trimmed)
      );
    }

    function splitGluedHeadingLine(trimmed) {
      const headingMatch = trimmed.match(/^(#{1,6})\s+(.+)$/);
      if (!headingMatch) {
        return [trimmed];
      }
      const rawBody = headingMatch[2].trim();

      function wordCount(value) {
        return String(value || "").trim().split(/\s+/).filter(Boolean).length;
      }

      function looksLikeEnumeratedHeading(value) {
        return /^(?:\d{1,3}|[A-Z]{1,2})\)\s+/.test(value);
      }

      function looksLikeTitleFragment(value) {
        const s = value.trim();
        if (!looksLikeEnumeratedHeading(s)) {
          return false;
        }
        if (s.length < 18 || s.length > 120 || /[.!?]$/.test(s)) {
          return false;
        }
        const count = wordCount(s);
        return count >= 3 && count <= 12;
      }

      function looksLikeParagraphStart(value) {
        const s = value.trim();
        if (s.length < 24 || wordCount(s) < 5) {
          return false;
        }
        if (/^(?:[-*+]\s+|```+|~~~+|#{1,6}\s+)/.test(s)) {
          return false;
        }
        return /^["'“‘(\[]?[A-Z0-9]/.test(s);
      }

      function splitAfterEmphasizedHeadingTitle(value) {
        const match = value.match(/^(.{1,120}\*[^*]{2,160}\*)\s+([A-Z][\s\S]{24,})$/);
        if (!match) {
          return null;
        }
        const title = match[1].trim();
        const body = match[2].trim();
        if (wordCount(title) <= 16 && looksLikeParagraphStart(body)) {
          return [title, body];
        }
        return null;
      }

      const emphasizedSplit = splitAfterEmphasizedHeadingTitle(rawBody);
      if (emphasizedSplit) {
        return [`${headingMatch[1]} ${emphasizedSplit[0]}`, emphasizedSplit[1]];
      }

      if (!looksLikeEnumeratedHeading(rawBody)) {
        return [trimmed];
      }
      const minTitleLength = 28;
      const maxTitleLength = Math.min(120, Math.max(0, rawBody.length - 24));
      for (let i = minTitleLength; i <= maxTitleLength; i += 1) {
        if (!/\s/.test(rawBody[i])) {
          continue;
        }
        const title = rawBody.slice(0, i).trim();
        const body = rawBody.slice(i).trim();
        if (looksLikeTitleFragment(title) && looksLikeParagraphStart(body)) {
          return [`${headingMatch[1]} ${title}`, body];
        }
      }
      return [trimmed];
    }

    function normalizeAnswerHeadingLine(trimmed) {
      const headingMatch = trimmed.match(/^(#{1,6})\s+(.+)$/);
      if (!headingMatch) {
        return trimmed;
      }
      const body = headingMatch[2].trim();
      if (/^(👤|🤖)\s+/.test(body)) {
        return trimmed;
      }
      if (/^\d{1,3}\)\s+/.test(body)) {
        return `### ${body}`;
      }
      if (/^[A-Z]{1,2}\)\s+/.test(body)) {
        return `#### ${body}`;
      }
      if (headingMatch[1].length < 3) {
        return `### ${body}`;
      }
      return trimmed;
    }

    function processLine(line) {
      const trimmed = line.trim();
      if (/^```+/.test(trimmed)) {
        inFence = !inFence;
        out.push(line);
        return;
      }
      if (!inFence) {
        if (/^#{1,6}\s+/.test(trimmed)) {
          const splitParts = splitGluedHeadingLine(trimmed);
          const safeHeading = normalizeAnswerHeadingLine(splitParts[0]);
          pushBlankLine();
          out.push(safeHeading);
          out.push("");
          if (splitParts.length > 1) {
            out.push(splitParts[1]);
            out.push("");
          }
          return;
        }
        if (
          trimmed.startsWith("**Sources:**") ||
          trimmed.startsWith("<summary>Sources / citations") ||
          trimmed.startsWith("**Images:**") ||
          trimmed.startsWith("**Links (from cards / metadata):**") ||
          trimmed.startsWith("**Code:**") ||
          trimmed.startsWith("**Output:**") ||
          trimmed.startsWith("**Error:**")
        ) {
          pushBlankLine();
          out.push(trimmed);
          out.push("");
          return;
        }
        if (isEnumeratedListLine(trimmed)) {
          pushBlankLine();
          out.push(trimmed);
          return;
        }
        if (
          /^\*\*[^*].*?\*\*\s+/.test(trimmed) &&
          !/^\*\*(Status|Sources|Images|Links \(from cards \/ metadata\)|Error|Tool run|Reasoning|Code|Output)/.test(
            trimmed,
          )
        ) {
          const m = trimmed.match(/^(\*\*[^*].*?\*\*)(\s+)(.+)$/);
          if (m) {
            pushBlankLine();
            out.push(m[1]);
            out.push("");
            out.push(m[3]);
            return;
          }
        }
        if (/^[A-Z][a-z].+\.[A-Z]/.test(trimmed)) {
          const dot = trimmed.indexOf(".");
          const rest = trimmed.slice(dot + 1).trim();
          if (rest) {
            pushBlankLine();
            out.push(trimmed.slice(0, dot + 1));
            out.push("");
            out.push(rest);
            return;
          }
        }
        if (/^---+$/.test(trimmed)) {
          pushBlankLine();
          out.push("---");
          out.push("");
          return;
        }
      }
      out.push(line);
    }

    for (const line of lines) processLine(line);
    const compact = out
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
    const compactLines = compact.split("\n");
    const merged = [];
    inFence = false;

    for (let i = 0; i < compactLines.length; i++) {
      const line = compactLines[i];
      const trimmed = line.trim();
      if (/^```+/.test(trimmed)) {
        inFence = !inFence;
        merged.push(line);
        continue;
      }
      if (!inFence && isEnumeratedListLine(trimmed)) {
        let item = trimmed;
        let j = i + 1;
        while (j < compactLines.length) {
          const nextTrimmed = compactLines[j].trim();
          if (isStructuralLine(nextTrimmed)) {
            break;
          }
          item += " " + nextTrimmed.replace(/\s+/g, " ").trim();
          j += 1;
        }
        merged.push(item);
        i = j - 1;
        continue;
      }
      merged.push(line);
    }

    return merged
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function guessFenceLangFromText(text) {
    const t = String(text || "").trim();
    if (!t) return "text";
    if (
      /^(import\s+\S+|from\s+\S+\s+import\s+|def\s+\w+\(|class\s+\w+\b|try:|with\s+\w+\b|for\s+\w+\s+in\b|if\s+.+:)/m.test(
        t,
      )
    ) {
      return "python";
    }
    if (/^\s*[{\[]/.test(t)) return "json";
    return "text";
  }

  function longestFenceRun(text, marker) {
    const escaped = marker === "`" ? "`" : "~";
    const re = new RegExp(`${escaped}+`, "g");
    const matches = String(text || "").match(re) || [];
    let maxLen = 0;
    for (const m of matches) {
      if (m.length > maxLen) maxLen = m.length;
    }
    return maxLen;
  }

  function longestBacktickRun(text) {
    return longestFenceRun(text, "`");
  }

  function renderFencedBlock(text, lang = "text") {
    if (!text) return "";
    const body = String(text).replace(/\r\n/g, "\n").trimEnd();
    const backtickLen = Math.max(3, longestFenceRun(body, "`") + 1);
    const tildeLen = Math.max(3, longestFenceRun(body, "~") + 1);
    const marker = backtickLen <= tildeLen ? "`" : "~";
    const fence = marker.repeat(marker === "`" ? backtickLen : tildeLen);
    const opener = lang ? `${fence}${lang}` : fence;
    return [opener, body, fence].join("\n");
  }


  function hardenNestedMarkdownFenceExamples(text) {
    if (!text) return text;
    const lines = String(text).replace(/\r\n/g, "\n").split("\n");
    for (let i = 0; i < lines.length - 3; i += 1) {
      const opener = lines[i].trim().match(/^(```+|~~~+)(markdown|md|text)\s*$/i);
      if (!opener) {
        continue;
      }
      const marker = opener[1][0];
      const outerLen = opener[1].length;
      const innerOpen = lines[i + 1].trim().match(/^(```+|~~~+)[A-Za-z0-9_-]*\s*$/);
      if (!innerOpen || innerOpen[1][0] !== marker || innerOpen[1].length < outerLen) {
        continue;
      }
      let innerClose = -1;
      for (let j = i + 2; j < lines.length; j += 1) {
        const close = lines[j].trim().match(/^(```+|~~~+)\s*$/);
        if (close && close[1][0] === marker && close[1].length >= innerOpen[1].length) {
          innerClose = j;
          break;
        }
      }
      if (innerClose < 0 || innerClose + 1 >= lines.length) {
        continue;
      }
      const outerClose = lines[innerClose + 1].trim().match(/^(```+|~~~+)\s*$/);
      if (!outerClose || outerClose[1][0] !== marker || outerClose[1].length < outerLen) {
        continue;
      }
      const body = lines.slice(i + 1, innerClose + 1).join("\n");
      const safeLen = Math.max(outerLen, longestFenceRun(body, marker) + 1);
      const safeFence = marker.repeat(safeLen);
      const suffix = lines[i].trim().slice(opener[1].length);
      lines[i] = `${safeFence}${suffix}`;
      lines[innerClose + 1] = safeFence;
      i = innerClose + 1;
    }
    return lines.join("\n");
  }

  function promoteOuterMarkdownExampleFences(text) {
    if (!text) return text;
    const lines = String(text).replace(/\r\n/g, "\n").split("\n");

    function nextShortFenceNonBlankIndex(start) {
      for (let i = start; i < lines.length; i += 1) {
        if (lines[i].trim()) return i;
      }
      return -1;
    }

    function isAfterExampleBoundary(trimmed) {
      return trimmed === "---" || /^#{1,6}\s+/.test(trimmed) || /^\*+\s*(Notes|Rules|Definition|Rendering|Reason|Include|Collapsed)\b/i.test(trimmed) || /^➡\s+Rendered as:/i.test(trimmed) || /^\|/.test(trimmed);
    }

    for (let i = 0; i < lines.length - 2; i += 1) {
      const opener = lines[i].trim().match(/^(```+|~~~+)(markdown|md|text)\s*$/i);
      if (!opener) {
        continue;
      }
      const marker = opener[1][0];
      const outerLen = opener[1].length;
      const fenceLines = [];
      let closeIdx = -1;
      const scanLimit = Math.min(lines.length, i + 160);
      for (let j = i + 1; j < scanLimit; j += 1) {
        const candidate = lines[j].trim().match(/^(```+|~~~+)[A-Za-z0-9_-]*\s*$/);
        if (candidate && candidate[1][0] === marker && candidate[1].length >= outerLen) {
          fenceLines.push(j);
          const nextIdx = nextShortFenceNonBlankIndex(j + 1);
          const nextTrimmed = nextIdx >= 0 ? lines[nextIdx].trim() : "";
          if (nextIdx < 0 || isAfterExampleBoundary(nextTrimmed)) {
            if (fenceLines.length >= 2) {
              closeIdx = j;
            }
            break;
          }
        }
      }
      if (closeIdx < 0) {
        continue;
      }
      const body = lines.slice(i + 1, closeIdx).join("\n");
      const safeLen = Math.max(outerLen, longestFenceRun(body, marker) + 1);
      const safeFence = marker.repeat(safeLen);
      const suffix = lines[i].trim().slice(opener[1].length);
      lines[i] = `${safeFence}${suffix}`;
      lines[closeIdx] = safeFence;
      i = closeIdx;
    }
    return lines.join("\n");
  }

  function closeShortFenceClosersBeforeMarkdownBoundaries(text) {
    if (!text) {
      return text;
    }
    const lines = String(text).replace(/\r\n/g, "\n").split("\n");
    let openMarker = "";
    let openLength = 0;
    let nestedShortFenceLength = 0;

    function cascadeShortFenceNextNonBlankIndex(start) {
      for (let i = start; i < lines.length; i += 1) {
        if (lines[i].trim()) {
          return i;
        }
      }
      return -1;
    }

    function cascadeShortFencePreviousNonBlankIndex(start) {
      for (let i = start; i >= 0; i -= 1) {
        if (lines[i].trim()) {
          return i;
        }
      }
      return -1;
    }

    function looksLikeShortFenceBoundary(trimmed) {
      return !trimmed || trimmed === "---" || /^#{1,6}\s+/.test(trimmed) || /^(?:before|after|and|into|to|becomes|renders? as|target(?: output)?):\s*$/i.test(trimmed) || /^that should become\b/i.test(trimmed) || /^#\s+[👤🤖]\s+(?:USER|COPILOT)\b/.test(trimmed) || /^<\/?details\b/i.test(trimmed) || /^<summary\b/i.test(trimmed) || /^\*\*(?:Status|Output|Error|Code|Sources|Images|Links)\*\*:?$/i.test(trimmed);
    }

    function looksLikeExampleLeadIn(trimmed) {
      return /^(?:before|after|and|into|to|becomes|renders? as|target(?: output)?):\s*$/i.test(trimmed) || /^that should become\b/i.test(trimmed) || /^(?:the fix turns cases like|examples?|sample|expected|actual|input|output):\s*$/i.test(trimmed);
    }

    for (let i = 0; i < lines.length; i += 1) {
      const match = lines[i].match(/^(\s*)(`{3,}|~{3,})([A-Za-z0-9_-]*)\s*$/);
      if (!match) {
        continue;
      }
      const leading = match[1];
      const marker = match[2][0];
      const length = match[2].length;
      const info = match[3] || "";
      if (!openMarker) {
        openMarker = marker;
        openLength = length;
        nestedShortFenceLength = 0;
        continue;
      }
      if (marker !== openMarker) {
        continue;
      }
      if (length >= openLength) {
        openMarker = "";
        openLength = 0;
        nestedShortFenceLength = 0;
        continue;
      }
      if (openLength > 3 && length >= 3) {
        if (info) {
          nestedShortFenceLength = length;
          continue;
        }
        if (nestedShortFenceLength && length >= nestedShortFenceLength) {
          nestedShortFenceLength = 0;
          continue;
        }
        const nextIdx = cascadeShortFenceNextNonBlankIndex(i + 1);
        const prevIdx = cascadeShortFencePreviousNonBlankIndex(i - 1);
        const nextTrimmed = nextIdx >= 0 ? lines[nextIdx].trim() : "";
        const prevTrimmed = prevIdx >= 0 ? lines[prevIdx].trim() : "";
        if (nextIdx < 0 || looksLikeShortFenceBoundary(nextTrimmed) || looksLikeExampleLeadIn(prevTrimmed)) {
          lines[i] = leading + marker.repeat(openLength);
          openMarker = "";
          openLength = 0;
          nestedShortFenceLength = 0;
        }
      }
    }
    return lines.join("\n");
  }

  function indentMarkdown(text, prefix) {
    return String(text || "")
      .split(/\r?\n/)
      .map((line) => `${prefix}${line}`)
      .join("\n");
  }

  function adaptiveBodyToFieldMap(cards) {
    const fields = {};
    const parts = [];
    for (const card of cards || []) {
      const body = Array.isArray(card?.body) ? card.body : [];
      for (const node of body) {
        const text = adaptiveBodyNodeToMarkdown(node).trim();
        if (!text) {
          continue;
        }
        if (node?.id) fields[node.id] = text;
        parts.push(text);
      }
    }
    return { fields, parts };
  }

  function normalizeReasoningText(text) {
    let t = String(text || "").trim();
    t = t.replace(/^\*\*(.*?)\*\*/, "$1 —");
    return t.replace(/\s+/g, " ").trim();
  }

  function renderReasoningSection(items) {
    const lines = [];
    for (const msg of items || []) {
      const t = normalizeReasoningText(msg?.text || "");
      if (t) lines.push(`- ${t}`);
    }
    if (!lines.length) return "";
    return [
      "<details>",
      "<summary>Reasoning / progress</summary>",
      "",
      ...lines,
      "</details>",
    ].join("\n");
  }


  function isCanonicalToolRunProgressMessage(msg) {
    return (
      msg?.author === "bot" &&
      msg?.messageType === "Progress" &&
      normalizeStatusToken(msg?.text || "") === "Coding and executing" &&
      Array.isArray(msg?.adaptiveCards) &&
      msg.adaptiveCards.length
    );
  }

  function isMeaningfulGenericProgressMessage(msg) {
    if (msg?.author !== "bot" || msg?.messageType !== "Progress") {
      return false;
    }
    if (msg?.contentOrigin === "ChainOfThoughtSummary") {
      return false;
    }
    if (isCanonicalToolRunProgressMessage(msg) || isLongFileMapReduceProgressMessage(msg)) {
      return false;
    }
    const text = normalizeReasoningText(msg?.text || "");
    if (!text) {
      return false;
    }
    const statusToken = normalizeStatusToken(text);
    if (statusToken === "No content returned" || statusToken === "Success" || statusToken === "Failure") {
      return false;
    }
    return true;
  }

  function renderProcessTimelineSection(entries) {
    const orderedEntries = (entries || [])
      .filter((entry) => entry && entry.rendered)
      .sort((a, b) => a.sortAt - b.sortAt);
    if (!orderedEntries.length) {
      return "";
    }
    const stepWord = orderedEntries.length === 1 ? "step" : "steps";
    const lines = [
      "<details>",
      `<summary>Reasoning completed in ${orderedEntries.length} ${stepWord}</summary>`,
      "",
    ];
    orderedEntries.forEach((entry) => {
      if (entry.kind === "detail") {
        lines.push(entry.rendered.trim(), "");
      } else {
        lines.push(entry.rendered.trim(), "");
      }
    });
    lines.push("</details>");
    return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  function normalizeToolRunSummaryTitle(text) {
    let title = String(text || "").trim().replace(/\s+/g, " ");
    title = title.replace(/^\*\*(.*?)\*\*$/, "$1").trim();
    return title || "Tool run";
  }

  function toolRunSummaryTitle(msg, fields, parts) {
    const cardTitle = fields?.Header || fields?.Title || (parts || []).find((part) => normalizeStatusToken(part || ""));
    return normalizeToolRunSummaryTitle(cardTitle || msg?.text || "Tool run");
  }

  function renderToolRunSection(msg, allItems) {
    const { fields, parts } = adaptiveBodyToFieldMap(msg?.adaptiveCards || []);
    const header = fields.Header || normalizeStatusToken(msg?.text || "");
    if (normalizeStatusToken(header) !== "Coding and executing") {
      return "";
    }

    const code = fields.Content || msg?.hiddenText || "";
    const status = fields.Status || "";
    const output = fields.ExecutionOutput || "";
    const error = fields.ExecutionError || "";
    const summaryTitle = toolRunSummaryTitle(msg, fields, parts);
    const lines = [
      "<details>",
      `<summary>${escapeHtml(summaryTitle)}${status ? ` — ${escapeHtml(status)}` : ""}</summary>`,
      "",
    ];

    if (code) {
      lines.push("**Code:**", "");
      lines.push(renderFencedBlock(code, guessFenceLangFromText(code)), "");
    }
    if (status) {
      lines.push(`**Status:** ${status}`, "");
    }
    if (output) {
      lines.push("**Output:**", "");
      lines.push(renderFencedBlock(output, "text"), "");
    }
    if (error) {
      lines.push("**Error:**", "");
      lines.push(renderFencedBlock(error, "text"), "");
    }

    const pluginFootnote =
      pluginFootnoteForMessage(msg) ||
      pyexecPluginFootnoteForToolRun(msg, allItems, code);
    if (pluginFootnote) {
      lines.push(pluginFootnote, "");
    }
    lines.push("</details>");
    return lines
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  // --------------------
  // Copilot body Markdown hardening (viewer safety)
  // --------------------
  function normalizeCopilotBodyMarkdown(md) {
    const text = String(md || "");
    if (!text) return "";

    const lines = text.split(/\r?\n/);
    const out = [];

    let inFence = false;
    let fenceDelim = "";

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmed = line.trim();

      const fenceMatch = trimmed.match(/^(\`\`\`+|\~\~\~+)(.*)$/);
      if (fenceMatch) {
        const delim = fenceMatch[1];
        if (!inFence) {
          inFence = true;
          fenceDelim = delim;
        } else if (
          delim.startsWith(fenceDelim[0]) &&
          delim.length >= fenceDelim.length
        ) {
          inFence = false;
          fenceDelim = "";
        }
        out.push(line);
        continue;
      }

      if (inFence) {
        out.push(line);
        continue;
      }

      // Prevent setext-heading surprises from "---" / "===" inside Copilot body.
      if (/^(-{3,}|={3,}|_{3,})$/.test(trimmed)) {
        out.push("* * *");
        continue;
      }

      // Prevent empty list markers from starting list context.
      if (trimmed === "*" || trimmed === "-") {
        out.push("\\" + trimmed);
        continue;
      }

      // Demote enumerators like "1)" / "A)" / "- D)" that some viewers treat as ordered lists or columns.
      const enumMatch = trimmed.match(
        /^(?:[-*+]\s+)?(([0-9]{1,3})|([A-Z]{1,2}))\)\s+(.+)$/,
      );
      if (enumMatch) {
        out.push(`* **${enumMatch[1]}) ${enumMatch[4].trim()}**`);
        out.push("");
        continue;
      }

      out.push(line);
    }

    return out
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }
  // --------------------
  // Uploaded filename derivation
  // --------------------
  const LOCAL_FILE_ANNOTATION_TYPES = new Set(["LocalFile"]);

  function normalizeUploadedFilenameCandidate(value) {
    const s = String(value || "").trim();
    if (!s) return "";
    if (s.length > 180) return "";
    if (/^https?:\/\//i.test(s)) return "";
    if (/[<>\n\r\t]/.test(s)) return "";
    if (/[/\\]/.test(s)) return "";
    if (
      !/\.(pdf|docx?|xlsx?|csv|zip(?:\.txt)?|txt|md|json|png|jpe?g|gif|pptx?|py|js)$/i.test(
        s,
      )
    ) {
      return "";
    }
    return s;
  }

  function pushUniqueUploadedFilename(outMap, file, score, order) {
    const normalized = normalizeUploadedFilenameCandidate(file);
    if (!normalized) {
      return;
    }
    const key = normalized.toLowerCase();
    const prev = outMap.get(key);
    if (
      !prev ||
      score < prev.score ||
      (score === prev.score && order < prev.order)
    ) {
      outMap.set(key, { file: normalized, score, order });
    }
  }

  function tryParseLooseJsonFromString(raw) {
    if (typeof raw !== "string") return null;
    const trimmed = raw.trim();
    if (!trimmed) return null;
    const fence = trimmed.match(/```json\s*([\s\S]*?)```/i);
    const candidate = fence ? fence[1].trim() : trimmed;
    if (candidate.startsWith("{") || candidate.startsWith("[")) {
      try {
        return JSON.parse(candidate);
      } catch (_) {
        /* ignore */
      }
    }
    const idxObj = candidate.indexOf("{");
    const idxArr = candidate.indexOf("[");
    let idx = -1;
    if (idxObj >= 0 && idxArr >= 0) idx = Math.min(idxObj, idxArr);
    else idx = Math.max(idxObj, idxArr);
    if (idx >= 0) {
      try {
        return JSON.parse(candidate.slice(idx).trim());
      } catch (_) {
        /* ignore */
      }
    }
    return null;
  }

  function collectUploadedFilenameCandidatesFromParsed(
    value,
    outMap,
    score,
    orderBase = 0,
  ) {
    if (value == null) {
      return;
    }
    if (typeof value === "string") {
      pushUniqueUploadedFilename(outMap, value, score, orderBase);
      return;
    }
    if (Array.isArray(value)) {
      for (let i = 0; i < value.length; i++) {
        collectUploadedFilenameCandidatesFromParsed(
          value[i],
          outMap,
          score,
          orderBase + i / 100,
        );
      }
      return;
    }
    if (typeof value === "object") {
      if (typeof value.fileName === "string") {
        pushUniqueUploadedFilename(outMap, value.fileName, score, orderBase);
      }
      if (typeof value.annotatedFileName === "string") {
        pushUniqueUploadedFilename(
          outMap,
          value.annotatedFileName,
          score,
          orderBase + 0.0001,
        );
      }
      if (
        typeof value.text === "string" &&
        typeof value.type === "string" &&
        /file/i.test(value.type)
      ) {
        pushUniqueUploadedFilename(
          outMap,
          value.text,
          score,
          orderBase + 0.0002,
        );
      }
      if (Array.isArray(value.files)) {
        collectUploadedFilenameCandidatesFromParsed(
          value.files,
          outMap,
          score,
          orderBase + 0.001,
        );
      }
      if (Array.isArray(value.messageAnnotations)) {
        collectUploadedFilenameCandidatesFromParsed(
          value.messageAnnotations,
          outMap,
          score,
          orderBase + 0.002,
        );
      }
      if (Array.isArray(value.children)) {
        collectUploadedFilenameCandidatesFromParsed(
          value.children,
          outMap,
          score,
          orderBase + 0.003,
        );
      }
    }
  }

  function isAttachmentCarrierMessage(msg) {
    if (!msg) return false;

    // Strict per-turn scoping: only USER or FileUpload App messages are allowed to carry
    // uploaded-filename metadata into the USER filename bubble.
    const author = msg.author || "";
    if (author !== "user" && author !== "FileUpload App") return false;

    // Only explicit metadata sources.
    if (msg.filesAndLinksContextMetadata) return true;
    if (Array.isArray(msg.messageAnnotations) && msg.messageAnnotations.length) return true;

    return false;
  }

  function attachmentFilenamePriority(msg) {
    if (msg?.filesAndLinksContextMetadata?.files) return 0;
    if (
      msg?.author === "system" &&
      msg?.messageType === "InternalAttachedContentInfo" &&
      Array.isArray(msg?.messageAnnotations)
    ) {
      return 1;
    }
    if (msg?.author === "FileUpload App") return 2;
    if (msg?.author === "Attached Content App") return 3;
    return 4;
  }

  function collectUploadedFilenameCandidatesFromAnnotations(
    annotations,
    outMap,
    score,
    orderBase = 0,
  ) {
    if (!Array.isArray(annotations)) {
      return;
    }
    for (let i = 0; i < annotations.length; i++) {
      const ann = annotations[i];
      const type =
        ann?.messageAnnotationType || ann?.annotationType || ann?.type || "";
      const source = ann?.messageAnnotationSource || ann?.source || "";
      const isLocalFile =
        LOCAL_FILE_ANNOTATION_TYPES.has(type) ||
        /LocalFile/i.test(type) ||
        /UserAnnotated/i.test(source);
      if (!isLocalFile) {
        continue;
      }
      pushUniqueUploadedFilename(
        outMap,
        ann?.text || ann?.fileName || ann?.annotatedFileName || "",
        score,
        orderBase + i / 100,
      );
    }
  }

  function collectUploadedFilenameCandidatesFromMetadata(
    metadata,
    outMap,
    score,
    orderBase = 0,
  ) {
    if (!metadata) {
      return;
    }
    const files = Array.isArray(metadata?.files) ? metadata.files : [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      pushUniqueUploadedFilename(
        outMap,
        file?.fileName || "",
        score,
        orderBase + i / 100,
      );
      pushUniqueUploadedFilename(
        outMap,
        file?.annotatedFileName || "",
        score,
        orderBase + i / 100 + 0.0001,
      );
    }
  }

  function collectUploadedFilenameCandidatesFromAttachmentText(
    raw,
    outMap,
    score,
    orderBase = 0,
  ) {
    const parsed = tryParseLooseJsonFromString(raw);
    if (!parsed) {
      return;
    }
    collectUploadedFilenameCandidatesFromParsed(
      parsed,
      outMap,
      score,
      orderBase,
    );
  }

  function extractUploadedFilesFromTurn(items) {
    const ranked = new Map();
    for (let idx = 0; idx < (items || []).length; idx++) {
      const msg = items[idx];
      if (!isAttachmentCarrierMessage(msg)) {
        continue;
      }
      const score = attachmentFilenamePriority(msg);
      collectUploadedFilenameCandidatesFromMetadata(
        msg?.filesAndLinksContextMetadata,
        ranked,
        score,
        idx,
      );
      collectUploadedFilenameCandidatesFromAnnotations(
        msg?.messageAnnotations,
        ranked,
        score + 0.1,
        idx + 0.01,
      );
      if (
        msg?.author === "system" &&
        msg?.messageType === "InternalAttachedContentInfo"
      ) {
        collectUploadedFilenameCandidatesFromAttachmentText(
          msg?.text,
          ranked,
          score + 0.2,
          idx + 0.02,
        );
      }
      if (
        msg?.author === "FileUpload App" ||
        msg?.author === "Attached Content App"
      ) {
        collectUploadedFilenameCandidatesFromAttachmentText(
          msg?.hiddenText,
          ranked,
          score + 0.3,
          idx + 0.03,
        );
      }
    }
    return Array.from(ranked.values())
      .sort(
        (a, b) =>
          a.score - b.score ||
          a.order - b.order ||
          a.file.localeCompare(b.file),
      )
      .map((x) => x.file)
      .slice(0, 6);
  }

  // --------------------
  // Rich extraction
  // --------------------
  function groupContiguous(turns) {
    return (turns || []).map((t) => ({ role: t.role, items: [t] }));
  }

  function spacer() {
    // Strong visual separation between turns, especially COPILOT → next USER.
    return ADD_RULE_BETWEEN_BLOCKS ? "\n\n\n---\n\n\n\n" : "\n\n\n\n";
  }

  function ensureLeadingDetailsSeparation(text) {
    const src = String(text || "").replace(/\r\n/g, "\n").trim();
    if (!src.startsWith("<details>")) {
      return src;
    }
    const lines = src.split("\n");
    let i = 0;
    let consumedLeadingDetails = false;
    while (i < lines.length) {
      while (i < lines.length && !lines[i].trim()) {
        i += 1;
      }
      if (lines[i]?.trim() !== "<details>") {
        break;
      }
      const closeIdx = lines.findIndex((line, idx) => idx > i && line.trim() === "</details>");
      if (closeIdx < 0) {
        return src;
      }
      consumedLeadingDetails = true;
      i = closeIdx + 1;
    }
    if (!consumedLeadingDetails || i >= lines.length) {
      return src;
    }
    const before = lines.slice(0, i).join("\n").trimEnd();
    const after = lines.slice(i).join("\n").trimStart();
    if (!after) {
      return before;
    }
    return `${before}\n\n\n${after}`;
  }

  function renderGroup(group) {
    const role = group.role;
    const textParts = [];
    const seenText = new Set();
    for (const item of group.items || []) {
      const raw = (item?.text || "").replace(/\r\n/g, "\n").trim();
      if (!raw) {
        continue;
      }
      const key = normalizeTextForCompare(raw);
      if (!key || seenText.has(key)) {
        continue;
      }
      seenText.add(key);
      textParts.push(raw);
    }
    const combinedText = textParts.join("\n\n");
    const uploadedFiles = uniqBy(
      group.items?.[0]?.uploadedFiles || [],
      (x) => x,
    ).slice(0, 6);

    const heading =
      role === "User"
        ? "👤 USER"
        : role === "Copilot"
          ? "🤖 COPILOT"
          : String(role).toUpperCase();
    const meta = buildTurnMeta(group.items?.[0] || { role });
    const lines = [];

    if (role === "User") {
      lines.push("");
      lines.push("");
      lines.push(`# ${heading}`);
      if (meta) lines.push(normalizeCopilotBodyMarkdown(meta));
      if (combinedText) {
        // USER text is intentionally treated as an opaque chat-bubble payload.
        // Do not run Markdown normalization on USER text; preserve line breaks via <pre>.
        lines.push('<div align="right" style="text-align:right">');
        lines.push(
          `<pre style="white-space:pre-wrap; font-size:14px; margin:0 0 0 auto; max-width:40%; text-align:left; background:#2a2a45; padding:8px 12px; border-radius:12px;">${escapeHtml(combinedText)}</pre>`,
        );
        lines.push("</div>");
      }
      if (uploadedFiles.length) {
        lines.push(
          '<div align="right" style="text-align:right; margin-top:6px;">',
        );
        lines.push(
          `<pre style="white-space:pre-wrap; font-size:13px; margin:0 0 0 auto; max-width:40%; text-align:left; background:#20364d; padding:8px 12px; border-radius:12px;">${escapeHtml(uploadedFiles.join("\n"))}</pre>`,
        );
        lines.push("</div>");
      }
      lines.push("");
      return lines.join("\n");
    }

    lines.push(`# ${heading}`);
    if (meta) lines.push(meta);
    lines.push("");
    lines.push(combinedText || "");

    const textForDedupe = combinedText || "";
    const textHasUrl = (u) => u && textForDedupe.includes(u);

    const images = uniqBy(
      group.items.flatMap((x) => x.images || []),
      (i) => i.url,
    ).filter((i) => !textHasUrl(i.url));
    const citations = uniqBy(
      group.items.flatMap((x) => x.citations || []),
      (c) => c.url,
    ).filter((c) => !textHasUrl(c.url));
    const additionalSourceAttributions = uniqBy(
      group.items.flatMap((x) => x.additionalSourceAttributions || []),
      sourceAttributionKey,
    ).filter((s) => !textHasUrl(s.url));
    const queries = uniqBy(
      group.items.flatMap((x) => x.queries || []),
      (q) => normalizeQueryCandidate(q).toLowerCase(),
    ).filter(Boolean);
    const links = uniqBy(
      group.items.flatMap((x) => x.links || []),
      (l) => l.url,
    ).filter((l) => !textHasUrl(l.url));

    if (images.length) {
      lines.push("");
      lines.push("**Images:**");
      for (const img of images) {
        lines.push(`![${img.alt || "image"}](${img.url})`);
      }
    }

    const searchProvenance = renderSearchProvenanceDetails({
      queries,
      citations,
      additionalSourceAttributions,
      searchPlugins: uniqBy(
        group.items.flatMap((x) => x.searchPlugins || []),
        (p) => p.name,
      ),
    });
    if (searchProvenance) {
      lines.push("", searchProvenance);
    }

    if (links.length) {
      const citeUrls = new Set(citations.map((x) => x.url));
      const extras = links.filter((x) => !citeUrls.has(x.url));
      if (extras.length) {
        lines.push("");
        lines.push("**Links (from cards / metadata):**");
        for (const l of extras) lines.push(`- [${l.title || l.url}](${l.url})`);
      }
    }

    lines.push("");
    return lines.join("\n");
  }

  function normalizeTextForCompare(s) {
    return String(s || "")
      .replace(/\r\n/g, "\n")
      .trim()
      .toLowerCase();
  }

  function richTextInlineToMarkdown(inline) {
    if (!inline) return "";
    let text = "";
    if (typeof inline === "string") text = inline;
    else if (typeof inline.text === "string") text = inline.text;
    if (!text) return "";

    const alreadyWrapped = /^\*\*.*\*\*$/.test(text) || /^\*.*\*$/.test(text);
    if (!alreadyWrapped) {
      if (inline.weight === "Bolder" || inline.bold === true) {
        text = `**${text}**`;
      }
      if (inline.italic === true) text = `*${text}*`;
    }
    return text;
  }

  function adaptiveBodyNodeToMarkdown(node) {
    if (!node || typeof node !== "object") return "";
    if (node.type === "TextBlock") return node.text || "";
    if (node.type === "RichTextBlock") {
      return (node.inlines || []).map(richTextInlineToMarkdown).join("");
    }
    return "";
  }

  function extractAdaptiveCardsMarkdown(cards) {
    const parts = [];
    for (const card of cards || []) {
      const body = Array.isArray(card?.body) ? card.body : [];
      for (const node of body) {
        const text = adaptiveBodyNodeToMarkdown(node).trim();
        if (text) parts.push(text);
      }
    }
    return parts.join("\n\n").trim();
  }

  function sourceAttributionUrl(s) {
    return s?.seeMoreUrl || s?.url || s?.href || "";
  }

  function sourceAttributionTitle(s) {
    return (
      s?.providerDisplayName ||
      s?.displayName ||
      s?.title ||
      sourceAttributionUrl(s)
    );
  }

  function sourceAttributionIsExplicitlyUncited(s) {
    if (!s || !Object.prototype.hasOwnProperty.call(s, "isCitedInResponse")) {
      return false;
    }
    return String(s.isCitedInResponse).trim().toLowerCase() === "false";
  }

  function sourceAttributionIsCitedOrUnspecified(s) {
    return !sourceAttributionIsExplicitlyUncited(s);
  }

  function sourceAttributionSearchQuery(s) {
    return s?.searchQuery || s?.query || "";
  }

  function sourceAttributionKey(s) {
    return [
      sourceAttributionUrl(s),
      sourceAttributionSearchQuery(s),
      sourceAttributionTitle(s),
    ]
      .join("")
      .toLowerCase();
  }

  function sourceAttributionsToCitations(msg) {
    const items = [];
    for (const s of msg?.sourceAttributions || []) {
      const url = sourceAttributionUrl(s);
      const title = sourceAttributionTitle(s);
      if (isHttpUrl(url) && sourceAttributionIsCitedOrUnspecified(s)) {
        items.push({ title, url });
      }
    }
    return uniqBy(items, (x) => x.url);
  }

  function sourceAttributionsToAdditionalItems(items, citations) {
    const citedUrls = new Set(
      (citations || []).map((c) => c.url).filter(Boolean),
    );
    const out = [];
    for (const msg of items || []) {
      for (const s of msg?.sourceAttributions || []) {
        const url = sourceAttributionUrl(s);
        if (!isHttpUrl(url)) {
          continue;
        }
        if (!sourceAttributionIsExplicitlyUncited(s) && citedUrls.has(url)) {
          continue;
        }
        out.push({
          title: sourceAttributionTitle(s),
          url,
          provider: s?.providerDisplayName || s?.displayName || "",
          searchQuery: sourceAttributionSearchQuery(s),
          cited: sourceAttributionIsExplicitlyUncited(s)
            ? "false"
            : "unspecified",
        });
      }
    }
    return uniqBy(out, sourceAttributionKey);
  }

  function renderAdditionalSourceAttributionsDetails(attributions) {
    const items = uniqBy(attributions || [], sourceAttributionKey).filter((s) =>
      isHttpUrl(s.url),
    );
    if (!items.length) {
      return "";
    }
    const lines = [
      "<details>",
      `<summary>Additional source attributions (${items.length})</summary>`,
      "",
    ];
    for (const s of items) {
      const title = s.title || s.url;
      const domain = domainFromUrl(s.url);
      const bits = [];
      if (domain) {
        bits.push(domain);
      }
      if (s.cited) {
        bits.push(`cited: ${s.cited}`);
      }
      if (s.searchQuery) {
        bits.push(`query: ${s.searchQuery}`);
      }
      lines.push(
        `- [${title}](${s.url})${bits.length ? ` — ${bits.join(" · ")}` : ""}`,
      );
    }
    lines.push("", "</details>");
    return lines.join("\n");
  }

  function searchPluginMessagesToSummaries(items) {
    const counts = new Map();
    for (const msg of items || []) {
      if (!isSearchProvenancePluginMessage(msg)) {
        continue;
      }
      const name = pluginDisplayName(msg.pluginInfo);
      const key = name.toLowerCase();
      const existing = counts.get(key) || { name, count: 0 };
      existing.count += 1;
      counts.set(key, existing);
    }
    return Array.from(counts.values());
  }

  function renderSearchProvenanceDetails(input) {
    const queries = uniqBy(input?.queries || [], (q) => normalizeQueryCandidate(q).toLowerCase()).filter(Boolean);
    const citedSources = uniqBy(input?.citations || [], (c) => c.url).filter((c) => isHttpUrl(c.url)).map((c) => ({ ...c, provenanceType: "cited" }));
    const additionalSources = uniqBy(input?.additionalSourceAttributions || [], sourceAttributionKey).filter((s) => isHttpUrl(s.url)).map((s) => ({ ...s, provenanceType: s.cited === "false" ? "search-only" : "uncited" }));
    const sourceMap = new Map();
    for (const s of additionalSources) {
      sourceMap.set(s.url, s);
    }
    for (const s of citedSources) {
      sourceMap.set(s.url, { ...(sourceMap.get(s.url) || {}), ...s, provenanceType: "cited" });
    }
    const sources = Array.from(sourceMap.values());
    const searchPlugins = (input?.searchPlugins || []).filter((p) => p && p.name);
    if (!queries.length && !sources.length && !searchPlugins.length) {
      return "";
    }
    const summaryBits = [];
    if (sources.length) {
      summaryBits.push(`${sources.length} ${sources.length === 1 ? "source" : "sources"}`);
    }
    if (queries.length) {
      summaryBits.push(`${queries.length} ${queries.length === 1 ? "query" : "queries"}`);
    }
    if (searchPlugins.length) {
      summaryBits.push(`${searchPlugins.length} ${searchPlugins.length === 1 ? "plugin" : "plugins"}`);
    }
    const lines = [
      "<details>",
      `<summary>Search provenance (${summaryBits.join(", ")})</summary>`,
      "",
    ];
    if (queries.length) {
      lines.push("**Queries:**", "");
      for (const q of queries) {
        lines.push(`- ${q}`);
      }
      lines.push("");
    }
    if (sources.length) {
      lines.push("**Sources:**", "");
      for (const s of sources) {
        const title = s.title || s.url;
        const domain = domainFromUrl(s.url);
        const bits = [];
        bits.push(s.provenanceType === "cited" ? "cited" : s.provenanceType || "uncited");
        if (domain) {
          bits.push(domain);
        }
        if (s.searchQuery) {
          bits.push(`query: ${s.searchQuery}`);
        }
        lines.push(`- [${title}](${s.url})${bits.length ? ` — ${bits.join(" · ")}` : ""}`);
      }
      lines.push("");
    }
    if (searchPlugins.length) {
      lines.push("**Web-search plugin activity:**", "");
      for (const p of searchPlugins) {
        lines.push(`- ${p.name}${p.count > 1 ? ` — ${p.count} records` : ""}`);
      }
      lines.push("");
    }
    lines.push("</details>");
    return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  function pluginDisplayName(info) {
    const id = info?.id || "";
    const source = info?.source || "";
    if (id && source) {
      return `${id} (${source})`;
    }
    return id || source || "Plugin";
  }

  function pluginFootnoteForMessage(msg) {
    const info = msg?.pluginInfo || {};
    const id = info.id;
    const source = info.source;
    const version = info.version;
    if (!id && !source && !version) {
      return "";
    }
    const bits = [];
    if (id) {
      bits.push(`Plugin: ${id}`);
    }
    if (source) {
      bits.push(`source: ${source}`);
    }
    if (version) {
      bits.push(`v${version}`);
    }
    return `_<sub>${bits.join(" · ")}</sub>_`;
  }

  function canonicalPluginId(msgOrInfo) {
    const info = msgOrInfo?.pluginInfo || msgOrInfo || {};
    return String(info.id || "").toLowerCase();
  }

  function isPyexecPluginMessage(msg) {
    return canonicalPluginId(msg) === "pyexec";
  }

  function isUpdateMemoryPluginMessage(msg) {
    return canonicalPluginId(msg) === "updatememory";
  }

  function isFileUploadPluginMessage(msg) {
    const id = canonicalPluginId(msg);
    return (
      id === "flux_fileupload_longfile_mapreduce_plugin" ||
      id === "flux_fileupload_fetch_plugin" ||
      id.includes("fileupload")
    );
  }

  function isClickUrlPluginMessage(msg) {
    return canonicalPluginId(msg) === "clickurl";
  }

  function isMultiBingWebSearchPluginMessage(msg) {
    return canonicalPluginId(msg) === "multibingwebsearch";
  }

  function isMarkdownHiddenTextOnlyPluginMessage(msg) {
    return (
      isClickUrlPluginMessage(msg) || isMultiBingWebSearchPluginMessage(msg)
    );
  }

  function isSearchProvenancePluginMessage(msg) {
    return isClickUrlPluginMessage(msg) || isMultiBingWebSearchPluginMessage(msg);
  }

  function isPythonCodeInterpreterPluginMessage(msg) {
    return canonicalPluginId(msg) === "pythoncodeinterpreter";
  }

  function pluginEvidenceFullText(value) {
    if (value == null) {
      return "";
    }
    const raw =
      typeof value === "string" ? value : JSON.stringify(value, null, 2);
    return String(raw || "")
      .replace(/\r\n/g, "\n")
      .trim();
  }

  function wrapPluginEvidenceLines(value, maxLen = 120) {
    const raw = String(value || "");
    return raw
      .split(/\r?\n/)
      .map((line) => {
        if (line.length <= maxLen) {
          return line;
        }
        const chunks = [];
        let rest = line;
        while (rest.length > maxLen) {
          let cut = Math.max(
            rest.lastIndexOf(", ", maxLen),
            rest.lastIndexOf("; ", maxLen),
            rest.lastIndexOf(" ", maxLen),
          );
          if (cut < Math.floor(maxLen * 0.55)) {
            cut = maxLen;
          }
          chunks.push(rest.slice(0, cut + 1).trimEnd());
          rest = rest.slice(cut + 1).trimStart();
        }
        if (rest) {
          chunks.push(rest);
        }
        return chunks.join("\n");
      })
      .join("\n");
  }

  function tryParsePluginJsonValue(text) {
    const raw = String(text || "").trim();
    if (!raw) {
      return { ok: false, value: null };
    }
    try {
      const parsed = JSON.parse(raw);
      if (typeof parsed === "string") {
        const inner = parsed.trim();
        if (inner && (inner.startsWith("{") || inner.startsWith("["))) {
          try {
            return { ok: true, value: JSON.parse(inner) };
          } catch {
            return { ok: true, value: parsed };
          }
        }
      }
      return { ok: true, value: parsed };
    } catch {
      return { ok: false, value: null };
    }
  }

  function normalizePluginJsonStringFields(value, depth = 0) {
    if (depth > 4) {
      return value;
    }
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (!trimmed || !(trimmed.startsWith("{") || trimmed.startsWith("["))) {
        return value;
      }
      try {
        return normalizePluginJsonStringFields(JSON.parse(trimmed), depth + 1);
      } catch {
        return value;
      }
    }
    if (Array.isArray(value)) {
      return value.map((item) =>
        normalizePluginJsonStringFields(item, depth + 1),
      );
    }
    if (value && typeof value === "object") {
      const out = {};
      for (const [key, item] of Object.entries(value)) {
        out[key] = normalizePluginJsonStringFields(item, depth + 1);
      }
      return out;
    }
    return value;
  }

  function prettyJsonForPluginEvidence(rawJson) {
    const parsed = tryParsePluginJsonValue(rawJson);
    if (!parsed.ok) {
      return "";
    }
    return JSON.stringify(
      normalizePluginJsonStringFields(parsed.value),
      null,
      2,
    );
  }

  function prettyPrintJsonFencesInPluginMarkdown(raw) {
    let changed = false;
    const rendered = String(raw || "").replace(
      /```json\s*\n([\s\S]*?)\n```/gi,
      (match, body) => {
        const pretty = prettyJsonForPluginEvidence(body);
        if (!pretty) {
          return match;
        }
        changed = true;
        return renderFencedBlock(pretty, "json");
      },
    );
    return changed ? rendered : "";
  }

  function readBalancedJsonCandidate(text, start) {
    const stack = [];
    let inString = false;
    let escaped = false;
    for (let i = start; i < text.length; i += 1) {
      const ch = text[i];
      if (inString) {
        if (escaped) {
          escaped = false;
        } else if (ch === "\\") {
          escaped = true;
        } else if (ch === '"') {
          inString = false;
        }
        continue;
      }
      if (ch === '"') {
        inString = true;
        continue;
      }
      if (ch === "{" || ch === "[") {
        stack.push(ch);
        continue;
      }
      if (ch === "}" || ch === "]") {
        const expected = ch === "}" ? "{" : "[";
        if (stack.pop() !== expected) {
          return "";
        }
        if (!stack.length) {
          return text.slice(start, i + 1);
        }
      }
    }
    return "";
  }

  function wrapperLooksTrivialJsonShell(text, candidate) {
    const wrapper = String(text || "")
      .replace(candidate, "")
      .replace(/```(?:json|javascript|js)?/gi, "")
      .replace(/`/g, "")
      .replace(/\b(json|result|output|response|data|payload|content)\s*:/gi, "")
      .replace(/^[\s>*#\-:]+|[\s>*#\-:]+$/g, "")
      .trim();
    return wrapper.length <= 80;
  }

  function tryExtractJsonFromMarkdownishText(text) {
    const raw = String(text || "").trim();
    const candidates = [];
    for (let i = 0; i < raw.length; i += 1) {
      if (raw[i] === "{" || raw[i] === "[") {
        const candidate = readBalancedJsonCandidate(raw, i);
        if (candidate) {
          candidates.push(candidate);
        }
      }
    }
    candidates.sort((a, b) => b.length - a.length);
    for (const candidate of candidates) {
      if (!wrapperLooksTrivialJsonShell(raw, candidate)) {
        continue;
      }
      const parsed = tryParsePluginJsonValue(candidate);
      if (parsed.ok) {
        return parsed;
      }
    }
    return { ok: false, value: null };
  }

  function pluginEvidencePrettyJson(value) {
    const raw = pluginEvidenceFullText(value);
    const direct = tryParsePluginJsonValue(raw);
    if (direct.ok) {
      return JSON.stringify(
        normalizePluginJsonStringFields(direct.value),
        null,
        2,
      );
    }
    const wrapped = tryExtractJsonFromMarkdownishText(raw);
    if (wrapped.ok) {
      return JSON.stringify(
        normalizePluginJsonStringFields(wrapped.value),
        null,
        2,
      );
    }
    return "";
  }

  function pluginEvidenceDisplayBlock(value, mode = "auto") {
    const raw = pluginEvidenceFullText(value);
    if (!raw) {
      return "";
    }
    if (mode === "python") {
      return renderFencedBlock(raw, "python");
    }
    if (mode === "wrapped-text") {
      return renderFencedBlock(wrapPluginEvidenceLines(raw), "text");
    }
    const prettyFencedMarkdown = prettyPrintJsonFencesInPluginMarkdown(raw);
    if (prettyFencedMarkdown) {
      return prettyFencedMarkdown;
    }
    const prettyJson = pluginEvidencePrettyJson(raw);
    if (prettyJson) {
      return renderFencedBlock(prettyJson, "json");
    }
    if (mode === "markdown") {
      return raw;
    }
    return renderFencedBlock(raw, "text");
  }

  function appendPluginEvidence(lines, label, value, mode = "auto") {
    const block = pluginEvidenceDisplayBlock(value, mode);
    if (block) {
      lines.push(`  - ${label}:`);
      lines.push(indentMarkdown(block, "    "));
    }
  }

  function appendPluginAuxiliaryEvidence(lines, msg) {
    if (!msg) {
      return;
    }
    appendPluginEvidence(
      lines,
      "filesAndLinksContextMetadata",
      msg.filesAndLinksContextMetadata,
      "auto",
    );
    appendPluginEvidence(
      lines,
      "messageAnnotations",
      msg.messageAnnotations,
      "auto",
    );
  }

  function sameToolRunScope(a, b) {
    if (!a || !b) {
      return false;
    }
    return (
      String(a.requestId || "") === String(b.requestId || "") &&
      String(a.turnCount ?? "") === String(b.turnCount ?? "")
    );
  }

  function pyexecMatchesToolRun(pyexecMsg, toolRunMsg, code) {
    if (
      !isPyexecPluginMessage(pyexecMsg) ||
      !sameToolRunScope(pyexecMsg, toolRunMsg)
    ) {
      return false;
    }
    const wanted = normalizeTextForCompare(code || "");
    if (!wanted) {
      return true;
    }
    const hidden = normalizeTextForCompare(pyexecMsg?.hiddenText || "");
    if (hidden && hidden === wanted) {
      return true;
    }
    const evidence = normalizeTextForCompare(
      `${pyexecMsg?.text || ""}\n${pyexecMsg?.invocation || ""}`,
    );
    return Boolean(evidence && evidence.includes(wanted.slice(0, 120)));
  }

  function pyexecPluginFootnoteForToolRun(toolRunMsg, allItems, code) {
    const matches = uniqBy(
      (allItems || []).filter((msg) =>
        pyexecMatchesToolRun(msg, toolRunMsg, code),
      ),
      pluginCarrierKey,
    );
    if (!matches.length) {
      return "";
    }
    const pluginName = pluginDisplayName(matches[0].pluginInfo);
    const recordWord = matches.length === 1 ? "record" : "records";
    return `_<sub>Plugin usage: ${pluginName} · ${matches.length} ${recordWord}</sub>_`;
  }

  function appendPluginFootnote(rendered, msg) {
    const body = rendered || "";
    const footnote = pluginFootnoteForMessage(msg);
    if (!body || !footnote) {
      return body;
    }
    return `${body}\n\n${footnote}`;
  }

  function pluginCarrierKey(msg) {
    return [
      msg?.messageId || "",
      msg?.requestId || "",
      msg?.turnCount ?? "",
      msg?.createdAt || msg?.timestamp || "",
      msg?.pluginInfo?.id || "",
      msg?.messageType || "",
      msg?.contentOrigin || "",
    ].join("|");
  }

  function fallbackToolRunMessages(items) {
    const generated = (items || [])
      .filter((m) => m.author === "bot" && m.messageType === "GeneratedCode")
      .slice(-1)[0];
    const internal = (items || [])
      .filter((m) => m.author === "bot" && m.messageType === "Internal")
      .slice(-1)[0];
    return [generated, internal].filter(Boolean);
  }

  function renderPluginMessageEvidence(msg, renderedSet) {
    const info = msg?.pluginInfo || {};
    const status = renderedSet?.has(msg)
      ? "rendered inline"
      : "not otherwise rendered";
    const bits = [
      `**${pluginDisplayName(info)}**`,
      `status: ${status}`,
      `messageType: ${msg?.messageType || "<missing>"}`,
      `contentOrigin: ${msg?.contentOrigin || "<missing>"}`,
    ];
    if (info.version) {
      bits.push(`pluginVersion: ${info.version}`);
    }
    if (msg?.createdAt || msg?.timestamp) {
      bits.push(`createdAt: ${msg.createdAt || msg.timestamp}`);
    }
    const lines = [`- ${bits.join(" · ")}`];

    if (isFileUploadPluginMessage(msg)) {
      appendPluginEvidence(lines, "text", msg?.text, "auto");
      return lines.join("\n");
    }

    if (isMarkdownHiddenTextOnlyPluginMessage(msg)) {
      appendPluginEvidence(lines, "hiddenText", msg?.hiddenText, "markdown");
      appendPluginAuxiliaryEvidence(lines, msg);
      return lines.join("\n");
    }

    if (isPythonCodeInterpreterPluginMessage(msg)) {
      appendPluginEvidence(
        lines,
        "invocation",
        msg?.invocation,
        "wrapped-text",
      );
      appendPluginEvidence(lines, "hiddenText", msg?.hiddenText, "python");
      appendPluginAuxiliaryEvidence(lines, msg);
      return lines.join("\n");
    }

    if (isUpdateMemoryPluginMessage(msg)) {
      appendPluginEvidence(lines, "text", msg?.text, "auto");
      appendPluginAuxiliaryEvidence(lines, msg);
      return lines.join("\n");
    }

    appendPluginEvidence(lines, "invocation", msg?.invocation, "auto");
    appendPluginEvidence(lines, "text", msg?.text, "auto");
    appendPluginAuxiliaryEvidence(lines, msg);
    return lines.join("\n");
  }


  function isLongFileMapReduceProgressMessage(msg) {
    const id = canonicalPluginId(msg);
    const mt = String(msg?.messageType || "");
    const co = String(msg?.contentOrigin || "");
    const text = typeof msg?.text === "string" ? msg.text.trim() : "";
    if (id !== "flux_fileupload_longfile_mapreduce_plugin" || !text) {
      return false;
    }
    return (mt === "InternalLoaderMessage" && co === "summarize-long-file-progress") || co === "summarize-file" || mt === "Internal";
  }

  function isLongFileMapReduceStatusOnlyMessage(msg) {
    const text = typeof msg?.text === "string" ? msg.text.trim() : "";
    return text === "Analyzing the file:";
  }

  function longFileMapReduceFileName(msg) {
    const fileCandidates = [];
    const addCandidate = (value) => {
      const text = String(value || "").trim();
      if (text) {
        fileCandidates.push(text);
      }
    };
    for (const file of msg?.filesAndLinksContextMetadata?.files || []) {
      addCandidate(file?.fileName || file?.annotatedFileName || file?.name || file?.text);
    }
    for (const annotation of msg?.messageAnnotations || []) {
      addCandidate(annotation?.text || annotation?.fileName || annotation?.annotatedFileName);
    }
    for (const value of [msg?.text, msg?.hiddenText]) {
      const raw = String(value || "").trim();
      if (!raw) {
        continue;
      }
      try {
        const parsed = JSON.parse(raw);
        addCandidate(parsed?.fileName || parsed?.FileName || parsed?.filename || parsed?.name);
      } catch {
        // Non-JSON loader text is expected for status-only records.
      }
    }
    return fileCandidates.find((name) => /\.(?:md|txt|json|zip|csv|xlsx?|docx?|pptx?|pdf)(?:\.txt)?$/i.test(name)) || fileCandidates[0] || "";
  }

  function renderLongFileMapReduceProgressDetails(messages) {
    const msgList = Array.isArray(messages) ? messages : [messages];
    const fileNames = uniqBy(msgList.map(longFileMapReduceFileName).filter(Boolean), (x) => x);
    if (!fileNames.length && msgList.every(isLongFileMapReduceStatusOnlyMessage)) {
      return "";
    }
    const footnotes = uniqBy(msgList.map(pluginFootnoteForMessage).filter(Boolean), (x) => x);
    const lines = ["<details>", "<summary>Analyzing the file:</summary>", ""];
    if (fileNames.length) {
      for (const fileName of fileNames) {
        lines.push(`- ${escapeMarkdown(fileName)}`);
      }
    } else {
      lines.push("- File analysis in progress");
    }
    for (const footnote of footnotes) {
      lines.push("", footnote);
    }
    lines.push("</details>");
    return lines.join("\n");
  }

  function renderLongFileMapReduceProgressSections(items) {
    const progressMessages = (items || []).filter(isLongFileMapReduceProgressMessage);
    const evidenceMessages = progressMessages.filter((msg) => !isLongFileMapReduceStatusOnlyMessage(msg));
    const selected = evidenceMessages.length ? evidenceMessages : progressMessages;
    const rendered = renderLongFileMapReduceProgressDetails(selected);
    return rendered ? [{ msgs: selected, msg: selected[0], rendered }] : [];
  }
  function renderPluginProvenanceSection(items, renderedSet) {
    const pluginMessages = uniqBy(
      (items || []).filter(
        (m) =>
          m?.pluginInfo &&
          (m.pluginInfo.id || m.pluginInfo.source) &&
          !isPyexecPluginMessage(m) &&
          !isSearchProvenancePluginMessage(m) &&
          !isFetchFileMessage(m) &&
          !isLongFileMapReduceProgressMessage(m),
      ),
      pluginCarrierKey,
    );
    if (!pluginMessages.length) {
      return "";
    }
    const renderedCount = pluginMessages.filter((m) =>
      renderedSet?.has(m),
    ).length;
    const orphanCount = pluginMessages.length - renderedCount;
    const pluginNames = uniqBy(
      pluginMessages.map((m) => pluginDisplayName(m.pluginInfo)),
      (x) => x,
    )
      .slice(0, 4)
      .join(", ");
    const headerBits = [`Plugin usage (${pluginMessages.length})`];
    if (orphanCount) {
      headerBits.push(`${orphanCount} internal/orphaned`);
    }
    if (pluginNames) {
      headerBits.push(pluginNames);
    }
    const lines = [
      "<details>",
      `<summary>${escapeHtml(`Plugin usage (${pluginMessages.length})`)}</summary>`,
      "",
    ];
    lines.push(`**${headerBits.join(" — ")}**`, "");
    for (const msg of pluginMessages) {
      lines.push(renderPluginMessageEvidence(msg, renderedSet), "");
    }
    lines.push("</details>");
    return lines
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function normalizeQueryCandidate(query) {
    const q = String(query || "")
      .replace(/\s+/g, " ")
      .trim();
    if (!q || q.length > 500) {
      return "";
    }
    return q;
  }

  function pushUniqueQuery(out, query) {
    const q = normalizeQueryCandidate(query);
    if (!q) {
      return;
    }
    const key = q.toLowerCase();
    if (!out.has(key)) {
      out.set(key, q);
    }
  }

  function tryParseJsonValue(value) {
    if (value == null) {
      return null;
    }
    if (typeof value !== "string") {
      return value;
    }
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }
    try {
      return JSON.parse(trimmed);
    } catch {
      return null;
    }
  }

  function collectSearchQueriesFromInvocation(invocation, out) {
    const values = Array.isArray(invocation) ? invocation : [invocation];
    for (const raw of values) {
      const parsed = tryParseJsonValue(raw);
      if (!parsed || typeof parsed !== "object") {
        continue;
      }
      const fn = parsed.function || parsed;
      if (fn?.name && fn.name !== "search_web") {
        continue;
      }
      const args = tryParseJsonValue(fn?.arguments) || fn?.arguments;
      if (Array.isArray(args?.queries)) {
        for (const q of args.queries) {
          pushUniqueQuery(out, q);
        }
      }
    }
  }

  function collectSearchQueriesFromPayloadText(text, out) {
    const parsed = tryParseJsonValue(text);
    if (Array.isArray(parsed)) {
      for (const item of parsed) {
        if (item?.query) {
          pushUniqueQuery(out, item.query);
        }
      }
      return;
    }
    if (parsed?.query) {
      pushUniqueQuery(out, parsed.query);
    }
  }

  function extractSearchQueriesFromMessage(msg) {
    const out = new Map();
    if (!msg || msg.messageType !== "InternalSearchResult") {
      return [];
    }
    collectSearchQueriesFromPayloadText(msg.text, out);
    collectSearchQueriesFromInvocation(msg.invocation, out);
    return Array.from(out.values());
  }

  function extractSearchQueriesFromItems(items) {
    const out = new Map();
    for (const msg of items || []) {
      for (const q of extractSearchQueriesFromMessage(msg)) {
        pushUniqueQuery(out, q);
      }
    }
    return Array.from(out.values());
  }

  function renderQueriesDetails(queries) {
    const items = uniqBy(queries || [], (q) =>
      normalizeQueryCandidate(q).toLowerCase(),
    ).filter(Boolean);
    if (!items.length) {
      return "";
    }
    const lines = [
      "<details>",
      `<summary>Queries (${items.length})</summary>`,
      "",
    ];
    for (const q of items) {
      lines.push(`- ${q}`);
    }
    lines.push("", "</details>");
    return lines.join("\n");
  }

  const KNOWN_CONTEXT_AUTHORS = new Set([
    "system",
    "FileUpload App",
    "Attached Content App",
  ]);

  function isUnclassifiedContextMessage(msg) {
    return !!(
      msg &&
      msg.messageType === "Context" &&
      !KNOWN_CONTEXT_AUTHORS.has(msg.author || "")
    );
  }

  function renderUnclassifiedContextSection(items) {
    if (!(items || []).length) {
      return "";
    }
    return [
      "<details>",
      `<summary>Unclassified context (${items.length})</summary>`,
      "",
      renderFencedBlock(JSON.stringify(items, null, 2), "json"),
      "</details>",
    ].join("\n");
  }

  function hasUnknownRecordEvidence(msg) {
    if (!msg) {
      return false;
    }
    if (typeof msg.text === "string" && msg.text.trim()) {
      return true;
    }
    if (typeof msg.hiddenText === "string" && msg.hiddenText.trim()) {
      return true;
    }
    if (typeof msg.invocation === "string" && msg.invocation.trim()) {
      return true;
    }
    if (Array.isArray(msg.adaptiveCards) && msg.adaptiveCards.length) {
      return true;
    }
    if (
      Array.isArray(msg.sourceAttributions) &&
      msg.sourceAttributions.length
    ) {
      return true;
    }
    if (msg.pluginInfo && (msg.pluginInfo.id || msg.pluginInfo.source)) {
      return true;
    }
    if (msg.filesAndLinksContextMetadata) {
      return true;
    }
    if (
      Array.isArray(msg.messageAnnotations) &&
      msg.messageAnnotations.length
    ) {
      return true;
    }
    return false;
  }

  function isNoContentReturnedMessage(msg) {
    const text = normalizeTextForCompare(msg?.text || "");
    const hidden = normalizeTextForCompare(msg?.hiddenText || "");
    return (
      text === "no content returned" &&
      (!hidden || hidden === "no content returned")
    );
  }

  function isFetchFileMessage(msg) {
    return (
      String(msg?.contentOrigin || "")
        .trim()
        .toLowerCase() === "fetch-file"
    );
  }

  // Known attachment context origins: UserAttachedContentContext, fileupload-rewrite.
  // Known attachment context authors: Attached Content App, FileUpload App.
  function isKnownLowValueAttachmentContextMessage(msg) {
    if (!msg) {
      return false;
    }
    const mtKey = String(msg.messageType ?? "").trim().toLowerCase();
    const coKey = String(msg.contentOrigin ?? "").trim().toLowerCase();
    const authorKey = String(msg.author ?? "").trim().toLowerCase();
    if (mtKey !== "context") {
      return false;
    }
    if (coKey === "userattachedcontentcontext" || coKey === "fileupload-rewrite") {
      return true;
    }
    if (authorKey === "attached content app" || authorKey === "fileupload app") {
      return true;
    }
    return false;
  }

  function isKnownInternalAttachedContentInfoMessage(msg) {
    if (!msg) {
      return false;
    }
    const authorKey = String(msg.author ?? "").trim().toLowerCase();
    const mtKey = String(msg.messageType ?? "").trim().toLowerCase();
    const coKey = String(msg.contentOrigin ?? "").trim().toLowerCase();
    return authorKey === "system" && mtKey === "internalattachedcontentinfo" && coKey === "deepleo";
  }

  // Known tool/attachment plumbing observed in older exports.
  function isKnownInternalCodeInterpreterMessage(msg) {
    if (!msg) {
      return false;
    }
    const authorKey = String(msg.author ?? "").trim().toLowerCase();
    const mtKey = String(msg.messageType ?? "").trim().toLowerCase();
    const coKey = String(msg.contentOrigin ?? "").trim().toLowerCase();
    return authorKey === "bot" && mtKey === "internal" && coKey === "codeinterpreter";
  }

  function isKnownSystemAttachmentActionMessage(msg) {
    if (!msg) {
      return false;
    }
    const authorKey = String(msg.author ?? "").trim().toLowerCase();
    const mtKey = String(msg.messageType ?? "").trim().toLowerCase();
    const coKey = String(msg.contentOrigin ?? "").trim();
    return authorKey === "system" && mtKey === "attachmentaction" && coKey === "";
  }

  function isKnownModelSelectorMessage(msg) {
    const authorKey = String(msg?.author ?? "").trim().toLowerCase();
    const mtKey = String(msg?.messageType ?? "").trim().toLowerCase();
    const coKey = String(msg?.contentOrigin ?? "").trim().toLowerCase();
    return authorKey === "system" && mtKey === "internal" && coKey === "modelselector";
  }

  function isKnownNoiseMessage(msg) {
    if (!msg) {
      return true;
    }
    const mt = String(msg.messageType ?? "").trim();
    const co = String(msg.contentOrigin ?? "").trim();
    const mtKey = mt.toLowerCase();
    const coKey = co.toLowerCase();
    const status = normalizeStatusToken(msg.text || "");
    if (
      mt === "Suggestion" ||
      mt === "InternalSuggestions" ||
      mt === "RenderCardRequest" ||
      mt === "GenerateContentQuery" ||
      mt === "AdsQuery"
    ) {
      return true;
    }
    if (mt === "MemoryUpdate" || mt === "InternalUserSummary") {
      return true;
    }
    if (isFetchFileMessage(msg)) {
      return true;
    }
    if (isKnownLowValueAttachmentContextMessage(msg)) {
      return true;
    }
    if (isKnownInternalAttachedContentInfoMessage(msg)) {
      return true;
    }
    if (isKnownInternalCodeInterpreterMessage(msg)) {
      return true;
    }
    if (isKnownSystemAttachmentActionMessage(msg)) {
      return true;
    }
    if (isKnownModelSelectorMessage(msg)) {
      return true;
    }
    if (
      mtKey === "invokeaction" &&
      coKey === "deepleo" &&
      isNoContentReturnedMessage(msg)
    ) {
      return true;
    }
    if (
      mt === "Progress" &&
      co !== "ChainOfThoughtSummary" &&
      status !== "Coding and executing"
    ) {
      return true;
    }
    if (
      mt === "InternalSearchResult" ||
      mt === "InternalSearchQuery" ||
      co === "citation-formatter"
    ) {
      return true;
    }
    if (
      looksLikeRawSearchResultPayload(String(msg.text || msg.hiddenText || ""))
    ) {
      return true;
    }
    return false;
  }

  function unknownRecordSummary(msg) {
    const author = msg?.author || "unknown-author";
    const mt = msg?.messageType || "missing-messageType";
    const co = msg?.contentOrigin || "missing-contentOrigin";
    return `${author} · ${mt} · ${co}`;
  }

  function renderUnknownRecordsSection(items, renderedSet) {
    if (!settings.includeUnclassifiedRecords) {
      return "";
    }
    const records = (items || []).filter((msg) => {
      if (!msg || renderedSet.has(msg)) {
        return false;
      }
      if (msg.author === "user") {
        return false;
      }
      if (isUnclassifiedContextMessage(msg)) {
        return false;
      }
      if (
        msg.pluginInfo &&
        (msg.pluginInfo.id || msg.pluginInfo.source)
      ) {
        return false;
      }
      if (isKnownNoiseMessage(msg)) {
        return false;
      }
      return hasUnknownRecordEvidence(msg);
    });
    if (!records.length) {
      return "";
    }
    const lines = [
      "<details>",
      `<summary>Unknown/unclassified records (${records.length})</summary>`,
      "",
    ];
    for (const msg of records) {
      lines.push(
        `- ${unknownRecordSummary(msg)}`,
        "",
        renderFencedBlock(JSON.stringify(msg, null, 2), "json"),
        "",
      );
    }
    lines.push("</details>");
    return lines
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function renderHiddenTextDetails(hiddenText) {
    if (!hiddenText) return "";
    return ["<details>", "<summary>hiddenText</summary>", "", renderFencedBlock(hiddenText, "text"), "</details>"].join("\n");
  }

  function shouldRenderHiddenText(msg, primary, hiddenText) {
    if (!hiddenText) return false;
    if (hiddenTextLooksLikeSearchInvocation(hiddenText)) return false;
    const mt = msg?.messageType ?? "";
    if (mt === "GeneratedCode" || mt === "Internal" || mt === "InvokeAction") return false;
    const hidden = normalizeTextForCompare(hiddenText);
    const visible = normalizeTextForCompare(primary);
    if (!hidden) return false;
    if (hidden === visible) return false;
    return true;
  }

  function stripMarkdownCodeFence(text) {
    if (!text) return "";
    return String(text)
      .replace(/^```[a-zA-Z0-9_-]*\n?/, "")
      .replace(/\n?```$/, "")
      .trim();
  }

  function parseInternalExecutionPayload(msg) {
    if (!msg) return null;
    const candidates = [msg?.text, msg?.hiddenText];
    for (const candidate of candidates) {
      if (!candidate || typeof candidate !== "string") {
        continue;
      }
      try {
        const parsed = JSON.parse(candidate);
        if (parsed && typeof parsed === "object") return parsed;
      } catch {
        // ignore
      }
    }
    return null;
  }

  function renderToolRunFallback(items) {
    const generated = items
      .filter((m) => m.author === "bot" && m.messageType === "GeneratedCode")
      .slice(-1)[0];
    const internal = items
      .filter((m) => m.author === "bot" && m.messageType === "Internal")
      .slice(-1)[0];
    if (!generated && !internal) return "";

    let code = "";
    if (generated) {
      const generatedCardText = extractAdaptiveCardsMarkdown(
        generated.adaptiveCards || [],
      );
      code = stripMarkdownCodeFence(
        generated.hiddenText || generatedCardText || generated.text || "",
      );
    }

    const payload = parseInternalExecutionPayload(internal);
    let status = "";
    let output = "";
    let error = "";

    if (payload) {
      status = payload.status || "";
      output = payload.stdout || payload.result || "";
      error = payload.stderr || "";
      if (!code) code = stripMarkdownCodeFence(payload.executedCode || "");
    } else if (internal) {
      const mapped =
        adaptiveBodyToFieldMap(internal.adaptiveCards || []).fields || {};
      status = mapped.Status || "";
      output = mapped.ExecutionOutput || "";
      error = mapped.ExecutionError || "";
      if (!code) {
        code = stripMarkdownCodeFence(
          mapped.Content || internal.hiddenText || "",
        );
      }
    }

    if (!code && !output && !error) return "";

    const lines = [
      "<details>",
      `<summary>Tool run${status ? ` — ${status}` : " — fallback"}</summary>`,
      "",
    ];
    if (code) lines.push(renderFencedBlock(code, guessFenceLangFromText(code)), "");
    if (status) lines.push(`**Status:** ${status}`, "");
    if (output) lines.push(renderFencedBlock(output, "text"), "");
    if (error) lines.push(renderFencedBlock(error, "text"), "");
    const pluginFootnote =
      pluginFootnoteForMessage(generated) || pluginFootnoteForMessage(internal);
    if (pluginFootnote) {
      lines.push(pluginFootnote, "");
    }
    lines.push("</details>");
    return lines
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function isMainAnswerCandidate(msg) {
    if (!msg || msg.author !== "bot") return false;
    if (isLongFileMapReduceProgressMessage(msg)) return false;
    const mt = msg.messageType ?? "";
    const co = msg.contentOrigin ?? "";
    if (mt === "MemoryUpdate" || mt === "InternalUserSummary") return false;
    if (
      mt === "InternalSearchResult" ||
      mt === "InternalSearchQuery" ||
      co === "citation-formatter"
    ) {
      return false;
    }
    if (mt === "GeneratedCode" || mt === "Internal" || mt === "InvokeAction") {
      return false;
    }
    if (
      mt === "Suggestion" ||
      mt === "RenderCardRequest" ||
      mt === "GenerateContentQuery" ||
      mt === "AdsQuery"
    ) {
      return false;
    }
    if (mt === "Progress") return false;
    const cardText = extractAdaptiveCardsMarkdown(msg.adaptiveCards || []);
    const visibleText =
      cardText ||
      (typeof msg.text === "string" ? msg.text.trim() : "") ||
      (typeof msg.hiddenText === "string" ? msg.hiddenText.trim() : "");
    if (looksLikeRawSearchResultPayload(visibleText)) return false;
    return !!(
      cardText ||
      (typeof msg.text === "string" && msg.text.trim()) ||
      (msg.sourceAttributions && msg.sourceAttributions.length)
    );
  }

  function extractMainAnswerPrimaryText(msg) {
    let primary = extractAdaptiveCardsMarkdown(msg?.adaptiveCards || []);
    if (!primary) primary = typeof msg?.text === "string" ? msg.text.trim() : "";
    if (!primary) primary = typeof msg?.hiddenText === "string" ? msg.hiddenText.trim() : "";
    primary = collapseRepeatedTrailingLinks(primary || "");
    primary = dropNoContentReturned(primary || "");
    primary = collapseDuplicateStatusLines(primary || "");
    primary = formatReasoningSteps(primary || "");
    primary = italicizeSystemishOutsideFences(primary || "");
    primary = normalizeLatexDisplayMath(primary);
    primary = normalizeInlineLatexMath(primary);
    primary = repairSplitMarkdownHeadingContinuations(primary);
    primary = normalizeMarkdownBlockBoundaries(primary);
    primary = normalizeCopilotBodyMarkdown(primary);
    primary = hardenNestedMarkdownFenceExamples(primary);
    primary = promoteOuterMarkdownExampleFences(primary);
    primary = closeShortFenceClosersBeforeMarkdownBoundaries(primary);
    primary = improveDisplayMathSpacing(primary);
    if (looksLikeRawSearchResultPayload(primary)) return "";
    return primary || "";
  }

  function renderMainAnswerMessage(msg, primary) {
    let rendered = primary || extractMainAnswerPrimaryText(msg);
    const hidden =
      typeof msg?.hiddenText === "string" ? msg.hiddenText.trim() : "";
    if (rendered && shouldRenderHiddenText(msg, rendered, hidden)) {
      rendered += `\n\n${renderHiddenTextDetails(hidden)}`;
    }
    return appendPluginFootnote(rendered || "", msg);
  }

  function isFallbackMainAnswerCandidate(msg) {
    if (!msg || msg.author !== "bot") return false;
    if (isLongFileMapReduceProgressMessage(msg)) return false;
    const status = normalizeStatusToken(msg.text || "");
    if (
      status === "No content returned" ||
      status === "Coding and executing" ||
      status === "Reviewing the data..." ||
      status === "Thinking..."
    ) {
      return false;
    }
    const mt = msg.messageType ?? "";
    const co = msg.contentOrigin ?? "";
    if (mt === "MemoryUpdate" || mt === "InternalUserSummary") return false;
    if (
      mt === "InternalSearchResult" ||
      mt === "InternalSearchQuery" ||
      co === "citation-formatter"
    ) {
      return false;
    }
    if (mt === "GeneratedCode" || mt === "Internal" || mt === "InvokeAction") {
      return false;
    }
    if (
      mt === "Suggestion" ||
      mt === "RenderCardRequest" ||
      mt === "GenerateContentQuery" ||
      mt === "AdsQuery"
    ) {
      return false;
    }
    if (mt === "Progress") return false;
    const cardText = extractAdaptiveCardsMarkdown(msg.adaptiveCards || []);
    const visibleText =
      cardText ||
      (typeof msg.text === "string" ? msg.text.trim() : "") ||
      (typeof msg.hiddenText === "string" ? msg.hiddenText.trim() : "");
    if (looksLikeRawSearchResultPayload(visibleText)) return false;
    return !!(cardText || (typeof msg.text === "string" && msg.text.trim()));
  }

  function pickMainAnswerMessages(items) {
    const primaryCandidates = (items || []).filter(isMainAnswerCandidate);
    const candidates = primaryCandidates.length
      ? primaryCandidates
      : (items || []).filter(isFallbackMainAnswerCandidate);
    const out = [];
    const seen = new Set();
    for (const msg of candidates) {
      const primary = extractMainAnswerPrimaryText(msg);
      const key = normalizeTextForCompare(primary).replace(/\s+/g, " ");
      if (!key || seen.has(key)) {
        continue;
      }
      seen.add(key);
      const rendered = renderMainAnswerMessage(msg, primary);
      if (rendered) out.push({ msg, rendered });
    }
    return out;
  }

  function messageTimestampMs(msg) {
    const raw = msg?.createdAt || msg?.timestamp || "";
    const ms = raw ? new Date(raw).getTime() : Number.NaN;
    return Number.isFinite(ms) ? ms : Number.MAX_SAFE_INTEGER;
  }

  function earliestMessageTimestampMs(messages) {
    const values = (messages || []).map((msg) => messageTimestampMs(msg)).filter((ms) => Number.isFinite(ms));
    return values.length ? Math.min(...values) : Number.MAX_SAFE_INTEGER;
  }

  function joinCopilotSections(sections) {
    return (sections || []).join("\n\n\n").replace(/<\/details>\n+(?=(?!<details>)\S)/g, "</details>\n\n<br>\n\n").trim();
  }

  function composeBotTurn(items) {
    const reasoningItems = items.filter(
      (m) =>
        m.author === "bot" &&
        m.messageType === "Progress" &&
        m.contentOrigin === "ChainOfThoughtSummary",
    );
    const genericProgressItems = items.filter(isMeaningfulGenericProgressMessage);
    const toolRunItems = items.filter(isCanonicalToolRunProgressMessage);
    const mainAnswers = pickMainAnswerMessages(items);
    const sections = [];
    const processEntries = [];
    const renderedSet = new Set();
    const citations = uniqBy(
      items.flatMap((m) => sourceAttributionsToCitations(m)),
      (x) => x.url,
    );
    const additionalSourceAttributions = sourceAttributionsToAdditionalItems(
      items,
      citations,
    );
    const queries = extractSearchQueriesFromItems(items);
    const searchPlugins = searchPluginMessagesToSummaries(items);
    const addProcessEntry = (entry, msgs) => {
      if (!entry || !entry.rendered) {
        return;
      }
      processEntries.push({
        ...entry,
        sortAt: Number.isFinite(entry.sortAt) ? entry.sortAt : Number.MAX_SAFE_INTEGER,
      });
      for (const msg of msgs || []) {
        if (msg) {
          renderedSet.add(msg);
        }
      }
    };

    const seenProgressText = new Set();
    const addTextProgressEntry = (msg) => {
      const rendered = normalizeReasoningText(msg?.text || "");
      if (!rendered) {
        return;
      }
      const key = rendered.toLowerCase();
      if (seenProgressText.has(key)) {
        renderedSet.add(msg);
        return;
      }
      seenProgressText.add(key);
      addProcessEntry(
        { kind: "text", rendered, sortAt: messageTimestampMs(msg) },
        [msg],
      );
    };

    for (const msg of reasoningItems) {
      addTextProgressEntry(msg);
    }
    for (const msg of genericProgressItems) {
      addTextProgressEntry(msg);
    }

    const renderedToolRuns = toolRunItems
      .map((msg) => ({ msg, rendered: renderToolRunSection(msg, items) }))
      .filter((x) => x.rendered);
    if (renderedToolRuns.length) {
      for (const toolRun of renderedToolRuns) {
        addProcessEntry(
          { kind: "detail", rendered: toolRun.rendered, sortAt: messageTimestampMs(toolRun.msg) },
          [toolRun.msg],
        );
      }
    } else {
      const fallbackToolRun = renderToolRunFallback(items);
      if (fallbackToolRun) {
        const fallbackMsgs = fallbackToolRunMessages(items);
        addProcessEntry(
          { kind: "detail", rendered: fallbackToolRun, sortAt: earliestMessageTimestampMs(fallbackMsgs) },
          fallbackMsgs,
        );
      }
    }

    const loaderProgressSections = renderLongFileMapReduceProgressSections(items);
    for (const loaderProgress of loaderProgressSections) {
      const loaderMsgs = loaderProgress.msgs || [loaderProgress.msg];
      addProcessEntry(
        { kind: "detail", rendered: loaderProgress.rendered, sortAt: earliestMessageTimestampMs(loaderMsgs) },
        loaderMsgs,
      );
    }

    const processTimeline = renderProcessTimelineSection(processEntries);
    if (processTimeline) {
      sections.push(processTimeline);
    }

    for (const answer of mainAnswers) {
      if (answer.rendered) {
        sections.push(answer.rendered);
        renderedSet.add(answer.msg);
      }
    }

    const pluginProvenance = renderPluginProvenanceSection(items, renderedSet);
    if (pluginProvenance) {
      sections.push(pluginProvenance);
    }

    const unknownRecords = renderUnknownRecordsSection(items, renderedSet);
    if (unknownRecords) {
      sections.push(unknownRecords);
    }

    if (!sections.length && !citations.length) {
      return null;
    }
    return {
      role: "Copilot",
      createdAt:
        mainAnswers[0]?.msg?.createdAt ??
        items.find((m) => m.author === "bot")?.createdAt ??
        null,
      turnCount: items.find((m) => m.turnCount != null)?.turnCount ?? null,
      sourceCount: citations.length,
      text: joinCopilotSections(sections),
      links: [],
      images: [],
      citations,
      additionalSourceAttributions,
      queries,
      searchPlugins,
    };
  }

  // Active readable pipeline (v0.1.23):
  // 1) group messages by turn
  // 2) render user prompt block
  // 3) compose Copilot lane sections (all distinct main answers + reasoning + tool runs)
  // 4) attach deduped citations / links / images
  function composeTurnAwareTurns(conversationJson) {
    // Keep all message records through grouping; known-noise exclusions happen in renderer lanes.
    const messages = (conversationJson?.messages || []).filter((m) => !!m);

    const ordered = messages.slice().sort((a, b) => {
      const at = a.createdAt ? new Date(a.createdAt).getTime() : 0;
      const bt = b.createdAt ? new Date(b.createdAt).getTime() : 0;
      return at - bt;
    });

    const groups = new Map();
    for (const msg of ordered) {
      const key =
        msg.turnCount ??
        `${msg.requestId || ""}:${msg.createdAt || msg.timestamp || Math.random()}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(msg);
    }

    const turns = [];
    for (const items of groups.values()) {
      const userItems = items.filter(
        (m) =>
          m.author === "user" && typeof m.text === "string" && m.text.trim(),
      );
      const seen = new Set();
      const parts = [];
      for (const m of userItems) {
        const t = m.text.replace(/\r\n/g, "\n").trim();
        const k = normalizeTextForCompare(t);
        if (t && k && !seen.has(k)) {
          seen.add(k);
          parts.push(t);
        }
      }
      const userText = parts.join("\n\n");
      const uploadedFiles = extractUploadedFilesFromTurn(items);
      const unclassifiedContextItems = items.filter(
        isUnclassifiedContextMessage,
      );
      if (userText || uploadedFiles.length) {
        turns.push({
          role: "User",
          createdAt: userItems[0]?.createdAt || null,
          turnCount: items[0]?.turnCount ?? null,
          sourceCount: 0,
          text: userText,
          uploadedFiles,
          links: [],
          images: [],
          citations: [],
        });
      }
      if (unclassifiedContextItems.length) {
        turns.push({
          role: "Context",
          createdAt: unclassifiedContextItems[0]?.createdAt || null,
          turnCount: items[0]?.turnCount ?? null,
          sourceCount: 0,
          text: renderUnclassifiedContextSection(unclassifiedContextItems),
          uploadedFiles: [],
          links: [],
          images: [],
          citations: [],
          queries: [],
        });
      }
      const botTurn = composeBotTurn(items);
      if (botTurn) turns.push(botTurn);
    }
    return turns;
  }

  function toMarkdownCardFirst(conversationJson, exportedAt = new Date().toISOString()) {
    const title = conversationJson?.chatName || "M365 Copilot Chat";
    const exported = exportedAt;
    const srcUrl = location.href;
    const turns = composeTurnAwareTurns(conversationJson);
    const groups = groupContiguous(turns);
    const lines = [];
    lines.push(`# ${title}`);
    lines.push("");
    lines.push(`- Exported: ${exported}`);
    lines.push(`- Source: ${srcUrl}`);
    if (conversationJson?.conversationId) lines.push(`- ConversationId: ${conversationJson.conversationId}`);
    lines.push(`- ExporterVersion: ${SCRIPT_VERSION}`);
    lines.push("");
    lines.push("---");
    lines.push("");
    for (let i = 0; i < groups.length; i++) {
      if (i > 0) lines.push(spacer());
      lines.push(renderGroup(groups[i]));
    }
    const full = lines.join("\n");

    // Grounding / searches used (v0.1.29)
    const allCitations = groups
      .flatMap((g) => g.items || [])
      .flatMap((x) => x.citations || []);
    if (allCitations.length) {
      lines.push("");
      lines.push("<details>");
      lines.push("<summary>Grounding / searches used</summary>");
      lines.push("");
      for (const c of allCitations) lines.push(`- ${c.url || ""}`);
      lines.push("</details>");
    }

    return (
      dropNoContentReturned(full)
        .replace(/\n{6,}/g, "\n\n\n\n")
        .trim() + "\n"
    );
  }

  function toRawJsonMarkdown(conversationJson, exportedAt = new Date().toISOString()) {
    const title = conversationJson?.chatName || "M365 Copilot Chat";
    const exported = exportedAt;
    const srcUrl = location.href;
    const lines = [];
    lines.push(`# ${title}`);
    lines.push("");
    lines.push(`- Exported: ${exported}`);
    lines.push(`- Source: ${srcUrl}`);
    if (conversationJson?.conversationId) lines.push(`- ConversationId: ${conversationJson.conversationId}`);
    lines.push(`- ExporterVersion: ${SCRIPT_VERSION}`);
    lines.push("");
    lines.push("---");
    lines.push("");
    lines.push(renderFencedBlock(JSON.stringify(conversationJson, null, 2), "json"));
    return lines.join("\n") + "\n";
  }

  // ---------------------------------------------------------------------------
  // MSAL token extraction
  // ---------------------------------------------------------------------------

  const getCookie = (key) =>
    document.cookie.match(`(^|;)\\s*${key}\\s*=\\s*([^;]+)`)?.pop() || "";

  function base64DecToArr(base64String) {
    let s = base64String.replace(/-/g, "+").replace(/_/g, "/");
    switch (s.length % 4) {
      case 2:
        s += "==";
        break;
      case 3:
        s += "=";
        break;
      default:
        break;
    }
    const bin = atob(s);
    return Uint8Array.from(bin, (c) => c.codePointAt(0) || 0);
  }

  function toArrayBuffer(bufferLike) {
    return Uint8Array.from(bufferLike).buffer;
  }

  async function deriveKey(baseKey, nonce, context) {
    return crypto.subtle.deriveKey(
      {
        name: "HKDF",
        salt: toArrayBuffer(nonce),
        hash: "SHA-256",
        info: new TextEncoder().encode(context),
      },
      baseKey,
      { name: "AES-GCM", length: 256 },
      false,
      ["encrypt", "decrypt"],
    );
  }

  async function decryptPayload(baseKey, nonce, context, encryptedData) {
    const encoded = base64DecToArr(encryptedData);
    const derived = await deriveKey(baseKey, base64DecToArr(nonce), context);
    const decrypted = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: new Uint8Array(12) },
      derived,
      toArrayBuffer(encoded),
    );
    return new TextDecoder().decode(decrypted);
  }

  async function getEncryptionCookie() {
    const raw = decodeURIComponent(getCookie("msal.cache.encryption"));
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch {
      throw new Error("Failed to parse msal.cache.encryption cookie");
    }
    if (!parsed?.key || !parsed?.id) {
      throw new Error("No encryption cookie found");
    }
    return {
      id: parsed.id,
      key: await crypto.subtle.importKey(
        "raw",
        toArrayBuffer(base64DecToArr(parsed.key)),
        "HKDF",
        false,
        ["deriveKey"],
      ),
    };
  }

  function getMsalIds() {
    const clientId = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";

    function walk(node, seen = new WeakSet(), depth = 0) {
      if (!node || typeof node !== "object" || depth > 12) return null;
      if (seen.has(node)) return null;
      seen.add(node);

      const objectId = node.objectId || node.oid;
      const tenantId = node.tenantId || node.tid || node.realm;
      const userPrincipalName =
        node.userPrincipalName ||
        node.upn ||
        node.preferred_username ||
        node.username ||
        node.email;
      if (objectId && tenantId) {
        return {
          localAccountId: objectId,
          tenantId,
          homeAccountId: `${objectId}.${tenantId}`,
          userPrincipalName: userPrincipalName || null,
          clientId,
        };
      }

      if (Array.isArray(node)) {
        for (const item of node) {
          const found = walk(item, seen, depth + 1);
          if (found) return found;
        }
        return null;
      }

      for (const key of Object.keys(node)) {
        try {
          const found = walk(node[key], seen, depth + 1);
          if (found) return found;
        } catch {
          // ignore access errors
        }
      }
      return null;
    }

    // Modern Copilot path: router hydration state.
    try {
      const hydration = window.__staticRouterHydrationData;
      const found = walk(hydration);
      if (found) return found;
    } catch {
      // ignore and continue
    }

    // Legacy path: older builds exposed identity JSON in the DOM.
    const el = document.getElementById("identity");
    if (el?.textContent) {
      try {
        const { objectId, tenantId, userPrincipalName } = JSON.parse(
          el.textContent,
        );
        if (objectId && tenantId) {
          return {
            localAccountId: objectId,
            tenantId,
            homeAccountId: `${objectId}.${tenantId}`,
            userPrincipalName: userPrincipalName || null,
            clientId,
          };
        }
      } catch {
        // continue to token-based fallback
      }
    }

    return { clientId };
  }

  function decodeJwtPayload(token) {
    try {
      const parts = String(token || "").split(".");
      if (parts.length < 2) return null;
      let payload = parts[1].replace(/-/g, "+").replace(/_/g, "/");
      while (payload.length % 4) payload += "=";
      return JSON.parse(atob(payload));
    } catch {
      return null;
    }
  }

  async function getAccessToken(msalIds) {
    const cookie = await getEncryptionCookie();
    const clientId =
      msalIds?.clientId || "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";
    const targetAud = "https://substrate.office.com/sydney";
    const candidates = [];

    function collectCandidates(storage, label) {
      try {
        for (let i = 0; i < storage.length; i++) {
          const key = storage.key(i);
          if (!key) {
            continue;
          }
          const raw = storage.getItem(key);
          if (!raw || raw.length < 20) {
            continue;
          }
          candidates.push({ key, raw, storage: label });
        }
      } catch {
        // ignore storage access issues
      }
    }

    collectCandidates(localStorage, "localStorage");
    try {
      collectCandidates(sessionStorage, "sessionStorage");
    } catch {
      // ignore
    }

    const audiences = [];

    for (const entry of candidates) {
      let payload;
      try {
        payload = JSON.parse(entry.raw);
      } catch {
        continue;
      }
      if (!payload?.nonce || !payload?.data) {
        continue;
      }

      let decrypted;
      try {
        decrypted = await decryptPayload(
          cookie.key,
          payload.nonce,
          clientId,
          payload.data,
        );
      } catch {
        continue;
      }

      let parsed;
      try {
        parsed = JSON.parse(decrypted);
      } catch {
        continue;
      }
      if (!parsed?.secret) {
        continue;
      }

      const jwt = decodeJwtPayload(parsed.secret);
      if (!jwt?.aud) {
        continue;
      }
      audiences.push({ key: entry.key, storage: entry.storage, aud: jwt.aud });

      if (jwt.aud === targetAud) {
        return {
          token: parsed.secret,
          tokenPayload: jwt,
          cacheItem: parsed,
          cacheKey: entry.key,
          cacheStorage: entry.storage,
          candidateAudiences: audiences,
        };
      }
    }

    const audSummary = audiences
      .map((x) => x.aud)
      .filter(Boolean)
      .join(", ");
    throw new Error(
      audSummary
        ? `No Substrate token found (candidate audiences: ${audSummary})`
        : "No decryptable MSAL access token found in browser storage",
    );
  }

  async function getTokenAndIds() {
    const hintedIds = getMsalIds();
    const tokenInfo = await getAccessToken(hintedIds);
    const jwt = tokenInfo?.tokenPayload || {};
    const cacheItem = tokenInfo?.cacheItem || {};

    const localAccountId =
      hintedIds.localAccountId ||
      jwt.oid ||
      cacheItem.localAccountId ||
      cacheItem.local_account_id;
    const tenantId =
      hintedIds.tenantId || jwt.tid || cacheItem.realm || cacheItem.tenantId;
    const homeAccountId =
      hintedIds.homeAccountId ||
      cacheItem.homeAccountId ||
      cacheItem.home_account_id ||
      (localAccountId && tenantId ? `${localAccountId}.${tenantId}` : null);
    const userPrincipalName =
      hintedIds.userPrincipalName ||
      jwt.preferred_username ||
      jwt.upn ||
      cacheItem.username ||
      null;

    if (!localAccountId || !tenantId) {
      throw new Error(
        "Failed to resolve identity from hydration state / token payload",
      );
    }

    return {
      token: tokenInfo.token,
      localAccountId,
      tenantId,
      homeAccountId,
      userPrincipalName,
      clientId: hintedIds.clientId,
      tokenDebug: tokenInfo.candidateAudiences || [],
    };
  }

  async function substrateGetConversation(auth, conversationId) {
    const request = {
      conversationId,
      source: "officeweb",
      traceId: crypto.randomUUID().replace(/-/g, ""),
    };
    const url = `${SUBSTRATE_BASE}/GetConversation?request=${encodeURIComponent(JSON.stringify(request))}`;
    const headers = {
      authorization: `Bearer ${auth.token}`,
      "content-type": "application/json",
      "x-anchormailbox":
        auth.userPrincipalName || `Oid:${auth.localAccountId}@${auth.tenantId}`,
      "x-tenant-id": auth.tenantId,
      "x-client-application": "M365CopilotChat",
      "x-clientrequestid": crypto.randomUUID().replace(/-/g, ""),
      "x-routingparameter-sessionkey": auth.localAccountId,
      "x-scenario": "OfficeWeb",
    };
    const resp = await fetch(url, { method: "GET", headers });
    if (!resp.ok) throw new Error(`GetConversation returned ${resp.status}`);
    const json = await resp.json();
    lastConversation = json;
    lastConversationId = json?.conversationId || lastConversationId;
    cacheConversationSummary(json);
    updateCurrentChatInfo();
    return json;
  }

  function inferConversationIdFromUrl() {
    const href = location.href;
    try {
      const u = new URL(href);
      const pathConversationMatch = u.pathname.match(
        /(?:^|\/)conversation\/([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})(?:$|\/)/,
      );
      if (pathConversationMatch) {
        return pathConversationMatch[1];
      }
      for (const key of ["conversationId", "chatId", "cid", "id"]) {
        const v = u.searchParams.get(key);
        if (v) {
          return v;
        }
      }
    } catch {
      // Fall through to generic legacy matching below.
    }
    const m1 = href.match(
      /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/,
    );
    if (m1) {
      return m1[0];
    }
    const m2 = href.match(/[0-9a-fA-F]{32}/);
    if (m2) {
      return m2[0];
    }
    return null;
  }


  function isSubstrateGetConversationUrl(rawUrl) {
    try {
      const parsedUrl = new URL(String(rawUrl || ""), location.href);
      return parsedUrl.hostname === "substrate.office.com" && parsedUrl.pathname.includes("GetConversation");
    } catch {
      return false;
    }
  }

  // ---------------------------------------------------------------------------
  // Passive interception
  // ---------------------------------------------------------------------------

  const originalFetch = window.fetch;
  window.fetch = async function (...args) {
    const url = typeof args[0] === "string" ? args[0] : args[0]?.url || "";
    const response = await originalFetch.apply(this, args);

    if (isSubstrateGetConversationUrl(url)) {
      try {
        const clone = response.clone();
        const text = await clone.text();
        if (text && text.length > 20) {
          const json = JSON.parse(text);
          if (json?.conversationId) {
            lastConversation = json;
            lastConversationId = json.conversationId;
            cacheConversationSummary(json);
            updateCurrentChatInfo();
          }
        }
      } catch {
        /* ignore */
      }
    }

    return response;
  };

  const origXHROpen = XMLHttpRequest.prototype.open;
  const origXHRSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (method, url, ...rest) {
    this._m365ce_url = url;
    return origXHROpen.call(this, method, url, ...rest);
  };
  XMLHttpRequest.prototype.send = function (...args) {
    this.addEventListener("load", function () {
      try {
        const url = this._m365ce_url || "";
        if (isSubstrateGetConversationUrl(url)) {
          const json = JSON.parse(this.responseText);
          if (json?.conversationId) {
            lastConversation = json;
            lastConversationId = json.conversationId;
            cacheConversationSummary(json);
            updateCurrentChatInfo();
          }
        }
      } catch {
        /* ignore */
      }
    });
    return origXHRSend.apply(this, args);
  };

  // ---------------------------------------------------------------------------
  // UI (checkboxes)
  // ---------------------------------------------------------------------------

  function createCheckbox(id, label, checked) {
    const row = document.createElement("label");
    row.style.cssText =
      "display:flex; gap:8px; align-items:center; margin:4px 0; cursor:pointer; user-select:none;";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.id = id;
    cb.checked = !!checked;

    const span = document.createElement("span");
    span.textContent = label;
    span.style.cssText = "color:#c9c9d9; font-size:12px;";

    row.appendChild(cb);
    row.appendChild(span);

    return { row, cb };
  }

  function createUI() {
    if (document.getElementById("m365ce-ui")) {
      return;
    }

    const wrap = document.createElement("div");
    wrap.id = "m365ce-ui";
    wrap.style.cssText = [
      "position:fixed",
      "bottom:16px",
      "right:16px",
      "z-index:999999",
      "font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif",
      "font-size:12px",
    ].join(";");

    const panel = document.createElement("div");
    panel.style.cssText = [
      "background:#1a1a2e",
      "color:#e0e0e0",
      "border-radius:10px",
      "padding:10px 12px",
      "box-shadow:0 4px 20px rgba(0,0,0,0.35)",
      "min-width:360px",
    ].join(";");

    const title = document.createElement("div");
    title.textContent = `M365 Copilot Chat Conversation Export v${SCRIPT_VERSION}`;
    title.style.cssText = "font-weight:700;color:#9d8aff;margin-bottom:6px;";

    const status = document.createElement("div");
    status.id = "m365ce-status";
    status.textContent = "";
    status.style.cssText = "min-height:0;margin-bottom:4px;color:#50fa7b;";

    const currentChat = document.createElement("div");
    currentChat.id = "m365ce-current-chat";
    currentChat.textContent = "Current chat: (resolving...)";
    currentChat.style.cssText =
      "min-height:16px;margin-bottom:8px;color:#8be9fd;font-size:11px;line-height:1.35;";

    const opts = document.createElement("div");
    opts.style.cssText =
      "border:1px solid #2a2a45; border-radius:8px; padding:8px; margin:8px 0;";

    const optTitle = document.createElement("div");
    optTitle.textContent = "Readable Markdown details";
    optTitle.style.cssText = "font-weight:700;color:#8be9fd;margin-bottom:6px;";

    const { row: rowUnclassified, cb: cbUnclassified } = createCheckbox("m365ce-opt-unclassified-records", "Include unclassified records", currentUnclassifiedRecordSetting());

    cbUnclassified.addEventListener("change", () => {
      applyUnclassifiedRecordSetting(cbUnclassified.checked);
      status.textContent = "";
    });

    opts.appendChild(optTitle);
    opts.appendChild(rowUnclassified);

    const btn = document.createElement("button");
    btn.textContent = "Export conversation → .md + .json.md";
    btn.style.cssText = [
      "width:100%",
      "padding:8px 10px",
      "border-radius:8px",
      "border:1px solid #333",
      "background:#1f3460",
      "color:#e0e0e0",
      "cursor:pointer",
      "font-weight:700",
      "text-align:left",
      "transition:transform 120ms ease, opacity 120ms ease, background 120ms ease",
    ].join(";");

    btn.addEventListener(
      "mouseenter",
      () => (btn.style.background = "#2a4a80"),
    );
    btn.addEventListener(
      "mouseleave",
      () => (btn.style.background = "#1f3460"),
    );

    btn.addEventListener("click", async () => {
      if (btn.disabled) {
        return;
      }
      btn.disabled = true;
      const oldButtonText = btn.textContent;
      const exportStartedAt = Date.now();
      btn.textContent = "⏳ Exporting…";
      btn.style.opacity = "0.85";
      btn.style.transform = "scale(0.985)";
      btn.style.cursor = "wait";
      btn.style.background = "#315f9f";
      btn.style.boxShadow =
        "0 0 0 2px rgba(80,250,123,0.35), 0 0 14px rgba(80,250,123,0.25)";
      btn.setAttribute("aria-busy", "true");
      await new Promise((resolve) => setTimeout(resolve, 90));
      try {
        settings = loadSettings();
        cbUnclassified.checked = currentUnclassifiedRecordSetting();
        applyUnclassifiedRecordSetting(cbUnclassified.checked);

        status.textContent = "Preparing export files...";
        invalidateConversationCacheIfNeeded();

        const currentUrlConvId = inferConversationIdFromUrl();
        let conv = lastConversation;
        let convId = currentUrlConvId || lastConversationId;

        if (currentUrlConvId && conv?.conversationId !== currentUrlConvId) {
          conv = null;
        }

        if (!conv && !convId) {
          setStatus("Open a conversation first (then retry).", true);
          updateCurrentChatInfo();
          return;
        }

        if (!conv || (convId && conv?.conversationId !== convId)) {
          setStatus("Fetching conversation...");
          const auth = await getTokenAndIds();
          conv = await substrateGetConversation(
            auth,
            convId || conv?.conversationId,
          );
        }

        if (!conv || !conv.messages || conv.messages.length === 0) {
          setStatus("No messages found to export.", true);
          return;
        }

        setStatus(
          `Building export files for: ${conv?.chatName || "current chat"}...`,
        );
        const exportedAt = new Date().toISOString();
        const timestampForFilename = exportedAt.replace("T", "_").replace(/:/g, "-");
        const baseName = `${sanitizeFilename(conv.chatName || "m365-copilot-chat")}_${timestampForFilename}`;
        const readableMd = toMarkdownCardFirst(conv, exportedAt);
        downloadText(`${baseName}.md`, readableMd);

        const rawMd = toRawJsonMarkdown(conv, exportedAt);
        downloadText(`${baseName}.json.md`, rawMd);

        setStatus("Exported ✔  (2 files: .md + .json.md)");
      } catch (e) {
        console.error("[M365 Export]", e);
        setStatus(`Error: ${e.message}`, true);
      } finally {
        const busyRemaining = Math.max(0, 450 - (Date.now() - exportStartedAt));
        if (busyRemaining > 0) {
          await new Promise((resolve) => setTimeout(resolve, busyRemaining));
        }
        btn.disabled = false;
        btn.textContent = oldButtonText;
        btn.style.opacity = "";
        btn.style.transform = "";
        btn.style.cursor = "pointer";
        btn.style.background = "#1f3460";
        btn.style.boxShadow = "";
        btn.removeAttribute("aria-busy");
      }
    });

    panel.appendChild(title);
    panel.appendChild(status);
    panel.appendChild(currentChat);
    panel.appendChild(opts);
    panel.appendChild(btn);

    wrap.appendChild(panel);
    document.body.appendChild(wrap);
    updateCurrentChatInfo();
    scheduleCurrentChatResolution(250);
  }

  function init() {
    createUI();
    updateCurrentChatInfo();
    scheduleCurrentChatResolution(250);
    let lastHref = location.href;
    setInterval(() => {
      if (location.href !== lastHref) {
        lastHref = location.href;
        invalidateConversationCacheIfNeeded();
        scheduleCurrentChatResolution(300);
        setTimeout(() => {
          createUI();
          updateCurrentChatInfo();
          scheduleCurrentChatResolution(300);
        }, 500);
        return;
      }
      invalidateConversationCacheIfNeeded();
    }, 800);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();

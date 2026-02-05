/* global document, Office, Word, console */
/* eslint-disable no-console */

import { AIService } from "./ai-service";
import { getGranularDiff } from "./diff-service";
import { handleToolEvent } from "./tool-renderer";
import TurndownService from "turndown";

const turndownService = new TurndownService();
let aiService: AIService | null = null;

// UI Elements
let chatHistory: HTMLElement;
let promptInput: HTMLTextAreaElement;
let sendBtn: HTMLButtonElement;
let settingsBtn: HTMLButtonElement;
let settingsPanel: HTMLElement;
let tokenInput: HTMLInputElement;
let apiKeyInput: HTMLInputElement;
let qwenApiKeyInput: HTMLInputElement;
let saveSettingsBtn: HTMLButtonElement;
let selectionIndicator: HTMLElement;
let selectionPreview: HTMLElement;
let clearSelectionBtn: HTMLButtonElement;

// Model Selection Elements
let modelSelectorBtn: HTMLElement;
let modelMenu: HTMLElement;
let currentModelName: HTMLElement;
let selectedModelId = "google/gemini-pro";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initUI();
    loadSettings();
    registerEventHandlers();
  }
});

function initUI() {
  chatHistory = document.getElementById("chat-history")!;
  promptInput = document.getElementById("prompt-input") as HTMLTextAreaElement;
  sendBtn = document.getElementById("send-btn") as HTMLButtonElement;
  settingsBtn = document.getElementById("settings-btn") as HTMLButtonElement;
  settingsPanel = document.getElementById("settings-panel")!;
  tokenInput = document.getElementById("token-input") as HTMLInputElement;
  apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
  qwenApiKeyInput = document.getElementById("qwen-api-key-input") as HTMLInputElement;
  saveSettingsBtn = document.getElementById("save-settings-btn") as HTMLButtonElement;
  selectionIndicator = document.getElementById("selection-indicator")!;
  selectionPreview = document.getElementById("selected-text-preview")!;
  clearSelectionBtn = document.getElementById("clear-selection-btn") as HTMLButtonElement;

  // Model Selection
  modelSelectorBtn = document.getElementById("model-selector-btn")!;
  modelMenu = document.getElementById("model-menu")!;
  currentModelName = document.getElementById("current-model-name")!;

  // Auto-resize textarea
  promptInput.addEventListener("input", function () {
    this.style.height = "auto";
    this.style.height = (this.scrollHeight) + "px";

    // Enable/disable send button
    sendBtn.disabled = this.value.trim() === "";
  });

  // Enter to send
  promptInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  });

  // Check initial selection
  checkSelection();
  // Poll for selection changes
  setInterval(checkSelection, 1000);
}

function registerEventHandlers() {
  settingsBtn.onclick = () => {
    settingsPanel.classList.toggle("hidden");
    if (!settingsPanel.classList.contains("hidden")) {
      tokenInput.focus();
    }
  };

  // Model Selector Toggle
  modelSelectorBtn.onclick = (e) => {
    e.stopPropagation();
    modelMenu.classList.toggle("hidden");
  };

  // Model Menu Item Click
  document.querySelectorAll(".menu-item").forEach(item => {
    item.addEventListener("click", (e) => {
      const target = e.currentTarget as HTMLElement;
      selectedModelId = target.getAttribute("data-model") || selectedModelId;
      currentModelName.textContent = target.textContent;
      modelMenu.classList.add("hidden");
      localStorage.setItem("openclaw_model", selectedModelId);

      // Update visual selected state
      document.querySelectorAll(".menu-item").forEach(mi => mi.classList.remove("selected"));
      target.classList.add("selected");
    });
  });

  // Close menu when clicking outside
  document.addEventListener("click", () => {
    if (modelMenu) modelMenu.classList.add("hidden");
  });

  if (tokenInput) {
    tokenInput.addEventListener("change", () => {
      localStorage.setItem("openclaw_token", tokenInput.value);
      if (tokenInput.value) {
        aiService = new AIService(tokenInput.value);
      }
    });
  }

  if (saveSettingsBtn) {
    saveSettingsBtn.onclick = async () => {
      if (!aiService) {
        aiService = new AIService(tokenInput.value || "test-token");
      }

      saveSettingsBtn.disabled = true;
      saveSettingsBtn.textContent = "Saving...";

      try {
        await aiService.updateConfig(apiKeyInput.value, qwenApiKeyInput.value);
        localStorage.setItem("openclaw_api_key", apiKeyInput.value);
        localStorage.setItem("openclaw_qwen_api_key", qwenApiKeyInput.value);
        appendMessage("system", "Settings saved and OpenClaw updated!");
        settingsPanel.classList.add("hidden");
      } catch (e: any) {
        appendMessage("system", "Error saving settings: " + e.message);
      } finally {
        saveSettingsBtn.disabled = false;
        saveSettingsBtn.textContent = "Save to OpenClaw";
      }
    };
  }

  sendBtn.onclick = handleSend;

  if (clearSelectionBtn) {
    clearSelectionBtn.onclick = () => {
      selectionIndicator.classList.add("hidden");
    };
  }
}

function loadSettings() {
  const savedToken = localStorage.getItem("openclaw_token");
  const savedApiKey = localStorage.getItem("openclaw_api_key");
  const savedQwenApiKey = localStorage.getItem("openclaw_qwen_api_key");
  const savedModel = localStorage.getItem("openclaw_model");

  if (savedToken && tokenInput) {
    tokenInput.value = savedToken;
    aiService = new AIService(savedToken);
  }
  if (savedApiKey && apiKeyInput) apiKeyInput.value = savedApiKey;
  if (savedQwenApiKey && qwenApiKeyInput) qwenApiKeyInput.value = savedQwenApiKey;

  if (savedModel) {
    selectedModelId = savedModel;
    const menuEl = document.querySelector(`.menu-item[data-model="${savedModel}"]`);
    if (menuEl && currentModelName) {
      currentModelName.textContent = menuEl.textContent;
      menuEl.classList.add("selected");
    }
  }

  if (!savedToken && settingsPanel) {
    settingsPanel.classList.remove("hidden");
  }
}

async function checkSelection() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (selection.text && selection.text.trim().length > 0) {
        selectionIndicator.classList.remove("hidden");
        selectionPreview.textContent = selection.text.substring(0, 50) + (selection.text.length > 50 ? "..." : "");
      } else {
        selectionIndicator.classList.add("hidden");
      }
    });
  } catch (e) {
    console.error("Error checking selection", e);
  }
}

async function handleSend() {
  const prompt = promptInput.value.trim();
  if (!prompt) return;

  if (!aiService) {
    appendMessage("system", "Please set your OpenClaw Gateway Token in settings first.");
    settingsPanel.classList.remove("hidden");
    return;
  }

  // 1. Add User Message
  appendMessage("user", prompt);
  promptInput.value = "";
  promptInput.style.height = "auto";
  sendBtn.disabled = true;

  // 2. Prepare for AI Processing
  let originalText = "";
  let originalMarkdown = "";
  let hasSelection = false;
  let msgId = "";
  let contentDiv: HTMLElement | null = null;

  try {
    await Word.run(async (context) => {
      // Step A: Check for selection (Optional now)
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (selection.text && selection.text.trim().length > 0) {
        hasSelection = true;
        originalText = selection.text;

        // Step B: Only fetch HTML if there is a selection
        const htmlSelection = selection.getHtml();
        await context.sync();
        const originalHtml = htmlSelection.value;
        originalMarkdown = turndownService.turndown(originalHtml);
      }
    });
  } catch (e: any) {
    console.warn("Selection check failed, proceeding as chat-only:", e);
    hasSelection = false;
  }

  // 3. UI Setup: AI Message Container (Streaming)
  msgId = appendMessage("ai", "");
  const msgEl = document.getElementById(msgId);
  if (msgEl) {
    contentDiv = msgEl.querySelector(".content");
    if (contentDiv) contentDiv.innerHTML = '<span class="typing-indicator">Thinking...</span>';
  }

  // 4. Call AI (Streaming context-aware)
  // If we have a selection, pass it. If not, just pass the prompt.
  const aiContext = hasSelection ? originalMarkdown : "";

  // Note: We might want to adjust the system prompt or user prompt structure slightly when there is no context
  // But for now, passing empty string as context is fine, the AI Service handles the formatting.

  let fullRewrittenText = "";
  let isFirstChunk = true;
  try {
    // Clear old content
    // contentDiv.innerHTML = ""; // DO NOT clear here, as we are appending to existing chat history div structure
    // But wait, contentDiv IS the message content div created in handleSend. So yes, clear it.

    let mdContainer: HTMLElement | null = null;
    let accumulatedText = "";

    // Simple Markdown Parser (can be improved or replaced with marked.js later)
    const parseMarkdown = (text: string) => {
      let html = text
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/\n/g, '<br>')
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/`(.*?)`/g, '<code>$1</code>')
        .replace(/\[(.*?)\]\((.*?)\)/g, '<a href="$2" target="_blank">$1</a>');
      return html;
    };

    fullRewrittenText = await aiService.streamRewrite(prompt, aiContext, selectedModelId, (chunk) => {
      if (!contentDiv) return;

      if (chunk.startsWith("__TOOL_EVENT__:")) {
        if (isFirstChunk) {
          contentDiv.innerHTML = ""; // Clear "Thinking..."
          isFirstChunk = false;
        }
        // If we were streaming text, maybe we should stop using the current mdContainer?
        // For now, let's just insert the tool after whatever content we have.

        const eventData = JSON.parse(chunk.split("__TOOL_EVENT__:")[1]);

        const handled = handleToolEvent(contentDiv, eventData, parseMarkdown);
        if (handled) {
          // Tool event handled (call or result)
          // Prepare state for text coming after this tool
          mdContainer = null;
          accumulatedText = "";
          return;
        }
      }

      // Handle Text Chunk
      if (isFirstChunk) {
        contentDiv.innerHTML = ""; // Clear "Thinking..."
        isFirstChunk = false;
      }

      if (!mdContainer) {
        mdContainer = document.createElement("div");
        mdContainer.className = "md-content";
        contentDiv.appendChild(mdContainer);
      }

      accumulatedText += chunk;
      mdContainer.innerHTML = parseMarkdown(accumulatedText);
      chatHistory.scrollTop = chatHistory.scrollHeight;

    });
    // Stream finished. No auto-apply to Word.
  } catch (error: any) {
    console.error(error);
    if (contentDiv) {
      contentDiv.innerHTML += `<div style="color:red; margin-top:8px;">Error: ${error.message}</div>`;
    } else {
      appendMessage("system", "Error: " + error.message);
    }
  } finally {
    sendBtn.disabled = false;
  }
}

/**
 * Applies granular revisions to the Word document using the "Remaining Range" strategy.
 * This ensures interwoven edits (red-line/green-underline) without bunching up at the start.
 */
async function applyRevisionsToWord(instructions: any[]) {
  await Word.run(async (context) => {
    const body = context.document.body;

    // Ensure Track Changes is ON
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

    // Maintain a pointer to the "unprocessed" part of the document
    // We start with the entire body
    let remainingRange = body.getRange();

    for (const part of instructions) {
      if (part.type === "keep") {
        // Search for the text within the remaining range
        // We use .search() which returns a range collection
        let searchRes = remainingRange.search(part.text, { matchCase: true });
        searchRes.load("items");
        await context.sync();

        if (searchRes.items.length > 0) {
          // Update remainingRange to start AFTER the found text
          // This effectively "skips" the kept text
          const foundRange = searchRes.items[0];
          remainingRange = foundRange.getRange(Word.RangeLocation.after).expandTo(body.getRange(Word.RangeLocation.end));
        }
      }
      else if (part.type === "insert") {
        // Insert at the START of the remaining area
        // In Track Changes mode, this shows as an insertion (underline)
        let insertedRange = remainingRange.insertText(part.text, Word.InsertLocation.start);
        await context.sync();

        // Move remainingRange to start AFTER the inserted text
        remainingRange = insertedRange.getRange(Word.RangeLocation.after).expandTo(body.getRange(Word.RangeLocation.end));
      }
      else if (part.type === "delete") {
        // Search for the text to delete within the remaining range
        let searchRes = remainingRange.search(part.text, { matchCase: true });
        searchRes.load("items");
        await context.sync();

        if (searchRes.items.length > 0) {
          let target = searchRes.items[0];
          // We need to capture the point AFTER the deletion target BEFORE we delete it
          // because deleting it might invalidate the range structure slightly differently
          let nextPoint = target.getRange(Word.RangeLocation.after);

          // Delete it (shows as strikethrough in Track Changes)
          target.delete();

          // Move remainingRange to start AFTER the deleted word
          remainingRange = nextPoint.expandTo(body.getRange(Word.RangeLocation.end));
        }
      }
      // Sync after every step to ensure the DOM is up to date for the next search
      await context.sync();
    }
  });
}

function appendMessage(role: "user" | "ai" | "system", text: string, isLoading = false): string {
  const msgId = "msg-" + Date.now();
  const msgDiv = document.createElement("div");
  msgDiv.id = msgId;
  msgDiv.className = `message ${role}-message`;

  // For User: Just text (no avatar based on new design)
  if (role === "user") {
    msgDiv.innerHTML = `<div class="content">${text}</div>`;
  }
  // For System/Simple AI:
  else {
    let contentHtml = text;
    if (isLoading) {
      contentHtml = '<i class="fa-solid fa-spinner fa-spin"></i> ' + text;
    }
    msgDiv.innerHTML = `<div class="content">${contentHtml}</div>`;
  }

  chatHistory.appendChild(msgDiv);
  chatHistory.scrollTop = chatHistory.scrollHeight;

  return msgId;
}

function removeMessage(id: string) {
  const el = document.getElementById(id);
  if (el) el.remove();
}

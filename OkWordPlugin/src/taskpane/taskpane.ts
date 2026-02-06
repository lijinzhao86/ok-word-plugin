/* global document, Office, Word, console */
/* eslint-disable no-console */

import { AIService } from "./ai-service";
import { getGranularDiff } from "./diff-service";
import { handleToolEvent } from "./tool-renderer";
import { executeClientTool } from "./tools/word-implementation";
import TurndownService from "turndown";
import { marked } from "marked";

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
let lastUserPrompt = "";

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
        if (aiService) aiService.close();
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

  // New Chat Button Handler
  const newChatBtn = document.getElementById("new-chat-btn") as HTMLButtonElement;
  if (newChatBtn) {
    newChatBtn.onclick = () => {
      if (!aiService || newChatBtn.disabled) return;
      aiService.startNewSession();
      chatHistory.innerHTML = `
            <div class="message system-message">
                <div class="avatar"><i class="fa-solid fa-robot"></i></div>
                <div class="content">
                    Ready to help you edit your document. Select some text and tell me what to do.
                </div>
            </div>`;
      // Disable button after reset
      newChatBtn.disabled = true;
    };
  }
}

/**
 * Checks if the chat has any user interaction and updates the New Chat button state.
 */
function updateNewChatButtonState() {
  const newChatBtn = document.getElementById("new-chat-btn") as HTMLButtonElement;
  if (!newChatBtn) return;

  // Check if there are any user messages or AI messages (excluding the initial system welcome)
  // The initial welcome message usually has class 'system-message' and is the only child.
  const messages = chatHistory.querySelectorAll('.message');
  let hasInteraction = false;

  if (messages.length > 1) {
    hasInteraction = true;
  } else if (messages.length === 1) {
    // If meant to be strict: only enable if USER has typed something
    const firstMsg = messages[0];
    // If the only message is user message (rare), or if it's NOT the default system message
    if (firstMsg.classList.contains('user-message')) {
      hasInteraction = true;
    }
  }

  newChatBtn.disabled = !hasInteraction;

  // Visual opacity handled by CSS :disabled usually, but ensure icon style correct if needed
  newChatBtn.style.opacity = hasInteraction ? "1" : "0.5";
  newChatBtn.style.cursor = hasInteraction ? "pointer" : "default";
}

async function loadHistory() {
  if (!aiService) return;

  // Show loading indicator
  const loaderId = appendMessage("system", "Restoring history...", true);

  try {
    const history = await aiService.getHistory();
    // Remove loader
    removeMessage(loaderId);

    // Clear default welcome message if we have history
    if (history.length > 0) {
      chatHistory.innerHTML = "";
    }

    for (const msg of history) {
      if (msg.role === "user") {
        appendMessage("user", msg.content);
      } else if (msg.role === "assistant") {
        // Only render text content for now as requested
        if (msg.content) {
          const msgId = appendMessage("ai", "");
          const msgEl = document.getElementById(msgId);
          if (msgEl) {
            const contentDiv = msgEl.querySelector(".content");
            // Parse markdown for historical messages
            if (contentDiv) contentDiv.innerHTML = marked.parse(msg.content) as string;
          }
        }
      }
      // TODO: Handle Tool Calls restoration
    }

    // Scroll to bottom
    chatHistory.scrollTop = chatHistory.scrollHeight;

    // Update button state based on loaded history
    updateNewChatButtonState();

  } catch (e) {
    console.error("Failed to load history", e);
    removeMessage(loaderId);
  }
}

function loadSettings() {
  const savedToken = localStorage.getItem("openclaw_token");
  const savedApiKey = localStorage.getItem("openclaw_api_key");
  const savedQwenApiKey = localStorage.getItem("openclaw_qwen_api_key");
  const savedModel = localStorage.getItem("openclaw_model");

  if (savedToken && tokenInput) {
    tokenInput.value = savedToken;
    if (aiService) aiService.close();
    aiService = new AIService(savedToken);

    // Load history after service is initialized
    loadHistory();
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
  lastUserPrompt = prompt;

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
  updateNewChatButtonState(); // Enable New Chat button now that user has interacted

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
  const aiContext = hasSelection ? originalMarkdown : "";

  try {
    if (contentDiv) {
      const handler = createStreamHandler(contentDiv, chatHistory);
      await aiService!.streamRewrite(prompt, aiContext, selectedModelId, handler);
    }
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
 * Recovers from a tool execution by feeding the data back to the AI.
 */
async function handleToolOutput(actionName: string, output: string) {
  const msg = `Original Context/Request: "${lastUserPrompt}"\n\n[System] Tool '${actionName}' executed successfully.\nOutput:\n\`\`\`\n${output}\n\`\`\`\nPlease continue fulfilling the original request using this information.`;

  // Create AI response placeholder for the CONTINUED response
  const aiMsgId = appendMessage("ai", "");
  const aiMsgEl = document.getElementById(aiMsgId);
  let contentDiv: HTMLElement | null = null;

  if (aiMsgEl) {
    contentDiv = aiMsgEl.querySelector(".content");
    if (contentDiv) contentDiv.innerHTML = '<span class="typing-indicator">Processing tool output...</span>';
  }

  try {
    if (contentDiv) {
      const handler = createStreamHandler(contentDiv, chatHistory);
      // We pass empty context because the tool output IS the context
      await aiService!.streamRewrite(msg, "", selectedModelId, handler);
    }
  } catch (e: any) {
    if (contentDiv) contentDiv.innerHTML += `<br><span style="color:red">Error: ${e.message}</span>`;
  }
}

function createStreamHandler(contentDiv: HTMLElement, scrollContainer: HTMLElement) {
  let isFirstChunk = true;
  let mdContainer: HTMLElement | null = null;
  let accumulatedText = "";

  const parseMarkdown = (text: string) => {
    return marked.parse(text) as string;
  };

  // Track which tool indices we have already rendered cards for
  const renderedToolIndices = new Set<number>();
  // Map index -> real tool_call_id (e.g. 0 -> call_xyz123)
  const toolIdMap = new Map<number, string>();

  return (chunk: string) => {
    if (!contentDiv) return;

    if (chunk.startsWith("__TOOL_EVENT__:")) {
      if (isFirstChunk) {
        contentDiv.innerHTML = "";
        isFirstChunk = false;
      }

      try {
        const rawPayload = chunk.split("__TOOL_EVENT__:")[1];
        const eventPayload = JSON.parse(rawPayload);
        console.log("🔥 [RAW EVENT]", JSON.stringify(eventPayload)); // DEBUG EVERYTHING

        // 1. Resolve Tool Index (if available) to track IDs
        const toolIndex = typeof eventPayload.index === 'number' ? eventPayload.index : 0;

        // Store real ID -> Index mapping when we see a tool_call
        if (eventPayload.custom_type === "tool_call" && eventPayload.id && !eventPayload.id.startsWith("chatcmpl")) {
          toolIdMap.set(toolIndex, eventPayload.id);
        }

        // 2. Resolve UI Tool ID
        // Strategy: 
        // A. If payload has specific tool_call_id, USE IT (This is what we want from backend)
        // B. If payload has index, look up the map (Standard OpenAI stream behavior)
        // C. Fallback: use eventPayload.id (which might be chatcmpl_xxx, causing the bug, but we have no choice until backend is fixed)

        let uiToolId = eventPayload.tool_call_id; // Priority 1

        if (!uiToolId) {
          const mappedId = toolIdMap.get(toolIndex);
          uiToolId = mappedId || `${eventPayload.id || Date.now()}_${toolIndex}`;
        }

        console.log(`🛠️ [Tool] Type: ${eventPayload.custom_type}, UI-ID: ${uiToolId}`, eventPayload);

        // Handle Proxy Tools (WORD_PLUGIN_ACTION)
        const resultObj = eventPayload.result?.details || eventPayload.result;
        if (eventPayload.custom_type === "tool_result" && resultObj && resultObj.__CLIENT_ACTION__ === "WORD_PLUGIN_ACTION") {
          const { action, args } = resultObj;
          const toolDomId = `tool-proxy-${uiToolId}`;

          if (!document.getElementById(toolDomId)) {
            // Render card
            contentDiv.insertAdjacentHTML('beforeend', `
                  <div id="${toolDomId}" class="tool-container searching">
                     <div class="tool-header">
                        <span class="icon fa-solid fa-bolt"></span>
                        <span class="label">Running ${action}...</span>
                     </div>
                  </div>
                `);

            executeClientTool(action, args).then(output => {
              const el = document.getElementById(toolDomId);
              if (el) {
                el.className = "tool-container completed";
                el.querySelector(".label")!.textContent = `Finished ${action}`;
              }
              handleToolOutput(action, output);
            }).catch(err => {
              const el = document.getElementById(toolDomId);
              if (el) {
                el.className = "tool-container error";
                el.querySelector(".label")!.textContent = `Failed: ${err.message}`;
              }
            });
          }
          // Reset text accumulation so subsequent AI text starts fresh
          mdContainer = null;
          accumulatedText = "";
          return;
        }

        // Handle Standard Tools: we inject our unique UI ID
        const handled = handleToolEvent(contentDiv, { ...eventPayload, id: uiToolId }, parseMarkdown);
        if (handled) {
          mdContainer = null;
          accumulatedText = "";
          return;
        }
      } catch (e) {
        console.error("Error parsing tool event", e);
      }
    }

    if (isFirstChunk) {
      contentDiv.innerHTML = "";
      isFirstChunk = false;
    }

    if (!mdContainer) {
      mdContainer = document.createElement("div");
      mdContainer.className = "md-content";
      contentDiv.appendChild(mdContainer);
    }

    accumulatedText += chunk;
    mdContainer.innerHTML = parseMarkdown(accumulatedText);
    scrollContainer.scrollTop = scrollContainer.scrollHeight;
  };
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

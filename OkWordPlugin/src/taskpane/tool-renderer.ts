
// Tool Abstraction: Configuration & Rendering Logic

// Helper to clean OpenClaw security tags
const cleanWebText = (text: string) => {
    if (!text) return "";
    return text
        .replace(/<<<EXTERNAL_UNTRUSTED_CONTENT>>>/g, "")
        .replace(/<<<END_EXTERNAL_UNTRUSTED_CONTENT>>>/g, "")
        .replace(/Source: Web Search/g, "")
        .replace(/---\s*/g, "")
        .trim();
};

interface ToolConfig {
    icon: string;
    verbs: [string, string]; // [Present, Past] e.g. ["Searching", "Searched"]
    renderLabelSuffix: (args: any) => string;
    renderContent: (res: any) => string;
}

const TOOL_REGISTRY: Record<string, ToolConfig> = {
    web_search: {
        icon: "icon-globe",
        verbs: ["Searching web for", "Searched web for"],
        renderLabelSuffix: (a) => ` "${a.query || '...'}"`,
        renderContent: (resObj) => {
            let results: any[] = [];
            // Strategy 1: OpenClaw embedded structure (details.results)
            if (resObj?.details?.results && Array.isArray(resObj.details.results)) results = resObj.details.results;
            // Strategy 2: OpenClaw content array
            else if (Array.isArray(resObj?.content)) {
                const textItem = resObj.content.find((item: any) => item.type === 'text');
                if (textItem?.text) { try { const inner = JSON.parse(textItem.text); if (inner.results) results = inner.results; } catch (e) { } }
            }
            // Strategy 3: Direct results array
            else if (resObj?.results && Array.isArray(resObj.results)) results = resObj.results;
            // Strategy 4: Brave API specific (web.results)
            else if (resObj?.web?.results && Array.isArray(resObj.web.results)) results = resObj.web.results;

            if (results.length > 0) {
                return results.map((r: any) => `• [${cleanWebText(r.title)}](${r.url})\n  ${cleanWebText(r.description)}`).join("\n\n");
            }
            return typeof resObj === 'object' ? ("```json\n" + JSON.stringify(resObj, null, 2) + "\n```") : String(resObj);
        }
    },
    read_file: {
        icon: "icon-file",
        verbs: ["Reading", "Read"],
        renderLabelSuffix: (a) => ` ${a.path || 'file'}`,
        renderContent: (res) => `\`\`\`\n${typeof res === 'string' ? res : (res.content || JSON.stringify(res, null, 2))}\n\`\`\``
    },
    edit_file: {
        icon: "icon-edit",
        verbs: ["Editing", "Edited"],
        renderLabelSuffix: (a) => ` ${a.path || 'file'}`,
        renderContent: (res) => typeof res === 'string' ? res : (res.diff ? `Diff:\n${res.diff}` : "Edit applied successfully.")
    }
};

export interface ToolEventData {
    custom_type: string; // "tool_call" | "tool_result"
    id?: string;
    name: string;
    args?: any;
    result?: any;
}

/**
 * Handles tool events (call visualization and result rendering)
 * @returns true if a tool event was handled, false otherwise
 */
export function handleToolEvent(
    contentDiv: HTMLElement,
    eventData: ToolEventData,
    parseMarkdown: (text: string) => string
): boolean {
    if (eventData.custom_type !== "tool_call" && eventData.custom_type !== "tool_result") {
        return false;
    }

    const toolDomId = `tool-${eventData.id || Date.now()}`;
    const name = eventData.name;

    const config = TOOL_REGISTRY[name] || {
        icon: "icon-globe",
        verbs: [`Using ${name}`, `Used ${name}`],
        renderLabelSuffix: (a: any) => "...",
        renderContent: (res: any) => JSON.stringify(res, null, 2)
    };

    if (eventData.custom_type === "tool_call") {
        if (!document.getElementById(toolDomId)) {
            const labelText = `${config.verbs[0]}${config.renderLabelSuffix(eventData.args || {})}`;
            const html = `
                  <div id="${toolDomId}" class="tool-container searching">
                      <div class="tool-header">
                          <span class="icon ${config.icon}"></span>
                          <span class="label">${labelText}...</span>
                      </div>
                      <div class="tool-content"></div>
                  </div>`;
            contentDiv.insertAdjacentHTML('beforeend', html);
        }
        return true;
    } else if (eventData.custom_type === "tool_result") {
        const el = document.getElementById(toolDomId);
        if (el) {
            el.className = "tool-container completed collapsed";
            const label = el.querySelector(".label");
            const content = el.querySelector(".tool-content");
            const header = el.querySelector(".tool-header") as HTMLElement;

            if (label && label.textContent) {
                // Try to smartly replace the verb
                const currentText = label.textContent;
                // If it starts with the progressive verb, replace it. Otherwise just use new label.
                if (currentText.startsWith(config.verbs[0])) {
                    label.textContent = currentText.replace(config.verbs[0], config.verbs[1]).replace("...", "");
                } else {
                    // Fallback if text content changed or doesn't match
                    label.textContent = config.verbs[1] + config.renderLabelSuffix(eventData.args || {});
                }
            }

            let resultText = "No results.";
            try {
                let resObj = eventData.result;
                if (typeof resObj === 'string') {
                    try { resObj = JSON.parse(resObj); } catch (e) { resObj = eventData.result; }
                }
                resultText = config.renderContent(resObj);
            } catch (e) { resultText = "Error: " + String(e); }

            if (content) content.innerHTML = parseMarkdown(resultText);
            if (header) {
                header.onclick = (e) => { e.stopPropagation(); el.classList.toggle("expanded"); };
            }
        }
        return true;
    }

    return false;
}


/* global Word, Office */
import TurndownService from "turndown";

const turndownService = new TurndownService();

// Add tables support if possible, or just standard markdown
// turndownService.addRule(...) 

export async function executeClientTool(action: string, args: any): Promise<any> {
    console.log(`[WordPlugin] Executing tool: ${action}`, args);
    try {
        switch (action) {
            case "read_document":
                return await readDocument(args);
            case "grep_document":
                return await grepDocument(args);
            default:
                throw new Error(`Unknown client action: ${action}`);
        }
    } catch (e: any) {
        console.error("Tool execution failed:", e);
        return `Error executing ${action}: ${e.message}`;
    }
}

async function readDocument(args: { scope: string; offset?: number; limit?: number }) {
    const { scope, offset = 0, limit = 50 } = args;

    return await Word.run(async (context) => {

        // 1. Structure (Outline)
        if (scope === "structure") {
            const paragraphs = context.document.body.paragraphs;
            // Only load outlineLevel and text to be lightweight
            paragraphs.load("items/outlineLevel,items/text");
            await context.sync();

            const headings = paragraphs.items
                .map((p, i) => ({ index: i, level: p.outlineLevel, text: p.text.trim() }))
                .filter(h => h.level < 9 && h.text.length > 0);

            if (headings.length === 0) {
                return "<system-info>No structured headings found.</system-info>";
            }

            // Calculate stats for each heading (Inclusive of sub-headings)
            const allItems = paragraphs.items;
            const totalParas = allItems.length;

            const enrichedHeadings = headings.map((h, i, arr) => {
                // Find end index (next heading with level <= current level)
                let endIndex = totalParas;
                for (let j = i + 1; j < arr.length; j++) {
                    if (arr[j].level <= h.level) {
                        endIndex = arr[j].index;
                        break;
                    }
                }

                // Calculate stats
                let charCount = 0;
                // Optimization: Slice once if possible, or iterate loop
                for (let k = h.index; k < endIndex; k++) {
                    charCount += allItems[k].text.length;
                }

                const paraCount = endIndex - h.index;
                return { ...h, paraCount, charCount };
            });

            return enrichedHeadings.map(h => `${'#'.repeat(h.level + 1)} ${h.text} (Para Index: ${h.index}, Size: ${h.paraCount} paras / ~${h.charCount} chars)`).join("\n");
        }

        // 2. Paragraphs (Content)
        if (scope === "paragraph") {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items/text");
            await context.sync();

            const total = paragraphs.items.length;
            const start = Math.min(offset, total);
            // Cap the limit to avoid huge payloads
            const safeLimit = Math.min(limit, 100);
            const end = Math.min(start + safeLimit, total);

            if (start >= total) return `Offset ${start} is out of bounds (Total: ${total}).`;

            const slice = paragraphs.items.slice(start, end);
            const content = slice.map((p, i) => `[Para ${start + i}]: ${p.text}`).join("\n\n");

            return content;
        }

        return "<system-info>Invalid scope. Use 'structure' to see outline, or 'paragraph' to read content.</system-info>";
    });
}

async function grepDocument(args: { pattern: string; output_mode?: string; match_case?: boolean; context_lines?: number }) {
    const { pattern, output_mode = "content", match_case = false } = args;

    return await Word.run(async (context) => {
        const searchOptions = { matchCase: match_case, matchWholeWord: false };
        const searchResults = context.document.body.search(pattern, searchOptions);
        searchResults.load("items");
        await context.sync();

        const count = searchResults.items.length;
        if (output_mode === "count") {
            return `Found ${count} matches for pattern "${pattern}".`;
        }

        if (count === 0) return `No matches found for pattern "${pattern}".`;

        // Optimized approach: Batch load paragraph text
        const limit = Math.min(count, 20);

        // Re-queue items proxy
        const matches = searchResults.items.slice(0, limit);

        // We need to load paragraph text for each match context.
        // Range.paragraphs returns a collection. We need to load it.
        // Loading 'paragraphs/items/text' on the Range object directly? 
        // Range has a 'paragraphs' property which is a collection.
        // We can load the collection itself.
        const rangesParagraphs = matches.map(m => m.paragraphs);
        rangesParagraphs.forEach(c => c.load("items/text"));

        await context.sync();

        const textOutput = matches.map((m, i) => {
            // Access the loaded paragraph collection
            const pText = m.paragraphs.items.length > 0 ? m.paragraphs.items[0].text : "(Context unreachable)";
            return `Match ${i + 1}: ...${pText.trim()}...`;
        }).join("\n---\n");

        return `Found ${count} matches. Showing first ${limit}:\n\n${textOutput}`;
    });
}

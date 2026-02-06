
import { Type } from "@sinclair/typebox";
import type { OpenClawConfig } from "../../config/config.js";
import type { AnyAgentTool } from "./common.js";
import { jsonResult } from "./common.js";

const WordReadSchema = Type.Object({
    scope: Type.String({
        description: "Scope to read: 'structure' (returns outline & Para Index) or 'paragraph' (returns content). Use 'structure' first to find Para Index.",
        enum: ["structure", "paragraph"]
    }),
    offset: Type.Optional(Type.Number({ description: "For 'paragraph' scope: Start paragraph index (from Para Index). Default 0." })),
    limit: Type.Optional(Type.Number({ description: "For 'paragraph' scope: Number of paragraphs to read. Default 50." })),
});

const WordGrepSchema = Type.Object({
    pattern: Type.String({ description: "Regex pattern or keyword to search." }),
    output_mode: Type.Optional(Type.String({
        description: "Output mode: 'content' (show matches) or 'count' (show count). Default 'content'.",
        enum: ["content", "count"]
    })),
    match_case: Type.Optional(Type.Boolean({ description: "Match case. Default false." })),
    context_lines: Type.Optional(Type.Number({ description: "Number of context paragraphs around match." })),
});

export function createWordReadTool(options?: {
    config?: OpenClawConfig;
    reverseRpc?: (action: string, args: unknown) => Promise<unknown>;
}): AnyAgentTool {
    return {
        label: "Word Read",
        name: "read_document",
        description: "Read content from the current Word document. Use pagination for large docs.",
        parameters: WordReadSchema,
        execute: async (_toolCallId, args) => {
            if (options?.reverseRpc) {
                const result = await options.reverseRpc("read_document", args);
                return jsonResult(result);
            }
            // fallback (original behavior)
            return jsonResult({
                __CLIENT_ACTION__: "WORD_PLUGIN_ACTION",
                action: "read_document",
                args: args,
            });
        },
    };
}

export function createWordGrepTool(options?: {
    config?: OpenClawConfig;
    reverseRpc?: (action: string, args: unknown) => Promise<unknown>;
}): AnyAgentTool {
    return {
        label: "Word Grep",
        name: "grep_document",
        description: "Search for text patterns within the Word document.",
        parameters: WordGrepSchema,
        execute: async (_toolCallId, args) => {
            if (options?.reverseRpc) {
                const result = await options.reverseRpc("grep_document", args);
                return jsonResult(result);
            }
            return jsonResult({
                __CLIENT_ACTION__: "WORD_PLUGIN_ACTION",
                action: "grep_document",
                args: args,
            });
        },
    };
}

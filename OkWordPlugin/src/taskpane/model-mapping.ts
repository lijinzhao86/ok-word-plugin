/**
 * Maps frontend model IDs to OpenClaw Agent IDs
 */
export function modelToAgentId(modelId: string): string {
    const mapping: Record<string, string> = {
        "google/gemini-pro": "gemini-pro",
        "google/gemini-flash": "gemini-flash",
        "anthropic/claude-3-5-sonnet": "claude",
        "openai/gpt-4o": "gpt4",
        "qwen-plus": "qwen",
    };
    return mapping[modelId] || "gemini-pro"; // fallback to gemini-pro
}

/* global fetch, console */
/**
 * Service to communicate with OpenClaw backend.
 */
export class AIService {
    private gatewayToken: string;
    private baseUrl: string = "/openclaw-api";

    constructor(token: string) {
        this.gatewayToken = token;
    }

    /**
     * Maps frontend model IDs to OpenClaw Agent IDs
     */
    private mapModelToAgent(model: string): string {
        const mapping: Record<string, string> = {
            "google/gemini-pro": "openclaw:gemini-pro",
            "google/gemini-flash": "openclaw:gemini-flash",
            "anthropic/claude-3-5-sonnet": "openclaw:claude",
            "openai/gpt-4o": "openclaw:gpt4",
            "qwen-plus": "openclaw:qwen",
        };
        return mapping[model] || "openclaw:gemini-pro";
    }

    /**
     * Sends a prompt and content to OpenClaw and returns the result.
     */
    /**
     * Sends a prompt and content to OpenClaw and returns the result.
     */
    async rewrite(prompt: string, content: string, model: string): Promise<string> {
        return this.streamRewrite(prompt, content, model, () => { });
    }

    /**
     * Sends a prompt and content to OpenClaw and streams the result back.
     */
    async streamRewrite(
        prompt: string,
        content: string,
        model: string,
        onChunk: (chunk: string) => void
    ): Promise<string> {
        let fullText = "";
        try {
            const response = await fetch(`${this.baseUrl}/v1/chat/completions`, {
                method: "POST",
                headers: {
                    Authorization: `Bearer ${this.gatewayToken}`,
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    messages: [
                        {
                            role: "system",
                            content: "You are a helpful assistant that rewrites text while preserving the meaning.",
                        },
                        { role: "user", content: `${prompt}\n\nText to rewrite:\n${content}` },
                    ],
                    model: this.mapModelToAgent(model),
                    stream: true,
                }),
            });

            if (!response.ok) throw new Error(`OpenClaw error: ${response.statusText}`);

            const reader = response.body?.getReader();
            if (!reader) throw new Error("Response body is not readable");

            const decoder = new TextDecoder("utf-8");
            let buffer = "";

            while (true) {
                const { done, value } = await reader.read();
                if (done) break;

                const chunk = decoder.decode(value, { stream: true });
                buffer += chunk;
                const lines = buffer.split("\n");

                // Keep the last partial line in the buffer
                buffer = lines.pop() || "";

                for (const line of lines) {
                    const trimmedLine = line.trim();
                    if (!trimmedLine || trimmedLine === "data: [DONE]") continue;

                    if (trimmedLine.startsWith("data: ")) {
                        try {
                            const dataStr = trimmedLine.slice(6);


                            const data = JSON.parse(dataStr);
                            const delta = data.choices[0]?.delta;

                            if (delta?.custom_type) {

                                // Inject ID from root object so frontend can correlate events
                                const eventPayload = { ...delta, id: data.id };
                                onChunk(`__TOOL_EVENT__:${JSON.stringify(eventPayload)}`);
                                continue;
                            }

                            const content = delta?.content || "";
                            if (content) {
                                fullText += content;
                                onChunk(content);
                            }
                        } catch (e) {
                            console.warn("Failed to parse SSE JSON:", e);
                        }
                    }
                }
            }
            return fullText;
        } catch (error) {
            console.error("Failed to stream from OpenClaw:", error);
            throw error;
        }
    }

    /**
     * Updates OpenClaw API Keys via RPC.
     */
    async updateConfig(apiKey: string, qwenApiKey: string): Promise<void> {
        try {
            const getResponse = await fetch(`${this.baseUrl}/rpc`, {
                method: "POST",
                headers: {
                    Authorization: `Bearer ${this.gatewayToken}`,
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    jsonrpc: "2.0",
                    method: "config.get",
                    params: {},
                    id: "get-config",
                }),
            });

            if (!getResponse.ok) throw new Error("Failed to get config hash");
            const getData = await getResponse.json();
            const configHash = getData.result.hash;

            const patchResponse = await fetch(`${this.baseUrl}/rpc`, {
                method: "POST",
                headers: {
                    Authorization: `Bearer ${this.gatewayToken}`,
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    jsonrpc: "2.0",
                    method: "config.patch",
                    params: {
                        raw: JSON.stringify({
                            env: {
                                vars: {
                                    ANTHROPIC_API_KEY: apiKey,
                                    GOOGLE_GENERATIVE_AI_API_KEY: apiKey,
                                    OPENAI_API_KEY: apiKey,
                                    DASHSCOPE_API_KEY: qwenApiKey
                                }
                            },
                        }),
                        baseHash: configHash,
                    },
                    id: "patch-config",
                }),
            });

            if (!patchResponse.ok) throw new Error("Failed to patch config");
            console.log("OpenClaw configuration updated successfully.");
        } catch (error) {
            console.error("Failed to update OpenClaw config:", error);
            throw error;
        }
    }
}

/* global fetch, console */
/**
 * Service to communicate with OpenClaw backend.
 */
import { executeClientTool } from "./tools/word-implementation";

/**
 * Service to communicate with OpenClaw backend.
 */
export class AIService {
    private gatewayToken: string;
    private baseUrl: string = "/openclaw-api";
    private sessionKey: string;
    private socket: WebSocket | null = null;
    private isConnected: boolean = false;
    private reconnectTimer: any = null;

    constructor(token: string) {
        this.gatewayToken = token;
        // Generate or load a stable session key for this plugin instance
        const savedKey = localStorage.getItem("openclaw_session_key");
        const baseKey = savedKey || ("word-plugin-" + Math.random().toString(36).substring(7));
        if (!savedKey) {
            localStorage.setItem("openclaw_session_key", baseKey);
        }

        // Force prefixing if missing to ensure routing to qwen agent
        this.sessionKey = baseKey.startsWith("agent:") ? baseKey : `agent:qwen:${baseKey}`;

        console.log(`[AIService][V2.4] Resolved Session Key: ${this.sessionKey}`);
        this.initWebSocket();
    }

    private initWebSocket() {
        if (this.reconnectTimer) {
            clearTimeout(this.reconnectTimer);
            this.reconnectTimer = null;
        }
        if (this.socket) {
            this.socket.onclose = null; // Prevent onclose firing during intentional close
            this.socket.close();
        }

        const protocol = window.location.protocol === "https:" ? "wss:" : "ws:";
        const host = window.location.host;
        // Use the proxied path /openclaw-api/ for WebSocket as well
        const wsUrl = `${protocol}//${host}/openclaw-api/`;

        console.log(`[AIService][V2.1] Connecting to WebSocket: ${wsUrl} (Session: ${this.sessionKey})`);
        this.socket = new WebSocket(wsUrl);

        this.socket.onopen = () => {
            console.log("[AIService] WebSocket Connected");
            const handshake = {
                type: "req",
                method: "connect",
                id: "handshake-" + Date.now(),
                params: {
                    minProtocol: 1,
                    maxProtocol: 100,
                    client: {
                        id: "webchat-ui",
                        version: "1.0.0",
                        platform: "browser",
                        mode: "webchat"
                    },
                    auth: {
                        token: this.gatewayToken
                    },
                    sessionKey: this.sessionKey
                }
            };
            console.log("[AIService][V2.3] Sending handshake:", handshake);
            this.socket?.send(JSON.stringify(handshake));
            this.isConnected = true;
        };

        this.socket.onmessage = async (event) => {
            try {
                const data = JSON.parse(event.data);
                if (data.type === "req" && data.method === "client.tool.invoke") {
                    const { action, args } = data.params;
                    console.log(`[AIService] Received reverse tool call: ${action}`, args);

                    try {
                        const result = await executeClientTool(action, args);
                        console.log(`[AIService] Tool ${action} finished, sending result`);
                        this.socket?.send(JSON.stringify({
                            type: "res",
                            id: data.id,
                            ok: true,
                            payload: result
                        }));
                    } catch (error: any) {
                        console.error(`[AIService] Tool ${action} failed:`, error);
                        this.socket?.send(JSON.stringify({
                            type: "res",
                            id: data.id,
                            ok: false,
                            error: { message: error.message || "Unknown tool error" }
                        }));
                    }
                }
            } catch (e) {
                console.error("[AIService] Error handling WebSocket message:", e);
            }
        };

        this.socket.onclose = (e) => {
            console.log("[AIService] WebSocket Closed", e.code, e.reason);
            this.isConnected = false;
            // Reconnect after 3 seconds if not intentionally closed
            this.reconnectTimer = setTimeout(() => this.initWebSocket(), 3000);
        };

        this.socket.onerror = (err) => {
            console.error("[AIService] WebSocket Error:", err);
        };
    }

    /**
     * Closes the service connections.
     */
    public close() {
        if (this.reconnectTimer) {
            clearTimeout(this.reconnectTimer);
            this.reconnectTimer = null;
        }
        if (this.socket) {
            this.socket.onclose = null;
            this.socket.close();
            this.socket = null;
        }
        this.isConnected = false;
    }

    /**
     * Maps frontend model IDs to OpenClaw Agent IDs
     */
    private mapModelToAgentId(model: string): string {
        const mapping: Record<string, string> = {
            "google/gemini-pro": "gemini-pro",
            "google/gemini-flash": "gemini-flash",
            "anthropic/claude-3-5-sonnet": "claude",
            "openai/gpt-4o": "gpt4",
            "qwen-plus": "qwen",
        };
        return mapping[model] || "qwen";
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
        const agentId = this.mapModelToAgentId(model);

        // Final Session Key: agent:qwen:word-plugin-xxx
        const cleanBaseKey = this.sessionKey.replace(/^agent:[^:]+:/, "");
        const activeSessionKey = `agent:${agentId}:${cleanBaseKey}`;

        console.log(`[AIService] Routing to Agent: ${agentId}, Session: ${activeSessionKey}`);

        try {
            const requestBody = {
                messages: [
                    {
                        role: "system",
                        content: "You are a helpful assistant that rewrites text while preserving the meaning.",
                    },
                    {
                        role: "user",
                        content: content ? `${prompt}\n\nText to rewrite:\n${content}` : prompt
                    },
                ],
                // Model ID for request resolution: openclaw:qwen
                model: `openclaw:${agentId}`,
                stream: true,
            };

            const response = await fetch(`${this.baseUrl}/v1/chat/completions`, {
                method: "POST",
                headers: {
                    Authorization: `Bearer ${this.gatewayToken}`,
                    "Content-Type": "application/json",
                    "X-OpenClaw-Session-Key": activeSessionKey
                },
                body: JSON.stringify(requestBody),
            });

            console.log("[AIService] Request sent. Status:", response.status);
            // console.log("[AIService] Full Body Preview:", prompt.substring(0, 50) + "...");


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
                            // console.log("[AIService] Raw Chunk:", dataStr.substring(0, 100) + "..."); 

                            const data = JSON.parse(dataStr);
                            const delta = data.choices[0]?.delta;

                            console.log("[AIService] Parsed Delta:", JSON.stringify(delta, null, 2));

                            // 1. Tool Call Start/Delta
                            if (delta.tool_calls) {
                                const toolCall = delta.tool_calls[0];
                                if (toolCall) {
                                    // Make sure we capture the ID if present (usually only in first chunk)
                                    // and the index to correlate chunks
                                    const eventPayload = {
                                        custom_type: "tool_call",
                                        index: toolCall.index,
                                        id: toolCall.id, // Only present in first chunk
                                        name: toolCall.function?.name,
                                        args: toolCall.function?.arguments
                                    };
                                    onChunk(`__TOOL_EVENT__:${JSON.stringify(eventPayload)}`);
                                    continue;
                                }
                            }

                            // 2. Tool Result / Custom Events from OpenClaw backend
                            if (delta.custom_type) {
                                // console.log("🔥 [AIService] Tool Event Detected:", delta);
                                // OpenClaw backend sends results with custom_type
                                const eventPayload = { id: data.id, ...delta };
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
                            console.log("Faulty Line:", trimmedLine);
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
     * Resets the session with a new Key.
     */
    public startNewSession(): string {
        const newKey = "word-plugin-" + Math.random().toString(36).substring(7);
        // Ensure prefix
        this.sessionKey = `agent:qwen:${newKey}`;
        console.log(`[AIService] Starting New Session: ${this.sessionKey}`);

        // Update LocalStorage
        localStorage.setItem("openclaw_session_key", newKey);

        // Re-initialize WebSocket with new session
        this.initWebSocket();
        return this.sessionKey;
    }

    /**
     * Helper to call Gateway JSON-RPC
     */
    private async callRpc(method: string, params: any = {}): Promise<any> {
        const rpcBody = {
            jsonrpc: "2.0",
            method: method,
            params: params,
            id: Date.now().toString()
        };

        const response = await fetch(`${this.baseUrl}/rpc`, {
            method: "POST",
            headers: {
                Authorization: `Bearer ${this.gatewayToken}`,
                "Content-Type": "application/json",
            },
            body: JSON.stringify(rpcBody),
        });

        if (!response.ok) {
            throw new Error(`RPC HTTP Error: ${response.status} ${response.statusText}`);
        }

        const data = await response.json();
        if (data.error) {
            throw new Error(`RPC Error: ${data.error.message || JSON.stringify(data.error)}`);
        }
        return data.result;
    }

    /**
     * Retrieves chat history for the current session.
     */
    async getHistory(): Promise<any[]> {
        try {
            // Note: The sessionKey needs to match what is used in chat/completions
            // Our sessionKey is internally stored with 'agent:qwen:' prefix, 
            // but the RPC might expect the full key or let the backend handle it.
            // Let's try passing the full sessionKey we use.
            const result = await this.callRpc("chat.history", {
                sessionKey: this.sessionKey
            });
            return Array.isArray(result) ? result : [];
        } catch (error) {
            console.error("[AIService] Failed to fetch history:", error);
            return [];
        }
    }

    /**
     * Updates OpenClaw API Keys via RPC.
     */
    async updateConfig(apiKey: string, qwenApiKey: string): Promise<void> {
        try {
            const getData = await this.callRpc("config.get", {});
            const configHash = getData.hash;

            await this.callRpc("config.patch", {
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
            });

            console.log("OpenClaw configuration updated successfully.");
        } catch (error) {
            console.error("Failed to update OpenClaw config:", error);
            throw error;
        }
    }
}

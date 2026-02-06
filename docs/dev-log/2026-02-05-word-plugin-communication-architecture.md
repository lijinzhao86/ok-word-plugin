# Word Plugin Communication & Session Management Debugging Log
**Date**: 2026-02-05
**Status**: Architecture Consensus Reached

## 1. Problem Description
The Word plugin experienced several communication issues between the Taskpane (Frontend) and OpenClaw (Backend):
- **Infinite Loops**: AI would repeatedly call tools or output "NO" when receiving tool results.
- **Out-of-order Responses**: Old AI text segments would mix with new ones after a tool execution.
- **Context Loss**: Each request (initial prompt vs tool result feedback) was treated as a completely new, stateless session, causing the AI to lose the original user request.
- **Unclean Logs**: Tool results were being recorded as "User Message" stubs containing JSON, making the conversation history messy.

## 2. Root Cause Analysis
- **Stateless Gateway**: The OpenClaw HTTP Gateway (`/v1/chat/completions`) generated a new UUID for every request that lacked an explicit `X-OpenClaw-Session-Key` header.
- **Execution Boundary**: OpenClaw's internal tools (like bash/fs) run on the backend (Node.js). Word tools must run in the browser (Office.js). This physical separation creates a "Network Gap".
- **Immediate Return (Stub Tool)**: To trigger the frontend, the backend `word-tool.ts` immediately returns a JSON instruction. OpenClaw registers this as a "completed" tool turn, even though the real operation just started on the client.

## 3. Investigated Solutions

### Option A: Blocking Backend Process (Rejected)
- **Idea**: Make `word-tool.ts` `await` until the client returns data.
- **Pros**: Clean logs (one tool call = one result), standard ReAct look.
- **Cons**: 
    - **High Complexity**: Requires building a persistent callback/event-bus system in the Gateway.
    - **Timeout Risk**: HTTP connections typically timeout after 60s. Client-side Word operations + user interaction might exceed this, crashing the request.

### Option B: Session ID + Distributed Loop (Final Decision)
- **Idea**: Use a consistent Session ID to link multiple HTTP requests into a single logical "Thought Stream".
- **Pros**:
    - **Infrastructure Compatible**: Respects the stateless nature of OpenClaw's OpenAI-compatible HTTP Gateway.
    - **Reliable**: No risk of hung connections.
    - **Clean Semantics**: By using `role: tool` for the second request, we preserve the proper Agent history even if it spans across two HTTP turns.

## 4. Current Architecture Consensus

1.  **Session Linkage**: Frontend will generate a unique `sessionId` and send it via `X-OpenClaw-Session-Key` header.
2.  **Tool Role**: feedback from the client will be sent with `role: "tool"` and the corresponding tool name. 
    *   *Note*: OpenClaw's HTTP adapter (`openai-http.ts`) converts these to text prompts without strict ID checking, allowing us to safely use a dummy `tool_call_id`.
3.  **Command Pattern**:
    - `User`: "Summarize the doc."
    - `Assistant`: Calls `read_document`.
    - `Tool (Stub)`: Returns JSON Instruction (Recorded as Command Log).
    - `Frontend`: Executes Word API.
    - `Tool (Real)`: Frontend sends actual content back to Backend using the SAME Session ID.
    - `Assistant`: Sees the full history and provides the final summary.

## 5. Next Implementation Steps
- [ ] Modify `ai-service.ts` to manage `sessionId` state and include the header in all fetches.
- [ ] Introduce a `streamToolResult` method in `ai-service.ts` to support `role: "tool"` messaging.
- [ ] Update `taskpane.ts` to use the new session-aware messaging for all interactions.

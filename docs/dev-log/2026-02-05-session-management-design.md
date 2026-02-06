# Session Management System Design

**Date**: 2026-02-05
**Status**: Proposal / In Progress
**Context**: Implement a Cursor-like session management experience in the OkWord Word Plugin.

## 1. Overview
The goal is to transition the Word Plugin from a "Single Persistent Session" model to a "Multi-Session" model, similar to Cursor or ChatGPT. Users should be able to create new chats, view history, and switch between sessions seamlessly.

## 2. User Interface Design

### 2.1 Header Refactoring
The top navigation bar (`header`) will be redesigned to house utility controls.

**Layout**:
- **Left**: `OkWord` Branding/Title.
- **Right** (Action Group):
    - `[+]` **New Chat**: Immediately clears the current view and generates a new Session Key.
    - `[🕒]` **History**: Opens the Session History Modal.
    - `[...]` **More**: Dropdown menu containing "Settings" (migrated from the standalone button) and future utilities.
    - `[×]` **Close**: Closes the side pane container (`Office.context.ui.closeContainer()`).

### 2.2 History Modal (Center Overlay)
- **Visuals**: A centered floating modal with a glassmorphism (semi-transparent blur) backdrop, covering the chat area.
- **Content**: A scrollable list of past sessions.
- **Item Design**:
    - **Title**: Derived from the session summary (OpenClaw) or the first user message.
    - **Time**: Relative timestamp (e.g., "Just now", "2 hrs ago").
- **Interaction**: Clicking an item closes the modal, switches the active `sessionKey`, and triggers a history reload.

## 3. Technical Architecture

### 3.1 Backend Capability (OpenClaw)
Detailed analysis of the OpenClaw Gateway source code reveals built-in JSON-RPC endpoints that natively support this feature. No backend code changes are required.

**Key APIs**:
*   `sessions.list`: Returns a list of all sessions for the current agent/user. (Requires `READ` scope - **OK**)
*   `chat.history`: Returns the full array of messages for a specific `sessionKey`. (Requires `READ` scope - **OK**)
*   `sessions.delete`: Deletes a session. (Requires `ADMIN` scope - **Blocked** for now, will omit delete button).

### 3.2 Frontend Architecture

#### `AIService` Extensions
We need to extend `AIService` to support generic Gateway RPC calls over HTTP.

```typescript
// Proposed RPC Method
async callGatewayRpc(method: string, params: any = {}) {
    // POST /openclaw-api/rpc
    // Body: { jsonrpc: "2.0", method, params, id }
}

// New High-Level Methods
async listSessions() { ... }
async loadSessionHistory(key: string) { ... }
startNewSession() { ... } // Generates new UUID key
```

#### rendering Logic Refactoring
**Challenge**: Current tool cards (e.g., "Analyzed document structure") are rendered during the streaming process using specific chunk markers. Static history JSON from `chat.history` does not contain these markers.

**Solution: High-Fidelity History Restoration**
We will implement a `HistoryRenderer` that:
1.  Iterates through the static message history.
2.  Detects `tool_calls` and `tool_result` messages.
3.  **Simulates** the frontend event pipeline by calling `handleToolEvent` directly with the stored data.
4.  This ensures that historical tool usage remains visually consistent (rendered as cards) rather than reverting to raw JSON text.

## 4. Implementation Plan

### Phase 1: API & Data Layer
1.  Implement `callGatewayRpc` in `AIService`.
2.  Verify `sessions.list` and `chat.history` return expected data structure.
3.  Implement `startNewSession` logic (UUID generation + state reset).

### Phase 2: UI Construction
1.  Update `taskpane.html` structure for the new Header.
2.  Create `taskpane.css` rules for the Centered Modal and Glassmorphism effects.
3.  Implement the "More" dropdown menu logic.

### Phase 3: Logic Wire-up
1.  Bind "New Chat" button to clearing UI and resetting state.
2.  Bind "History" button to fetching list and rendering Modal.
3.  Implement **History Replay**:
    - Fetch history -> Clear Chat -> Loop & Render Messages.
    - specialized handling for markdown and tool events.

## 5. Future Considerations
- **Delete Session**: Once we upgrade the Token scope to Admin, enable the delete button in the history list.
- **Search**: Add a search bar to the History Modal to filter sessions by title/content.

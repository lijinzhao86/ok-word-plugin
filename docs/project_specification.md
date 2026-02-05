# Word AI Plugin Project Specification

This document outlines the architecture and implementation strategy for the Word AI Plugin, using **OpenClaw** as the backend AI agent and **Office.js** for the frontend Word integration.

## 1. Project Overview

The objective is to create a professional Word add-in that leverages AI for document drafting, rewriting, and refinement. A core requirement is to provide a seamless human-in-the-loop experience using Word's native track changes (revision mode) for AI-suggested edits.

## 2. Technical Architecture

### 2.1 Backend: OpenClaw
- **Role:** AI Agentic Control Plane.
- **Service:** Local Node.js service running on port `18789`.
- **Interface:** Communicates via REST API (secured with a Gateway Token).
- **Core Features:** Utilizing OpenClaw's Memory and Skills to provide contextual AI responses.

### 2.2 Frontend: Word Office Add-in
- **Framework:** Office.js (HTML/TypeScript).
- **UI:** Task Pane integration (recommended: Fluent UI for native look).
- **Data Exchange Format:** Markdown (for AI understanding) and HTML (for document rendering).

## 3. Core Workflow: AI Rewriting & Review

To ensure AI suggestions are human-reviewable and formatting is preserved, the following "Double-Track Diff" strategy will be used:

### Step 1: Context Extraction
- **Format:** Fetch selection via `Office.CoercionType.Html`.
- **Processing:** Convert HTML to **Markdown** (using `turndown`) to provide clean, structured text to the AI, reducing token noise and costs.

### Step 2: AI Processing (OpenClaw)
- Send Markdown to OpenClaw.
- OpenClaw performs reasoning and returns the optimized **Markdown**.

### Step 3: Granular Diff Calculation
- **Logic:** Instead of replacing the entire selection (which results in a messy "delete all, insert all" visual), the plugin calculates a granular diff between the original text and the AI output using a library like `diff-match-patch`.
- **Output:** A sequence of instructions: `Keep`, `Insert`, `Delete`.

### Step 4: Instructional Writing (Native Track Changes)
1. **Enable Track Changes:** Programmatically enable Word's revision mode:
   `context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;`
2. **Apply Changes:** Execute the diff instructions sequentially:
   - `Keep`: Move selection/cursor.
   - `Delete`: Delete the specific word/character (appears as a red strike-through).
   - `Insert`: Insert the new content (appears as a green underlined text).
3. **Format Recovery:** During insertion, parse basic Markdown (bold, italic) and apply the corresponding Word styles to the inserted range.

## 4. Key Advantages

- **Formatting Preservation:** By using granular text-level edits, we avoid overwriting Word's complex document metadata and styles.
- **Superior UI/UX:** Users see precisely what the AI changed (word-by-word) rather than having to re-read the entire paragraph.
- **Privacy First:** Since OpenClaw runs locally, sensitive document data remains within the user's controlled environment.

## 5. Potential Challenges & Solutions

- **CORS:** Ensure OpenClaw allows requests from the Office Add-in origins.
- **Complex Objects:** For tables or images, the plugin should detect their presence and either handle them via specialized logic or inform the user that only textual content is being optimized.
- **Performance:** For very large documents, chunking or limiting AI scope to the active selection is recommended.

---
*Status: Initial Research & Architecture Finalized. Ready for Prototype development.*

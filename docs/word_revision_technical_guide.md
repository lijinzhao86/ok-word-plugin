# Word Revision (Track Changes) Technical Guide

This document describes the implementation details and testing procedures for the **Granular Revision** feature in the Word AI Plugin.

## 1. Overview

The goal of this feature is to provide a "word-by-word" or "phrase-by-phrase" revision experience in Word. By using granular diffs and programmatic control over Word's **Track Changes** mode, the AI can show precisely what it changed (deletions and insertions) rather than replacing entire paragraphs.

## 2. Core Implementation Strategy

To achieve interwoven "red-line" edits, we follow these steps:

1.  **Calculate Diff**: Use `diff-match-patch` to get a list of instructions (`KEEP`, `DELETE`, `INSERT`).
2.  **Enable Tracking**: Set `context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;`.
3.  **Sequential Execution**: Iterate through instructions while maintaining a "Remaining Range" pointer.
    -   **KEEP**: Search for the text within the `remainingRange`, then update `remainingRange` to start after the found text.
    -   **DELETE**: Search for the text within the `remainingRange`, call `.delete()` on the match, and update `remainingRange` to start after the deletion point.
    -   **INSERT**: Call `.insertText(text, "Start")` on the `remainingRange` and update `remainingRange` to start after the new text.

## 3. Why we use `remainingRange`

Directly moving a cursor or searching the entire document body for every step can lead to errors (e.g., finding the same word earlier in the document). By always narrowing the search to the **unprocessed portion** of the document, we ensure that:
1.  Edits are applied in the correct sequence.
2.  Multiple occurrences of the same word are handled correctly.
3.  Positioning remains precise even as the document length changes during edits.

## 4. Testing & Verification

You can test this logic directly in the Word Add-in developer console using the script below.

### 4.1 Preparation
1. Open the Word Add-in.
2. Open the browser/taskpane developer tools (F12).
3. Ensure you have a blank document or one you are willing to clear for testing.

### 4.2 Test Script (Granular Revision Simulator)

```javascript
/* 
 * This script will:
 * 1. Clear the document and insert a base text.
 * 2. Enable Track Changes.
 * 3. Apply fine-grained edits word-by-word.
 */
(async () => {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            
            // Step 1: Initialize Clean Environment
            context.document.changeTrackingMode = "Off";
            body.clear();
            const baseText = "The quick brown fox jumps over the lazy dog. It was a suny day in the forest and everything felt peaceful.";
            body.insertText(baseText, "Start");
            await context.sync();

            // Step 2: Enable Track Changes
            context.document.changeTrackingMode = "TrackAll";

            // Step 3: Define Simulated AI Diff Rules
            const instructions = [
                { type: "keep", text: "The " },
                { type: "delete", text: "quick" },
                { type: "insert", text: "fast" },
                { type: "keep", text: " brown fox " },
                { type: "delete", text: "jumps" },
                { type: "insert", text: "leaped" },
                { type: "keep", text: " over the lazy dog. It was a " },
                { type: "delete", text: "suny" },
                { type: "insert", text: "sunny" },
                { type: "keep", text: " day in the " },
                { type: "insert", text: "lush " },
                { type: "keep", text: "forest and everything felt peaceful." }
            ];

            // Maintain a pointer to the "unprocessed" part of the document
            let remainingRange = body.getRange();

            for (const part of instructions) {
                if (part.type === "keep") {
                    let searchRes = remainingRange.search(part.text);
                    searchRes.load("items");
                    await context.sync();
                    if (searchRes.items.length > 0) {
                        // Move remainingRange to start AFTER the kept text
                        remainingRange = searchRes.items[0].getRange("After").expandTo(body.getRange("End"));
                    }
                } 
                else if (part.type === "insert") {
                    // Insert at the START of the remaining area
                    let insertedRange = remainingRange.insertText(part.text, "Start");
                    await context.sync();
                    // Move remainingRange to start AFTER the inserted text
                    remainingRange = insertedRange.getRange("After").expandTo(body.getRange("End"));
                } 
                else if (part.type === "delete") {
                    let searchRes = remainingRange.search(part.text);
                    searchRes.load("items");
                    await context.sync();
                    if (searchRes.items.length > 0) {
                        let target = searchRes.items[0];
                        let nextPoint = target.getRange("After");
                        target.delete(); // Mark as deleted in Track Changes
                        // Move remainingRange to start AFTER the deleted word
                        remainingRange = nextPoint.expandTo(body.getRange("End"));
                    }
                }
                await context.sync();
            }
            
            console.log("Granular Revision Test Completed Successfully!");
        });
    } catch (e) {
        console.error("Test execution failed:", e);
    }
})();
```

## 5. Visual Reference

When correctly executed, the Word document will show:
- **Red Strikethroughs** on the "quick", "jumps", and "suny".
- **Red Underlines** on the "fast", "leaped", "sunny", and "lush".
- The changes are interwoven (e.g., `fast` appears right after the struck-through `quick`).

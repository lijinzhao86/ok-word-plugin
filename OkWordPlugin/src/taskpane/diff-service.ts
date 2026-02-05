import { diff_match_patch, DIFF_DELETE, DIFF_EQUAL } from "diff-match-patch";

export interface DiffInstruction {
    type: "keep" | "delete" | "insert";
    text: string;
}

/**
 * Calculates a granular diff between two strings and returns a sequence of instructions.
 */
export function getGranularDiff(oldText: string, newText: string): DiffInstruction[] {
    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(oldText, newText);
    dmp.diff_cleanupSemantic(diffs);

    return diffs.map(([type, text]) => {
        let instructionType: "keep" | "delete" | "insert";
        // diff-match-patch uses -1: delete, 1: insert, 0: equal
        if (type === DIFF_EQUAL) instructionType = "keep";
        else if (type === DIFF_DELETE) instructionType = "delete";
        else instructionType = "insert";

        return {
            type: instructionType,
            text: text,
        };
    });
}

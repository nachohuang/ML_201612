/** Shared cell -> plain-text extraction, used by both the table (zh/en column) and
 * freeform (whole-sheet cell-by-cell) Excel comparison paths. */

export function cellText(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "object") {
    const obj = value as Record<string, unknown>;
    // Rich-text cells (mixed fonts/colors within one cell, common in free-form
    // requirement docs) come back as { richText: [{text, font}, ...] } rather than a
    // plain string — concatenate the runs to get the cell's actual visible text.
    if (Array.isArray(obj.richText)) {
      return obj.richText.map((run) => String((run as { text?: unknown }).text ?? "")).join("");
    }
    if ("text" in obj) return String(obj.text ?? "");
    if ("result" in obj) return String(obj.result ?? "");
  }
  return String(value);
}

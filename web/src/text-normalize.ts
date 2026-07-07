/** Mirrors diffing/text_normalize.py. String.prototype.normalize("NFKC") follows the
 * same Unicode standard as Python's unicodedata.normalize, so behavior matches exactly. */

export function normalize(text: string | null | undefined): string {
  if (text == null) return "";
  return text.normalize("NFKC").trim().split(/\s+/).filter(Boolean).join(" ");
}

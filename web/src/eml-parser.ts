/** Parses Outlook's .eml export format (plain-text RFC 822/MIME, not the binary OLE
 * .msg format) to auto-fill sender/date instead of asking the user to type them in.
 * .eml is just text with "Header: value" lines followed by a blank line and the body,
 * so this is a small hand-rolled parser rather than a dependency — no need for a
 * full MIME library just to read three headers.
 */

export interface EmailMeta {
  subject: string | null;
  sender: string | null;
  receivedDate: string | null; // YYYY-MM-DD if the Date header parsed, else null
}

/** Decodes MIME "encoded-word" syntax (=?charset?B?...?= or ?Q?...?=), which is how
 * non-ASCII display names (e.g. Chinese sender names) typically show up in From/Subject
 * headers written by Outlook. */
function decodeMimeWords(text: string): string {
  return text.replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, (whole, charset: string, enc: string, encoded: string) => {
    try {
      let bytes: Uint8Array;
      if (enc.toUpperCase() === "B") {
        const binary = atob(encoded);
        bytes = Uint8Array.from(binary, (c) => c.charCodeAt(0));
      } else {
        const withSpaces = encoded.replace(/_/g, " ");
        const values: number[] = [];
        for (let i = 0; i < withSpaces.length; i++) {
          if (withSpaces[i] === "=" && i + 2 < withSpaces.length) {
            values.push(parseInt(withSpaces.slice(i + 1, i + 3), 16));
            i += 2;
          } else {
            values.push(withSpaces.charCodeAt(i));
          }
        }
        bytes = Uint8Array.from(values);
      }
      return new TextDecoder(charset).decode(bytes);
    } catch {
      return whole; // unrecognized charset/encoding — leave the raw token as-is
    }
  });
}

function unfoldHeaderLines(headerBlock: string): string[] {
  const lines = headerBlock.split(/\r\n|\r|\n/);
  const unfolded: string[] = [];
  for (const line of lines) {
    if (/^[ \t]/.test(line) && unfolded.length > 0) {
      unfolded[unfolded.length - 1] += " " + line.trim();
    } else if (line.trim() !== "") {
      unfolded.push(line);
    }
  }
  return unfolded;
}

function getHeader(lines: string[], name: string): string | null {
  const prefix = `${name.toLowerCase()}:`;
  for (const line of lines) {
    if (line.toLowerCase().startsWith(prefix)) return line.slice(prefix.length).trim();
  }
  return null;
}

/** "陳宜筠" <a@b.com> / 陳宜筠 <a@b.com> / a@b.com -> a readable display string. */
function extractDisplayName(fromHeader: string): string {
  const match = fromHeader.match(/^"?([^"<]*)"?\s*<([^>]+)>$/);
  if (!match) return fromHeader.trim();
  const name = match[1].trim();
  return name || match[2];
}

function formatDate(dateHeader: string): string | null {
  const parsed = new Date(dateHeader);
  if (Number.isNaN(parsed.getTime())) return null;
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
}

export function parseEml(rawText: string): EmailMeta {
  const headerEndIdx = rawText.search(/\r?\n\r?\n/);
  const headerBlock = headerEndIdx === -1 ? rawText : rawText.slice(0, headerEndIdx);
  const lines = unfoldHeaderLines(headerBlock);

  const subjectRaw = getHeader(lines, "subject");
  const fromRaw = getHeader(lines, "from");
  const dateRaw = getHeader(lines, "date");

  return {
    subject: subjectRaw ? decodeMimeWords(subjectRaw) : null,
    sender: fromRaw ? extractDisplayName(decodeMimeWords(fromRaw)) : null,
    receivedDate: dateRaw ? formatDate(dateRaw) : null,
  };
}

export async function parseEmlFile(file: Blob): Promise<EmailMeta> {
  return parseEml(await file.text());
}

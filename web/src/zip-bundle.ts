/** Packages the run's output files into one downloadable zip (Path B has no folder
 * access, so there's no "copy into archive/output" step — the user unzips and files
 * things away manually; see the generated README entry for exact instructions). */

import type JSZipNS from "jszip";

export interface ZipEntry {
  name: string;
  bytes: ArrayBuffer;
}

export async function buildZip(JSZip: typeof JSZipNS, entries: ZipEntry[]): Promise<Blob> {
  const zip = new JSZip();
  for (const entry of entries) {
    zip.file(entry.name, entry.bytes);
  }
  return zip.generateAsync({ type: "blob" });
}

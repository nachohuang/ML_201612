/** Cell-by-cell comparison for sheets that aren't a simple zh/en table — free-form
 * screen mockups, report layouts with merged cells, etc. Real requirement documents
 * often have far more of these than actual tables, and they're still meaningful
 * requirement content that must be tracked for changes (see plan Context), just not
 * ones with a clean "one column is Chinese, one is English" structure to translate.
 *
 * Every non-blank cell address in either version is compared directly; a changed cell
 * is reported for human review same as a table-mode row, but — unlike table-mode
 * Segments — there is no separate English column to write a translation into, so these
 * records intentionally leave newEnLocationId unset. The translation-table filter in
 * app.ts already only picks up records that have one, so freeform changes show up in
 * the diff report but never demand a translation input.
 */

import type ExcelJSNS from "exceljs";

import { cellText } from "./cell-text.ts";
import { ChangeType } from "./models.ts";
import type { ChangeRecord } from "./models.ts";

export function diffFreeformSheet(
  oldWs: ExcelJSNS.Worksheet | undefined,
  newWs: ExcelJSNS.Worksheet
): ChangeRecord[] {
  const records: ChangeRecord[] = [];
  const maxRow = Math.max(oldWs?.rowCount ?? 0, oldWs?.actualRowCount ?? 0, newWs.rowCount, newWs.actualRowCount);
  const maxCol = Math.max(oldWs?.columnCount ?? 0, newWs.columnCount);

  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const newCell = newWs.getRow(r).getCell(c);
      const newText = cellText(newCell.value).trim();
      const oldText = oldWs ? cellText(oldWs.getRow(r).getCell(c).value).trim() : "";
      if (oldText === newText) continue;

      const address = newCell.address;
      if (oldText === "") {
        records.push({ locationId: `${newWs.name}!${address}`, fieldName: null, changeType: ChangeType.ADDED, newZh: newText });
      } else if (newText === "") {
        records.push({
          locationId: `${(oldWs as ExcelJSNS.Worksheet).name}!${address}`,
          fieldName: null,
          changeType: ChangeType.DELETED,
          oldZh: oldText,
        });
      } else {
        records.push({
          locationId: `${newWs.name}!${address}`,
          fieldName: null,
          changeType: ChangeType.MODIFIED,
          oldZh: oldText,
          newZh: newText,
        });
      }
    }
  }
  return records;
}

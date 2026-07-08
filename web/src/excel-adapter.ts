/** Mirrors file_types/excel_adapter.py, backed by ExcelJS instead of openpyxl.
 *
 * Real customer workbooks have multiple tabs with genuinely different layouts, so a
 * mapping is applied per worksheet name (Record<sheetName, SheetMapping>) rather than
 * once for "the" sheet — every configured sheet that exists in the workbook gets its
 * own segments extracted, and they're merged into one flat list (each Segment's
 * location already carries its sheet name, so diff reports/translation tables can
 * still tell which tab a change came from).
 *
 * The ExcelJS namespace is passed in rather than imported at the module level: in the
 * browser it's loaded as a global via <script src="libs/exceljs.min.js"> (this tool
 * ships zero bundled copies of exceljs so the static files stay small and cacheable),
 * while unit tests running under Node inject the real npm package. Only `import type`
 * is used below, which esbuild/tsc erase completely, so no bundler ever needs to
 * actually resolve the "exceljs" module at build time.
 */

import type ExcelJSNS from "exceljs";

import { cellText } from "./cell-text.ts";
import type { SheetMapping } from "./config.ts";
import type { Segment, SegmentUpdate } from "./models.ts";
import { toArrayBuffer } from "./xlsx-buffer.ts";

export class ExcelAdapter {
  private readonly ExcelJS: typeof ExcelJSNS;
  private readonly sheetMappings: Record<string, SheetMapping>;

  constructor(ExcelJS: typeof ExcelJSNS, sheetMappings: Record<string, SheetMapping>) {
    this.ExcelJS = ExcelJS;
    this.sheetMappings = sheetMappings;
  }

  async load(bytes: ArrayBuffer): Promise<ExcelJSNS.Workbook> {
    const wb = new this.ExcelJS.Workbook();
    await wb.xlsx.load(bytes);
    return wb;
  }

  /** `referenceWb` is only used as a fallback when `wb` doesn't have a sheet with the
   * mapping's exact name — real customer workbooks sometimes rename tabs between
   * versions (e.g. stripping an English suffix) while keeping the same tab order, so
   * this falls back to matching by position within `referenceWb` (typically the newly
   * uploaded document, since mappings are configured against its sheet names) rather
   * than silently dropping the whole sheet from the diff. */
  extractSegments(wb: ExcelJSNS.Workbook, referenceWb?: ExcelJSNS.Workbook): Segment[] {
    const segments: Segment[] = [];
    for (const mapping of Object.values(this.sheetMappings)) {
      if (mapping.mode !== "table") continue; // freeform sheets go through diffFreeformSheet instead
      const ws = resolveWorksheet(wb, mapping.sheetName, referenceWb);
      if (!ws) continue; // this workbook doesn't have that tab (even by position) — skip rather than fail
      segments.push(...extractFromSheet(ws, mapping));
    }
    return segments;
  }

  applyTranslations(wb: ExcelJSNS.Workbook, updates: SegmentUpdate[]): void {
    for (const update of updates) {
      const sepIdx = update.enLocationId.indexOf("!");
      const sheetName = update.enLocationId.slice(0, sepIdx);
      const cellRef = update.enLocationId.slice(sepIdx + 1);
      const ws = wb.getWorksheet(sheetName);
      if (!ws) throw new Error(`找不到工作表: ${sheetName}`);
      ws.getCell(cellRef).value = update.enText;
    }
  }

  async save(wb: ExcelJSNS.Workbook): Promise<ArrayBuffer> {
    const buffer = await wb.xlsx.writeBuffer();
    return toArrayBuffer(buffer as ArrayBuffer | ArrayBufferView);
  }
}

/** `referenceWb` is only used as a fallback when `wb` doesn't have a sheet with this
 * exact name — real customer workbooks sometimes rename tabs between versions (e.g.
 * stripping an English suffix) while keeping the same tab order, so this falls back to
 * matching by position within `referenceWb` (typically the newly uploaded document,
 * since mappings/freeform sheet choices are made against its sheet names) rather than
 * silently dropping the whole sheet from the diff. Exported so the freeform diff path
 * in app.ts (which isn't driven by ExcelAdapter) can resolve sheets the same way. */
export function resolveWorksheet(
  wb: ExcelJSNS.Workbook,
  sheetName: string,
  referenceWb?: ExcelJSNS.Workbook
): ExcelJSNS.Worksheet | undefined {
  const exact = wb.getWorksheet(sheetName);
  if (exact) return exact;
  if (!referenceWb) return undefined;
  const refIndex = referenceWb.worksheets.findIndex((w) => w.name === sheetName);
  if (refIndex === -1) return undefined;
  return wb.worksheets[refIndex];
}

function extractFromSheet(ws: ExcelJSNS.Worksheet, mapping: SheetMapping): Segment[] {
  const headerRow = mapping.headerRow;
  const headers = new Map<string, unknown>();
  for (const col of mapping.zhColumns) {
    headers.set(col, ws.getCell(`${col}${headerRow}`).value);
  }

  const segments: Segment[] = [];
  const lastRow = lastNonBlankRow(ws, headerRow, mapping);

  for (let row = headerRow + 1; row <= lastRow; row++) {
    const zhValues = new Map<string, unknown>();
    for (const col of mapping.zhColumns) {
      zhValues.set(col, ws.getCell(`${col}${row}`).value);
    }

    let rowKey: string | null = null;
    if (mapping.keyColumns.length > 0) {
      const keyParts = mapping.keyColumns.map((col) => cellText(ws.getCell(`${col}${row}`).value).trim());
      rowKey = keyParts.some((p) => p !== "") ? keyParts.join("|") : null;
    }

    const rowHasContent = [...zhValues.values()].some((v) => cellText(v).trim() !== "");
    if (!rowHasContent && rowKey === null) continue; // fully blank row

    for (let i = 0; i < mapping.zhColumns.length; i++) {
      const zhCol = mapping.zhColumns[i];
      const enCol = mapping.enColumns[i];
      segments.push({
        locationId: `${ws.name}!${zhCol}${row}`,
        enLocationId: `${ws.name}!${enCol}${row}`,
        rowKey,
        fieldName: headers.get(zhCol) != null ? cellText(headers.get(zhCol)) : null,
        zhText: cellText(zhValues.get(zhCol)),
        enText: cellText(ws.getCell(`${enCol}${row}`).value),
      });
    }
  }
  return segments;
}

/** ExcelJS's worksheet.rowCount can include trailing rows with only style/no value, so
 * we scan for the last row that actually has content in a mapped column instead of
 * trusting it blindly — mirrors the Python adapter iterating to ws.max_row and letting
 * the blank-row check drop trailing whitespace rows. */
function lastNonBlankRow(ws: ExcelJSNS.Worksheet, headerRow: number, mapping: SheetMapping): number {
  const candidateCols = [...mapping.zhColumns, ...mapping.keyColumns];
  let last = headerRow;
  const scanLimit = Math.max(ws.rowCount, ws.actualRowCount, headerRow);
  for (let row = headerRow + 1; row <= scanLimit; row++) {
    const hasContent = candidateCols.some((col) => cellText(ws.getCell(`${col}${row}`).value).trim() !== "");
    if (hasContent) last = row;
  }
  return last;
}

export interface InspectColumn {
  column: string;
  header: unknown;
  samples: unknown[];
  hasFormula: boolean;
}

export interface InspectSheet {
  sheetName: string;
  columns: InspectColumn[];
}

/** Lists every worksheet's columns (header + sample values), used to power the
 * point-and-click mapping UI instead of asking users to type raw column letters. */
export async function inspectWorkbook(
  ExcelJS: typeof ExcelJSNS,
  bytes: ArrayBuffer,
  opts: { headerRow?: number; sampleRows?: number } = {}
): Promise<InspectSheet[]> {
  const headerRow = opts.headerRow ?? 1;
  const sampleRows = opts.sampleRows ?? 3;

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);

  return wb.worksheets.map((ws) => {
    const columns: InspectColumn[] = [];
    for (let colIdx = 1; colIdx <= ws.columnCount; colIdx++) {
      const columnLetter = ws.getColumn(colIdx).letter;
      const header = ws.getRow(headerRow).getCell(colIdx).value;
      const samples: unknown[] = [];
      let hasFormula = false;
      for (let row = headerRow + 1; row < headerRow + 1 + sampleRows; row++) {
        const cell = ws.getRow(row).getCell(colIdx);
        if (cell.type === ExcelJS.ValueType.Formula) hasFormula = true;
        samples.push(cell.value);
      }
      columns.push({ column: columnLetter, header, samples, hasFormula });
    }
    return { sheetName: ws.name, columns };
  });
}

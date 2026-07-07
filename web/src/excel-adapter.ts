/** Mirrors file_types/excel_adapter.py, backed by ExcelJS instead of openpyxl.
 *
 * The ExcelJS namespace is passed in rather than imported at the module level: in the
 * browser it's loaded as a global via <script src="libs/exceljs.min.js"> (this tool
 * ships zero bundled copies of exceljs so the static files stay small and cacheable),
 * while unit tests running under Node inject the real npm package. Only `import type`
 * is used below, which esbuild/tsc erase completely, so no bundler ever needs to
 * actually resolve the "exceljs" module at build time.
 */

import type ExcelJSNS from "exceljs";

import type { ExcelMapping } from "./config.ts";
import type { Segment, SegmentUpdate } from "./models.ts";
import { toArrayBuffer } from "./xlsx-buffer.ts";

export class ExcelAdapter {
  private readonly ExcelJS: typeof ExcelJSNS;
  private readonly mapping: ExcelMapping;

  constructor(ExcelJS: typeof ExcelJSNS, mapping: ExcelMapping) {
    this.ExcelJS = ExcelJS;
    this.mapping = mapping;
  }

  async load(bytes: ArrayBuffer): Promise<ExcelJSNS.Workbook> {
    const wb = new this.ExcelJS.Workbook();
    await wb.xlsx.load(bytes);
    return wb;
  }

  private sheet(wb: ExcelJSNS.Workbook): ExcelJSNS.Worksheet {
    const ws = this.mapping.sheetName === "auto" ? wb.worksheets[0] : wb.getWorksheet(this.mapping.sheetName);
    if (!ws) throw new Error(`找不到工作表: ${this.mapping.sheetName}`);
    return ws;
  }

  extractSegments(wb: ExcelJSNS.Workbook): Segment[] {
    const ws = this.sheet(wb);
    const headerRow = this.mapping.headerRow;
    const headers = new Map<string, unknown>();
    for (const col of this.mapping.zhColumns) {
      headers.set(col, ws.getCell(`${col}${headerRow}`).value);
    }

    const segments: Segment[] = [];
    const lastRow = lastNonBlankRow(ws, headerRow, this.mapping);

    for (let row = headerRow + 1; row <= lastRow; row++) {
      const zhValues = new Map<string, unknown>();
      for (const col of this.mapping.zhColumns) {
        zhValues.set(col, ws.getCell(`${col}${row}`).value);
      }

      let rowKey: string | null = null;
      if (this.mapping.keyColumns.length > 0) {
        const keyParts = this.mapping.keyColumns.map((col) => cellText(ws.getCell(`${col}${row}`).value).trim());
        rowKey = keyParts.some((p) => p !== "") ? keyParts.join("|") : null;
      }

      const rowHasContent = [...zhValues.values()].some((v) => cellText(v).trim() !== "");
      if (!rowHasContent && rowKey === null) continue; // fully blank row

      for (let i = 0; i < this.mapping.zhColumns.length; i++) {
        const zhCol = this.mapping.zhColumns[i];
        const enCol = this.mapping.enColumns[i];
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

function cellText(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "object" && "text" in (value as Record<string, unknown>)) {
    return String((value as Record<string, unknown>).text ?? "");
  }
  if (typeof value === "object" && "result" in (value as Record<string, unknown>)) {
    return String((value as Record<string, unknown>).result ?? "");
  }
  return String(value);
}

/** ExcelJS's worksheet.rowCount can include trailing rows with only style/no value, so
 * we scan for the last row that actually has content in a mapped column instead of
 * trusting it blindly — mirrors the Python adapter iterating to ws.max_row and letting
 * the blank-row check drop trailing whitespace rows. */
function lastNonBlankRow(ws: ExcelJSNS.Worksheet, headerRow: number, mapping: ExcelMapping): number {
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

export async function inspectExcel(
  ExcelJS: typeof ExcelJSNS,
  bytes: ArrayBuffer,
  opts: { sheetName?: string; headerRow?: number; sampleRows?: number } = {}
): Promise<InspectColumn[]> {
  const sheetName = opts.sheetName ?? "auto";
  const headerRow = opts.headerRow ?? 1;
  const sampleRows = opts.sampleRows ?? 3;

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = sheetName === "auto" ? wb.worksheets[0] : wb.getWorksheet(sheetName);
  if (!ws) throw new Error(`找不到工作表: ${sheetName}`);

  const result: InspectColumn[] = [];
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
    result.push({ column: columnLetter, header, samples, hasFormula });
  }
  return result;
}

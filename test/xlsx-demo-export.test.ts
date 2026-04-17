import { describe, expect, it } from "vitest";
import type { WorksheetPivotTableDefinition } from "@flexsheet/core";
import { parseCellRef } from "../packages/import-export/src/a1.js";
import { pivotOutputExtents } from "../packages/import-export/src/export-xlsx-pivot.js";
import { exportWorkbookToXlsxBytes, importXlsxToWorkbook, unzipToMap } from "@flexsheet/import-export";
import { createDemoWorkbook } from "../packages/flexsheet/src/demo/demo-workbook.js";
import { CreatePivotTableCommand } from "../packages/flexsheet/src/pivot/pivot-table-command.js";

function parseA1Range(ref: string): { minR: number; maxR: number; minC: number; maxC: number } {
  const parts = ref.split(":");
  const a = parseCellRef(parts[0]!.trim());
  const b = parseCellRef((parts[1] ?? parts[0])!.trim());
  if (a === null || b === null) {
    throw new Error(`bad ref ${ref}`);
  }
  return {
    minR: Math.min(a.row, b.row),
    maxR: Math.max(a.row, b.row),
    minC: Math.min(a.col, b.col),
    maxC: Math.max(a.col, b.col),
  };
}

function findSheetByName(wb: ReturnType<typeof createDemoWorkbook>, name: string) {
  for (let i = 0; i < wb.sheetCount; i++) {
    const sh = wb.getSheet(i);
    if (sh?.name === name) {
      return sh;
    }
  }
  return undefined;
}

describe("demo workbook xlsx export", () => {
  it("pivotOutputExtents clamps zero row/col span so dimension can match pivot location", () => {
    const stub = {
      destinationRow: 5,
      destinationCol: 2,
      outputRowCount: 0,
      outputColCount: 0,
    } as unknown as WorksheetPivotTableDefinition;
    expect(pivotOutputExtents(stub)).toEqual({
      minR: 5,
      maxR: 5,
      minC: 2,
      maxC: 2,
    });
  });

  it("every XML entry parses as well-formed and import roundtrips (no pivot)", async () => {
    const wb = createDemoWorkbook();
    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const parser = new DOMParser();
    for (const [path, u8] of map) {
      if (!path.endsWith(".xml")) {
        continue;
      }
      const text = new TextDecoder().decode(u8);
      const doc = parser.parseFromString(text, "application/xml");
      const pe = doc.getElementsByTagName("parsererror");
      expect(pe.length, `parse ${path}`).toBe(0);
    }
    await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
  });

  it("every XML entry parses with pivot sheet like demo main.ts", async () => {
    const wb = createDemoWorkbook();
    const pivotSource = findSheetByName(wb, "透视数据源");
    expect(pivotSource).toBeDefined();
    const cmd = new CreatePivotTableCommand(wb, pivotSource!, {
      sourceRange: { startRow: 0, endRow: 6, startCol: 0, endCol: 2 },
      hasHeaders: true,
      rowFieldCols: [0],
      columnFieldCols: [],
      filterFieldCols: [],
      valueFields: [{ col: 2, aggregate: "sum" }],
      destination: { kind: "newSheet", preferredName: "透视表示例" },
    });
    cmd.execute();
    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const parser = new DOMParser();
    for (const [path, u8] of map) {
      if (!path.endsWith(".xml")) {
        continue;
      }
      const text = new TextDecoder().decode(u8);
      const doc = parser.parseFromString(text, "application/xml");
      const pe = doc.getElementsByTagName("parsererror");
      expect(pe.length, `parse ${path}`).toBe(0);
    }
    await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    const pivotSheetXml = new TextDecoder().decode(map.get("xl/worksheets/sheet8.xml"));
    expect(pivotSheetXml).toContain("<pivotTables ");
    expect(pivotSheetXml).toContain("<row ");
    expect(pivotSheetXml).not.toMatch(/<sheetData>\s*<\/sheetData>/);

    const dimM = /dimension ref="([^"]+)"/.exec(pivotSheetXml);
    expect(dimM).not.toBeNull();
    const pivotTableXml = new TextDecoder().decode(map.get("xl/pivotTables/pivotTable1.xml"));
    const locM = /location ref="([^"]+)"/.exec(pivotTableXml);
    expect(locM).not.toBeNull();
    const dimBox = parseA1Range(dimM![1]!);
    const locBox = parseA1Range(locM![1]!);
    expect(dimBox.minR).toBeLessThanOrEqual(locBox.minR);
    expect(dimBox.minC).toBeLessThanOrEqual(locBox.minC);
    expect(dimBox.maxR).toBeGreaterThanOrEqual(locBox.maxR);
    expect(dimBox.maxC).toBeGreaterThanOrEqual(locBox.maxC);
  });
});

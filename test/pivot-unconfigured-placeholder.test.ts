import { describe, expect, it } from "vitest";
import {
  Workbook,
  Worksheet,
  buildUnconfiguredPivotPlaceholderMatrix,
  isUnconfiguredPivotDefinition,
  writeUnconfiguredPivotPlaceholderToSheet,
} from "@flexsheet/core";
import { exportWorkbookToXlsxBytes, importXlsxToWorkbook } from "@flexsheet/import-export";

describe("pivot unconfigured placeholder", () => {
  it("buildUnconfiguredPivotPlaceholderMatrix fills title and hint", () => {
    const m = buildUnconfiguredPivotPlaceholderMatrix(4, 3, "数据透视表1");
    expect(m.rowCount).toBe(4);
    expect(m.colCount).toBe(3);
    expect(m.values[0]![0]).toBe("数据透视表1");
    expect(m.values[1]![0]).toBe("若要生成报表，请从数据透视表字段列表中选择字段");
    expect(m.values[3]![2]).toBe(null);
    expect(m.styles[0]![1]?.fillArgb).toBe("FFBDD7EE");
  });

  it("isUnconfiguredPivotDefinition detects empty layout", () => {
    const def = {
      id: "x",
      name: "P",
      sourceSheetIndex: 0,
      sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 0 },
      hasHeaders: true,
      rowFieldCols: [],
      columnFieldCols: [],
      filterFieldCols: [],
      filterSelectedKeys: [],
      valueFields: [],
      destinationRow: 2,
      destinationCol: 0,
      outputRowCount: 3,
      outputColCount: 2,
    };
    expect(isUnconfiguredPivotDefinition(def)).toBe(true);
    expect(
      isUnconfiguredPivotDefinition({
        ...def,
        rowFieldCols: [0],
      }),
    ).toBe(false);
  });

  it("writeUnconfiguredPivotPlaceholderToSheet writes at layout origin", () => {
    const wb = new Workbook();
    const sh = new Worksheet("S", 20, 8);
    wb.addSheet(sh);
    const def = {
      id: "p",
      name: "透视1",
      sourceSheetIndex: 0,
      sourceRange: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      hasHeaders: true,
      rowFieldCols: [],
      columnFieldCols: [],
      filterFieldCols: [],
      filterSelectedKeys: [],
      valueFields: [],
      destinationRow: 5,
      destinationCol: 1,
      outputRowCount: 3,
      outputColCount: 2,
    };
    writeUnconfiguredPivotPlaceholderToSheet(sh, def);
    expect(sh.getCell(5, 1).value).toBe("透视1");
    expect(sh.getCell(6, 1).value).toContain("字段列表");
  });

  it("imports xlsx with empty dataFields as registered pivot and roundtrips name", async () => {
    const wb = new Workbook();
    const src = new Worksheet("Data", 10, 6);
    wb.addSheet(src);
    src.setCellLiteral(0, 0, "A");
    src.setCellLiteral(0, 1, "B");
    src.setCellLiteral(1, 0, "x");
    src.setCellLiteral(1, 1, 1);
    const pivot = new Worksheet("Pivot", 24, 8);
    wb.addSheet(pivot);
    pivot.registerPivotTableDefinition({
      id: "empty-pvt",
      name: "数据透视表1",
      sourceSheetIndex: 0,
      sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
      hasHeaders: true,
      rowFieldCols: [],
      columnFieldCols: [],
      filterFieldCols: [],
      filterSelectedKeys: [],
      valueFields: [],
      destinationRow: 6,
      destinationCol: 0,
      outputRowCount: 10,
      outputColCount: 3,
    });
    const bytes = exportWorkbookToXlsxBytes(wb);
    const back = await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    const importedPivot = back.getSheet(1);
    expect(importedPivot).toBeDefined();
    const defs = importedPivot!.getPivotTableDefinitionsSnapshot();
    expect(defs.length).toBe(1);
    expect(defs[0]!.valueFields.length).toBe(0);
    expect(defs[0]!.rowFieldCols.length).toBe(0);
    expect(defs[0]!.name).toBe("数据透视表1");
    expect(importedPivot!.getCell(6, 0).value).toBe("数据透视表1");
    expect(String(importedPivot!.getCell(7, 0).value ?? "")).toContain("字段列表");
  });
});

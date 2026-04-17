import { describe, expect, it } from "vitest";
import { Workbook, Worksheet } from "@flexsheet/core";

describe("Workbook", () => {
  it("starts with no sheets and active index 0", () => {
    const wb = new Workbook();
    expect(wb.sheetCount).toBe(0);
    expect(wb.activeSheetIndex).toBe(0);
    expect(wb.getActiveSheet()).toBeUndefined();
  });

  it("addSheet appends and getSheet returns the sheet", () => {
    const wb = new Workbook();
    const s1 = new Worksheet("Sheet1");
    const s2 = new Worksheet("Sheet2");
    wb.addSheet(s1);
    wb.addSheet(s2);
    expect(wb.sheetCount).toBe(2);
    expect(wb.getSheet(0)).toBe(s1);
    expect(wb.getSheet(1)).toBe(s2);
  });

  it("getSheet returns undefined for out-of-range index", () => {
    const wb = new Workbook();
    wb.addSheet(new Worksheet("A"));
    expect(wb.getSheet(-1)).toBeUndefined();
    expect(wb.getSheet(1)).toBeUndefined();
  });

  it("getActiveSheet follows activeSheetIndex", () => {
    const wb = new Workbook();
    wb.addSheet(new Worksheet("A"));
    wb.addSheet(new Worksheet("B"));
    wb.activeSheetIndex = 1;
    expect(wb.getActiveSheet()?.name).toBe("B");
    wb.activeSheetIndex = 0;
    expect(wb.getActiveSheet()?.name).toBe("A");
  });

  it("getSheets exposes readonly list matching internal order", () => {
    const wb = new Workbook();
    const a = new Worksheet("A");
    wb.addSheet(a);
    const sheets = wb.getSheets();
    expect(sheets.length).toBe(1);
    expect(sheets[0]).toBe(a);
  });

  it("subscribe receives events when a sheet cell changes", () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    let n = 0;
    wb.subscribe(() => {
      n++;
    });
    s.setCellLiteral(0, 0, 1);
    expect(n).toBe(1);
  });

  it("subscribe fires on addSheet and when activeSheetIndex changes", () => {
    const wb = new Workbook();
    const events: string[] = [];
    wb.subscribe(() => {
      events.push("x");
    });
    wb.addSheet(new Worksheet("A"));
    wb.addSheet(new Worksheet("B"));
    expect(events.length).toBe(2);
    wb.activeSheetIndex = 1;
    expect(events.length).toBe(3);
    wb.activeSheetIndex = 1;
    expect(events.length).toBe(3);
  });

  it("activeSheetIndex setter clamps to valid range", () => {
    const wb = new Workbook();
    wb.addSheet(new Worksheet("A"));
    wb.addSheet(new Worksheet("B"));
    wb.activeSheetIndex = 99;
    expect(wb.activeSheetIndex).toBe(1);
    wb.activeSheetIndex = -3;
    expect(wb.activeSheetIndex).toBe(0);
  });

  it("removeSheetAt removes sheet and tightens active index", () => {
    const wb = new Workbook();
    wb.addSheet(new Worksheet("A"));
    wb.addSheet(new Worksheet("B"));
    wb.activeSheetIndex = 1;
    expect(wb.removeSheetAt(0)).toBe(true);
    expect(wb.sheetCount).toBe(1);
    expect(wb.getSheet(0)?.name).toBe("B");
    expect(wb.activeSheetIndex).toBe(0);
  });

  it("removeSheetAt returns false when only one sheet or bad index", () => {
    const wb = new Workbook();
    wb.addSheet(new Worksheet("Only"));
    expect(wb.removeSheetAt(0)).toBe(false);
    expect(wb.sheetCount).toBe(1);
    wb.addSheet(new Worksheet("Two"));
    expect(wb.removeSheetAt(-1)).toBe(false);
    expect(wb.removeSheetAt(2)).toBe(false);
  });

  it("insertSheetAt inserts at index and shifts activeSheetIndex when needed", () => {
    const wb = new Workbook();
    const a = new Worksheet("A");
    const b = new Worksheet("B");
    wb.addSheet(a);
    wb.addSheet(b);
    wb.activeSheetIndex = 1;
    const mid = new Worksheet("Mid");
    wb.insertSheetAt(1, mid);
    expect(wb.sheetCount).toBe(3);
    expect(wb.getSheet(1)).toBe(mid);
    expect(wb.activeSheetIndex).toBe(2);
  });

  it("insertSheetAt bumps pivotTableDefinition sourceSheetIndex at or after insert", () => {
    const wb = new Workbook();
    const s0 = new Worksheet("S0");
    const s1 = new Worksheet("S1");
    wb.addSheet(s0);
    wb.addSheet(s1);
    s1.registerPivotTableDefinition({
      id: "p1",
      name: "PivotTable1",
      sourceSheetIndex: 1,
      sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
      hasHeaders: true,
      rowFieldCols: [0],
      columnFieldCols: [],
      filterFieldCols: [],
      filterSelectedKeys: [],
      valueFields: [{ col: 1, aggregate: "sum" }],
      destinationRow: 0,
      destinationCol: 0,
      outputRowCount: 3,
      outputColCount: 2,
    });
    wb.insertSheetAt(1, new Worksheet("Pivot"));
    expect(s1.getPivotTableDefinitionsSnapshot()[0]?.sourceSheetIndex).toBe(2);
  });

  it("removeSheetAt drops pivot defs whose sourceSheetIndex equals removed index", () => {
    const wb = new Workbook();
    const s0 = new Worksheet("S0");
    const s1 = new Worksheet("S1");
    wb.addSheet(s0);
    wb.addSheet(s1);
    s0.registerPivotTableDefinition({
      id: "self-src",
      name: "PivotTable1",
      sourceSheetIndex: 0,
      sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
      hasHeaders: true,
      rowFieldCols: [0],
      columnFieldCols: [],
      filterFieldCols: [],
      filterSelectedKeys: [],
      valueFields: [{ col: 1, aggregate: "sum" }],
      destinationRow: 0,
      destinationCol: 0,
      outputRowCount: 1,
      outputColCount: 1,
    });
    wb.removeSheetAt(0);
    expect(s1.getPivotTableDefinitionsSnapshot().length).toBe(0);
  });

  it("removeSheetAt decrements pivot sourceSheetIndex when it was above removed index", () => {
    const wb = new Workbook();
    const s0 = new Worksheet("S0");
    const s1 = new Worksheet("S1");
    wb.addSheet(s0);
    wb.addSheet(s1);
    s1.registerPivotTableDefinition({
      id: "p",
      name: "PivotTable1",
      sourceSheetIndex: 1,
      sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
      hasHeaders: true,
      rowFieldCols: [0],
      columnFieldCols: [],
      filterFieldCols: [],
      filterSelectedKeys: [],
      valueFields: [{ col: 1, aggregate: "sum" }],
      destinationRow: 0,
      destinationCol: 0,
      outputRowCount: 1,
      outputColCount: 1,
    });
    wb.removeSheetAt(0);
    expect(s1.getPivotTableDefinitionsSnapshot()[0]?.sourceSheetIndex).toBe(0);
  });
});

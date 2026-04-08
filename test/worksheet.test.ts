import { describe, expect, it } from "vitest";
import { type Cell, Worksheet } from "@flexsheet/core";
import { setCellValueAndRecalc } from "@flexsheet/formula";

describe("Worksheet", () => {
  it("setName trims and notifies subscribers", () => {
    const ws = new Worksheet("Old");
    let n = 0;
    ws.subscribe(() => {
      n++;
    });
    ws.setName("  New  ");
    expect(ws.name).toBe("New");
    expect(n).toBe(1);
    ws.setName("New");
    expect(n).toBe(1);
    ws.setName("");
    expect(ws.name).toBe("New");
  });

  it("constructor sets name and dimensions with defaults", () => {
    const ws = new Worksheet("Data");
    expect(ws.name).toBe("Data");
    expect(ws.rowCount).toBe(1000);
    expect(ws.colCount).toBe(26);
    expect(ws.defaultRowHeight).toBe(20);
    expect(ws.defaultColWidth).toBe(64);
  });

  it("constructor accepts custom row and column counts", () => {
    const ws = new Worksheet("Small", 10, 5);
    expect(ws.rowCount).toBe(10);
    expect(ws.colCount).toBe(5);
  });

  it("constructor clamps row and column counts to at least 1", () => {
    const ws = new Worksheet("Tiny", 0, -3);
    expect(ws.rowCount).toBe(1);
    expect(ws.colCount).toBe(1);
  });

  it("hasCell is false until cell is materialized", () => {
    const ws = new Worksheet("S", 100, 10);
    expect(ws.hasCell(0, 0)).toBe(false);
    ws.getCell(0, 0);
    expect(ws.hasCell(0, 0)).toBe(true);
  });

  it("getCell returns same instance for the same address", () => {
    const ws = new Worksheet("S");
    const c1 = ws.getCell(3, 4);
    const c2 = ws.getCell(3, 4);
    expect(c1).toBe(c2);
    expect(c1.row).toBe(3);
    expect(c1.col).toBe(4);
  });

  it("setCellLiteral clears formula and sets scalar value", () => {
    const ws = new Worksheet("S");
    ws.setCellFormula(0, 0, "=1+1");
    ws.setCellLiteral(0, 0, 42);
    const cell = ws.getCell(0, 0);
    expect(cell.formula).toBeNull();
    expect(cell.value).toBe(42);
  });

  it("setCellFormula trims and stores formula string", () => {
    const ws = new Worksheet("S");
    ws.setCellFormula(1, 2, "  =SUM(A1:A2)  ");
    expect(ws.getCell(1, 2).formula).toBe("=SUM(A1:A2)");
  });

  it("setCellValue treats leading = as formula", () => {
    const ws = new Worksheet("S");
    ws.setCellValue(0, 0, "  =A1+B1  ");
    expect(ws.getCell(0, 0).formula).toBe("=A1+B1");
  });

  it("setCellValue uses literal for non-formula strings and other scalars", () => {
    const ws = new Worksheet("S");
    ws.setCellValue(0, 0, "plain");
    expect(ws.getCell(0, 0).formula).toBeNull();
    expect(ws.getCell(0, 0).value).toBe("plain");
    ws.setCellValue(1, 1, 3.14);
    expect(ws.getCell(1, 1).value).toBe(3.14);
    ws.setCellValue(2, 2, null);
    expect(ws.getCell(2, 2).value).toBeNull();
  });

  it("iterateCells visits only created cells", () => {
    const ws = new Worksheet("S");
    ws.setCellLiteral(0, 0, 1);
    ws.setCellLiteral(2, 2, 2);
    const keys: string[] = [];
    ws.iterateCells((c: Cell) => keys.push(`${c.row},${c.col}`));
    expect(keys.sort()).toEqual(["0,0", "2,2"]);
  });

  it("getCell alone does not notify subscribers", () => {
    const ws = new Worksheet("S");
    let n = 0;
    ws.subscribe(() => {
      n++;
    });
    ws.getCell(0, 0);
    expect(n).toBe(0);
  });

  it("setCellLiteral notifies and increments revision", () => {
    const ws = new Worksheet("S");
    let n = 0;
    ws.subscribe(() => {
      n++;
    });
    const r0 = ws.revision;
    ws.setCellLiteral(0, 0, "x");
    expect(n).toBe(1);
    expect(ws.revision).toBe(r0 + 1);
  });

  it("batch coalesces multiple writes into one notification", () => {
    const ws = new Worksheet("S");
    let n = 0;
    ws.subscribe(() => {
      n++;
    });
    ws.batch(() => {
      ws.setCellLiteral(0, 0, 1);
      ws.setCellLiteral(1, 1, 2);
    });
    expect(n).toBe(1);
  });

  it("setCellStyle updates cell and notifies", () => {
    const ws = new Worksheet("S");
    ws.setCellStyle(0, 0, { bold: true, fgArgb: "FFFF0000" });
    expect(ws.getCell(0, 0).style).toEqual({ bold: true, fgArgb: "FFFF0000" });
    ws.setCellStyle(0, 0, null);
    expect(ws.getCell(0, 0).style).toBeNull();
  });

  it("setGridSize updates dimensions when changed", () => {
    const ws = new Worksheet("S", 10, 10);
    ws.setGridSize(20, 15);
    expect(ws.rowCount).toBe(20);
    expect(ws.colCount).toBe(15);
  });

  it("Cell.isFormulaCell reflects formula field", () => {
    const ws = new Worksheet("S");
    ws.setCellLiteral(0, 0, 1);
    expect(ws.getCell(0, 0).isFormulaCell()).toBe(false);
    ws.setCellFormula(0, 0, "=1");
    expect(ws.getCell(0, 0).isFormulaCell()).toBe(true);
  });

  it("setCellValueAndRecalc emits once after write and recalc", () => {
    const ws = new Worksheet("S", 50, 10);
    let n = 0;
    ws.subscribe(() => {
      n++;
    });
    setCellValueAndRecalc(ws, 0, 0, 2);
    setCellValueAndRecalc(ws, 1, 0, "=A1*2");
    expect(ws.getCell(1, 0).value).toBe(4);
    expect(n).toBe(2);
  });

  it("mergeCenter clears indent and sets horizontal center on master", () => {
    const ws = new Worksheet("S");
    ws.getCell(0, 0).style = { indentLevel: 2 };
    ws.applyMergeForSelection({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }, "mergeCenter");
    expect(ws.getCell(0, 0).style?.indentLevel).toBeUndefined();
    expect(ws.getCell(0, 0).style?.hAlign).toBe("center");
  });
});

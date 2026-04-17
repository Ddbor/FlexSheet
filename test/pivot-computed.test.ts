import { describe, expect, it } from "vitest";
import { Worksheet } from "@flexsheet/core";
import { buildPivotRender } from "../packages/flexsheet/src/pivot/pivot-table-command.js";

describe("buildPivotRender computed value fields", () => {
  it("bucketRatio: CTR as sum(点击)/sum(展现) per group and in total", () => {
    const sheet = new Worksheet("S", 10, 8);
    sheet.setCellLiteral(0, 0, "场景");
    sheet.setCellLiteral(0, 1, "展现");
    sheet.setCellLiteral(0, 2, "点击");
    sheet.setCellLiteral(1, 0, "A");
    sheet.setCellLiteral(1, 1, 100);
    sheet.setCellLiteral(1, 2, 10);
    sheet.setCellLiteral(2, 0, "B");
    sheet.setCellLiteral(2, 1, 200);
    sheet.setCellLiteral(2, 2, 30);

    const out = buildPivotRender(sheet, {
      sourceRange: { startRow: 0, endRow: 2, startCol: 0, endCol: 2 },
      hasHeaders: true,
      rowFieldCols: [0],
      columnFieldCols: [],
      filterFieldCols: [],
      valueFields: [
        { col: 2, aggregate: "sum", computed: { kind: "bucketRatio", denominatorCol: 1 } },
      ],
      destination: { kind: "existingSheet", startRow: 0, startCol: 0 },
    });

    expect(out.values[1]?.[1]).toBeCloseTo(10 / 100, 10);
    expect(out.values[2]?.[1]).toBeCloseTo(30 / 200, 10);
    expect(out.values[3]?.[1]).toBeCloseTo(40 / 300, 10);
  });

  it("shareOfGrandTotal: cost share sums to 100% in total row", () => {
    const sheet = new Worksheet("S", 10, 8);
    sheet.setCellLiteral(0, 0, "场景");
    sheet.setCellLiteral(0, 1, "花费");
    sheet.setCellLiteral(1, 0, "A");
    sheet.setCellLiteral(1, 1, 50);
    sheet.setCellLiteral(2, 0, "B");
    sheet.setCellLiteral(2, 1, 70);

    const out = buildPivotRender(sheet, {
      sourceRange: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 },
      hasHeaders: true,
      rowFieldCols: [0],
      columnFieldCols: [],
      filterFieldCols: [],
      valueFields: [{ col: 1, aggregate: "sum", computed: { kind: "shareOfGrandTotal" } }],
      destination: { kind: "existingSheet", startRow: 0, startCol: 0 },
    });

    expect(out.values[1]?.[1]).toBeCloseTo(50 / 120, 10);
    expect(out.values[2]?.[1]).toBeCloseTo(70 / 120, 10);
    expect(out.values[3]?.[1]).toBeCloseTo(1, 10);
  });

  it("filterFieldCols: renders filter row and excludes non-matching source rows", () => {
    const sheet = new Worksheet("S", 12, 8);
    sheet.setCellLiteral(0, 0, "类");
    sheet.setCellLiteral(0, 1, "额");
    sheet.setCellLiteral(0, 2, "组");
    sheet.setCellLiteral(1, 0, "A");
    sheet.setCellLiteral(1, 1, 10);
    sheet.setCellLiteral(1, 2, "G1");
    sheet.setCellLiteral(2, 0, "B");
    sheet.setCellLiteral(2, 1, 20);
    sheet.setCellLiteral(2, 2, "G2");

    const out = buildPivotRender(sheet, {
      sourceRange: { startRow: 0, endRow: 2, startCol: 0, endCol: 2 },
      hasHeaders: true,
      rowFieldCols: [2],
      columnFieldCols: [],
      filterFieldCols: [0],
      filterSelectedKeys: [["A"]],
      valueFields: [{ col: 1, aggregate: "sum" }],
      destination: { kind: "existingSheet", startRow: 0, startCol: 0 },
    });

    expect(out.values[0]?.[0]).toBe("类");
    expect(out.values[0]?.[1]).toBe("A");
    expect(out.values[1]?.[0]).toBe("组");
    expect(String(out.values[1]?.[1])).toContain("额");
    expect(out.values[2]?.[0]).toBe("G1");
    expect(out.values[2]?.[1]).toBe(10);
    expect(out.values[3]?.[0]).toBe("总计");
    expect(out.values[3]?.[1]).toBe(10);
  });

  it("valueFieldsOnRows: stacks multiple measures as sub-rows with 汇总 column", () => {
    const sheet = new Worksheet("S", 20, 8);
    sheet.setCellLiteral(0, 0, "G");
    sheet.setCellLiteral(0, 1, "A");
    sheet.setCellLiteral(0, 2, "B");
    sheet.setCellLiteral(1, 0, "X");
    sheet.setCellLiteral(1, 1, 1);
    sheet.setCellLiteral(1, 2, 2);

    const out = buildPivotRender(sheet, {
      sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 2 },
      hasHeaders: true,
      rowFieldCols: [0],
      columnFieldCols: [],
      filterFieldCols: [],
      filterSelectedKeys: [],
      valueFields: [
        { col: 1, aggregate: "sum" },
        { col: 2, aggregate: "sum" },
      ],
      valueFieldsOnRows: true,
      destination: { kind: "existingSheet", startRow: 0, startCol: 0 },
    });

    expect(out.colCount).toBe(2);
    expect(out.rowCount).toBe(5);
    expect(out.values[0]?.[0]).toBe("G");
    expect(out.values[0]?.[1]).toBe("汇总");
    expect(out.values[1]?.[0]).toBe("X");
    expect(out.values[1]?.[1]).toBe(1);
    expect(out.values[2]?.[0]).toBe("　求和项:B");
    expect(out.values[2]?.[1]).toBe(2);
    expect(out.values[3]?.[0]).toBe("总计 求和项:A");
    expect(out.values[3]?.[1]).toBe(1);
    expect(out.values[4]?.[0]).toBe("总计 求和项:B");
    expect(out.values[4]?.[1]).toBe(2);
  });
});

import { describe, expect, it } from "vitest";
import { Workbook, Worksheet } from "@flexsheet/core";
import {
  CreatePivotTableCommand,
  refreshPivotTableDefinition,
} from "../packages/flexsheet/src/pivot/pivot-table-command.js";

describe("refreshPivotTableDefinition", () => {
  it("rebuilds pivot values when source cells change", () => {
    const wb = new Workbook();
    const source = new Worksheet("源数据", 20, 8);
    const pivot = new Worksheet("透视", 20, 8);
    wb.addSheet(source);
    wb.addSheet(pivot);

    source.setCellLiteral(0, 0, "类别");
    source.setCellLiteral(0, 1, "金额");
    source.setCellLiteral(1, 0, "A");
    source.setCellLiteral(1, 1, 10);
    source.setCellLiteral(2, 0, "B");
    source.setCellLiteral(2, 1, 20);

    const cmd = new CreatePivotTableCommand(wb, source, {
      sourceRange: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 },
      hasHeaders: true,
      rowFieldCols: [0],
      columnFieldCols: [],
      filterFieldCols: [],
      valueFields: [{ col: 1, aggregate: "sum" }],
      destination: { kind: "existingSheet", targetSheet: pivot, startRow: 0, startCol: 0 },
    });
    cmd.execute();

    expect(pivot.getCell(1, 1).value).toBe(10);
    expect(pivot.getCell(3, 1).value).toBe(30);

    source.setCellLiteral(1, 1, 15);
    const refreshed = refreshPivotTableDefinition(wb, pivot, cmd.getPivotDefinitionId());

    expect(refreshed).toBe(true);
    expect(pivot.getCell(1, 1).value).toBe(15);
    expect(pivot.getCell(3, 1).value).toBe(35);
  });
});

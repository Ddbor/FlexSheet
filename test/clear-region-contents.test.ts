import { describe, expect, it } from "vitest";
import { CommandManager, Worksheet } from "@flexsheet/core";
import { ClearRegionContentsCommand } from "@flexsheet/formula";

describe("ClearRegionContentsCommand", () => {
  it("clears values and formulas in range but keeps styles", () => {
    const sheet = new Worksheet("s", 5, 5);
    sheet.setCellValue(0, 0, "a");
    sheet.setCellStyle(0, 0, { bold: true });
    sheet.setCellFormula(0, 1, "=1+1");
    sheet.setCellValue(1, 0, 42);

    const cmd = new ClearRegionContentsCommand(sheet, {
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });
    expect(cmd.hasChanges).toBe(true);
    cmd.execute();

    expect(sheet.getCell(0, 0).value).toBeNull();
    expect(sheet.getCell(0, 0).formula).toBeNull();
    expect(sheet.getCell(0, 0).style).toEqual({ bold: true });

    expect(sheet.getCell(0, 1).formula).toBeNull();
    expect(sheet.getCell(0, 1).value).toBeNull();

    expect(sheet.getCell(1, 0).value).toBeNull();
  });

  it("undo restores content", () => {
    const sheet = new Worksheet("s", 3, 3);
    sheet.setCellValue(0, 0, "x");
    const mgr = new CommandManager();
    const cmd = new ClearRegionContentsCommand(sheet, {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    });
    mgr.execute(cmd);
    expect(sheet.getCell(0, 0).value).toBeNull();
    expect(mgr.undo()).toBe(true);
    expect(sheet.getCell(0, 0).value).toBe("x");
  });

  it("hasChanges false when range already empty", () => {
    const sheet = new Worksheet("s", 2, 2);
    const cmd = new ClearRegionContentsCommand(sheet, {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    });
    expect(cmd.hasChanges).toBe(false);
  });
});

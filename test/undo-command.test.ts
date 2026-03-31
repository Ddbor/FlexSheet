import { CommandManager, Worksheet } from "@flexsheet/core";
import { recalcWorksheet, SetCellValueCommand } from "@flexsheet/formula";
import { describe, expect, it } from "vitest";

describe("SetCellValueCommand + CommandManager", () => {
  it("undo/redo restores cell content", () => {
    const sheet = new Worksheet("S", 10, 10);
    const mgr = new CommandManager();
    mgr.execute(new SetCellValueCommand(sheet, 0, 0, "a"));
    expect(sheet.getCell(0, 0).value).toBe("a");
    mgr.undo();
    expect(sheet.getCell(0, 0).value).toBeNull();
    mgr.redo();
    expect(sheet.getCell(0, 0).value).toBe("a");
  });

  it("undo formula cell restores prior formula", () => {
    const sheet = new Worksheet("S", 10, 10);
    recalcWorksheet(sheet);
    const mgr = new CommandManager();
    mgr.execute(new SetCellValueCommand(sheet, 1, 0, "=1+1"));
    expect(sheet.getCell(1, 0).formula).toBe("=1+1");
    mgr.undo();
    expect(sheet.getCell(1, 0).formula).toBeNull();
  });
});

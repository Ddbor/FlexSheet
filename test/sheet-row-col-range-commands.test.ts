import { describe, expect, it } from "vitest";
import { CommandManager, Worksheet } from "@flexsheet/core";
import { SetColWidthsInRangeCommand, SetRowHeightsInRangeCommand } from "flexsheet";

describe("SetRowHeightsInRangeCommand / SetColWidthsInRangeCommand", () => {
  it("sets one undo step and restores previous heights", () => {
    const sheet = new Worksheet("S", 5, 3);
    const cm = new CommandManager();
    sheet.setRowHeight(1, 18);
    sheet.setRowHeight(2, 22);
    cm.execute(new SetRowHeightsInRangeCommand(sheet, 1, 2, 40));
    expect(sheet.getRowHeight(1)).toBe(40);
    expect(sheet.getRowHeight(2)).toBe(40);
    expect(cm.undo()).toBe(true);
    expect(sheet.getRowHeight(1)).toBe(18);
    expect(sheet.getRowHeight(2)).toBe(22);
  });

  it("sets one undo step and restores previous widths", () => {
    const sheet = new Worksheet("S", 3, 5);
    const cm = new CommandManager();
    sheet.setColWidth(0, 50);
    sheet.setColWidth(1, 70);
    cm.execute(new SetColWidthsInRangeCommand(sheet, 0, 1, 100));
    expect(sheet.getColWidth(0)).toBe(100);
    expect(sheet.getColWidth(1)).toBe(100);
    expect(cm.undo()).toBe(true);
    expect(sheet.getColWidth(0)).toBe(50);
    expect(sheet.getColWidth(1)).toBe(70);
  });
});

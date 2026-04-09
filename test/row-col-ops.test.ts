import { CommandManager, Worksheet } from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";
import { SelectionModel } from "@flexsheet/selection";
import { describe, expect, it } from "vitest";

import {
  DeleteCellsShiftLeftCommand,
  DeleteCellsShiftUpCommand,
  DeleteColsCommand,
  DeleteRowsCommand,
  InsertCellsShiftDownCommand,
  InsertCellsShiftRightCommand,
  InsertColsCommand,
  InsertRowsCommand,
  SetColHiddenCommand,
  SetColWidthCommand,
  SetRowHeightCommand,
  SetRowHiddenCommand,
} from "../packages/flexsheet/src/sheet-structure-commands";

describe("row/col structural commands", () => {
  it("insert/delete rows should shift formula and support undo redo", () => {
    const sheet = new Worksheet("S1", 10, 10);
    const selection = new SelectionModel(() => sheet);
    const cm = new CommandManager();
    sheet.setCellLiteral(0, 0, 1);
    sheet.setCellFormula(1, 0, "=A1");
    recalcWorksheet(sheet);
    expect(sheet.getCell(1, 0).value).toBe(1);

    cm.execute(new InsertRowsCommand(sheet, selection, 0, 1));
    expect(sheet.getCell(2, 0).formula).toBe("=A2");
    cm.undo();
    expect(sheet.getCell(1, 0).formula).toBe("=A1");
    cm.redo();
    expect(sheet.getCell(2, 0).formula).toBe("=A2");

    cm.execute(new DeleteRowsCommand(sheet, selection, 0, 1));
    expect(sheet.getCell(1, 0).formula).toBe("=A1");
    cm.undo();
    expect(sheet.getCell(2, 0).formula).toBe("=A2");
  });

  it("insert/delete cols should shift formula and support undo redo", () => {
    const sheet = new Worksheet("S1", 10, 10);
    const selection = new SelectionModel(() => sheet);
    const cm = new CommandManager();
    sheet.setCellLiteral(0, 0, 1);
    sheet.setCellFormula(0, 1, "=A1");
    recalcWorksheet(sheet);

    cm.execute(new InsertColsCommand(sheet, selection, 0, 1));
    expect(sheet.getCell(0, 2).formula).toBe("=B1");
    cm.execute(new DeleteColsCommand(sheet, selection, 0, 1));
    expect(sheet.getCell(0, 1).formula).toBe("=A1");
  });

  it("hide and resize commands should be undoable", () => {
    const sheet = new Worksheet("S1", 10, 10);
    const cm = new CommandManager();
    cm.execute(new SetRowHiddenCommand(sheet, 2, true));
    cm.execute(new SetColHiddenCommand(sheet, 3, true));
    cm.execute(new SetRowHeightCommand(sheet, 1, 40));
    cm.execute(new SetColWidthCommand(sheet, 1, 120));
    expect(sheet.isRowHidden(2)).toBe(true);
    expect(sheet.isColHidden(3)).toBe(true);
    expect(sheet.getRowHeight(1)).toBe(40);
    expect(sheet.getColWidth(1)).toBe(120);
    cm.undo();
    cm.undo();
    expect(sheet.getRowHeight(1)).toBe(sheet.defaultRowHeight);
    expect(sheet.getColWidth(1)).toBe(sheet.defaultColWidth);
  });

  it("insert cells shift right/down should support undo redo", () => {
    const sheet = new Worksheet("S1", 4, 4);
    const selection = new SelectionModel(() => sheet);
    const cm = new CommandManager();
    sheet.setCellLiteral(1, 1, "B2");
    sheet.setCellLiteral(1, 2, "C2");
    sheet.setCellLiteral(2, 1, "B3");
    selection.selectCell(1, 1);

    cm.execute(new InsertCellsShiftRightCommand(sheet, selection, 1, 1, 1));
    expect(sheet.getCell(1, 1).value).toBeNull();
    expect(sheet.getCell(1, 2).value).toBe("B2");
    expect(sheet.getCell(1, 3).value).toBe("C2");
    cm.undo();
    expect(sheet.getCell(1, 1).value).toBe("B2");
    expect(sheet.getCell(1, 2).value).toBe("C2");
    cm.redo();
    expect(sheet.getCell(1, 2).value).toBe("B2");

    cm.execute(new InsertCellsShiftDownCommand(sheet, selection, 1, 1, 1));
    expect(sheet.getCell(1, 1).value).toBeNull();
    expect(sheet.getCell(2, 1).value).toBeNull();
    expect(sheet.getCell(3, 1).value).toBe("B3");
    cm.undo();
    expect(sheet.getCell(2, 1).value).toBe("B3");
  });

  it("delete cells shift left/up should support undo redo", () => {
    const sheet = new Worksheet("S1", 4, 4);
    const selection = new SelectionModel(() => sheet);
    const cm = new CommandManager();
    sheet.setCellLiteral(1, 1, "B2");
    sheet.setCellLiteral(1, 2, "C2");
    sheet.setCellLiteral(1, 3, "D2");
    selection.selectCell(1, 1);

    cm.execute(
      new DeleteCellsShiftLeftCommand(sheet, selection, {
        startRow: 1,
        startCol: 1,
        endRow: 1,
        endCol: 1,
      }),
    );
    expect(sheet.getCell(1, 1).value).toBe("C2");
    expect(sheet.getCell(1, 2).value).toBe("D2");
    expect(sheet.getCell(1, 3).value).toBeNull();
    cm.undo();
    expect(sheet.getCell(1, 1).value).toBe("B2");
    expect(sheet.getCell(1, 2).value).toBe("C2");

    sheet.setCellLiteral(1, 1, "B2");
    sheet.setCellLiteral(2, 1, "B3");
    sheet.setCellLiteral(3, 1, "B4");
    selection.selectCell(1, 1);
    cm.execute(
      new DeleteCellsShiftUpCommand(sheet, selection, {
        startRow: 1,
        startCol: 1,
        endRow: 1,
        endCol: 1,
      }),
    );
    expect(sheet.getCell(1, 1).value).toBe("B3");
    expect(sheet.getCell(2, 1).value).toBe("B4");
    expect(sheet.getCell(3, 1).value).toBeNull();
    cm.undo();
    expect(sheet.getCell(1, 1).value).toBe("B2");
    expect(sheet.getCell(2, 1).value).toBe("B3");
  });
});

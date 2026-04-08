import { describe, expect, it } from "vitest";
import { Worksheet } from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";

import {
  CutClearRegionCommand,
  CutRangeExceptRectCommand,
  PasteRegionCommand,
} from "flexsheet";
import { parseTsv, serializeTsv } from "../packages/flexsheet/src/clipboard/tsv-io.js";

describe("clipboard TSV", () => {
  it("roundtrips simple grid", () => {
    const rows = [
      ["a", "b"],
      ["1", "2"],
    ];
    expect(parseTsv(serializeTsv(rows))).toEqual(rows);
  });

  it("escapes tab and newline in fields", () => {
    const rows = [["x\ty", "line\n2"]];
    const s = serializeTsv(rows);
    expect(s).toContain('"');
    expect(parseTsv(s)).toEqual(rows);
  });

  it("parses quoted double quotes", () => {
    const rows = [[`say "hi"`]];
    expect(parseTsv(serializeTsv(rows))).toEqual(rows);
  });
});

describe("PasteRegionCommand bounds", () => {
  it("clips paste to sheet size", () => {
    const sheet = new Worksheet("s", 2, 2);
    const cmd = new PasteRegionCommand(sheet, 0, 0, [["a", "b", "c"], ["d", "e", "f"]], null);
    cmd.execute();
    expect(sheet.getCell(0, 0).value).toBe("a");
    expect(sheet.getCell(0, 1).value).toBe("b");
    expect(sheet.getCell(1, 0).value).toBe("d");
    expect(sheet.getCell(1, 1).value).toBe("e");
    cmd.undo();
    expect(sheet.getCell(0, 0).value).toBeNull();
  });

  it("pastes from active corner with partial overlap", () => {
    const sheet = new Worksheet("s", 3, 3);
    const cmd = new PasteRegionCommand(sheet, 1, 1, [["z"]], null);
    cmd.execute();
    expect(sheet.getCell(1, 1).value).toBe("z");
  });
});

describe("PasteRegionCommand + internal styles", () => {
  it("applies styles when matrix matches", () => {
    const sheet = new Worksheet("s", 2, 2);
    const styles = [[{ bold: true, fgArgb: "FFFF0000" } as const]];
    const cmd = new PasteRegionCommand(sheet, 0, 0, [["x"]], styles);
    cmd.execute();
    expect(sheet.getCell(0, 0).style).toEqual({ bold: true, fgArgb: "FFFF0000" });
  });
});

describe("CutRangeExceptRectCommand", () => {
  it("clears cut cells outside paste rectangle only", () => {
    const sheet = new Worksheet("s", 3, 3);
    sheet.setCellValue(0, 0, "a");
    sheet.setCellValue(0, 1, "b");
    sheet.setCellValue(1, 0, "c");
    sheet.setCellValue(1, 1, "d");
    recalcWorksheet(sheet);
    const cmd = new CutRangeExceptRectCommand(
      sheet,
      { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      1,
      1,
      1,
      1,
    );
    cmd.execute();
    expect(sheet.getCell(0, 0).value).toBeNull();
    expect(sheet.getCell(0, 1).value).toBeNull();
    expect(sheet.getCell(1, 0).value).toBeNull();
    expect(sheet.getCell(1, 1).value).toBe("d");
    cmd.undo();
    expect(sheet.getCell(0, 0).value).toBe("a");
    expect(sheet.getCell(1, 1).value).toBe("d");
  });
});

describe("CutClearRegionCommand", () => {
  it("clears and restores", () => {
    const sheet = new Worksheet("s", 2, 2);
    sheet.setCellValue(0, 0, "v");
    sheet.setCellStyle(0, 0, { bold: true });
    recalcWorksheet(sheet);
    const cmd = new CutClearRegionCommand(sheet, {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    });
    cmd.execute();
    expect(sheet.getCell(0, 0).value).toBeNull();
    expect(sheet.getCell(0, 0).style).toBeNull();
    cmd.undo();
    expect(sheet.getCell(0, 0).value).toBe("v");
    expect(sheet.getCell(0, 0).style).toEqual({ bold: true });
  });
});

import { Workbook, Worksheet } from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";
import {
  DEFAULT_FLEXSHEET_JSON_EXPORT_OPTIONS,
  DEFAULT_FLEXSHEET_JSON_IMPORT_OPTIONS,
  FLEXSHEET_JSON_FORMAT,
  FLEXSHEET_JSON_GENERATOR_APP,
  parseFlexSheetJson,
  serializeWorkbookToJsonDocument,
  workbookFromFlexSheetJsonDocument,
} from "@flexsheet/import-export";
import { describe, expect, it } from "vitest";

describe("flexsheet JSON", () => {
  it("parseFlexSheetJson rejects non-object JSON", () => {
    const r = parseFlexSheetJson("[]");
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.length).toBeGreaterThan(0);
    }
  });

  it("parseFlexSheetJson rejects missing format", () => {
    const r = parseFlexSheetJson(JSON.stringify({ workbook: { sheets: [] } }));
    expect(r.ok).toBe(false);
  });

  it("parseFlexSheetJson rejects wrong generator.app", () => {
    const r = parseFlexSheetJson(
      JSON.stringify({
        format: FLEXSHEET_JSON_FORMAT,
        formatVersion: 1,
        generator: { app: "OtherApp" },
        workbook: {
          activeSheetIndex: 0,
          sheets: [
            {
              name: "S1",
              rowCount: 2,
              colCount: 2,
              cells: [],
            },
          ],
        },
      }),
    );
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error).toContain("generator");
    }
  });

  it("roundtrips workbook with value, style, and formula", () => {
    const wb = new Workbook();
    const sh = new Worksheet("表1", 10, 8);
    sh.setCellLiteral(1, 2, 42);
    sh.setCellStyle(1, 2, { bold: true, fgArgb: "FFFF0000" });
    sh.setCellFormula(2, 2, "=1+2");
    recalcWorksheet(sh);
    wb.addSheet(sh);
    wb.activeSheetIndex = 0;

    const doc = serializeWorkbookToJsonDocument(wb, DEFAULT_FLEXSHEET_JSON_EXPORT_OPTIONS, {
      generatorVersion: "test",
    });
    expect(doc.format).toBe(FLEXSHEET_JSON_FORMAT);
    expect(doc.generator.app).toBe(FLEXSHEET_JSON_GENERATOR_APP);

    const text = JSON.stringify(doc);
    const parsed = parseFlexSheetJson(text);
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) {
      return;
    }
    const wb2 = workbookFromFlexSheetJsonDocument(parsed.doc);
    expect(wb2.sheetCount).toBe(1);
    const s2 = wb2.getSheet(0);
    expect(s2).toBeDefined();
    if (s2 === undefined) {
      return;
    }
    expect(s2.getCell(1, 2).value).toBe(42);
    expect(s2.getCell(1, 2).style?.bold).toBe(true);
    expect(s2.getCell(2, 2).formula).toBe("=1+2");
    expect(s2.getCell(2, 2).value).toBe(3);
  });

  it("import skips styles and formulas when options say so", () => {
    const doc = {
      format: FLEXSHEET_JSON_FORMAT,
      formatVersion: 1 as const,
      generator: { app: FLEXSHEET_JSON_GENERATOR_APP },
      workbook: {
        activeSheetIndex: 0,
        sheets: [
          {
            name: "S",
            rowCount: 5,
            colCount: 5,
            cells: [
              { r: 0, c: 0, v: 1, s: { bold: true } },
              { r: 1, c: 0, f: "=A1+1", v: 2 },
            ],
          },
        ],
      },
    };
    const parsed = parseFlexSheetJson(JSON.stringify(doc));
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) {
      return;
    }
    const wb = workbookFromFlexSheetJsonDocument(parsed.doc, {
      ...DEFAULT_FLEXSHEET_JSON_IMPORT_OPTIONS,
      includeStyles: false,
      includeFormulas: false,
      recalcAfterImport: false,
    });
    const sh = wb.getSheet(0);
    expect(sh).toBeDefined();
    if (sh === undefined) {
      return;
    }
    expect(sh.getCell(0, 0).style).toBeNull();
    expect(sh.getCell(0, 0).value).toBe(1);
    expect(sh.getCell(1, 0).formula).toBeNull();
    expect(sh.getCell(1, 0).value).toBe(2);
  });

  it("export without formulas keeps displayed values only", () => {
    const wb = new Workbook();
    const sh = new Worksheet("S", 5, 5);
    sh.setCellFormula(0, 0, "=2*3");
    recalcWorksheet(sh);
    wb.addSheet(sh);

    const doc = serializeWorkbookToJsonDocument(wb, {
      ...DEFAULT_FLEXSHEET_JSON_EXPORT_OPTIONS,
      includeFormulas: false,
    });
    const cell0 = doc.workbook.sheets[0]?.cells.find((c) => c.r === 0 && c.c === 0);
    expect(cell0?.f).toBeUndefined();
    expect(cell0?.v).toBe(6);
  });
});

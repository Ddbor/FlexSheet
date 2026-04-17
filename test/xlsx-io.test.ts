import { describe, expect, it } from "vitest";
import { type ConditionalFormatRule, Workbook, Worksheet } from "@flexsheet/core";
import {
  crc32,
  exportWorkbookToXlsxBytes,
  importXlsxToWorkbook,
  unzipToMap,
} from "@flexsheet/import-export";

describe("crc32", () => {
  it("matches IEEE / PKZIP test vector 123456789", () => {
    const enc = new TextEncoder();
    expect(crc32(enc.encode("123456789"))).toBe(0xcbf43926);
  });

  it("empty payload is 0", () => {
    expect(crc32(new Uint8Array(0))).toBe(0);
  });
});

describe("xlsx export/import", () => {
  it("roundtrips per-row height and per-column width", async () => {
    const wb = new Workbook();
    const s = new Worksheet("S", 12, 12);
    wb.addSheet(s);
    s.getCell(0, 0).value = "a";
    s.setRowHeight(0, 40);
    // 117 = 16*7+5，与 OOXML 字符宽度往返无舍入误差
    s.setColWidth(0, 117);

    const bytes = exportWorkbookToXlsxBytes(wb);
    const back = await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    const sh = back.getSheet(0);
    expect(sh?.getRowHeight(0)).toBe(40);
    expect(sh?.getColWidth(0)).toBe(117);
  });

  it("roundtrips multi-sheet, values, formula, style", async () => {
    const wb = new Workbook();
    const a = new Worksheet("第一页");
    wb.addSheet(a);
    a.getCell(0, 0).value = "标题";
    a.getCell(1, 0).value = 42;
    a.getCell(1, 1).formula = "=A2*2";
    a.getCell(2, 0).style = { bold: true, fgArgb: "FFFF0000" };

    const b = new Worksheet("Sheet2");
    wb.addSheet(b);
    b.getCell(0, 0).value = "b1";
    b.getCell(0, 1).formula = "='第一页'!A2+1";

    const bytes = exportWorkbookToXlsxBytes(wb);
    const back = await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );

    expect(back.sheetCount).toBe(2);
    const s0 = back.getSheet(0);
    expect(s0?.name).toBe("第一页");
    expect(s0?.getCell(0, 0).value).toBe("标题");
    expect(s0?.getCell(1, 0).value).toBe(42);
    expect(s0?.getCell(1, 1).formula).toBe("=A2*2");
    expect(s0?.getCell(2, 0).style?.bold).toBe(true);
    expect(s0?.getCell(2, 0).style?.fgArgb?.toUpperCase()).toBe("FFFF0000");

    const s1 = back.getSheet(1);
    expect(s1?.getCell(0, 0).value).toBe("b1");
    expect(s1?.getCell(0, 1).formula).toContain("第一页");
  });

  it("sharedStrings count is total refs, uniqueCount is distinct strings (Excel SST)", () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "dup";
    s.getCell(0, 1).value = "dup";
    s.getCell(0, 2).value = "dup";
    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const sst = new TextDecoder().decode(map.get("xl/sharedStrings.xml"));
    expect(sst).toContain('count="3"');
    expect(sst).toContain('uniqueCount="1"');
  });

  it("ZIP is valid OPC with required parts", () => {
    const wb = new Workbook();
    wb.addSheet(new Worksheet("S"));
    wb.getSheet(0)!.getCell(0, 0).value = "x";
    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    expect(map.has("[Content_Types].xml")).toBe(true);
    expect(map.has("_rels/.rels")).toBe(true);
    expect(map.has("xl/workbook.xml")).toBe(true);
    expect(map.has("xl/_rels/workbook.xml.rels")).toBe(true);
    expect(map.has("xl/sharedStrings.xml")).toBe(true);
    expect(map.has("xl/styles.xml")).toBe(true);
    expect(map.has("xl/worksheets/sheet1.xml")).toBe(true);
  });

  it("roundtrips strikethrough and superscript/subscript font", async () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "strike";
    s.getCell(0, 0).style = { strikethrough: true };
    s.getCell(0, 1).value = "sup";
    s.getCell(0, 1).style = { fontScript: "superscript" };
    s.getCell(0, 2).value = "sub";
    s.getCell(0, 2).style = { fontScript: "subscript" };

    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const stylesXml = new TextDecoder().decode(map.get("xl/styles.xml"));
    expect(stylesXml).toContain("<strike/>");
    expect(stylesXml).toContain('vertAlign val="superscript"');
    expect(stylesXml).toContain('vertAlign val="subscript"');

    const back = await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    const sh = back.getSheet(0);
    expect(sh?.getCell(0, 0).style?.strikethrough).toBe(true);
    expect(sh?.getCell(0, 1).style?.fontScript).toBe("superscript");
    expect(sh?.getCell(0, 2).style?.fontScript).toBe("subscript");
  });

  it("roundtrips font (family, size, italic, underline) and borders", async () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "a";
    s.getCell(0, 0).style = {
      fontFamily: "Arial",
      fontSizePt: 12,
      italic: true,
      underline: "single",
      fgArgb: "FF0066CC",
      borderTop: { kind: "thin", colorArgb: "FF000000" },
      borderBottom: { kind: "medium", colorArgb: "FFFF0000" },
    };

    const bytes = exportWorkbookToXlsxBytes(wb);
    const back = await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    const sh = back.getSheet(0);
    const st = sh?.getCell(0, 0).style;
    expect(st?.fontFamily).toBe("Arial");
    expect(st?.fontSizePt).toBe(12);
    expect(st?.italic).toBe(true);
    expect(st?.underline).toBe("single");
    expect(st?.fgArgb?.toUpperCase()).toBe("FF0066CC");
    expect(st?.borderTop?.kind).toBe("thin");
    expect(st?.borderBottom?.kind).toBe("medium");
    expect(st?.borderBottom?.colorArgb?.toUpperCase()).toBe("FFFF0000");
  });

  it("roundtrips merge cells", async () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "merged";
    s.applyMergeForSelection({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }, "mergeCells");

    const bytes = exportWorkbookToXlsxBytes(wb);
    const back = await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    const sh = back.getSheet(0);
    expect(sh?.getCell(0, 0).value).toBe("merged");
    expect(sh?.isMergeCoveredCell(1, 0)).toBe(true);
    expect(sh?.isMergeCoveredCell(0, 1)).toBe(true);
    const info = sh?.getMergedRectInfo(0, 0);
    expect(info?.rowSpan).toBe(2);
    expect(info?.colSpan).toBe(2);
  });

  it("exports default vertical align as center to match FlexSheet (unset vAlign renders middle)", () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "x";
    s.getCell(0, 0).style = { bold: true };

    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const stylesXml = new TextDecoder().decode(map.get("xl/styles.xml"));
    expect(stylesXml).toContain('vertical="center"');
  });

  it("exports indent with horizontal=left so Excel applies indent (OOXML requires horizontal)", () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "indented";
    s.getCell(0, 0).style = { indentLevel: 2 };

    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const stylesXml = new TextDecoder().decode(map.get("xl/styles.xml"));
    expect(stylesXml).toContain('horizontal="left"');
    expect(stylesXml).toContain('indent="2"');
  });

  it("roundtrips text orientation (textRotation)", async () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "a";
    s.getCell(0, 0).style = { textOrientation: "angleUp45" };
    s.getCell(0, 1).value = "b";
    s.getCell(0, 1).style = { textOrientation: "verticalStack" };
    s.getCell(0, 2).value = "c";
    s.getCell(0, 2).style = { textOrientation: "rotateDown90" };

    const bytes = exportWorkbookToXlsxBytes(wb);
    const back = await importXlsxToWorkbook(
      new Blob([new Uint8Array(bytes)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    const sh = back.getSheet(0);
    expect(sh?.getCell(0, 0).style?.textOrientation).toBe("angleUp45");
    expect(sh?.getCell(0, 1).style?.textOrientation).toBe("verticalStack");
    expect(sh?.getCell(0, 2).style?.textOrientation).toBe("rotateDown90");

    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const stylesXml = new TextDecoder().decode(map.get("xl/styles.xml"));
    expect(stylesXml).toContain('textRotation="45"');
    expect(stylesXml).toContain('textRotation="255"');
    expect(stylesXml).toContain('textRotation="180"');
  });

  it("exports explicit vertical top/bottom to OOXML vertical top/bottom", () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "a";
    s.getCell(0, 0).style = { vAlign: "top" };
    s.getCell(0, 1).value = "b";
    s.getCell(0, 1).style = { vAlign: "bottom" };

    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const stylesXml = new TextDecoder().decode(map.get("xl/styles.xml"));
    expect(stylesXml).toContain('vertical="top"');
    expect(stylesXml).toContain('vertical="bottom"');
  });

  it("writes perimeter cells so merged outer borders span the full merge in Excel", () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = "x";
    s.getCell(0, 0).style = {
      borderTop: { kind: "thin", colorArgb: "FF000000" },
      borderLeft: { kind: "thin", colorArgb: "FF000000" },
      borderBottom: { kind: "thin", colorArgb: "FF000000" },
      borderRight: { kind: "thin", colorArgb: "FF000000" },
    };
    s.applyMergeForSelection({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }, "mergeCells");

    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const sheetXml = new TextDecoder().decode(map.get("xl/worksheets/sheet1.xml"));
    expect(sheetXml).toMatch(/<c r="A1"[^>]*\ss="/);
    expect(sheetXml).toMatch(/<c r="B1"[^>]*\ss="/);
    expect(sheetXml).toMatch(/<c r="A2"[^>]*\ss="/);
    expect(sheetXml).toMatch(/<c r="B2"[^>]*\ss="/);
  });

  it("exports conditional formatting (color scale) for Excel worksheet OOXML", () => {
    const wb = new Workbook();
    const s = new Worksheet("S", 6, 6);
    wb.addSheet(s);
    s.getCell(0, 0).value = 1;
    s.getCell(0, 1).value = 5;
    const rule: ConditionalFormatRule = {
      id: "cf-test-1",
      range: { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      uiFamily: "twoColorScale",
      classicType: "colorScale",
      formatPreset: "none",
      cfTwoColorMin: { type: "lowest", value: "", colorArgb: "FFFFFFFF" },
      cfTwoColorMax: { type: "highest", value: "", colorArgb: "FF000000" },
    };
    s.addConditionalFormatRule(rule);

    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const sheetXml = new TextDecoder().decode(map.get("xl/worksheets/sheet1.xml"));
    expect(sheetXml).toContain("<conditionalFormatting");
    expect(sheetXml).toContain("colorScale");
    expect(sheetXml).toContain('sqref="A1:B1"');

    const stylesXml = new TextDecoder().decode(map.get("xl/styles.xml"));
    expect(stylesXml).toContain("<dxfs ");
  });

  it("exports classic conditional formatting with dxfs for Excel", () => {
    const wb = new Workbook();
    const s = new Worksheet("S");
    wb.addSheet(s);
    s.getCell(0, 0).value = 10;
    const rule: ConditionalFormatRule = {
      id: "cf-classic-1",
      range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      uiFamily: "classic",
      classicType: "cellsThatContain",
      cellsThatContainKind: "cellValue",
      valueOperator: "greaterThan",
      value1: "5",
      formatPreset: "lightRedFillDarkRedText",
    };
    s.addConditionalFormatRule(rule);

    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const sheetXml = new TextDecoder().decode(map.get("xl/worksheets/sheet1.xml"));
    expect(sheetXml).toContain('type="cellIs"');
    expect(sheetXml).toContain('operator="greaterThan"');
    expect(sheetXml).toContain("dxfId=");

    const stylesXml = new TextDecoder().decode(map.get("xl/styles.xml"));
    expect(stylesXml).toMatch(/<dxfs count="[1-9]/);
    expect(stylesXml).toContain("<dxf>");
  });

  it("exports pivot table with pivotCacheRecords and OOXML-valid dataField attributes", () => {
    const wb = new Workbook();
    const src = new Worksheet("Data", 20, 10);
    wb.addSheet(src);
    src.setCellLiteral(0, 0, "A");
    src.setCellLiteral(0, 1, "B");
    src.setCellLiteral(1, 0, "x");
    src.setCellLiteral(1, 1, 1);
    const dest = new Worksheet("Pivot", 30, 10);
    wb.addSheet(dest);
    dest.registerPivotTableDefinition({
      id: "p1",
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
      outputRowCount: 5,
      outputColCount: 3,
    });

    const bytes = exportWorkbookToXlsxBytes(wb);
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    expect(map.has("xl/pivotCache/pivotCacheRecords1.xml")).toBe(true);
    expect(map.has("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels")).toBe(true);
    const records = new TextDecoder().decode(map.get("xl/pivotCache/pivotCacheRecords1.xml"));
    expect(records).toContain('count="0"');
    const pivotTable = new TextDecoder().decode(map.get("xl/pivotTables/pivotTable1.xml"));
    expect(pivotTable).toContain("<dataField ");
    expect(pivotTable).not.toContain('baseField="0"');
    expect(pivotTable).not.toContain('baseItem="0"');
  });
});

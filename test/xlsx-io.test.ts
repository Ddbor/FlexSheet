import { describe, expect, it } from "vitest";
import { Workbook, Worksheet } from "@flexsheet/core";
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
});

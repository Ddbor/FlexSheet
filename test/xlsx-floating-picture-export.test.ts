import { describe, expect, it } from "vitest";
import { Workbook, Worksheet } from "@flexsheet/core";
import {
  exportWorkbookToXlsxBytes,
  unzipToMap,
  type XlsxFloatingPictureExport,
} from "@flexsheet/import-export";

/** 1×1 透明 PNG */
const TINY_PNG_DATA_URL =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

describe("xlsx export floating pictures", () => {
  it("writes drawings, media, worksheet rel and drawing element", () => {
    const wb = new Workbook();
    const sh = new Worksheet("S1", 12, 12);
    wb.addSheet(sh);
    const pic: XlsxFloatingPictureExport = {
      sheetName: sh.name,
      anchorRow: 0,
      anchorCol: 0,
      relCX: 0,
      relCY: 0,
      width: 64,
      height: 64,
      rotationRad: 0,
      dataUrl: TINY_PNG_DATA_URL,
    };
    const bytes = exportWorkbookToXlsxBytes(wb, {
      includeStyles: true,
      includeFormulas: true,
      includeSparseStyledEmpty: true,
      viewZoom: 1,
      floatingPictures: [pic],
    });
    const map = unzipToMap(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    );
    const keys = [...map.keys()].sort();
    expect(keys.some((k) => k === "xl/drawings/drawing1.xml")).toBe(true);
    expect(keys.some((k) => k === "xl/drawings/_rels/drawing1.xml.rels")).toBe(true);
    expect(keys.some((k) => k === "xl/media/image1.png")).toBe(true);
    const sheet1 = new TextDecoder().decode(map.get("xl/worksheets/sheet1.xml"));
    expect(sheet1).toContain("<drawing ");
    const rels = new TextDecoder().decode(map.get("xl/worksheets/_rels/sheet1.xml.rels"));
    expect(rels).toContain("relationships/drawing");
    expect(rels).toContain("drawings/drawing1.xml");
  });
});

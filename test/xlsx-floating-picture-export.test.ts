import { describe, expect, it } from "vitest";
import { Workbook, Worksheet } from "@flexsheet/core";
import {
  collectSheetFloatingPicturesFromXlsx,
  exportWorkbookToXlsxBytes,
  floatingPictureNeedsFrameCompositeForXlsx,
  floatingPictureNeedsRasterForXlsxExport,
  floatingPictureSrcRectSides,
  importXlsx,
  unzipToMap,
  type XlsxFloatingPictureExport,
} from "@flexsheet/import-export";

/** 1×1 透明 PNG */
const TINY_PNG_DATA_URL =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

describe("xlsx export floating pictures", () => {
  it("detects when frame has empty margin vs image (needs PNG composite)", () => {
    const base: Pick<
      XlsxFloatingPictureExport,
      "sheetName" | "anchorRow" | "anchorCol" | "relCX" | "relCY" | "rotationRad" | "dataUrl"
    > = {
      sheetName: "S",
      anchorRow: 0,
      anchorCol: 0,
      relCX: 0,
      relCY: 0,
      rotationRad: 0,
      dataUrl: TINY_PNG_DATA_URL,
    };
    expect(
      floatingPictureNeedsFrameCompositeForXlsx({
        ...base,
        width: 100,
        height: 100,
        imgBoxX: 0,
        imgBoxY: 0,
        imgBoxW: 100,
        imgBoxH: 100,
      }),
    ).toBe(false);
    expect(
      floatingPictureNeedsFrameCompositeForXlsx({
        ...base,
        width: 100,
        height: 100,
        imgBoxX: 0,
        imgBoxY: 0,
        imgBoxW: 50,
        imgBoxH: 50,
      }),
    ).toBe(true);
    expect(
      floatingPictureNeedsFrameCompositeForXlsx({
        ...base,
        width: 100,
        height: 100,
        imgBoxX: 25,
        imgBoxY: 25,
        imgBoxW: 50,
        imgBoxH: 50,
      }),
    ).toBe(true);
  });

  it("raster export when solid frame fill even if image fills geometry", () => {
    const base: Pick<
      XlsxFloatingPictureExport,
      "sheetName" | "anchorRow" | "anchorCol" | "relCX" | "relCY" | "rotationRad" | "dataUrl"
    > = {
      sheetName: "S",
      anchorRow: 0,
      anchorCol: 0,
      relCX: 0,
      relCY: 0,
      rotationRad: 0,
      dataUrl: TINY_PNG_DATA_URL,
    };
    const full: XlsxFloatingPictureExport = {
      ...base,
      width: 100,
      height: 100,
      imgBoxX: 0,
      imgBoxY: 0,
      imgBoxW: 100,
      imgBoxH: 100,
      frameFill: { kind: "solid", solidColor: "#fde9d9", solidTransparencyPct: 0 },
    };
    expect(floatingPictureNeedsFrameCompositeForXlsx(full)).toBe(false);
    expect(floatingPictureNeedsRasterForXlsxExport(full)).toBe(true);
  });

  it("writes a:solidFill in drawing when frameFill is solid (sync export)", () => {
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
      frameFill: { kind: "solid", solidColor: "#FDE9D9", solidTransparencyPct: 0 },
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
    const drawing = new TextDecoder().decode(map.get("xl/drawings/drawing1.xml"));
    expect(drawing).toContain("<a:solidFill>");
    expect(drawing).toContain("<a:srgbClr");
    expect(drawing).toContain('val="FDE9D9"');
  });

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

  it("writes srcRect when crop geometry differs from full bitmap", () => {
    const wb = new Workbook();
    const sh = new Worksheet("S1", 12, 12);
    wb.addSheet(sh);
    const pic: XlsxFloatingPictureExport = {
      sheetName: sh.name,
      anchorRow: 0,
      anchorCol: 0,
      relCX: 0,
      relCY: 0,
      width: 50,
      height: 50,
      rotationRad: 0,
      dataUrl: TINY_PNG_DATA_URL,
      naturalWidth: 100,
      naturalHeight: 100,
      imgBoxX: 0,
      imgBoxY: 0,
      imgBoxW: 100,
      imgBoxH: 100,
    };
    const sides = floatingPictureSrcRectSides(pic);
    expect(sides).toBeDefined();
    expect(sides!.l).toBe("0%");
    expect(sides!.t).toBe("0%");
    expect(sides!.r).toBe("50%");
    expect(sides!.b).toBe("50%");

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
    const drawing = new TextDecoder().decode(map.get("xl/drawings/drawing1.xml"));
    expect(drawing).toContain("<a:srcRect ");
    expect(drawing).toContain('l="0%"');
    expect(drawing).toContain('r="50%"');
  });

  it("collectSheetFloatingPicturesFromXlsx reads picture and solid fill from exported zip", () => {
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
      frameFill: { kind: "solid", solidColor: "#FDE9D9", solidTransparencyPct: 0 },
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
    const found = collectSheetFloatingPicturesFromXlsx(map, "xl/worksheets/sheet1.xml", sh, "S1", 0);
    expect(found.length).toBe(1);
    expect(found[0]?.frameFill?.solidColor.toUpperCase()).toBe("#FDE9D9");
  });

  it("importXlsx restores floating picture and solid frame fill from export", async () => {
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
      frameFill: { kind: "solid", solidColor: "#FDE9D9", solidTransparencyPct: 0 },
    };
    const bytes = exportWorkbookToXlsxBytes(wb, {
      includeStyles: true,
      includeFormulas: true,
      includeSparseStyledEmpty: true,
      viewZoom: 1,
      floatingPictures: [pic],
    });
    const result = await importXlsx(
      new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    expect(result.floatingPictures.length).toBe(1);
    const p = result.floatingPictures[0]!;
    expect(p.dataUrl.startsWith("data:image/png")).toBe(true);
    expect(p.sheetName).toBe("S1");
    expect(p.frameFill?.kind).toBe("solid");
    expect(p.frameFill?.solidColor.toUpperCase()).toBe("#FDE9D9");
  });
});

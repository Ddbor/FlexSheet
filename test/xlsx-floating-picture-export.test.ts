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
import { REL_WORKSHEET_DRAWING } from "../packages/import-export/src/export-xlsx-drawing.js";
import { buildZipArchive } from "../packages/import-export/src/zip-writer.js";

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

  it("solid frame fill does not force raster merge (OOXML layered spPr + blip)", () => {
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
    expect(floatingPictureNeedsRasterForXlsxExport(full)).toBe(false);
  });

  it("writes full spPr xfrm and integer blip srcRect for imgBox (Excel crop mode; not stretch/fillRect)", () => {
    const EMU = 9525;
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
      imgBoxX: 10,
      imgBoxY: 10,
      imgBoxW: 32,
      imgBoxH: 32,
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
    expect(drawing).toContain(`<a:off x="0" y="0"/>`);
    expect(drawing).toContain(`<a:ext cx="${64 * EMU}" cy="${64 * EMU}"/>`);
    expect(drawing).toContain("<a:srcRect ");
    expect(drawing).toContain('l="-31250"');
    expect(drawing).toContain('r="-68750"');
    expect(drawing).not.toContain("fillRect");
    expect(drawing).toContain("<a:solidFill>");
    expect(drawing).toContain("<a:stretch/>");

    const roundtrip = collectSheetFloatingPicturesFromXlsx(
      map,
      "xl/worksheets/sheet1.xml",
      sh,
      "S1",
      0,
    );
    expect(roundtrip.length).toBe(1);
    expect(roundtrip[0]?.imgBoxX).toBeCloseTo(10, 5);
    expect(roundtrip[0]?.imgBoxY).toBeCloseTo(10, 5);
    expect(roundtrip[0]?.imgBoxW).toBeCloseTo(32, 5);
    expect(roundtrip[0]?.imgBoxH).toBeCloseTo(32, 5);
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

  it("writes a:gradFill (linear) and round-trips stops and angle", () => {
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
      frameFill: {
        kind: "gradient",
        solidColor: "#000000",
        solidTransparencyPct: 0,
        gradientType: "linear",
        gradientAngleDeg: 90,
        gradientRotateWithShape: true,
        gradientPresetId: null,
        gradientStops: [
          { positionPct: 0, color: "#ff0000", transparencyPct: 0, brightnessPct: 0 },
          { positionPct: 100, color: "#0000ff", transparencyPct: 0, brightnessPct: 0 },
        ],
      },
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
    const drawing = new TextDecoder().decode(map.get("xl/drawings/drawing1.xml")!);
    expect(drawing).toContain("<a:gradFill");
    expect(drawing).toContain("<a:gsLst>");
    expect(drawing).toContain('<a:lin ang="5400000" scaled="1"/>');
    expect(drawing).toContain('val="FF0000"');
    expect(drawing).toContain('val="0000FF"');

    const found = collectSheetFloatingPicturesFromXlsx(
      map,
      "xl/worksheets/sheet1.xml",
      sh,
      "S1",
      0,
    );
    expect(found.length).toBe(1);
    const ff = found[0]?.frameFill;
    expect(ff?.kind).toBe("gradient");
    if (ff?.kind !== "gradient") {
      return;
    }
    expect(ff.gradientAngleDeg).toBe(90);
    expect(ff.gradientStops?.length).toBe(2);
    expect(ff.gradientStops?.[0]?.color.toLowerCase()).toBe("#ff0000");
    expect(ff.gradientStops?.[1]?.color.toLowerCase()).toBe("#0000ff");
  });

  it("writes a:gradFill (radial a:path circle) and round-trips fillToRect / tileRect", () => {
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
      frameFill: {
        kind: "gradient",
        solidColor: "#000000",
        solidTransparencyPct: 0,
        gradientType: "radial",
        gradientAngleDeg: 90,
        gradientRotateWithShape: true,
        gradientPresetId: null,
        radialFillLtrb: { l: 0, t: 0, r: 100000, b: 100000 },
        radialTileLtrb: { l: -100000, t: -100000, r: 0, b: 0 },
        gradientStops: [
          { positionPct: 0, color: "#00ff00", transparencyPct: 0, brightnessPct: 0 },
          { positionPct: 100, color: "#000000", transparencyPct: 0, brightnessPct: 0 },
        ],
      },
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
    const drawing = new TextDecoder().decode(map.get("xl/drawings/drawing1.xml")!);
    expect(drawing).toContain("<a:gradFill");
    expect(drawing).toContain('<a:path path="circle">');
    expect(drawing).toContain('<a:fillToRect l="0" t="0" r="100000" b="100000"/>');
    expect(drawing).toContain('<a:tileRect l="-100000" t="-100000" r="0" b="0"/>');
    expect(drawing).toContain('val="00FF00"');
    expect(drawing).not.toContain("<a:lin ");

    const found = collectSheetFloatingPicturesFromXlsx(
      map,
      "xl/worksheets/sheet1.xml",
      sh,
      "S1",
      0,
    );
    expect(found.length).toBe(1);
    const ff = found[0]?.frameFill;
    expect(ff?.kind).toBe("gradient");
    if (ff?.kind !== "gradient") {
      return;
    }
    expect(ff.gradientType).toBe("radial");
    expect(ff.radialFillLtrb).toEqual({ l: 0, t: 0, r: 100000, b: 100000 });
    expect(ff.radialTileLtrb).toEqual({ l: -100000, t: -100000, r: 0, b: 0 });
    expect(ff.gradientStops?.[0]?.color.toLowerCase()).toBe("#00ff00");
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
    const found = collectSheetFloatingPicturesFromXlsx(
      map,
      "xl/worksheets/sheet1.xml",
      sh,
      "S1",
      0,
    );
    expect(found.length).toBe(1);
    expect(found[0]?.frameFill?.solidColor.toUpperCase()).toBe("#FDE9D9");
  });

  it("twoCellAnchor uses pic spPr xfrm absolute off/ext as frame, not from-to span", () => {
    const EMU = 9525;
    const wb = new Workbook();
    const sh = new Worksheet("S1", 60, 60);
    wb.addSheet(sh);
    const pngB64 = TINY_PNG_DATA_URL.split(",")[1] ?? "";
    const pngBytes = Uint8Array.from(atob(pngB64), (ch) => ch.charCodeAt(0));
    const offX = 2 * sh.defaultColWidth * EMU;
    const offY = 3 * sh.defaultRowHeight * EMU;
    const cx = 100 * EMU;
    const cy = 80 * EMU;
    const drawingXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<xdr:twoCellAnchor>` +
      `<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>25</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>50</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
      `<xdr:pic>` +
      `<xdr:nvPicPr><xdr:cNvPr id="1" name="P"/>` +
      `<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
      `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>` +
      `<a:stretch/></xdr:blipFill>` +
      `<xdr:spPr><a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>` +
      `</xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
    const relsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image0.png"/>` +
      `</Relationships>`;
    const sheetRels =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_WORKSHEET_DRAWING}" Target="../drawings/drawing1.xml"/>` +
      `</Relationships>`;
    const files = new Map<string, Uint8Array>([
      ["xl/drawings/drawing1.xml", new TextEncoder().encode(drawingXml)],
      ["xl/drawings/_rels/drawing1.xml.rels", new TextEncoder().encode(relsXml)],
      ["xl/worksheets/_rels/sheet1.xml.rels", new TextEncoder().encode(sheetRels)],
      ["xl/media/image0.png", pngBytes],
    ]);
    const pics = collectSheetFloatingPicturesFromXlsx(
      files,
      "xl/worksheets/sheet1.xml",
      sh,
      "S1",
      0,
    );
    expect(pics.length).toBe(1);
    const p = pics[0]!;
    expect(p.width).toBeCloseTo(100, 2);
    expect(p.height).toBeCloseTo(80, 2);
    expect(p.imgBoxX).toBeCloseTo(0, 3);
    expect(p.imgBoxY).toBeCloseTo(0, 3);
    expect(p.imgBoxW).toBeCloseTo(100, 2);
    expect(p.imgBoxH).toBeCloseTo(80, 2);
  });

  it("import honors spPr xfrm inner rect smaller than anchor (fill margins)", async () => {
    const EMU = 9525;
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
    const drawPath = "xl/drawings/drawing1.xml";
    const drawXml0 = new TextDecoder().decode(map.get(drawPath)!);
    const outer = 64 * EMU;
    const margin = 10 * EMU;
    const inner = 32 * EMU;
    const oldXfrm = `<a:off x="0" y="0"/><a:ext cx="${outer}" cy="${outer}"/>`;
    const newXfrm = `<a:off x="${margin}" y="${margin}"/><a:ext cx="${inner}" cy="${inner}"/>`;
    expect(drawXml0).toContain(oldXfrm);
    const drawXml1 = drawXml0.replace(oldXfrm, newXfrm);
    const next = new Map(map);
    next.set(drawPath, new TextEncoder().encode(drawXml1));
    const entries = [...next.entries()].map(([path, data]) => ({ path, data }));
    const patched = buildZipArchive(entries);
    const result = await importXlsx(
      new Blob([patched], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    expect(result.floatingPictures.length).toBe(1);
    const p = result.floatingPictures[0]!;
    expect(p.width).toBeCloseTo(64, 3);
    expect(p.height).toBeCloseTo(64, 3);
    expect(p.imgBoxX).toBeCloseTo(10, 3);
    expect(p.imgBoxY).toBeCloseTo(10, 3);
    expect(p.imgBoxW).toBeCloseTo(32, 3);
    expect(p.imgBoxH).toBeCloseTo(32, 3);
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

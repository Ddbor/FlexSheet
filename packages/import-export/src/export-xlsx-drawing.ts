import type { Workbook, Worksheet } from "@flexsheet/core";
import { escapeXml } from "./xml-escape.js";

/** 浮动图片导出描述（与 FlexSheet 浮动层模型一致；尺寸与偏移为画布 CSS 像素，需配合 `viewZoom`）。 */
export interface XlsxFloatingPictureExport {
  readonly sheetName: string;
  readonly anchorRow: number;
  readonly anchorCol: number;
  readonly relCX: number;
  readonly relCY: number;
  readonly width: number;
  readonly height: number;
  readonly rotationRad: number;
  readonly dataUrl: string;
}

const EMU_PER_PX = 9525;

const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

export const REL_WORKSHEET_DRAWING =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
export const REL_DRAWING_IMAGE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

function mergedCellSheetRectPx(
  sheet: Worksheet,
  row: number,
  col: number,
): { x: number; y: number; w: number; h: number } {
  const info = sheet.getMergedRectInfo(row, col);
  let x = 0;
  for (let c = 0; c < info.anchorCol; c++) {
    x += sheet.getColWidth(c);
  }
  let y = 0;
  for (let r = 0; r < info.anchorRow; r++) {
    y += sheet.getRowHeight(r);
  }
  let w = 0;
  for (let c = info.anchorCol; c < info.anchorCol + info.colSpan; c++) {
    w += sheet.getColWidth(c);
  }
  let h = 0;
  for (let r = info.anchorRow; r < info.anchorRow + info.rowSpan; r++) {
    h += sheet.getRowHeight(r);
  }
  return { x, y, w, h };
}

function pictureSheetBoxPx(
  sheet: Worksheet,
  pic: XlsxFloatingPictureExport,
  viewZoom: number,
): { left: number; top: number; w: number; h: number } {
  const z = Math.max(1e-6, viewZoom);
  const cr = mergedCellSheetRectPx(sheet, pic.anchorRow, pic.anchorCol);
  const cx = cr.x + cr.w / 2 + pic.relCX / z;
  const cy = cr.y + cr.h / 2 + pic.relCY / z;
  const w = pic.width / z;
  const h = pic.height / z;
  return { left: cx - w / 2, top: cy - h / 2, w, h };
}

function pxToFromAnchor(
  sheet: Worksheet,
  xPx: number,
  yPx: number,
): { col: number; colOff: number; row: number; rowOff: number } {
  const x0 = Math.max(0, xPx);
  const y0 = Math.max(0, yPx);
  let x = x0;
  let col = 0;
  const lastCol = Math.max(0, sheet.colCount - 1);
  while (col < lastCol) {
    const cw = sheet.getColWidth(col);
    if (x < cw) {
      break;
    }
    x -= cw;
    col++;
  }
  const colW = sheet.getColWidth(col);
  const colOffEmu = Math.round(Math.min(x, colW) * EMU_PER_PX);

  let y = y0;
  let row = 0;
  const lastRow = Math.max(0, sheet.rowCount - 1);
  while (row < lastRow) {
    const rh = sheet.getRowHeight(row);
    if (y < rh) {
      break;
    }
    y -= rh;
    row++;
  }
  const rowH = sheet.getRowHeight(row);
  const rowOffEmu = Math.round(Math.min(y, rowH) * EMU_PER_PX);
  return { col, colOff: colOffEmu, row, rowOff: rowOffEmu };
}

export function decodeDataUrlToImagePart(
  dataUrl: string,
): { bytes: Uint8Array; extension: "png" | "jpeg" } | null {
  const m = /^data:([^;,]+);base64,(.+)$/i.exec(dataUrl.trim());
  if (m === null) {
    return null;
  }
  const mime = m[1]!.toLowerCase().split(";")[0]!.trim();
  const b64 = m[2]!;
  let extension: "png" | "jpeg";
  if (mime === "image/png") {
    extension = "png";
  } else if (mime === "image/jpeg" || mime === "image/jpg") {
    extension = "jpeg";
  } else {
    return null;
  }
  const bytes = base64ToUint8Array(b64);
  return { bytes, extension };
}

function base64ToUint8Array(b64: string): Uint8Array {
  if (typeof atob === "function") {
    const binary = atob(b64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      out[i] = binary.charCodeAt(i);
    }
    return out;
  }
  const g = globalThis as unknown as {
    Buffer?: { from(data: string, enc: "base64"): Uint8Array };
  };
  if (g.Buffer !== undefined) {
    return g.Buffer.from(b64, "base64");
  }
  throw new Error("无法解码 base64：缺少 atob/Buffer");
}

function rotationToOoxmlRot60000(rotationRad: number): number {
  const deg = (rotationRad * 180) / Math.PI;
  return Math.round(deg * 60000);
}

function buildOnePictureAnchorXml(
  pic: XlsxFloatingPictureExport,
  sheet: Worksheet,
  viewZoom: number,
  shapeId: number,
  embedRid: string,
): string {
  const box = pictureSheetBoxPx(sheet, pic, viewZoom);
  const from = pxToFromAnchor(sheet, box.left, box.top);
  const cxEmu = Math.max(1, Math.round(box.w * EMU_PER_PX));
  const cyEmu = Math.max(1, Math.round(box.h * EMU_PER_PX));
  const rot = rotationToOoxmlRot60000(pic.rotationRad);
  const rotAttr = rot === 0 ? "" : ` rot="${rot}"`;
  return (
    `<xdr:oneCellAnchor>` +
    `<xdr:from>` +
    `<xdr:col>${from.col}</xdr:col>` +
    `<xdr:colOff>${from.colOff}</xdr:colOff>` +
    `<xdr:row>${from.row}</xdr:row>` +
    `<xdr:rowOff>${from.rowOff}</xdr:rowOff>` +
    `</xdr:from>` +
    `<xdr:ext cx="${cxEmu}" cy="${cyEmu}"/>` +
    `<xdr:pic>` +
    `<xdr:nvPicPr>` +
    `<xdr:cNvPr id="${shapeId}" name="${escapeXml(`Picture ${shapeId}`)}"/>` +
    `<xdr:cNvPicPr><a:picLocks noChangeAspect="0"/></xdr:cNvPicPr>` +
    `</xdr:nvPicPr>` +
    `<xdr:blipFill>` +
    `<a:blip xmlns:r="${REL_NS}" r:embed="${embedRid}"/>` +
    `<a:stretch><a:fillRect/></a:stretch>` +
    `</xdr:blipFill>` +
    `<xdr:spPr>` +
    `<a:xfrm${rotAttr}>` +
    `<a:off x="0" y="0"/>` +
    `<a:ext cx="${cxEmu}" cy="${cyEmu}"/>` +
    `</a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</xdr:spPr>` +
    `</xdr:pic>` +
    `<xdr:clientData/>` +
    `</xdr:oneCellAnchor>`
  );
}

export interface SheetDrawingBuildResult {
  readonly drawingXml: string;
  readonly drawingRelsXml: string;
  readonly mediaEntries: readonly { path: string; data: Uint8Array }[];
  /** 下一可用 `xl/media/image{N}` 序号（本次已占用至 N-1）。 */
  readonly nextMediaFileIndex: number;
}

/**
 * 为单张工作表构建 `drawingN.xml`、其 `.rels` 及待写入的 `xl/media` 条目。
 * `embedRid` 自 `rId1` 递增，与 `drawingRelsXml` 一致。
 */
export function buildSheetDrawingPackage(
  sheet: Worksheet,
  pictures: readonly XlsxFloatingPictureExport[],
  viewZoom: number,
  mediaFileIndexStart: number,
): SheetDrawingBuildResult | null {
  if (pictures.length === 0) {
    return null;
  }
  const anchors: string[] = [];
  const rels: string[] = [];
  const mediaEntries: { path: string; data: Uint8Array }[] = [];
  let mediaIdx = mediaFileIndexStart;
  let shapeId = 1;
  let embedNum = 1;
  for (const pic of pictures) {
    const decoded = decodeDataUrlToImagePart(pic.dataUrl);
    if (decoded === null) {
      continue;
    }
    const ext = decoded.extension;
    const mediaPath = `xl/media/image${mediaIdx}.${ext}`;
    mediaIdx += 1;
    const embedRid = `rId${embedNum}`;
    embedNum += 1;
    const anchorXml = buildOnePictureAnchorXml(pic, sheet, viewZoom, shapeId, embedRid);
    shapeId += 1;
    anchors.push(anchorXml);
    rels.push(
      `<Relationship Id="${embedRid}" Type="${REL_DRAWING_IMAGE}" Target="../media/image${mediaIdx - 1}.${ext}"/>`,
    );
    mediaEntries.push({ path: mediaPath, data: decoded.bytes });
  }
  if (anchors.length === 0) {
    return null;
  }
  const drawingXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${REL_NS}">` +
    anchors.join("") +
    `</xdr:wsDr>`;
  const drawingRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="${PKG_REL}">${rels.join("")}</Relationships>`;
  return { drawingXml, drawingRelsXml, mediaEntries, nextMediaFileIndex: mediaIdx };
}

export function groupFloatingPicturesBySheetIndex(
  workbook: Workbook,
  pictures: readonly XlsxFloatingPictureExport[],
): Map<number, XlsxFloatingPictureExport[]> {
  const map = new Map<number, XlsxFloatingPictureExport[]>();
  for (const p of pictures) {
    let idx = -1;
    for (let i = 0; i < workbook.sheetCount; i++) {
      const sh = workbook.getSheet(i);
      if (sh !== undefined && sh.name === p.sheetName) {
        idx = i;
        break;
      }
    }
    if (idx < 0) {
      continue;
    }
    const arr = map.get(idx);
    if (arr === undefined) {
      map.set(idx, [p]);
    } else {
      arr.push(p);
    }
  }
  return map;
}

export function parseRelationshipSelfClosingTags(xml: string): string[] {
  return xml.match(/<Relationship\b[^>]+\/>/g) ?? [];
}

export function maxRIdNumberFromRelationshipTags(tags: readonly string[]): number {
  let max = 0;
  for (const rel of tags) {
    const m = /Id="rId(\d+)"/.exec(rel);
    if (m !== null) {
      max = Math.max(max, Number(m[1]));
    }
  }
  return max;
}

export function mergeWorksheetRelsXml(
  pivotOrExistingRelsXml: string | undefined,
  drawingTarget: string | undefined,
): { relsXml: string; drawingRid?: string } | undefined {
  const existingTags =
    pivotOrExistingRelsXml !== undefined && pivotOrExistingRelsXml.trim() !== ""
      ? parseRelationshipSelfClosingTags(pivotOrExistingRelsXml)
      : [];
  if (drawingTarget === undefined && existingTags.length === 0) {
    return undefined;
  }
  const tags = [...existingTags];
  let drawingRid: string | undefined;
  if (drawingTarget !== undefined) {
    const next = maxRIdNumberFromRelationshipTags(tags) + 1;
    drawingRid = `rId${next}`;
    tags.push(
      `<Relationship Id="${drawingRid}" Type="${REL_WORKSHEET_DRAWING}" Target="${escapeXml(drawingTarget)}"/>`,
    );
  }
  const relsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="${PKG_REL}">${tags.join("")}</Relationships>`;
  return { relsXml, drawingRid };
}

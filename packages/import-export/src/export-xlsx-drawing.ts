import type { Workbook, Worksheet } from "@flexsheet/core";
import { escapeXml } from "./xml-escape.js";

/** 与 FlexSheet「图片格式 → 填充」一致；仅 `solid` 参与 XLSX 导出。 */
export interface XlsxFloatingPictureFrameFill {
  readonly kind: "none" | "solid" | "gradient" | "picture" | "pattern";
  /** `solid` 时 `#rrggbb` */
  readonly solidColor: string;
  /** 纯色填充透明度 0～100（100 为全透明） */
  readonly solidTransparencyPct: number;
}

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
  /** 源图像素宽；与 `imgBox*` 一并用于 OOXML `srcRect` 裁剪；缺省或 0 则整张图拉伸进占位（旧行为）。 */
  readonly naturalWidth?: number;
  readonly naturalHeight?: number;
  /** 相对裁剪框左上角的图片内容矩形（画布像素）；缺省视为铺满 `width`×`height`。 */
  readonly imgBoxX?: number;
  readonly imgBoxY?: number;
  readonly imgBoxW?: number;
  readonly imgBoxH?: number;
  /**
   * 裁剪框内纯色填充。写入 `xdr:spPr` 的 `a:solidFill`（边距区），与 `blipFill` 主体图分层，不合并栅格。
   */
  readonly frameFill?: XlsxFloatingPictureFrameFill;
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

function clamp01(n: number): number {
  return Math.min(1, Math.max(0, n));
}

function normalizeSolidColorHexForExport(input: string): string | null {
  const t = input.trim();
  if (/^#[0-9a-fA-F]{6}$/.test(t)) {
    return t.slice(1).toUpperCase();
  }
  if (/^#[0-9a-fA-F]{3}$/.test(t)) {
    const r = t[1]!;
    const g = t[2]!;
    const b = t[3]!;
    return `${r}${r}${g}${g}${b}${b}`.toUpperCase();
  }
  if (/^[0-9a-fA-F]{6}$/.test(t)) {
    return t.toUpperCase();
  }
  return null;
}

/**
 * `xdr:pic` / `a:spPr` 下纯色底（Excel 在几何区背后绘制；与 FlexSheet `frameFill` 对齐）。
 * `CT_PositiveFixedPercentage`：100000 = 100% 不透明度。
 */
function buildPicSpPrSolidFillXml(fill: XlsxFloatingPictureFrameFill | undefined): string {
  if (fill === undefined || fill.kind !== "solid") {
    return "";
  }
  const op = clamp01(1 - fill.solidTransparencyPct / 100);
  if (op <= 0) {
    return "";
  }
  const val = normalizeSolidColorHexForExport(fill.solidColor);
  if (val === null) {
    return "";
  }
  const alphaVal = Math.round(op * 100000);
  const alphaXml = alphaVal >= 100000 ? "" : `<a:alpha val="${alphaVal}"/>`;
  return `<a:solidFill><a:srgbClr val="${val}">${alphaXml}</a:srgbClr></a:solidFill>`;
}

/** OOXML `CT_RelativeRect`：从各边裁掉的占比（ST_Percentage，如 `"12.5%"`）。 */
function formatOoxmlPercentage(fraction: number): string {
  const p = Math.max(0, Math.min(100, fraction * 100));
  if (p <= 0) {
    return "0%";
  }
  if (p >= 100) {
    return "100%";
  }
  const s = p.toFixed(4).replace(/\.?0+$/, "");
  return `${s}%`;
}

/**
 * FlexSheet 侧是否仍需在导出前用 Canvas 处理位图（当前始终为 false）。
 * 填充与边距由 `spPr/a:solidFill` 与内嵌 `a:xfrm` 表达；源图滤镜在宿主侧按需烘焙。
 */
export function floatingPictureNeedsRasterForXlsxExport(_pic: XlsxFloatingPictureExport): boolean {
  return false;
}

/** 裁剪框内是否存在未被图片覆盖的留白（几何判断；导出用内嵌 `xfrm` + `solidFill`，不合并栅格）。 */
export function floatingPictureNeedsFrameCompositeForXlsx(pic: XlsxFloatingPictureExport): boolean {
  const fw = pic.width;
  const fh = pic.height;
  if (!(fw > 0) || !(fh > 0)) {
    return false;
  }
  const ibx = pic.imgBoxX ?? 0;
  const iby = pic.imgBoxY ?? 0;
  const ibw = pic.imgBoxW ?? fw;
  const ibh = pic.imgBoxH ?? fh;
  if (!(ibw > 0) || !(ibh > 0)) {
    return true;
  }
  const visL = Math.max(0, ibx);
  const visT = Math.max(0, iby);
  const visR = Math.min(fw, ibx + ibw);
  const visB = Math.min(fh, iby + ibh);
  const eps = 1e-3;
  if (visR <= visL + eps || visB <= visT + eps) {
    return true;
  }
  const cw = visR - visL;
  const ch = visB - visT;
  return cw < fw - eps || ch < fh - eps;
}

/**
 * 由 FlexSheet 裁剪模型推算 `blipFill/srcRect`（与 DOM：`object-fit: fill` + `img-wrap` 几何一致）。
 * 若无法推算或等价于未裁剪，返回 `undefined`（不写 `srcRect`）。
 */
export function floatingPictureSrcRectSides(
  pic: XlsxFloatingPictureExport,
): { readonly l: string; readonly t: string; readonly r: string; readonly b: string } | undefined {
  const nw = pic.naturalWidth ?? 0;
  const nh = pic.naturalHeight ?? 0;
  if (!(nw > 0) || !(nh > 0)) {
    return undefined;
  }
  const fw = pic.width;
  const fh = pic.height;
  if (!(fw > 0) || !(fh > 0)) {
    return undefined;
  }
  const ibx = pic.imgBoxX ?? 0;
  const iby = pic.imgBoxY ?? 0;
  const ibw = pic.imgBoxW ?? fw;
  const ibh = pic.imgBoxH ?? fh;
  if (!(ibw > 0) || !(ibh > 0)) {
    return undefined;
  }
  const visL = Math.max(0, ibx);
  const visT = Math.max(0, iby);
  const visR = Math.min(fw, ibx + ibw);
  const visB = Math.min(fh, iby + ibh);
  if (visR <= visL || visB <= visT) {
    return undefined;
  }
  const u0 = (nw * (visL - ibx)) / ibw;
  const u1 = (nw * (visR - ibx)) / ibw;
  const v0 = (nh * (visT - iby)) / ibh;
  const v1 = (nh * (visB - iby)) / ibh;
  const u0c = Math.max(0, Math.min(nw, u0));
  const u1c = Math.max(0, Math.min(nw, u1));
  const v0c = Math.max(0, Math.min(nh, v0));
  const v1c = Math.max(0, Math.min(nh, v1));
  if (u1c <= u0c || v1c <= v0c) {
    return undefined;
  }
  const l = u0c / nw;
  const r = (nw - u1c) / nw;
  const t = v0c / nh;
  const b = (nh - v1c) / nh;
  const eps = 1e-6;
  if (l <= eps && r <= eps && t <= eps && b <= eps) {
    return undefined;
  }
  return {
    l: formatOoxmlPercentage(l),
    t: formatOoxmlPercentage(t),
    r: formatOoxmlPercentage(r),
    b: formatOoxmlPercentage(b),
  };
}

function picHasImgBoxMargins(pic: XlsxFloatingPictureExport, viewZoom: number): boolean {
  const z = Math.max(1e-6, viewZoom);
  const W = pic.width / z;
  const H = pic.height / z;
  const ibx = (pic.imgBoxX ?? 0) / z;
  const iby = (pic.imgBoxY ?? 0) / z;
  const ibw = (pic.imgBoxW ?? pic.width) / z;
  const ibh = (pic.imgBoxH ?? pic.height) / z;
  const eps = 1e-3;
  return (
    ibx > eps || iby > eps || W - ibx - ibw > eps || H - iby - ibh > eps
  );
}

/**
 * 框内留白（`imgBox` 未铺满裁剪框）用 `a:srcRect` 的 **100000 分度整数** 表达（可为负），与 Excel / SpreadJS 一致；
 * 与 `import-xlsx-drawing` 中 `tryParsePicture` 的 `du/imgBoxW` 公式对偶。
 * 勿用 `stretch/fillRect`：进入裁剪模式时 Excel 易丢弃 fillRect，导致位图突然铺满 `spPr`。
 */
function buildMarginSrcRectXml(pic: XlsxFloatingPictureExport, viewZoom: number): string {
  const z = Math.max(1e-6, viewZoom);
  const W = pic.width / z;
  const H = pic.height / z;
  const ibx = (pic.imgBoxX ?? 0) / z;
  const iby = (pic.imgBoxY ?? 0) / z;
  const ibw = (pic.imgBoxW ?? pic.width) / z;
  const ibh = (pic.imgBoxH ?? pic.height) / z;
  const fl = -ibx / ibw;
  const fr = (ibw + ibx - W) / ibw;
  const ft = -iby / ibh;
  const fb = (ibh + iby - H) / ibh;
  const q = (x: number): string => String(Math.round(x * 100000));
  return `<a:srcRect l="${q(fl)}" t="${q(ft)}" r="${q(fr)}" b="${q(fb)}"/>`;
}

function buildBlipSrcRectAndStretchXml(
  pic: XlsxFloatingPictureExport,
  viewZoom: number,
): { readonly srcRectXml: string; readonly stretchXml: string } {
  const stretchXml = `<a:stretch/>`;
  if (picHasImgBoxMargins(pic, viewZoom)) {
    /** 与源图裁剪 `floatingPictureSrcRectSides` 并存时优先留白编码（极少数场景会损失额外源裁剪在 OOXML 中的独立表达）。 */
    return { srcRectXml: buildMarginSrcRectXml(pic, viewZoom), stretchXml };
  }
  const src = floatingPictureSrcRectSides(pic);
  if (src === undefined) {
    return { srcRectXml: `<a:srcRect l="0" t="0" r="0" b="0"/>`, stretchXml };
  }
  return {
    srcRectXml: `<a:srcRect l="${src.l}" t="${src.t}" r="${src.r}" b="${src.b}"/>`,
    stretchXml,
  };
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
  const cxFullEmu = Math.max(1, Math.round(box.w * EMU_PER_PX));
  const cyFullEmu = Math.max(1, Math.round(box.h * EMU_PER_PX));
  const rot = rotationToOoxmlRot60000(pic.rotationRad);
  const rotAttr = rot === 0 ? "" : ` rot="${rot}"`;
  const { srcRectXml, stretchXml } = buildBlipSrcRectAndStretchXml(pic, viewZoom);
  const spPrFillXml = buildPicSpPrSolidFillXml(pic.frameFill);
  /**
   * Excel 将 `solidFill` 铺在 `spPr/a:xfrm` 的整个几何上；`xfrm` 与锚点外框同大。
   * 框内图片几何用 `blipFill/srcRect`（整数分度，见 SpreadJS），勿用 `stretch/fillRect` 以免裁剪模式错乱。
   */
  const spPrLnXml = `<a:ln cmpd="sng" cap="flat"><a:noFill/><a:prstDash val="solid"/><a:round/></a:ln>`;
  return (
    `<xdr:oneCellAnchor>` +
    `<xdr:from>` +
    `<xdr:col>${from.col}</xdr:col>` +
    `<xdr:colOff>${from.colOff}</xdr:colOff>` +
    `<xdr:row>${from.row}</xdr:row>` +
    `<xdr:rowOff>${from.rowOff}</xdr:rowOff>` +
    `</xdr:from>` +
    `<xdr:ext cx="${cxFullEmu}" cy="${cyFullEmu}"/>` +
    `<xdr:pic>` +
    `<xdr:nvPicPr>` +
    `<xdr:cNvPr id="${shapeId}" name="${escapeXml(`Picture ${shapeId}`)}"/>` +
    `<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>` +
    `</xdr:nvPicPr>` +
    `<xdr:blipFill>` +
    `<a:blip xmlns:r="${REL_NS}" r:embed="${embedRid}"/>` +
    srcRectXml +
    stretchXml +
    `</xdr:blipFill>` +
    `<xdr:spPr>` +
    `<a:xfrm${rotAttr}>` +
    `<a:off x="0" y="0"/>` +
    `<a:ext cx="${cxFullEmu}" cy="${cyFullEmu}"/>` +
    `</a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    spPrFillXml +
    spPrLnXml +
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

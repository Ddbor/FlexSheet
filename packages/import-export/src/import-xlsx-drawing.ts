import type { Worksheet } from "@flexsheet/core";
import { REL_WORKSHEET_DRAWING, type XlsxFloatingPictureFrameFill } from "./export-xlsx-drawing.js";

const EMU_PER_PX = 9525;
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** 自 XLSX 工作表绘图解析出的浮动图（几何为工作表逻辑像素，加载 FlexSheet 时再乘 `viewZoom`）。 */
export interface XlsxImportedFloatingPicture {
  readonly sheetName: string;
  readonly sheetIndex: number;
  readonly anchorRow: number;
  readonly anchorCol: number;
  readonly relCX: number;
  readonly relCY: number;
  readonly width: number;
  readonly height: number;
  readonly rotationRad: number;
  readonly dataUrl: string;
  readonly naturalWidth: number;
  readonly naturalHeight: number;
  readonly imgBoxX: number;
  readonly imgBoxY: number;
  readonly imgBoxW: number;
  readonly imgBoxH: number;
  readonly frameFill?: XlsxFloatingPictureFrameFill;
}

/** `FlexSheet.loadWorkbook` 可选参数：导入 XLSX 时附带浮动图。 */
export interface FlexSheetLoadWorkbookOptions {
  readonly importedFloatingPictures?: readonly XlsxImportedFloatingPicture[];
  /** 默认取 `CanvasRenderer.getViewZoom()`；传入时可覆盖。 */
  readonly floatingPictureViewZoom?: number;
}

function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  if (root === null || root.localName === "parsererror") {
    throw new Error("XLSX：绘图 XML 解析失败");
  }
  return doc;
}

function childrenLocal(el: Element, local: string): Element[] {
  return [...el.children].filter((c) => c.localName === local);
}

function firstLocal(el: Element, local: string): Element | undefined {
  return childrenLocal(el, local)[0];
}

function normalizePartPath(path: string): string {
  const raw = path.replace(/\\/g, "/").replace(/^\/+/, "");
  const segs = raw.split("/");
  const out: string[] = [];
  for (const seg of segs) {
    if (seg === "" || seg === ".") {
      continue;
    }
    if (seg === "..") {
      if (out.length > 0) {
        out.pop();
      }
      continue;
    }
    out.push(seg);
  }
  return out.join("/");
}

function dirname(path: string): string {
  const p = normalizePartPath(path);
  const idx = p.lastIndexOf("/");
  if (idx < 0) {
    return "";
  }
  return p.slice(0, idx);
}

function resolvePartTarget(basePartPath: string, target: string): string {
  if (target.startsWith("/")) {
    return normalizePartPath(target.slice(1));
  }
  const baseDir = dirname(basePartPath);
  const joined = baseDir === "" ? target : `${baseDir}/${target}`;
  return normalizePartPath(joined);
}

function relsPartPathFor(partPath: string): string {
  const dir = dirname(partPath);
  const file = partPath.slice(partPath.lastIndexOf("/") + 1);
  return normalizePartPath(`${dir}/_rels/${file}.rels`);
}

function findRelTargetsByType(relsXml: string, type: string): string[] {
  const doc = parseXml(relsXml);
  const targets: string[] = [];
  for (const rel of [...doc.documentElement.children]) {
    if (rel.localName === "Relationship" && rel.getAttribute("Type") === type) {
      const target = rel.getAttribute("Target");
      if (target !== null) {
        targets.push(target);
      }
    }
  }
  return targets;
}

function parseDrawingRelsIdToTarget(relsXml: string): Map<string, string> {
  const doc = parseXml(relsXml);
  const map = new Map<string, string>();
  for (const rel of [...doc.documentElement.children]) {
    if (rel.localName !== "Relationship") {
      continue;
    }
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id !== null && target !== null) {
      map.set(id, target);
    }
  }
  return map;
}

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

function sheetDistToPx(
  sheet: Worksheet,
  col: number,
  colOffEmu: number,
  row: number,
  rowOffEmu: number,
): { x: number; y: number } {
  let x = 0;
  for (let c = 0; c < col; c++) {
    x += sheet.getColWidth(c);
  }
  x += colOffEmu / EMU_PER_PX;
  let y = 0;
  for (let r = 0; r < row; r++) {
    y += sheet.getRowHeight(r);
  }
  y += rowOffEmu / EMU_PER_PX;
  return { x, y };
}

function parseGridPos(el: Element): { col: number; colOff: number; row: number; rowOff: number } {
  const col = Math.max(0, Math.trunc(Number(firstLocal(el, "col")?.textContent ?? "0")));
  const colOff = Math.max(0, Math.trunc(Number(firstLocal(el, "colOff")?.textContent ?? "0")));
  const row = Math.max(0, Math.trunc(Number(firstLocal(el, "row")?.textContent ?? "0")));
  const rowOff = Math.max(0, Math.trunc(Number(firstLocal(el, "rowOff")?.textContent ?? "0")));
  return { col, colOff, row, rowOff };
}

function parseGeomOneCell(
  anchor: Element,
  sheet: Worksheet,
): { left: number; top: number; w: number; h: number } | null {
  const from = firstLocal(anchor, "from");
  const ext = firstLocal(anchor, "ext");
  if (from === undefined || ext === undefined) {
    return null;
  }
  const g = parseGridPos(from);
  const p0 = sheetDistToPx(sheet, g.col, g.colOff, g.row, g.rowOff);
  const cx = Number(ext.getAttribute("cx") ?? "0");
  const cy = Number(ext.getAttribute("cy") ?? "0");
  const w = Math.max(1, cx / EMU_PER_PX);
  const h = Math.max(1, cy / EMU_PER_PX);
  return { left: p0.x, top: p0.y, w, h };
}

function parseGeomTwoCell(
  anchor: Element,
  sheet: Worksheet,
): { left: number; top: number; w: number; h: number } | null {
  const from = firstLocal(anchor, "from");
  const to = firstLocal(anchor, "to");
  if (from === undefined || to === undefined) {
    return null;
  }
  const a = parseGridPos(from);
  const b = parseGridPos(to);
  const p0 = sheetDistToPx(sheet, a.col, a.colOff, a.row, a.rowOff);
  const p1 = sheetDistToPx(sheet, b.col, b.colOff, b.row, b.rowOff);
  const left = Math.min(p0.x, p1.x);
  const top = Math.min(p0.y, p1.y);
  const w = Math.max(1, Math.abs(p1.x - p0.x));
  const h = Math.max(1, Math.abs(p1.y - p0.y));
  return { left, top, w, h };
}

function parseGeomAbsolute(anchor: Element): { left: number; top: number; w: number; h: number } | null {
  const pos = firstLocal(anchor, "pos");
  const ext = firstLocal(anchor, "ext");
  if (pos === undefined || ext === undefined) {
    return null;
  }
  const x = Number(pos.getAttribute("x") ?? "0");
  const y = Number(pos.getAttribute("y") ?? "0");
  const cx = Number(ext.getAttribute("cx") ?? "0");
  const cy = Number(ext.getAttribute("cy") ?? "0");
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    return null;
  }
  return {
    left: x / EMU_PER_PX,
    top: y / EMU_PER_PX,
    w: cx / EMU_PER_PX,
    h: cy / EMU_PER_PX,
  };
}

function sheetPxToTopLeftCell(sheet: Worksheet, xPx: number, yPx: number): { col: number; row: number } {
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
  return { col, row };
}

function parseOoxmlPercentAttr(raw: string | null): number {
  if (raw === null || raw === "") {
    return 0;
  }
  const t = raw.trim();
  if (t.endsWith("%")) {
    const n = Number(t.slice(0, -1));
    return Number.isFinite(n) ? Math.min(1, Math.max(0, n / 100)) : 0;
  }
  const n = Number(t);
  if (!Number.isFinite(n)) {
    return 0;
  }
  if (n > 1 && n <= 100000) {
    return Math.min(1, Math.max(0, n / 100000));
  }
  return Math.min(1, Math.max(0, n));
}

function sniffImageMime(bytes: Uint8Array): "image/png" | "image/jpeg" | null {
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    return "image/png";
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return "image/jpeg";
  }
  return null;
}

function readPngDimensions(bytes: Uint8Array): { w: number; h: number } | null {
  if (bytes.length < 24) {
    return null;
  }
  const w = (bytes[16]! << 24) | (bytes[17]! << 16) | (bytes[18]! << 8) | bytes[19]!;
  const h = (bytes[20]! << 24) | (bytes[21]! << 16) | (bytes[22]! << 8) | bytes[23]!;
  if (!(w > 0) || !(h > 0)) {
    return null;
  }
  return { w, h };
}

function readJpegDimensions(bytes: Uint8Array): { w: number; h: number } | null {
  let i = 0;
  if (bytes[i++] !== 0xff || bytes[i++] !== 0xd8) {
    return null;
  }
  while (i < bytes.length - 1) {
    if (bytes[i++] !== 0xff) {
      continue;
    }
    const m = bytes[i++];
    if (m === undefined) {
      return null;
    }
    if (m === 0xd8 || m === 0xd9) {
      continue;
    }
    const lenHi = bytes[i];
    const lenLo = bytes[i + 1];
    if (lenHi === undefined || lenLo === undefined) {
      return null;
    }
    const segLen = (lenHi << 8) | lenLo;
    if (m >= 0xc0 && m <= 0xc3) {
      const h = (bytes[i + 3]! << 8) | bytes[i + 4]!;
      const w = (bytes[i + 5]! << 8) | bytes[i + 6]!;
      if (w > 0 && h > 0) {
        return { w, h };
      }
      return null;
    }
    i += segLen;
  }
  return null;
}

function readImageDimensions(bytes: Uint8Array, mime: "image/png" | "image/jpeg"): { w: number; h: number } {
  if (mime === "image/png") {
    return readPngDimensions(bytes) ?? { w: 0, h: 0 };
  }
  return readJpegDimensions(bytes) ?? { w: 0, h: 0 };
}

function bytesToDataUrl(bytes: Uint8Array, mime: "image/png" | "image/jpeg"): string {
  if (typeof btoa === "function") {
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }
    return `data:${mime};base64,${btoa(binary)}`;
  }
  const g = globalThis as unknown as {
    Buffer?: { from(data: Uint8Array): { toString(enc: "base64"): string } };
  };
  if (g.Buffer !== undefined) {
    return `data:${mime};base64,${g.Buffer.from(bytes).toString("base64")}`;
  }
  throw new Error("无法 base64 编码图片");
}

function parseSpPrSolidFill(spPr: Element): XlsxFloatingPictureFrameFill | undefined {
  const solid = firstLocal(spPr, "solidFill");
  if (solid === undefined) {
    return undefined;
  }
  const srgb = firstLocal(solid, "srgbClr");
  if (srgb === undefined) {
    return undefined;
  }
  const val = srgb.getAttribute("val");
  if (val === null || val === "") {
    return undefined;
  }
  const hex = val.length === 6 ? `#${val.toUpperCase()}` : `#${val.toUpperCase()}`;
  const alphaEl = firstLocal(srgb, "alpha");
  const alphaVal = Number(alphaEl?.getAttribute("val") ?? "100000");
  const op = Number.isFinite(alphaVal) ? Math.min(1, Math.max(0, alphaVal / 100000)) : 1;
  const solidTransparencyPct = Math.round((1 - op) * 100);
  return { kind: "solid", solidColor: hex, solidTransparencyPct };
}

function tryParsePicture(
  picEl: Element,
  geom: { left: number; top: number; w: number; h: number },
  embedIdToTarget: Map<string, string>,
  drawingBasePath: string,
  files: ReadonlyMap<string, Uint8Array>,
  sheet: Worksheet,
  sheetName: string,
  sheetIndex: number,
): XlsxImportedFloatingPicture | null {
  const blipFill = firstLocal(picEl, "blipFill");
  if (blipFill === undefined) {
    return null;
  }
  const blip = firstLocal(blipFill, "blip");
  if (blip === undefined) {
    return null;
  }
  const embedId =
    blip.getAttributeNS(REL_NS, "embed") ??
    blip.getAttribute("r:embed") ??
    blip.getAttribute("embed");
  if (embedId === null || embedId === "") {
    return null;
  }
  const relTarget = embedIdToTarget.get(embedId);
  if (relTarget === undefined) {
    return null;
  }
  const mediaPath = resolvePartTarget(drawingBasePath, relTarget);
  const bytes = files.get(mediaPath);
  if (bytes === undefined || bytes.length === 0) {
    return null;
  }
  const mime = sniffImageMime(bytes);
  if (mime === null) {
    return null;
  }
  const dataUrl = bytesToDataUrl(bytes, mime);
  const { w: nw, h: nh } = readImageDimensions(bytes, mime);

  const srcRect = firstLocal(blipFill, "srcRect");
  let imgBoxX = 0;
  let imgBoxY = 0;
  let imgBoxW = geom.w;
  let imgBoxH = geom.h;
  if (srcRect !== undefined && nw > 0 && nh > 0) {
    const fl = parseOoxmlPercentAttr(srcRect.getAttribute("l"));
    const ft = parseOoxmlPercentAttr(srcRect.getAttribute("t"));
    const fr = parseOoxmlPercentAttr(srcRect.getAttribute("r"));
    const fb = parseOoxmlPercentAttr(srcRect.getAttribute("b"));
    const u0 = fl * nw;
    const u1 = (1 - fr) * nw;
    const v0 = ft * nh;
    const v1 = (1 - fb) * nh;
    const du = u1 - u0;
    const dv = v1 - v0;
    if (du > 1e-6 && dv > 1e-6) {
      imgBoxW = (geom.w * nw) / du;
      imgBoxH = (geom.h * nh) / dv;
      imgBoxX = (-u0 * geom.w) / du;
      imgBoxY = (-v0 * geom.h) / dv;
    }
  }

  const spPr = firstLocal(picEl, "spPr");
  const frameFill = spPr !== undefined ? parseSpPrSolidFill(spPr) : undefined;

  const xfrm = spPr !== undefined ? firstLocal(spPr, "xfrm") : undefined;
  const rotAttr = xfrm?.getAttribute("rot");
  const rot60000 = rotAttr !== null && rotAttr !== undefined && rotAttr !== "" ? Number(rotAttr) : 0;
  const rotationRad =
    Number.isFinite(rot60000) && rot60000 !== 0 ? (rot60000 / 60000) * (Math.PI / 180) : 0;

  const tl = sheetPxToTopLeftCell(sheet, geom.left, geom.top);
  const merged = sheet.getMergeAnchorCell(tl.row, tl.col);
  const anchorRow = merged.row;
  const anchorCol = merged.col;
  const cr = mergedCellSheetRectPx(sheet, anchorRow, anchorCol);
  const centerX = geom.left + geom.w / 2;
  const centerY = geom.top + geom.h / 2;
  const relCX = centerX - (cr.x + cr.w / 2);
  const relCY = centerY - (cr.y + cr.h / 2);

  return {
    sheetName,
    sheetIndex,
    anchorRow,
    anchorCol,
    relCX,
    relCY,
    width: geom.w,
    height: geom.h,
    rotationRad,
    dataUrl,
    naturalWidth: nw,
    naturalHeight: nh,
    imgBoxX,
    imgBoxY,
    imgBoxW,
    imgBoxH,
    ...(frameFill !== undefined ? { frameFill } : {}),
  };
}

function walkElementSubtree(root: Element, visit: (el: Element) => void): void {
  visit(root);
  for (let i = 0; i < root.children.length; i++) {
    const c = root.children[i];
    if (c !== undefined) {
      walkElementSubtree(c, visit);
    }
  }
}

function collectPicturesUnder(
  anchorRoot: Element,
  geom: { left: number; top: number; w: number; h: number },
  embedIdToTarget: Map<string, string>,
  drawingBasePath: string,
  files: ReadonlyMap<string, Uint8Array>,
  sheet: Worksheet,
  sheetName: string,
  sheetIndex: number,
  out: XlsxImportedFloatingPicture[],
): void {
  walkElementSubtree(anchorRoot, (el) => {
    if (el.localName.toLowerCase() !== "pic") {
      return;
    }
    const pic = tryParsePicture(el, geom, embedIdToTarget, drawingBasePath, files, sheet, sheetName, sheetIndex);
    if (pic !== null) {
      out.push(pic);
    }
  });
}

function parseDrawingRoot(
  root: Element,
  embedIdToTarget: Map<string, string>,
  drawingBasePath: string,
  files: ReadonlyMap<string, Uint8Array>,
  sheet: Worksheet,
  sheetName: string,
  sheetIndex: number,
  out: XlsxImportedFloatingPicture[],
): void {
  for (const anchor of root.children) {
    const ln = anchor.localName;
    let geom: { left: number; top: number; w: number; h: number } | null = null;
    if (ln === "oneCellAnchor") {
      geom = parseGeomOneCell(anchor, sheet);
    } else if (ln === "twoCellAnchor") {
      geom = parseGeomTwoCell(anchor, sheet);
    } else if (ln === "absoluteAnchor") {
      geom = parseGeomAbsolute(anchor);
    } else {
      continue;
    }
    if (geom === null) {
      continue;
    }
    collectPicturesUnder(anchor, geom, embedIdToTarget, drawingBasePath, files, sheet, sheetName, sheetIndex, out);
  }
}

/**
 * 从已解压的 OPC 映射中读取单张工作表的 `drawing` 关系，解析 `xdr:pic`、嵌入图与 `spPr` 纯色填充。
 */
export function collectSheetFloatingPicturesFromXlsx(
  files: ReadonlyMap<string, Uint8Array>,
  sheetPartPath: string,
  sheet: Worksheet,
  sheetName: string,
  sheetIndex: number,
): XlsxImportedFloatingPicture[] {
  const relsPath = relsPartPathFor(sheetPartPath);
  const relsBytes = files.get(relsPath);
  if (relsBytes === undefined) {
    return [];
  }
  const relsXml = new TextDecoder("utf-8").decode(relsBytes);
  const drawingTargets = findRelTargetsByType(relsXml, REL_WORKSHEET_DRAWING);
  const out: XlsxImportedFloatingPicture[] = [];
  for (const target of drawingTargets) {
    const drawingPath = resolvePartTarget(sheetPartPath, target);
    const drawingBytes = files.get(drawingPath);
    if (drawingBytes === undefined) {
      continue;
    }
    const drawingXml = new TextDecoder("utf-8").decode(drawingBytes);
    const drawingRelsPath = relsPartPathFor(drawingPath);
    const drawingRelsBytes = files.get(drawingRelsPath);
    const embedIdToTarget =
      drawingRelsBytes !== undefined
        ? parseDrawingRelsIdToTarget(new TextDecoder("utf-8").decode(drawingRelsBytes))
        : new Map<string, string>();
    const doc = parseXml(drawingXml);
    const root = doc.documentElement;
    if (root === null) {
      continue;
    }
    if (root.localName === "wsDr" || root.localName.toLowerCase() === "wsdr") {
      parseDrawingRoot(root, embedIdToTarget, drawingPath, files, sheet, sheetName, sheetIndex, out);
    }
  }
  return out;
}

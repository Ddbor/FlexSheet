import type { Worksheet } from "@flexsheet/core";
import {
  REL_WORKSHEET_DRAWING,
  type XlsxFloatingPictureFrameFill,
  type XlsxFloatingPictureGradientRelativeRect,
  type XlsxFloatingPictureGradientStop,
} from "./export-xlsx-drawing.js";

const EMU_PER_PX = 9525;
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_WORKBOOK_THEME =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

/** Office 默认主题基色（6 位 hex，无 `#`）；`schemeClr` 无主题 part 时的兜底。 */
const DEFAULT_OFFICE_SCHEME_SRGB: Readonly<Record<string, string>> = {
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "44546A",
  lt2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
  hlink: "0563C1",
  folHlink: "954F72",
  tx1: "000000",
  tx2: "44546A",
  bg1: "FFFFFF",
  bg2: "E7E6E6",
};

/** 自 XLSX 工作表绘图解析出的浮动图（几何为工作表逻辑像素，加载 FlexSheet 时再乘 `viewZoom`）。 */
export interface XlsxImportedFloatingPicture {
  readonly sheetName: string;
  readonly sheetIndex: number;
  readonly anchorRow: number;
  readonly anchorCol: number;
  /** 相对合并锚点单元格左上角：图片左上角在工作表逻辑像素中的偏移（加载时再乘 `viewZoom`）。 */
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

/**
 * 自 `xl/theme/theme1.xml` 解析 `clrScheme`（accent1…），供绘图 `schemeClr` + 变换与渐变导入。
 */
export function parseWorkbookThemeSchemeColors(
  files: ReadonlyMap<string, Uint8Array>,
): ReadonlyMap<string, string> {
  const out = new Map<string, string>();
  const relsBytes = files.get("xl/_rels/workbook.xml.rels");
  if (relsBytes === undefined) {
    return out;
  }
  const relsText = new TextDecoder("utf-8").decode(relsBytes);
  const themeTargets = findRelTargetsByType(relsText, REL_WORKBOOK_THEME);
  const themeTarget = themeTargets[0];
  if (themeTarget === undefined) {
    return out;
  }
  const themePath = resolvePartTarget("xl/workbook.xml", themeTarget);
  const themeBytes = files.get(themePath);
  if (themeBytes === undefined) {
    return out;
  }
  const doc = parseXml(new TextDecoder("utf-8").decode(themeBytes));
  const root = doc.documentElement;
  if (root === null) {
    return out;
  }
  const themeElements = firstLocal(root, "themeElements") ?? root;
  const clrScheme = firstLocal(themeElements, "clrScheme");
  if (clrScheme === undefined) {
    return out;
  }
  for (const child of clrScheme.children) {
    const name = child.localName.toLowerCase();
    const srgb = firstLocal(child, "srgbClr");
    if (srgb !== undefined) {
      const v = srgb.getAttribute("val");
      if (v !== null) {
        const digits = v.replace(/^#/, "").toUpperCase();
        if (digits.length >= 6) {
          out.set(name, digits.slice(0, 6));
        }
      }
      continue;
    }
    const sysClr = firstLocal(child, "sysClr");
    if (sysClr !== undefined) {
      const last = sysClr.getAttribute("lastClr");
      if (last !== null) {
        const digits = last.replace(/^#/, "").toUpperCase();
        if (digits.length >= 6) {
          out.set(name, digits.slice(0, 6));
        }
      }
    }
  }
  return out;
}

function impClamp01(x: number): number {
  return Math.min(1, Math.max(0, x));
}

function impRgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
  const r1 = r / 255;
  const g1 = g / 255;
  const b1 = b / 255;
  const max = Math.max(r1, g1, b1);
  const min = Math.min(r1, g1, b1);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r1:
        h = (g1 - b1) / d + (g1 < b1 ? 6 : 0);
        break;
      case g1:
        h = (b1 - r1) / d + 2;
        break;
      default:
        h = (r1 - g1) / d + 4;
    }
    h /= 6;
  }
  return { h, s, l };
}

function impHslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  if (s === 0) {
    const v = Math.round(l * 255);
    return { r: v, g: v, b: v };
  }
  const hue2rgb = (p: number, q: number, t: number): number => {
    let tt = t;
    if (tt < 0) {
      tt += 1;
    }
    if (tt > 1) {
      tt -= 1;
    }
    if (tt < 1 / 6) {
      return p + (q - p) * 6 * tt;
    }
    if (tt < 1 / 2) {
      return q;
    }
    if (tt < 2 / 3) {
      return p + (q - p) * (2 / 3 - tt) * 6;
    }
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const rf = hue2rgb(p, q, h + 1 / 3);
  const gf = hue2rgb(p, q, h);
  const bf = hue2rgb(p, q, h - 1 / 3);
  return {
    r: Math.min(255, Math.max(0, Math.round(rf * 255))),
    g: Math.min(255, Math.max(0, Math.round(gf * 255))),
    b: Math.min(255, Math.max(0, Math.round(bf * 255))),
  };
}

/** `a:gs` 下 `srgbClr` / `schemeClr` 的子变换（tint/shade/satMod/lumMod/lumOff/alpha），顺序与 OOXML 一致。 */
function applyDrawingmlColorModifiers(
  r: number,
  g: number,
  b: number,
  a: number,
  parent: Element,
): { r: number; g: number; b: number; a: number } {
  let rf = r;
  let gf = g;
  let bf = b;
  let af = a;
  for (const ch of parent.children) {
    const ln = ch.localName;
    const raw = ch.getAttribute("val");
    const v = raw !== null && raw !== "" ? Number(raw) : NaN;
    const f = Number.isFinite(v) ? v / 100000 : 1;
    if (ln === "alpha") {
      af *= f;
      continue;
    }
    if (ln === "tint") {
      rf = rf + (255 - rf) * f;
      gf = gf + (255 - gf) * f;
      bf = bf + (255 - bf) * f;
      continue;
    }
    if (ln === "shade") {
      rf *= 1 - f;
      gf *= 1 - f;
      bf *= 1 - f;
      continue;
    }
    if (ln === "satMod" || ln === "lumMod" || ln === "lumOff") {
      const hsl = impRgbToHsl(rf, gf, bf);
      if (ln === "satMod") {
        hsl.s = impClamp01(hsl.s * f);
      }
      if (ln === "lumMod") {
        hsl.l = impClamp01(hsl.l * f);
      }
      if (ln === "lumOff") {
        hsl.l = impClamp01(hsl.l + f);
      }
      const nx = impHslToRgb(hsl.h, hsl.s, hsl.l);
      rf = nx.r;
      gf = nx.g;
      bf = nx.b;
    }
  }
  return {
    r: Math.min(255, Math.max(0, Math.round(rf))),
    g: Math.min(255, Math.max(0, Math.round(gf))),
    b: Math.min(255, Math.max(0, Math.round(bf))),
    a: impClamp01(af),
  };
}

function resolveSchemeClrBaseHexFromMap(
  schemeVal: string | null,
  schemeColors: ReadonlyMap<string, string>,
): string | undefined {
  if (schemeVal === null || schemeVal === "") {
    return undefined;
  }
  const k = schemeVal.trim().toLowerCase();
  if (k === "phclr") {
    return undefined;
  }
  const fromTheme = schemeColors.get(k);
  if (fromTheme !== undefined && fromTheme.length >= 6) {
    return fromTheme.slice(0, 6).toUpperCase();
  }
  return DEFAULT_OFFICE_SCHEME_SRGB[k];
}

function resolveGradientStopColor(
  el: Element,
  schemeColors: ReadonlyMap<string, string>,
): { r: number; g: number; b: number; a: number } | null {
  const ln = el.localName;
  if (ln === "srgbClr") {
    const raw = el.getAttribute("val")?.replace(/^#/, "") ?? "";
    if (raw.length < 6) {
      return null;
    }
    const r = parseInt(raw.slice(0, 2), 16);
    const g = parseInt(raw.slice(2, 4), 16);
    const b = parseInt(raw.slice(4, 6), 16);
    if (![r, g, b].every((x) => Number.isFinite(x))) {
      return null;
    }
    return applyDrawingmlColorModifiers(r, g, b, 1, el);
  }
  if (ln === "schemeClr") {
    const key = el.getAttribute("val")?.trim().toLowerCase() ?? "";
    const base = resolveSchemeClrBaseHexFromMap(key, schemeColors);
    if (base === undefined) {
      return null;
    }
    const r = parseInt(base.slice(0, 2), 16);
    const g = parseInt(base.slice(2, 4), 16);
    const b = parseInt(base.slice(4, 6), 16);
    if (![r, g, b].every((x) => Number.isFinite(x))) {
      return null;
    }
    return applyDrawingmlColorModifiers(r, g, b, 1, el);
  }
  return null;
}

function parseOoxmlRelativeRectLtrb(
  el: Element | undefined,
): XlsxFloatingPictureGradientRelativeRect | undefined {
  if (el === undefined) {
    return undefined;
  }
  const q = (a: string | null): number => {
    if (a === null || a === "") {
      return 0;
    }
    const n = Number(a);
    return Number.isFinite(n) ? Math.round(n) : 0;
  };
  return {
    l: q(el.getAttribute("l")),
    t: q(el.getAttribute("t")),
    r: q(el.getAttribute("r")),
    b: q(el.getAttribute("b")),
  };
}

function parseGradientStopsFromGsLst(
  gsLst: Element,
  schemeColors: ReadonlyMap<string, string>,
): XlsxFloatingPictureGradientStop[] | undefined {
  const stops: XlsxFloatingPictureGradientStop[] = [];
  for (const gs of childrenLocal(gsLst, "gs")) {
    const posAttr = gs.getAttribute("pos");
    const posRaw = posAttr !== null && posAttr !== "" ? Number(posAttr) : 0;
    const positionPct = Number.isFinite(posRaw) ? Math.min(100, Math.max(0, posRaw / 1000)) : 0;
    const colorEl = [...gs.children].find((c) => {
      const n = c.localName;
      return n === "srgbClr" || n === "schemeClr";
    });
    if (colorEl === undefined) {
      continue;
    }
    const rgba = resolveGradientStopColor(colorEl, schemeColors);
    if (rgba === null) {
      continue;
    }
    const transparencyPct = Math.round((1 - rgba.a) * 100);
    const hex =
      `#${rgba.r.toString(16).padStart(2, "0")}${rgba.g.toString(16).padStart(2, "0")}${rgba.b
        .toString(16)
        .padStart(2, "0")}`.toLowerCase();
    stops.push({
      positionPct,
      color: hex,
      transparencyPct,
      brightnessPct: 0,
    });
  }
  if (stops.length < 2) {
    return undefined;
  }
  stops.sort((a, b) => a.positionPct - b.positionPct);
  return stops;
}

function parseSpPrGradientFill(
  spPr: Element,
  schemeColors: ReadonlyMap<string, string>,
): XlsxFloatingPictureFrameFill | undefined {
  const grad = firstLocal(spPr, "gradFill");
  if (grad === undefined) {
    return undefined;
  }
  const gsLst = firstLocal(grad, "gsLst");
  if (gsLst === undefined) {
    return undefined;
  }
  const stops = parseGradientStopsFromGsLst(gsLst, schemeColors);
  if (stops === undefined) {
    return undefined;
  }
  const rot = grad.getAttribute("rotWithShape");
  const gradientRotateWithShape = rot !== "0" && rot !== "false";

  const lin = firstLocal(grad, "lin");
  if (lin !== undefined) {
    const angAttr = lin.getAttribute("ang");
    const ang60000 = angAttr !== null && angAttr !== "" ? Number(angAttr) : 5400000;
    const gradientAngleDeg = Number.isFinite(ang60000)
      ? (((ang60000 / 60000) % 360) + 360) % 360
      : 90;
    return {
      kind: "gradient",
      solidColor: "#000000",
      solidTransparencyPct: 0,
      gradientType: "linear",
      gradientAngleDeg,
      gradientStops: stops,
      gradientRotateWithShape,
      gradientPresetId: null,
    };
  }

  const pathEl = firstLocal(grad, "path");
  if (pathEl !== undefined) {
    const ftrEl = firstLocal(pathEl, "fillToRect");
    const tileEl = firstLocal(grad, "tileRect");
    const radialFillLtrb = parseOoxmlRelativeRectLtrb(ftrEl);
    const radialTileLtrb = parseOoxmlRelativeRectLtrb(tileEl);
    return {
      kind: "gradient",
      solidColor: "#000000",
      solidTransparencyPct: 0,
      gradientType: "radial",
      gradientAngleDeg: 90,
      gradientStops: stops,
      gradientRotateWithShape,
      gradientPresetId: null,
      ...(radialFillLtrb !== undefined ? { radialFillLtrb } : {}),
      ...(radialTileLtrb !== undefined ? { radialTileLtrb } : {}),
    };
  }

  return undefined;
}

function parseSpPrFrameFill(
  spPr: Element,
  schemeColors: ReadonlyMap<string, string>,
): XlsxFloatingPictureFrameFill | undefined {
  const g = parseSpPrGradientFill(spPr, schemeColors);
  if (g !== undefined) {
    return g;
  }
  return parseSpPrSolidFill(spPr, schemeColors);
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

/** 工作表水平方向：自左端起累计列宽（px→EMU），将绝对 x 坐标（EMU）换为逻辑 px（与 `sheetDistToPx` 一致）。 */
function sheetAbsEmuToPxX(sheet: Worksheet, xEmu: number): number {
  if (!Number.isFinite(xEmu)) {
    return 0;
  }
  if (xEmu <= 0) {
    return 0;
  }
  let remaining = xEmu;
  let xPx = 0;
  const lastCol = Math.max(0, sheet.colCount - 1);
  for (let c = 0; c <= lastCol; c++) {
    const wPx = sheet.getColWidth(c);
    const wEmu = wPx * EMU_PER_PX;
    if (remaining < wEmu || c === lastCol) {
      return xPx + remaining / EMU_PER_PX;
    }
    remaining -= wEmu;
    xPx += wPx;
  }
  return xPx + remaining / EMU_PER_PX;
}

/** 工作表垂直方向：自顶端起累计行高，将绝对 y 坐标（EMU）换为逻辑 px。 */
function sheetAbsEmuToPxY(sheet: Worksheet, yEmu: number): number {
  if (!Number.isFinite(yEmu)) {
    return 0;
  }
  if (yEmu <= 0) {
    return 0;
  }
  let remaining = yEmu;
  let yPx = 0;
  const lastRow = Math.max(0, sheet.rowCount - 1);
  for (let r = 0; r <= lastRow; r++) {
    const hPx = sheet.getRowHeight(r);
    const hEmu = hPx * EMU_PER_PX;
    if (remaining < hEmu || r === lastRow) {
      return yPx + remaining / EMU_PER_PX;
    }
    remaining -= hEmu;
    yPx += hPx;
  }
  return yPx + remaining / EMU_PER_PX;
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

function parseGeomAbsolute(
  anchor: Element,
): { left: number; top: number; w: number; h: number } | null {
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

/**
 * `twoCellAnchor`：宿主用 `from`/`to` 表示拖拽范围，可见矩形由 `pic/xdr:spPr/a:xfrm` 的绝对 EMU `off`/`ext` 给出（与 `from`/`to` 同一工作表坐标系）。
 */
function parsePicShapeGeomTwoCellFromSpPr(
  sheet: Worksheet,
  spPr: Element | undefined,
): { left: number; top: number; w: number; h: number } | null {
  if (spPr === undefined) {
    return null;
  }
  const xfrm = firstLocal(spPr, "xfrm");
  if (xfrm === undefined) {
    return null;
  }
  const offEl = firstLocal(xfrm, "off");
  const extEl = firstLocal(xfrm, "ext");
  if (offEl === undefined || extEl === undefined) {
    return null;
  }
  const cx = Number(extEl.getAttribute("cx") ?? "0");
  const cy = Number(extEl.getAttribute("cy") ?? "0");
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    return null;
  }
  const ox = Number(offEl.getAttribute("x") ?? "0");
  const oy = Number(offEl.getAttribute("y") ?? "0");
  if (!Number.isFinite(ox) || !Number.isFinite(oy)) {
    return null;
  }
  return {
    left: sheetAbsEmuToPxX(sheet, ox),
    top: sheetAbsEmuToPxY(sheet, oy),
    w: cx / EMU_PER_PX,
    h: cy / EMU_PER_PX,
  };
}

function sheetPxToTopLeftCell(
  sheet: Worksheet,
  xPx: number,
  yPx: number,
): { col: number; row: number } {
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

/**
 * DrawingML `CT_RelativeRect`（如 `a:srcRect`）侧：`ST_Percentage` 可为 `n%` 或整数（100000=100%，可负，见 ECMA-376）。
 */
function parseOoxmlRelativeRectRatio(raw: string | null): number {
  if (raw === null || raw === "") {
    return 0;
  }
  const t = raw.trim();
  if (t.endsWith("%")) {
    const n = Number(t.slice(0, -1));
    return Number.isFinite(n) ? n / 100 : 0;
  }
  const n = Number(t);
  if (!Number.isFinite(n)) {
    return 0;
  }
  return n / 100000;
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

function readImageDimensions(
  bytes: Uint8Array,
  mime: "image/png" | "image/jpeg",
): { w: number; h: number } {
  if (mime === "image/png") {
    return readPngDimensions(bytes) ?? { w: 0, h: 0 };
  }
  return readJpegDimensions(bytes) ?? { w: 0, h: 0 };
}

/**
 * `xdr:pic/xdr:spPr/a:xfrm` 内嵌矩形：Excel 用锚点 `xdr:ext` 表示外框（含填充边距），
 * 用 `a:off`/`a:ext` 表示位图拉伸目标区；缺省则与锚点同大、偏移 0。
 */
function parsePicSpPrInnerDestPx(
  spPr: Element | undefined,
  anchorW: number,
  anchorH: number,
): { offX: number; offY: number; destW: number; destH: number } {
  let offX = 0;
  let offY = 0;
  let destW = anchorW;
  let destH = anchorH;
  if (spPr === undefined) {
    return { offX, offY, destW, destH };
  }
  const xfrm = firstLocal(spPr, "xfrm");
  if (xfrm === undefined) {
    return { offX, offY, destW, destH };
  }
  const offEl = firstLocal(xfrm, "off");
  if (offEl !== undefined) {
    const ox = Number(offEl.getAttribute("x") ?? "0");
    const oy = Number(offEl.getAttribute("y") ?? "0");
    if (Number.isFinite(ox)) {
      offX = ox / EMU_PER_PX;
    }
    if (Number.isFinite(oy)) {
      offY = oy / EMU_PER_PX;
    }
  }
  const extEl = firstLocal(xfrm, "ext");
  if (extEl !== undefined) {
    const icx = Number(extEl.getAttribute("cx") ?? "0");
    const icy = Number(extEl.getAttribute("cy") ?? "0");
    if (Number.isFinite(icx) && icx > 0) {
      destW = icx / EMU_PER_PX;
    }
    if (Number.isFinite(icy) && icy > 0) {
      destH = icy / EMU_PER_PX;
    }
  }
  destW = Math.max(1e-6, destW);
  destH = Math.max(1e-6, destH);
  return { offX, offY, destW, destH };
}

/**
 * 历史导出曾用 `blipFill/stretch/fillRect` 表达框内留白；新导出改用 `srcRect` 整数（Excel 裁剪模式兼容）。
 * 有非零 fillRect 时仍解析，以读入旧包。
 */
function applyStretchFillRectToDest(
  blipFill: Element,
  outerW: number,
  outerH: number,
  inner: { offX: number; offY: number; destW: number; destH: number },
): { offX: number; offY: number; destW: number; destH: number } {
  const stretch = firstLocal(blipFill, "stretch");
  if (stretch === undefined) {
    return inner;
  }
  const fillRect = firstLocal(stretch, "fillRect");
  if (fillRect === undefined) {
    return inner;
  }
  const fl = parseOoxmlRelativeRectRatio(fillRect.getAttribute("l"));
  const ft = parseOoxmlRelativeRectRatio(fillRect.getAttribute("t"));
  const fr = parseOoxmlRelativeRectRatio(fillRect.getAttribute("r"));
  const fb = parseOoxmlRelativeRectRatio(fillRect.getAttribute("b"));
  const eps = 1e-9;
  if (fl <= eps && ft <= eps && fr <= eps && fb <= eps) {
    return inner;
  }
  const w = Math.max(1e-6, (1 - fl - fr) * outerW);
  const h = Math.max(1e-6, (1 - ft - fb) * outerH);
  return { offX: fl * outerW, offY: ft * outerH, destW: w, destH: h };
}

function solidFillFromHexAttrParent(
  hexRaw: string,
  alphaParent: Element,
): XlsxFloatingPictureFrameFill | undefined {
  const digits = hexRaw.replace(/^#/, "").toUpperCase();
  if (digits.length !== 6) {
    return undefined;
  }
  const hex = `#${digits}`;
  const alphaEl = firstLocal(alphaParent, "alpha");
  const alphaVal = Number(alphaEl?.getAttribute("val") ?? "100000");
  const op = Number.isFinite(alphaVal) ? Math.min(1, Math.max(0, alphaVal / 100000)) : 1;
  const solidTransparencyPct = Math.round((1 - op) * 100);
  return { kind: "solid", solidColor: hex, solidTransparencyPct };
}

function solidFillFromSrgbElement(srgb: Element): XlsxFloatingPictureFrameFill | undefined {
  const raw = srgb.getAttribute("val");
  if (raw === null || raw === "") {
    return undefined;
  }
  return solidFillFromHexAttrParent(raw, srgb);
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

function parseSpPrSolidFill(
  spPr: Element,
  schemeColors: ReadonlyMap<string, string>,
): XlsxFloatingPictureFrameFill | undefined {
  const solid = firstLocal(spPr, "solidFill");
  if (solid === undefined) {
    return undefined;
  }
  const srgb = firstLocal(solid, "srgbClr");
  if (srgb !== undefined) {
    return solidFillFromSrgbElement(srgb);
  }
  const sysClr = firstLocal(solid, "sysClr");
  if (sysClr !== undefined) {
    const last = sysClr.getAttribute("lastClr");
    if (last !== null && last.replace(/^#/, "").length >= 6) {
      return solidFillFromHexAttrParent(last, sysClr);
    }
  }
  const scheme = firstLocal(solid, "schemeClr");
  if (scheme !== undefined) {
    const base = resolveSchemeClrBaseHexFromMap(scheme.getAttribute("val"), schemeColors);
    if (base !== undefined) {
      return solidFillFromHexAttrParent(base, scheme);
    }
  }
  return undefined;
}

function tryParsePicture(
  picEl: Element,
  anchorKind: string,
  anchorGeom: { left: number; top: number; w: number; h: number },
  embedIdToTarget: Map<string, string>,
  drawingBasePath: string,
  files: ReadonlyMap<string, Uint8Array>,
  sheet: Worksheet,
  sheetName: string,
  sheetIndex: number,
  themeScheme: ReadonlyMap<string, string>,
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

  const spPr = firstLocal(picEl, "spPr");
  let geom: { left: number; top: number; w: number; h: number };
  let baseInner: { offX: number; offY: number; destW: number; destH: number };
  if (anchorKind === "twoCellAnchor") {
    const g2 = parsePicShapeGeomTwoCellFromSpPr(sheet, spPr);
    if (g2 !== null) {
      geom = g2;
      baseInner = { offX: 0, offY: 0, destW: geom.w, destH: geom.h };
    } else {
      geom = anchorGeom;
      baseInner = parsePicSpPrInnerDestPx(spPr, geom.w, geom.h);
    }
  } else {
    geom = anchorGeom;
    baseInner = parsePicSpPrInnerDestPx(spPr, geom.w, geom.h);
  }
  const { offX, offY, destW, destH } = applyStretchFillRectToDest(
    blipFill,
    geom.w,
    geom.h,
    baseInner,
  );

  const srcRect = firstLocal(blipFill, "srcRect");
  let imgBoxX = offX;
  let imgBoxY = offY;
  let imgBoxW = destW;
  let imgBoxH = destH;
  if (srcRect !== undefined && nw > 0 && nh > 0) {
    const fl = parseOoxmlRelativeRectRatio(srcRect.getAttribute("l"));
    const ft = parseOoxmlRelativeRectRatio(srcRect.getAttribute("t"));
    const fr = parseOoxmlRelativeRectRatio(srcRect.getAttribute("r"));
    const fb = parseOoxmlRelativeRectRatio(srcRect.getAttribute("b"));
    const u0 = fl * nw;
    const u1 = (1 - fr) * nw;
    const v0 = ft * nh;
    const v1 = (1 - fb) * nh;
    const du = u1 - u0;
    const dv = v1 - v0;
    if (du > 1e-6 && dv > 1e-6) {
      imgBoxW = (destW * nw) / du;
      imgBoxH = (destH * nh) / dv;
      imgBoxX = offX + (-u0 * destW) / du;
      imgBoxY = offY + (-v0 * destH) / dv;
    }
  }

  const frameFill = spPr !== undefined ? parseSpPrFrameFill(spPr, themeScheme) : undefined;

  const xfrm = spPr !== undefined ? firstLocal(spPr, "xfrm") : undefined;
  const rotAttr = xfrm?.getAttribute("rot");
  const rot60000 =
    rotAttr !== null && rotAttr !== undefined && rotAttr !== "" ? Number(rotAttr) : 0;
  const rotationRad =
    Number.isFinite(rot60000) && rot60000 !== 0 ? (rot60000 / 60000) * (Math.PI / 180) : 0;

  const tl = sheetPxToTopLeftCell(sheet, geom.left, geom.top);
  const merged = sheet.getMergeAnchorCell(tl.row, tl.col);
  const anchorRow = merged.row;
  const anchorCol = merged.col;
  const cr = mergedCellSheetRectPx(sheet, anchorRow, anchorCol);
  const relCX = geom.left - cr.x;
  const relCY = geom.top - cr.y;

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
  anchorKind: string,
  geom: { left: number; top: number; w: number; h: number },
  embedIdToTarget: Map<string, string>,
  drawingBasePath: string,
  files: ReadonlyMap<string, Uint8Array>,
  sheet: Worksheet,
  sheetName: string,
  sheetIndex: number,
  themeScheme: ReadonlyMap<string, string>,
  out: XlsxImportedFloatingPicture[],
): void {
  walkElementSubtree(anchorRoot, (el) => {
    if (el.localName.toLowerCase() !== "pic") {
      return;
    }
    const pic = tryParsePicture(
      el,
      anchorKind,
      geom,
      embedIdToTarget,
      drawingBasePath,
      files,
      sheet,
      sheetName,
      sheetIndex,
      themeScheme,
    );
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
  themeScheme: ReadonlyMap<string, string>,
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
    collectPicturesUnder(
      anchor,
      ln,
      geom,
      embedIdToTarget,
      drawingBasePath,
      files,
      sheet,
      sheetName,
      sheetIndex,
      themeScheme,
      out,
    );
  }
}

/**
 * 从已解压的 OPC 映射中读取单张工作表的 `drawing` 关系，解析 `xdr:pic`、嵌入图与 `spPr` 填充（纯色 / 线性渐变）。
 */
export function collectSheetFloatingPicturesFromXlsx(
  files: ReadonlyMap<string, Uint8Array>,
  sheetPartPath: string,
  sheet: Worksheet,
  sheetName: string,
  sheetIndex: number,
  themeScheme: ReadonlyMap<string, string> = new Map(),
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
      parseDrawingRoot(
        root,
        embedIdToTarget,
        drawingPath,
        files,
        sheet,
        sheetName,
        sheetIndex,
        themeScheme,
        out,
      );
    }
  }
  return out;
}

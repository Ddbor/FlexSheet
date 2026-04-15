import type { CellScalar, CellStyle } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import {
  SELECTION_OUTLINE_VISUAL_SCALE,
  VIEW_ZOOM_MAX,
  VIEW_ZOOM_MIN,
} from "./canvas-renderer-constants.js";

/** OOXML 风格 8 位 ARGB → CSS 颜色（含 alpha）。 */
export function argbToCss(argb: string): string | undefined {
  const s = argb.trim();
  if (!/^[\dA-Fa-f]{8}$/.test(s)) {
    return undefined;
  }
  const a = parseInt(s.slice(0, 2), 16) / 255;
  const R = s.slice(2, 4);
  const G = s.slice(4, 6);
  const B = s.slice(6, 8);
  const r = parseInt(R, 16);
  const g = parseInt(G, 16);
  const b = parseInt(B, 16);
  if (a >= 1 - 1e-6) {
    return `#${R}${G}${B}`;
  }
  return `rgba(${r},${g},${b},${a})`;
}

/**
 * 全表统一 1px 描边坐标：对齐到 CSS 像素中心，与 Excel 类表格一致。
 */
export function snapLine(v: number): number {
  return Math.round(v) + 0.5;
}

/**
 * 激活单元格/选区描边与行列标题强调带共用线宽（CSS 像素），随 viewZoom 线性缩放。
 */
export function viewScaledSelectionOutlineWidth(viewZoom: number): number {
  const z = Math.max(VIEW_ZOOM_MIN, Math.min(VIEW_ZOOM_MAX, viewZoom));
  return Math.max(1, 2 * SELECTION_OUTLINE_VISUAL_SCALE * z);
}

export function scaledColW(sheet: Worksheet, viewZoom: number): number {
  return sheet.defaultColWidth * viewZoom;
}

export function scaledRowH(sheet: Worksheet, viewZoom: number): number {
  return sheet.defaultRowHeight * viewZoom;
}

export function scaledColWidthAt(sheet: Worksheet, col: number, viewZoom: number): number {
  if (sheet.isColHidden(col)) {
    return 0;
  }
  return sheet.getColWidth(col) * viewZoom;
}

export function scaledRowHeightAt(sheet: Worksheet, row: number, viewZoom: number): number {
  if (sheet.isRowHidden(row)) {
    return 0;
  }
  return sheet.getRowHeight(row) * viewZoom;
}

export function scaledFontSizePx(base: number, viewZoom: number): number {
  return Math.max(6, Math.min(256, Math.round(base * viewZoom)));
}

const DEFAULT_CELL_FONT_SIZE_BASE_PX = 13;
const DEFAULT_CELL_FONT_FAMILY = "system-ui, -apple-system, sans-serif";

/** 未设置 `fontSizePt` 时与历史画布默认一致（约 13px 逻辑高）；否则按 pt→px（96dpi）换算。 */
export function cellStyleLogicalFontSizeBasePx(style: CellStyle | null | undefined): number {
  const pt = style?.fontSizePt;
  if (pt !== undefined && pt > 0) {
    return (pt * 96) / 72;
  }
  return DEFAULT_CELL_FONT_SIZE_BASE_PX;
}

export function cellStyleFontFamilyCss(style: CellStyle | null | undefined): string {
  const f = style?.fontFamily;
  if (f !== undefined && f.trim() !== "") {
    return f.trim();
  }
  return DEFAULT_CELL_FONT_FAMILY;
}

/** 供 `ctx.font` 与单元格内联编辑器同步。 */
export function buildCellCanvasFont(style: CellStyle | null | undefined, viewZoom: number): string {
  const basePx = cellStyleLogicalFontSizeBasePx(style);
  return buildCellCanvasFontWithLogicalPx(style, basePx, viewZoom);
}

/** 指定逻辑字号（px）构造 `ctx.font`（用于缩小字体填充等）。 */
export function buildCellCanvasFontWithLogicalPx(
  style: CellStyle | null | undefined,
  logicalPx: number,
  viewZoom: number,
): string {
  const italic = style?.italic === true ? "italic " : "";
  const weight = style?.bold === true ? "600" : "400";
  const px = scaledFontSizePx(logicalPx, viewZoom);
  const family = cellStyleFontFamilyCss(style);
  return `${italic}${weight} ${px}px ${family}`;
}

export function formatCellDisplay(value: CellScalar): string {
  if (value === null || value === "") {
    return "";
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return String(value);
}

export function truncateText(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxWidth: number,
): string {
  if (ctx.measureText(text).width <= maxWidth) {
    return text;
  }
  let s = text;
  const ell = "…";
  while (s.length > 0 && ctx.measureText(s + ell).width > maxWidth) {
    s = s.slice(0, -1);
  }
  return s + ell;
}

/** 与表体绘制一致的按宽度折行（用于自动行高等）。 */
export function wrapCellLines(ctx: CanvasRenderingContext2D, text: string, maxW: number): string[] {
  if (maxW <= 0) {
    return text.split("\n");
  }
  const out: string[] = [];
  for (const para of text.split("\n")) {
    if (para.length === 0) {
      out.push("");
      continue;
    }
    let rest = para;
    while (rest.length > 0) {
      if (ctx.measureText(rest).width <= maxW) {
        out.push(rest);
        break;
      }
      let lo = 1;
      let hi = rest.length;
      let best = 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (ctx.measureText(rest.slice(0, mid)).width <= maxW) {
          best = mid;
          lo = mid + 1;
        } else {
          hi = mid - 1;
        }
      }
      if (best < 1) {
        best = 1;
      }
      out.push(rest.slice(0, best));
      rest = rest.slice(best);
    }
  }
  return out;
}

/** 与 `Worksheet.getColumnAutoFilterSortHint` 一致：`null` 表示无排序箭头。 */
export type AutoFilterGlyphSortHint = "asc" | "desc" | null;

const AUTO_FILTER_TRI_BLUE = "#185abd";
const AUTO_FILTER_FUNNEL = "#323130";

function drawFilledChevronDown(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  halfW: number,
  halfH: number,
  fill: string,
): void {
  ctx.fillStyle = fill;
  ctx.beginPath();
  ctx.moveTo(cx - halfW, cy - halfH);
  ctx.lineTo(cx + halfW, cy - halfH);
  ctx.lineTo(cx, cy + halfH);
  ctx.closePath();
  ctx.fill();
}

function drawSortArrow(
  ctx: CanvasRenderingContext2D,
  cx: number,
  midY: number,
  direction: "asc" | "desc",
  color: string,
): void {
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.35;
  ctx.lineCap = "round";
  ctx.lineJoin = "miter";
  const stem = 5.2;
  const head = 2.4;
  if (direction === "asc") {
    const y0 = midY + stem * 0.35;
    const y1 = midY - stem * 0.55;
    ctx.beginPath();
    ctx.moveTo(cx, y0);
    ctx.lineTo(cx, y1);
    ctx.moveTo(cx - head, y1 + head * 0.9);
    ctx.lineTo(cx, y1);
    ctx.lineTo(cx + head, y1 + head * 0.9);
    ctx.stroke();
  } else {
    const y0 = midY - stem * 0.35;
    const y1 = midY + stem * 0.55;
    ctx.beginPath();
    ctx.moveTo(cx, y0);
    ctx.lineTo(cx, y1);
    ctx.moveTo(cx - head, y1 - head * 0.9);
    ctx.lineTo(cx, y1);
    ctx.lineTo(cx + head, y1 - head * 0.9);
    ctx.stroke();
  }
}

function drawFunnel(ctx: CanvasRenderingContext2D, cx: number, midY: number, fill: string): void {
  ctx.fillStyle = fill;
  const topW = 7;
  const botW = 4.2;
  const h = 4.2;
  const y0 = midY - h * 0.35;
  const y1 = midY + h * 0.45;
  ctx.beginPath();
  ctx.moveTo(cx - topW / 2, y0);
  ctx.lineTo(cx + topW / 2, y0);
  ctx.lineTo(cx + botW / 2, y1);
  ctx.lineTo(cx - botW / 2, y1);
  ctx.closePath();
  ctx.fill();
  ctx.fillRect(cx - 0.75, y1, 1.5, 3.2);
}

export interface AutoFilterGlyphPaintOptions {
  readonly narrowed: boolean;
  readonly sortHint: AutoFilterGlyphSortHint;
  readonly borderColor: string;
}

/**
 * 列筛选按钮四态：1 默认（实心下拉）2 升序（下拉+上箭头）3 降序（下拉+下箭头）4 筛选（漏斗，可叠加排序箭头）。
 */
export function paintAutoFilterDropdownGlyph(
  ctx: CanvasRenderingContext2D,
  bx: number,
  by: number,
  innerW: number,
  innerH: number,
  options: AutoFilterGlyphPaintOptions,
): void {
  const bxs = snapLine(bx);
  const bys = snapLine(by);
  const { narrowed, sortHint, borderColor } = options;
  const midY = bys + innerH / 2;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(Math.floor(bxs), Math.floor(bys), Math.ceil(innerW), Math.ceil(innerH));
  ctx.strokeStyle = borderColor;
  ctx.lineWidth = 1;
  ctx.strokeRect(bxs, bys, innerW, innerH);

  const showSortArrow = sortHint === "asc" || sortHint === "desc";
  const arrowX = bxs + innerW - 4.2;

  if (narrowed) {
    const funnelX = showSortArrow ? bxs + innerW * 0.34 : bxs + innerW / 2;
    drawFunnel(ctx, funnelX, midY, AUTO_FILTER_FUNNEL);
    if (sortHint === "asc") {
      drawSortArrow(ctx, arrowX, midY, "asc", AUTO_FILTER_TRI_BLUE);
    } else if (sortHint === "desc") {
      drawSortArrow(ctx, arrowX, midY, "desc", AUTO_FILTER_TRI_BLUE);
    }
    return;
  }

  const triX = showSortArrow ? bxs + innerW * 0.3 : bxs + innerW / 2;
  const triY = midY + 0.8;
  drawFilledChevronDown(ctx, triX, triY, 3.6, 2.8, AUTO_FILTER_TRI_BLUE);

  if (sortHint === "asc") {
    drawSortArrow(ctx, arrowX, midY, "asc", AUTO_FILTER_TRI_BLUE);
  } else if (sortHint === "desc") {
    drawSortArrow(ctx, arrowX, midY, "desc", AUTO_FILTER_TRI_BLUE);
  }
}

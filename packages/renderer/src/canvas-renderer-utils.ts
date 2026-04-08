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
  const italic = style?.italic === true ? "italic " : "";
  const weight = style?.bold === true ? "600" : "400";
  const basePx = cellStyleLogicalFontSizeBasePx(style);
  const px = scaledFontSizePx(basePx, viewZoom);
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

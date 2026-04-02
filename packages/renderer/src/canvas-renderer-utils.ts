import type { CellScalar } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import { SELECTION_OUTLINE_VISUAL_SCALE, VIEW_ZOOM_MAX, VIEW_ZOOM_MIN } from "./canvas-renderer-constants.js";

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

export function scaledFontSizePx(base: number, viewZoom: number): number {
  return Math.max(8, Math.min(36, Math.round(base * viewZoom)));
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

export function truncateText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string {
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

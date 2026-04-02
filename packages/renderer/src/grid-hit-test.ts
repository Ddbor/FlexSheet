import type { Worksheet } from "@flexsheet/core";
import type { FrozenLayout } from "./viewport.js";
import { clampScroll, type ViewportScrollLimits } from "./viewport.js";

/** 行列标题区命中（不含表体单元格）。 */
export type HeadingHit =
  | { readonly kind: "selectAllCorner" }
  | { readonly kind: "columnHeader"; readonly col: number }
  | { readonly kind: "rowHeader"; readonly row: number };

/**
 * 画布坐标命中左上角全选角、列标题或行标题（与 `hitTestCell` 坐标系一致）。
 * 点在表体或网格外返回 null。
 */
export function hitTestHeadingPointer(
  canvasX: number,
  canvasY: number,
  sheet: Worksheet,
  layout: FrozenLayout,
  scrollX: number,
  scrollY: number,
  scale = 1,
): HeadingHit | null {
  const { headerW, headerH } = layout;
  const colW = sheet.defaultColWidth * scale;
  const rowH = sheet.defaultRowHeight * scale;

  if (canvasX < headerW && canvasY < headerH) {
    return { kind: "selectAllCorner" };
  }

  if (canvasX >= headerW && canvasY < headerH) {
    const gx = canvasX - headerW;
    const col = columnIndexFromGx(gx, sheet, layout, scrollX, colW);
    if (col === null) {
      return null;
    }
    return { kind: "columnHeader", col };
  }

  if (canvasX < headerW && canvasY >= headerH) {
    const gy = canvasY - headerH;
    const row = rowIndexFromGy(gy, sheet, layout, scrollY, rowH);
    if (row === null) {
      return null;
    }
    return { kind: "rowHeader", row };
  }

  return null;
}

function columnIndexFromGx(
  gx: number,
  sheet: Worksheet,
  layout: FrozenLayout,
  scrollX: number,
  colW: number,
): number | null {
  const { frozenCols, frozenWidthPx } = layout;
  let col: number;
  if (frozenCols > 0 && gx < frozenWidthPx) {
    col = clampIndex(Math.floor(gx / colW), 0, frozenCols - 1);
  } else {
    const gxScroll = gx - frozenWidthPx + scrollX;
    col = frozenCols + Math.floor(gxScroll / colW);
  }
  if (col < 0 || col >= sheet.colCount) {
    return null;
  }
  return col;
}

function rowIndexFromGy(
  gy: number,
  sheet: Worksheet,
  layout: FrozenLayout,
  scrollY: number,
  rowH: number,
): number | null {
  const { frozenRows, frozenHeightPx } = layout;
  let row: number;
  if (frozenRows > 0 && gy < frozenHeightPx) {
    row = clampIndex(Math.floor(gy / rowH), 0, frozenRows - 1);
  } else {
    const gyScroll = gy - frozenHeightPx + scrollY;
    row = frozenRows + Math.floor(gyScroll / rowH);
  }
  if (row < 0 || row >= sheet.rowCount) {
    return null;
  }
  return row;
}

/**
 * 画布坐标（相对 canvas 的 CSS 像素，与 CellLeftX/CellTopY 一致）→ 单元格索引。
 * 点在行列标题或左上角装饰区内时返回 null。
 */
export function hitTestCell(
  canvasX: number,
  canvasY: number,
  sheet: Worksheet,
  layout: FrozenLayout,
  scrollX: number,
  scrollY: number,
  scale = 1,
): { row: number; col: number } | null {
  const { headerW, headerH } = layout;
  if (canvasX < headerW || canvasY < headerH) {
    return null;
  }

  const colW = sheet.defaultColWidth * scale;
  const rowH = sheet.defaultRowHeight * scale;
  const gx = canvasX - headerW;
  const gy = canvasY - headerH;

  const col = columnIndexFromGx(gx, sheet, layout, scrollX, colW);
  const row = rowIndexFromGy(gy, sheet, layout, scrollY, rowH);
  if (col === null || row === null) {
    return null;
  }
  return { row, col };
}

function clampIndex(v: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, v));
}

/**
 * 调整 scroll 使 (row,col) 在可滚动区内尽可能可见（冻结区无需改 scroll）。
 */
export function scrollToRevealCell(
  sheet: Worksheet,
  layout: FrozenLayout,
  limits: ViewportScrollLimits,
  row: number,
  col: number,
  scrollX: number,
  scrollY: number,
  scale = 1,
): { scrollX: number; scrollY: number } {
  const colW = sheet.defaultColWidth * scale;
  const rowH = sheet.defaultRowHeight * scale;
  const { frozenCols, frozenRows, scrollViewportW, scrollViewportH } = layout;

  let sx = scrollX;
  let sy = scrollY;

  if (col >= frozenCols) {
    const offset = (col - frozenCols) * colW;
    const lo = offset - scrollViewportW + colW;
    const hi = offset;
    if (lo <= hi) {
      sx = Math.max(lo, Math.min(hi, sx));
    } else {
      sx = offset;
    }
  }

  if (row >= frozenRows) {
    const offset = (row - frozenRows) * rowH;
    const lo = offset - scrollViewportH + rowH;
    const hi = offset;
    if (lo <= hi) {
      sy = Math.max(lo, Math.min(hi, sy));
    } else {
      sy = offset;
    }
  }

  return clampScroll(sx, sy, limits);
}

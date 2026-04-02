import type { Worksheet } from "@flexsheet/core";
import { scaledColWidthAt, scaledRowHeightAt } from "./canvas-renderer-utils.js";
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

  if (canvasX < headerW && canvasY < headerH) {
    return { kind: "selectAllCorner" };
  }

  if (canvasX >= headerW && canvasY < headerH) {
    const gx = canvasX - headerW;
    const col = columnIndexFromGx(gx, sheet, layout, scrollX, scale);
    if (col === null) {
      return null;
    }
    return { kind: "columnHeader", col };
  }

  if (canvasX < headerW && canvasY >= headerH) {
    const gy = canvasY - headerH;
    const row = rowIndexFromGy(gy, sheet, layout, scrollY, scale);
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
  scale: number,
): number | null {
  const { frozenCols, frozenWidthPx } = layout;
  const col =
    gx < frozenWidthPx
      ? locateByOffset(sheet, 0, gx, scale, "col")
      : locateByOffset(sheet, frozenCols, gx - frozenWidthPx + scrollX, scale, "col");
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
  scale: number,
): number | null {
  const { frozenRows, frozenHeightPx } = layout;
  const row =
    gy < frozenHeightPx
      ? locateByOffset(sheet, 0, gy, scale, "row")
      : locateByOffset(sheet, frozenRows, gy - frozenHeightPx + scrollY, scale, "row");
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

  const gx = canvasX - headerW;
  const gy = canvasY - headerH;

  const col = columnIndexFromGx(gx, sheet, layout, scrollX, scale);
  const row = rowIndexFromGy(gy, sheet, layout, scrollY, scale);
  if (col === null || row === null) {
    return null;
  }
  return { row, col };
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
  const { frozenCols, frozenRows, scrollViewportW, scrollViewportH } = layout;

  let sx = scrollX;
  let sy = scrollY;

  if (col >= frozenCols) {
    const offset = sumSizes(sheet, frozenCols, col, scale, "col");
    const colW = scaledColWidthAt(sheet, col, scale);
    const lo = offset - scrollViewportW + colW;
    const hi = offset;
    if (lo <= hi) {
      sx = Math.max(lo, Math.min(hi, sx));
    } else {
      sx = offset;
    }
  }

  if (row >= frozenRows) {
    const offset = sumSizes(sheet, frozenRows, row, scale, "row");
    const rowH = scaledRowHeightAt(sheet, row, scale);
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

function locateByOffset(
  sheet: Worksheet,
  start: number,
  offset: number,
  scale: number,
  axis: "row" | "col",
): number {
  let remain = Math.max(0, offset);
  const count = axis === "row" ? sheet.rowCount : sheet.colCount;
  for (let i = start; i < count; i++) {
    const size =
      axis === "row" ? scaledRowHeightAt(sheet, i, scale) : scaledColWidthAt(sheet, i, scale);
    if (remain < size) {
      return i;
    }
    remain -= size;
  }
  return Math.max(start, count - 1);
}

function sumSizes(
  sheet: Worksheet,
  start: number,
  endExcl: number,
  scale: number,
  axis: "row" | "col",
): number {
  let s = 0;
  for (let i = start; i < endExcl; i++) {
    s += axis === "row" ? scaledRowHeightAt(sheet, i, scale) : scaledColWidthAt(sheet, i, scale);
  }
  return s;
}

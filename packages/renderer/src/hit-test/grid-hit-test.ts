import type { Worksheet } from "@flexsheet/core";
import { cellLeftX, cellTopY } from "../canvas/canvas-renderer-geometry.js";
import { scaledColWidthAt, scaledRowHeightAt } from "../canvas/canvas-renderer-utils.js";
import type { FrozenLayout } from "../layout/viewport.js";
import { clampScroll, type ViewportScrollLimits } from "../layout/viewport.js";

/** 列标题右侧筛选按钮占用宽度（与 `paintColumnHeaderCell` 一致，画布 CSS 像素）。 */
export const COLUMN_HEADER_FILTER_BUTTON_CSS_PX = 20;

/**
 * 表体「列筛选」锚点格右侧为下拉按钮预留的宽度（与 `paintBodyCellTexts` / `paintBodyAutoFilterAnchors` 一致）。
 * 非锚点格或未启用列筛选时为 0。
 */
export function bodyColumnAutoFilterTextReservePx(
  sheet: Worksheet,
  row: number,
  col: number,
): number {
  const meta = sheet.getColumnAutoFilterMeta(col);
  if (meta?.uiKind !== "body" || meta.bodyAnchorRow !== row) {
    return 0;
  }
  const anchor = sheet.getMergeAnchorCell(row, col);
  if (anchor.row !== row || anchor.col !== col) {
    return 0;
  }
  return COLUMN_HEADER_FILTER_BUTTON_CSS_PX + 4;
}

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

/**
 * 命中列标题上的筛选下拉按钮时返回列索引；否则返回 null（与绘制区域一致）。
 */
export function hitTestColumnHeaderFilterButton(
  canvasX: number,
  canvasY: number,
  sheet: Worksheet,
  layout: FrozenLayout,
  scrollX: number,
  _scrollY: number,
  scale = 1,
): number | null {
  const { headerW, headerH } = layout;
  if (canvasX < headerW || canvasY < 0 || canvasY > headerH + 0.5) {
    return null;
  }
  const gx = canvasX - headerW;
  const col = columnIndexFromGx(gx, sheet, layout, scrollX, scale);
  if (col === null || !sheet.hasColumnAutoFilter(col)) {
    return null;
  }
  const headerMeta = sheet.getColumnAutoFilterMeta(col);
  if (headerMeta?.uiKind !== "header") {
    return null;
  }
  const x0 = cellLeftX(sheet, layout, col, scale, scrollX);
  const colW = scaledColWidthAt(sheet, col, scale);
  const pad = 2;
  const btnW = COLUMN_HEADER_FILTER_BUTTON_CSS_PX + pad;
  const rel = canvasX - x0;
  if (rel >= colW - btnW && rel <= colW) {
    return col;
  }
  return null;
}

/**
 * 命中表体内「列筛选」锚点格上的下拉按钮时返回列索引（与 `paintBodyAutoFilterAnchors` 一致）。
 */
export function hitTestBodyCellAutoFilterButton(
  canvasX: number,
  canvasY: number,
  sheet: Worksheet,
  layout: FrozenLayout,
  scrollX: number,
  scrollY: number,
  scale = 1,
): number | null {
  const { headerW, headerH } = layout;
  if (canvasX < headerW || canvasY < headerH) {
    return null;
  }
  const cellHit = hitTestCell(canvasX, canvasY, sheet, layout, scrollX, scrollY, scale);
  if (cellHit === null) {
    return null;
  }
  const meta = sheet.getColumnAutoFilterMeta(cellHit.col);
  if (meta?.uiKind !== "body" || meta.bodyAnchorRow !== cellHit.row) {
    return null;
  }
  const anchor = sheet.getMergeAnchorCell(cellHit.row, cellHit.col);
  if (anchor.row !== cellHit.row || anchor.col !== cellHit.col) {
    return null;
  }
  const x0 = cellLeftX(sheet, layout, cellHit.col, scale, scrollX);
  const y0 = cellTopY(sheet, layout, cellHit.row, scale, scrollY);
  const colW = scaledColWidthAt(sheet, cellHit.col, scale);
  const rowH = scaledRowHeightAt(sheet, cellHit.row, scale);
  const pad = 2;
  const btnW = COLUMN_HEADER_FILTER_BUTTON_CSS_PX + pad;
  const relX = canvasX - x0;
  const relY = canvasY - y0;
  const bandH = Math.min(rowH, 22);
  if (relX >= colW - btnW && relX <= colW && relY >= 0 && relY <= bandH) {
    return cellHit.col;
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

import type { Worksheet } from "@flexsheet/core";
import type { FrozenLayout } from "./viewport.js";
import { scaledColW, scaledRowH } from "./canvas-renderer-utils.js";

/** 列 `c` 左边缘画布 X（含冻结与滚动）。 */
export function cellLeftX(
  sheet: Worksheet,
  layout: FrozenLayout,
  c: number,
  viewZoom: number,
  scrollX: number,
): number {
  const colW = scaledColW(sheet, viewZoom);
  const { headerW, frozenCols, frozenWidthPx } = layout;
  if (c < frozenCols) {
    return headerW + c * colW;
  }
  return headerW + frozenWidthPx + (c - frozenCols) * colW - scrollX;
}

/** 行 `r` 顶边画布 Y（含冻结与滚动）。 */
export function cellTopY(
  sheet: Worksheet,
  layout: FrozenLayout,
  r: number,
  viewZoom: number,
  scrollY: number,
): number {
  const rowH = scaledRowH(sheet, viewZoom);
  const { headerH, frozenRows, frozenHeightPx } = layout;
  if (r < frozenRows) {
    return headerH + r * rowH;
  }
  return headerH + frozenHeightPx + (r - frozenRows) * rowH - scrollY;
}

export function cellIntersectsCanvas(
  x: number,
  y: number,
  colW: number,
  rowH: number,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
): boolean {
  return x + colW > headerW && x < canvasW && y + rowH > headerH && y < canvasH;
}

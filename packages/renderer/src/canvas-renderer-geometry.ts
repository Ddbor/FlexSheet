import type { Worksheet } from "@flexsheet/core";
import type { FrozenLayout } from "./viewport.js";
import { scaledColWidthAt, scaledRowHeightAt } from "./canvas-renderer-utils.js";

/** 列 `c` 左边缘画布 X（含冻结与滚动）。 */
export function cellLeftX(
  sheet: Worksheet,
  layout: FrozenLayout,
  c: number,
  viewZoom: number,
  scrollX: number,
): number {
  const { headerW, frozenCols, frozenWidthPx } = layout;
  const sumColWidths = (start: number, endExcl: number): number => {
    let s = 0;
    for (let i = start; i < endExcl; i++) {
      s += scaledColWidthAt(sheet, i, viewZoom);
    }
    return s;
  };
  if (c < frozenCols) {
    return headerW + sumColWidths(0, c);
  }
  return headerW + frozenWidthPx + sumColWidths(frozenCols, c) - scrollX;
}

/** 行 `r` 顶边画布 Y（含冻结与滚动）。 */
export function cellTopY(
  sheet: Worksheet,
  layout: FrozenLayout,
  r: number,
  viewZoom: number,
  scrollY: number,
): number {
  const { headerH, frozenRows, frozenHeightPx } = layout;
  const sumRowHeights = (start: number, endExcl: number): number => {
    let s = 0;
    for (let i = start; i < endExcl; i++) {
      s += scaledRowHeightAt(sheet, i, viewZoom);
    }
    return s;
  };
  if (r < frozenRows) {
    return headerH + sumRowHeights(0, r);
  }
  return headerH + frozenHeightPx + sumRowHeights(frozenRows, r) - scrollY;
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

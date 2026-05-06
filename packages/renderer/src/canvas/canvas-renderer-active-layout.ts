import type { Workbook } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import { buildFrozenLayout, computeScrollLimits, type FrozenLayout, type ViewportScrollLimits } from "../layout/viewport.js";
import { HEADER_STRIP_BASE_HEIGHT, HEADER_STRIP_BASE_WIDTH } from "./canvas-renderer-constants.js";

export interface ActiveFrozenContext {
  readonly sheet: Worksheet;
  readonly w: number;
  readonly h: number;
  readonly hw: number;
  readonly hh: number;
  readonly layout: FrozenLayout;
  readonly limits: ViewportScrollLimits;
}

/**
 * 活动表 + 与 `paint` / `clampScrollToActiveSheet` 一致的冻结布局（含表头占位）。
 */
export function tryActiveFrozenContext(
  workbook: Workbook,
  canvas: HTMLCanvasElement,
  showHeadings: boolean,
  viewZoom: number,
  frozenRows: number,
  frozenCols: number,
): ActiveFrozenContext | null {
  const sheet = workbook.getActiveSheet();
  if (sheet === undefined) {
    return null;
  }
  const w = canvas.clientWidth;
  const h = canvas.clientHeight;
  const hw = showHeadings ? HEADER_STRIP_BASE_WIDTH * viewZoom : 0;
  const hh = showHeadings ? HEADER_STRIP_BASE_HEIGHT * viewZoom : 0;
  const layout = buildFrozenLayout(sheet, hw, hh, w, h, frozenRows, frozenCols, viewZoom);
  const limits = computeScrollLimits(sheet, layout, viewZoom);
  return { sheet, w, h, hw, hh, layout, limits };
}

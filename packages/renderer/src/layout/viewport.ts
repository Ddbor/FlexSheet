import type { Worksheet } from "@flexsheet/core";
import { scaledColWidthAt, scaledRowHeightAt } from "../canvas/canvas-renderer-utils.js";

/** 视口与虚拟滚动：scroll 仅作用于「非冻结」行列区域（像素）。 */
export interface ViewportScrollLimits {
  readonly maxScrollX: number;
  readonly maxScrollY: number;
}

export interface FrozenLayout {
  readonly headerW: number;
  readonly headerH: number;
  readonly frozenCols: number;
  readonly frozenRows: number;
  readonly frozenWidthPx: number;
  readonly frozenHeightPx: number;
  /** 可滚动区（表体右下角象限）像素宽高。 */
  readonly scrollViewportW: number;
  readonly scrollViewportH: number;
}

/**
 * @param scale 视图缩放（与 CanvasRenderer.viewZoom 一致），用于行列像素尺寸。
 */
export function buildFrozenLayout(
  sheet: Worksheet,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  frozenRows: number,
  frozenCols: number,
  scale = 1,
): FrozenLayout {
  const fc = Math.max(0, Math.min(frozenCols, sheet.colCount));
  const fr = Math.max(0, Math.min(frozenRows, sheet.rowCount));
  let fw = 0;
  for (let c = 0; c < fc; c++) {
    fw += scaledColWidthAt(sheet, c, scale);
  }
  let fh = 0;
  for (let r = 0; r < fr; r++) {
    fh += scaledRowHeightAt(sheet, r, scale);
  }
  const scrollViewportW = Math.max(0, canvasW - headerW - fw);
  const scrollViewportH = Math.max(0, canvasH - headerH - fh);
  return {
    headerW,
    headerH,
    frozenCols: fc,
    frozenRows: fr,
    frozenWidthPx: fw,
    frozenHeightPx: fh,
    scrollViewportW,
    scrollViewportH,
  };
}

export function computeScrollLimits(
  sheet: Worksheet,
  layout: FrozenLayout,
  scale = 1,
): ViewportScrollLimits {
  const { scrollViewportW, scrollViewportH } = layout;
  let contentW = 0;
  for (let c = layout.frozenCols; c < sheet.colCount; c++) {
    contentW += scaledColWidthAt(sheet, c, scale);
  }
  let contentH = 0;
  for (let r = layout.frozenRows; r < sheet.rowCount; r++) {
    contentH += scaledRowHeightAt(sheet, r, scale);
  }
  return {
    maxScrollX: Math.max(0, contentW - scrollViewportW),
    maxScrollY: Math.max(0, contentH - scrollViewportH),
  };
}

export function clampScroll(
  scrollX: number,
  scrollY: number,
  limits: ViewportScrollLimits,
): { scrollX: number; scrollY: number } {
  return {
    scrollX: Math.max(0, Math.min(scrollX, limits.maxScrollX)),
    scrollY: Math.max(0, Math.min(scrollY, limits.maxScrollY)),
  };
}

export interface VisibleScrollRange {
  readonly startCol: number;
  readonly endCol: number;
  readonly startRow: number;
  readonly endRow: number;
}

/**
 * 右下角「双轴滚动」区在 buffer 下的可见行列索引（含端点）。
 */
export function visibleScrollableCellRange(
  sheet: Worksheet,
  layout: FrozenLayout,
  scrollX: number,
  scrollY: number,
  buffer: number,
  scale = 1,
): VisibleScrollRange {
  const { frozenCols, frozenRows, scrollViewportW, scrollViewportH } = layout;

  if (sheet.colCount <= 0 || sheet.rowCount <= 0) {
    return { startCol: 0, endCol: -1, startRow: 0, endRow: -1 };
  }

  let startCol = frozenCols;
  let endCol = sheet.colCount - 1;
  if (frozenCols >= sheet.colCount) {
    startCol = 0;
    endCol = sheet.colCount - 1;
  } else if (scrollViewportW <= 0) {
    startCol = frozenCols;
    endCol = frozenCols - 1;
  } else {
    const col = locateIndexByOffset(sheet, frozenCols, scrollX, scale, "col");
    const end = locateIndexByOffset(sheet, frozenCols, scrollX + scrollViewportW, scale, "col");
    startCol = Math.max(frozenCols, col - buffer);
    endCol = Math.min(sheet.colCount - 1, end + buffer);
  }

  let startRow = frozenRows;
  let endRow = sheet.rowCount - 1;
  if (frozenRows >= sheet.rowCount) {
    startRow = 0;
    endRow = sheet.rowCount - 1;
  } else if (scrollViewportH <= 0) {
    startRow = frozenRows;
    endRow = frozenRows - 1;
  } else {
    const row = locateIndexByOffset(sheet, frozenRows, scrollY, scale, "row");
    const end = locateIndexByOffset(sheet, frozenRows, scrollY + scrollViewportH, scale, "row");
    startRow = Math.max(frozenRows, row - buffer);
    endRow = Math.min(sheet.rowCount - 1, end + buffer);
  }

  return { startCol, endCol, startRow, endRow };
}

function locateIndexByOffset(
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

/**
 * 表体绘制/网格线用的行列闭区间：冻结区 ∪ 可见滚动区。
 */
export function bodyPaintExtents(
  sheet: Worksheet,
  layout: FrozenLayout,
  scrollX: number,
  scrollY: number,
  buffer: number,
  scale = 1,
): VisibleScrollRange {
  const vr = visibleScrollableCellRange(sheet, layout, scrollX, scrollY, buffer, scale);
  if (sheet.colCount <= 0 || sheet.rowCount <= 0) {
    return { startCol: 0, endCol: -1, startRow: 0, endRow: -1 };
  }

  let c0: number;
  let c1: number;
  if (layout.frozenCols >= sheet.colCount) {
    c0 = 0;
    c1 = sheet.colCount - 1;
  } else if (layout.frozenCols > 0) {
    c0 = 0;
    c1 = Math.max(layout.frozenCols - 1, vr.endCol);
  } else {
    c0 = vr.startCol;
    c1 = vr.endCol;
  }

  let r0: number;
  let r1: number;
  if (layout.frozenRows >= sheet.rowCount) {
    r0 = 0;
    r1 = sheet.rowCount - 1;
  } else if (layout.frozenRows > 0) {
    r0 = 0;
    r1 = Math.max(layout.frozenRows - 1, vr.endRow);
  } else {
    r0 = vr.startRow;
    r1 = vr.endRow;
  }

  return { startCol: c0, endCol: c1, startRow: r0, endRow: r1 };
}

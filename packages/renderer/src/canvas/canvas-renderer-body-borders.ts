import { resolveCellBorderStroke } from "@flexsheet/core";
import type {
  CellBorderSide,
  CellStyle,
  ConditionalFormattingOverlay,
  Worksheet,
} from "@flexsheet/core";
import { getConditionalFormattingCellOverlayCached } from "./canvas-renderer-cf-overlay.js";
import { cellIntersectsCanvas, cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import {
  argbToCss,
  scaledColWidthAt,
  scaledRowHeightAt,
  snapLine,
} from "./canvas-renderer-utils.js";
import type { FrozenLayout } from "../layout/viewport.js";

/** 与 `BodyPaintEnv` 中画布边框绘制所需字段一致（避免与 body 模块循环引用）。 */
export interface CellBorderPaintEnv {
  readonly ctx: CanvasRenderingContext2D;
  readonly viewZoom: number;
  readonly scrollX: number;
  readonly scrollY: number;
}

function mergeBorderFromCf(
  base: CellStyle | null,
  cf: ConditionalFormattingOverlay | null,
): CellStyle | null {
  if (cf === null) {
    return base;
  }
  const out: CellStyle = { ...(base ?? {}) };
  if (cf.borderTop !== undefined) {
    out.borderTop = cf.borderTop;
  }
  if (cf.borderLeft !== undefined) {
    out.borderLeft = cf.borderLeft;
  }
  if (cf.borderBottom !== undefined) {
    out.borderBottom = cf.borderBottom;
  }
  if (cf.borderRight !== undefined) {
    out.borderRight = cf.borderRight;
  }
  return Object.keys(out).length > 0 ? out : base;
}

function effectiveStyle(
  sheet: Worksheet,
  row: number,
  col: number,
  cfOverlayCellCache: Map<string, ConditionalFormattingOverlay | null> | undefined,
): CellStyle | null {
  const a = sheet.getMergeAnchorCell(row, col);
  const base = sheet.getCell(a.row, a.col).style;
  const cf = getConditionalFormattingCellOverlayCached(sheet, a.row, a.col, cfOverlayCellCache);
  return mergeBorderFromCf(base, cf);
}

function cellRightX(
  sheet: Worksheet,
  layout: FrozenLayout,
  row: number,
  col: number,
  viewZoom: number,
  scrollX: number,
): number {
  const info = sheet.getMergedRectInfo(row, col);
  let x = cellLeftX(sheet, layout, info.anchorCol, viewZoom, scrollX);
  for (let cc = info.anchorCol; cc < info.anchorCol + info.colSpan; cc++) {
    x += scaledColWidthAt(sheet, cc, viewZoom);
  }
  return x;
}

function cellBottomY(
  sheet: Worksheet,
  layout: FrozenLayout,
  row: number,
  col: number,
  viewZoom: number,
  scrollY: number,
): number {
  const info = sheet.getMergedRectInfo(row, col);
  let y = cellTopY(sheet, layout, info.anchorRow, viewZoom, scrollY);
  for (let rr = info.anchorRow; rr < info.anchorRow + info.rowSpan; rr++) {
    y += scaledRowHeightAt(sheet, rr, viewZoom);
  }
  return y;
}

function borderCssColor(side: CellBorderSide, fallback: string): string {
  const a = side.colorArgb?.trim();
  if (a !== undefined && a !== "" && /^[\dA-Fa-f]{8}$/i.test(a)) {
    return argbToCss(a) ?? fallback;
  }
  return fallback;
}

function strokeHorizontal(
  ctx: CanvasRenderingContext2D,
  x1: number,
  x2: number,
  y: number,
  side: CellBorderSide,
  viewZoom: number,
): void {
  const color = borderCssColor(side, "#000000");
  const sty = resolveCellBorderStroke(side, viewZoom);
  const ys = snapLine(y);
  ctx.strokeStyle = color;
  ctx.lineCap = sty.lineCap;
  if (sty.double) {
    const half = sty.gap / 2;
    ctx.lineWidth = sty.lineWidth;
    ctx.setLineDash([]);
    ctx.beginPath();
    ctx.moveTo(x1, ys - half);
    ctx.lineTo(x2, ys - half);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(x1, ys + half);
    ctx.lineTo(x2, ys + half);
    ctx.stroke();
  } else {
    ctx.lineWidth = sty.lineWidth;
    if (sty.lineDash !== null) {
      ctx.setLineDash([...sty.lineDash]);
    } else {
      ctx.setLineDash([]);
    }
    ctx.beginPath();
    ctx.moveTo(x1, ys);
    ctx.lineTo(x2, ys);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

function strokeVertical(
  ctx: CanvasRenderingContext2D,
  x: number,
  y1: number,
  y2: number,
  side: CellBorderSide,
  viewZoom: number,
): void {
  const color = borderCssColor(side, "#000000");
  const sty = resolveCellBorderStroke(side, viewZoom);
  const xs = snapLine(x);
  ctx.strokeStyle = color;
  ctx.lineCap = sty.lineCap;
  if (sty.double) {
    const half = sty.gap / 2;
    ctx.lineWidth = sty.lineWidth;
    ctx.setLineDash([]);
    ctx.beginPath();
    ctx.moveTo(xs - half, y1);
    ctx.lineTo(xs - half, y2);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(xs + half, y1);
    ctx.lineTo(xs + half, y2);
    ctx.stroke();
  } else {
    ctx.lineWidth = sty.lineWidth;
    if (sty.lineDash !== null) {
      ctx.setLineDash([...sty.lineDash]);
    } else {
      ctx.setLineDash([]);
    }
    ctx.beginPath();
    ctx.moveTo(xs, y1);
    ctx.lineTo(xs, y2);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

function hasAnyBorder(st: CellStyle | null | undefined): boolean {
  if (st === null || st === undefined) {
    return false;
  }
  return (
    st.borderTop !== undefined ||
    st.borderLeft !== undefined ||
    st.borderBottom !== undefined ||
    st.borderRight !== undefined
  );
}

/**
 * 在网格线之后、文本之前绘制单元格边框（与默认网格线独立）。
 */
export function paintBodyCellBorders(
  env: CellBorderPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  r0: number,
  r1: number,
  c0: number,
  c1: number,
  cfOverlayCellCache?: Map<string, ConditionalFormattingOverlay | null>,
): void {
  const { ctx } = env;
  ctx.save();
  ctx.lineCap = "butt";
  ctx.lineJoin = "miter";

  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      if (sheet.isMergeCoveredCell(r, c)) {
        continue;
      }
      const info = sheet.getMergedRectInfo(r, c);
      let colW = scaledColWidthAt(sheet, c, env.viewZoom);
      let rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
      let x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      let y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      if (info.rowSpan > 1 || info.colSpan > 1) {
        colW = 0;
        for (let cc = info.anchorCol; cc < info.anchorCol + info.colSpan; cc++) {
          colW += scaledColWidthAt(sheet, cc, env.viewZoom);
        }
        rowH = 0;
        for (let rr = info.anchorRow; rr < info.anchorRow + info.rowSpan; rr++) {
          rowH += scaledRowHeightAt(sheet, rr, env.viewZoom);
        }
        x = cellLeftX(sheet, layout, info.anchorCol, env.viewZoom, env.scrollX);
        y = cellTopY(sheet, layout, info.anchorRow, env.viewZoom, env.scrollY);
      }
      if (colW <= 0 || rowH <= 0) {
        continue;
      }
      if (!cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
        continue;
      }

      const ar = info.anchorRow;
      const ac = info.anchorCol;
      const st = sheet.getCell(ar, ac).style;
      if (!hasAnyBorder(st)) {
        continue;
      }

      const leftX = x;
      const topY = y;
      const rightX = x + colW;
      const bottomY = y + rowH;

      const bt = st?.borderTop;
      const bl = st?.borderLeft;
      const bb = st?.borderBottom;
      const br = st?.borderRight;

      // 共享边只描一次：水平线由「上方格子的 borderBottom」绘制，下方格子的 borderTop 在上方已有 bottom 时跳过；
      // 垂直线由「左侧格子的 borderRight」绘制，右侧格子的 borderLeft 在左侧已有 right 时跳过。
      // 若上下/左右两侧同时跳过（旧逻辑），则内部网格线会整条消失。

      if (bt !== undefined) {
        let skip = false;
        if (ar > 0) {
          const na = sheet.getMergeAnchorCell(ar - 1, ac);
          const nb = effectiveStyle(sheet, na.row, na.col, cfOverlayCellCache);
          const nbBottom = cellBottomY(sheet, layout, na.row, na.col, env.viewZoom, env.scrollY);
          if (Math.abs(nbBottom - topY) < 0.75 && nb?.borderBottom !== undefined) {
            skip = true;
          }
        }
        if (!skip) {
          strokeHorizontal(ctx, leftX, rightX, topY, bt, env.viewZoom);
        }
      }

      if (bb !== undefined) {
        strokeHorizontal(ctx, leftX, rightX, bottomY, bb, env.viewZoom);
      }

      if (bl !== undefined) {
        let skip = false;
        if (ac > 0) {
          const na = sheet.getMergeAnchorCell(ar, ac - 1);
          const nb = effectiveStyle(sheet, na.row, na.col, cfOverlayCellCache);
          const nbRight = cellRightX(sheet, layout, na.row, na.col, env.viewZoom, env.scrollX);
          if (Math.abs(nbRight - leftX) < 0.75 && nb?.borderRight !== undefined) {
            skip = true;
          }
        }
        if (!skip) {
          strokeVertical(ctx, leftX, topY, bottomY, bl, env.viewZoom);
        }
      }

      if (br !== undefined) {
        strokeVertical(ctx, rightX, topY, bottomY, br, env.viewZoom);
      }
    }
  }

  ctx.restore();
}

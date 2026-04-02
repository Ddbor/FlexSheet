import type { SelectionPaintSnapshot } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import { columnIndexToLabel } from "@flexsheet/shared";
import type { SheetTheme } from "@flexsheet/theme";
import { cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import { getClampedSelectionSpan } from "./canvas-renderer-selection-span.js";
import {
  scaledColW,
  scaledColWidthAt,
  scaledFontSizePx,
  scaledRowH,
  scaledRowHeightAt,
  snapLine,
  viewScaledSelectionOutlineWidth,
} from "./canvas-renderer-utils.js";
import type { FrozenLayout } from "./viewport.js";

export interface HeaderPaintEnv {
  readonly ctx: CanvasRenderingContext2D;
  readonly theme: SheetTheme;
  readonly viewZoom: number;
  readonly viewportBuffer: number;
  readonly scrollX: number;
  readonly scrollY: number;
}

export function drawCorner(
  ctx: CanvasRenderingContext2D,
  theme: SheetTheme,
  headerW: number,
  headerH: number,
): void {
  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(0, 0, headerW, headerH);
  ctx.strokeStyle = theme.headerLineColor;
  ctx.lineWidth = 1;
  ctx.beginPath();
  const xt = snapLine(0);
  const yt = snapLine(0);
  const xr = snapLine(headerW);
  const yb = snapLine(headerH);
  ctx.moveTo(xt, yt);
  ctx.lineTo(xr, yt);
  ctx.moveTo(xt, yt);
  ctx.lineTo(xt, yb);
  ctx.stroke();

  const inset = 4;
  const leg = Math.max(5, Math.min(headerW, headerH) * 0.36);
  const brx = headerW - inset;
  const bry = headerH - inset;
  ctx.beginPath();
  ctx.moveTo(brx - leg, bry);
  ctx.lineTo(brx, bry);
  ctx.lineTo(brx, bry - leg);
  ctx.closePath();
  ctx.fillStyle = theme.headerColor;
  ctx.globalAlpha = 0.38;
  ctx.fill();
  ctx.globalAlpha = 1;
  ctx.strokeStyle = theme.headerLineColor;
  ctx.lineWidth = 1;
  ctx.stroke();
}

export function drawRulerOverlay(
  env: HeaderPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
): void {
  const { ctx } = env;
  const yRuler = snapLine(headerH);
  ctx.strokeStyle = env.theme.gridLineColor;
  ctx.lineWidth = 1;
  const { frozenCols } = layout;
  const maxC = sheet.colCount - 1;
  const fw = layout.frozenWidthPx;
  const sx0 = headerW + fw;
  const tickH = 4;
  const drawTicks = (c0: number, c1: number): void => {
    for (let c = c0; c <= c1; c++) {
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const colW = scaledColWidthAt(sheet, c, env.viewZoom);
      if (x + colW <= headerW || x >= canvasW) {
        continue;
      }
      ctx.beginPath();
      ctx.moveTo(snapLine(x), yRuler);
      ctx.lineTo(snapLine(x), yRuler - tickH);
      ctx.stroke();
    }
  };
  if (frozenCols > 0) {
    drawTicks(0, Math.min(frozenCols - 1, maxC));
  }
  if (layout.scrollViewportW > 0 && frozenCols <= maxC) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(sx0, 0, canvasW - sx0, headerH);
    ctx.clip();
    const colW = scaledColW(sheet, env.viewZoom);
    const first = frozenCols + Math.floor(env.scrollX / colW) - 1;
    const last = frozenCols + Math.ceil((env.scrollX + layout.scrollViewportW) / colW) + 1;
    drawTicks(Math.max(frozenCols, first), Math.min(maxC, last));
    ctx.restore();
  }
}

function paintColumnHeaderCell(
  ctx: CanvasRenderingContext2D,
  theme: SheetTheme,
  viewZoom: number,
  c: number,
  x: number,
  colW: number,
  headerH: number,
  isSelectionColumn: boolean,
): void {
  ctx.fillStyle = isSelectionColumn ? theme.headerActiveBg : theme.headerBg;
  ctx.fillRect(x, 0, colW, headerH);
  ctx.fillStyle = isSelectionColumn ? theme.activeCellBorderColor : theme.headerColor;
  ctx.font = `${isSelectionColumn ? "bold " : ""}${scaledFontSizePx(12, viewZoom)}px system-ui, -apple-system, sans-serif`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(columnIndexToLabel(c), x + colW / 2, headerH / 2);
}

function paintRowHeaderCell(
  ctx: CanvasRenderingContext2D,
  theme: SheetTheme,
  viewZoom: number,
  r: number,
  y: number,
  rowH: number,
  headerW: number,
  isSelectionRow: boolean,
): void {
  ctx.fillStyle = isSelectionRow ? theme.headerActiveBg : theme.headerBg;
  ctx.fillRect(0, y, headerW, rowH);
  ctx.fillStyle = isSelectionRow ? theme.activeCellBorderColor : theme.headerColor;
  ctx.font = `${isSelectionRow ? "bold " : ""}${scaledFontSizePx(12, viewZoom)}px system-ui, -apple-system, sans-serif`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(String(r + 1), headerW / 2, y + rowH / 2);
}

export function drawAllColumnHeaders(
  env: HeaderPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  selectionSnap: SelectionPaintSnapshot | null,
): void {
  const { ctx } = env;
  const selSpan = getClampedSelectionSpan(sheet, selectionSnap);
  const columnInSelection = (c: number): boolean =>
    selSpan !== null && c >= selSpan.startCol && c <= selSpan.endCol;
  const { frozenCols } = layout;
  const buf = env.viewportBuffer;
  const maxC = sheet.colCount - 1;
  const fw = layout.frozenWidthPx;
  const sx0 = headerW + fw;

  const collectFrozenCols = (): number[] => {
    const out: number[] = [];
    for (let c = 0; c < frozenCols && c <= maxC; c++) {
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const colW = scaledColWidthAt(sheet, c, env.viewZoom);
      if (x + colW <= headerW || x >= canvasW) {
        continue;
      }
      out.push(c);
    }
    return out;
  };

  const collectScrollCols = (): number[] => {
    const out: number[] = [];
    if (frozenCols > maxC || layout.scrollViewportW <= 0) {
      return out;
    }
    const colW2 = scaledColW(sheet, env.viewZoom);
    const first = frozenCols + Math.floor(env.scrollX / colW2) - buf;
    const last = frozenCols + Math.ceil((env.scrollX + layout.scrollViewportW) / colW2) - 1 + buf;
    const c0 = Math.max(frozenCols, first);
    const c1 = Math.min(maxC, last);
    for (let c = c0; c <= c1; c++) {
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const colW = scaledColWidthAt(sheet, c, env.viewZoom);
      if (x + colW <= headerW || x >= canvasW) {
        continue;
      }
      out.push(c);
    }
    return out;
  };

  const strokeColumnHeaderGrid = (visibleCols: number[], clipLeft: number): void => {
    if (visibleCols.length === 0) {
      return;
    }
    const yTop = snapLine(0);
    const yBot = snapLine(headerH);
    const xClip = snapLine(clipLeft);
    ctx.strokeStyle = env.theme.headerLineColor;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(xClip, yTop);
    ctx.lineTo(xClip, yBot);
    ctx.moveTo(xClip, yTop);
    ctx.lineTo(snapLine(canvasW), yTop);
    for (const c of visibleCols) {
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const xs = snapLine(x);
      ctx.moveTo(xs, yTop);
      ctx.lineTo(xs, yBot);
    }
    const lastC = visibleCols[visibleCols.length - 1];
    const lastW = scaledColWidthAt(sheet, lastC, env.viewZoom);
    const xe = snapLine(cellLeftX(sheet, layout, lastC, env.viewZoom, env.scrollX) + lastW);
    ctx.moveTo(xe, yTop);
    ctx.lineTo(xe, yBot);
    ctx.stroke();
  };

  const scrollCols = collectScrollCols();
  if (scrollCols.length > 0 && frozenCols <= maxC && layout.scrollViewportW > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(sx0, 0, canvasW - sx0, headerH);
    ctx.clip();
    for (const c of scrollCols) {
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const colW = scaledColWidthAt(sheet, c, env.viewZoom);
      paintColumnHeaderCell(ctx, env.theme, env.viewZoom, c, x, colW, headerH, columnInSelection(c));
    }
    strokeColumnHeaderGrid(scrollCols, sx0);
    ctx.restore();
  }

  const frozenColsList = collectFrozenCols();
  if (frozenColsList.length > 0 && frozenCols > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(headerW, 0, fw, headerH);
    ctx.clip();
    for (const c of frozenColsList) {
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const colW = scaledColWidthAt(sheet, c, env.viewZoom);
      paintColumnHeaderCell(ctx, env.theme, env.viewZoom, c, x, colW, headerH, columnInSelection(c));
    }
    strokeColumnHeaderGrid(frozenColsList, headerW);
    ctx.restore();
  }
}

export function drawAllRowHeaders(
  env: HeaderPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasH: number,
  selectionSnap: SelectionPaintSnapshot | null,
): void {
  const { ctx } = env;
  const { frozenRows } = layout;
  const buf = env.viewportBuffer;
  const maxR = sheet.rowCount - 1;
  const fh = layout.frozenHeightPx;
  const sy0 = headerH + fh;
  const selSpanRows = getClampedSelectionSpan(sheet, selectionSnap);
  const rowInSelection = (r: number): boolean =>
    selSpanRows !== null && r >= selSpanRows.startRow && r <= selSpanRows.endRow;

  const collectFrozenRows = (): number[] => {
    const out: number[] = [];
    for (let r = 0; r < frozenRows && r <= maxR; r++) {
      const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
      if (y + rowH <= headerH || y >= canvasH) {
        continue;
      }
      out.push(r);
    }
    return out;
  };

  const collectScrollRows = (): number[] => {
    const out: number[] = [];
    if (frozenRows > maxR || layout.scrollViewportH <= 0) {
      return out;
    }
    const rowH2 = scaledRowH(sheet, env.viewZoom);
    const first = frozenRows + Math.floor(env.scrollY / rowH2) - buf;
    const last = frozenRows + Math.ceil((env.scrollY + layout.scrollViewportH) / rowH2) - 1 + buf;
    const r0 = Math.max(frozenRows, first);
    const r1 = Math.min(maxR, last);
    for (let r = r0; r <= r1; r++) {
      const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
      if (y + rowH <= headerH || y >= canvasH) {
        continue;
      }
      out.push(r);
    }
    return out;
  };

  const strokeRowHeaderGrid = (visibleRows: number[], clipTop: number): void => {
    if (visibleRows.length === 0) {
      return;
    }
    const xLeft = snapLine(0);
    const xRight = snapLine(headerW);
    const yClip = snapLine(clipTop);
    const yEnd = snapLine(canvasH);
    ctx.strokeStyle = env.theme.headerLineColor;
    ctx.lineWidth = 1;
    ctx.beginPath();
    if (clipTop !== headerH) {
      ctx.moveTo(xLeft, yClip);
      ctx.lineTo(xRight, yClip);
    }
    ctx.moveTo(xLeft, yClip);
    ctx.lineTo(xLeft, yEnd);
    for (const r of visibleRows) {
      const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      const ys = snapLine(y);
      ctx.moveTo(xLeft, ys);
      ctx.lineTo(xRight, ys);
    }
    const lastR = visibleRows[visibleRows.length - 1];
    const lastH = scaledRowHeightAt(sheet, lastR, env.viewZoom);
    const ye = snapLine(cellTopY(sheet, layout, lastR, env.viewZoom, env.scrollY) + lastH);
    ctx.moveTo(xLeft, ye);
    ctx.lineTo(xRight, ye);
    ctx.stroke();
  };

  const scrollRows = collectScrollRows();
  if (scrollRows.length > 0 && frozenRows <= maxR && layout.scrollViewportH > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, sy0, headerW, canvasH - sy0);
    ctx.clip();
    for (const r of scrollRows) {
      const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
      paintRowHeaderCell(ctx, env.theme, env.viewZoom, r, y, rowH, headerW, rowInSelection(r));
    }
    strokeRowHeaderGrid(scrollRows, sy0);
    ctx.restore();
  }

  const frozenRowsList = collectFrozenRows();
  if (frozenRowsList.length > 0 && frozenRows > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, headerH, headerW, fh);
    ctx.clip();
    for (const r of frozenRowsList) {
      const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
      paintRowHeaderCell(ctx, env.theme, env.viewZoom, r, y, rowH, headerW, rowInSelection(r));
    }
    strokeRowHeaderGrid(frozenRowsList, headerH);
    ctx.restore();
  }
}

export function paintAllHeaderSelectionAccents(
  env: HeaderPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  selectionSnap: SelectionPaintSnapshot | null,
): void {
  if (selectionSnap === null) {
    return;
  }
  const selSpan = getClampedSelectionSpan(sheet, selectionSnap);
  if (selSpan === null) {
    return;
  }
  const { ctx } = env;
  const t = viewScaledSelectionOutlineWidth(env.viewZoom);
  const omitColBottom = selSpan.startRow === 0;
  const omitRowRight = selSpan.startCol === 0;

  ctx.save();
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";

  if (!omitColBottom) {
    ctx.beginPath();
    ctx.rect(headerW, 0, canvasW - headerW, headerH);
    ctx.clip();
    ctx.fillStyle = env.theme.activeCellBorderColor;
    for (let c = selSpan.startCol; c <= selSpan.endCol; c++) {
      if (c < 0 || c >= sheet.colCount) {
        continue;
      }
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const colW = scaledColWidthAt(sheet, c, env.viewZoom);
      if (x + colW <= headerW || x >= canvasW) {
        continue;
      }
      ctx.fillRect(x, headerH - t, colW, t);
    }
  }

  ctx.restore();

  ctx.save();
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";

  if (!omitRowRight) {
    ctx.beginPath();
    ctx.rect(0, headerH, headerW, canvasH - headerH);
    ctx.clip();
    ctx.fillStyle = env.theme.activeCellBorderColor;
    for (let r = selSpan.startRow; r <= selSpan.endRow; r++) {
      if (r < 0 || r >= sheet.rowCount) {
        continue;
      }
      const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
      if (y + rowH <= headerH || y >= canvasH) {
        continue;
      }
      ctx.fillRect(headerW - t, y, t, rowH);
    }
  }

  ctx.restore();
}

export function drawHeadingBodyDividerLines(
  ctx: CanvasRenderingContext2D,
  theme: SheetTheme,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
): void {
  if (headerW <= 0 || headerH <= 0) {
    return;
  }
  const x0 = snapLine(0);
  const xHw = snapLine(headerW);
  const y0 = snapLine(headerH);
  const x1 = snapLine(canvasW);
  const y1 = snapLine(canvasH);
  ctx.save();
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";
  ctx.strokeStyle = theme.gridLineColor;
  ctx.lineWidth = 1;
  ctx.lineCap = "butt";
  ctx.lineJoin = "miter";
  ctx.beginPath();
  ctx.moveTo(x0, y0);
  ctx.lineTo(x1, y0);
  ctx.moveTo(xHw, y0);
  ctx.lineTo(xHw, y1);
  ctx.stroke();
  ctx.restore();
}

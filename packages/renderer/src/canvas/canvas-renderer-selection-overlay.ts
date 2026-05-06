import {
  normalizeSelectionRange,
  selectionRangesEqualNormalized,
  type SelectionPaintSnapshot,
} from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import { expandSelectionRangeForMergePaint } from "./canvas-renderer-selection-span.js";
import type { SheetTheme } from "@flexsheet/theme";
import { SELECTION_OUTLINE_VISUAL_SCALE } from "./canvas-renderer-constants.js";
import { cellIntersectsCanvas, cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import {
  scaledColWidthAt,
  scaledRowHeightAt,
  viewScaledSelectionOutlineWidth,
} from "./canvas-renderer-utils.js";

/** 选区描边与整像素格对齐，避免非整数线宽/坐标导致的发糊、虚影。 */
function strokeSnappedSelectionRect(
  ctx: CanvasRenderingContext2D,
  x0: number,
  y0: number,
  bw: number,
  bh: number,
  lineW: number,
): void {
  const leftI = Math.round(x0);
  const topI = Math.round(y0);
  const rightI = Math.round(x0 + bw);
  const bottomI = Math.round(y0 + bh);
  const rw = Math.max(0, rightI - leftI);
  const rh = Math.max(0, bottomI - topI);
  const inset = lineW / 2;
  ctx.lineWidth = lineW;
  ctx.strokeRect(
    leftI + inset,
    topI + inset,
    Math.max(0, rw - lineW),
    Math.max(0, rh - lineW),
  );
}
import { collectFrozenBodyQuadrantPasses } from "../layout/frozen-body-quadrants.js";
import type { FrozenLayout } from "../layout/viewport.js";

export interface SelectionOverlayEnv {
  readonly ctx: CanvasRenderingContext2D;
  readonly theme: SheetTheme;
  readonly viewZoom: number;
  readonly viewportBuffer: number;
  readonly scrollX: number;
  readonly scrollY: number;
}

export function drawSelectionOverlay(
  env: SelectionOverlayEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  snap: SelectionPaintSnapshot | null,
): void {
  if (snap === null) {
    return;
  }
  const { range: snapRange, activeRow, activeCol } = snap;
  const range = expandSelectionRangeForMergePaint(sheet, snapRange);
  const activeAnchor = sheet.getMergeAnchorCell(activeRow, activeCol);
  const { ctx } = env;
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";

  const passes = collectFrozenBodyQuadrantPasses(
    sheet,
    layout,
    headerW,
    headerH,
    canvasW,
    canvasH,
    env.scrollX,
    env.scrollY,
    env.viewZoom,
    env.viewportBuffer,
  );

  const drawFillPass = (
    clipX: number,
    clipY: number,
    clipW: number,
    clipH: number,
    pr0: number,
    pr1: number,
    pc0: number,
    pc1: number,
  ): void => {
    const r0 = Math.max(pr0, range.startRow);
    const r1 = Math.min(pr1, range.endRow);
    const c0 = Math.max(pc0, range.startCol);
    const c1 = Math.min(pc1, range.endCol);
    if (r0 > r1 || c0 > c1) {
      return;
    }
    ctx.save();
    ctx.beginPath();
    ctx.rect(clipX, clipY, clipW, clipH);
    ctx.clip();
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const here = sheet.getMergeAnchorCell(r, c);
        if (here.row === activeAnchor.row && here.col === activeAnchor.col) {
          continue;
        }
        const colW = scaledColWidthAt(sheet, c, env.viewZoom);
        const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
        if (colW <= 0 || rowH <= 0) {
          continue;
        }
        const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
        const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
        if (!cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
          continue;
        }
        ctx.fillStyle = env.theme.selectionFillColor;
        ctx.fillRect(x, y, colW, rowH);
      }
    }
    ctx.restore();
  };

  for (const p of passes) {
    drawFillPass(p.clipX, p.clipY, p.clipW, p.clipH, p.r0, p.r1, p.c0, p.c1);
  }

  const drawBorderPass = (
    clipX: number,
    clipY: number,
    clipW: number,
    clipH: number,
    pr0: number,
    pr1: number,
    pc0: number,
    pc1: number,
  ): void => {
    const r0 = Math.max(pr0, range.startRow);
    const r1 = Math.min(pr1, range.endRow);
    const c0 = Math.max(pc0, range.startCol);
    const c1 = Math.min(pc1, range.endCol);
    if (r0 > r1 || c0 > c1) {
      return;
    }

    ctx.save();
    ctx.beginPath();
    ctx.rect(clipX, clipY, clipW, clipH);
    ctx.clip();
    ctx.shadowColor = "transparent";
    ctx.shadowBlur = 0;
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;

    const x0 = cellLeftX(sheet, layout, c0, env.viewZoom, env.scrollX);
    const y0 = cellTopY(sheet, layout, r0, env.viewZoom, env.scrollY);
    let bw = 0;
    for (let c = c0; c <= c1; c++) {
      bw += scaledColWidthAt(sheet, c, env.viewZoom);
    }
    let bh = 0;
    for (let r = r0; r <= r1; r++) {
      bh += scaledRowHeightAt(sheet, r, env.viewZoom);
    }
    ctx.strokeStyle = env.theme.selectionBorderColor;
    strokeSnappedSelectionRect(
      ctx,
      x0,
      y0,
      bw,
      bh,
      viewScaledSelectionOutlineWidth(env.viewZoom),
    );

    ctx.restore();
  };

  for (const p of passes) {
    drawBorderPass(p.clipX, p.clipY, p.clipW, p.clipH, p.r0, p.r1, p.c0, p.c1);
  }

  const formulaRefRaw = snap.formulaReferencePreviewRange ?? null;
  const formulaRef =
    formulaRefRaw === null
      ? null
      : expandSelectionRangeForMergePaint(sheet, normalizeSelectionRange(formulaRefRaw));

  if (formulaRef !== null && !selectionRangesEqualNormalized(formulaRef, range)) {
    const frFillPass = (
      clipX: number,
      clipY: number,
      clipW: number,
      clipH: number,
      pr0: number,
      pr1: number,
      pc0: number,
      pc1: number,
    ): void => {
      const r0 = Math.max(pr0, formulaRef.startRow);
      const r1 = Math.min(pr1, formulaRef.endRow);
      const c0 = Math.max(pc0, formulaRef.startCol);
      const c1 = Math.min(pc1, formulaRef.endCol);
      if (r0 > r1 || c0 > c1) {
        return;
      }
      const anchorR = activeAnchor.row;
      const anchorC = activeAnchor.col;
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipW, clipH);
      ctx.clip();
      ctx.globalAlpha = 1;
      for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
          const here = sheet.getMergeAnchorCell(r, c);
          if (here.row === anchorR && here.col === anchorC) {
            continue;
          }
          const colW = scaledColWidthAt(sheet, c, env.viewZoom);
          const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
          if (colW <= 0 || rowH <= 0) {
            continue;
          }
          const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
          const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
          if (!cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
            continue;
          }
          ctx.fillStyle = "rgba(0, 120, 215, 0.12)";
          ctx.fillRect(x, y, colW, rowH);
        }
      }
      ctx.restore();
    };
    for (const p of passes) {
      frFillPass(p.clipX, p.clipY, p.clipW, p.clipH, p.r0, p.r1, p.c0, p.c1);
    }
    const frDashPass = (
      clipX: number,
      clipY: number,
      clipW: number,
      clipH: number,
      pr0: number,
      pr1: number,
      pc0: number,
      pc1: number,
    ): void => {
      const r0 = Math.max(pr0, formulaRef.startRow);
      const r1 = Math.min(pr1, formulaRef.endRow);
      const c0 = Math.max(pc0, formulaRef.startCol);
      const c1 = Math.min(pc1, formulaRef.endCol);
      if (r0 > r1 || c0 > c1) {
        return;
      }
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipW, clipH);
      ctx.clip();
      const x0 = cellLeftX(sheet, layout, c0, env.viewZoom, env.scrollX);
      const y0 = cellTopY(sheet, layout, r0, env.viewZoom, env.scrollY);
      let bw = 0;
      for (let c = c0; c <= c1; c++) {
        bw += scaledColWidthAt(sheet, c, env.viewZoom);
      }
      let bh = 0;
      for (let r = r0; r <= r1; r++) {
        bh += scaledRowHeightAt(sheet, r, env.viewZoom);
      }
      ctx.strokeStyle = "#0078d7";
      const frLw = Math.max(1, Math.round(1.5 * env.viewZoom));
      ctx.setLineDash([4 * Math.max(1, env.viewZoom * 0.5), 3 * Math.max(1, env.viewZoom * 0.5)]);
      strokeSnappedSelectionRect(ctx, x0, y0, bw, bh, frLw);
      ctx.setLineDash([]);
      ctx.restore();
    };
    for (const p of passes) {
      frDashPass(p.clipX, p.clipY, p.clipW, p.clipH, p.r0, p.r1, p.c0, p.c1);
    }
  }

  const fillPreviewRaw = snap.fillPreviewRange ?? null;
  const fillPreview =
    fillPreviewRaw === null
      ? null
      : expandSelectionRangeForMergePaint(sheet, normalizeSelectionRange(fillPreviewRaw));
  if (fillPreview !== null && !selectionRangesEqualNormalized(fillPreview, range)) {
    const drawDashedBorderPass = (
      clipX: number,
      clipY: number,
      clipW: number,
      clipH: number,
      pr0: number,
      pr1: number,
      pc0: number,
      pc1: number,
    ): void => {
      const r0 = Math.max(pr0, fillPreview.startRow);
      const r1 = Math.min(pr1, fillPreview.endRow);
      const c0 = Math.max(pc0, fillPreview.startCol);
      const c1 = Math.min(pc1, fillPreview.endCol);
      if (r0 > r1 || c0 > c1) {
        return;
      }
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipW, clipH);
      ctx.clip();
      const x0 = cellLeftX(sheet, layout, c0, env.viewZoom, env.scrollX);
      const y0 = cellTopY(sheet, layout, r0, env.viewZoom, env.scrollY);
      let bw = 0;
      for (let c = c0; c <= c1; c++) {
        bw += scaledColWidthAt(sheet, c, env.viewZoom);
      }
      let bh = 0;
      for (let r = r0; r <= r1; r++) {
        bh += scaledRowHeightAt(sheet, r, env.viewZoom);
      }
      ctx.strokeStyle = env.theme.selectionBorderColor;
      ctx.setLineDash([4, 3]);
      strokeSnappedSelectionRect(
        ctx,
        x0,
        y0,
        bw,
        bh,
        viewScaledSelectionOutlineWidth(env.viewZoom),
      );
      ctx.setLineDash([]);
      ctx.restore();
    };
    for (const p of passes) {
      drawDashedBorderPass(p.clipX, p.clipY, p.clipW, p.clipH, p.r0, p.r1, p.c0, p.c1);
    }
  }

  const handleRange = fillPreview !== null && !selectionRangesEqualNormalized(fillPreview, range)
    ? fillPreview
    : range;
  const handleCenterX =
    cellLeftX(sheet, layout, handleRange.endCol, env.viewZoom, env.scrollX) +
    scaledColWidthAt(sheet, handleRange.endCol, env.viewZoom);
  const handleCenterY =
    cellTopY(sheet, layout, handleRange.endRow, env.viewZoom, env.scrollY) +
    scaledRowHeightAt(sheet, handleRange.endRow, env.viewZoom);
  const handleSize = Math.max(4, 6 * SELECTION_OUTLINE_VISUAL_SCALE * env.viewZoom);
  const handleHalf = handleSize / 2;
  const bodyX = headerW;
  const bodyY = headerH;
  const bodyW = Math.max(0, canvasW - headerW);
  const bodyH = Math.max(0, canvasH - headerH);
  if (bodyW > 0 && bodyH > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(bodyX, bodyY, bodyW, bodyH);
    ctx.clip();
    ctx.shadowColor = "transparent";
    ctx.shadowBlur = 0;
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;
    ctx.fillStyle = env.theme.selectionBorderColor;
    ctx.fillRect(handleCenterX - handleHalf, handleCenterY - handleHalf, handleSize, handleSize);
    ctx.restore();
  }
}

import type { SelectionPaintSnapshot } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import type { SheetTheme } from "@flexsheet/theme";
import { SELECTION_OUTLINE_VISUAL_SCALE } from "./canvas-renderer-constants.js";
import { cellIntersectsCanvas, cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import {
  scaledColWidthAt,
  scaledRowHeightAt,
  viewScaledSelectionOutlineWidth,
} from "./canvas-renderer-utils.js";
import { collectFrozenBodyQuadrantPasses } from "./frozen-body-quadrants.js";
import type { FrozenLayout } from "./viewport.js";

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
  const { range, activeRow, activeCol } = snap;
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
        if (r === activeRow && c === activeCol) {
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
    ctx.lineWidth = viewScaledSelectionOutlineWidth(env.viewZoom);
    ctx.strokeRect(x0 + 0.5, y0 + 0.5, bw - 1, bh - 1);

    ctx.restore();
  };

  for (const p of passes) {
    drawBorderPass(p.clipX, p.clipY, p.clipW, p.clipH, p.r0, p.r1, p.c0, p.c1);
  }

  const handleCenterX =
    cellLeftX(sheet, layout, range.endCol, env.viewZoom, env.scrollX) +
    scaledColWidthAt(sheet, range.endCol, env.viewZoom);
  const handleCenterY =
    cellTopY(sheet, layout, range.endRow, env.viewZoom, env.scrollY) +
    scaledRowHeightAt(sheet, range.endRow, env.viewZoom);
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

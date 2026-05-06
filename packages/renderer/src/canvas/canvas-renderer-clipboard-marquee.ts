import type { SelectionRange } from "@flexsheet/core";
import { normalizeSelectionRange } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import { cellIntersectsCanvas, cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import { scaledColWidthAt, scaledRowHeightAt } from "./canvas-renderer-utils.js";
import { collectFrozenBodyQuadrantPasses } from "../layout/frozen-body-quadrants.js";
import type { FrozenLayout } from "../layout/viewport.js";
import type { SelectionOverlayEnv } from "./canvas-renderer-selection-overlay.js";

/**
 * 复制/剪切后的「走马灯」虚线框（黑白双色错位 stroke，风格接近 Excel），
 * 绘于选区覆盖层之上，不改变选区填充与边框逻辑。
 */
export function drawClipboardMarqueeOverlay(
  env: SelectionOverlayEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  range: SelectionRange,
  /** 用于 `lineDashOffset` 的动画相位（像素，单调递增即可）。 */
  phasePx: number,
): void {
  const n = normalizeSelectionRange(range);
  if (n.startRow > n.endRow || n.startCol > n.endCol) {
    return;
  }
  const { ctx } = env;
  ctx.save();
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";
  ctx.shadowColor = "transparent";
  ctx.shadowBlur = 0;
  ctx.lineJoin = "miter";
  ctx.lineCap = "square";

  const dash = Math.max(3, 3 * env.viewZoom);
  const period = dash * 2;
  const offset = ((phasePx % period) + period) % period;
  const lineW = Math.max(1, 1 * env.viewZoom);

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

  for (const p of passes) {
    const r0 = Math.max(p.r0, n.startRow);
    const r1 = Math.min(p.r1, n.endRow);
    const c0 = Math.max(p.c0, n.startCol);
    const c1 = Math.min(p.c1, n.endCol);
    if (r0 > r1 || c0 > c1) {
      continue;
    }

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
    if (!cellIntersectsCanvas(x0, y0, bw, bh, headerW, headerH, canvasW, canvasH)) {
      continue;
    }

    ctx.save();
    ctx.beginPath();
    ctx.rect(p.clipX, p.clipY, p.clipW, p.clipH);
    ctx.clip();

    const rx = x0 + 0.5;
    const ry = y0 + 0.5;
    const rw = Math.max(0, bw - 1);
    const rh = Math.max(0, bh - 1);

    ctx.lineWidth = lineW;
    ctx.setLineDash([dash, dash]);

    ctx.strokeStyle = "#ffffff";
    ctx.lineDashOffset = offset;
    ctx.strokeRect(rx, ry, rw, rh);

    ctx.strokeStyle = "#000000";
    ctx.lineDashOffset = offset + dash;
    ctx.strokeRect(rx, ry, rw, rh);

    ctx.restore();
  }

  ctx.restore();
}

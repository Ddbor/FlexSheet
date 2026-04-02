import type { SheetTheme } from "@flexsheet/theme";
import { snapLine } from "./canvas-renderer-utils.js";
import type { FrozenLayout } from "./viewport.js";

/**
 * 冻结分隔线：只画在表体区域 [headerW,canvasW)×[headerH,canvasH)，不贯穿行列标题。
 */
export function drawFreezeLines(
  ctx: CanvasRenderingContext2D,
  theme: SheetTheme,
  layout: FrozenLayout,
  canvasW: number,
  canvasH: number,
): void {
  const { headerW, headerH, frozenCols, frozenRows, frozenWidthPx, frozenHeightPx } = layout;
  if (frozenCols === 0 && frozenRows === 0) {
    return;
  }
  ctx.save();
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";
  ctx.strokeStyle = theme.freezeLineColor;
  ctx.lineWidth = 1;
  const xBody0 = snapLine(headerW);
  const yBody0 = snapLine(headerH);
  const xBody1 = snapLine(canvasW);
  const yBody1 = snapLine(canvasH);
  if (frozenCols > 0) {
    const x = headerW + frozenWidthPx;
    ctx.beginPath();
    ctx.moveTo(snapLine(x), yBody0);
    ctx.lineTo(snapLine(x), yBody1);
    ctx.stroke();
  }
  if (frozenRows > 0) {
    const y = headerH + frozenHeightPx;
    ctx.beginPath();
    ctx.moveTo(xBody0, snapLine(y));
    ctx.lineTo(xBody1, snapLine(y));
    ctx.stroke();
  }
  ctx.restore();
}

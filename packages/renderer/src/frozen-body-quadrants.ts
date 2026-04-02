import type { Worksheet } from "@flexsheet/core";
import type { FrozenLayout } from "./viewport.js";
import { visibleScrollableCellRange } from "./viewport.js";

/** 冻结表体分象限绘制/裁剪参数（滚动区 + 三条冻结带 + 角区）。 */
export interface BodyQuadrantPass {
  readonly clipX: number;
  readonly clipY: number;
  readonly clipW: number;
  readonly clipH: number;
  readonly r0: number;
  readonly r1: number;
  readonly c0: number;
  readonly c1: number;
}

/**
 * 与 `drawBody` / 选区覆盖层相同的四象限遍历顺序与行列区间，避免两处逻辑漂移。
 */
export function collectFrozenBodyQuadrantPasses(
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  scrollX: number,
  scrollY: number,
  viewZoom: number,
  buf: number,
): BodyQuadrantPass[] {
  const passes: BodyQuadrantPass[] = [];
  const colW = sheet.defaultColWidth * viewZoom;
  const fw = layout.frozenWidthPx;
  const fh = layout.frozenHeightPx;
  const sx0 = headerW + fw;
  const sy0 = headerH + fh;
  const fr = layout.frozenRows;
  const fc = layout.frozenCols;
  const vr = visibleScrollableCellRange(sheet, layout, scrollX, scrollY, buf, viewZoom);

  if (fr < sheet.rowCount && fc < sheet.colCount && vr.startRow <= vr.endRow && vr.startCol <= vr.endCol) {
    passes.push({
      clipX: sx0,
      clipY: sy0,
      clipW: canvasW - sx0,
      clipH: canvasH - sy0,
      r0: vr.startRow,
      r1: vr.endRow,
      c0: vr.startCol,
      c1: vr.endCol,
    });
  }

  if (fr > 0 && fc < sheet.colCount && layout.scrollViewportW > 0) {
    const maxC = sheet.colCount - 1;
    const firstCol = fc + Math.floor(scrollX / colW) - buf;
    const lastCol = fc + Math.ceil((scrollX + layout.scrollViewportW) / colW) - 1 + buf;
    const c0 = Math.max(fc, firstCol);
    const c1 = Math.min(maxC, lastCol);
    const r0 = 0;
    const r1 = Math.min(fr - 1, sheet.rowCount - 1);
    if (c0 <= c1 && r0 <= r1) {
      passes.push({
        clipX: sx0,
        clipY: headerH,
        clipW: canvasW - sx0,
        clipH: fh,
        r0,
        r1,
        c0,
        c1,
      });
    }
  }

  if (fc > 0 && fr < sheet.rowCount && vr.startRow <= vr.endRow && layout.scrollViewportH > 0) {
    const c0 = 0;
    const c1 = Math.min(fc - 1, sheet.colCount - 1);
    passes.push({
      clipX: headerW,
      clipY: sy0,
      clipW: fw,
      clipH: canvasH - sy0,
      r0: vr.startRow,
      r1: vr.endRow,
      c0,
      c1,
    });
  }

  if (fr > 0 && fc > 0) {
    const r0 = 0;
    const r1 = Math.min(fr - 1, sheet.rowCount - 1);
    const c0 = 0;
    const c1 = Math.min(fc - 1, sheet.colCount - 1);
    if (r0 <= r1 && c0 <= c1) {
      passes.push({
        clipX: headerW,
        clipY: headerH,
        clipW: fw,
        clipH: fh,
        r0,
        r1,
        c0,
        c1,
      });
    }
  }

  return passes;
}

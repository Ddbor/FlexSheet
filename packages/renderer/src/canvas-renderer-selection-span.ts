import {
  normalizeSelectionRange,
  type SelectionPaintSnapshot,
  type SelectionRange,
} from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";

/**
 * 将选区矩形按合并区域外扩（选区绘制、行列标题高亮与 `getClampedSelectionSpan` 共用）。
 */
export function expandSelectionRangeForMergePaint(sheet: Worksheet, range: SelectionRange): SelectionRange {
  const n = normalizeSelectionRange(range);
  let sr = n.startRow;
  let er = n.endRow;
  let sc = n.startCol;
  let ec = n.endCol;
  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      const info = sheet.getMergedRectInfo(r, c);
      sr = Math.min(sr, info.anchorRow);
      er = Math.max(er, info.anchorRow + info.rowSpan - 1);
      sc = Math.min(sc, info.anchorCol);
      ec = Math.max(ec, info.anchorCol + info.colSpan - 1);
    }
  }
  return normalizeSelectionRange({ startRow: sr, endRow: er, startCol: sc, endCol: ec });
}

/** 当前选区在表内的行列闭区间（与框选矩形一致）；无快照时返回 null。 */
export function getClampedSelectionSpan(
  sheet: Worksheet,
  selectionSnap: SelectionPaintSnapshot | null,
): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
  if (selectionSnap === null) {
    return null;
  }
  const n = normalizeSelectionRange(expandSelectionRangeForMergePaint(sheet, selectionSnap.range));
  const maxC = sheet.colCount - 1;
  const maxR = sheet.rowCount - 1;
  if (maxC < 0 || maxR < 0) {
    return null;
  }
  return {
    startCol: Math.max(0, Math.min(n.startCol, maxC)),
    endCol: Math.max(0, Math.min(n.endCol, maxC)),
    startRow: Math.max(0, Math.min(n.startRow, maxR)),
    endRow: Math.max(0, Math.min(n.endRow, maxR)),
  };
}

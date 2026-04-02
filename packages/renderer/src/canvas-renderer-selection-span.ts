import { normalizeSelectionRange, type SelectionPaintSnapshot } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";

/** 当前选区在表内的行列闭区间（与框选矩形一致）；无快照时返回 null。 */
export function getClampedSelectionSpan(
  sheet: Worksheet,
  selectionSnap: SelectionPaintSnapshot | null,
): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
  if (selectionSnap === null) {
    return null;
  }
  const n = normalizeSelectionRange(selectionSnap.range);
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

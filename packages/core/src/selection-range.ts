/**
 * 选区矩形（行列均为包含端点的索引）。
 * 置于 core 供 Render / Interaction 共用，避免 Render 依赖 Interaction。
 */

export interface SelectionRange {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

/** 供 Render 绘制的选区快照（只读）。 */
export interface SelectionPaintSnapshot {
  readonly range: SelectionRange;
  readonly activeRow: number;
  readonly activeCol: number;
  /**
   * 拖拽填充柄时，当前预览的外包矩形（与 `range` 取并集后规范化；无预览为 `undefined`）。
   */
  readonly fillPreviewRange?: SelectionRange | null;
}

export function normalizeSelectionRange(range: SelectionRange): SelectionRange {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startCol: Math.min(range.startCol, range.endCol),
    endCol: Math.max(range.startCol, range.endCol),
  };
}

export function selectionRangeContains(range: SelectionRange, row: number, col: number): boolean {
  const n = normalizeSelectionRange(range);
  return row >= n.startRow && row <= n.endRow && col >= n.startCol && col <= n.endCol;
}

export function selectionRangesIntersect(a: SelectionRange, b: SelectionRange): boolean {
  const na = normalizeSelectionRange(a);
  const nb = normalizeSelectionRange(b);
  return !(na.endRow < nb.startRow || na.startRow > nb.endRow || na.endCol < nb.startCol || na.startCol > nb.endCol);
}

/** 两选区规范化后是否为同一矩形（用于条件格式等同选区覆盖等）。 */
export function selectionRangesEqualNormalized(a: SelectionRange, b: SelectionRange): boolean {
  const na = normalizeSelectionRange(a);
  const nb = normalizeSelectionRange(b);
  return (
    na.startRow === nb.startRow &&
    na.endRow === nb.endRow &&
    na.startCol === nb.startCol &&
    na.endCol === nb.endCol
  );
}

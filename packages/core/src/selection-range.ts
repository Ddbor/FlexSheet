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

import type { ConditionalFormattingOverlay, Worksheet } from "@flexsheet/core";

/**
 * 单次象限绘制内对同一格只解析一次条件格式叠加（填底与文字共用），避免重复调用
 * `resolveConditionalFormattingCellOverlay`。
 */
export function getConditionalFormattingCellOverlayCached(
  sheet: Worksheet,
  row: number,
  col: number,
  cache: Map<string, ConditionalFormattingOverlay | null> | undefined,
): ConditionalFormattingOverlay | null {
  if (cache === undefined) {
    return sheet.resolveConditionalFormattingCellOverlay(row, col);
  }
  const key = `${row},${col}`;
  if (!cache.has(key)) {
    cache.set(key, sheet.resolveConditionalFormattingCellOverlay(row, col));
  }
  return cache.get(key) ?? null;
}

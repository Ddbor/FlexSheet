/**
 * 单元格填充图案（与 OOXML `patternFill/@patternType` 及 Excel「图案样式」子集一致）。
 * `none` 表示无图案，仅由 `fillArgb` 决定纯色底纹。
 * 注：数组顺序为类型全集；对话框 3×6 网格展示顺序由 flexsheet 层 `FORMAT_CELLS_PATTERN_GRID_ORDER` 单独定义。
 */

export const CELL_FILL_PATTERN_TYPES = [
  "none",
  "gray125",
  "gray0625",
  "darkGray",
  "mediumGray",
  "lightGray",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
] as const;

export type CellFillPatternType = (typeof CELL_FILL_PATTERN_TYPES)[number];

export function isCellFillPatternType(v: string): v is CellFillPatternType {
  return (CELL_FILL_PATTERN_TYPES as readonly string[]).includes(v);
}

import type { SelectionRange } from "./selection-range.js";

/** 与 FlexSheet 数据透视对话框及 OOXML 汇总方式对齐。 */
export type PivotAggregateKind = "sum" | "count" | "average" | "max" | "min";

/** 值区域中单个度量：源列 + 汇总方式。 */
export interface PivotValueFieldSpec {
  readonly col: number;
  readonly aggregate: PivotAggregateKind;
}

/**
 * 工作表上的「原生透视」导出元数据：由创建/更新透视命令写入，供 XLSX 导出生成
 * `pivotCacheDefinition` / `pivotTableDefinition` 等部件；不在渲染层解析。
 */
export interface WorksheetPivotTableDefinition {
  /** 稳定 id（撤销/重做与删除用）。 */
  readonly id: string;
  /** Excel 透视表名称（工作簿内建议唯一）。 */
  readonly name: string;
  /** 数据源所在工作表索引。 */
  readonly sourceSheetIndex: number;
  readonly sourceRange: SelectionRange;
  readonly hasHeaders: boolean;
  /** 数据源中作为「行」的列号（绝对列索引），顺序与输出嵌套一致。 */
  readonly rowFieldCols: readonly number[];
  /** 数据源中作为「列」的列号；空数组表示无列字段。 */
  readonly columnFieldCols: readonly number[];
  /** 筛选器区域字段列（当前主要供 UI/导出占位；刷新时可选扩展为行过滤）。 */
  readonly filterFieldCols: readonly number[];
  /** 值区域：每项对应输出中的一列或一块数据列。 */
  readonly valueFields: readonly PivotValueFieldSpec[];
  /** 透视表显示区域左上角（与 FlexSheet 写入预览格一致）。 */
  readonly destinationRow: number;
  readonly destinationCol: number;
  readonly outputRowCount: number;
  readonly outputColCount: number;
}

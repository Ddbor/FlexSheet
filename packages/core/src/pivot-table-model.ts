import type { SelectionRange } from "./selection-range.js";

/** 与 FlexSheet 数据透视对话框及 OOXML 汇总方式对齐。 */
export type PivotAggregateKind = "sum" | "count" | "average" | "max" | "min";

/**
 * 透视「值」区域的计算类型：在分组内先对分子/分母列分别求和再相除，
 * 避免对源表中预置的比率、占比、均价列做简单加总导致「总计」错误。
 */
export type PivotValueComputed =
  /** 同分组内 sum(本列)/sum(分母列)；总计为全表分子合计/分母合计（如点击率、平均点击花费）。 */
  | { readonly kind: "bucketRatio"; readonly denominatorCol: number }
  /** sum(本分组)/数据源中该列全体合计；总计为 1（如花费占总花费比例）。 */
  | { readonly kind: "shareOfGrandTotal" };

/** 值区域中单个度量：源列 + 汇总方式，可选计算语义。 */
export interface PivotValueFieldSpec {
  readonly col: number;
  readonly aggregate: PivotAggregateKind;
  readonly computed?: PivotValueComputed;
}

function pivotAggregateLabel(kind: PivotAggregateKind): string {
  switch (kind) {
    case "sum":
      return "求和";
    case "count":
      return "计数";
    case "average":
      return "平均值";
    case "max":
      return "最大值";
    case "min":
      return "最小值";
    default:
      return "汇总";
  }
}

/** 透视表值区域列标题（与 Excel「求和项:xxx」风格一致，含计算字段）。 */
export function getPivotValueFieldCaption(
  spec: PivotValueFieldSpec,
  fieldName: string,
  denominatorFieldName?: string,
): string {
  if (spec.computed?.kind === "shareOfGrandTotal") {
    return `占比项:${fieldName}`;
  }
  if (spec.computed?.kind === "bucketRatio") {
    const den =
      denominatorFieldName !== undefined && denominatorFieldName.length > 0
        ? denominatorFieldName
        : `列${spec.computed.denominatorCol + 1}`;
    return `比率项:${fieldName}/${den}`;
  }
  return `${pivotAggregateLabel(spec.aggregate)}项:${fieldName}`;
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
  /** 筛选器区域字段列（与 `filterSelectedKeys` 逐项对应）。 */
  readonly filterFieldCols: readonly number[];
  /**
   * 与 `filterFieldCols` 等长：每项为允许参与汇总的透视键（与透视行/列键规则一致）；
   * 空数组表示该字段不限制（全部）。
   */
  readonly filterSelectedKeys: readonly (readonly string[])[];
  /** 值区域：每项对应输出中的一列或一块数据列。 */
  readonly valueFields: readonly PivotValueFieldSpec[];
  /**
   * 多个值字段且无列字段时：与 Excel「数值」维度一致——`true` 表示度量在行方向展开（每个度量一行），
   * `false` 表示度量在列方向展开（每个度量一列）。单值字段或存在列字段时忽略。
   */
  readonly valueFieldsOnRows?: boolean;
  /** 透视表显示区域左上角（与 FlexSheet 写入预览格一致）。 */
  readonly destinationRow: number;
  readonly destinationCol: number;
  readonly outputRowCount: number;
  readonly outputColCount: number;
}

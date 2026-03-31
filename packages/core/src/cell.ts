/**
 * 单元格数据模型（Data 层最小单元），不含绘制与公式求值。
 *
 * **写入约定**：业务侧应通过 `Worksheet.setCellValue` / `setCellLiteral` / `setCellFormula` /
 * `setCellStyle` 修改内容与样式，以便触发「数据驱动视图」的变更通知。
 * 公式引擎在重算过程中会直接写入 `value`，由 `recalcWorksheet` 末尾统一 `notifyDataChanged`。
 */

export interface CellAddress {
  readonly row: number;
  readonly col: number;
}

export type CellScalar = string | number | boolean | null;

/** XLSX 往返用最小样式（ARGB 含 alpha，如 FFFF0000）。 */
export interface CellStyle {
  bold?: boolean;
  fgArgb?: string;
  fillArgb?: string;
}

export class Cell implements CellAddress {
  /**
   * 以 `=` 开头的公式原文；非公式单元格为 `null`，此时 `value` 为缓存的标量结果或字面量。
   */
  formula: string | null = null;

  /** 可选显示样式（渲染 / 导出 xlsx）。 */
  style: CellStyle | null = null;

  constructor(
    public readonly row: number,
    public readonly col: number,
    public value: CellScalar = null,
  ) {}

  static key(row: number, col: number): string {
    return `${row},${col}`;
  }

  /** 是否为公式单元格（`formula` 非空）。 */
  isFormulaCell(): boolean {
    return this.formula !== null && this.formula.length > 0;
  }
}

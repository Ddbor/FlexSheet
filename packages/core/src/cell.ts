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

/** 水平对齐（与 Ribbon / XLSX 常用子集一致；未设置视为左对齐）。 */
export type CellHorizontalAlign = "left" | "center" | "right";

/** 垂直对齐（未设置视为垂直居中，与默认绘制一致）。 */
export type CellVerticalAlign = "top" | "middle" | "bottom";

/**
 * 文本方向（与 Ribbon「方向」及 OOXML `alignment/@textRotation` 子集对应）。
 * 未设置视为 `horizontal`。
 */
export type CellTextOrientation =
  | "horizontal"
  /** 逆时针倾斜约 45°（与 Excel「逆时针角度」一致）。 */
  | "angleUp45"
  /** 顺时针倾斜约 45°。 */
  | "angleDown45"
  /** 竖排堆叠字符（与 Excel 竖排文字 / textRotation=255 一致）。 */
  | "verticalStack"
  /** 整行文字逆时针 90°（自下而上读）。 */
  | "rotateUp90"
  /** 整行文字顺时针 90°（自上而下读）。 */
  | "rotateDown90";

/** 单元格单边边框线型（与 Ribbon 边框预设子集一致）。 */
export type CellBorderKind = "thin" | "medium" | "thick" | "double" | "hairline";

/** 单边边框（颜色缺省为黑色，由渲染层解析）。 */
export interface CellBorderSide {
  readonly kind: CellBorderKind;
  /** 8 位 ARGB，如 FF000000；缺省为黑色。 */
  readonly colorArgb?: string;
}

/** XLSX 往返用最小样式（ARGB 含 alpha，如 FFFF0000）。 */
export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  /** CSS `font-family` 栈（与 Ribbon 字体列表一致）。 */
  fontFamily?: string;
  /** 字号，与 Excel 一致为磅（pt）。 */
  fontSizePt?: number;
  /** 下划线；未设置表示无下划线。 */
  underline?: "single" | "double";
  fgArgb?: string;
  fillArgb?: string;
  /** 文本水平对齐。 */
  hAlign?: CellHorizontalAlign;
  /** 文本垂直对齐。 */
  vAlign?: CellVerticalAlign;
  /**
   * 左缩进等级（与 OOXML `alignment/@indent` 一致，0–255；未设置视为 0）。
   * Ribbon 每点击一次增加/减少 1 级。
   */
  indentLevel?: number;
  /** 自动换行（单元格内按列宽折行）。 */
  wrapText?: boolean;
  /** 文本方向；未设置为水平。 */
  textOrientation?: CellTextOrientation;
  /** 单元格上/左/下/右边框（合并格仅存于主格）。 */
  borderTop?: CellBorderSide;
  borderLeft?: CellBorderSide;
  borderBottom?: CellBorderSide;
  borderRight?: CellBorderSide;
  /**
   * Excel 兼容的数字格式码（如 `0.00%`、`#,##0.00`、`yyyy/m/d`）。
   * 未设置或 `General` 表示常规。
   */
  numberFormat?: string;
}

/**
 * 选区样式补丁：`undefined` 表示不修改该项，`null` 表示清除该项（恢复默认）。
 */
export type CellStylePatch = {
  readonly bold?: boolean | null;
  readonly italic?: boolean | null;
  readonly fontFamily?: string | null;
  readonly fontSizePt?: number | null;
  readonly underline?: "single" | "double" | null;
  readonly fgArgb?: string | null;
  readonly fillArgb?: string | null;
  readonly hAlign?: CellHorizontalAlign | null;
  readonly vAlign?: CellVerticalAlign | null;
  readonly indentLevel?: number | null;
  readonly wrapText?: boolean | null;
  readonly textOrientation?: CellTextOrientation | null;
  readonly borderTop?: CellBorderSide | null;
  readonly borderLeft?: CellBorderSide | null;
  readonly borderBottom?: CellBorderSide | null;
  readonly borderRight?: CellBorderSide | null;
  readonly numberFormat?: string | null;
};

export function applyCellStylePatch(
  prev: CellStyle | null,
  patch: CellStylePatch,
): CellStyle | null {
  const next: CellStyle = { ...(prev ?? {}) };
  if (patch.bold !== undefined) {
    if (patch.bold === null) {
      delete next.bold;
    } else {
      next.bold = patch.bold;
    }
  }
  if (patch.italic !== undefined) {
    if (patch.italic === null) {
      delete next.italic;
    } else {
      next.italic = patch.italic;
    }
  }
  if (patch.fontFamily !== undefined) {
    if (patch.fontFamily === null) {
      delete next.fontFamily;
    } else {
      next.fontFamily = patch.fontFamily;
    }
  }
  if (patch.fontSizePt !== undefined) {
    if (patch.fontSizePt === null) {
      delete next.fontSizePt;
    } else {
      next.fontSizePt = patch.fontSizePt;
    }
  }
  if (patch.underline !== undefined) {
    if (patch.underline === null) {
      delete next.underline;
    } else {
      next.underline = patch.underline;
    }
  }
  if (patch.fgArgb !== undefined) {
    if (patch.fgArgb === null) {
      delete next.fgArgb;
    } else {
      next.fgArgb = patch.fgArgb;
    }
  }
  if (patch.fillArgb !== undefined) {
    if (patch.fillArgb === null) {
      delete next.fillArgb;
    } else {
      next.fillArgb = patch.fillArgb;
    }
  }
  if (patch.hAlign !== undefined) {
    if (patch.hAlign === null) {
      delete next.hAlign;
    } else {
      next.hAlign = patch.hAlign;
    }
  }
  if (patch.vAlign !== undefined) {
    if (patch.vAlign === null) {
      delete next.vAlign;
    } else {
      next.vAlign = patch.vAlign;
    }
  }
  if (patch.indentLevel !== undefined) {
    if (patch.indentLevel === null) {
      delete next.indentLevel;
    } else {
      const n = Math.round(patch.indentLevel);
      if (!Number.isFinite(n) || n <= 0) {
        delete next.indentLevel;
      } else {
        next.indentLevel = Math.min(255, n);
      }
    }
  }
  if (patch.wrapText !== undefined) {
    if (patch.wrapText === null) {
      delete next.wrapText;
    } else {
      next.wrapText = patch.wrapText;
    }
  }
  if (patch.textOrientation !== undefined) {
    if (patch.textOrientation === null) {
      delete next.textOrientation;
    } else {
      next.textOrientation = patch.textOrientation;
    }
  }
  if (patch.borderTop !== undefined) {
    if (patch.borderTop === null) {
      delete next.borderTop;
    } else {
      next.borderTop = patch.borderTop;
    }
  }
  if (patch.borderLeft !== undefined) {
    if (patch.borderLeft === null) {
      delete next.borderLeft;
    } else {
      next.borderLeft = patch.borderLeft;
    }
  }
  if (patch.borderBottom !== undefined) {
    if (patch.borderBottom === null) {
      delete next.borderBottom;
    } else {
      next.borderBottom = patch.borderBottom;
    }
  }
  if (patch.borderRight !== undefined) {
    if (patch.borderRight === null) {
      delete next.borderRight;
    } else {
      next.borderRight = patch.borderRight;
    }
  }
  if (patch.numberFormat !== undefined) {
    if (patch.numberFormat === null) {
      delete next.numberFormat;
    } else {
      const t = patch.numberFormat.trim();
      if (t === "" || t.toLowerCase() === "general") {
        delete next.numberFormat;
      } else {
        next.numberFormat = t;
      }
    }
  }
  return Object.keys(next).length === 0 ? null : next;
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

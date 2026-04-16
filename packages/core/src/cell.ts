import type { CellFillPatternType } from "./cell-fill-pattern.js";

export {
  CELL_FILL_PATTERN_TYPES,
  isCellFillPatternType,
  type CellFillPatternType,
} from "./cell-fill-pattern.js";

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

/** 水平对齐（与 Ribbon / OOXML `horizontal` 子集一致；未设置视为常规/左对齐）。 */
export type CellHorizontalAlign =
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "distributed"
  /** 跨列居中（OOXML `centerContinuous`）。单格内与居中表现一致。 */
  | "centerContinuous";

/** 垂直对齐（与 OOXML `vertical` 子集一致；未设置视为垂直居中）。 */
export type CellVerticalAlign = "top" | "middle" | "bottom" | "justify" | "distributed";

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

/**
 * 边框笔画图案（与「设置单元格格式」线型网格一致）；未设置时仅按 `kind` 画实线。
 */
export type CellBorderLinePattern =
  | "hairlineDots"
  | "shortDash"
  | "dashDot"
  | "dashDotDot"
  | "mediumDash"
  | "mediumDashDotDot"
  | "slantedDash"
  | "thickDash"
  | "thickDashDot"
  | "thinSolid"
  | "mediumSolid"
  | "thickSolid"
  | "doubleLine";

/** 单边边框（颜色缺省为黑色，由渲染层解析）。 */
export interface CellBorderSide {
  readonly kind: CellBorderKind;
  /** 8 位 ARGB，如 FF000000；缺省为黑色。 */
  readonly colorArgb?: string;
  /** 虚线/点线等图案；缺省时仅按 `kind` 绘制实线（兼容旧数据）。 */
  readonly linePattern?: CellBorderLinePattern;
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
  /** 删除线。 */
  strikethrough?: boolean;
  /**
   * 整格上标/下标（水平单行文本时生效；与竖排/任意角旋转等互斥时由渲染层忽略）。
   */
  fontScript?: "superscript" | "subscript";
  fgArgb?: string;
  fillArgb?: string;
  /**
   * 填充图案（OOXML `patternType`）；未设置或 `none` 表示无图案，仅 `fillArgb` 纯色。
   */
  fillPatternType?: CellFillPatternType;
  /**
   * 图案前景色（线/点），8 位 ARGB；未设置表示「自动」（渲染为近黑色，与 Excel 一致）。
   */
  fillPatternFgArgb?: string;
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
  /**
   * 任意旋转角（度），逆时针为正、顺时针为负（约 -90°～90°）。
   * 与 `textOrientation` 互斥：设置本字段时渲染优先按角度绘制，并忽略非 `horizontal` 的 `textOrientation`。
   */
  textRotationDegrees?: number;
  /** 缩小字体填充，使内容在单元格宽度内显示。 */
  shrinkToFit?: boolean;
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
  /**
   * 工作表受保护时是否锁定单元格（不可编辑）。未设置视为 `true`（与 Excel 默认一致）。
   */
  locked?: boolean;
  /**
   * 工作表受保护时是否在公式栏隐藏公式；仅对含公式单元格有意义。
   */
  formulaHidden?: boolean;
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
  readonly strikethrough?: boolean | null;
  readonly fontScript?: "superscript" | "subscript" | null;
  readonly fgArgb?: string | null;
  readonly fillArgb?: string | null;
  readonly fillPatternType?: CellFillPatternType | null;
  readonly fillPatternFgArgb?: string | null;
  readonly hAlign?: CellHorizontalAlign | null;
  readonly vAlign?: CellVerticalAlign | null;
  readonly indentLevel?: number | null;
  readonly wrapText?: boolean | null;
  readonly textRotationDegrees?: number | null;
  readonly shrinkToFit?: boolean | null;
  readonly textOrientation?: CellTextOrientation | null;
  readonly borderTop?: CellBorderSide | null;
  readonly borderLeft?: CellBorderSide | null;
  readonly borderBottom?: CellBorderSide | null;
  readonly borderRight?: CellBorderSide | null;
  readonly numberFormat?: string | null;
  readonly locked?: boolean | null;
  readonly formulaHidden?: boolean | null;
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
  if (patch.strikethrough !== undefined) {
    if (patch.strikethrough === null) {
      delete next.strikethrough;
    } else {
      next.strikethrough = patch.strikethrough;
    }
  }
  if (patch.fontScript !== undefined) {
    if (patch.fontScript === null) {
      delete next.fontScript;
    } else {
      next.fontScript = patch.fontScript;
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
  if (patch.fillPatternType !== undefined) {
    if (patch.fillPatternType === null) {
      delete next.fillPatternType;
    } else {
      next.fillPatternType = patch.fillPatternType;
    }
  }
  if (patch.fillPatternFgArgb !== undefined) {
    if (patch.fillPatternFgArgb === null) {
      delete next.fillPatternFgArgb;
    } else {
      next.fillPatternFgArgb = patch.fillPatternFgArgb;
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
  if (patch.textRotationDegrees !== undefined) {
    if (patch.textRotationDegrees === null) {
      delete next.textRotationDegrees;
    } else {
      const d = Math.round(patch.textRotationDegrees);
      if (!Number.isFinite(d) || d === 0) {
        delete next.textRotationDegrees;
      } else {
        next.textRotationDegrees = Math.max(-90, Math.min(90, d));
      }
      if (next.textRotationDegrees !== undefined) {
        delete next.textOrientation;
      }
    }
  }
  if (patch.shrinkToFit !== undefined) {
    if (patch.shrinkToFit === null) {
      delete next.shrinkToFit;
    } else {
      next.shrinkToFit = patch.shrinkToFit;
    }
  }
  if (patch.textOrientation !== undefined) {
    if (patch.textOrientation === null) {
      delete next.textOrientation;
    } else {
      next.textOrientation = patch.textOrientation;
      if (patch.textOrientation !== "horizontal") {
        delete next.textRotationDegrees;
      }
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
  if (patch.locked !== undefined) {
    if (patch.locked === null) {
      delete next.locked;
    } else {
      next.locked = patch.locked;
    }
  }
  if (patch.formulaHidden !== undefined) {
    if (patch.formulaHidden === null) {
      delete next.formulaHidden;
    } else {
      next.formulaHidden = patch.formulaHidden;
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

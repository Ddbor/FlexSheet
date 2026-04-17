import { CELL_FILL_PATTERN_TYPES, type CellFillPatternType, type CellStyle, type CellStylePatch } from "@flexsheet/core";

/** 可变的样式补丁对象（`CellStylePatch` 字段为 readonly）。 */
type CellStylePatchMutable = { -readonly [K in keyof CellStylePatch]: CellStylePatch[K] };

/**
 * 「设置单元格格式」填充页图案网格顺序（3 行 × 6 列），与 Excel 一致：
 * 第 1 行：无 / 灰度 75%→6.25%；第 2 行：粗线；第 3 行：细线。
 */
export const FORMAT_CELLS_PATTERN_GRID_ORDER: readonly CellFillPatternType[] = [
  "none",
  "darkGray",
  "mediumGray",
  "lightGray",
  "gray125",
  "gray0625",
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

export interface FormatCellsFillState {
  /** 背景色，`null` 为无颜色。 */
  bgArgb: string | null;
  patternType: CellFillPatternType;
  /** 图案颜色是否为「自动」（近黑色）。 */
  patternFgAuto: boolean;
  /** 非自动时的图案前景 ARGB。 */
  patternFgArgb: string | null;
}

function normalizeArgb8(v: string | undefined): string | null {
  const t = v?.trim();
  if (t === undefined || t === "" || !/^[\dA-Fa-f]{8}$/i.test(t)) {
    return null;
  }
  return t.toUpperCase();
}

export function inferFormatCellsFillState(style: CellStyle | null): FormatCellsFillState {
  const bgArgb = normalizeArgb8(style?.fillArgb);
  const rawPt = style?.fillPatternType;
  const patternType: CellFillPatternType =
    rawPt !== undefined && (CELL_FILL_PATTERN_TYPES as readonly string[]).includes(rawPt)
      ? rawPt
      : "none";
  const fgArgb = normalizeArgb8(style?.fillPatternFgArgb);
  /** 无 ARGB 表示「自动」；可先选图案色再选样式（对话框内暂存）。 */
  const patternFgAuto = fgArgb === null;
  return {
    bgArgb,
    patternType,
    patternFgAuto,
    patternFgArgb: fgArgb,
  };
}

export function formatCellsFillStateEqual(a: FormatCellsFillState, b: FormatCellsFillState): boolean {
  return (
    a.bgArgb === b.bgArgb &&
    a.patternType === b.patternType &&
    a.patternFgAuto === b.patternFgAuto &&
    a.patternFgArgb === b.patternFgArgb
  );
}

/** 写入单元格模型的图案前景：无图案或「自动」时均为 null。 */
function cellFgForModel(s: FormatCellsFillState): string | null {
  if (s.patternType === "none" || s.patternFgAuto) {
    return null;
  }
  return s.patternFgArgb;
}

export function fillPatchIfChanged(cur: FormatCellsFillState, ini: FormatCellsFillState): CellStylePatch {
  const p: CellStylePatchMutable = {};
  if (cur.bgArgb !== ini.bgArgb) {
    p.fillArgb = cur.bgArgb;
  }
  if (cur.patternType !== ini.patternType) {
    p.fillPatternType = cur.patternType === "none" ? null : cur.patternType;
  }
  const curM = cellFgForModel(cur);
  const iniM = cellFgForModel(ini);
  if (curM !== iniM || cur.patternFgAuto !== ini.patternFgAuto) {
    if (cur.patternType === "none") {
      p.fillPatternFgArgb = null;
    } else if (cur.patternFgAuto) {
      p.fillPatternFgArgb = null;
    } else {
      p.fillPatternFgArgb = cur.patternFgArgb;
    }
  }
  return p;
}

/**
 * 条件格式：规则模型与求值（纯 Data 层，不含 UI）。
 * 渲染层仅消费 `resolveConditionalFormattingOverlay` 的叠加样式。
 */

import type { CellBorderSide, CellStylePatch } from "./cell.js";
import { excelSerialToUtcDate, formatCellDisplayWithStyle } from "./excel-number-format.js";
import {
  normalizeSelectionRange,
  selectionRangeContains,
  type SelectionRange,
} from "./selection-range.js";
import type { Worksheet } from "./worksheet.js";

/** Ribbon「样式」下拉中的条件格式 UI 大类（与 Excel 对齐）。 */
export type ConditionalFormatUiFamily =
  | "twoColorScale"
  | "threeColorScale"
  | "dataBar"
  | "iconSet"
  | "classic";

/** 「经典」下主规则类型；`iconSet` 仅与 `uiFamily: "iconSet"` 联用；`colorScale` 与色阶 `uiFamily` 联用；`dataBar` 与 `uiFamily: "dataBar"` 联用。 */
export type ConditionalFormatClassicRuleType =
  | "cellsThatContain"
  | "topBottomRanked"
  | "aboveBelowAverage"
  | "uniqueOrDuplicate"
  | "formula"
  | "iconSet"
  | "colorScale"
  | "dataBar";

/** 色阶端点「类型」列（与 Excel 中文界面一致）。 */
export type CfColorScaleEndpointType =
  | "lowest"
  | "highest"
  | "number"
  | "percent"
  | "formula"
  | "percentile";

/** 数据条「最小值」列类型（与 Excel 中文界面一致）。 */
export type CfDataBarMinEndpointType =
  | "automatic"
  | "lowest"
  | "number"
  | "percent"
  | "formula"
  | "percentile";

/** 数据条「最大值」列类型（与 Excel 中文界面一致）。 */
export type CfDataBarMaxEndpointType =
  | "automatic"
  | "highest"
  | "number"
  | "percent"
  | "formula"
  | "percentile";

/** 数据条方向。 */
export type CfDataBarDirection = "context" | "leftToRight" | "rightToLeft";

/** 数据条填充：纯色 / 渐变。 */
export type CfDataBarFillKind = "solid" | "gradient";

/** 数据条边框：纯色描边 / 无。 */
export type CfDataBarBorderKind = "solid" | "none";

/** 数据条坐标轴位置。 */
export type CfDataBarAxisPosition = "automatic" | "midpoint" | "none";

/** 双色刻度一端：类型 + 可选数值 + 颜色（`FFRRGGBB`）。 */
export interface CfColorScaleEndpoint {
  readonly type: CfColorScaleEndpointType;
  /** `number` / `percent` / `percentile` / `formula` 时使用；`lowest`/`highest` 可为空。 */
  readonly value: string;
  readonly colorArgb: string;
}

/** 「只为包含以下内容的单元格设置格式」子类型。 */
export type CfCellsThatContainKind =
  | "cellValue"
  | "specificText"
  | "dateOccurring"
  | "blanks"
  | "noBlanks"
  | "errors"
  | "noErrors";

/** 单元格值比较运算符（与 Excel 中文界面一致）。 */
export type CfValueOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";

export type CfTextOperator = "contains" | "notContains" | "beginsWith" | "endsWith";

export type CfDateOccurring =
  | "today"
  | "yesterday"
  | "tomorrow"
  | "thisWeek"
  | "lastWeek"
  | "nextWeek"
  | "thisMonth"
  | "lastMonth"
  | "nextMonth";

export type CfTopBottomKind = "top" | "bottom" | "topPercent" | "bottomPercent";

export type CfAverageKind = "above" | "below" | "equalOrAbove" | "equalOrBelow";

export type CfUniqueKind = "duplicate" | "unique";

/** 预设格式（与 Excel 内置名称对应）。 */
export type CfFormatPresetId =
  | "lightRedFillDarkRedText"
  | "yellowFillDarkYellowText"
  | "greenFillDarkGreenText"
  | "lightRedFill"
  | "redText"
  | "redBorder"
  | "custom"
  /** 图标集等不叠加填充/字体时使用。 */
  | "none";

/** 条件格式图标集：单枚图标标识（Canvas 绘制与 Ribbon 选择器共用）。 */
export type CfIconGlyphId =
  | "none"
  | "circle_green"
  | "circle_yellow"
  | "circle_red"
  | "arrow_up_green"
  | "arrow_right_yellow"
  | "arrow_down_red"
  | "arrow_up_gray"
  | "arrow_right_gray"
  | "arrow_down_gray"
  | "flag_green"
  | "flag_yellow"
  | "flag_red"
  | "tri_up_green"
  | "tri_right_yellow"
  | "tri_down_red"
  | "sign_check_green"
  | "sign_exclaim_yellow"
  | "sign_cross_red"
  | "star_full"
  | "star_half"
  | "star_empty"
  | "bar4_1"
  | "bar4_2"
  | "bar4_3"
  | "bar4_4"
  | "bar5_1"
  | "bar5_2"
  | "bar5_3"
  | "bar5_4"
  | "bar5_5";

/** 内置图标集样式（与 Ribbon 下拉项一一对应）。 */
export type CfIconSetId =
  | "traffic3"
  | "arrows3"
  | "arrows3_gray"
  | "flags3"
  | "shapes3"
  | "signs3"
  | "stars3"
  | "bars4"
  | "bars5";

/** 图标集阈值行的「类型」列（与 Excel 中文界面一致）。 */
export type CfIconThresholdValueType = "number" | "percent" | "percentile" | "formula";

export interface CfIconThresholdRow {
  readonly operator: "greaterThanOrEqual";
  readonly value: string;
  readonly valueType: CfIconThresholdValueType;
}

export interface CfIconSetCatalogEntry {
  readonly id: CfIconSetId;
  readonly label: string;
  readonly iconCount: 3 | 4 | 5;
  readonly defaultGlyphs: readonly CfIconGlyphId[];
}

/** 内置图标集目录。 */
export const CF_ICON_SET_CATALOG: readonly CfIconSetCatalogEntry[] = [
  {
    id: "traffic3",
    label: "三向信号灯（无边框）",
    iconCount: 3,
    defaultGlyphs: ["circle_green", "circle_yellow", "circle_red"],
  },
  {
    id: "arrows3",
    label: "三色箭头",
    iconCount: 3,
    defaultGlyphs: ["arrow_up_green", "arrow_right_yellow", "arrow_down_red"],
  },
  {
    id: "arrows3_gray",
    label: "三色灰箭头",
    iconCount: 3,
    defaultGlyphs: ["arrow_up_gray", "arrow_right_gray", "arrow_down_gray"],
  },
  {
    id: "flags3",
    label: "三色旗",
    iconCount: 3,
    defaultGlyphs: ["flag_green", "flag_yellow", "flag_red"],
  },
  {
    id: "shapes3",
    label: "三色几何",
    iconCount: 3,
    defaultGlyphs: ["tri_up_green", "tri_right_yellow", "tri_down_red"],
  },
  {
    id: "signs3",
    label: "对勾 / 叹号 / 叉",
    iconCount: 3,
    defaultGlyphs: ["sign_check_green", "sign_exclaim_yellow", "sign_cross_red"],
  },
  {
    id: "stars3",
    label: "星级",
    iconCount: 3,
    defaultGlyphs: ["star_full", "star_half", "star_empty"],
  },
  {
    id: "bars4",
    label: "四格信号条",
    iconCount: 4,
    defaultGlyphs: ["bar4_4", "bar4_3", "bar4_2", "bar4_1"],
  },
  {
    id: "bars5",
    label: "五格信号条",
    iconCount: 5,
    defaultGlyphs: ["bar5_5", "bar5_4", "bar5_3", "bar5_2", "bar5_1"],
  },
];

export function findCfIconSetCatalogEntry(id: CfIconSetId): CfIconSetCatalogEntry | undefined {
  return CF_ICON_SET_CATALOG.find((e) => e.id === id);
}

/** 供图标选择器：常用图标平面列表（含「无」）。 */
export const CF_ICON_GLYPH_PICKER_ORDER: readonly CfIconGlyphId[] = [
  "none",
  "circle_green",
  "circle_yellow",
  "circle_red",
  "arrow_up_green",
  "arrow_right_yellow",
  "arrow_down_red",
  "arrow_up_gray",
  "arrow_right_gray",
  "arrow_down_gray",
  "flag_green",
  "flag_yellow",
  "flag_red",
  "tri_up_green",
  "tri_right_yellow",
  "tri_down_red",
  "sign_check_green",
  "sign_exclaim_yellow",
  "sign_cross_red",
  "star_full",
  "star_half",
  "star_empty",
  "bar4_4",
  "bar4_3",
  "bar4_2",
  "bar4_1",
  "bar5_5",
  "bar5_4",
  "bar5_3",
  "bar5_2",
  "bar5_1",
];

export interface ConditionalFormattingCellIcon {
  readonly glyphId: CfIconGlyphId;
  readonly hideCellValue?: boolean;
}

export interface ConditionalFormatRule {
  readonly id: string;
  readonly range: SelectionRange;
  readonly uiFamily: ConditionalFormatUiFamily;
  readonly classicType: ConditionalFormatClassicRuleType;
  readonly cellsThatContainKind?: CfCellsThatContainKind;
  readonly valueOperator?: CfValueOperator;
  readonly textOperator?: CfTextOperator;
  readonly dateOccurring?: CfDateOccurring;
  readonly value1?: string;
  readonly value2?: string;
  readonly topBottomKind?: CfTopBottomKind;
  /** 前/后 N 项或 N%（1–1000）。 */
  readonly topBottomN?: number;
  readonly averageKind?: CfAverageKind;
  readonly uniqueKind?: CfUniqueKind;
  readonly formulaExpression?: string;
  readonly formatPreset: CfFormatPresetId;
  /** `formatPreset === "custom"` 时使用。 */
  readonly customFormat?: CellStylePatch;
  readonly cfIconSetId?: CfIconSetId;
  readonly cfIconGlyphs?: readonly CfIconGlyphId[];
  readonly cfIconReverseOrder?: boolean;
  readonly cfIconShowIconOnly?: boolean;
  readonly cfIconThresholds?: readonly CfIconThresholdRow[];
  /** `uiFamily === "twoColorScale"` 时：最小值端。 */
  readonly cfTwoColorMin?: CfColorScaleEndpoint;
  /** `uiFamily === "twoColorScale"` 时：最大值端。 */
  readonly cfTwoColorMax?: CfColorScaleEndpoint;
  /** `uiFamily === "threeColorScale"` 时：最小值端。 */
  readonly cfThreeColorMin?: CfColorScaleEndpoint;
  /** `uiFamily === "threeColorScale"` 时：中间值端。 */
  readonly cfThreeColorMid?: CfColorScaleEndpoint;
  /** `uiFamily === "threeColorScale"` 时：最大值端。 */
  readonly cfThreeColorMax?: CfColorScaleEndpoint;
  /** `uiFamily === "dataBar"`：最小值端类型与值。 */
  readonly cfDataBarMin?: { readonly type: CfDataBarMinEndpointType; readonly value: string };
  /** `uiFamily === "dataBar"`：最大值端类型与值。 */
  readonly cfDataBarMax?: { readonly type: CfDataBarMaxEndpointType; readonly value: string };
  readonly cfDataBarDirection?: CfDataBarDirection;
  /** 「仅显示数据栏」：不绘制单元格数值文本（与 Excel showBarOnly 对齐）。 */
  readonly cfDataBarShowBarOnly?: boolean;
  readonly cfDataBarFillKind?: CfDataBarFillKind;
  readonly cfDataBarPositiveFillArgb?: string;
  readonly cfDataBarNegativeFillArgb?: string;
  readonly cfDataBarBorderKind?: CfDataBarBorderKind;
  readonly cfDataBarPositiveBorderArgb?: string;
  readonly cfDataBarNegativeBorderArgb?: string;
  readonly cfDataBarAxisPosition?: CfDataBarAxisPosition;
  readonly cfDataBarAxisColorArgb?: string;
}

/** 与单元格基础样式合并：仅包含有值的字段。 */
/** 供 Canvas 绘制的数据条一帧描述（相对单元格内边距框 0–1 分数坐标，原点左上角）。 */
export interface ConditionalFormattingDataBarPaint {
  /** 数据条矩形左边界（fraction）。 */
  readonly barX0Frac: number;
  /** 数据条矩形右边界（fraction）。 */
  readonly barX1Frac: number;
  /** 相对行高，条带高度占比（0–1）。 */
  readonly barHeightFrac: number;
  readonly fillKind: CfDataBarFillKind;
  readonly posFillArgb: string;
  readonly negFillArgb: string;
  /** 本条带是否使用「正值」颜色（边框取色时用）。 */
  readonly usePositiveFill: boolean;
  readonly border:
    | { readonly kind: "none" }
    | { readonly kind: "solid"; readonly posArgb: string; readonly negArgb: string };
  /** 坐标轴竖线位置（fraction），不绘制时为 `null`。 */
  readonly axisXFrac: number | null;
  readonly axisColorArgb: string;
  /** 为 `true` 时文本层应跳过该格数值。 */
  readonly hideCellValue: boolean;
  /** 为 `true` 时水平镜像条与轴位置（从右到左 / RTL 上下文）。 */
  readonly rtl: boolean;
  /**
   * 当前值在 [min,max] 刻度上的归一化位置（0=最小端，1=最大端）。
   * 与「仅彩色段」的 barX 无关，用于渲染层做从起点到数值端的渐变过渡。
   */
  readonly valueNormFrac: number;
}

export interface ConditionalFormattingOverlay {
  readonly fillArgb?: string;
  readonly fgArgb?: string;
  readonly borderTop?: CellBorderSide;
  readonly borderLeft?: CellBorderSide;
  readonly borderBottom?: CellBorderSide;
  readonly borderRight?: CellBorderSide;
  readonly icon?: ConditionalFormattingCellIcon;
  readonly dataBar?: ConditionalFormattingDataBarPaint;
}

const RED_BORDER: CellBorderSide = { kind: "thin", colorArgb: "FFFF0000" };

export function cfFormatPresetToOverlay(
  preset: CfFormatPresetId,
  custom?: CellStylePatch,
): ConditionalFormattingOverlay {
  if (preset === "custom" && custom !== undefined) {
    return {
      ...(custom.fillArgb !== undefined && custom.fillArgb !== null
        ? { fillArgb: custom.fillArgb }
        : {}),
      ...(custom.fgArgb !== undefined && custom.fgArgb !== null ? { fgArgb: custom.fgArgb } : {}),
      ...(custom.borderTop !== undefined && custom.borderTop !== null
        ? { borderTop: custom.borderTop }
        : {}),
      ...(custom.borderLeft !== undefined && custom.borderLeft !== null
        ? { borderLeft: custom.borderLeft }
        : {}),
      ...(custom.borderBottom !== undefined && custom.borderBottom !== null
        ? { borderBottom: custom.borderBottom }
        : {}),
      ...(custom.borderRight !== undefined && custom.borderRight !== null
        ? { borderRight: custom.borderRight }
        : {}),
    } as ConditionalFormattingOverlay;
  }
  switch (preset) {
    case "lightRedFillDarkRedText":
      return { fillArgb: "FFFFC7CE", fgArgb: "FF9C0006" };
    case "yellowFillDarkYellowText":
      return { fillArgb: "FFFFEB9C", fgArgb: "FF9C6500" };
    case "greenFillDarkGreenText":
      return { fillArgb: "FFC6EFCE", fgArgb: "FF006100" };
    case "lightRedFill":
      return { fillArgb: "FFFFC7CE" };
    case "redText":
      return { fgArgb: "FFFF0000" };
    case "redBorder":
      return {
        borderTop: RED_BORDER,
        borderLeft: RED_BORDER,
        borderBottom: RED_BORDER,
        borderRight: RED_BORDER,
      };
    case "none":
      return {};
    default:
      return {};
  }
}

function isBlankValue(v: unknown): boolean {
  return v === null || v === "";
}

function cellLooksLikeError(sheet: Worksheet, row: number, col: number): boolean {
  const cell = sheet.getCell(row, col);
  if (cell.formula === null || cell.formula.trim() === "") {
    return false;
  }
  const disp = formatCellDisplayWithStyle(cell.value, cell.style);
  return /^#[A-Z]{3,5}[!/]?$/i.test(disp.trim()) || disp.includes("#N/A");
}

function utcDayStart(d: Date): number {
  return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}

function excelSerialDayFloor(serial: number): number {
  return Math.floor(serial);
}

function todaySerialUtc(): number {
  const d = new Date();
  return Math.round(utcDayStart(d) / 86400000) + 25569;
}

function cellDateSerialDay(cellVal: unknown): number | null {
  if (typeof cellVal !== "number" || !Number.isFinite(cellVal)) {
    return null;
  }
  return excelSerialDayFloor(cellVal);
}

function startOfUtcMonth(serial: number): number {
  const d = excelSerialToUtcDate(serial);
  const t = Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1);
  return Math.round(t / 86400000) + 25569;
}

function endOfUtcMonth(serial: number): number {
  const d = excelSerialToUtcDate(serial);
  const t = Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0);
  return Math.round(t / 86400000) + 25569;
}

function startOfIsoWeek(serial: number): number {
  const d = excelSerialToUtcDate(serial);
  const day = d.getUTCDay();
  const diff = (day + 6) % 7;
  const t = utcDayStart(d) - diff * 86400000;
  return Math.round(t / 86400000) + 25569;
}

function matchesDateOccurring(serialDay: number, kind: CfDateOccurring): boolean {
  const t0 = todaySerialUtc();
  if (kind === "today") {
    return serialDay === t0;
  }
  if (kind === "yesterday") {
    return serialDay === t0 - 1;
  }
  if (kind === "tomorrow") {
    return serialDay === t0 + 1;
  }
  if (kind === "thisWeek") {
    const ws = startOfIsoWeek(t0);
    return serialDay >= ws && serialDay < ws + 7;
  }
  if (kind === "lastWeek") {
    const ws = startOfIsoWeek(t0);
    const prev = ws - 7;
    return serialDay >= prev && serialDay < ws;
  }
  if (kind === "nextWeek") {
    const ws = startOfIsoWeek(t0);
    const nx = ws + 7;
    return serialDay >= nx && serialDay < nx + 7;
  }
  if (kind === "thisMonth") {
    const sm = startOfUtcMonth(t0);
    const em = endOfUtcMonth(t0);
    return serialDay >= sm && serialDay <= em;
  }
  if (kind === "lastMonth") {
    const d = excelSerialToUtcDate(t0);
    const prevMonth = Date.UTC(d.getUTCFullYear(), d.getUTCMonth() - 1, 15);
    const mid = Math.round(prevMonth / 86400000) + 25569;
    const sm = startOfUtcMonth(mid);
    const em = endOfUtcMonth(mid);
    return serialDay >= sm && serialDay <= em;
  }
  if (kind === "nextMonth") {
    const d = excelSerialToUtcDate(t0);
    const nextMonth = Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 15);
    const mid = Math.round(nextMonth / 86400000) + 25569;
    const sm = startOfUtcMonth(mid);
    const em = endOfUtcMonth(mid);
    return serialDay >= sm && serialDay <= em;
  }
  return false;
}

function compareValues(
  op: CfValueOperator,
  cellVal: unknown,
  bound1: string,
  bound2: string | undefined,
): boolean {
  const a = cellVal;
  const b1 = bound1.trim();
  const b2 = (bound2 ?? "").trim();
  const numB1 = Number(b1);
  const numB2 = Number(b2);
  const b1Num = b1 !== "" && Number.isFinite(numB1);
  const b2Num = b2 !== "" && Number.isFinite(numB2);
  const naRaw = typeof a === "number" ? a : Number(String(a).trim());
  const aIsNum =
    typeof a === "number" || (typeof a === "string" && a.trim() !== "" && Number.isFinite(naRaw));
  const na = aIsNum ? (typeof a === "number" ? a : naRaw) : NaN;

  if (aIsNum && b1Num && Number.isFinite(na)) {
    if ((op === "between" || op === "notBetween") && b2Num) {
      const lo = Math.min(numB1, numB2);
      const hi = Math.max(numB1, numB2);
      if (op === "between") {
        return na >= lo && na <= hi;
      }
      return na < lo || na > hi;
    }
    if (op === "between" || op === "notBetween") {
      return false;
    }
    switch (op) {
      case "equal":
        return na === numB1;
      case "notEqual":
        return na !== numB1;
      case "greaterThan":
        return na > numB1;
      case "lessThan":
        return na < numB1;
      case "greaterThanOrEqual":
        return na >= numB1;
      case "lessThanOrEqual":
        return na <= numB1;
      default:
        return false;
    }
  }

  const sa = a === null || a === undefined ? "" : String(a);
  switch (op) {
    case "equal":
      return sa === b1;
    case "notEqual":
      return sa !== b1;
    case "greaterThan":
      return sa > b1;
    case "lessThan":
      return sa < b1;
    case "greaterThanOrEqual":
      return sa >= b1;
    case "lessThanOrEqual":
      return sa <= b1;
    case "between":
    case "notBetween":
    default:
      return false;
  }
}

function matchesText(op: CfTextOperator, cellText: string, needle: string): boolean {
  const n = needle;
  switch (op) {
    case "contains":
      return cellText.includes(n);
    case "notContains":
      return !cellText.includes(n);
    case "beginsWith":
      return cellText.startsWith(n);
    case "endsWith":
      return cellText.endsWith(n);
    default:
      return false;
  }
}

function collectNumericCellsInRange(sheet: Worksheet, range: SelectionRange): number[] {
  const n = normalizeSelectionRange(range);
  const nums: number[] = [];
  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      const v = sheet.getCell(r, c).value;
      if (typeof v === "number" && Number.isFinite(v)) {
        nums.push(v);
      }
    }
  }
  return nums;
}

function averageOf(nums: number[]): number | null {
  if (nums.length === 0) {
    return null;
  }
  let s = 0;
  for (const x of nums) {
    s += x;
  }
  return s / nums.length;
}

function topBottomThreshold(nums: number[], kind: CfTopBottomKind, nn: number): number | null {
  if (nums.length === 0 || nn <= 0) {
    return null;
  }
  const sorted = [...nums].sort((a, b) => a - b);
  if (kind === "top" || kind === "topPercent") {
    const k =
      kind === "topPercent"
        ? Math.max(1, Math.ceil((sorted.length * Math.min(100, nn)) / 100))
        : Math.min(sorted.length, Math.floor(nn));
    const cut = sorted[sorted.length - k];
    return cut ?? null;
  }
  const k =
    kind === "bottomPercent"
      ? Math.max(1, Math.ceil((sorted.length * Math.min(100, nn)) / 100))
      : Math.min(sorted.length, Math.floor(nn));
  const cut = sorted[k - 1];
  return cut ?? null;
}

function duplicateKeyForCell(sheet: Worksheet, row: number, col: number): string {
  const v = sheet.getCell(row, col).value;
  const disp = formatCellDisplayWithStyle(v, sheet.getCell(row, col).style);
  return disp.trim() === "" ? "__BLANK__" : disp;
}

function buildDuplicateKeyCounts(sheet: Worksheet, range: SelectionRange): Map<string, number> {
  const n = normalizeSelectionRange(range);
  const m = new Map<string, number>();
  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      const k = duplicateKeyForCell(sheet, r, c);
      m.set(k, (m.get(k) ?? 0) + 1);
    }
  }
  return m;
}

/**
 * 依赖整段选区扫描的规则，在同 `Worksheet.revision` 下对每个 `rule.id` 只预计算一次，
 * 避免「全表选区 + 重复值/排名/平均值」时对每个单元格重复 O(区域) 扫描导致卡顿。
 */
interface ConditionalFormatRuleRangeEvalCache {
  readonly duplicateKeyCounts?: ReadonlyMap<string, number>;
  readonly rangeAverage?: number | null;
  readonly topBottomThresholdValue?: number | null;
  readonly iconSetSortedNums?: readonly number[];
  readonly iconSetBounds?: readonly number[];
  readonly colorScaleSortedNums?: readonly number[];
}

interface ConditionalFormatRuleEvalBucket {
  revision: number;
  readonly byRuleId: Map<string, ConditionalFormatRuleRangeEvalCache>;
}

const conditionalFormatRuleEvalCaches = new WeakMap<Worksheet, ConditionalFormatRuleEvalBucket>();

function getCfRuleEvalBucket(sheet: Worksheet): ConditionalFormatRuleEvalBucket {
  const rev = sheet.revision;
  let b = conditionalFormatRuleEvalCaches.get(sheet);
  if (b === undefined || b.revision !== rev) {
    b = { revision: rev, byRuleId: new Map() };
    conditionalFormatRuleEvalCaches.set(sheet, b);
  }
  return b;
}

function percentileInclusive(sorted: readonly number[], p: number): number {
  if (sorted.length === 0) {
    return NaN;
  }
  if (sorted.length === 1) {
    return sorted[0]!;
  }
  const pp = Math.max(0, Math.min(100, p));
  const rank = (pp / 100) * (sorted.length - 1);
  const lo = Math.floor(rank);
  const hi = Math.ceil(rank);
  const t = rank - lo;
  const a = sorted[lo]!;
  const b = sorted[hi]!;
  return a * (1 - t) + b * t;
}

function parseArgbToRgbChannels(argb: string): { r: number; g: number; b: number } | null {
  const t = argb.trim();
  if (!/^[\dA-Fa-f]{8}$/i.test(t)) {
    return null;
  }
  const r = parseInt(t.slice(2, 4), 16);
  const g = parseInt(t.slice(4, 6), 16);
  const b = parseInt(t.slice(6, 8), 16);
  if (![r, g, b].every((n) => Number.isFinite(n))) {
    return null;
  }
  return { r, g, b };
}

function cfLerpFillArgb(a: string, b: string, t: number): string {
  const ca = parseArgbToRgbChannels(a);
  const cb = parseArgbToRgbChannels(b);
  const u = Math.max(0, Math.min(1, t));
  if (ca === null || cb === null) {
    const g = Math.round(u * 255)
      .toString(16)
      .padStart(2, "0");
    return `FF${g}${g}${g}`.toUpperCase();
  }
  const r = Math.round(ca.r + (cb.r - ca.r) * u);
  const gg = Math.round(ca.g + (cb.g - ca.g) * u);
  const bb = Math.round(ca.b + (cb.b - ca.b) * u);
  const rr = Math.max(0, Math.min(255, r)).toString(16).padStart(2, "0");
  const g2 = Math.max(0, Math.min(255, gg)).toString(16).padStart(2, "0");
  const b2 = Math.max(0, Math.min(255, bb)).toString(16).padStart(2, "0");
  return `FF${rr}${g2}${b2}`.toUpperCase();
}

function resolveColorScaleEndpointBound(
  ep: CfColorScaleEndpoint,
  sortedNums: readonly number[],
  dataMin: number,
  dataMax: number,
  fallbackNumber: number,
): number {
  const span = dataMax - dataMin;
  switch (ep.type) {
    case "lowest":
      return dataMin;
    case "highest":
      return dataMax;
    case "number":
    case "formula": {
      const raw = ep.value.trim().replace(/^=/, "");
      if (raw === "") {
        return fallbackNumber;
      }
      const n = Number(raw);
      return Number.isFinite(n) ? n : fallbackNumber;
    }
    case "percent": {
      const p = Number(ep.value.trim());
      const pp = Number.isFinite(p) ? Math.max(0, Math.min(100, p)) : 0;
      return dataMin + span * (pp / 100);
    }
    case "percentile": {
      if (sortedNums.length === 0) {
        return fallbackNumber;
      }
      const p = Number(ep.value.trim());
      const pp = Number.isFinite(p) ? Math.max(0, Math.min(100, p)) : 0;
      return percentileInclusive(sortedNums, pp);
    }
    default:
      return fallbackNumber;
  }
}

/** 双色刻度：在规则范围内且单元格为有限数字时返回插值填充色，否则 `null`。 */
export function resolveTwoColorScaleFillArgb(
  sheet: Worksheet,
  row: number,
  col: number,
  rule: ConditionalFormatRule,
): string | null {
  if (rule.uiFamily !== "twoColorScale") {
    return null;
  }
  if (!selectionRangeContains(rule.range, row, col)) {
    return null;
  }
  const val = sheet.getCell(row, col).value;
  if (typeof val !== "number" || !Number.isFinite(val)) {
    return null;
  }
  const cache = getRuleRangeEvalCache(sheet, rule);
  const sorted = cache.colorScaleSortedNums ?? [];
  if (sorted.length === 0) {
    return null;
  }
  const dataMin = sorted[0]!;
  const dataMax = sorted[sorted.length - 1]!;
  const minEp: CfColorScaleEndpoint =
    rule.cfTwoColorMin ?? { type: "lowest", value: "", colorArgb: "FFFF6600" };
  const maxEp: CfColorScaleEndpoint =
    rule.cfTwoColorMax ?? { type: "highest", value: "", colorArgb: "FFFFFFCC" };
  let lo = resolveColorScaleEndpointBound(minEp, sorted, dataMin, dataMax, dataMin);
  let hi = resolveColorScaleEndpointBound(maxEp, sorted, dataMin, dataMax, dataMax);
  let cLo = minEp.colorArgb;
  let cHi = maxEp.colorArgb;
  if (lo > hi) {
    const tmp = lo;
    lo = hi;
    hi = tmp;
    const tc = cLo;
    cLo = cHi;
    cHi = tc;
  }
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) {
    return null;
  }
  if (hi === lo) {
    return cfLerpFillArgb(cLo, cHi, 0.5);
  }
  const t = (val - lo) / (hi - lo);
  return cfLerpFillArgb(cLo, cHi, t);
}

/** 三色刻度：三端点按数值排序后分段线性插值；非数字或无数值格返回 `null`。 */
export function resolveThreeColorScaleFillArgb(
  sheet: Worksheet,
  row: number,
  col: number,
  rule: ConditionalFormatRule,
): string | null {
  if (rule.uiFamily !== "threeColorScale") {
    return null;
  }
  if (!selectionRangeContains(rule.range, row, col)) {
    return null;
  }
  const val = sheet.getCell(row, col).value;
  if (typeof val !== "number" || !Number.isFinite(val)) {
    return null;
  }
  const cache = getRuleRangeEvalCache(sheet, rule);
  const sorted = cache.colorScaleSortedNums ?? [];
  if (sorted.length === 0) {
    return null;
  }
  const dataMin = sorted[0]!;
  const dataMax = sorted[sorted.length - 1]!;
  const minEp: CfColorScaleEndpoint =
    rule.cfThreeColorMin ?? { type: "lowest", value: "", colorArgb: "FFFF6600" };
  const midEp: CfColorScaleEndpoint =
    rule.cfThreeColorMid ?? { type: "percentile", value: "50", colorArgb: "FFFFFFCC" };
  const maxEp: CfColorScaleEndpoint =
    rule.cfThreeColorMax ?? { type: "highest", value: "", colorArgb: "FFFFFFCC" };
  const b0 = resolveColorScaleEndpointBound(minEp, sorted, dataMin, dataMax, dataMin);
  const b1 = resolveColorScaleEndpointBound(
    midEp,
    sorted,
    dataMin,
    dataMax,
    (dataMin + dataMax) / 2,
  );
  const b2 = resolveColorScaleEndpointBound(maxEp, sorted, dataMin, dataMax, dataMax);
  const stops = [
    { b: b0, c: minEp.colorArgb },
    { b: b1, c: midEp.colorArgb },
    { b: b2, c: maxEp.colorArgb },
  ].sort((x, y) => x.b - y.b);
  const p0 = stops[0]!;
  const p1 = stops[1]!;
  const p2 = stops[2]!;
  if (!Number.isFinite(p0.b) || !Number.isFinite(p1.b) || !Number.isFinite(p2.b)) {
    return null;
  }
  if (p0.b === p2.b) {
    return cfLerpFillArgb(p0.c, p2.c, 0.5);
  }
  if (val <= p0.b) {
    return p0.c;
  }
  if (val >= p2.b) {
    return p2.c;
  }
  if (val <= p1.b) {
    if (p1.b === p0.b) {
      return p1.c;
    }
    return cfLerpFillArgb(p0.c, p1.c, (val - p0.b) / (p1.b - p0.b));
  }
  if (p2.b === p1.b) {
    return p1.c;
  }
  return cfLerpFillArgb(p1.c, p2.c, (val - p1.b) / (p2.b - p1.b));
}

function cfDataBarMinToScaleEp(ep: { readonly type: CfDataBarMinEndpointType; readonly value: string }): CfColorScaleEndpoint {
  switch (ep.type) {
    case "automatic":
    case "lowest":
      return { type: "lowest", value: "", colorArgb: "FF000000" };
    case "number":
      return { type: "number", value: ep.value, colorArgb: "FF000000" };
    case "percent":
      return { type: "percent", value: ep.value, colorArgb: "FF000000" };
    case "formula":
      return { type: "formula", value: ep.value, colorArgb: "FF000000" };
    case "percentile":
      return { type: "percentile", value: ep.value, colorArgb: "FF000000" };
    default:
      return { type: "lowest", value: "", colorArgb: "FF000000" };
  }
}

function cfDataBarMaxToScaleEp(ep: { readonly type: CfDataBarMaxEndpointType; readonly value: string }): CfColorScaleEndpoint {
  switch (ep.type) {
    case "automatic":
    case "highest":
      return { type: "highest", value: "", colorArgb: "FF000000" };
    case "number":
      return { type: "number", value: ep.value, colorArgb: "FF000000" };
    case "percent":
      return { type: "percent", value: ep.value, colorArgb: "FF000000" };
    case "formula":
      return { type: "formula", value: ep.value, colorArgb: "FF000000" };
    case "percentile":
      return { type: "percentile", value: ep.value, colorArgb: "FF000000" };
    default:
      return { type: "highest", value: "", colorArgb: "FF000000" };
  }
}

const DEFAULT_CF_DATA_BAR_POS = "FF638EC6";
const DEFAULT_CF_DATA_BAR_NEG = "FFFF0000";
const DEFAULT_CF_DATA_BAR_AXIS = "FF000000";
/** 正值数据条默认描边（深蓝，与 Excel「实心边框」观感接近）。 */
const DEFAULT_CF_DATA_BAR_POS_BORDER = "FF2F5597";

/** 求当前格数据条绘制参数；非数字或无数值样本时返回 `null`。 */
export function resolveDataBarPaintForCell(
  sheet: Worksheet,
  row: number,
  col: number,
  rule: ConditionalFormatRule,
): ConditionalFormattingDataBarPaint | null {
  if (rule.uiFamily !== "dataBar" || rule.classicType !== "dataBar") {
    return null;
  }
  if (!selectionRangeContains(rule.range, row, col)) {
    return null;
  }
  const cellVal = sheet.getCell(row, col).value;
  if (typeof cellVal !== "number" || !Number.isFinite(cellVal)) {
    return null;
  }
  const cache = getRuleRangeEvalCache(sheet, rule);
  const sorted = cache.colorScaleSortedNums ?? [];
  if (sorted.length === 0) {
    return null;
  }
  const dataMin = sorted[0]!;
  const dataMax = sorted[sorted.length - 1]!;

  const minEp = rule.cfDataBarMin ?? { type: "automatic", value: "" };
  const maxEp = rule.cfDataBarMax ?? { type: "automatic", value: "" };
  const lo = resolveColorScaleEndpointBound(
    cfDataBarMinToScaleEp(minEp),
    sorted,
    dataMin,
    dataMax,
    dataMin,
  );
  const hi = resolveColorScaleEndpointBound(
    cfDataBarMaxToScaleEp(maxEp),
    sorted,
    dataMin,
    dataMax,
    dataMax,
  );
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) {
    return null;
  }
  let loB = lo;
  let hiB = hi;
  if (loB > hiB) {
    const tmp = loB;
    loB = hiB;
    hiB = tmp;
  }

  const axisMode = rule.cfDataBarAxisPosition ?? "automatic";
  const span = hiB - loB;

  let tAxisFrac: number;
  if (span === 0 || !Number.isFinite(span)) {
    tAxisFrac = loB >= 0 ? 0 : 1;
  } else if (loB < 0 && hiB > 0) {
    tAxisFrac = (0 - loB) / span;
  } else if (loB >= 0) {
    tAxisFrac = 0;
  } else {
    tAxisFrac = 1;
  }

  let tValFrac: number;
  if (span === 0 || !Number.isFinite(span)) {
    tValFrac = 1;
  } else {
    tValFrac = (cellVal - loB) / span;
    tValFrac = Math.max(0, Math.min(1, tValFrac));
  }

  let barX0 = Math.min(tAxisFrac, tValFrac);
  let barX1 = Math.max(tAxisFrac, tValFrac);
  if (span === 0 || !Number.isFinite(span)) {
    barX0 = 0;
    barX1 = 1;
  }

  let axisDisplayX: number | null;
  if (axisMode === "none") {
    axisDisplayX = null;
  } else if (axisMode === "midpoint") {
    axisDisplayX = 0.5;
  } else {
    axisDisplayX = tAxisFrac;
  }

  const dir = rule.cfDataBarDirection ?? "context";
  const rtl = dir === "rightToLeft";

  const posArgb = rule.cfDataBarPositiveFillArgb ?? DEFAULT_CF_DATA_BAR_POS;
  const negArgb = rule.cfDataBarNegativeFillArgb ?? DEFAULT_CF_DATA_BAR_NEG;
  const axisArgb = rule.cfDataBarAxisColorArgb ?? DEFAULT_CF_DATA_BAR_AXIS;
  const fillKind = rule.cfDataBarFillKind ?? "solid";
  const brKind = rule.cfDataBarBorderKind ?? "solid";

  const borderPaint: ConditionalFormattingDataBarPaint["border"] =
    brKind === "solid"
      ? {
          kind: "solid",
          posArgb: rule.cfDataBarPositiveBorderArgb ?? DEFAULT_CF_DATA_BAR_POS_BORDER,
          negArgb: rule.cfDataBarNegativeBorderArgb ?? "FF000000",
        }
      : { kind: "none" };

  const mirrorX = (u: number): number => (rtl ? 1 - u : u);

  /** 全为正：统一用正值条色；全为负：负值条色；跨 0：按单元格符号（与 Excel 一致）。 */
  const usePositiveFill: boolean =
    span === 0 || !Number.isFinite(span)
      ? true
      : loB >= 0
        ? true
        : hiB <= 0
          ? false
          : cellVal >= 0;

  const ax0 = mirrorX(barX0);
  const ax1 = mirrorX(barX1);

  return {
    barX0Frac: Math.min(ax0, ax1),
    barX1Frac: Math.max(ax0, ax1),
    /** 条带占内框高度比例（与 Excel 默认数据条相近：近满高、数字叠在条上）。 */
    barHeightFrac: 0.92,
    fillKind,
    posFillArgb: posArgb,
    negFillArgb: negArgb,
    usePositiveFill,
    border: borderPaint,
    axisXFrac: axisDisplayX === null ? null : mirrorX(axisDisplayX),
    axisColorArgb: axisArgb,
    hideCellValue: rule.cfDataBarShowBarOnly === true,
    rtl,
    valueNormFrac: tValFrac,
  };
}

function thresholdRowToBound(sortedNums: readonly number[], row: CfIconThresholdRow): number {
  const raw = row.value.trim();
  const n = Number(raw);
  const t = row.valueType;
  if (t === "number" || t === "formula") {
    return Number.isFinite(n) ? n : NaN;
  }
  if (sortedNums.length === 0) {
    return NaN;
  }
  if (!Number.isFinite(n)) {
    return NaN;
  }
  return percentileInclusive(sortedNums, Math.max(0, Math.min(100, n)));
}

function materializeIconSetGlyphs(rule: ConditionalFormatRule): CfIconGlyphId[] {
  const id = rule.cfIconSetId ?? "traffic3";
  const cat = findCfIconSetCatalogEntry(id);
  const glyphs = [...(cat?.defaultGlyphs ?? ["circle_green", "circle_yellow", "circle_red"])];
  const over = rule.cfIconGlyphs;
  if (over !== undefined) {
    for (let i = 0; i < glyphs.length; i++) {
      if (over[i] !== undefined) {
        glyphs[i] = over[i]!;
      }
    }
  }
  if (rule.cfIconReverseOrder === true) {
    glyphs.reverse();
  }
  return glyphs;
}

function buildRuleRangeEvalCache(
  sheet: Worksheet,
  rule: ConditionalFormatRule,
): ConditionalFormatRuleRangeEvalCache {
  if (rule.uiFamily === "iconSet") {
    const sorted = collectNumericCellsInRange(sheet, rule.range).sort((a, b) => a - b);
    const thresholds = rule.cfIconThresholds ?? [];
    const bounds = thresholds.map((row) => thresholdRowToBound(sorted, row));
    return { iconSetSortedNums: sorted, iconSetBounds: bounds };
  }
  if (rule.uiFamily === "twoColorScale" || rule.uiFamily === "threeColorScale") {
    const sorted = collectNumericCellsInRange(sheet, rule.range).sort((a, b) => a - b);
    return { colorScaleSortedNums: sorted };
  }
  if (rule.uiFamily === "dataBar") {
    const sorted = collectNumericCellsInRange(sheet, rule.range).sort((a, b) => a - b);
    return { colorScaleSortedNums: sorted };
  }
  if (rule.uiFamily !== "classic") {
    return {};
  }
  switch (rule.classicType) {
    case "uniqueOrDuplicate":
      return { duplicateKeyCounts: buildDuplicateKeyCounts(sheet, rule.range) };
    case "aboveBelowAverage":
      return { rangeAverage: averageOf(collectNumericCellsInRange(sheet, rule.range)) };
    case "topBottomRanked": {
      const nums = collectNumericCellsInRange(sheet, rule.range);
      const nn = Math.max(1, Math.min(1000, rule.topBottomN ?? 10));
      const kind = rule.topBottomKind ?? "top";
      return { topBottomThresholdValue: topBottomThreshold(nums, kind, nn) };
    }
    case "iconSet":
      return {};
    default:
      return {};
  }
}

function getRuleRangeEvalCache(
  sheet: Worksheet,
  rule: ConditionalFormatRule,
): ConditionalFormatRuleRangeEvalCache {
  const b = getCfRuleEvalBucket(sheet);
  const hit = b.byRuleId.get(rule.id);
  if (hit !== undefined) {
    return hit;
  }
  const built = buildRuleRangeEvalCache(sheet, rule);
  b.byRuleId.set(rule.id, built);
  return built;
}

/** 求当前格在图标集规则下的图标；非数字或不在档内返回 `null`。 */
export function resolveIconSetGlyphForCell(
  sheet: Worksheet,
  row: number,
  col: number,
  rule: ConditionalFormatRule,
): CfIconGlyphId | null {
  if (rule.uiFamily !== "iconSet") {
    return null;
  }
  const val = sheet.getCell(row, col).value;
  if (typeof val !== "number" || !Number.isFinite(val)) {
    return null;
  }
  const glyphs = materializeIconSetGlyphs(rule);
  const n = glyphs.length;
  if (n < 2) {
    return null;
  }
  const cache = getRuleRangeEvalCache(sheet, rule);
  const bounds = cache.iconSetBounds ?? [];
  if (bounds.length !== n - 1) {
    return null;
  }
  for (let tier = 0; tier < n - 1; tier++) {
    const b = bounds[tier];
    if (b === undefined || !Number.isFinite(b)) {
      continue;
    }
    if (tier === 0) {
      if (val >= b) {
        return glyphs[0] ?? null;
      }
    } else {
      const prev = bounds[tier - 1];
      if (prev !== undefined && Number.isFinite(prev) && val < prev && val >= b) {
        return glyphs[tier] ?? null;
      }
    }
  }
  return glyphs[n - 1] ?? null;
}

export function cellMatchesConditionalFormatRule(
  sheet: Worksheet,
  row: number,
  col: number,
  rule: ConditionalFormatRule,
): boolean {
  if (rule.uiFamily !== "classic") {
    return false;
  }
  if (!selectionRangeContains(rule.range, row, col)) {
    return false;
  }
  switch (rule.classicType) {
    case "iconSet":
      return false;
    case "formula":
      return false;
    case "uniqueOrDuplicate": {
      const kind = rule.uniqueKind ?? "duplicate";
      const counts = getRuleRangeEvalCache(sheet, rule).duplicateKeyCounts;
      if (counts === undefined) {
        return false;
      }
      const k = duplicateKeyForCell(sheet, row, col);
      const cnt = counts.get(k) ?? 0;
      return kind === "duplicate" ? cnt > 1 : cnt === 1;
    }
    case "aboveBelowAverage": {
      const avg = getRuleRangeEvalCache(sheet, rule).rangeAverage;
      if (avg == null) {
        return false;
      }
      const v = sheet.getCell(row, col).value;
      if (typeof v !== "number" || !Number.isFinite(v)) {
        return false;
      }
      const ak = rule.averageKind ?? "above";
      if (ak === "above") {
        return v > avg;
      }
      if (ak === "below") {
        return v < avg;
      }
      if (ak === "equalOrAbove") {
        return v >= avg;
      }
      return v <= avg;
    }
    case "topBottomRanked": {
      const th = getRuleRangeEvalCache(sheet, rule).topBottomThresholdValue;
      if (th == null) {
        return false;
      }
      const kind = rule.topBottomKind ?? "top";
      const v = sheet.getCell(row, col).value;
      if (typeof v !== "number" || !Number.isFinite(v)) {
        return false;
      }
      if (kind === "top" || kind === "topPercent") {
        return v >= th;
      }
      return v <= th;
    }
    case "cellsThatContain": {
      const sub = rule.cellsThatContainKind ?? "cellValue";
      const cell = sheet.getCell(row, col);
      const val = cell.value;
      if (sub === "blanks") {
        return isBlankValue(val);
      }
      if (sub === "noBlanks") {
        return !isBlankValue(val);
      }
      if (sub === "errors") {
        return cellLooksLikeError(sheet, row, col);
      }
      if (sub === "noErrors") {
        return !cellLooksLikeError(sheet, row, col);
      }
      if (sub === "specificText") {
        const disp = formatCellDisplayWithStyle(val, cell.style);
        const op = rule.textOperator ?? "contains";
        return matchesText(op, disp, (rule.value1 ?? "").trim());
      }
      if (sub === "dateOccurring") {
        const serialDay = cellDateSerialDay(val);
        if (serialDay === null) {
          return false;
        }
        return matchesDateOccurring(serialDay, rule.dateOccurring ?? "today");
      }
      const op = rule.valueOperator ?? "greaterThan";
      return compareValues(op, val, rule.value1 ?? "", rule.value2);
    }
    default:
      return false;
  }
}

export function resolveConditionalFormattingOverlay(
  sheet: Worksheet,
  row: number,
  col: number,
): ConditionalFormattingOverlay | null {
  const rules = sheet.getConditionalFormatRules();
  let base: ConditionalFormattingOverlay | null = null;
  let icon: ConditionalFormattingCellIcon | null = null;
  let scaleFillArgb: string | null = null;
  let dataBar: ConditionalFormattingDataBarPaint | null = null;
  for (const rule of rules) {
    if (rule.uiFamily === "iconSet") {
      if (!selectionRangeContains(rule.range, row, col)) {
        continue;
      }
      if (icon !== null) {
        continue;
      }
      const g = resolveIconSetGlyphForCell(sheet, row, col, rule);
      if (g !== null && g !== "none") {
        icon = { glyphId: g, hideCellValue: rule.cfIconShowIconOnly === true };
      } else if (rule.cfIconShowIconOnly === true) {
        icon = { glyphId: "none", hideCellValue: true };
      }
    } else if (rule.uiFamily === "twoColorScale") {
      if (scaleFillArgb === null) {
        const sf = resolveTwoColorScaleFillArgb(sheet, row, col, rule);
        if (sf !== null) {
          scaleFillArgb = sf;
        }
      }
    } else if (rule.uiFamily === "threeColorScale") {
      if (scaleFillArgb === null) {
        const sf = resolveThreeColorScaleFillArgb(sheet, row, col, rule);
        if (sf !== null) {
          scaleFillArgb = sf;
        }
      }
    } else if (rule.uiFamily === "dataBar") {
      if (dataBar === null) {
        const db = resolveDataBarPaintForCell(sheet, row, col, rule);
        if (db !== null) {
          dataBar = db;
        }
      }
    } else if (rule.uiFamily === "classic") {
      if (base !== null) {
        continue;
      }
      if (cellMatchesConditionalFormatRule(sheet, row, col, rule)) {
        base = cfFormatPresetToOverlay(rule.formatPreset, rule.customFormat);
      }
    }
  }
  if (base === null && icon === null && scaleFillArgb === null && dataBar === null) {
    return null;
  }
  const merged: ConditionalFormattingOverlay = {
    ...(base ?? {}),
    ...(icon !== null ? { icon } : {}),
    ...(base?.fillArgb === undefined && scaleFillArgb !== null ? { fillArgb: scaleFillArgb } : {}),
    ...(dataBar !== null ? { dataBar } : {}),
  };
  if (Object.keys(merged).length === 0) {
    return null;
  }
  return merged as ConditionalFormattingOverlay;
}

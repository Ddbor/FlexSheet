import {
  type CellHorizontalAlign,
  type CellScalar,
  type CellStyle,
  type CellStylePatch,
  type CellTextOrientation,
  type CellVerticalAlign,
  formatCellDisplayWithStyle,
} from "@flexsheet/core";
import { RIBBON_NUMBER_FORMAT_PRESETS } from "@flexsheet/toolbar";

import type { FlexSheet } from "./flex-sheet.js";
import { ensureFsSheetPromptStyles } from "./fs-dialog-styles.js";

export type FormatCellsMainTabId =
  | "number"
  | "alignment"
  | "font"
  | "border"
  | "fill"
  | "protection";

export type NumberCategoryId =
  | "general"
  | "number"
  | "currency"
  | "accounting"
  | "date"
  | "time"
  | "percentage"
  | "fraction"
  | "scientific"
  | "text"
  | "special"
  | "custom";

type LocaleId = "zh-CN" | "en-US";

interface FormatCellsNumberState {
  readonly category: NumberCategoryId;
  readonly decimals: number;
  readonly useThousands: boolean;
  readonly negStyle: 0 | 1 | 2 | 3 | 4;
  readonly currencySymbol: "CNY" | "USD";
  readonly currencyNegStyle: 0 | 1 | 2 | 3 | 4;
  readonly accountingSymbol: "CNY" | "USD";
  readonly dateLocale: LocaleId;
  readonly dateTypeIndex: number;
  readonly timeLocale: LocaleId;
  readonly timeTypeIndex: number;
  readonly calendarKind: "gregorian";
  readonly fractionTypeIndex: number;
  readonly scientificDecimals: number;
  readonly percentageDecimals: number;
  readonly specialLocale: LocaleId;
  readonly specialTypeIndex: number;
  readonly customCode: string;
}

const CATEGORY_LABELS: readonly { readonly id: NumberCategoryId; readonly label: string }[] = [
  { id: "general", label: "常规" },
  { id: "number", label: "数值" },
  { id: "currency", label: "货币" },
  { id: "accounting", label: "会计专用" },
  { id: "date", label: "日期" },
  { id: "time", label: "时间" },
  { id: "percentage", label: "百分比" },
  { id: "fraction", label: "分数" },
  { id: "scientific", label: "科学记数" },
  { id: "text", label: "文本" },
  { id: "special", label: "特殊" },
  { id: "custom", label: "自定义" },
];

const MAIN_TAB_LABELS: readonly { readonly id: FormatCellsMainTabId; readonly label: string }[] = [
  { id: "number", label: "数字" },
  { id: "alignment", label: "对齐" },
  { id: "font", label: "字体" },
  { id: "border", label: "边框" },
  { id: "fill", label: "填充" },
  { id: "protection", label: "保护" },
];

const LOCALE_OPTIONS: readonly { readonly id: LocaleId; readonly label: string }[] = [
  { id: "zh-CN", label: "简体中文 (中国大陆)" },
  { id: "en-US", label: "英语 (美国)" },
];

const CURRENCY_SYMBOL_OPTIONS: readonly { readonly id: "CNY" | "USD"; readonly label: string }[] = [
  { id: "CNY", label: "¥" },
  { id: "USD", label: "$" },
];

const DATE_TYPES_ZH: readonly { readonly label: string; readonly code: string }[] = [
  { label: "*2012/3/14", code: "*yyyy/m/d" },
  { label: "*2012年3月14日 星期三", code: '*yyyy"年"m"月"d"日"' },
  { label: "2012-03-14", code: "yyyy-mm-dd" },
  { label: "二〇一二年三月十四日", code: 'yyyy"年"m"月"d"日"' },
  { label: "二〇一二年三月", code: 'yyyy"年"m"月"' },
  { label: "三月十四日", code: 'm"月"d"日"' },
  { label: "2012年3月14日", code: 'yyyy"年"m"月"d"日"' },
  { label: "2012年3月", code: 'yyyy"年"m"月"' },
];

const DATE_TYPES_EN: readonly { readonly label: string; readonly code: string }[] = [
  { label: "*3/14/2012", code: "*m/d/yyyy" },
  { label: "*Wednesday, March 14, 2012", code: '*dddd, mmmm d, yyyy' },
  { label: "3/14/2012", code: "m/d/yyyy" },
  { label: "3/14/12", code: "m/d/yy" },
  { label: "03/14/12", code: "mm/dd/yy" },
  { label: "3-14-12", code: "m-d-yy" },
  { label: "Mar-12", code: "mmm-yy" },
  { label: "March-12", code: "mmmm-yy" },
];

const TIME_TYPES_ZH: readonly { readonly label: string; readonly code: string }[] = [
  { label: "*13:30:55", code: "*h:mm:ss" },
  { label: "13:30", code: "h:mm" },
  { label: "1:30 PM", code: "h:mm AM/PM" },
  { label: "13:30:55", code: "h:mm:ss" },
  { label: "1:30:55 PM", code: "h:mm:ss AM/PM" },
  { label: "13时30分", code: 'h"时"mm"分"' },
  { label: "13时30分55秒", code: 'h"时"mm"分"ss"秒"' },
  { label: "下午1时30分", code: '上午/下午h"时"mm"分"' },
];

const TIME_TYPES_EN: readonly { readonly label: string; readonly code: string }[] = [
  { label: "*1:30:55 PM", code: "*h:mm:ss AM/PM" },
  { label: "1:30 PM", code: "h:mm AM/PM" },
  { label: "13:30", code: "h:mm" },
  { label: "1:30:55 PM", code: "h:mm:ss AM/PM" },
  { label: "13:30:55", code: "h:mm:ss" },
  { label: "30:55.0", code: "mm:ss.0" },
  { label: "[h]:mm:ss", code: "[h]:mm:ss" },
];

const FRACTION_TYPES: readonly { readonly label: string; readonly code: string }[] = [
  { label: "分母为一位数 (1/4)", code: "# ?/?" },
  { label: "分母为两位数 (21/25)", code: "# ??/??" },
  { label: "分母为三位数 (312/943)", code: "# ???/???" },
  { label: "以 2 为分母 (1/2)", code: "# ?/2" },
  { label: "以 4 为分母 (2/4)", code: "# ?/4" },
  { label: "以 8 为分母 (4/8)", code: "# ?/8" },
  { label: "以 16 为分母 (8/16)", code: "# ?/16" },
  { label: "以 10 为分母 (3/10)", code: "# ?/10" },
];

const SPECIAL_TYPES_ZH: readonly { readonly label: string; readonly code: string }[] = [
  { label: "邮政编码", code: "000000" },
  { label: "中文小写数字", code: "[DBNum1]General" },
  { label: "中文大写数字", code: "[DBNum2]General" },
];

const SPECIAL_TYPES_EN: readonly { readonly label: string; readonly code: string }[] = [
  { label: "Zip Code", code: "00000" },
];

const CATEGORY_HELP: Record<NumberCategoryId, string> = {
  general: "常规单元格格式不包含任何特定的数字格式。",
  number:
    "数值格式用于一般数字的表示。货币和会计格式则提供货币值计算的专用格式。",
  currency: "货币格式用于表示一般货币数值。会计格式可以对一列数值进行小数点对齐。",
  accounting: "会计格式可对一列数值进行货币符号和小数点对齐。",
  date: "日期格式将日期和时间系列数值显示为日期值。以星号 (*) 开头的日期格式响应操作系统特定的区域日期设置的更改。不带星号的格式不受操作系统设置的影响。",
  time: "时间格式将日期和时间系列数值显示为时间值。以星号 (*) 开头的时间格式响应操作系统特定的区域日期和时间设置的更改。不带星号的格式不受操作系统设置的影响。",
  percentage: "百分比格式将单元格中数值乘以 100，并以百分数形式显示。",
  fraction: "分数格式以分数形式显示小数。",
  scientific: "科学记数格式用指数显示数字。",
  text: "在文本单元格格式中，数字作为文本处理。单元格显示的内容与输入的内容完全一致。",
  special: "特殊格式可用于跟踪数据列表及数据库的值。",
  custom: "以现有格式为基础，生成自定义的数字格式。",
};

const PLACEHOLDER_TAB_TEXT =
  "此选项卡尚未实现，后续版本将提供字体、边框、填充与保护等选项。";

const ALIGNMENT_TAB_DESC =
  "设置单元格内文本的水平与垂直对齐、缩进、方向，以及自动换行、缩小字体填充等选项。";

type AlignmentHorizontalUi =
  | "general"
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerContinuous"
  | "distributed";

interface FormatCellsAlignmentState {
  horizontal: AlignmentHorizontalUi;
  vertical: CellVerticalAlign;
  indent: number;
  wrap: boolean;
  shrink: boolean;
  mergeCells: boolean;
  orientationVertical: boolean;
  degrees: number;
}

const HORIZONTAL_OPTIONS: readonly { readonly id: AlignmentHorizontalUi; readonly label: string }[] =
  [
    { id: "general", label: "常规" },
    { id: "left", label: "靠左" },
    { id: "center", label: "居中" },
    { id: "right", label: "靠右" },
    { id: "fill", label: "填充" },
    { id: "justify", label: "两端对齐" },
    { id: "centerContinuous", label: "跨列居中" },
    { id: "distributed", label: "分散对齐" },
  ];

const VERTICAL_OPTIONS: readonly { readonly id: CellVerticalAlign; readonly label: string }[] = [
  { id: "top", label: "靠上" },
  { id: "middle", label: "居中" },
  { id: "bottom", label: "靠下" },
  { id: "justify", label: "两端对齐" },
  { id: "distributed", label: "分散对齐" },
];

function orientationToDegrees(orient: CellTextOrientation | undefined): number {
  switch (orient) {
    case "angleUp45":
      return 45;
    case "angleDown45":
      return -45;
    case "rotateUp90":
      return 90;
    case "rotateDown90":
      return -90;
    default:
      return 0;
  }
}

function inferAlignmentState(style: CellStyle | null, flex: FlexSheet): FormatCellsAlignmentState {
  const sheet = flex.workbook.getActiveSheet();
  let merge = false;
  if (sheet !== undefined) {
    const { row, col } = flex.selection.getActiveCell();
    const info = sheet.getMergedRectInfo(row, col);
    merge = info.rowSpan > 1 || info.colSpan > 1;
  }
  const h = style?.hAlign;
  let horizontal: AlignmentHorizontalUi = "general";
  if (
    h === "left" ||
    h === "center" ||
    h === "right" ||
    h === "fill" ||
    h === "justify" ||
    h === "distributed" ||
    h === "centerContinuous"
  ) {
    horizontal = h;
  }
  const v: CellVerticalAlign = style?.vAlign ?? "middle";
  const indent = style?.indentLevel ?? 0;
  const wrap = style?.wrapText === true;
  const shrink = style?.shrinkToFit === true;
  const orient = style?.textOrientation ?? "horizontal";
  const trd = style?.textRotationDegrees;
  let orientationVertical = orient === "verticalStack";
  let degrees = 0;
  if (trd !== undefined && Number.isFinite(trd) && trd !== 0) {
    degrees = Math.max(-90, Math.min(90, Math.round(trd)));
    orientationVertical = false;
  } else if (!orientationVertical && orient !== "horizontal") {
    degrees = orientationToDegrees(orient);
  }
  return {
    horizontal,
    vertical: v,
    indent,
    wrap,
    shrink,
    mergeCells: merge,
    orientationVertical,
    degrees,
  };
}

function alignmentPatchIfChanged(
  cur: FormatCellsAlignmentState,
  ini: FormatCellsAlignmentState,
): CellStylePatch {
  const p = {} as CellStylePatch;
  if (cur.horizontal !== ini.horizontal) {
    p.hAlign = cur.horizontal === "general" ? null : (cur.horizontal as CellHorizontalAlign);
  }
  if (cur.vertical !== ini.vertical) {
    p.vAlign = cur.vertical;
  }
  if (cur.indent !== ini.indent) {
    p.indentLevel = cur.indent <= 0 ? null : Math.min(255, Math.round(cur.indent));
  }
  if (cur.wrap !== ini.wrap) {
    p.wrapText = cur.wrap ? true : null;
  }
  if (cur.shrink !== ini.shrink) {
    p.shrinkToFit = cur.shrink ? true : null;
  }
  const orientSame =
    cur.orientationVertical === ini.orientationVertical && cur.degrees === ini.degrees;
  if (!orientSame) {
    if (cur.orientationVertical) {
      p.textOrientation = "verticalStack";
      p.textRotationDegrees = null;
    } else if (cur.degrees !== 0) {
      p.textOrientation = "horizontal";
      p.textRotationDegrees = cur.degrees;
    } else {
      p.textOrientation = null;
      p.textRotationDegrees = null;
    }
  }
  return p;
}

/**
 * 将指针映射到文字转角：半圆在圆心右侧，上为 +90°、右为 0°、下为 −90°（与 Excel 对话框一致）。
 * `arcCx, arcCy` 为圆弧所在圆的圆心；指针为从圆心到弧上一点的半径。
 */
function textRotationDegFromGaugePointer(
  clientX: number,
  clientY: number,
  svg: SVGSVGElement,
  arcCx: number,
  arcCy: number,
): number {
  const pt = svg.createSVGPoint();
  pt.x = clientX;
  pt.y = clientY;
  const ctm = svg.getScreenCTM();
  if (!ctm) {
    return 0;
  }
  const loc = pt.matrixTransform(ctm.inverse());
  const vx = loc.x - arcCx;
  const vy = loc.y - arcCy;
  let phi = Math.atan2(vy, vx);
  if (vx < 0) {
    phi = vy >= 0 ? Math.PI / 2 : -Math.PI / 2;
  } else {
    phi = Math.max(-Math.PI / 2, Math.min(Math.PI / 2, phi));
  }
  const raw = -(phi * 180) / Math.PI;
  if (!Number.isFinite(raw)) {
    return 0;
  }
  return Math.max(-90, Math.min(90, Math.round(raw)));
}

let formatCellsStylesInjected = false;

function ensureFormatCellsDialogStyles(): void {
  ensureFsSheetPromptStyles();
  if (formatCellsStylesInjected) {
    return;
  }
  formatCellsStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-format-cells", "1");
  style.textContent = `
.fs-format-cells-overlay {
  position: fixed;
  inset: 0;
  z-index: 10002;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(0, 0, 0, 0.35);
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
}
.fs-format-cells {
  width: min(560px, calc(100vw - 24px));
  max-height: min(520px, calc(100vh - 24px));
  box-sizing: border-box;
  background: #f3f2f1;
  border-radius: 10px;
  box-shadow: 0 10px 40px rgba(0, 0, 0, 0.22);
  display: flex;
  flex-direction: column;
  overflow: hidden;
}
.fs-format-cells__header {
  position: relative;
  padding: 12px 40px 10px 14px;
  background: #f3f2f1;
  border-bottom: 1px solid #e1dfdd;
}
.fs-format-cells__title {
  font-size: 15px;
  font-weight: 600;
  color: #323130;
  text-align: center;
}
.fs-format-cells__close {
  position: absolute;
  right: 8px;
  top: 6px;
  width: 32px;
  height: 32px;
  border: none;
  background: transparent;
  font-size: 20px;
  line-height: 1;
  color: #605e5c;
  cursor: pointer;
  border-radius: 4px;
}
.fs-format-cells__close:hover {
  background: #edebe9;
  color: #323130;
}
.fs-format-cells__tabs {
  display: flex;
  gap: 0;
  padding: 0 10px;
  background: #edebe9;
  border-bottom: 1px solid #d2d0ce;
}
.fs-format-cells__tab {
  flex: 1;
  min-width: 0;
  padding: 8px 4px;
  border: none;
  border-radius: 6px 6px 0 0;
  background: transparent;
  font-size: 13px;
  color: #323130;
  cursor: pointer;
  font-family: inherit;
}
.fs-format-cells__tab:hover {
  background: rgba(255, 255, 255, 0.45);
}
.fs-format-cells__tab--active {
  background: #fff;
  font-weight: 600;
  box-shadow: 0 -1px 0 #fff;
}
.fs-format-cells__body {
  flex: 1;
  min-height: 280px;
  overflow: auto;
  background: #fff;
  padding: 12px 14px;
}
.fs-format-cells__placeholder {
  font-size: 13px;
  color: #605e5c;
  line-height: 1.5;
  padding: 24px 8px;
}
.fs-format-cells__number-layout {
  display: flex;
  gap: 12px;
  align-items: stretch;
  min-height: 240px;
}
.fs-format-cells__cat-col {
  display: flex;
  flex-direction: column;
  gap: 4px;
  width: 132px;
  flex-shrink: 0;
}
.fs-format-cells__cat-label {
  font-size: 12px;
  color: #323130;
  font-weight: 600;
}
.fs-format-cells__cat-list {
  flex: 1;
  list-style: none;
  margin: 0;
  padding: 2px 0;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  overflow-y: auto;
  max-height: 320px;
}
.fs-format-cells__cat-item {
  width: 100%;
  padding: 5px 10px;
  border: none;
  background: transparent;
  text-align: left;
  font-size: 13px;
  cursor: pointer;
  font-family: inherit;
  color: #323130;
}
.fs-format-cells__cat-item:hover:not(.fs-format-cells__cat-item--active) {
  background: #e8f5e9;
  color: #201f1e;
}
.fs-format-cells__cat-item--active {
  background: #217346;
  color: #fff;
  font-weight: 500;
}
.fs-format-cells__cat-item--active:hover {
  background: #1a5c38;
  color: #fff;
}
.fs-format-cells__detail {
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.fs-format-cells__row {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #323130;
}
.fs-format-cells__row label {
  min-width: 88px;
  flex-shrink: 0;
}
.fs-format-cells__sample {
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #f3f2f1;
  min-height: 28px;
  padding: 4px 10px;
  font-size: 13px;
  color: #201f1e;
}
.fs-format-cells__sample--red {
  color: #c50f1f;
}
.fs-format-cells__input-num {
  width: 72px;
  padding: 4px 6px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  font-size: 13px;
}
.fs-format-cells__checkbox {
  accent-color: #217346;
}
.fs-format-cells__select {
  flex: 1;
  min-width: 0;
  padding: 5px 8px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  font-size: 13px;
  font-family: inherit;
  background: #fff;
}
.fs-format-cells__listbox {
  flex: 1;
  min-height: 120px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  overflow-y: auto;
  background: #fff;
}
.fs-format-cells__listbox button {
  display: block;
  width: 100%;
  padding: 5px 8px;
  border: none;
  background: transparent;
  text-align: left;
  font-size: 12px;
  cursor: pointer;
  font-family: inherit;
}
.fs-format-cells__listbox button:hover {
  background: #e8f5e9;
}
.fs-format-cells__listbox button.fs-format-cells__list-sel {
  background: #c8e6c9;
}
.fs-format-cells__listbox--neg button.fs-format-cells__list-sel {
  background: #c8e6c9;
}
.fs-format-cells__text-input {
  width: 100%;
  box-sizing: border-box;
  padding: 6px 8px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  font-size: 13px;
  font-family: ui-monospace, monospace;
}
.fs-format-cells__custom-list {
  min-height: 140px;
  max-height: 180px;
}
.fs-format-cells__footer {
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding: 8px 14px 14px 14px;
  background: #fff;
  border-top: 1px solid #edebe9;
}
.fs-format-cells__desc {
  font-size: 12px;
  color: #605e5c;
  line-height: 1.45;
  min-height: 36px;
}
.fs-format-cells__actions {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
}
.fs-format-cells__btn {
  min-width: 76px;
  padding: 7px 16px;
  font-size: 13px;
  border-radius: 4px;
  cursor: pointer;
  font-family: inherit;
}
.fs-format-cells__btn--secondary {
  border: 1px solid #c8c6c4;
  background: #fff;
  color: #323130;
}
.fs-format-cells__btn--secondary:hover {
  background: #f3f2f1;
}
.fs-format-cells__btn--primary {
  border: none;
  background: #217346;
  color: #fff;
  font-weight: 500;
}
.fs-format-cells__btn--primary:hover {
  background: #1a5c38;
}
.fs-format-cells__align-outer {
  display: flex;
  flex-direction: column;
  gap: 0;
  min-height: 280px;
}
.fs-format-cells__align-grid {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 1px 220px;
  gap: 14px 16px;
  align-items: start;
}
.fs-format-cells__align-left {
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 14px;
}
.fs-format-cells__align-distributed-row {
  margin-top: 4px;
}
.fs-format-cells__align-sep {
  width: 1px;
  align-self: stretch;
  min-height: 160px;
  background: #d2d0ce;
}
.fs-format-cells__align-right {
  min-width: 0;
}
.fs-format-cells__orient-row {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.fs-format-cells__orient-label {
  font-size: 12px;
  font-weight: 600;
  color: #323130;
}
.fs-format-cells__orient-preview {
  display: flex;
  gap: 10px;
  align-items: stretch;
}
.fs-format-cells__orient-vertical {
  width: 40px;
  min-height: 100px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  display: flex;
  align-items: center;
  justify-content: center;
  background: #fff;
  cursor: pointer;
  font-size: 12px;
  color: #323130;
  writing-mode: vertical-rl;
  text-orientation: mixed;
  user-select: none;
}
.fs-format-cells__orient-vertical:hover {
  background: #e8f5e9;
}
.fs-format-cells__orient-vertical--on {
  outline: 2px solid #217346;
  outline-offset: 1px;
}
.fs-format-cells__orient-gauge {
  flex: 1;
  min-height: 118px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: flex-start;
  padding: 6px 6px 8px;
  gap: 6px;
}
.fs-format-cells__orient-gauge-surface {
  width: 100%;
  max-width: 200px;
  touch-action: none;
  cursor: grab;
  user-select: none;
}
.fs-format-cells__orient-gauge-surface--dragging {
  cursor: grabbing;
}
.fs-format-cells__orient-gauge svg {
  display: block;
  width: 100%;
  height: auto;
  max-height: 86px;
}
.fs-format-cells__control-block {
  margin-top: 18px;
  padding-top: 14px;
  border-top: 1px solid #edebe9;
}
.fs-format-cells__control-title {
  font-size: 12px;
  font-weight: 600;
  color: #323130;
  margin-bottom: 10px;
}
.fs-format-cells__checkbox-row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin: 8px 0;
  font-size: 13px;
  color: #323130;
}
.fs-format-cells__checkbox-row input:disabled + span {
  color: #a19f9d;
}
`;
  document.head.appendChild(style);
}

function readActiveScalar(flex: FlexSheet): CellScalar {
  const sheet = flex.workbook.getActiveSheet();
  if (sheet === undefined) {
    return 1234.1;
  }
  const { row, col } = flex.selection.getActiveCell();
  const a = sheet.getMergeAnchorCell(row, col);
  return sheet.getCell(a.row, a.col).value;
}

function sampleNumberForPreview(scalar: CellScalar): number {
  if (typeof scalar === "number" && Number.isFinite(scalar)) {
    return scalar;
  }
  if (typeof scalar === "string" && scalar.trim() !== "") {
    const n = Number(scalar);
    if (Number.isFinite(n)) {
      return n;
    }
  }
  return 1234.1;
}

function posNumberPattern(decimals: number, useThousands: boolean): string {
  const dec = decimals > 0 ? `.${"0".repeat(decimals)}` : "";
  if (useThousands) {
    return decimals > 0 ? `#,##0${dec}` : "#,##0";
  }
  return decimals > 0 ? `0${dec}` : "0";
}

function withNegatives(pos: string, style: 0 | 1 | 2 | 3 | 4): string {
  switch (style) {
    case 0:
      return `${pos};[Red](${pos})`;
    case 1:
      return `${pos};(${pos})`;
    case 2:
      return `${pos};[Red]${pos}`;
    case 3:
      return `${pos};-${pos}`;
    case 4:
      return `${pos};[Red]-${pos}`;
    default:
      return pos;
  }
}

function currencyChar(sym: "CNY" | "USD"): string {
  return sym === "CNY" ? "¥" : "$";
}

function posCurrencyPattern(sym: "CNY" | "USD", decimals: number, useThousands: boolean): string {
  const c = currencyChar(sym);
  const core = posNumberPattern(decimals, useThousands);
  return `${c}${core}`;
}

function buildFormatFromNumberState(s: FormatCellsNumberState): string | null {
  switch (s.category) {
    case "general":
      return null;
    case "number": {
      const pos = posNumberPattern(s.decimals, s.useThousands);
      return withNegatives(pos, s.negStyle);
    }
    case "currency": {
      const pos = posCurrencyPattern(s.currencySymbol, s.decimals, true);
      return withNegatives(pos, s.currencyNegStyle);
    }
    case "accounting": {
      const pos = posCurrencyPattern(s.accountingSymbol, s.decimals, true);
      return pos;
    }
    case "date": {
      const list = s.dateLocale === "zh-CN" ? DATE_TYPES_ZH : DATE_TYPES_EN;
      const row = list[s.dateTypeIndex] ?? list[0];
      return row?.code ?? "yyyy/m/d";
    }
    case "time": {
      const list = s.timeLocale === "zh-CN" ? TIME_TYPES_ZH : TIME_TYPES_EN;
      const row = list[s.timeTypeIndex] ?? list[0];
      return row?.code ?? "h:mm:ss";
    }
    case "percentage": {
      const dec = s.percentageDecimals > 0 ? `.${"0".repeat(s.percentageDecimals)}` : "";
      return `0${dec}%`;
    }
    case "fraction": {
      return FRACTION_TYPES[s.fractionTypeIndex]?.code ?? FRACTION_TYPES[0]!.code;
    }
    case "scientific": {
      const d = Math.max(2, s.scientificDecimals);
      return `0.${"0".repeat(d)}E+00`;
    }
    case "text":
      return "@";
    case "special": {
      const list =
        s.specialLocale === "zh-CN" ? SPECIAL_TYPES_ZH : SPECIAL_TYPES_EN;
      return list[s.specialTypeIndex]?.code ?? list[0]!.code;
    }
    case "custom": {
      const t = s.customCode.trim();
      if (t === "" || t === "General" || t === "G/通用格式") {
        return null;
      }
      return t;
    }
    default:
      return null;
  }
}

function inferInitialNumberState(nfRaw: string | undefined): FormatCellsNumberState {
  const nf = nfRaw?.trim() ?? "";
  const base: FormatCellsNumberState = {
    category: "general",
    decimals: 2,
    useThousands: false,
    negStyle: 3,
    currencySymbol: "CNY",
    currencyNegStyle: 3,
    accountingSymbol: "CNY",
    dateLocale: "zh-CN",
    dateTypeIndex: 0,
    timeLocale: "zh-CN",
    timeTypeIndex: 0,
    calendarKind: "gregorian",
    fractionTypeIndex: 0,
    scientificDecimals: 2,
    percentageDecimals: 2,
    specialLocale: "zh-CN",
    specialTypeIndex: 0,
    customCode: "G/通用格式",
  };
  if (nf === "" || nf === "General") {
    return { ...base, category: "general" };
  }
  if (nf === "@") {
    return { ...base, category: "text" };
  }
  if (/DBNum1/i.test(nf)) {
    return { ...base, category: "special", specialLocale: "zh-CN", specialTypeIndex: 1 };
  }
  if (/DBNum2/i.test(nf)) {
    return { ...base, category: "special", specialLocale: "zh-CN", specialTypeIndex: 2 };
  }
  if (/^0{5,6}$/.test(nf)) {
    return {
      ...base,
      category: "special",
      specialLocale: nf.length >= 6 ? "zh-CN" : "en-US",
      specialTypeIndex: 0,
    };
  }
  if (nf.includes("?") && nf.includes("/") && nf.includes("#")) {
    const i = FRACTION_TYPES.findIndex((t) => t.code === nf);
    return { ...base, category: "fraction", fractionTypeIndex: i >= 0 ? i : 0 };
  }
  if (/e\+00/i.test(nf) || /0e\+/i.test(nf)) {
    const m = nf.match(/\.(0+)e/i);
    const d = m !== null && m[1] !== undefined ? m[1].length : 2;
    return { ...base, category: "scientific", scientificDecimals: d };
  }
  if (nf.includes("%")) {
    const m = nf.match(/\.(0+)%/);
    const d = m !== null && m[1] !== undefined ? m[1].length : 0;
    return { ...base, category: "percentage", percentageDecimals: d };
  }
  const dateIdxZh = DATE_TYPES_ZH.findIndex((t) => t.code === nf);
  if (dateIdxZh >= 0) {
    return { ...base, category: "date", dateLocale: "zh-CN", dateTypeIndex: dateIdxZh };
  }
  const dateIdxEn = DATE_TYPES_EN.findIndex((t) => t.code === nf);
  if (dateIdxEn >= 0) {
    return { ...base, category: "date", dateLocale: "en-US", dateTypeIndex: dateIdxEn };
  }
  const timeIdxZh = TIME_TYPES_ZH.findIndex((t) => t.code === nf);
  if (timeIdxZh >= 0) {
    return { ...base, category: "time", timeLocale: "zh-CN", timeTypeIndex: timeIdxZh };
  }
  const timeIdxEn = TIME_TYPES_EN.findIndex((t) => t.code === nf);
  if (timeIdxEn >= 0) {
    return { ...base, category: "time", timeLocale: "en-US", timeTypeIndex: timeIdxEn };
  }
  const preset = RIBBON_NUMBER_FORMAT_PRESETS.find((p) => p.format === nf);
  if (preset !== undefined) {
    if (preset.id === "home.number.format.accounting") {
      return { ...base, category: "accounting", accountingSymbol: "CNY" };
    }
    if (preset.id === "home.number.format.currency") {
      return { ...base, category: "currency", currencySymbol: "CNY" };
    }
    if (preset.id === "home.number.format.number") {
      return { ...base, category: "number", decimals: 2, useThousands: false };
    }
  }
  if ((nf.includes("¥") || nf.includes("￥")) && !nf.includes(";")) {
    return { ...base, category: "currency", currencySymbol: "CNY", decimals: 2 };
  }
  if (nf.includes("$") && !nf.includes(";")) {
    return { ...base, category: "currency", currencySymbol: "USD", decimals: 2 };
  }
  if (/^0\.0+$/.test(nf) || nf === "0") {
    const d = nf === "0" ? 0 : nf.length - nf.indexOf(".") - 1;
    return { ...base, category: "number", decimals: d, useThousands: false };
  }
  return { ...base, category: "custom", customCode: nf === "" ? "G/通用格式" : nf };
}

function customFormatBuiltinList(): readonly string[] {
  const fromRibbon = RIBBON_NUMBER_FORMAT_PRESETS.map((p) => p.format).filter((f) => f !== "");
  const extra = [
    "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
    "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
    "#,##0_);(#,##0)",
    "#,##0.00_);(#,##0.00)",
    "[Red]#,##0",
  ];
  return ["G/通用格式", "0", "0.00", ...fromRibbon, ...extra];
}

function previewDisplay(sample: number, formatCode: string | null): { readonly text: string; readonly red: boolean } {
  if (formatCode === null || formatCode === "" || formatCode === "General") {
    return { text: String(sample), red: false };
  }
  const style: CellStyle = { numberFormat: formatCode };
  const neg = sample < 0 ? sample : -Math.abs(sample);
  const use = /;/.test(formatCode) ? neg : sample;
  const t = formatCellDisplayWithStyle(use, style);
  const red = /\[Red\]/i.test(formatCode) && use < 0;
  return { text: t, red };
}

export interface MountFormatCellsDialogOptions {
  readonly flex: FlexSheet;
  readonly onClose?: () => void;
}

/** 挂载「设置单元格格式」对话框，返回 overlay 根节点（供与右键菜单 prompt 生命周期对齐）。 */
export function mountFormatCellsDialog(options: MountFormatCellsDialogOptions): HTMLDivElement {
  ensureFormatCellsDialogStyles();
  const { flex, onClose } = options;
  const initialNf = flex.getActiveCellStyle()?.numberFormat ?? undefined;
  let mainTab: FormatCellsMainTabId = "number";
  let numState = inferInitialNumberState(initialNf);

  const overlay = document.createElement("div");
  overlay.className = "fs-format-cells-overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-format-cells";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-format-cells-title");

  const header = document.createElement("div");
  header.className = "fs-format-cells__header";
  const titleEl = document.createElement("div");
  titleEl.id = "fs-format-cells-title";
  titleEl.className = "fs-format-cells__title";
  titleEl.textContent = "设置单元格格式";
  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "fs-format-cells__close";
  closeBtn.setAttribute("aria-label", "关闭");
  closeBtn.textContent = "×";
  header.appendChild(titleEl);
  header.appendChild(closeBtn);

  const tabsRow = document.createElement("div");
  tabsRow.className = "fs-format-cells__tabs";
  const tabButtons = new Map<FormatCellsMainTabId, HTMLButtonElement>();
  for (const t of MAIN_TAB_LABELS) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "fs-format-cells__tab";
    b.textContent = t.label;
    b.dataset.tab = t.id;
    tabButtons.set(t.id, b);
    tabsRow.appendChild(b);
  }

  const body = document.createElement("div");
  body.className = "fs-format-cells__body";

  const footer = document.createElement("div");
  footer.className = "fs-format-cells__footer";
  const descEl = document.createElement("div");
  descEl.className = "fs-format-cells__desc";
  const actions = document.createElement("div");
  actions.className = "fs-format-cells__actions";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "fs-format-cells__btn fs-format-cells__btn--secondary";
  cancelBtn.textContent = "取消";
  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.className = "fs-format-cells__btn fs-format-cells__btn--primary";
  okBtn.textContent = "确定";
  actions.appendChild(cancelBtn);
  actions.appendChild(okBtn);
  footer.appendChild(descEl);
  footer.appendChild(actions);

  panel.appendChild(header);
  panel.appendChild(tabsRow);
  panel.appendChild(body);
  panel.appendChild(footer);
  overlay.appendChild(panel);
  document.body.appendChild(overlay);

  const scalar = readActiveScalar(flex);
  const sampleNum = sampleNumberForPreview(scalar);

  const initialAlignState = inferAlignmentState(flex.getActiveCellStyle(), flex);
  let alignState: FormatCellsAlignmentState = { ...initialAlignState };

  const numberLayout = document.createElement("div");
  numberLayout.className = "fs-format-cells__number-layout";

  const catCol = document.createElement("div");
  catCol.className = "fs-format-cells__cat-col";
  const catLabel = document.createElement("div");
  catLabel.className = "fs-format-cells__cat-label";
  catLabel.textContent = "类别:";
  const catList = document.createElement("div");
  catList.className = "fs-format-cells__cat-list";
  catList.setAttribute("role", "listbox");
  const catItemBtns = new Map<NumberCategoryId, HTMLButtonElement>();
  for (const c of CATEGORY_LABELS) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "fs-format-cells__cat-item";
    btn.textContent = c.label;
    btn.dataset.cat = c.id;
    btn.setAttribute("role", "option");
    catItemBtns.set(c.id, btn);
    catList.appendChild(btn);
  }
  catCol.appendChild(catLabel);
  catCol.appendChild(catList);

  const detail = document.createElement("div");
  detail.className = "fs-format-cells__detail";

  numberLayout.appendChild(catCol);
  numberLayout.appendChild(detail);

  const placeholderEl = document.createElement("div");
  placeholderEl.className = "fs-format-cells__placeholder";
  placeholderEl.textContent = PLACEHOLDER_TAB_TEXT;

  const syncCategoryHighlight = (): void => {
    for (const [id, b] of catItemBtns) {
      b.classList.toggle("fs-format-cells__cat-item--active", id === numState.category);
    }
  };

  const syncTabHighlight = (): void => {
    for (const [id, b] of tabButtons) {
      b.classList.toggle("fs-format-cells__tab--active", id === mainTab);
    }
  };

  const updateDesc = (): void => {
    if (mainTab === "number") {
      descEl.textContent = CATEGORY_HELP[numState.category] ?? "";
      return;
    }
    if (mainTab === "alignment") {
      descEl.textContent = ALIGNMENT_TAB_DESC;
      return;
    }
    descEl.textContent = "";
  };

  const updateSample = (): void => {
    const sampleEl = detail.querySelector("[data-role='sample']");
    if (!(sampleEl instanceof HTMLElement)) {
      return;
    }
    const code = buildFormatFromNumberState(numState);
    const pv = previewDisplay(sampleNum, code);
    sampleEl.textContent = pv.text;
    sampleEl.classList.toggle("fs-format-cells__sample--red", pv.red);
  };

  const rebuildNumberDetail = (): void => {
    detail.replaceChildren();
    const sampleRow = document.createElement("div");
    sampleRow.className = "fs-format-cells__row";
    const sl = document.createElement("label");
    sl.textContent = "示例:";
    const sampleBox = document.createElement("div");
    sampleBox.className = "fs-format-cells__sample";
    sampleBox.dataset.role = "sample";
    sampleRow.appendChild(sl);
    sampleRow.appendChild(sampleBox);
    detail.appendChild(sampleRow);

    const cat = numState.category;

    const addDecimals = (): void => {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      const lab = document.createElement("label");
      lab.textContent = "小数位数:";
      const inp = document.createElement("input");
      inp.type = "number";
      inp.className = "fs-format-cells__input-num";
      inp.min = "0";
      inp.max = "15";
      inp.value = String(numState.decimals);
      inp.addEventListener("change", () => {
        const n = Math.max(0, Math.min(15, Math.floor(Number(inp.value) || 0)));
        numState = { ...numState, decimals: n };
        inp.value = String(n);
        updateSample();
      });
      row.appendChild(lab);
      row.appendChild(inp);
      detail.appendChild(row);
    };

    const addThousands = (): void => {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "fs-format-cells__checkbox";
      cb.checked = numState.useThousands;
      cb.addEventListener("change", () => {
        numState = { ...numState, useThousands: cb.checked };
        updateSample();
      });
      const span = document.createElement("span");
      span.textContent = "使用千位分隔符(,)";
      row.appendChild(cb);
      row.appendChild(span);
      detail.appendChild(row);
    };

    const addSymbolSelect = (field: "currencySymbol" | "accountingSymbol"): void => {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      const lab = document.createElement("label");
      lab.textContent = "符号:";
      const sel = document.createElement("select");
      sel.className = "fs-format-cells__select";
      for (const o of CURRENCY_SYMBOL_OPTIONS) {
        const opt = document.createElement("option");
        opt.value = o.id;
        opt.textContent = o.label;
        sel.appendChild(opt);
      }
      sel.value = numState[field];
      sel.addEventListener("change", () => {
        const v = sel.value === "USD" ? "USD" : "CNY";
        numState = { ...numState, [field]: v } as FormatCellsNumberState;
        updateSample();
      });
      row.appendChild(lab);
      row.appendChild(sel);
      detail.appendChild(row);
    };

    const addLocaleSelect = (
      field: "dateLocale" | "timeLocale" | "specialLocale",
      onChange: () => void,
    ): void => {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      const lab = document.createElement("label");
      lab.textContent = "语言(位置):";
      const sel = document.createElement("select");
      sel.className = "fs-format-cells__select";
      for (const o of LOCALE_OPTIONS) {
        const opt = document.createElement("option");
        opt.value = o.id;
        opt.textContent = o.label;
        sel.appendChild(opt);
      }
      sel.value = numState[field];
      sel.addEventListener("change", () => {
        const v = sel.value === "en-US" ? "en-US" : "zh-CN";
        numState = { ...numState, [field]: v } as FormatCellsNumberState;
        onChange();
      });
      row.appendChild(lab);
      row.appendChild(sel);
      detail.appendChild(row);
    };

    const addTypeList = (
      items: readonly { readonly label: string; readonly code: string }[],
      selectedIndex: number,
      onPick: (index: number) => void,
    ): void => {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      row.style.alignItems = "flex-start";
      const lab = document.createElement("label");
      lab.textContent = "类型:";
      const box = document.createElement("div");
      box.className = "fs-format-cells__listbox";
      box.style.flex = "1";
      items.forEach((it, idx) => {
        const b = document.createElement("button");
        b.type = "button";
        b.textContent = it.label;
        if (idx === selectedIndex) {
          b.classList.add("fs-format-cells__list-sel");
        }
        b.addEventListener("click", () => {
          for (const c of box.querySelectorAll("button")) {
            c.classList.remove("fs-format-cells__list-sel");
          }
          b.classList.add("fs-format-cells__list-sel");
          onPick(idx);
        });
        box.appendChild(b);
      });
      row.appendChild(lab);
      row.appendChild(box);
      detail.appendChild(row);
    };

    const addNegList = (currency: boolean): void => {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      row.style.alignItems = "flex-start";
      const lab = document.createElement("label");
      lab.textContent = "负数:";
      const box = document.createElement("div");
      box.className = "fs-format-cells__listbox fs-format-cells__listbox--neg";
      const posPlain = posNumberPattern(numState.decimals, numState.useThousands);
      const posCur = posCurrencyPattern(numState.currencySymbol, numState.decimals, true);
      const pos = currency ? posCur : posPlain;
      const curStyle = currency ? numState.currencyNegStyle : numState.negStyle;
      const demo = -1234.1;
      for (let st = 0; st < 5; st++) {
        const style = st as 0 | 1 | 2 | 3 | 4;
        const code = withNegatives(pos, style);
        const label = formatCellDisplayWithStyle(demo, { numberFormat: code });
        const b = document.createElement("button");
        b.type = "button";
        b.textContent = label;
        if (style === curStyle) {
          b.classList.add("fs-format-cells__list-sel");
        }
        if (style === 0 || style === 2 || style === 4) {
          b.style.color = "#c50f1f";
        }
        b.addEventListener("click", () => {
          for (const c of box.querySelectorAll("button")) {
            c.classList.remove("fs-format-cells__list-sel");
          }
          b.classList.add("fs-format-cells__list-sel");
          if (currency) {
            numState = { ...numState, currencyNegStyle: style };
          } else {
            numState = { ...numState, negStyle: style };
          }
          updateSample();
        });
        box.appendChild(b);
      }
      row.appendChild(lab);
      row.appendChild(box);
      detail.appendChild(row);
    };

    if (cat === "general" || cat === "text") {
      // only sample
    } else if (cat === "number") {
      addDecimals();
      addThousands();
      addNegList(false);
    } else if (cat === "currency") {
      addDecimals();
      addSymbolSelect("currencySymbol");
      addNegList(true);
    } else if (cat === "accounting") {
      addDecimals();
      addSymbolSelect("accountingSymbol");
    } else if (cat === "date") {
      const list = numState.dateLocale === "zh-CN" ? DATE_TYPES_ZH : DATE_TYPES_EN;
      addTypeList(list, numState.dateTypeIndex, (i) => {
        numState = { ...numState, dateTypeIndex: i };
        updateSample();
      });
      addLocaleSelect("dateLocale", () => {
        numState = { ...numState, dateTypeIndex: 0 };
        rebuildNumberDetail();
        syncCategoryHighlight();
        updateDesc();
        updateSample();
      });
      const calRow = document.createElement("div");
      calRow.className = "fs-format-cells__row";
      const lab = document.createElement("label");
      lab.textContent = "日历类型:";
      const sel = document.createElement("select");
      sel.className = "fs-format-cells__select";
      const opt = document.createElement("option");
      opt.value = "gregorian";
      opt.textContent = "公历";
      sel.appendChild(opt);
      sel.disabled = true;
      calRow.appendChild(lab);
      calRow.appendChild(sel);
      detail.appendChild(calRow);
    } else if (cat === "time") {
      const list = numState.timeLocale === "zh-CN" ? TIME_TYPES_ZH : TIME_TYPES_EN;
      addTypeList(list, numState.timeTypeIndex, (i) => {
        numState = { ...numState, timeTypeIndex: i };
        updateSample();
      });
      addLocaleSelect("timeLocale", () => {
        numState = { ...numState, timeTypeIndex: 0 };
        rebuildNumberDetail();
        syncCategoryHighlight();
        updateDesc();
        updateSample();
      });
    } else if (cat === "percentage") {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      const lab = document.createElement("label");
      lab.textContent = "小数位数:";
      const inp = document.createElement("input");
      inp.type = "number";
      inp.className = "fs-format-cells__input-num";
      inp.min = "0";
      inp.max = "10";
      inp.value = String(numState.percentageDecimals);
      inp.addEventListener("change", () => {
        const n = Math.max(0, Math.min(10, Math.floor(Number(inp.value) || 0)));
        numState = { ...numState, percentageDecimals: n };
        inp.value = String(n);
        updateSample();
      });
      row.appendChild(lab);
      row.appendChild(inp);
      detail.appendChild(row);
    } else if (cat === "fraction") {
      addTypeList(FRACTION_TYPES, numState.fractionTypeIndex, (i) => {
        numState = { ...numState, fractionTypeIndex: i };
        updateSample();
      });
    } else if (cat === "scientific") {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      const lab = document.createElement("label");
      lab.textContent = "小数位数:";
      const inp = document.createElement("input");
      inp.type = "number";
      inp.className = "fs-format-cells__input-num";
      inp.min = "2";
      inp.max = "15";
      inp.value = String(numState.scientificDecimals);
      inp.addEventListener("change", () => {
        const n = Math.max(2, Math.min(15, Math.floor(Number(inp.value) || 2)));
        numState = { ...numState, scientificDecimals: n };
        inp.value = String(n);
        updateSample();
      });
      row.appendChild(lab);
      row.appendChild(inp);
      detail.appendChild(row);
    } else if (cat === "special") {
      const list =
        numState.specialLocale === "zh-CN" ? SPECIAL_TYPES_ZH : SPECIAL_TYPES_EN;
      addTypeList(list, numState.specialTypeIndex, (i) => {
        numState = { ...numState, specialTypeIndex: i };
        updateSample();
      });
      addLocaleSelect("specialLocale", () => {
        numState = { ...numState, specialTypeIndex: 0 };
        rebuildNumberDetail();
        syncCategoryHighlight();
        updateDesc();
        updateSample();
      });
    } else if (cat === "custom") {
      const row = document.createElement("div");
      row.className = "fs-format-cells__row";
      row.style.alignItems = "flex-start";
      const lab = document.createElement("label");
      lab.textContent = "类型:";
      const wrap = document.createElement("div");
      wrap.style.flex = "1";
      wrap.style.minWidth = "0";
      const ti = document.createElement("input");
      ti.type = "text";
      ti.className = "fs-format-cells__text-input";
      ti.value = numState.customCode;
      ti.addEventListener("input", () => {
        numState = { ...numState, customCode: ti.value };
        updateSample();
      });
      wrap.appendChild(ti);
      const listBox = document.createElement("div");
      listBox.className = "fs-format-cells__listbox fs-format-cells__custom-list";
      for (const line of customFormatBuiltinList()) {
        const b = document.createElement("button");
        b.type = "button";
        b.textContent = line;
        b.addEventListener("click", () => {
          ti.value = line;
          numState = { ...numState, customCode: line };
          updateSample();
        });
        listBox.appendChild(b);
      }
      wrap.appendChild(listBox);
      row.appendChild(lab);
      row.appendChild(wrap);
      detail.appendChild(row);
    }

    updateSample();
  };

  const alignmentRoot = document.createElement("div");
  alignmentRoot.className = "fs-format-cells__align-outer";

  const rebuildAlignmentPanel = (): void => {
    alignmentRoot.replaceChildren();
    const topGrid = document.createElement("div");
    topGrid.className = "fs-format-cells__align-grid";

    const left = document.createElement("div");
    left.className = "fs-format-cells__align-left";

    const rowH = document.createElement("div");
    rowH.className = "fs-format-cells__row";
    const labH = document.createElement("label");
    labH.textContent = "水平对齐方式:";
    const selH = document.createElement("select");
    selH.className = "fs-format-cells__select";
    for (const o of HORIZONTAL_OPTIONS) {
      const opt = document.createElement("option");
      opt.value = o.id;
      opt.textContent = o.label;
      selH.appendChild(opt);
    }
    selH.value = alignState.horizontal;
    selH.addEventListener("change", () => {
      alignState = { ...alignState, horizontal: selH.value as AlignmentHorizontalUi };
    });
    rowH.appendChild(labH);
    rowH.appendChild(selH);
    left.appendChild(rowH);

    const rowV = document.createElement("div");
    rowV.className = "fs-format-cells__row";
    const labV = document.createElement("label");
    labV.textContent = "垂直对齐方式:";
    const selV = document.createElement("select");
    selV.className = "fs-format-cells__select";
    for (const o of VERTICAL_OPTIONS) {
      const opt = document.createElement("option");
      opt.value = o.id;
      opt.textContent = o.label;
      selV.appendChild(opt);
    }
    selV.value = alignState.vertical;
    selV.addEventListener("change", () => {
      alignState = { ...alignState, vertical: selV.value as CellVerticalAlign };
    });
    rowV.appendChild(labV);
    rowV.appendChild(selV);
    left.appendChild(rowV);

    const rowIndent = document.createElement("div");
    rowIndent.className = "fs-format-cells__row";
    const labI = document.createElement("label");
    labI.textContent = "缩进:";
    const inpI = document.createElement("input");
    inpI.type = "number";
    inpI.className = "fs-format-cells__input-num";
    inpI.min = "0";
    inpI.max = "255";
    inpI.value = String(alignState.indent);
    inpI.addEventListener("change", () => {
      const n = Math.max(0, Math.min(255, Math.floor(Number(inpI.value) || 0)));
      alignState = { ...alignState, indent: n };
      inpI.value = String(n);
    });
    rowIndent.appendChild(labI);
    rowIndent.appendChild(inpI);
    left.appendChild(rowIndent);

    const rowDist = document.createElement("div");
    rowDist.className = "fs-format-cells__checkbox-row fs-format-cells__align-distributed-row";
    const cbDist = document.createElement("input");
    cbDist.type = "checkbox";
    cbDist.disabled = true;
    cbDist.className = "fs-format-cells__checkbox";
    const spDist = document.createElement("span");
    spDist.textContent = "两端分散对齐";
    rowDist.appendChild(cbDist);
    rowDist.appendChild(spDist);
    left.appendChild(rowDist);

    const sep = document.createElement("div");
    sep.className = "fs-format-cells__align-sep";

    const right = document.createElement("div");
    right.className = "fs-format-cells__align-right";
    const orientCol = document.createElement("div");
    orientCol.className = "fs-format-cells__orient-row";
    const ol = document.createElement("div");
    ol.className = "fs-format-cells__orient-label";
    ol.textContent = "方向:";
    orientCol.appendChild(ol);

    const preview = document.createElement("div");
    preview.className = "fs-format-cells__orient-preview";

    const vertBox = document.createElement("button");
    vertBox.type = "button";
    vertBox.className = "fs-format-cells__orient-vertical";
    vertBox.textContent = "文本";
    vertBox.title = "恢复水平文字方向";
    if (alignState.orientationVertical) {
      vertBox.classList.add("fs-format-cells__orient-vertical--on");
    }
    vertBox.addEventListener("click", () => {
      alignState = {
        ...alignState,
        orientationVertical: false,
        degrees: 0,
      };
      rebuildAlignmentPanel();
    });

    /** 圆心即指针枢轴；右半圆经上(90°)、右(0°)、下(−90°)，与 Excel 半径针一致。 */
    const arcR = 50;
    const arcCx = 78;
    const arcCy = 50;
    const pivotX = arcCx;
    const pivotY = arcCy;
    const arcNorthX = arcCx;
    const arcNorthY = arcCy - arcR;
    const arcSouthX = arcCx;
    const arcSouthY = arcCy + arcR;
    const tickR = arcR - 3;

    const gauge = document.createElement("div");
    gauge.className = "fs-format-cells__orient-gauge";
    const surface = document.createElement("div");
    surface.className = "fs-format-cells__orient-gauge-surface";
    const svgNS = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(svgNS, "svg");
    svg.setAttribute("viewBox", "0 0 204 102");
    svg.setAttribute("role", "img");
    svg.setAttribute("aria-label", "沿半圆拖动：上为 90 度，右为 0 度，下为负 90 度");

    const arcPath = document.createElementNS(svgNS, "path");
    arcPath.setAttribute(
      "d",
      `M ${arcNorthX} ${arcNorthY} A ${arcR} ${arcR} 0 0 1 ${arcSouthX} ${arcSouthY}`,
    );
    arcPath.setAttribute("fill", "none");
    arcPath.setAttribute("stroke", "#a19f9d");
    arcPath.setAttribute("stroke-width", "2");
    svg.appendChild(arcPath);

    const mainTicks = new Set([-90, -45, 0, 45, 90]);
    for (let d = -90; d <= 90; d += 15) {
      const phi = (-d * Math.PI) / 180;
      const mx = arcCx + tickR * Math.cos(phi);
      const my = arcCy + tickR * Math.sin(phi);
      if (mainTicks.has(d)) {
        const dm = document.createElementNS(svgNS, "path");
        dm.setAttribute(
          "d",
          `M ${mx} ${my - 3.5} L ${mx + 3} ${my} L ${mx} ${my + 3.5} L ${mx - 3} ${my} Z`,
        );
        dm.setAttribute("fill", "#8a8886");
        const tickRot = (phi * 180) / Math.PI + 90;
        dm.setAttribute("transform", `rotate(${tickRot}, ${mx}, ${my})`);
        svg.appendChild(dm);
      } else {
        const dot = document.createElementNS(svgNS, "circle");
        dot.setAttribute("cx", String(mx));
        dot.setAttribute("cy", String(my));
        dot.setAttribute("r", "1.35");
        dot.setAttribute("fill", "#b3b0ad");
        svg.appendChild(dot);
      }
    }

    const needleGroup = document.createElementNS(svgNS, "g");
    const needleLine = document.createElementNS(svgNS, "line");
    needleLine.setAttribute("stroke", "#323130");
    needleLine.setAttribute("stroke-width", "2");
    needleLine.setAttribute("stroke-linecap", "round");
    const handleDiamond = document.createElementNS(svgNS, "path");
    handleDiamond.setAttribute("d", "M 0 -6.2 L 5.4 0 L 0 6.2 L -5.4 0 Z");
    handleDiamond.setAttribute("fill", "#c50f1f");
    handleDiamond.setAttribute("stroke", "#a4262c");
    handleDiamond.setAttribute("stroke-width", "0.35");
    const needleText = document.createElementNS(svgNS, "text");
    needleText.setAttribute("font-size", "11");
    needleText.setAttribute("font-family", "system-ui,sans-serif");
    needleText.setAttribute("fill", "#323130");
    needleText.textContent = "文本";
    needleText.setAttribute("text-anchor", "middle");
    needleText.setAttribute("dominant-baseline", "middle");
    needleGroup.appendChild(needleLine);
    needleGroup.appendChild(handleDiamond);
    needleGroup.appendChild(needleText);
    svg.appendChild(needleGroup);

    const tipFromDegrees = (deg: number): { x: number; y: number } => {
      const phi = (-deg * Math.PI) / 180;
      return {
        x: arcCx + arcR * Math.cos(phi),
        y: arcCy + arcR * Math.sin(phi),
      };
    };

    const applyDegVisual = (): void => {
      const deg = alignState.orientationVertical ? 0 : alignState.degrees;
      const tip = tipFromDegrees(deg);
      const lineInset = 7;
      const ux = (tip.x - arcCx) / arcR;
      const uy = (tip.y - arcCy) / arcR;
      const lineX2 = tip.x - ux * lineInset;
      const lineY2 = tip.y - uy * lineInset;
      needleLine.setAttribute("x1", String(pivotX));
      needleLine.setAttribute("y1", String(pivotY));
      needleLine.setAttribute("x2", String(lineX2));
      needleLine.setAttribute("y2", String(lineY2));
      const phi = (-deg * Math.PI) / 180;
      const tangentDeg = (phi * 180) / Math.PI + 90;
      handleDiamond.setAttribute("transform", `translate(${tip.x},${tip.y}) rotate(${tangentDeg})`);
      const t = 0.42;
      const midX = arcCx + (tip.x - arcCx) * t;
      const midY = arcCy + (tip.y - arcCy) * t;
      needleText.setAttribute("transform", `translate(${midX},${midY}) rotate(${-deg})`);
    };
    applyDegVisual();

    surface.appendChild(svg);

    const degRow = document.createElement("div");
    degRow.className = "fs-format-cells__row";
    degRow.style.marginTop = "4px";
    const degLab = document.createElement("label");
    degLab.textContent = "度:";
    const degInp = document.createElement("input");
    degInp.type = "number";
    degInp.className = "fs-format-cells__input-num";
    degInp.min = "-90";
    degInp.max = "90";
    degInp.value = String(alignState.degrees);
    degInp.disabled = alignState.orientationVertical;
    degInp.addEventListener("change", () => {
      let n = Math.round(Number(degInp.value) || 0);
      n = Math.max(-90, Math.min(90, n));
      alignState = {
        ...alignState,
        degrees: n,
        orientationVertical: false,
      };
      rebuildAlignmentPanel();
    });

    let gaugeDragging = false;
    const setGaugeDragging = (on: boolean): void => {
      gaugeDragging = on;
      surface.classList.toggle("fs-format-cells__orient-gauge-surface--dragging", on);
    };
    const applyPointerToDegrees = (clientX: number, clientY: number): void => {
      const nextDeg = textRotationDegFromGaugePointer(clientX, clientY, svg, arcCx, arcCy);
      alignState = { ...alignState, orientationVertical: false, degrees: nextDeg };
      applyDegVisual();
      degInp.value = String(nextDeg);
      degInp.disabled = false;
    };
    surface.addEventListener("pointerdown", (ev: PointerEvent) => {
      if (ev.button !== 0) {
        return;
      }
      ev.preventDefault();
      try {
        surface.setPointerCapture(ev.pointerId);
      } catch {
        /* ignore */
      }
      setGaugeDragging(true);
      applyPointerToDegrees(ev.clientX, ev.clientY);
    });
    surface.addEventListener("pointermove", (ev: PointerEvent) => {
      if (!gaugeDragging) {
        return;
      }
      ev.preventDefault();
      applyPointerToDegrees(ev.clientX, ev.clientY);
    });
    const endGaugeDrag = (ev: PointerEvent): void => {
      if (!gaugeDragging) {
        return;
      }
      setGaugeDragging(false);
      try {
        if (surface.hasPointerCapture(ev.pointerId)) {
          surface.releasePointerCapture(ev.pointerId);
        }
      } catch {
        /* ignore */
      }
    };
    surface.addEventListener("pointerup", endGaugeDrag);
    surface.addEventListener("pointercancel", endGaugeDrag);
    surface.addEventListener("lostpointercapture", () => setGaugeDragging(false));

    gauge.appendChild(surface);
    degRow.appendChild(degLab);
    degRow.appendChild(degInp);
    gauge.appendChild(degRow);

    preview.appendChild(vertBox);
    preview.appendChild(gauge);
    orientCol.appendChild(preview);
    right.appendChild(orientCol);

    topGrid.appendChild(left);
    topGrid.appendChild(sep);
    topGrid.appendChild(right);
    alignmentRoot.appendChild(topGrid);

    const ctrl = document.createElement("div");
    ctrl.className = "fs-format-cells__control-block";
    const ct = document.createElement("div");
    ct.className = "fs-format-cells__control-title";
    ct.textContent = "文本控制:";
    ctrl.appendChild(ct);

    const rowWrap = document.createElement("div");
    rowWrap.className = "fs-format-cells__checkbox-row";
    const cbW = document.createElement("input");
    cbW.type = "checkbox";
    cbW.className = "fs-format-cells__checkbox";
    cbW.checked = alignState.wrap;
    cbW.addEventListener("change", () => {
      alignState = { ...alignState, wrap: cbW.checked };
    });
    const spW = document.createElement("span");
    spW.textContent = "自动换行";
    rowWrap.appendChild(cbW);
    rowWrap.appendChild(spW);
    ctrl.appendChild(rowWrap);

    const rowSh = document.createElement("div");
    rowSh.className = "fs-format-cells__checkbox-row";
    const cbS = document.createElement("input");
    cbS.type = "checkbox";
    cbS.className = "fs-format-cells__checkbox";
    cbS.checked = alignState.shrink;
    cbS.addEventListener("change", () => {
      alignState = { ...alignState, shrink: cbS.checked };
    });
    const spS = document.createElement("span");
    spS.textContent = "缩小字体填充";
    rowSh.appendChild(cbS);
    rowSh.appendChild(spS);
    ctrl.appendChild(rowSh);

    const rowM = document.createElement("div");
    rowM.className = "fs-format-cells__checkbox-row";
    const cbM = document.createElement("input");
    cbM.type = "checkbox";
    cbM.className = "fs-format-cells__checkbox";
    cbM.checked = alignState.mergeCells;
    cbM.addEventListener("change", () => {
      alignState = { ...alignState, mergeCells: cbM.checked };
    });
    const spM = document.createElement("span");
    spM.textContent = "合并单元格";
    rowM.appendChild(cbM);
    rowM.appendChild(spM);
    ctrl.appendChild(rowM);

    alignmentRoot.appendChild(ctrl);
  };

  const renderBody = (): void => {
    body.replaceChildren();
    if (mainTab === "number") {
      body.appendChild(numberLayout);
      syncCategoryHighlight();
      rebuildNumberDetail();
    } else if (mainTab === "alignment") {
      rebuildAlignmentPanel();
      body.appendChild(alignmentRoot);
    } else {
      body.appendChild(placeholderEl);
    }
    updateDesc();
  };

  for (const [id, btn] of tabButtons) {
    btn.addEventListener("click", () => {
      mainTab = id;
      syncTabHighlight();
      renderBody();
    });
  }

  for (const [id, btn] of catItemBtns) {
    btn.addEventListener("click", () => {
      numState = { ...numState, category: id };
      syncCategoryHighlight();
      rebuildNumberDetail();
      updateDesc();
    });
  }

  const close = (): void => {
    overlay.remove();
    onClose?.();
  };

  closeBtn.addEventListener("click", close);
  cancelBtn.addEventListener("click", close);
  overlay.addEventListener("pointerdown", (ev) => {
    if (ev.target === overlay) {
      close();
    }
  });

  okBtn.addEventListener("click", () => {
    const code = buildFormatFromNumberState(numState);
    const alignPatch = alignmentPatchIfChanged(alignState, initialAlignState);
    flex.applySelectionStylePatch({
      ...alignPatch,
      numberFormat: code === null ? null : code,
    });
    if (alignState.mergeCells !== initialAlignState.mergeCells) {
      if (alignState.mergeCells) {
        const r = flex.selection.getNormalizedRange();
        const multi = r.endRow > r.startRow || r.endCol > r.startCol;
        if (multi) {
          flex.applySelectionMerge("mergeCells");
        }
      } else {
        flex.applySelectionMerge("unmerge");
      }
    }
    close();
  });

  syncTabHighlight();
  renderBody();

  return overlay;
}

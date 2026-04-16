import {
  type CellHorizontalAlign,
  type CellScalar,
  type CellStyle,
  type CellStylePatch,
  type CellTextOrientation,
  type CellVerticalAlign,
  formatCellDisplayWithStyle,
} from "@flexsheet/core";
import {
  appendRibbonColorPaletteContent,
  RIBBON_NUMBER_FORMAT_PRESETS,
  RIBBON_FONT_FAMILY_DEFAULT_PREVIEW,
  RIBBON_FONT_FAMILY_ITEMS,
  argb8ToCssHex6,
  cellStyleToRibbonHomeFontChrome,
  cssHexToFillArgb,
  showRibbonColorDialog,
} from "@flexsheet/toolbar";
import { paintCellFillPatternOverlay } from "@flexsheet/renderer";
import type { FlexSheet } from "./flex-sheet.js";
import {
  FORMAT_CELLS_LINE_STYLES,
  type FormatCellsBorderState,
  formatCellsBorderStateEqual,
  inferFormatCellsBorderState,
} from "./format-cells-border.js";
import {
  createFormatCellsBorderPreviewSvg,
  createFormatCellsLineSwatchHost,
} from "./format-cells-line-swatch.js";
import { ensureFsSheetPromptStyles } from "./fs-dialog-styles.js";
import {
  fillPatchIfChanged,
  FORMAT_CELLS_PATTERN_GRID_ORDER,
  inferFormatCellsFillState,
  type FormatCellsFillState,
} from "./format-cells-fill.js";
import {
  inferFormatCellsProtectionState,
  inferFormatCellsProtectionUiState,
  protectionPatchIfChanged,
  type FormatCellsProtectionState,
  type FormatCellsProtectionUiState,
} from "./format-cells-protection.js";

/** 可变的样式补丁对象，用于在对话框内逐步写入（`CellStylePatch` 字段为 readonly）。 */
type CellStylePatchMutable = { -readonly [K in keyof CellStylePatch]: CellStylePatch[K] };

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

const PROTECTION_TAB_DESC =
  "要锁定单元格或隐藏公式，请保护工作表。在「审阅」选项卡上，单击「保护工作表」。";

const FILL_TAB_DESC = "设置单元格背景色与图案样式；图案颜色可为自动或自定义。示例预览反映当前选择。";

const ALIGNMENT_TAB_DESC =
  "设置单元格内文本的水平与垂直对齐、缩进、方向，以及自动换行、缩小字体填充等选项。";

const FONT_TAB_DESC =
  "设置单元格内文字的字体、字号、字形、下划线、颜色，以及删除线与上标、下标等效果。";

const BORDER_TAB_DESC =
  "单击线型和颜色，然后单击预设边框图案或各个边框按钮。您还可以通过在预览框中单击来应用边框。";

function formatCellsBorderPreviewColor(state: FormatCellsBorderState): string {
  if (state.colorAuto) {
    return "#323130";
  }
  const a = state.colorArgb?.trim();
  if (a !== undefined && /^[\dA-Fa-f]{8}$/i.test(a)) {
    return argb8ToCssHex6(a);
  }
  return "#323130";
}

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
  const p: CellStylePatchMutable = {};
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

interface FormatCellsFontState {
  normalFont: boolean;
  selectedFontItemIndex: number;
  customFontFamilyStack: string;
  fontSizePt: number;
  bold: boolean;
  italic: boolean;
  underline: "none" | "single" | "double";
  fgArgb: string | null;
  strikethrough: boolean;
  fontScript: "none" | "superscript" | "subscript";
}

const FONT_SIZE_CHOICES: readonly number[] = [
  8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
];

const UNDERLINE_OPTIONS: readonly { readonly id: "none" | "single" | "double"; readonly label: string }[] =
  [
    { id: "none", label: "无" },
    { id: "single", label: "单下划线" },
    { id: "double", label: "双下划线" },
  ];

function hasExplicitFontStyle(style: CellStyle | null): boolean {
  if (style === null) {
    return false;
  }
  const fg = style.fgArgb?.trim();
  return (
    style.bold === true ||
    style.italic === true ||
    (style.fontFamily !== undefined && String(style.fontFamily).trim() !== "") ||
    style.fontSizePt !== undefined ||
    style.underline !== undefined ||
    (fg !== undefined && fg !== "") ||
    style.strikethrough === true ||
    (style.fontScript !== undefined && style.fontScript !== null)
  );
}

function inferFormatCellsFontState(style: CellStyle | null): FormatCellsFontState {
  const chrome = cellStyleToRibbonHomeFontChrome(style);
  const idx = RIBBON_FONT_FAMILY_ITEMS.findIndex((it) => it.label === chrome.fontLabel);
  const rawFam = style?.fontFamily?.trim();
  const u = style?.underline;
  const underline: "none" | "single" | "double" =
    u === "double" ? "double" : u === "single" ? "single" : "none";
  const fs = style?.fontScript;
  const fontScript: FormatCellsFontState["fontScript"] =
    fs === "superscript" ? "superscript" : fs === "subscript" ? "subscript" : "none";
  return {
    normalFont: !hasExplicitFontStyle(style),
    selectedFontItemIndex: idx,
    customFontFamilyStack: rawFam !== undefined && rawFam !== "" ? rawFam : chrome.fontPreviewCss,
    fontSizePt: style?.fontSizePt ?? (Number(chrome.sizeLabel) || 11),
    bold: style?.bold === true,
    italic: style?.italic === true,
    underline,
    fgArgb: style?.fgArgb?.trim() && /^[\dA-Fa-f]{8}$/.test(style.fgArgb.trim())
      ? style.fgArgb.trim().toUpperCase()
      : null,
    strikethrough: style?.strikethrough === true,
    fontScript,
  };
}

function resolveFontFamilyCss(state: FormatCellsFontState): string {
  if (state.selectedFontItemIndex >= 0) {
    const it = RIBBON_FONT_FAMILY_ITEMS[state.selectedFontItemIndex]!;
    if (it.previewFontFamily !== undefined && it.previewFontFamily !== "") {
      return it.previewFontFamily;
    }
    return `"${it.label}", sans-serif`;
  }
  const t = state.customFontFamilyStack.trim();
  return t !== "" ? t : RIBBON_FONT_FAMILY_DEFAULT_PREVIEW;
}

function fontPatchIfChanged(
  cur: FormatCellsFontState,
  ini: FormatCellsFontState,
): CellStylePatch {
  if (cur.normalFont) {
    if (!ini.normalFont) {
      return {
        bold: null,
        italic: null,
        fontFamily: null,
        fontSizePt: null,
        underline: null,
        fgArgb: null,
        strikethrough: null,
        fontScript: null,
      };
    }
    return {};
  }
  if (ini.normalFont) {
    const p: CellStylePatchMutable = {};
    p.bold = cur.bold ? true : null;
    p.italic = cur.italic ? true : null;
    p.fontFamily = resolveFontFamilyCss(cur);
    p.fontSizePt = cur.fontSizePt;
    p.underline = cur.underline === "none" ? null : cur.underline;
    p.fgArgb = cur.fgArgb;
    p.strikethrough = cur.strikethrough ? true : null;
    p.fontScript = cur.fontScript === "none" ? null : cur.fontScript;
    return p;
  }
  const p: CellStylePatchMutable = {};
  if (resolveFontFamilyCss(cur) !== resolveFontFamilyCss(ini)) {
    p.fontFamily = resolveFontFamilyCss(cur);
  }
  if (cur.fontSizePt !== ini.fontSizePt) {
    p.fontSizePt = cur.fontSizePt;
  }
  if (cur.bold !== ini.bold) {
    p.bold = cur.bold ? true : null;
  }
  if (cur.italic !== ini.italic) {
    p.italic = cur.italic ? true : null;
  }
  if (cur.underline !== ini.underline) {
    p.underline = cur.underline === "none" ? null : cur.underline;
  }
  if (cur.fgArgb !== ini.fgArgb) {
    p.fgArgb = cur.fgArgb;
  }
  if (cur.strikethrough !== ini.strikethrough) {
    p.strikethrough = cur.strikethrough ? true : null;
  }
  if (cur.fontScript !== ini.fontScript) {
    p.fontScript = cur.fontScript === "none" ? null : cur.fontScript;
  }
  return p;
}

function stylePresetLabel(bold: boolean, italic: boolean): string {
  if (bold && italic) {
    return "加粗 倾斜";
  }
  if (bold) {
    return "加粗";
  }
  if (italic) {
    return "倾斜";
  }
  return "常规";
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
  width: min(600px, calc(100vw - 24px));
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
.fs-format-cells__font-outer {
  display: flex;
  flex-direction: column;
  gap: 12px;
  min-height: 260px;
}
.fs-format-cells__font-grid {
  display: grid;
  grid-template-columns: minmax(0, 1.1fr) minmax(0, 0.75fr) minmax(0, 0.55fr);
  gap: 10px 12px;
  align-items: start;
}
.fs-format-cells__font-col {
  display: flex;
  flex-direction: column;
  gap: 4px;
  min-width: 0;
}
.fs-format-cells__font-col-label {
  font-size: 12px;
  font-weight: 600;
  color: #323130;
}
.fs-format-cells__font-field {
  width: 100%;
  box-sizing: border-box;
  padding: 4px 6px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  font-size: 13px;
  font-family: inherit;
  background: #fff;
}
.fs-format-cells__font-field:disabled {
  background: #f3f2f1;
  color: #605e5c;
}
.fs-format-cells__font-listbox {
  min-height: 120px;
  max-height: 160px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  overflow-y: auto;
  background: #fff;
}
.fs-format-cells__font-listbox button {
  display: block;
  width: 100%;
  padding: 4px 8px;
  border: none;
  background: transparent;
  text-align: left;
  font-size: 12px;
  cursor: pointer;
  font-family: inherit;
}
.fs-format-cells__font-listbox button:hover {
  background: #e8f5e9;
}
.fs-format-cells__font-listbox button.fs-format-cells__font-sel {
  background: #c8e6c9;
}
.fs-format-cells__font-listbox--name button {
  font-size: 12px;
}
.fs-format-cells__font-row2 {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 10px 16px;
  font-size: 13px;
  color: #323130;
}
.fs-format-cells__font-row2 label {
  display: flex;
  align-items: center;
  gap: 6px;
}
.fs-format-cells__font-preview-box {
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  min-height: 52px;
  padding: 10px 12px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 14px;
  color: #201f1e;
}
.fs-format-cells__border-wrap {
  display: grid;
  grid-template-columns: minmax(0, 200px) minmax(0, 1fr);
  gap: 14px 18px;
  align-items: start;
  min-height: 260px;
}
.fs-format-cells__border-line-label {
  font-size: 12px;
  font-weight: 600;
  color: #323130;
  margin-bottom: 6px;
}
.fs-format-cells__border-line-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0;
  margin-bottom: 12px;
  border: 1px solid #c8c6c4;
  background: #fff;
  box-sizing: border-box;
}
.fs-format-cells__border-line-grid > .fs-format-cells__border-line-btn {
  border-right: 1px solid #e1dfdd;
  border-bottom: 1px solid #e1dfdd;
  margin: 0;
}
.fs-format-cells__border-line-grid > .fs-format-cells__border-line-btn:nth-child(2n) {
  border-right: none;
}
.fs-format-cells__border-line-grid > .fs-format-cells__border-line-btn:nth-child(n + 13) {
  border-bottom: none;
}
.fs-format-cells__border-line-btn {
  min-height: 28px;
  padding: 5px 6px;
  border: none;
  border-radius: 0;
  background: #fff;
  cursor: pointer;
  font-family: inherit;
  display: flex;
  align-items: center;
  justify-content: center;
}
.fs-format-cells__border-line-btn:hover:not(:disabled) {
  background: #f3f2f1;
}
.fs-format-cells__border-line-btn--active {
  outline: 1px dotted #000000;
  outline-offset: -1px;
  background: #ffffff;
}
.fs-format-cells__border-line-swatch {
  display: block;
  width: 100%;
  max-width: 96px;
  min-height: 16px;
  box-sizing: border-box;
}
.fs-format-cells__border-presets {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.fs-format-cells__border-presets-label {
  font-size: 12px;
  font-weight: 600;
  color: #323130;
}
.fs-format-cells__border-presets-row {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
}
.fs-format-cells__border-preset {
  min-width: 72px;
  padding: 10px 8px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  font-size: 12px;
  color: #323130;
  font-family: inherit;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 4px;
}
.fs-format-cells__border-preset:hover:not(:disabled) {
  background: #e8f5e9;
}
.fs-format-cells__border-preset:disabled {
  opacity: 0.45;
  cursor: not-allowed;
}
.fs-format-cells__border-preset-icon {
  width: 36px;
  height: 28px;
  box-sizing: border-box;
  border: 1px solid #a19f9d;
}
.fs-format-cells__border-preset-icon--none {
  border-style: dashed;
  border-color: #c8c6c4;
}
.fs-format-cells__border-preset-icon--outline {
  border-width: 2px;
  border-color: #323130;
}
.fs-format-cells__border-preset-icon--inside {
  background: linear-gradient(#323130, #323130) center/100% 1px no-repeat,
    linear-gradient(#323130, #323130) center/1px 100% no-repeat;
  border-width: 1px;
}
.fs-format-cells__border-field-label {
  font-size: 12px;
  font-weight: 600;
  color: #323130;
  margin-bottom: 6px;
}
.fs-format-cells__border-preview-shell {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 6px;
}
.fs-format-cells__border-preview-row {
  display: flex;
  align-items: center;
  gap: 6px;
}
.fs-format-cells__border-edge-fab {
  width: 26px;
  height: 26px;
  padding: 0;
  border: 1px solid #c8c6c4;
  border-radius: 3px;
  background: #fff;
  cursor: pointer;
  font-size: 11px;
  color: #323130;
  line-height: 1;
}
.fs-format-cells__border-edge-fab:hover:not(:disabled) {
  background: #e8f5e9;
}
.fs-format-cells__border-edge-fab:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}
.fs-format-cells__border-edge-fab--on {
  background: #c8e6c9;
  border-color: #217346;
}
.fs-format-cells__border-preview {
  width: 132px;
  height: 88px;
  box-sizing: border-box;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 13px;
  color: #201f1e;
  background: #fff;
  cursor: pointer;
  user-select: none;
}
.fs-format-cells__border-bottom-controls {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 4px;
  flex-wrap: wrap;
  max-width: 280px;
}
.fs-format-cells__fill-outer {
  display: flex;
  flex-direction: column;
  gap: 14px;
  min-height: 260px;
}
.fs-format-cells__protection-outer {
  display: flex;
  flex-direction: column;
  gap: 12px;
  min-height: 200px;
}
.fs-format-cells__protection-tip {
  font-size: 13px;
  color: #323130;
  line-height: 1.5;
  max-width: 440px;
}
/* 纵向排列三行：背景色、图案颜色、图案样式（单列避免窄屏漏项） */
.fs-format-cells__fill-top {
  display: flex;
  flex-direction: column;
  gap: 10px;
  min-width: 0;
}
.fs-format-cells__fill-dd-row {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #323130;
}
.fs-format-cells__fill-dd-row > label {
  min-width: 64px;
  flex-shrink: 0;
}
.fs-format-cells__fill-dd {
  flex: 1;
  min-width: 0;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 10px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  font-family: inherit;
  font-size: 13px;
  text-align: left;
}
.fs-format-cells__fill-dd:hover {
  background: #f3f2f1;
}
.fs-format-cells__fill-swatch {
  width: 22px;
  height: 16px;
  border: 1px solid #a19f9d;
  border-radius: 2px;
  flex-shrink: 0;
  box-sizing: border-box;
}
.fs-format-cells__fill-swatch--empty {
  background: linear-gradient(
    to top right,
    transparent calc(50% - 0.5px),
    #c50f1f calc(50% - 0.5px),
    #c50f1f calc(50% + 0.5px),
    transparent calc(50% + 0.5px)
  );
  background-color: #fff;
}
.fs-format-cells__fill-dd-label {
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.fs-format-cells__fill-dd-arrow {
  width: 0;
  height: 0;
  border-left: 5px solid transparent;
  border-right: 5px solid transparent;
  border-top: 6px solid #217346;
  flex-shrink: 0;
}
.fs-format-cells__fill-dd canvas {
  pointer-events: none;
}
.fs-format-cells__fill-popover.fs-color-menu {
  position: fixed;
  z-index: 10005;
  min-width: 200px;
  max-width: min(320px, calc(100vw - 24px));
  max-height: min(420px, calc(100vh - 48px));
  overflow: auto;
  background: #fff;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  box-shadow: 0 6px 24px rgba(0, 0, 0, 0.18);
  /* 覆盖 Ribbon .fs-color-menu 的 padding，保证左右内边距 */
  padding: 10px 14px;
  box-sizing: border-box;
}
.fs-format-cells__fill-popover .fs-color-menu__heading {
  padding: 6px 2px 4px;
  font-size: 11px;
  color: #605e5c;
}
.fs-format-cells__fill-popover .fs-color-menu__row--top,
.fs-format-cells__fill-popover .fs-color-menu__row--standard {
  padding: 0 2px;
  box-sizing: border-box;
}
.fs-format-cells__fill-popover .fs-color-menu__grid {
  padding: 0 2px 8px;
  box-sizing: border-box;
}
.fs-format-cells__fill-tip {
  padding: 8px 0 10px;
  font-size: 13px;
  color: #323130;
  line-height: 1.45;
  max-width: 280px;
}
.fs-format-cells__fill-auto-btn {
  display: block;
  width: 100%;
  box-sizing: border-box;
  margin: 4px 0;
  padding: 6px 8px;
  border: 1px solid #edebe9;
  border-radius: 4px;
  background: #fff;
  font-size: 13px;
  cursor: pointer;
  font-family: inherit;
  text-align: left;
}
.fs-format-cells__fill-auto-btn:hover {
  background: #e8f5e9;
}
.fs-format-cells__fill-pattern-grid {
  display: grid;
  grid-template-columns: repeat(6, 1fr);
  gap: 4px;
  padding: 4px 0 8px;
  margin: 0;
  box-sizing: border-box;
}
.fs-format-cells__fill-pattern-btn {
  width: 100%;
  aspect-ratio: 1.15;
  min-height: 26px;
  padding: 0;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  background: #fff;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
}
.fs-format-cells__fill-pattern-btn:hover {
  outline: 1px solid #217346;
}
.fs-format-cells__fill-pattern-btn--on {
  outline: 2px solid #217346;
  outline-offset: 0;
}
.fs-format-cells__fill-sample-wrap {
  display: flex;
  flex-direction: column;
  gap: 6px;
  padding-top: 4px;
  border-top: 1px solid #edebe9;
}
.fs-format-cells__fill-sample-wrap > label {
  font-size: 12px;
  font-weight: 600;
  color: #323130;
}
.fs-format-cells__fill-sample {
  width: 100%;
  height: 88px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  display: block;
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
  /** 仅渲染给定标签页；未传时显示全部。 */
  readonly visibleTabs?: readonly FormatCellsMainTabId[];
  /** 隐藏底部帮助说明文字（仍保留按钮区域）。 */
  readonly hideDescription?: boolean;
  /**
   * 自定义「确定」行为；未传时应用到当前选区（默认行为）。
   * 用于“新建表样式”等只需复用弹窗 UI 的场景。
   */
  readonly onApply?: (payload: {
    readonly basePatch: CellStylePatch;
    readonly border: { readonly apply: boolean; readonly state: FormatCellsBorderState };
    readonly mergeCellsChanged: boolean;
    readonly mergeCells: boolean;
  }) => void;
}

/** 挂载「设置单元格格式」对话框，返回 overlay 根节点（供与右键菜单 prompt 生命周期对齐）。 */
export function mountFormatCellsDialog(options: MountFormatCellsDialogOptions): HTMLDivElement {
  ensureFormatCellsDialogStyles();
  const { flex, onClose, hideDescription = false, onApply } = options;
  const visibleTabSet = new Set<FormatCellsMainTabId>(options.visibleTabs ?? MAIN_TAB_LABELS.map((t) => t.id));
  const visibleTabs = MAIN_TAB_LABELS.filter((t) => visibleTabSet.has(t.id));
  const mainTabs = visibleTabs.length > 0 ? visibleTabs : MAIN_TAB_LABELS;
  const initialNf = flex.getActiveCellStyle()?.numberFormat ?? undefined;
  let mainTab: FormatCellsMainTabId = mainTabs[0]?.id ?? "number";
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
  for (const t of mainTabs) {
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
  if (hideDescription) {
    descEl.hidden = true;
  }
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

  const initialFontState = inferFormatCellsFontState(flex.getActiveCellStyle());
  let fontState: FormatCellsFontState = { ...initialFontState };
  const fontRoot = document.createElement("div");
  fontRoot.className = "fs-format-cells__font-outer";

  const borderRoot = document.createElement("div");
  borderRoot.className = "fs-format-cells__border-wrap";

  const fillRoot = document.createElement("div");
  fillRoot.className = "fs-format-cells__fill-outer";

  const initialFillState = inferFormatCellsFillState(flex.getActiveCellStyle());
  let fillState: FormatCellsFillState = { ...initialFillState };
  let fillPopoverCleanup: (() => void) | null = null;
  let fillSampleResizeObserver: ResizeObserver | null = null;

  const closeFillPopover = (): void => {
    if (fillPopoverCleanup !== null) {
      fillPopoverCleanup();
      fillPopoverCleanup = null;
    }
  };

  const normSel0 = flex.selection.getNormalizedRange();
  const selectionMultiCell =
    normSel0.endRow > normSel0.startRow || normSel0.endCol > normSel0.startCol;

  const sheetForFormat = flex.workbook.getActiveSheet();
  const initialProtectionState: FormatCellsProtectionState =
    sheetForFormat !== undefined
      ? inferFormatCellsProtectionState(sheetForFormat, normSel0)
      : {
          locked: true,
          lockedMixed: false,
          hidden: false,
          hiddenMixed: false,
        };
  let protectionUi: FormatCellsProtectionUiState =
    sheetForFormat !== undefined
      ? inferFormatCellsProtectionUiState(sheetForFormat, normSel0)
      : {
          locked: true,
          lockedIndeterminate: false,
          hidden: false,
          hiddenIndeterminate: false,
        };

  const protectionRoot = document.createElement("div");
  protectionRoot.className = "fs-format-cells__protection-outer";

  const rebuildProtectionPanel = (): void => {
    protectionRoot.replaceChildren();
    const tip = document.createElement("div");
    tip.className = "fs-format-cells__protection-tip";
    tip.textContent = PROTECTION_TAB_DESC;
    const row1 = document.createElement("label");
    row1.className = "fs-format-cells__checkbox-row";
    const cbLocked = document.createElement("input");
    cbLocked.type = "checkbox";
    cbLocked.className = "fs-format-cells__checkbox";
    cbLocked.checked = protectionUi.locked;
    cbLocked.indeterminate = protectionUi.lockedIndeterminate;
    cbLocked.addEventListener("change", () => {
      protectionUi = {
        ...protectionUi,
        locked: cbLocked.checked,
        lockedIndeterminate: false,
      };
    });
    const sp1 = document.createElement("span");
    sp1.textContent = "锁定";
    row1.appendChild(cbLocked);
    row1.appendChild(sp1);
    const row2 = document.createElement("label");
    row2.className = "fs-format-cells__checkbox-row";
    const cbHidden = document.createElement("input");
    cbHidden.type = "checkbox";
    cbHidden.className = "fs-format-cells__checkbox";
    cbHidden.checked = protectionUi.hidden;
    cbHidden.indeterminate = protectionUi.hiddenIndeterminate;
    cbHidden.addEventListener("change", () => {
      protectionUi = {
        ...protectionUi,
        hidden: cbHidden.checked,
        hiddenIndeterminate: false,
      };
    });
    const sp2 = document.createElement("span");
    sp2.textContent = "隐藏";
    row2.appendChild(cbHidden);
    row2.appendChild(sp2);
    protectionRoot.appendChild(tip);
    protectionRoot.appendChild(row1);
    protectionRoot.appendChild(row2);
  };

  const initialBorderState = inferFormatCellsBorderState(flex.getActiveCellStyle());
  let borderState: FormatCellsBorderState = {
    ...initialBorderState,
    edges: { ...initialBorderState.edges },
  };

  /** 防止「其他颜色」在异步未结束时重复点击叠多个取色遮罩 */
  let otherColorDialogBusy = false;
  let borderOtherColorBusy = false;

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
    if (hideDescription) {
      descEl.textContent = "";
      return;
    }
    if (mainTab === "number") {
      descEl.textContent = CATEGORY_HELP[numState.category] ?? "";
      return;
    }
    if (mainTab === "alignment") {
      descEl.textContent = ALIGNMENT_TAB_DESC;
      return;
    }
    if (mainTab === "font") {
      descEl.textContent = FONT_TAB_DESC;
      return;
    }
    if (mainTab === "border") {
      descEl.textContent = BORDER_TAB_DESC;
      return;
    }
    if (mainTab === "fill") {
      descEl.textContent = FILL_TAB_DESC;
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

  const updateFontPreview = (): void => {
    const el = fontRoot.querySelector("[data-role='font-preview']");
    if (!(el instanceof HTMLElement)) {
      return;
    }
    el.style.textDecoration = "none";
    el.style.verticalAlign = "baseline";
    if (fontState.normalFont) {
      el.textContent = "AaBbYyZz";
      el.style.fontFamily = RIBBON_FONT_FAMILY_DEFAULT_PREVIEW;
      el.style.fontSize = "11pt";
      el.style.fontWeight = "400";
      el.style.fontStyle = "normal";
      el.style.color = "#000000";
      return;
    }
    el.textContent = "AaBbYyZz";
    el.style.fontFamily = resolveFontFamilyCss(fontState);
    el.style.fontSize = `${fontState.fontSizePt}pt`;
    el.style.fontWeight = fontState.bold ? "700" : "400";
    el.style.fontStyle = fontState.italic ? "italic" : "normal";
    let deco = "";
    if (fontState.underline === "single") {
      deco = "underline";
    } else if (fontState.underline === "double") {
      deco = "underline double";
    }
    if (fontState.strikethrough) {
      deco = deco === "" ? "line-through" : `${deco} line-through`;
    }
    el.style.textDecoration = deco === "" ? "none" : deco;
    el.style.color =
      fontState.fgArgb !== null ? argb8ToCssHex6(fontState.fgArgb) : "#000000";
    if (fontState.fontScript === "superscript") {
      el.style.fontSize = `${Math.max(6, fontState.fontSizePt * 0.65)}pt`;
      el.style.verticalAlign = "super";
    } else if (fontState.fontScript === "subscript") {
      el.style.fontSize = `${Math.max(6, fontState.fontSizePt * 0.65)}pt`;
      el.style.verticalAlign = "sub";
    }
  };

  const rebuildFontPanel = (): void => {
    fontRoot.replaceChildren();
    const disabled = fontState.normalFont;

    const grid = document.createElement("div");
    grid.className = "fs-format-cells__font-grid";

    const colFont = document.createElement("div");
    colFont.className = "fs-format-cells__font-col";
    const lf = document.createElement("div");
    lf.className = "fs-format-cells__font-col-label";
    lf.textContent = "字体(F):";
    const fontNameInput = document.createElement("input");
    fontNameInput.type = "text";
    fontNameInput.className = "fs-format-cells__font-field";
    fontNameInput.disabled = disabled;
    fontNameInput.autocomplete = "off";
    fontNameInput.value =
      fontState.selectedFontItemIndex >= 0
        ? RIBBON_FONT_FAMILY_ITEMS[fontState.selectedFontItemIndex]!.label
        : fontState.customFontFamilyStack.split(",")[0]!.replace(/^["']|["']$/g, "").trim();
    fontNameInput.addEventListener("change", () => {
      fontState = {
        ...fontState,
        normalFont: false,
        selectedFontItemIndex: -1,
        customFontFamilyStack: fontNameInput.value.trim() || RIBBON_FONT_FAMILY_DEFAULT_PREVIEW,
      };
      rebuildFontPanel();
    });
    const fontList = document.createElement("div");
    fontList.className = "fs-format-cells__font-listbox fs-format-cells__font-listbox--name";
    for (let i = 0; i < RIBBON_FONT_FAMILY_ITEMS.length; i++) {
      const it = RIBBON_FONT_FAMILY_ITEMS[i]!;
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = it.label;
      const preview =
        it.previewFontFamily !== undefined && it.previewFontFamily !== ""
          ? it.previewFontFamily
          : `"${it.label}", sans-serif`;
      b.style.fontFamily = preview;
      if (i === fontState.selectedFontItemIndex) {
        b.classList.add("fs-format-cells__font-sel");
      }
      b.addEventListener("click", () => {
        fontState = {
          ...fontState,
          normalFont: false,
          selectedFontItemIndex: i,
          customFontFamilyStack: preview,
        };
        rebuildFontPanel();
      });
      fontList.appendChild(b);
    }
    colFont.appendChild(lf);
    colFont.appendChild(fontNameInput);
    colFont.appendChild(fontList);

    const colShape = document.createElement("div");
    colShape.className = "fs-format-cells__font-col";
    const ls = document.createElement("div");
    ls.className = "fs-format-cells__font-col-label";
    ls.textContent = "字形(O):";
    const shapeInput = document.createElement("input");
    shapeInput.type = "text";
    shapeInput.className = "fs-format-cells__font-field";
    shapeInput.readOnly = true;
    shapeInput.disabled = disabled;
    shapeInput.value = stylePresetLabel(fontState.bold, fontState.italic);
    const shapeList = document.createElement("div");
    shapeList.className = "fs-format-cells__font-listbox";
    const shapeRows: readonly { readonly bold: boolean; readonly italic: boolean; readonly label: string }[] = [
      { bold: false, italic: false, label: "常规" },
      { bold: false, italic: true, label: "倾斜" },
      { bold: true, italic: false, label: "加粗" },
      { bold: true, italic: true, label: "加粗 倾斜" },
    ];
    for (const sr of shapeRows) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = sr.label;
      if (fontState.bold === sr.bold && fontState.italic === sr.italic) {
        b.classList.add("fs-format-cells__font-sel");
      }
      b.addEventListener("click", () => {
        fontState = {
          ...fontState,
          normalFont: false,
          bold: sr.bold,
          italic: sr.italic,
        };
        rebuildFontPanel();
      });
      shapeList.appendChild(b);
    }
    colShape.appendChild(ls);
    colShape.appendChild(shapeInput);
    colShape.appendChild(shapeList);

    const colSize = document.createElement("div");
    colSize.className = "fs-format-cells__font-col";
    const lz = document.createElement("div");
    lz.className = "fs-format-cells__font-col-label";
    lz.textContent = "字号:";
    const sizeInput = document.createElement("input");
    sizeInput.type = "number";
    sizeInput.className = "fs-format-cells__font-field";
    sizeInput.min = "1";
    sizeInput.max = "409";
    sizeInput.disabled = disabled;
    sizeInput.value = String(fontState.fontSizePt);
    sizeInput.addEventListener("change", () => {
      let n = Math.round(Number(sizeInput.value) || 11);
      n = Math.max(1, Math.min(409, n));
      fontState = { ...fontState, normalFont: false, fontSizePt: n };
      rebuildFontPanel();
    });
    const sizeList = document.createElement("div");
    sizeList.className = "fs-format-cells__font-listbox";
    for (const sz of FONT_SIZE_CHOICES) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = String(sz);
      if (sz === fontState.fontSizePt) {
        b.classList.add("fs-format-cells__font-sel");
      }
      b.addEventListener("click", () => {
        fontState = { ...fontState, normalFont: false, fontSizePt: sz };
        rebuildFontPanel();
      });
      sizeList.appendChild(b);
    }
    colSize.appendChild(lz);
    colSize.appendChild(sizeInput);
    colSize.appendChild(sizeList);

    grid.appendChild(colFont);
    grid.appendChild(colShape);
    grid.appendChild(colSize);

    const row2 = document.createElement("div");
    row2.className = "fs-format-cells__font-row2";

    const ulLab = document.createElement("div");
    ulLab.style.display = "flex";
    ulLab.style.alignItems = "center";
    ulLab.style.gap = "6px";
    const ulSpan = document.createElement("span");
    ulSpan.textContent = "下划线(U):";
    const ulSel = document.createElement("select");
    ulSel.className = "fs-format-cells__select";
    ulSel.disabled = disabled;
    ulSel.style.minWidth = "120px";
    for (const u of UNDERLINE_OPTIONS) {
      const o = document.createElement("option");
      o.value = u.id;
      o.textContent = u.label;
      ulSel.appendChild(o);
    }
    ulSel.value = fontState.underline;
    ulSel.addEventListener("change", () => {
      const v = ulSel.value === "double" ? "double" : ulSel.value === "single" ? "single" : "none";
      fontState = { ...fontState, normalFont: false, underline: v };
      rebuildFontPanel();
    });
    ulLab.appendChild(ulSpan);
    ulLab.appendChild(ulSel);

    const colorWrap = document.createElement("div");
    colorWrap.style.display = "flex";
    colorWrap.style.alignItems = "center";
    colorWrap.style.gap = "6px";
    const cSpan = document.createElement("span");
    cSpan.textContent = "颜色(C):";
    const colorInp = document.createElement("input");
    colorInp.type = "color";
    colorInp.disabled = disabled;
    colorInp.style.width = "44px";
    colorInp.style.height = "26px";
    colorInp.style.padding = "0";
    colorInp.style.border = "1px solid #c8c6c4";
    colorInp.style.borderRadius = "4px";
    colorInp.value = argb8ToCssHex6(fontState.fgArgb ?? "FF000000");
    colorInp.addEventListener("input", () => {
      fontState = {
        ...fontState,
        normalFont: false,
        fgArgb: cssHexToFillArgb(colorInp.value),
      };
      updateFontPreview();
    });
    const moreColBtn = document.createElement("button");
    moreColBtn.type = "button";
    moreColBtn.className = "fs-format-cells__btn fs-format-cells__btn--secondary";
    moreColBtn.textContent = "其他颜色…";
    moreColBtn.disabled = disabled;
    moreColBtn.addEventListener("click", () => {
      if (otherColorDialogBusy || disabled) {
        return;
      }
      otherColorDialogBusy = true;
      moreColBtn.disabled = true;
      void (async () => {
        try {
          const picked = await showRibbonColorDialog(
            argb8ToCssHex6(fontState.fgArgb ?? "FF000000"),
          );
          if (picked !== null) {
            fontState = {
              ...fontState,
              normalFont: false,
              fgArgb: cssHexToFillArgb(picked),
            };
            rebuildFontPanel();
          }
        } finally {
          otherColorDialogBusy = false;
          moreColBtn.disabled = disabled;
        }
      })();
    });
    colorWrap.appendChild(cSpan);
    colorWrap.appendChild(colorInp);
    colorWrap.appendChild(moreColBtn);

    const nfRow = document.createElement("label");
    const nfCb = document.createElement("input");
    nfCb.type = "checkbox";
    nfCb.className = "fs-format-cells__checkbox";
    nfCb.checked = fontState.normalFont;
    nfCb.addEventListener("change", () => {
      if (nfCb.checked) {
        fontState = { ...inferFormatCellsFontState(null), normalFont: true };
      } else {
        fontState = { ...fontState, normalFont: false };
      }
      rebuildFontPanel();
    });
    const nfSp = document.createElement("span");
    nfSp.textContent = "普通字体(N)";
    nfRow.appendChild(nfCb);
    nfRow.appendChild(nfSp);

    row2.appendChild(ulLab);
    row2.appendChild(colorWrap);
    row2.appendChild(nfRow);

    const fxTitle = document.createElement("div");
    fxTitle.className = "fs-format-cells__control-title";
    fxTitle.textContent = "效果";
    const fxRow = document.createElement("div");
    fxRow.className = "fs-format-cells__font-row2";

    const mkFx = (key: "strike" | "sup" | "sub", lab: string): HTMLElement => {
      const labEl = document.createElement("label");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "fs-format-cells__checkbox";
      cb.disabled = disabled;
      if (key === "strike") {
        cb.checked = fontState.strikethrough;
        cb.addEventListener("change", () => {
          fontState = { ...fontState, normalFont: false, strikethrough: cb.checked };
          rebuildFontPanel();
        });
      } else if (key === "sup") {
        cb.checked = fontState.fontScript === "superscript";
        cb.addEventListener("change", () => {
          const on = cb.checked;
          fontState = {
            ...fontState,
            normalFont: false,
            fontScript: on ? "superscript" : "none",
          };
          rebuildFontPanel();
        });
      } else {
        cb.checked = fontState.fontScript === "subscript";
        cb.addEventListener("change", () => {
          const on = cb.checked;
          fontState = {
            ...fontState,
            normalFont: false,
            fontScript: on ? "subscript" : "none",
          };
          rebuildFontPanel();
        });
      }
      const sp = document.createElement("span");
      sp.textContent = lab;
      labEl.appendChild(cb);
      labEl.appendChild(sp);
      return labEl;
    };
    fxRow.appendChild(mkFx("strike", "删除线"));
    fxRow.appendChild(mkFx("sup", "上标"));
    fxRow.appendChild(mkFx("sub", "下标"));

    const prevLab = document.createElement("div");
    prevLab.className = "fs-format-cells__font-col-label";
    prevLab.textContent = "预览";
    const prevBox = document.createElement("div");
    prevBox.className = "fs-format-cells__font-preview-box";
    prevBox.dataset.role = "font-preview";

    fontRoot.appendChild(grid);
    fontRoot.appendChild(row2);
    fontRoot.appendChild(fxTitle);
    fontRoot.appendChild(fxRow);
    fontRoot.appendChild(prevLab);
    fontRoot.appendChild(prevBox);
    updateFontPreview();
  };

  const rebuildBorderPanel = (): void => {
    borderRoot.replaceChildren();

    const leftCol = document.createElement("div");
    const lineLab = document.createElement("div");
    lineLab.className = "fs-format-cells__border-line-label";
    lineLab.textContent = "线型:";

    const lineGrid = document.createElement("div");
    lineGrid.className = "fs-format-cells__border-line-grid";

    for (let i = 0; i < FORMAT_CELLS_LINE_STYLES.length; i++) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fs-format-cells__border-line-btn";
      if (i === borderState.lineStyleIndex) {
        b.classList.add("fs-format-cells__border-line-btn--active");
      }
      const entry = FORMAT_CELLS_LINE_STYLES[i]!;
      const sw = createFormatCellsLineSwatchHost(entry.swatch);
      b.appendChild(sw);
      b.addEventListener("click", () => {
        borderState = { ...borderState, lineStyleIndex: i };
        rebuildBorderPanel();
      });
      lineGrid.appendChild(b);
    }

    leftCol.appendChild(lineLab);
    leftCol.appendChild(lineGrid);

    const colorRow = document.createElement("div");
    colorRow.className = "fs-format-cells__row";
    colorRow.style.flexWrap = "wrap";
    const cl = document.createElement("label");
    cl.textContent = "线条颜色:";
    const selCol = document.createElement("select");
    selCol.className = "fs-format-cells__select";
    selCol.style.minWidth = "100px";
    const optAuto = document.createElement("option");
    optAuto.value = "auto";
    optAuto.textContent = "自动";
    const optCustom = document.createElement("option");
    optCustom.value = "custom";
    optCustom.textContent = "自定义";
    selCol.appendChild(optAuto);
    selCol.appendChild(optCustom);
    selCol.value = borderState.colorAuto ? "auto" : "custom";

    const colorInp = document.createElement("input");
    colorInp.type = "color";
    colorInp.style.width = "44px";
    colorInp.style.height = "26px";
    colorInp.style.padding = "0";
    colorInp.style.border = "1px solid #c8c6c4";
    colorInp.style.borderRadius = "4px";
    colorInp.disabled = borderState.colorAuto;
    colorInp.value = argb8ToCssHex6(borderState.colorArgb ?? "FF000000");

    selCol.addEventListener("change", () => {
      if (selCol.value === "auto") {
        borderState = { ...borderState, colorAuto: true, colorArgb: null };
      } else {
        borderState = {
          ...borderState,
          colorAuto: false,
          colorArgb: cssHexToFillArgb(colorInp.value),
        };
      }
      rebuildBorderPanel();
    });
    colorInp.addEventListener("input", () => {
      borderState = {
        ...borderState,
        colorAuto: false,
        colorArgb: cssHexToFillArgb(colorInp.value),
      };
      selCol.value = "custom";
      rebuildBorderPanel();
    });

    const moreColBtn = document.createElement("button");
    moreColBtn.type = "button";
    moreColBtn.className = "fs-format-cells__btn fs-format-cells__btn--secondary";
    moreColBtn.textContent = "其他颜色…";
    moreColBtn.addEventListener("click", () => {
      if (borderOtherColorBusy) {
        return;
      }
      borderOtherColorBusy = true;
      void (async () => {
        try {
          const picked = await showRibbonColorDialog(
            argb8ToCssHex6(borderState.colorArgb ?? "FF000000"),
          );
          if (picked !== null) {
            borderState = {
              ...borderState,
              colorAuto: false,
              colorArgb: cssHexToFillArgb(picked),
            };
            rebuildBorderPanel();
          }
        } finally {
          borderOtherColorBusy = false;
        }
      })();
    });

    colorRow.appendChild(cl);
    colorRow.appendChild(selCol);
    colorRow.appendChild(colorInp);
    colorRow.appendChild(moreColBtn);
    leftCol.appendChild(colorRow);

    const rightCol = document.createElement("div");
    rightCol.style.display = "flex";
    rightCol.style.flexDirection = "column";
    rightCol.style.gap = "12px";
    rightCol.style.minWidth = "0";

    const pres = document.createElement("div");
    pres.className = "fs-format-cells__border-presets";
    const pl = document.createElement("div");
    pl.className = "fs-format-cells__border-presets-label";
    pl.textContent = "预设:";
    const prow = document.createElement("div");
    prow.className = "fs-format-cells__border-presets-row";

    const mkPreset = (
      label: string,
      iconClass: string,
      onClick: () => void,
      disabled: boolean,
    ): void => {
      const pb = document.createElement("button");
      pb.type = "button";
      pb.className = "fs-format-cells__border-preset";
      pb.disabled = disabled;
      const ic = document.createElement("div");
      ic.className = `fs-format-cells__border-preset-icon ${iconClass}`;
      const tx = document.createElement("span");
      tx.textContent = label;
      pb.appendChild(ic);
      pb.appendChild(tx);
      pb.addEventListener("click", onClick);
      prow.appendChild(pb);
    };

    mkPreset(
      "无",
      "fs-format-cells__border-preset-icon--none",
      () => {
        borderState = {
          ...borderState,
          edges: {
            top: false,
            bottom: false,
            left: false,
            right: false,
            insideH: false,
            insideV: false,
          },
        };
        rebuildBorderPanel();
      },
      false,
    );

    mkPreset(
      "外边框",
      "fs-format-cells__border-preset-icon--outline",
      () => {
        borderState = {
          ...borderState,
          edges: {
            ...borderState.edges,
            top: true,
            bottom: true,
            left: true,
            right: true,
            insideH: false,
            insideV: false,
          },
        };
        rebuildBorderPanel();
      },
      false,
    );

    mkPreset(
      "内部",
      "fs-format-cells__border-preset-icon--inside",
      () => {
        borderState = {
          ...borderState,
          edges: {
            ...borderState.edges,
            top: false,
            bottom: false,
            left: false,
            right: false,
            insideH: true,
            insideV: true,
          },
        };
        rebuildBorderPanel();
      },
      !selectionMultiCell,
    );

    pres.appendChild(pl);
    pres.appendChild(prow);
    rightCol.appendChild(pres);

    const toggleEdge = (key: keyof FormatCellsBorderState["edges"]): void => {
      const line = FORMAT_CELLS_LINE_STYLES[borderState.lineStyleIndex];
      const hasKind = line?.kind != null;
      const next = { ...borderState.edges };
      if (!hasKind) {
        next[key] = false;
      } else {
        next[key] = !next[key];
      }
      borderState = { ...borderState, edges: next };
      rebuildBorderPanel();
    };

    const blab = document.createElement("div");
    blab.className = "fs-format-cells__border-field-label";
    blab.textContent = "边框:";

    const hasPen = FORMAT_CELLS_LINE_STYLES[borderState.lineStyleIndex]?.kind != null;
    const pvColor = formatCellsBorderPreviewColor(borderState);
    const pvSwatch = FORMAT_CELLS_LINE_STYLES[borderState.lineStyleIndex]?.swatch ?? "none";

    const pv = document.createElement("div");
    pv.className = "fs-format-cells__border-preview";
    pv.style.position = "relative";
    pv.style.overflow = "hidden";
    if (hasPen) {
      pv.appendChild(
        createFormatCellsBorderPreviewSvg({
          swatch: pvSwatch,
          colorCss: pvColor,
          top: borderState.edges.top,
          bottom: borderState.edges.bottom,
          left: borderState.edges.left,
          right: borderState.edges.right,
          insideH: borderState.edges.insideH,
          insideV: borderState.edges.insideV,
          multiCell: selectionMultiCell,
        }),
      );
    }
    const pvText = document.createElement("span");
    pvText.textContent = "文本";
    pvText.style.position = "relative";
    pvText.style.zIndex = "1";
    pv.appendChild(pvText);

    pv.addEventListener("click", (ev) => {
      const t = ev.currentTarget as HTMLElement;
      const r = t.getBoundingClientRect();
      const x = ev.clientX - r.left;
      const y = ev.clientY - r.top;
      const ew = 14;
      const { width: W, height: H } = r;
      if (y < ew) {
        toggleEdge("top");
      } else if (y > H - ew) {
        toggleEdge("bottom");
      } else if (x < ew) {
        toggleEdge("left");
      } else if (x > W - ew) {
        toggleEdge("right");
      } else if (selectionMultiCell) {
        const mx = x / W;
        const my = y / H;
        if (Math.abs(mx - 0.5) < 0.14) {
          toggleEdge("insideV");
        }
        if (Math.abs(my - 0.5) < 0.14) {
          toggleEdge("insideH");
        }
      }
    });

    const mkFab = (
      label: string,
      title: string,
      edge: keyof FormatCellsBorderState["edges"],
      disabled: boolean,
    ): HTMLButtonElement => {
      const fb = document.createElement("button");
      fb.type = "button";
      fb.className = "fs-format-cells__border-edge-fab";
      fb.textContent = label;
      fb.title = title;
      fb.disabled = disabled;
      fb.classList.toggle(
        "fs-format-cells__border-edge-fab--on",
        !disabled && borderState.edges[edge] && hasPen,
      );
      fb.addEventListener("click", () => {
        if (!disabled) {
          toggleEdge(edge);
        }
      });
      return fb;
    };

    const rowMid = document.createElement("div");
    rowMid.className = "fs-format-cells__border-preview-row";
    const vstack = document.createElement("div");
    vstack.style.display = "flex";
    vstack.style.flexDirection = "column";
    vstack.style.gap = "4px";
    vstack.appendChild(mkFab("上", "上框线", "top", false));
    vstack.appendChild(mkFab("—", "内部横线", "insideH", !selectionMultiCell));
    vstack.appendChild(mkFab("下", "下框线", "bottom", false));

    rowMid.appendChild(vstack);
    rowMid.appendChild(pv);

    const bottomCtr = document.createElement("div");
    bottomCtr.className = "fs-format-cells__border-bottom-controls";
    const d1 = document.createElement("button");
    d1.type = "button";
    d1.className = "fs-format-cells__border-edge-fab";
    d1.textContent = "╲";
    d1.title = "斜线边框（暂不支持）";
    d1.disabled = true;
    bottomCtr.appendChild(d1);
    bottomCtr.appendChild(mkFab("左", "左边框", "left", false));
    bottomCtr.appendChild(mkFab("│", "内部竖线", "insideV", !selectionMultiCell));
    bottomCtr.appendChild(mkFab("右", "右边框", "right", false));
    const d2 = document.createElement("button");
    d2.type = "button";
    d2.className = "fs-format-cells__border-edge-fab";
    d2.textContent = "╱";
    d2.title = "斜线边框（暂不支持）";
    d2.disabled = true;
    bottomCtr.appendChild(d2);

    const shell = document.createElement("div");
    shell.className = "fs-format-cells__border-preview-shell";
    shell.appendChild(blab);
    shell.appendChild(rowMid);
    shell.appendChild(bottomCtr);

    rightCol.appendChild(shell);

    borderRoot.appendChild(leftCol);
    borderRoot.appendChild(rightCol);
  };

  const syncFillSampleBox = (cv: HTMLCanvasElement): void => {
    const rect = cv.getBoundingClientRect();
    let w = Math.floor(rect.width);
    let h = Math.floor(rect.height);
    if (w < 2) {
      w = 320;
    }
    if (h < 2) {
      h = 88;
    }
    const dpr = window.devicePixelRatio || 1;
    cv.width = Math.max(2, Math.floor(w * dpr));
    cv.height = Math.max(2, Math.floor(h * dpr));
    const ctx = cv.getContext("2d");
    if (ctx === null) {
      return;
    }
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    const bg = fillState.bgArgb !== null ? argb8ToCssHex6(fillState.bgArgb) : "#ffffff";
    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, w, h);
    if (fillState.patternType !== "none") {
      const fgA =
        fillState.patternFgAuto || fillState.patternFgArgb === null
          ? "FF000000"
          : fillState.patternFgArgb;
      paintCellFillPatternOverlay(
        ctx,
        0,
        0,
        w,
        h,
        fillState.patternType,
        argb8ToCssHex6(fgA),
      );
    }
  };

  const rebuildFillPanel = (): void => {
    closeFillPopover();
    if (fillSampleResizeObserver !== null) {
      fillSampleResizeObserver.disconnect();
      fillSampleResizeObserver = null;
    }
    fillRoot.replaceChildren();

    const fillUiId = `fs-fcf-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

    const mountPalettePopover = (
      anchor: HTMLButtonElement,
      build: (pop: HTMLDivElement) => void,
    ): void => {
      closeFillPopover();
      const pop = document.createElement("div");
      pop.className = "fs-format-cells__fill-popover fs-color-menu";
      build(pop);
      document.body.appendChild(pop);
      const position = (): void => {
        const r = anchor.getBoundingClientRect();
        const pw = pop.offsetWidth;
        const left = Math.max(8, Math.min(r.left, window.innerWidth - pw - 8));
        pop.style.position = "fixed";
        pop.style.left = `${left}px`;
        pop.style.top = `${r.bottom + 4}px`;
      };
      requestAnimationFrame(position);
      const onDoc = (ev: PointerEvent): void => {
        const t = ev.target as Node | null;
        if (t !== null && (pop.contains(t) || anchor.contains(t))) {
          return;
        }
        closeFillPopover();
      };
      fillPopoverCleanup = (): void => {
        pop.remove();
        document.removeEventListener("pointerdown", onDoc, true);
      };
      setTimeout(() => document.addEventListener("pointerdown", onDoc, true), 0);
    };

    const top = document.createElement("div");
    top.className = "fs-format-cells__fill-top";

    const rowBg = document.createElement("div");
    rowBg.className = "fs-format-cells__fill-dd-row";
    const labBg = document.createElement("label");
    labBg.textContent = "背景色:";
    const btnBg = document.createElement("button");
    btnBg.id = `${fillUiId}-bg`;
    labBg.htmlFor = btnBg.id;
    btnBg.type = "button";
    btnBg.className = "fs-format-cells__fill-dd";
    const swBg = document.createElement("span");
    swBg.className =
      "fs-format-cells__fill-swatch" +
      (fillState.bgArgb === null ? " fs-format-cells__fill-swatch--empty" : "");
    if (fillState.bgArgb !== null) {
      swBg.style.backgroundColor = argb8ToCssHex6(fillState.bgArgb);
    }
    const txtBg = document.createElement("span");
    txtBg.className = "fs-format-cells__fill-dd-label";
    txtBg.textContent = fillState.bgArgb === null ? "无颜色" : "自定义颜色";
    const arBg = document.createElement("span");
    arBg.className = "fs-format-cells__fill-dd-arrow";
    btnBg.appendChild(swBg);
    btnBg.appendChild(txtBg);
    btnBg.appendChild(arBg);
    btnBg.addEventListener("click", () => {
      mountPalettePopover(btnBg, (pop) => {
        appendRibbonColorPaletteContent(pop, {
          themeHeading: "主题颜色",
          standardHeading: "标准色",
          includeNoneRow: true,
          onNone: () => {
            fillState = { ...fillState, bgArgb: null };
            closeFillPopover();
            rebuildFillPanel();
          },
          onPickHex: (hex: string) => {
            fillState = { ...fillState, bgArgb: cssHexToFillArgb(hex) };
            closeFillPopover();
            rebuildFillPanel();
          },
          onMoreColors: () => {
            closeFillPopover();
            void (async () => {
              const cur =
                fillState.bgArgb !== null ? argb8ToCssHex6(fillState.bgArgb) : "#ffffff";
              const picked = await showRibbonColorDialog(cur);
              if (picked !== null) {
                fillState = { ...fillState, bgArgb: cssHexToFillArgb(picked) };
                rebuildFillPanel();
              }
            })();
          },
        });
      });
    });
    rowBg.appendChild(labBg);
    rowBg.appendChild(btnBg);

    const rowPC = document.createElement("div");
    rowPC.className = "fs-format-cells__fill-dd-row";
    const labPC = document.createElement("label");
    labPC.textContent = "图案颜色:";
    const btnPC = document.createElement("button");
    btnPC.id = `${fillUiId}-patfg`;
    labPC.htmlFor = btnPC.id;
    btnPC.type = "button";
    btnPC.className = "fs-format-cells__fill-dd";
    const swPC = document.createElement("span");
    swPC.className = "fs-format-cells__fill-swatch";
    const txtPC = document.createElement("span");
    txtPC.className = "fs-format-cells__fill-dd-label";
    if (fillState.patternFgAuto || fillState.patternFgArgb === null) {
      swPC.classList.add("fs-format-cells__fill-swatch--empty");
      txtPC.textContent = "自动";
    } else {
      swPC.style.backgroundColor = argb8ToCssHex6(fillState.patternFgArgb);
      txtPC.textContent = "自定义颜色";
    }
    const arPC = document.createElement("span");
    arPC.className = "fs-format-cells__fill-dd-arrow";
    btnPC.appendChild(swPC);
    btnPC.appendChild(txtPC);
    btnPC.appendChild(arPC);
    btnPC.addEventListener("click", () => {
      mountPalettePopover(btnPC, (pop) => {
        const autoBtn = document.createElement("button");
        autoBtn.type = "button";
        autoBtn.className = "fs-format-cells__fill-auto-btn";
        autoBtn.textContent = "自动";
        autoBtn.addEventListener("click", () => {
          fillState = { ...fillState, patternFgAuto: true, patternFgArgb: null };
          closeFillPopover();
          rebuildFillPanel();
        });
        pop.appendChild(autoBtn);
        appendRibbonColorPaletteContent(pop, {
          themeHeading: "主题颜色",
          standardHeading: "标准色",
          includeNoneRow: false,
          onPickHex: (hex: string) => {
            fillState = {
              ...fillState,
              patternFgAuto: false,
              patternFgArgb: cssHexToFillArgb(hex),
            };
            closeFillPopover();
            rebuildFillPanel();
          },
          onMoreColors: () => {
            closeFillPopover();
            void (async () => {
              const cur =
                fillState.patternFgArgb !== null && !fillState.patternFgAuto
                  ? argb8ToCssHex6(fillState.patternFgArgb)
                  : "#323130";
              const picked = await showRibbonColorDialog(cur);
              if (picked !== null) {
                fillState = {
                  ...fillState,
                  patternFgAuto: false,
                  patternFgArgb: cssHexToFillArgb(picked),
                };
                rebuildFillPanel();
              }
            })();
          },
        });
      });
    });
    rowPC.appendChild(labPC);
    rowPC.appendChild(btnPC);

    const rowPS = document.createElement("div");
    rowPS.className = "fs-format-cells__fill-dd-row";
    const labPS = document.createElement("label");
    labPS.textContent = "图案样式:";
    const btnPS = document.createElement("button");
    btnPS.id = `${fillUiId}-pat`;
    labPS.htmlFor = btnPS.id;
    btnPS.type = "button";
    btnPS.className = "fs-format-cells__fill-dd";
    const mini = document.createElement("canvas");
    mini.width = 40;
    mini.height = 28;
    mini.style.width = "40px";
    mini.style.height = "28px";
    mini.style.pointerEvents = "none";
    const mtx = mini.getContext("2d");
    if (mtx !== null) {
      mtx.fillStyle = "#ffffff";
      mtx.fillRect(0, 0, 40, 28);
      if (fillState.patternType !== "none") {
        const fga =
          fillState.patternFgAuto || fillState.patternFgArgb === null
            ? "FF000000"
            : fillState.patternFgArgb;
        paintCellFillPatternOverlay(
          mtx,
          0,
          0,
          40,
          28,
          fillState.patternType,
          argb8ToCssHex6(fga),
        );
      }
    }
    const txtPS = document.createElement("span");
    txtPS.className = "fs-format-cells__fill-dd-label";
    txtPS.textContent = fillState.patternType === "none" ? "无" : "图案";
    const arPS = document.createElement("span");
    arPS.className = "fs-format-cells__fill-dd-arrow";
    btnPS.appendChild(mini);
    btnPS.appendChild(txtPS);
    btnPS.appendChild(arPS);
    btnPS.addEventListener("click", () => {
      mountPalettePopover(btnPS, (pop) => {
        const grid = document.createElement("div");
        grid.className = "fs-format-cells__fill-pattern-grid";
        for (const pid of FORMAT_CELLS_PATTERN_GRID_ORDER) {
          const pb = document.createElement("button");
          pb.type = "button";
          pb.className = "fs-format-cells__fill-pattern-btn";
          if (pid === fillState.patternType) {
            pb.classList.add("fs-format-cells__fill-pattern-btn--on");
          }
          const cvs = document.createElement("canvas");
          cvs.width = 32;
          cvs.height = 24;
          cvs.style.pointerEvents = "none";
          const cx = cvs.getContext("2d");
          if (cx !== null) {
            cx.fillStyle = "#ffffff";
            cx.fillRect(0, 0, 32, 24);
            if (pid !== "none") {
              paintCellFillPatternOverlay(cx, 0, 0, 32, 24, pid, "#323130");
            }
          }
          pb.appendChild(cvs);
          grid.appendChild(pb);
          pb.addEventListener("click", () => {
            fillState = {
              ...fillState,
              patternType: pid,
              patternFgAuto: pid === "none" ? true : fillState.patternFgAuto,
              patternFgArgb: pid === "none" ? null : fillState.patternFgArgb,
            };
            closeFillPopover();
            rebuildFillPanel();
          });
        }
        pop.appendChild(grid);
      });
    });
    rowPS.appendChild(labPS);
    rowPS.appendChild(btnPS);

    top.appendChild(rowBg);
    top.appendChild(rowPC);
    top.appendChild(rowPS);

    const sampleWrap = document.createElement("div");
    sampleWrap.className = "fs-format-cells__fill-sample-wrap";
    const sampleLab = document.createElement("label");
    sampleLab.textContent = "示例";
    const sampleCv = document.createElement("canvas");
    sampleCv.className = "fs-format-cells__fill-sample";

    fillRoot.appendChild(top);
    fillRoot.appendChild(sampleWrap);
    sampleWrap.appendChild(sampleLab);
    sampleWrap.appendChild(sampleCv);

    fillSampleResizeObserver = new ResizeObserver(() => {
      syncFillSampleBox(sampleCv);
    });
    fillSampleResizeObserver.observe(sampleCv);
    requestAnimationFrame(() => syncFillSampleBox(sampleCv));
  };

  const renderBody = (): void => {
    closeFillPopover();
    body.replaceChildren();
    if (mainTab === "number") {
      body.appendChild(numberLayout);
      syncCategoryHighlight();
      rebuildNumberDetail();
    } else if (mainTab === "alignment") {
      rebuildAlignmentPanel();
      body.appendChild(alignmentRoot);
    } else if (mainTab === "font") {
      rebuildFontPanel();
      body.appendChild(fontRoot);
    } else if (mainTab === "border") {
      rebuildBorderPanel();
      body.appendChild(borderRoot);
    } else if (mainTab === "fill") {
      rebuildFillPanel();
      body.appendChild(fillRoot);
    } else {
      rebuildProtectionPanel();
      body.appendChild(protectionRoot);
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
    closeFillPopover();
    if (fillSampleResizeObserver !== null) {
      fillSampleResizeObserver.disconnect();
      fillSampleResizeObserver = null;
    }
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
    const fontPatch = fontPatchIfChanged(fontState, initialFontState);
    const fillPatch = fillPatchIfChanged(fillState, initialFillState);
    const protectionPatch = protectionPatchIfChanged(protectionUi, initialProtectionState);
    const basePatch: CellStylePatch = {
      ...alignPatch,
      ...fontPatch,
      ...fillPatch,
      ...protectionPatch,
      numberFormat: code === null ? null : code,
    };
    const borderPatch = {
      apply: !formatCellsBorderStateEqual(borderState, initialBorderState),
      state: borderState,
    } as const;
    const mergeCellsChanged = alignState.mergeCells !== initialAlignState.mergeCells;
    if (onApply !== undefined) {
      onApply({
        basePatch,
        border: borderPatch,
        mergeCellsChanged,
        mergeCells: alignState.mergeCells,
      });
    } else {
      flex.applyFormatCellsDialogOk(basePatch, borderPatch);
      if (mergeCellsChanged) {
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
    }
    close();
  });

  syncTabHighlight();
  renderBody();

  return overlay;
}

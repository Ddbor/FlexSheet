import type { CellStyle } from "@flexsheet/core";

/** 「开始」→「数字」组：下拉项与 Excel 兼容的格式码。 */
export interface RibbonNumberFormatPreset {
  readonly id: string;
  readonly label: string;
  /** 空字符串表示常规（清除 numberFormat）。 */
  readonly format: string;
}

export const RIBBON_NUMBER_FORMAT_PRESETS: readonly RibbonNumberFormatPreset[] = [
  { id: "home.number.format.general", label: "常规", format: "" },
  { id: "home.number.format.number", label: "数字", format: "0.00" },
  { id: "home.number.format.currency", label: "货币", format: "¥#,##0.00" },
  {
    id: "home.number.format.accounting",
    label: "会计专用",
    format: '_ * #,##0.00_ ;_ * (#,##0.00)_ ;_ * "-"??_ ;_ @_ @',
  },
  { id: "home.number.format.shortDate", label: "短日期", format: "yyyy/m/d" },
  { id: "home.number.format.longDate", label: "长日期", format: 'yyyy"年"m"月"d"日"' },
  { id: "home.number.format.time", label: "时间", format: "h:mm:ss" },
  { id: "home.number.format.percent", label: "百分比", format: "0%" },
  { id: "home.number.format.fraction", label: "分数", format: "# ?/?" },
  { id: "home.number.format.scientific", label: "科学记数", format: "0.00E+00" },
  { id: "home.number.format.text", label: "文本", format: "@" },
];

const PRESET_BY_ID = new Map(RIBBON_NUMBER_FORMAT_PRESETS.map((p) => [p.id, p] as const));

export function getNumberFormatPresetByCommandId(id: string): RibbonNumberFormatPreset | undefined {
  return PRESET_BY_ID.get(id);
}

/** 与活动单元格 `numberFormat` 对齐的下拉展示文案。 */
export interface RibbonHomeNumberFormatChromeState {
  readonly categoryLabel: string;
}

export function cellStyleToRibbonHomeNumberFormat(style: CellStyle | null): RibbonHomeNumberFormatChromeState {
  const nf = style?.numberFormat?.trim() ?? "";
  for (const p of RIBBON_NUMBER_FORMAT_PRESETS) {
    if (p.format === nf) {
      return { categoryLabel: p.label };
    }
  }
  if (nf === "") {
    return { categoryLabel: "常规" };
  }
  return { categoryLabel: "自定义" };
}

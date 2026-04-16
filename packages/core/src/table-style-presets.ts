import type { CellBorderSide, CellStyle } from "./cell.js";
import { normalizeSelectionRange, type SelectionRange } from "./selection-range.js";

/** 7 列主题色（与 Excel 表格样式库列顺序一致：灰 / 蓝 / 橙 / 银灰 / 黄 / 浅蓝 / 绿） */
export interface TableAccentPalette {
  readonly border: string;
  readonly lightHeader: string;
  readonly lightStripe: string;
  readonly mediumHeader: string;
  readonly mediumStripeA: string;
  readonly mediumStripeB: string;
  readonly darkSolid: string;
  readonly darkHeader: string;
  readonly darkStripeA: string;
  readonly darkStripeB: string;
}

export const TABLE_ACCENT_PALETTES: readonly TableAccentPalette[] = [
  {
    border: "FF7F7F7F",
    lightHeader: "FFD9D9D9",
    lightStripe: "FFF2F2F2",
    mediumHeader: "FFA6A6A6",
    mediumStripeA: "FFE7E6E6",
    mediumStripeB: "FFFFFFFF",
    darkSolid: "FF3F3F3F",
    darkHeader: "FF525252",
    darkStripeA: "FF4A4A4A",
    darkStripeB: "FF5F5F5F",
  },
  {
    border: "FF4472C4",
    lightHeader: "FFBDD7EE",
    lightStripe: "FFDDEBF7",
    mediumHeader: "FF5B9BD5",
    mediumStripeA: "FFDDEBF7",
    mediumStripeB: "FFFFFFFF",
    darkSolid: "FF1F4E78",
    darkHeader: "FF2E5597",
    darkStripeA: "FF2F5597",
    darkStripeB: "FF3F63A3",
  },
  {
    border: "FFC65911",
    lightHeader: "FFF8CBAD",
    lightStripe: "FFFCE4D6",
    mediumHeader: "FFF4B084",
    mediumStripeA: "FFFCE4D6",
    mediumStripeB: "FFFFFFFF",
    darkSolid: "FF833C0C",
    darkHeader: "FFA65E1F",
    darkStripeA: "FFB35A12",
    darkStripeB: "FFC46828",
  },
  {
    border: "FF7F7F7F",
    lightHeader: "FFDCE6F1",
    lightStripe: "FFF2F2F2",
    mediumHeader: "FFB4C6E7",
    mediumStripeA: "FFE7E6E6",
    mediumStripeB: "FFFFFFFF",
    darkSolid: "FF4F4F4F",
    darkHeader: "FF6B6B6B",
    darkStripeA: "FF5C5C5C",
    darkStripeB: "FF6E6E6E",
  },
  {
    border: "FFBF8F00",
    lightHeader: "FFFFE699",
    lightStripe: "FFFFF2CC",
    mediumHeader: "FFFFC000",
    mediumStripeA: "FFFFF2CC",
    mediumStripeB: "FFFFFFFF",
    darkSolid: "FF806000",
    darkHeader: "FFA67C00",
    darkStripeA: "FFB38F00",
    darkStripeB: "FFC19E1A",
  },
  {
    border: "FF2E75B6",
    lightHeader: "FF9BC2E6",
    lightStripe: "FFDDEBF7",
    mediumHeader: "FF9BC2E6",
    mediumStripeA: "FFDDEBF7",
    mediumStripeB: "FFFFFFFF",
    darkSolid: "FF1F4E78",
    darkHeader: "FF2E75B6",
    darkStripeA: "FF3A6FB0",
    darkStripeB: "FF4A7FC0",
  },
  {
    border: "FF548235",
    lightHeader: "FFC6E0B4",
    lightStripe: "FFE2EFDA",
    mediumHeader: "FFA9D08E",
    mediumStripeA: "FFE2EFDA",
    mediumStripeB: "FFFFFFFF",
    darkSolid: "FF375623",
    darkHeader: "FF548235",
    darkStripeA: "FF5F8F40",
    darkStripeB: "FF6FA050",
  },
];

export type TableStyleSection = "light" | "medium" | "dark";

export interface ParsedTableStyleCommand {
  readonly section: TableStyleSection;
  readonly row: number;
  readonly col: number;
}

const TABLE_STYLE_CMD_RE = /^home\.style\.table\.(light|medium|dark)\.r(\d+)c(\d+)$/;

export function parseTableStyleRibbonCommand(id: string): ParsedTableStyleCommand | null {
  const m = TABLE_STYLE_CMD_RE.exec(id);
  if (m === null) {
    return null;
  }
  const section = m[1] as TableStyleSection;
  const row = Number(m[2]);
  const col = Number(m[3]);
  if (!Number.isInteger(row) || !Number.isInteger(col) || col < 0 || col >= TABLE_ACCENT_PALETTES.length) {
    return null;
  }
  const maxRow = section === "light" ? 2 : section === "medium" ? 3 : 1;
  if (row < 0 || row > maxRow) {
    return null;
  }
  return { section, row, col };
}

/**
 * 首行是否使用「表头」强调色（与 Excel 各缩略图一致：仅斑马纹/无表头样式时首行也参与斑马纹）。
 */
export function tableStyleUsesDistinctHeaderRow(parsed: ParsedTableStyleCommand): boolean {
  const { section, row } = parsed;
  if (section === "light") {
    return row === 0;
  }
  if (section === "medium") {
    return row !== 1;
  }
  return section === "dark" && row === 1;
}

function thinBorder(color: string): CellBorderSide {
  return { kind: "thin", colorArgb: color };
}

function mediumBorder(color: string): CellBorderSide {
  return { kind: "medium", colorArgb: color };
}

function uniformBorders(side: CellBorderSide): Pick<CellStyle, "borderTop" | "borderLeft" | "borderBottom" | "borderRight"> {
  return { borderTop: side, borderLeft: side, borderBottom: side, borderRight: side };
}

/**
 * 根据套用表格格式预设计算单元格完整样式（含四边边框）。
 */
export function computeTableFormatCellStyle(
  parsed: ParsedTableStyleCommand,
  palette: TableAccentPalette,
  range: SelectionRange,
  hasHeaders: boolean,
  row: number,
  _col: number,
): CellStyle {
  const n = normalizeSelectionRange(range);
  const { section, row: patRow } = parsed;
  const borderThin = thinBorder(palette.border);
  const borderMed = mediumBorder(palette.border);
  const fgDark = "FF000000";
  const fgLight = "FFFFFFFF";

  const distinctHeader = tableStyleUsesDistinctHeaderRow(parsed);
  const isHeaderRow = hasHeaders && row === n.startRow && distinctHeader;
  const bodyIndex = hasHeaders && distinctHeader ? row - n.startRow - 1 : row - n.startRow;
  const bandOdd = bodyIndex >= 0 && bodyIndex % 2 === 0;

  const base = (): CellStyle => ({
    fontSizePt: 11,
    fgArgb: fgDark,
    hAlign: "center",
    vAlign: "middle",
  });

  if (section === "light") {
    if (patRow === 0) {
      if (isHeaderRow) {
        return {
          ...base(),
          fillArgb: palette.lightHeader,
          ...uniformBorders(borderThin),
        };
      }
      return {
        ...base(),
        fillArgb: "FFFFFFFF",
        ...uniformBorders(borderThin),
      };
    }
    if (patRow === 1) {
      return {
        ...base(),
        fillArgb: "FFFFFFFF",
        ...uniformBorders(borderThin),
      };
    }
    return {
      ...base(),
      fillArgb: bandOdd ? palette.lightStripe : "FFFFFFFF",
      ...uniformBorders(borderThin),
    };
  }

  if (section === "medium") {
    if (patRow === 0) {
      if (isHeaderRow) {
        return {
          ...base(),
          fillArgb: palette.mediumHeader,
          ...uniformBorders(borderThin),
        };
      }
      return {
        ...base(),
        fillArgb: "FFFFFFFF",
        ...uniformBorders(borderThin),
      };
    }
    if (patRow === 1) {
      return {
        ...base(),
        fillArgb: bandOdd ? palette.mediumStripeA : palette.mediumStripeB,
        ...uniformBorders(borderThin),
      };
    }
    if (patRow === 2) {
      if (isHeaderRow) {
        return {
          ...base(),
          fillArgb: palette.mediumHeader,
          ...uniformBorders(borderThin),
        };
      }
      return {
        ...base(),
        fillArgb: bandOdd ? palette.mediumStripeA : palette.mediumStripeB,
        ...uniformBorders(borderThin),
      };
    }
    if (isHeaderRow) {
      return {
        ...base(),
        fillArgb: palette.mediumHeader,
        ...uniformBorders(borderMed),
      };
    }
    return {
      ...base(),
      fillArgb: bandOdd ? palette.mediumStripeA : palette.mediumStripeB,
      ...uniformBorders(borderMed),
    };
  }

  if (patRow === 0) {
    return {
      ...base(),
      fgArgb: fgLight,
      fillArgb: palette.darkSolid,
      ...uniformBorders(thinBorder(palette.border)),
    };
  }
  if (isHeaderRow) {
    return {
      ...base(),
      fgArgb: fgLight,
      fillArgb: palette.darkHeader,
      ...uniformBorders(borderThin),
    };
  }
  return {
    ...base(),
    fgArgb: fgLight,
    fillArgb: bandOdd ? palette.darkStripeA : palette.darkStripeB,
    ...uniformBorders(borderThin),
  };
}

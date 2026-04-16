import type { CellStylePatch } from "@flexsheet/core";
import type { FlexSheetLike } from "./ribbon-types.js";

/** 与 Excel「常规」一致：清除主要显式格式（含边框与数字格式）。 */
export const CELL_STYLE_NORMAL_PATCH: CellStylePatch = {
  bold: null,
  italic: null,
  underline: null,
  strikethrough: null,
  fontScript: null,
  fontFamily: null,
  fontSizePt: null,
  fgArgb: null,
  fillArgb: null,
  fillPatternType: null,
  fillPatternFgArgb: null,
  hAlign: null,
  vAlign: null,
  indentLevel: null,
  wrapText: null,
  textRotationDegrees: null,
  shrinkToFit: null,
  textOrientation: null,
  borderTop: null,
  borderLeft: null,
  borderBottom: null,
  borderRight: null,
  numberFormat: null,
};

const THEME_ACCENT_HEX = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"] as const;
const THEME_ROW_TINT = [0.2, 0.4, 0.6, 1.0] as const;

function blendAccentWithWhite(hex: string, t: number): string {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const mix = (c: number): number => Math.round(255 * (1 - t) + c * t);
  const toHex = (n: number): string => n.toString(16).padStart(2, "0").toUpperCase();
  return `FF${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
}

const THIN_GREY = "FFBFBFBF";
const BLACK = "FF000000";

function clearDecorationPatch(): Pick<
  CellStylePatch,
  | "italic"
  | "underline"
  | "strikethrough"
  | "fontScript"
  | "textOrientation"
  | "textRotationDegrees"
  | "numberFormat"
> {
  return {
    italic: null,
    underline: null,
    strikethrough: null,
    fontScript: null,
    textOrientation: null,
    textRotationDegrees: null,
    numberFormat: null,
  };
}

/**
 * 处理「开始 → 单元格样式」库命令；已识别则返回 true。
 */
export function applyRibbonCellStyleCommand(fs: FlexSheetLike, id: string): boolean {
  const themeM = /^home\.style\.cell\.theme\.r([0-3])\.c([0-5])$/.exec(id);
  if (themeM !== null) {
    const row = Number(themeM[1]);
    const col = Number(themeM[2]);
    const accent = THEME_ACCENT_HEX[col];
    const t = THEME_ROW_TINT[row];
    const fillArgb = blendAccentWithWhite(accent, t);
    const solid = row === 3;
    fs.applySelectionStylePatch({
      ...CELL_STYLE_NORMAL_PATCH,
      ...clearDecorationPatch(),
      bold: solid ? true : null,
      fontSizePt: null,
      fgArgb: solid ? "FFFFFFFF" : "FF000000",
      fillArgb,
      fillPatternType: null,
      fillPatternFgArgb: null,
    });
    return true;
  }

  let patch: CellStylePatch | null = null;

  switch (id) {
    case "home.style.cell.normal":
      patch = { ...CELL_STYLE_NORMAL_PATCH };
      break;
    case "home.style.cell.good":
      patch = {
        ...clearDecorationPatch(),
        bold: true,
        fillArgb: "FFC6EFCE",
        fgArgb: "FF006100",
        borderTop: { kind: "thin", colorArgb: THIN_GREY },
        borderLeft: { kind: "thin", colorArgb: THIN_GREY },
        borderBottom: { kind: "thin", colorArgb: THIN_GREY },
        borderRight: { kind: "thin", colorArgb: THIN_GREY },
        fontSizePt: null,
      };
      break;
    case "home.style.cell.bad":
      patch = {
        ...clearDecorationPatch(),
        bold: true,
        fillArgb: "FFFFC7CE",
        fgArgb: "FF9C0006",
        borderTop: { kind: "thin", colorArgb: THIN_GREY },
        borderLeft: { kind: "thin", colorArgb: THIN_GREY },
        borderBottom: { kind: "thin", colorArgb: THIN_GREY },
        borderRight: { kind: "thin", colorArgb: THIN_GREY },
        fontSizePt: null,
      };
      break;
    case "home.style.cell.neutral":
      patch = {
        ...clearDecorationPatch(),
        bold: false,
        fillArgb: "FFFFEB9C",
        fgArgb: "FF9C6500",
        borderTop: { kind: "thin", colorArgb: THIN_GREY },
        borderLeft: { kind: "thin", colorArgb: THIN_GREY },
        borderBottom: { kind: "thin", colorArgb: THIN_GREY },
        borderRight: { kind: "thin", colorArgb: THIN_GREY },
        fontSizePt: null,
      };
      break;
    case "home.style.cell.calculation":
      patch = {
        ...clearDecorationPatch(),
        bold: true,
        fillArgb: "FFF2F2F2",
        fgArgb: "FFC65911",
        borderTop: { kind: "thin", colorArgb: THIN_GREY },
        borderLeft: { kind: "thin", colorArgb: THIN_GREY },
        borderBottom: { kind: "thin", colorArgb: THIN_GREY },
        borderRight: { kind: "thin", colorArgb: THIN_GREY },
        fontSizePt: null,
      };
      break;
    case "home.style.cell.checkCell":
      patch = {
        ...clearDecorationPatch(),
        bold: true,
        fillArgb: "FF595959",
        fgArgb: "FFFFFFFF",
        borderTop: { kind: "thick", colorArgb: BLACK },
        borderLeft: { kind: "thick", colorArgb: BLACK },
        borderBottom: { kind: "thick", colorArgb: BLACK },
        borderRight: { kind: "thick", colorArgb: BLACK },
        fontSizePt: null,
      };
      break;
    case "home.style.cell.explanatory":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        italic: true,
        fgArgb: "FF7F7F7F",
        fillArgb: null,
        fontSizePt: null,
      };
      break;
    case "home.style.cell.warningText":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        fgArgb: "FFFF0000",
        fillArgb: null,
        fontSizePt: null,
      };
      break;
    case "home.style.cell.linkedCell":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        fgArgb: "FFFF6600",
        underline: "double",
        fillArgb: null,
        fontSizePt: null,
      };
      break;
    case "home.style.cell.output":
      patch = {
        ...clearDecorationPatch(),
        bold: true,
        fillArgb: "FFF2F2F2",
        fgArgb: "FF000000",
        borderTop: { kind: "thin", colorArgb: BLACK },
        borderLeft: { kind: "thin", colorArgb: BLACK },
        borderBottom: { kind: "thin", colorArgb: BLACK },
        borderRight: { kind: "thin", colorArgb: BLACK },
        fontSizePt: null,
      };
      break;
    case "home.style.cell.input":
      patch = {
        ...clearDecorationPatch(),
        fillArgb: "FFFDE9D9",
        fgArgb: "FF974706",
        borderTop: { kind: "thin", colorArgb: THIN_GREY },
        borderLeft: { kind: "thin", colorArgb: THIN_GREY },
        borderBottom: { kind: "thin", colorArgb: THIN_GREY },
        borderRight: { kind: "thin", colorArgb: THIN_GREY },
        fontSizePt: null,
      };
      break;
    case "home.style.cell.note":
      patch = {
        ...clearDecorationPatch(),
        fillArgb: "FFFFF2CC",
        fgArgb: "FF000000",
        borderTop: { kind: "thin", colorArgb: THIN_GREY },
        borderLeft: { kind: "thin", colorArgb: THIN_GREY },
        borderBottom: { kind: "thin", colorArgb: THIN_GREY },
        borderRight: { kind: "thin", colorArgb: THIN_GREY },
        fontSizePt: null,
      };
      break;
    case "home.style.cell.title":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        bold: true,
        fontSizePt: 18,
        fgArgb: "FF1F497D",
        fillArgb: null,
      };
      break;
    case "home.style.cell.heading1":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        bold: true,
        fontSizePt: 14,
        fgArgb: "FF1F497D",
        borderBottom: { kind: "thick", colorArgb: "FF4F81BD" },
        borderTop: null,
        borderLeft: null,
        borderRight: null,
      };
      break;
    case "home.style.cell.heading2":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        bold: true,
        fontSizePt: 13,
        fgArgb: "FF1F497D",
        borderBottom: { kind: "medium", colorArgb: "FF8EAADC" },
        borderTop: null,
        borderLeft: null,
        borderRight: null,
      };
      break;
    case "home.style.cell.heading3":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        bold: true,
        fontSizePt: 12,
        fgArgb: "FF1F497D",
        borderBottom: { kind: "thin", colorArgb: "FF8EAADC" },
        borderTop: null,
        borderLeft: null,
        borderRight: null,
      };
      break;
    case "home.style.cell.heading4":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        bold: true,
        fontSizePt: 11,
        fgArgb: "FF1F497D",
      };
      break;
    case "home.style.cell.total":
      patch = {
        ...clearDecorationPatch(),
        bold: true,
        fgArgb: "FF000000",
        fillArgb: null,
        borderTop: { kind: "thin", colorArgb: BLACK },
        borderBottom: { kind: "double", colorArgb: BLACK },
        borderLeft: null,
        borderRight: null,
        fontSizePt: null,
      };
      break;
    case "home.style.cell.num.percent":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        numberFormat: "0%",
      };
      break;
    case "home.style.cell.num.currency":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        numberFormat: "¥#,##0.00",
      };
      break;
    case "home.style.cell.num.currency0":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        numberFormat: "¥#,##0",
      };
      break;
    case "home.style.cell.num.comma":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        numberFormat: "#,##0.00",
      };
      break;
    case "home.style.cell.num.comma0":
      patch = {
        ...CELL_STYLE_NORMAL_PATCH,
        numberFormat: "#,##0",
      };
      break;
    case "home.style.cell.newStyle": {
      if (fs.openFormatCellsDialog === undefined) {
        return false;
      }
      fs.openFormatCellsDialog();
      return true;
    }
    default:
      return false;
  }

  if (patch !== null) {
    fs.applySelectionStylePatch(patch);
    return true;
  }
  return false;
}

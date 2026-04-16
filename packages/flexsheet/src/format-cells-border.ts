import {
  applyCellStylePatch,
  normalizeSelectionRange,
  type CellBorderKind,
  type CellBorderLinePattern,
  type CellBorderSide,
  type CellStyle,
  type CellStylePatch,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";

import { cloneCellStyle } from "./border-ribbon-preset.js";

/**
 * 与 Excel 边框线型网格一致的 2 列 × 7 行顺序（先行后列：每行左细、右粗）。
 * `swatch` 仅用于对话框 SVG 预览；`kind` 为写入单元格后在 Canvas 上绘制的近似线宽。
 */
export type FormatCellsLineSwatchId =
  | "none"
  | "mediumDashDotDot"
  | "hairlineDots"
  | "slantedDash"
  | "shortDash"
  | "thickDash"
  | "dashDot"
  | "thickDashDot"
  | "dashDotDot"
  | "mediumSolid"
  | "mediumDash"
  | "thickSolid"
  | "thinSolid"
  | "doubleLine";

/** 先行后列：与 Excel 线型网格行对齐（含「无」与粗线型成对）。 */
export const FORMAT_CELLS_LINE_STYLES: readonly {
  readonly swatch: FormatCellsLineSwatchId;
  readonly kind: CellBorderKind | null;
}[] = [
  { swatch: "none", kind: null },
  { swatch: "mediumDashDotDot", kind: "medium" },
  { swatch: "hairlineDots", kind: "hairline" },
  { swatch: "thickDashDot", kind: "thick" },
  { swatch: "shortDash", kind: "thin" },
  { swatch: "thickDash", kind: "medium" },
  { swatch: "dashDotDot", kind: "thin" },
  { swatch: "mediumSolid", kind: "medium" },
  { swatch: "dashDot", kind: "thin" },
  { swatch: "mediumSolid", kind: "medium" },
  { swatch: "mediumDash", kind: "medium" },
  { swatch: "thickSolid", kind: "thick" },
  { swatch: "thinSolid", kind: "thin" },
  { swatch: "doubleLine", kind: "double" },
];

export interface FormatCellsBorderEdgeFlags {
  top: boolean;
  bottom: boolean;
  left: boolean;
  right: boolean;
  insideH: boolean;
  insideV: boolean;
}

export interface FormatCellsBorderState {
  lineStyleIndex: number;
  colorAuto: boolean;
  /** 非自动时 8 位 ARGB */
  colorArgb: string | null;
  edges: FormatCellsBorderEdgeFlags;
}

export function formatCellsBorderStateEqual(
  a: FormatCellsBorderState,
  b: FormatCellsBorderState,
): boolean {
  return (
    a.lineStyleIndex === b.lineStyleIndex &&
    a.colorAuto === b.colorAuto &&
    a.colorArgb === b.colorArgb &&
    a.edges.top === b.edges.top &&
    a.edges.bottom === b.edges.bottom &&
    a.edges.left === b.edges.left &&
    a.edges.right === b.edges.right &&
    a.edges.insideH === b.edges.insideH &&
    a.edges.insideV === b.edges.insideV
  );
}

function effectiveColorArgb(state: FormatCellsBorderState): string | undefined {
  if (state.colorAuto) {
    return undefined;
  }
  const t = state.colorArgb?.trim();
  if (t !== undefined && t !== "" && /^[\dA-Fa-f]{8}$/i.test(t)) {
    return t.toUpperCase();
  }
  return undefined;
}

function makePenSide(
  kind: CellBorderKind,
  colorArgb: string | undefined,
  swatch: FormatCellsLineSwatchId,
): CellBorderSide {
  const linePattern: CellBorderLinePattern | undefined =
    swatch === "none" ? undefined : (swatch as CellBorderLinePattern);
  if (linePattern === undefined) {
    return colorArgb === undefined ? { kind } : { kind, colorArgb };
  }
  return colorArgb === undefined
    ? { kind, linePattern }
    : { kind, colorArgb, linePattern };
}

/**
 * 根据「设置单元格格式」边框选项卡状态，计算选区内某一主格应用后的完整样式（先清空四边再按几何写入）。
 */
export function computeStyleForFormatCellsBorder(
  before: CellStyle | null,
  sheet: Worksheet,
  range: SelectionRange,
  row: number,
  col: number,
  state: FormatCellsBorderState,
): CellStyle | null {
  const n = normalizeSelectionRange(range);
  const style = FORMAT_CELLS_LINE_STYLES[state.lineStyleIndex];
  const penKind = style?.kind ?? null;

  const mid0 = cloneCellStyle(before);
  const mid: CellStyle = mid0 !== null ? { ...mid0 } : {};
  delete mid.borderTop;
  delete mid.borderLeft;
  delete mid.borderBottom;
  delete mid.borderRight;

  if (penKind === null) {
    return Object.keys(mid).length > 0 ? mid : null;
  }

  const colorArgb = effectiveColorArgb(state);
  const pen = makePenSide(penKind, colorArgb, style.swatch);

  const info = sheet.getMergedRectInfo(row, col);
  const ar = info.anchorRow;
  const ac = info.anchorCol;
  const endR = ar + info.rowSpan - 1;
  const endC = ac + info.colSpan - 1;

  const { edges } = state;
  const multiRow = n.endRow > n.startRow;
  const multiCol = n.endCol > n.startCol;

  if (edges.top && ar === n.startRow) {
    mid.borderTop = pen;
  }
  if (edges.bottom && endR === n.endRow) {
    mid.borderBottom = pen;
  }
  if (edges.left && ac === n.startCol) {
    mid.borderLeft = pen;
  }
  if (edges.right && endC === n.endCol) {
    mid.borderRight = pen;
  }

  if (multiRow && edges.insideH && ar < n.endRow) {
    mid.borderBottom = pen;
  }
  if (multiCol && edges.insideV && ac < n.endCol) {
    mid.borderRight = pen;
  }

  return Object.keys(mid).length > 0 ? mid : null;
}

/** 从已有边框映射到线型网格索引（优先 linePattern，否则按 kind）。 */
function inferLineStyleIndexFromBorderSide(side: CellBorderSide | undefined): number {
  if (side === undefined) {
    return 12;
  }
  if (side.linePattern !== undefined) {
    const idx = FORMAT_CELLS_LINE_STYLES.findIndex((e) => e.swatch === side.linePattern);
    if (idx >= 0) {
      return idx;
    }
  }
  const k = side.kind;
  switch (k) {
    case "hairline":
      return 2;
    case "thin":
      return 12;
    case "medium":
      return 7;
    case "thick":
      return 11;
    case "double":
      return 13;
    default:
      return 12;
  }
}

/**
 * 从当前活动格样式推断边框选项卡初始状态（预览边线；多格选区时内部线默认关闭）。
 */
export function inferFormatCellsBorderState(style: CellStyle | null): FormatCellsBorderState {
  const bt = style?.borderTop;
  const bb = style?.borderBottom;
  const bl = style?.borderLeft;
  const br = style?.borderRight;

  let lineIndex = 12;
  for (const s of [bt, bb, bl, br]) {
    if (s !== undefined) {
      lineIndex = inferLineStyleIndexFromBorderSide(s);
      break;
    }
  }

  let colorAuto = true;
  let colorArgb: string | null = null;
  for (const s of [bt, bb, bl, br]) {
    if (s?.colorArgb !== undefined && s.colorArgb.trim() !== "") {
      const a = s.colorArgb.trim();
      if (/^[\dA-Fa-f]{8}$/i.test(a)) {
        colorAuto = false;
        colorArgb = a.toUpperCase();
        break;
      }
    }
  }

  return {
    lineStyleIndex: lineIndex,
    colorAuto,
    colorArgb,
    edges: {
      top: bt !== undefined,
      bottom: bb !== undefined,
      left: bl !== undefined,
      right: br !== undefined,
      insideH: false,
      insideV: false,
    },
  };
}

export function stripBorderKeysFromPatch(patch: CellStylePatch): CellStylePatch {
  const {
    borderTop: _t,
    borderLeft: _l,
    borderBottom: _b,
    borderRight: _r,
    ...rest
  } = patch;
  return rest;
}

/** 合并「数字/对齐/字体」补丁与可选的边框重算（单条撤销单元）。 */
export function mergeFormatCellsDialogStyle(
  before: CellStyle | null,
  sheet: Worksheet,
  range: SelectionRange,
  row: number,
  col: number,
  basePatch: CellStylePatch,
  applyBorder: boolean,
  borderState: FormatCellsBorderState,
): CellStyle | null {
  const noBorderPatch = stripBorderKeysFromPatch(basePatch);
  const mid = applyCellStylePatch(before === null ? null : { ...before }, noBorderPatch);
  if (!applyBorder) {
    return mid;
  }
  return computeStyleForFormatCellsBorder(mid, sheet, range, row, col, borderState);
}

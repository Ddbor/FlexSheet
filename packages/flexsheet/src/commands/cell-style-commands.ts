import {
  applyCellStylePatch,
  computeTableFormatCellStyle,
  normalizeSelectionRange,
  TABLE_ACCENT_PALETTES,
  type CellStyle,
  type CellStylePatch,
  type ICommand,
  type ParsedTableStyleCommand,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";
import { cloneCellStyle, computeBorderStyleForRibbonCommand } from "../format-cells/border-ribbon-preset.js";
import { mergeFormatCellsDialogStyle, type FormatCellsBorderState } from "../format-cells/format-cells-border.js";

interface CellStyleSnapshot {
  readonly row: number;
  readonly col: number;
  readonly before: CellStyle | null;
  readonly after: CellStyle | null;
}

const FONT_SIZE_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 26, 28, 36, 48, 72] as const;

function nearestFontStepIndex(pt: number): number {
  if (pt <= FONT_SIZE_STEPS[0]) {
    return 0;
  }
  const last = FONT_SIZE_STEPS[FONT_SIZE_STEPS.length - 1];
  if (pt >= last) {
    return FONT_SIZE_STEPS.length - 1;
  }
  let best = 0;
  let bestD = Infinity;
  for (let i = 0; i < FONT_SIZE_STEPS.length; i++) {
    const d = Math.abs(FONT_SIZE_STEPS[i] - pt);
    if (d < bestD) {
      bestD = d;
      best = i;
    }
  }
  return best;
}

function steppedFontSizePt(fromPt: number, dir: 1 | -1): number {
  const i = nearestFontStepIndex(fromPt);
  const j = dir === 1 ? Math.min(FONT_SIZE_STEPS.length - 1, i + 1) : Math.max(0, i - 1);
  return FONT_SIZE_STEPS[j];
}

const MAX_INDENT_LEVEL = 255;

/** 对规范化选区内每个单元格合并样式补丁（支持撤销/重做）。 */
export class ApplySelectionCellStylePatchCommand implements ICommand {
  readonly id = "cell.applySelectionStylePatch";
  readonly label = "设置单元格样式";
  private readonly snapshots: CellStyleSnapshot[];

  constructor(
    private readonly sheet: Worksheet,
    range: SelectionRange,
    patch: CellStylePatch,
  ) {
    const n = normalizeSelectionRange(range);
    const list: CellStyleSnapshot[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        const before = sheet.getCell(r, c).style;
        const beforeClone = before === null ? null : { ...before };
        const after = applyCellStylePatch(beforeClone, patch);
        list.push({ row: r, col: c, before: beforeClone, after });
      }
    }
    this.snapshots = list;
  }

  execute(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, s.after === null ? null : { ...s.after });
    }
  }

  undo(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, s.before === null ? null : { ...s.before });
    }
  }
}

/** 「设置单元格格式」确定：合并数字/对齐/字体与可选的边框几何（单条撤销）。 */
export class ApplySelectionFormatCellsDialogCommand implements ICommand {
  readonly id = "cell.applyFormatCellsDialog";
  readonly label = "设置单元格格式";
  private readonly snapshots: CellStyleSnapshot[];

  constructor(
    private readonly sheet: Worksheet,
    range: SelectionRange,
    basePatch: CellStylePatch,
    applyBorder: boolean,
    borderState: FormatCellsBorderState,
  ) {
    const n = normalizeSelectionRange(range);
    const list: CellStyleSnapshot[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        if (this.sheet.isMergeCoveredCell(r, c)) {
          continue;
        }
        const raw = this.sheet.getCell(r, c).style;
        const before = cloneCellStyle(raw);
        const after = mergeFormatCellsDialogStyle(raw, this.sheet, range, r, c, basePatch, applyBorder, borderState);
        list.push({ row: r, col: c, before, after });
      }
    }
    this.snapshots = list;
  }

  execute(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, s.after === null ? null : { ...s.after });
    }
  }

  undo(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, cloneCellStyle(s.before));
    }
  }
}

/** 「套用表格格式」：整区写入表样式（单条撤销）。 */
export class ApplyFormatAsTableCommand implements ICommand {
  readonly id = "cell.applyFormatAsTable";
  readonly label = "套用表格格式";
  private readonly snapshots: CellStyleSnapshot[];
  private readonly normalizedRange: SelectionRange;

  constructor(
    private readonly sheet: Worksheet,
    range: SelectionRange,
    private readonly parsed: ParsedTableStyleCommand,
    private readonly hasHeaders: boolean,
  ) {
    const n = normalizeSelectionRange(range);
    this.normalizedRange = n;
    const palette = TABLE_ACCENT_PALETTES[parsed.col];
    const list: CellStyleSnapshot[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        if (this.sheet.isMergeCoveredCell(r, c)) {
          continue;
        }
        const raw = this.sheet.getCell(r, c).style;
        const before = cloneCellStyle(raw);
        const after = computeTableFormatCellStyle(parsed, palette, n, hasHeaders, r, c);
        list.push({ row: r, col: c, before, after });
      }
    }
    this.snapshots = list;
  }

  execute(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, s.after === null ? null : { ...s.after });
    }
    this.sheet.registerTableStyleRegion(this.normalizedRange, this.parsed, this.hasHeaders);
  }

  undo(): void {
    this.sheet.unregisterTableStyleRegion(this.normalizedRange);
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, cloneCellStyle(s.before));
    }
  }
}

const BORDER_RIBBON_IDS = new Set<string>([
  "home.font.border",
  "home.font.border.bottom",
  "home.font.border.top",
  "home.font.border.left",
  "home.font.border.right",
  "home.font.border.none",
  "home.font.border.all",
  "home.font.border.outside",
  "home.font.border.thickBox",
  "home.font.border.doubleBottom",
  "home.font.border.thickBottom",
  "home.font.border.topBottom",
  "home.font.border.topThickBottom",
  "home.font.border.topDoubleBottom",
]);

export function isRibbonBorderCommandId(id: string): boolean {
  return BORDER_RIBBON_IDS.has(id);
}

/** Ribbon 边框下拉 / 主按钮：按命令 id 对选区内每个主格计算样式（可撤销）。 */
export class ApplySelectionBorderRibbonCommand implements ICommand {
  readonly id = "cell.applySelectionBorderRibbon";
  readonly label = "设置边框";
  private readonly snapshots: CellStyleSnapshot[];

  constructor(
    private readonly sheet: Worksheet,
    range: SelectionRange,
    readonly commandId: string,
  ) {
    const n = normalizeSelectionRange(range);
    const list: CellStyleSnapshot[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        if (this.sheet.isMergeCoveredCell(r, c)) {
          continue;
        }
        const before = this.sheet.getCell(r, c).style;
        const beforeClone = cloneCellStyle(before);
        const after = computeBorderStyleForRibbonCommand(
          this.sheet,
          range,
          r,
          c,
          commandId,
          beforeClone,
        );
        list.push({ row: r, col: c, before: beforeClone, after });
      }
    }
    this.snapshots = list;
  }

  execute(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, cloneCellStyle(s.after));
    }
  }

  undo(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, cloneCellStyle(s.before));
    }
  }
}

/** 选区内按 Ribbon 字号阶梯逐格增大或减小字号（每单元格独立参考当前字号，可撤销）。 */
export class ApplySelectionFontSizeStepCommand implements ICommand {
  readonly id = "cell.applySelectionFontSizeStep";
  readonly label = "调整字号";
  private readonly snapshots: CellStyleSnapshot[];

  constructor(
    private readonly sheet: Worksheet,
    range: SelectionRange,
    dir: 1 | -1,
  ) {
    const n = normalizeSelectionRange(range);
    const list: CellStyleSnapshot[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        const before = sheet.getCell(r, c).style;
        const beforeClone = before === null ? null : { ...before };
        const curPt = beforeClone?.fontSizePt ?? 11;
        const nextPt = steppedFontSizePt(curPt, dir);
        const after = applyCellStylePatch(beforeClone, { fontSizePt: nextPt });
        list.push({ row: r, col: c, before: beforeClone, after });
      }
    }
    this.snapshots = list;
  }

  execute(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, s.after === null ? null : { ...s.after });
    }
  }

  undo(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, s.before === null ? null : { ...s.before });
    }
  }
}

/** 选区内逐格增加或减少缩进等级（每单元格独立参考当前等级，可撤销）。 */
export class ApplySelectionIndentStepCommand implements ICommand {
  readonly id = "cell.applySelectionIndentStep";
  readonly label = "调整缩进";
  private readonly snapshots: CellStyleSnapshot[];

  constructor(
    private readonly sheet: Worksheet,
    range: SelectionRange,
    dir: 1 | -1,
  ) {
    const n = normalizeSelectionRange(range);
    const list: CellStyleSnapshot[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        const before = sheet.getCell(r, c).style;
        const beforeClone = before === null ? null : { ...before };
        const cur = beforeClone?.indentLevel ?? 0;
        const nextLv = Math.max(0, Math.min(MAX_INDENT_LEVEL, cur + dir));
        const patch: CellStylePatch =
          nextLv === 0
            ? { indentLevel: null }
            : { indentLevel: nextLv, hAlign: null };
        const after = applyCellStylePatch(beforeClone, patch);
        list.push({ row: r, col: c, before: beforeClone, after });
      }
    }
    this.snapshots = list;
  }

  execute(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, s.after === null ? null : { ...s.after });
    }
  }

  undo(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, s.before === null ? null : { ...s.before });
    }
  }
}

/** 清除选区内所有单元格格式（保留值/公式，可撤销）。 */
export class ClearSelectionFormatsCommand implements ICommand {
  readonly id = "cell.clearSelectionFormats";
  readonly label = "清除格式";
  readonly hasChanges: boolean;
  private readonly snapshots: CellStyleSnapshot[];

  constructor(
    private readonly sheet: Worksheet,
    range: SelectionRange,
  ) {
    const n = normalizeSelectionRange(range);
    const list: CellStyleSnapshot[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        const before = cloneCellStyle(this.sheet.getCell(r, c).style);
        list.push({ row: r, col: c, before, after: null });
      }
    }
    this.snapshots = list;
    this.hasChanges = list.some((s) => s.before !== null);
  }

  execute(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, null);
    }
  }

  undo(): void {
    for (const s of this.snapshots) {
      this.sheet.setCellStyle(s.row, s.col, cloneCellStyle(s.before));
    }
  }
}

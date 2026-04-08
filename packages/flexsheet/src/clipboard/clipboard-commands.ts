import type { CellScalar, CellStyle, ICommand, SelectionRange, Worksheet } from "@flexsheet/core";
import { normalizeSelectionRange } from "@flexsheet/core";
import { parseEditString } from "@flexsheet/editor";
import { recalcWorksheet } from "@flexsheet/formula";

interface CellFullSnapshot {
  readonly row: number;
  readonly col: number;
  readonly formula: string | null;
  readonly value: CellScalar;
  readonly style: CellStyle | null;
}

function takeSnapshot(sheet: Worksheet, row: number, col: number): CellFullSnapshot {
  const cell = sheet.getCell(row, col);
  return {
    row,
    col,
    formula: cell.formula,
    value: cell.value,
    style: cell.style === null ? null : { ...cell.style },
  };
}

function applySnapshot(sheet: Worksheet, s: CellFullSnapshot): void {
  if (s.formula !== null && s.formula.length > 0) {
    sheet.setCellFormula(s.row, s.col, s.formula);
  } else {
    sheet.setCellLiteral(s.row, s.col, s.value);
  }
  sheet.setCellStyle(s.row, s.col, s.style === null ? null : { ...s.style });
}

function normalizePasteRows(values: readonly (readonly string[])[]): string[][] {
  if (values.length === 0) {
    return [];
  }
  const colCount = Math.max(...values.map((r) => r.length));
  return values.map((row) => {
    const next = [...row];
    while (next.length < colCount) {
      next.push("");
    }
    return next;
  });
}

/** 与 `PasteRegionCommand` 一致：按表边界裁剪后的粘贴占用行/列数。 */
export function getPasteClippedRect(
  sheet: Worksheet,
  startRow: number,
  startCol: number,
  values: readonly (readonly string[])[],
): { h: number; w: number } {
  const grid = normalizePasteRows(values);
  const pasteRows = grid.length;
  const pasteCols = pasteRows > 0 ? grid[0]!.length : 0;
  const h = Math.min(pasteRows, Math.max(0, sheet.rowCount - startRow));
  const w = Math.min(pasteCols, Math.max(0, sheet.colCount - startCol));
  return { h, w };
}

/**
 * 从活动格左上角粘贴矩形；execute 前根据表边界裁剪行列。
 */
export class PasteRegionCommand implements ICommand {
  readonly id = "clipboard.paste";
  readonly label = "粘贴";

  private readonly sheet: Worksheet;
  private readonly startRow: number;
  private readonly startCol: number;
  private readonly values: string[][];
  private readonly styles: readonly (readonly (CellStyle | null)[])[] | null;
  private readonly h: number;
  private readonly w: number;
  private readonly before: CellFullSnapshot[];

  constructor(
    sheet: Worksheet,
    startRow: number,
    startCol: number,
    values: readonly (readonly string[])[],
    styles: readonly (readonly (CellStyle | null)[])[] | null,
  ) {
    this.sheet = sheet;
    this.startRow = startRow;
    this.startCol = startCol;
    this.values = normalizePasteRows(values);
    const dim = getPasteClippedRect(sheet, startRow, startCol, values);
    this.h = dim.h;
    this.w = dim.w;

    let styleGrid: readonly (readonly (CellStyle | null)[])[] | null = styles;
    if (styleGrid !== null) {
      if (
        styleGrid.length !== this.values.length ||
        styleGrid.some((r, ri) => r.length !== this.values[ri]!.length)
      ) {
        styleGrid = null;
      }
    }
    this.styles = styleGrid;

    this.before = [];
    for (let i = 0; i < this.h; i++) {
      for (let j = 0; j < this.w; j++) {
        this.before.push(takeSnapshot(sheet, startRow + i, startCol + j));
      }
    }
  }

  execute(): void {
    const { sheet, startRow, startCol, values, styles } = this;
    sheet.batch(() => {
      for (let i = 0; i < this.h; i++) {
        for (let j = 0; j < this.w; j++) {
          const r = startRow + i;
          const c = startCol + j;
          const raw = values[i]![j] ?? "";
          sheet.setCellValue(r, c, parseEditString(raw));
          if (styles !== null) {
            const st = styles[i]![j] ?? null;
            sheet.setCellStyle(r, c, st === null ? null : { ...st });
          }
        }
      }
    });
    recalcWorksheet(sheet);
  }

  undo(): void {
    let k = 0;
    this.sheet.batch(() => {
      for (let i = 0; i < this.h; i++) {
        for (let j = 0; j < this.w; j++) {
          applySnapshot(this.sheet, this.before[k]!);
          k++;
        }
      }
    });
    recalcWorksheet(this.sheet);
  }
}

/**
 * 延迟剪切完成后：清空「剪切源」矩形内、且落在粘贴占用矩形之外的单元格
 * （避免与 `PasteRegionCommand` 重叠区域被误清空）。
 */
export class CutRangeExceptRectCommand implements ICommand {
  readonly id = "clipboard.cutDeferredClear";
  readonly label = "剪切后清空源区";

  private readonly sheet: Worksheet;
  private readonly before: CellFullSnapshot[];

  constructor(
    sheet: Worksheet,
    cutRange: SelectionRange,
    pasteTop: number,
    pasteLeft: number,
    pasteH: number,
    pasteW: number,
  ) {
    this.sheet = sheet;
    const n = normalizeSelectionRange(cutRange);
    const pr1 = pasteTop + pasteH - 1;
    const pc1 = pasteLeft + pasteW - 1;
    const snaps: CellFullSnapshot[] = [];
    const hasPaste = pasteH > 0 && pasteW > 0;
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        const insidePaste =
          hasPaste && r >= pasteTop && r <= pr1 && c >= pasteLeft && c <= pc1;
        if (!insidePaste) {
          snaps.push(takeSnapshot(sheet, r, c));
        }
      }
    }
    this.before = snaps;
  }

  execute(): void {
    if (this.before.length === 0) {
      return;
    }
    const sheet = this.sheet;
    sheet.batch(() => {
      for (const s of this.before) {
        sheet.setCellValue(s.row, s.col, null);
        sheet.setCellStyle(s.row, s.col, null);
      }
    });
    recalcWorksheet(sheet);
  }

  undo(): void {
    this.sheet.batch(() => {
      for (const s of this.before) {
        applySnapshot(this.sheet, s);
      }
    });
    recalcWorksheet(this.sheet);
  }
}

/** 剪切：立即清空选区内单元格的值与样式（仍可用于右键等需即时清空的场景）。 */
export class CutClearRegionCommand implements ICommand {
  readonly id = "clipboard.cut";
  readonly label = "剪切";

  private readonly sheet: Worksheet;
  private readonly before: CellFullSnapshot[];

  constructor(sheet: Worksheet, range: SelectionRange) {
    this.sheet = sheet;
    const n = normalizeSelectionRange(range);
    const snaps: CellFullSnapshot[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        snaps.push(takeSnapshot(sheet, r, c));
      }
    }
    this.before = snaps;
  }

  execute(): void {
    const sheet = this.sheet;
    const n = this.before;
    if (n.length === 0) {
      return;
    }
    sheet.batch(() => {
      for (const s of n) {
        sheet.setCellValue(s.row, s.col, null);
        sheet.setCellStyle(s.row, s.col, null);
      }
    });
    recalcWorksheet(sheet);
  }

  undo(): void {
    this.sheet.batch(() => {
      for (const s of this.before) {
        applySnapshot(this.sheet, s);
      }
    });
    recalcWorksheet(this.sheet);
  }
}

import {
  normalizeSelectionRange,
  type CellStyle,
  type ICommand,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";
import type { SelectionModel } from "@flexsheet/selection";

interface CellSnapshot {
  readonly row: number;
  readonly col: number;
  readonly formula: string | null;
  readonly value: string | number | boolean | null;
}

interface RowMetaSnapshot {
  readonly row: number;
  readonly height: number;
  readonly hidden: boolean;
}

interface ColMetaSnapshot {
  readonly col: number;
  readonly width: number;
  readonly hidden: boolean;
}

interface CellStateSnapshot {
  readonly formula: string | null;
  readonly value: string | number | boolean | null;
  readonly style: CellStyle | null;
}

export class InsertRowsCommand implements ICommand {
  readonly id = "sheet.insertRows";
  readonly label = "插入行";
  constructor(
    private readonly sheet: Worksheet,
    private readonly selection: SelectionModel,
    private readonly atRow: number,
    private readonly count: number,
  ) {}
  execute(): void {
    this.sheet.insertRows(this.atRow, this.count);
    shiftAllFormulasForRows(this.sheet, "insert", this.atRow, this.count);
    shiftSelectionRows(this.selection, this.atRow, this.count);
    recalcWorksheet(this.sheet);
  }
  undo(): void {
    this.sheet.deleteRows(this.atRow, this.count);
    shiftAllFormulasForRows(this.sheet, "delete", this.atRow, this.count);
    shiftSelectionRows(this.selection, this.atRow, -this.count);
    recalcWorksheet(this.sheet);
  }
}

export class DeleteRowsCommand implements ICommand {
  readonly id = "sheet.deleteRows";
  readonly label = "删除行";
  private readonly removedCells: CellSnapshot[] = [];
  private readonly removedMeta: RowMetaSnapshot[] = [];
  constructor(
    private readonly sheet: Worksheet,
    private readonly selection: SelectionModel,
    private readonly atRow: number,
    private readonly count: number,
  ) {
    const end = Math.min(sheet.rowCount - 1, atRow + count - 1);
    for (let r = atRow; r <= end; r++) {
      this.removedMeta.push({
        row: r,
        height: sheet.getRowHeight(r),
        hidden: sheet.isRowHidden(r),
      });
      for (let c = 0; c < sheet.colCount; c++) {
        const cell = sheet.getCell(r, c);
        this.removedCells.push({ row: r, col: c, formula: cell.formula, value: cell.value });
      }
    }
  }
  execute(): void {
    this.sheet.deleteRows(this.atRow, this.count);
    shiftAllFormulasForRows(this.sheet, "delete", this.atRow, this.count);
    shiftSelectionRows(this.selection, this.atRow, -this.count);
    recalcWorksheet(this.sheet);
  }
  undo(): void {
    this.sheet.insertRows(this.atRow, this.count);
    shiftAllFormulasForRows(this.sheet, "insert", this.atRow, this.count);
    this.sheet.batch(() => {
      for (const m of this.removedMeta) {
        this.sheet.setRowHeight(m.row, m.height);
        this.sheet.setRowHidden(m.row, m.hidden);
      }
      for (const cell of this.removedCells) {
        if (cell.formula !== null) {
          this.sheet.setCellFormula(cell.row, cell.col, cell.formula);
        } else {
          this.sheet.setCellLiteral(cell.row, cell.col, cell.value);
        }
      }
    });
    shiftSelectionRows(this.selection, this.atRow, this.count);
    recalcWorksheet(this.sheet);
  }
}

export class InsertColsCommand implements ICommand {
  readonly id = "sheet.insertCols";
  readonly label = "插入列";
  constructor(
    private readonly sheet: Worksheet,
    private readonly selection: SelectionModel,
    private readonly atCol: number,
    private readonly count: number,
  ) {}
  execute(): void {
    this.sheet.insertCols(this.atCol, this.count);
    shiftAllFormulasForCols(this.sheet, "insert", this.atCol, this.count);
    shiftSelectionCols(this.selection, this.atCol, this.count);
    recalcWorksheet(this.sheet);
  }
  undo(): void {
    this.sheet.deleteCols(this.atCol, this.count);
    shiftAllFormulasForCols(this.sheet, "delete", this.atCol, this.count);
    shiftSelectionCols(this.selection, this.atCol, -this.count);
    recalcWorksheet(this.sheet);
  }
}

export class DeleteColsCommand implements ICommand {
  readonly id = "sheet.deleteCols";
  readonly label = "删除列";
  private readonly removedCells: CellSnapshot[] = [];
  private readonly removedMeta: ColMetaSnapshot[] = [];
  constructor(
    private readonly sheet: Worksheet,
    private readonly selection: SelectionModel,
    private readonly atCol: number,
    private readonly count: number,
  ) {
    const end = Math.min(sheet.colCount - 1, atCol + count - 1);
    for (let c = atCol; c <= end; c++) {
      this.removedMeta.push({
        col: c,
        width: sheet.getColWidth(c),
        hidden: sheet.isColHidden(c),
      });
      for (let r = 0; r < sheet.rowCount; r++) {
        const cell = sheet.getCell(r, c);
        this.removedCells.push({ row: r, col: c, formula: cell.formula, value: cell.value });
      }
    }
  }
  execute(): void {
    this.sheet.deleteCols(this.atCol, this.count);
    shiftAllFormulasForCols(this.sheet, "delete", this.atCol, this.count);
    shiftSelectionCols(this.selection, this.atCol, -this.count);
    recalcWorksheet(this.sheet);
  }
  undo(): void {
    this.sheet.insertCols(this.atCol, this.count);
    shiftAllFormulasForCols(this.sheet, "insert", this.atCol, this.count);
    this.sheet.batch(() => {
      for (const m of this.removedMeta) {
        this.sheet.setColWidth(m.col, m.width);
        this.sheet.setColHidden(m.col, m.hidden);
      }
      for (const cell of this.removedCells) {
        if (cell.formula !== null) {
          this.sheet.setCellFormula(cell.row, cell.col, cell.formula);
        } else {
          this.sheet.setCellLiteral(cell.row, cell.col, cell.value);
        }
      }
    });
    shiftSelectionCols(this.selection, this.atCol, this.count);
    recalcWorksheet(this.sheet);
  }
}

export class SetRowHiddenCommand implements ICommand {
  readonly id = "sheet.setRowHidden";
  readonly label: string;
  private readonly before: boolean;
  constructor(private readonly sheet: Worksheet, private readonly row: number, private readonly hidden: boolean) {
    this.before = sheet.isRowHidden(row);
    this.label = hidden ? "隐藏行" : "取消隐藏行";
  }
  execute(): void {
    this.sheet.setRowHidden(this.row, this.hidden);
  }
  undo(): void {
    this.sheet.setRowHidden(this.row, this.before);
  }
}

export class SetColHiddenCommand implements ICommand {
  readonly id = "sheet.setColHidden";
  readonly label: string;
  private readonly before: boolean;
  constructor(private readonly sheet: Worksheet, private readonly col: number, private readonly hidden: boolean) {
    this.before = sheet.isColHidden(col);
    this.label = hidden ? "隐藏列" : "取消隐藏列";
  }
  execute(): void {
    this.sheet.setColHidden(this.col, this.hidden);
  }
  undo(): void {
    this.sheet.setColHidden(this.col, this.before);
  }
}

export class SetRowHeightCommand implements ICommand {
  readonly id = "sheet.setRowHeight";
  readonly label = "调整行高";
  private readonly before: number;
  constructor(private readonly sheet: Worksheet, private readonly row: number, private readonly height: number) {
    this.before = sheet.getRowHeight(row);
  }
  execute(): void {
    this.sheet.setRowHeight(this.row, this.height);
  }
  undo(): void {
    this.sheet.setRowHeight(this.row, this.before);
  }
}

export class SetColWidthCommand implements ICommand {
  readonly id = "sheet.setColWidth";
  readonly label = "调整列宽";
  private readonly before: number;
  constructor(private readonly sheet: Worksheet, private readonly col: number, private readonly width: number) {
    this.before = sheet.getColWidth(col);
  }
  execute(): void {
    this.sheet.setColWidth(this.col, this.width);
  }
  undo(): void {
    this.sheet.setColWidth(this.col, this.before);
  }
}

/** 将闭区间 [startRow,endRow] 各行设为同一高度，单次撤销。 */
export class SetRowHeightsInRangeCommand implements ICommand {
  readonly id = "sheet.setRowHeightsInRange";
  readonly label = "调整行高";
  private readonly sheet: Worksheet;
  private readonly startRow: number;
  private readonly endRow: number;
  private readonly height: number;
  private readonly before: number[];

  constructor(sheet: Worksheet, startRow: number, endRow: number, height: number) {
    this.sheet = sheet;
    const lo = Math.min(startRow, endRow);
    const hi = Math.max(startRow, endRow);
    this.startRow = lo;
    this.endRow = hi;
    this.height = height;
    this.before = [];
    for (let r = lo; r <= hi; r++) {
      this.before.push(sheet.getRowHeight(r));
    }
  }

  execute(): void {
    for (let r = this.startRow; r <= this.endRow; r++) {
      this.sheet.setRowHeight(r, this.height);
    }
  }

  undo(): void {
    let i = 0;
    for (let r = this.startRow; r <= this.endRow; r++) {
      this.sheet.setRowHeight(r, this.before[i]!);
      i++;
    }
  }
}

/** 将闭区间 [startCol,endCol] 各列设为同一宽度，单次撤销。 */
export class SetColWidthsInRangeCommand implements ICommand {
  readonly id = "sheet.setColWidthsInRange";
  readonly label = "调整列宽";
  private readonly sheet: Worksheet;
  private readonly startCol: number;
  private readonly endCol: number;
  private readonly width: number;
  private readonly before: number[];

  constructor(sheet: Worksheet, startCol: number, endCol: number, width: number) {
    this.sheet = sheet;
    const lo = Math.min(startCol, endCol);
    const hi = Math.max(startCol, endCol);
    this.startCol = lo;
    this.endCol = hi;
    this.width = width;
    this.before = [];
    for (let c = lo; c <= hi; c++) {
      this.before.push(sheet.getColWidth(c));
    }
  }

  execute(): void {
    for (let c = this.startCol; c <= this.endCol; c++) {
      this.sheet.setColWidth(c, this.width);
    }
  }

  undo(): void {
    let i = 0;
    for (let c = this.startCol; c <= this.endCol; c++) {
      this.sheet.setColWidth(c, this.before[i]!);
      i++;
    }
  }
}

export class InsertCellsShiftRightCommand implements ICommand {
  readonly id = "sheet.insertCellsShiftRight";
  readonly label = "插入单元格右移";
  private readonly beforeRow: CellStateSnapshot[];
  private readonly beforeSelection: SelectionRange;

  constructor(
    private readonly sheet: Worksheet,
    private readonly selection: SelectionModel,
    private readonly row: number,
    private readonly col: number,
    private readonly count: number,
  ) {
    this.beforeRow = captureRow(this.sheet, this.row);
    this.beforeSelection = this.selection.getNormalizedRange();
  }

  execute(): void {
    const n = clampCount(this.count);
    this.sheet.batch(() => {
      for (let c = this.sheet.colCount - 1; c >= this.col; c--) {
        const target = this.sheet.getCell(this.row, c);
        const sourceCol = c - n;
        if (sourceCol >= this.col) {
          const source = this.sheet.getCell(this.row, sourceCol);
          target.formula = source.formula;
          target.value = source.value;
          target.style = source.style;
        } else {
          target.formula = null;
          target.value = null;
          target.style = null;
        }
      }
      this.sheet.notifyDataChanged();
    });
    this.selection.selectCell(this.row, this.col);
    recalcWorksheet(this.sheet);
  }

  undo(): void {
    restoreRow(this.sheet, this.row, this.beforeRow);
    this.selection.setNormalizedRange(this.beforeSelection);
    recalcWorksheet(this.sheet);
  }
}

export class InsertCellsShiftDownCommand implements ICommand {
  readonly id = "sheet.insertCellsShiftDown";
  readonly label = "插入单元格下移";
  private readonly beforeCol: CellStateSnapshot[];
  private readonly beforeSelection: SelectionRange;

  constructor(
    private readonly sheet: Worksheet,
    private readonly selection: SelectionModel,
    private readonly row: number,
    private readonly col: number,
    private readonly count: number,
  ) {
    this.beforeCol = captureCol(this.sheet, this.col);
    this.beforeSelection = this.selection.getNormalizedRange();
  }

  execute(): void {
    const n = clampCount(this.count);
    this.sheet.batch(() => {
      for (let r = this.sheet.rowCount - 1; r >= this.row; r--) {
        const target = this.sheet.getCell(r, this.col);
        const sourceRow = r - n;
        if (sourceRow >= this.row) {
          const source = this.sheet.getCell(sourceRow, this.col);
          target.formula = source.formula;
          target.value = source.value;
          target.style = source.style;
        } else {
          target.formula = null;
          target.value = null;
          target.style = null;
        }
      }
      this.sheet.notifyDataChanged();
    });
    this.selection.selectCell(this.row, this.col);
    recalcWorksheet(this.sheet);
  }

  undo(): void {
    restoreCol(this.sheet, this.col, this.beforeCol);
    this.selection.setNormalizedRange(this.beforeSelection);
    recalcWorksheet(this.sheet);
  }
}

/** 删除选区单元格，同行内右侧左移填补（与 Excel「右侧单元格左移」一致）。 */
export class DeleteCellsShiftLeftCommand implements ICommand {
  readonly id = "sheet.deleteCellsShiftLeft";
  readonly label = "删除单元格右移左填";
  private readonly range: SelectionRange;
  private readonly beforeTailByRow: readonly CellStateSnapshot[][];
  private readonly beforeSelection: SelectionRange;

  constructor(
    private readonly sheet: Worksheet,
    private readonly selection: SelectionModel,
    range: SelectionRange,
  ) {
    const n = normalizeSelectionRange(range);
    this.range = n;
    this.beforeSelection = this.selection.getNormalizedRange();
    const tails: CellStateSnapshot[][] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      const seg: CellStateSnapshot[] = [];
      for (let c = n.startCol; c < sheet.colCount; c++) {
        const cell = sheet.getCell(r, c);
        seg.push({
          formula: cell.formula,
          value: cell.value,
          style: cell.style,
        });
      }
      tails.push(seg);
    }
    this.beforeTailByRow = tails;
  }

  execute(): void {
    const sheet = this.sheet;
    const { startRow, endRow, startCol, endCol } = this.range;
    const w = endCol - startCol + 1;
    sheet.batch(() => {
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= sheet.colCount - 1 - w; c++) {
          const target = sheet.getCell(r, c);
          const source = sheet.getCell(r, c + w);
          target.formula = source.formula;
          target.value = source.value;
          target.style = source.style;
        }
        for (let c = sheet.colCount - w; c < sheet.colCount; c++) {
          const target = sheet.getCell(r, c);
          target.formula = null;
          target.value = null;
          target.style = null;
        }
      }
      sheet.notifyDataChanged();
    });
    this.selection.selectCell(startRow, startCol);
    recalcWorksheet(sheet);
  }

  undo(): void {
    const sheet = this.sheet;
    const { startRow, endRow, startCol } = this.range;
    let rowIdx = 0;
    for (let r = startRow; r <= endRow; r++) {
      const seg = this.beforeTailByRow[rowIdx]!;
      for (let j = 0; j < seg.length; j++) {
        const c = startCol + j;
        if (c >= sheet.colCount) {
          break;
        }
        const target = sheet.getCell(r, c);
        const st = seg[j]!;
        target.formula = st.formula;
        target.value = st.value;
        target.style = st.style;
      }
      rowIdx++;
    }
    sheet.notifyDataChanged();
    this.selection.setNormalizedRange(this.beforeSelection);
    recalcWorksheet(sheet);
  }
}

/** 删除选区单元格，同列内下方上移填补（与 Excel「下方单元格上移」一致）。 */
export class DeleteCellsShiftUpCommand implements ICommand {
  readonly id = "sheet.deleteCellsShiftUp";
  readonly label = "删除单元格下移上填";
  private readonly range: SelectionRange;
  private readonly beforeTailByCol: readonly CellStateSnapshot[][];
  private readonly beforeSelection: SelectionRange;

  constructor(
    private readonly sheet: Worksheet,
    private readonly selection: SelectionModel,
    range: SelectionRange,
  ) {
    const n = normalizeSelectionRange(range);
    this.range = n;
    this.beforeSelection = this.selection.getNormalizedRange();
    const tails: CellStateSnapshot[][] = [];
    for (let c = n.startCol; c <= n.endCol; c++) {
      const seg: CellStateSnapshot[] = [];
      for (let r = n.startRow; r < sheet.rowCount; r++) {
        const cell = sheet.getCell(r, c);
        seg.push({
          formula: cell.formula,
          value: cell.value,
          style: cell.style,
        });
      }
      tails.push(seg);
    }
    this.beforeTailByCol = tails;
  }

  execute(): void {
    const sheet = this.sheet;
    const { startRow, endRow, startCol, endCol } = this.range;
    const h = endRow - startRow + 1;
    sheet.batch(() => {
      for (let c = startCol; c <= endCol; c++) {
        for (let r = startRow; r <= sheet.rowCount - 1 - h; r++) {
          const target = sheet.getCell(r, c);
          const source = sheet.getCell(r + h, c);
          target.formula = source.formula;
          target.value = source.value;
          target.style = source.style;
        }
        for (let r = sheet.rowCount - h; r < sheet.rowCount; r++) {
          const target = sheet.getCell(r, c);
          target.formula = null;
          target.value = null;
          target.style = null;
        }
      }
      sheet.notifyDataChanged();
    });
    this.selection.selectCell(startRow, startCol);
    recalcWorksheet(sheet);
  }

  undo(): void {
    const sheet = this.sheet;
    const { startRow, startCol, endCol } = this.range;
    let colIdx = 0;
    for (let c = startCol; c <= endCol; c++) {
      const seg = this.beforeTailByCol[colIdx]!;
      for (let j = 0; j < seg.length; j++) {
        const r = startRow + j;
        if (r >= sheet.rowCount) {
          break;
        }
        const target = sheet.getCell(r, c);
        const st = seg[j]!;
        target.formula = st.formula;
        target.value = st.value;
        target.style = st.style;
      }
      colIdx++;
    }
    sheet.notifyDataChanged();
    this.selection.setNormalizedRange(this.beforeSelection);
    recalcWorksheet(sheet);
  }
}

function shiftSelectionRows(selection: SelectionModel, atRow: number, delta: number): void {
  const r = selection.getNormalizedRange();
  const shifted = {
    startRow: shiftIndex(r.startRow, atRow, delta),
    endRow: shiftIndex(r.endRow, atRow, delta),
    startCol: r.startCol,
    endCol: r.endCol,
  };
  selection.setNormalizedRange(shifted);
}

function shiftSelectionCols(selection: SelectionModel, atCol: number, delta: number): void {
  const r = selection.getNormalizedRange();
  const shifted = {
    startRow: r.startRow,
    endRow: r.endRow,
    startCol: shiftIndex(r.startCol, atCol, delta),
    endCol: shiftIndex(r.endCol, atCol, delta),
  };
  selection.setNormalizedRange(shifted);
}

function shiftIndex(value: number, at: number, delta: number): number {
  if (delta > 0) {
    return value >= at ? value + delta : value;
  }
  const remove = Math.abs(delta);
  if (value < at) {
    return value;
  }
  if (value >= at + remove) {
    return value - remove;
  }
  return at;
}

type ShiftKind = "insert" | "delete";

function shiftAllFormulasForRows(sheet: Worksheet, kind: ShiftKind, atRow: number, count: number): void {
  if (count <= 0) return;
  sheet.batch(() => {
    sheet.iterateCells((cell) => {
      if (cell.formula === null) return;
      cell.formula = shiftFormulaRows(cell.formula, kind, atRow, count);
    });
  });
}

function shiftAllFormulasForCols(sheet: Worksheet, kind: ShiftKind, atCol: number, count: number): void {
  if (count <= 0) return;
  sheet.batch(() => {
    sheet.iterateCells((cell) => {
      if (cell.formula === null) return;
      cell.formula = shiftFormulaCols(cell.formula, kind, atCol, count);
    });
  });
}

function shiftFormulaRows(formula: string, kind: ShiftKind, atRow: number, count: number): string {
  return formula.replace(/\b([A-Za-z]+)(\d+)\b/g, (_m, letters: string, digits: string) => {
    const row = Number(digits) - 1;
    if (!Number.isFinite(row) || row < 0) return `${letters}${digits}`;
    const shifted = shiftIndexByKind(row, kind, atRow, count);
    return `${letters}${shifted + 1}`;
  });
}

function shiftFormulaCols(formula: string, kind: ShiftKind, atCol: number, count: number): string {
  return formula.replace(/\b([A-Za-z]+)(\d+)\b/g, (_m, letters: string, digits: string) => {
    const col = colLabelToIndex(letters);
    const shifted = shiftIndexByKind(col, kind, atCol, count);
    return `${indexToColLabel(shifted)}${digits}`;
  });
}

function shiftIndexByKind(value: number, kind: ShiftKind, at: number, count: number): number {
  if (kind === "insert") {
    return value >= at ? value + count : value;
  }
  if (value < at) {
    return value;
  }
  if (value >= at + count) {
    return value - count;
  }
  return at;
}

function colLabelToIndex(label: string): number {
  let v = 0;
  const s = label.toUpperCase();
  for (let i = 0; i < s.length; i++) {
    v = v * 26 + (s.charCodeAt(i) - 64);
  }
  return v - 1;
}

function indexToColLabel(col: number): string {
  let x = Math.max(0, col);
  let out = "";
  while (x >= 0) {
    out = String.fromCharCode((x % 26) + 65) + out;
    x = Math.floor(x / 26) - 1;
  }
  return out;
}

function clampCount(count: number): number {
  return Math.max(1, Math.trunc(count));
}

function captureRow(sheet: Worksheet, row: number): CellStateSnapshot[] {
  const out: CellStateSnapshot[] = [];
  for (let c = 0; c < sheet.colCount; c++) {
    const cell = sheet.getCell(row, c);
    out.push({
      formula: cell.formula,
      value: cell.value,
      style: cell.style,
    });
  }
  return out;
}

function captureCol(sheet: Worksheet, col: number): CellStateSnapshot[] {
  const out: CellStateSnapshot[] = [];
  for (let r = 0; r < sheet.rowCount; r++) {
    const cell = sheet.getCell(r, col);
    out.push({
      formula: cell.formula,
      value: cell.value,
      style: cell.style,
    });
  }
  return out;
}

function restoreRow(sheet: Worksheet, row: number, states: readonly CellStateSnapshot[]): void {
  sheet.batch(() => {
    const limit = Math.min(states.length, sheet.colCount);
    for (let c = 0; c < limit; c++) {
      const target = sheet.getCell(row, c);
      const state = states[c];
      if (state === undefined) {
        continue;
      }
      target.formula = state.formula;
      target.value = state.value;
      target.style = state.style;
    }
    sheet.notifyDataChanged();
  });
}

function restoreCol(sheet: Worksheet, col: number, states: readonly CellStateSnapshot[]): void {
  sheet.batch(() => {
    const limit = Math.min(states.length, sheet.rowCount);
    for (let r = 0; r < limit; r++) {
      const target = sheet.getCell(r, col);
      const state = states[r];
      if (state === undefined) {
        continue;
      }
      target.formula = state.formula;
      target.value = state.value;
      target.style = state.style;
    }
    sheet.notifyDataChanged();
  });
}

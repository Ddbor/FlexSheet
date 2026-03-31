import { Cell, type CellScalar, type CellStyle } from "./cell.js";

/** 工作表数据或网格规模变更时触发；由 `Workbook` 汇聚后通知宿主刷新视图。 */
export type WorksheetChangeListener = () => void;

/**
 * 工作表：稀疏存储单元格，行列规模可扩展。
 * 通过 `subscribe` / `notifyDataChanged` 与 `batch` 支持数据驱动渲染。
 */
export class Worksheet {
  readonly name: string;
  private readonly cells = new Map<string, Cell>();

  /** 逻辑行数（≥1）。规模变更请用 `setGridSize`。 */
  rowCount: number;
  /** 逻辑列数（≥1）。规模变更请用 `setGridSize`。 */
  colCount: number;
  defaultRowHeight = 24;
  defaultColWidth = 80;

  private readonly changeListeners = new Set<WorksheetChangeListener>();
  private batchDepth = 0;
  private pendingNotify = false;
  private _revision = 0;

  constructor(name: string, rowCount = 1000, colCount = 26) {
    this.name = name;
    this.rowCount = Math.max(1, rowCount);
    this.colCount = Math.max(1, colCount);
  }

  /** 单调递增，每次成功向监听者发出变更通知后 +1（可用于外部缓存失效）。 */
  get revision(): number {
    return this._revision;
  }

  subscribe(listener: WorksheetChangeListener): () => void {
    this.changeListeners.add(listener);
    return () => {
      this.changeListeners.delete(listener);
    };
  }

  /**
   * 在批量操作内合并多次数据修改为一次通知（嵌套时仅最外层结束后再派发）。
   */
  batch<T>(fn: () => T): T {
    this.batchDepth++;
    try {
      return fn();
    } finally {
      this.batchDepth--;
      if (this.batchDepth === 0 && this.pendingNotify) {
        this.pendingNotify = false;
        this.flushNotify();
      }
    }
  }

  /**
   * 在直接修改 `Cell` 字段的大批量路径（如导入）结束后调用一次，以驱动视图更新。
   */
  notifyDataChanged(): void {
    if (this.batchDepth > 0) {
      this.pendingNotify = true;
      return;
    }
    this.flushNotify();
  }

  private flushNotify(): void {
    this._revision++;
    for (const fn of this.changeListeners) {
      fn();
    }
  }

  private touchData(): void {
    if (this.batchDepth > 0) {
      this.pendingNotify = true;
      return;
    }
    this.flushNotify();
  }

  /** 更新逻辑网格大小并通知监听者。 */
  setGridSize(rowCount: number, colCount: number): void {
    const r = Math.max(1, Math.trunc(rowCount));
    const c = Math.max(1, Math.trunc(colCount));
    if (r === this.rowCount && c === this.colCount) {
      return;
    }
    this.rowCount = r;
    this.colCount = c;
    this.touchData();
  }

  getCell(row: number, col: number): Cell {
    const key = Cell.key(row, col);
    let cell = this.cells.get(key);
    if (cell === undefined) {
      cell = new Cell(row, col, null);
      this.cells.set(key, cell);
    }
    return cell;
  }

  /** 遍历已创建单元格（稀疏）。 */
  iterateCells(callback: (cell: Cell) => void): void {
    for (const c of this.cells.values()) {
      callback(c);
    }
  }

  /** 字面量：清除公式并写入标量（不含公式重算，请在外层调用 `recalcWorksheet`）。 */
  setCellLiteral(row: number, col: number, value: CellScalar): void {
    const cell = this.getCell(row, col);
    cell.formula = null;
    cell.value = value;
    this.touchData();
  }

  /** 公式串须含前导 `=`（不含重算，请在外层调用 `recalcWorksheet`）。 */
  setCellFormula(row: number, col: number, formula: string): void {
    const cell = this.getCell(row, col);
    cell.formula = formula.trim();
    this.touchData();
  }

  /**
   * 字符串若以 `=` 开头则视为公式，否则为字面量。
   * 修改公式依赖后请在外层调用 `recalcWorksheet`（或使用 `setCellValueAndRecalc`）。
   */
  setCellValue(row: number, col: number, value: CellScalar): void {
    if (typeof value === "string") {
      const t = value.trim();
      if (t.startsWith("=")) {
        this.setCellFormula(row, col, t);
        return;
      }
    }
    this.setCellLiteral(row, col, value);
  }

  /** 写入单元格样式（`null` 表示清除样式）。 */
  setCellStyle(row: number, col: number, style: CellStyle | null): void {
    const cell = this.getCell(row, col);
    cell.style = style;
    this.touchData();
  }

  hasCell(row: number, col: number): boolean {
    return this.cells.has(Cell.key(row, col));
  }
}

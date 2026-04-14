import {
  AUTO_FILTER_BLANK_KEY,
  cellToAutoFilterDisplayKey,
  compareAutoFilterDisplayKeys,
} from "./column-auto-filter-keys.js";
import { Cell, applyCellStylePatch, type CellScalar, type CellStyle } from "./cell.js";
import type {
  ConditionalFormatRule,
  ConditionalFormattingOverlay,
} from "./conditional-formatting.js";
import { resolveConditionalFormattingOverlay } from "./conditional-formatting.js";
import { formatCellDisplayWithStyle } from "./excel-number-format.js";
import {
  normalizeSelectionRange,
  selectionRangesEqualNormalized,
  selectionRangesIntersect,
  type SelectionRange,
} from "./selection-range.js";

/** 列筛选按钮绘制位置：列标题栏或选区上一行的表体单元格。 */
export type ColumnAutoFilterUiKind = "header" | "body";

/** 该列在筛选菜单中最后一次应用的排序方向（用于表头图标）。 */
export type ColumnAutoFilterSortHint = "asc" | "desc" | null;

interface ColumnAutoFilterMutable {
  checkedKeys: Set<string>;
  includeBlank: boolean;
  fontColorArgb: string | null;
  /** 筛选/排序生效的闭区间行号（仅这些行会被隐藏或参与排序）。 */
  rowStart: number;
  rowEnd: number;
  uiKind: ColumnAutoFilterUiKind;
  /** `uiKind === "body"` 时：筛选按钮所在表体行（通常为选区顶行的上一行）。 */
  bodyAnchorRow: number;
  /** 由「升序/降序/按颜色排序」写入；清除筛选时随条目删除。 */
  lastSortDirection: ColumnAutoFilterSortHint;
}

function cloneColumnAutoFilterState(st: ColumnAutoFilterMutable): ColumnAutoFilterMutable {
  return {
    checkedKeys: new Set(st.checkedKeys),
    includeBlank: st.includeBlank,
    fontColorArgb: st.fontColorArgb,
    rowStart: st.rowStart,
    rowEnd: st.rowEnd,
    uiKind: st.uiKind,
    bodyAnchorRow: st.bodyAnchorRow,
    lastSortDirection: st.lastSortDirection,
  };
}

/** 工作表数据或网格规模变更时触发；由 `Workbook` 汇聚后通知宿主刷新视图。 */
export type WorksheetChangeListener = () => void;

/**
 * 工作表：稀疏存储单元格，行列规模可扩展。
 * 通过 `subscribe` / `notifyDataChanged` 与 `batch` 支持数据驱动渲染。
 */
export class Worksheet {
  private _name: string;
  private readonly cells = new Map<string, Cell>();

  get name(): string {
    return this._name;
  }

  /** 逻辑行数（≥1）。规模变更请用 `setGridSize`。 */
  rowCount: number;
  /** 逻辑列数（≥1）。规模变更请用 `setGridSize`。 */
  colCount: number;
  defaultRowHeight = 20;
  defaultColWidth = 64;
  private readonly rowHeights = new Map<number, number>();
  private readonly colWidths = new Map<number, number>();
  private readonly hiddenRows = new Set<number>();
  private readonly hiddenCols = new Set<number>();
  /** 由列自动筛选临时隐藏的行（与手动 `hiddenRows` 叠加）。 */
  private readonly autoFilterConcealedRows = new Set<number>();
  private readonly autoFilterByCol = new Map<number, ColumnAutoFilterMutable>();

  /** 合并区域：主格键 `row,col` → 跨度（≥2 格）。 */
  private mergeRegionsByMaster = new Map<
    string,
    { readonly rowSpan: number; readonly colSpan: number }
  >();
  /** 格键 → 主格键（含主格自身）。 */
  private mergeCellToMaster = new Map<string, string>();

  private readonly changeListeners = new Set<WorksheetChangeListener>();
  private batchDepth = 0;
  private pendingNotify = false;
  private _revision = 0;

  /** 条件格式规则（自上而下求值，首条匹配生效）。 */
  private conditionalFormatRules: ConditionalFormatRule[] = [];

  constructor(name: string, rowCount = 1000, colCount = 26) {
    this._name = name;
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
    this.reindexMetadataOnGridClamp();
    this.clampAutoFiltersToGridSize();
    this.touchData();
  }

  getRowHeight(row: number): number {
    if (!Number.isInteger(row) || row < 0 || row >= this.rowCount) {
      return this.defaultRowHeight;
    }
    return this.rowHeights.get(row) ?? this.defaultRowHeight;
  }

  getColWidth(col: number): number {
    if (!Number.isInteger(col) || col < 0 || col >= this.colCount) {
      return this.defaultColWidth;
    }
    return this.colWidths.get(col) ?? this.defaultColWidth;
  }

  setRowHeight(row: number, height: number): void {
    if (!Number.isInteger(row) || row < 0 || row >= this.rowCount) {
      return;
    }
    const h = Math.max(2, Math.trunc(height));
    if (h === this.defaultRowHeight) {
      this.rowHeights.delete(row);
    } else {
      this.rowHeights.set(row, h);
    }
    this.touchData();
  }

  setColWidth(col: number, width: number): void {
    if (!Number.isInteger(col) || col < 0 || col >= this.colCount) {
      return;
    }
    const w = Math.max(2, Math.trunc(width));
    if (w === this.defaultColWidth) {
      this.colWidths.delete(col);
    } else {
      this.colWidths.set(col, w);
    }
    this.touchData();
  }

  isRowHidden(row: number): boolean {
    return this.hiddenRows.has(row) || this.autoFilterConcealedRows.has(row);
  }

  isColHidden(col: number): boolean {
    return this.hiddenCols.has(col);
  }

  setRowHidden(row: number, hidden: boolean): void {
    if (!Number.isInteger(row) || row < 0 || row >= this.rowCount) {
      return;
    }
    if (hidden) {
      this.hiddenRows.add(row);
    } else {
      this.hiddenRows.delete(row);
    }
    this.touchData();
  }

  setColHidden(col: number, hidden: boolean): void {
    if (!Number.isInteger(col) || col < 0 || col >= this.colCount) {
      return;
    }
    if (hidden) {
      this.hiddenCols.add(col);
    } else {
      this.hiddenCols.delete(col);
    }
    this.touchData();
  }

  /** 该列是否已启用自动筛选。 */
  hasColumnAutoFilter(col: number): boolean {
    return (
      Number.isInteger(col) && col >= 0 && col < this.colCount && this.autoFilterByCol.has(col)
    );
  }

  /**
   * 该列筛选是否已缩小可见行（值/颜色条件排除部分行），用于图标「筛选生效」态。
   */
  isColumnAutoFilterNarrowed(col: number): boolean {
    if (!this.hasColumnAutoFilter(col)) {
      return false;
    }
    const st = this.autoFilterByCol.get(col);
    if (st === undefined) {
      return false;
    }
    if (st.fontColorArgb !== null) {
      return true;
    }
    for (let r = st.rowStart; r <= st.rowEnd; r++) {
      if (r < 0 || r >= this.rowCount) {
        continue;
      }
      if (!this.rowPassesColumnAutoFilter(r, col, st)) {
        return true;
      }
    }
    return false;
  }

  /** 该列筛选按钮上应显示的排序提示（无筛选列时 `undefined`）。 */
  getColumnAutoFilterSortHint(col: number): ColumnAutoFilterSortHint | undefined {
    const st = this.autoFilterByCol.get(col);
    if (st === undefined) {
      return undefined;
    }
    return st.lastSortDirection;
  }

  /**
   * 列筛选元数据：作用行范围与按钮位置（标题栏 / 表体锚点格）。
   */
  getColumnAutoFilterMeta(col: number):
    | {
        readonly rowStart: number;
        readonly rowEnd: number;
        readonly uiKind: ColumnAutoFilterUiKind;
        readonly bodyAnchorRow: number;
      }
    | undefined {
    const st = this.autoFilterByCol.get(col);
    if (st === undefined) {
      return undefined;
    }
    return {
      rowStart: st.rowStart,
      rowEnd: st.rowEnd,
      uiKind: st.uiKind,
      bodyAnchorRow: st.bodyAnchorRow,
    };
  }

  /** 供 UI 读取当前列筛选勾选与颜色条件（返回副本）。 */
  getColumnAutoFilterState(col: number):
    | {
        readonly checkedKeys: ReadonlySet<string>;
        readonly includeBlank: boolean;
        readonly fontColorArgb: string | null;
        readonly rowStart: number;
        readonly rowEnd: number;
        readonly uiKind: ColumnAutoFilterUiKind;
        readonly bodyAnchorRow: number;
      }
    | undefined {
    const st = this.autoFilterByCol.get(col);
    if (st === undefined) {
      return undefined;
    }
    return {
      checkedKeys: new Set(st.checkedKeys),
      includeBlank: st.includeBlank,
      fontColorArgb: st.fontColorArgb,
      rowStart: st.rowStart,
      rowEnd: st.rowEnd,
      uiKind: st.uiKind,
      bodyAnchorRow: st.bodyAnchorRow,
    };
  }

  /**
   * 按当前选区启用列自动筛选：作用域为选区行范围；选区含第 1 行时按钮在列标题，否则在选区顶行上一格。
   * 初始为范围内**全选**（所有去重显示值及空白项均勾选）。`col` 为筛选列；`_valueRow` 为右键格，仅作合法性校验。
   */
  enableColumnAutoFilterFromSelection(
    _valueRow: number,
    col: number,
    selection: SelectionRange,
  ): void {
    if (
      !Number.isInteger(_valueRow) ||
      !Number.isInteger(col) ||
      _valueRow < 0 ||
      _valueRow >= this.rowCount ||
      col < 0 ||
      col >= this.colCount
    ) {
      return;
    }
    const n = normalizeSelectionRange(selection);
    const rowStart = clampIndex(n.startRow, 0, this.rowCount - 1);
    const rowEnd = clampIndex(n.endRow, rowStart, this.rowCount - 1);
    const selectionIncludesFirstRow = n.startRow <= 0 && n.endRow >= 0;
    const uiKind: ColumnAutoFilterUiKind = selectionIncludesFirstRow ? "header" : "body";
    const bodyAnchorRow = selectionIncludesFirstRow ? 0 : Math.max(0, n.startRow - 1);

    const checkedKeys = new Set<string>();
    let includeBlank = false;
    const seenNonBlank = new Set<string>();
    for (let r = rowStart; r <= rowEnd; r++) {
      const key = cellToAutoFilterDisplayKey(this.getCell(r, col));
      if (key === AUTO_FILTER_BLANK_KEY) {
        includeBlank = true;
      } else if (!seenNonBlank.has(key)) {
        seenNonBlank.add(key);
        checkedKeys.add(key);
      }
    }
    this.autoFilterByCol.set(col, {
      checkedKeys,
      includeBlank,
      fontColorArgb: null,
      rowStart,
      rowEnd,
      uiKind,
      bodyAnchorRow,
      lastSortDirection: null,
    });
    this.refreshAutoFilterConcealment();
  }

  /** 更新列的值列表勾选与是否包含空白。 */
  updateColumnAutoFilterValueSelection(
    col: number,
    checkedKeys: ReadonlySet<string>,
    includeBlank: boolean,
  ): void {
    const st = this.autoFilterByCol.get(col);
    if (st === undefined) {
      return;
    }
    st.checkedKeys = new Set(checkedKeys);
    st.includeBlank = includeBlank;
    this.refreshAutoFilterConcealment();
  }

  /** 按字体颜色 ARGB（大写）筛选；`null` 表示不按颜色筛选。 */
  setColumnAutoFilterFontColorArgb(col: number, fontColorArgb: string | null): void {
    const st = this.autoFilterByCol.get(col);
    if (st === undefined) {
      return;
    }
    st.fontColorArgb = fontColorArgb;
    this.refreshAutoFilterConcealment();
  }

  /** 移除该列的自动筛选。 */
  removeColumnAutoFilter(col: number): void {
    if (!this.autoFilterByCol.delete(col)) {
      return;
    }
    this.refreshAutoFilterConcealment();
  }

  /** 列中曾出现过的字体前景色 ARGB（去重，用于「按颜色」子菜单）；仅扫描该列筛选作用行范围。 */
  collectUniqueFontColorArgbsInColumn(col: number): string[] {
    if (!Number.isInteger(col) || col < 0 || col >= this.colCount) {
      return [];
    }
    const st = this.autoFilterByCol.get(col);
    const r0 = st !== undefined ? st.rowStart : 0;
    const r1 = st !== undefined ? st.rowEnd : this.rowCount - 1;
    const seen = new Set<string>();
    const out: string[] = [];
    for (let r = r0; r <= r1; r++) {
      if (r < 0 || r >= this.rowCount) {
        continue;
      }
      const anchor = this.getMergeAnchorCell(r, col);
      if (anchor.row !== r || anchor.col !== col) {
        continue;
      }
      const fg = this.getCell(r, col).style?.fgArgb;
      if (fg === undefined || fg === "") {
        continue;
      }
      const u = fg.toUpperCase();
      if (!seen.has(u)) {
        seen.add(u);
        out.push(u);
      }
    }
    return out.sort();
  }

  /**
   * 在闭区间行范围内按指定列排序（仅重排区间内行，区间外不动）。
   */
  sortRowsInRangeByColumn(
    rowStart: number,
    rowEnd: number,
    sortCol: number,
    direction: "asc" | "desc",
  ): void {
    if (!Number.isInteger(sortCol) || sortCol < 0 || sortCol >= this.colCount) {
      return;
    }
    const lo = clampIndex(rowStart, 0, this.rowCount - 1);
    const hi = clampIndex(rowEnd, lo, this.rowCount - 1);
    const len = hi - lo + 1;
    if (len <= 1) {
      return;
    }
    const dir = direction === "asc" ? 1 : -1;
    const sortedOld = Array.from({ length: len }, (_, i) => lo + i);
    sortedOld.sort((a, b) => this.compareRowsForColumnSort(a, b, sortCol, dir));
    const inv = this.buildIdentityRowMap();
    for (let i = 0; i < len; i++) {
      inv[sortedOld[i]!] = lo + i;
    }
    const stCol = this.autoFilterByCol.get(sortCol);
    if (stCol !== undefined) {
      stCol.lastSortDirection = direction;
    }
    this.applyRowPermutation(inv);
  }

  /**
   * 在闭区间行范围内按字体颜色排序；`targetArgb === null` 表示「自动」优先。
   */
  sortRowsInRangeByColumnFontColor(
    rowStart: number,
    rowEnd: number,
    sortCol: number,
    targetArgb: string | null,
    direction: "asc" | "desc",
  ): void {
    if (!Number.isInteger(sortCol) || sortCol < 0 || sortCol >= this.colCount) {
      return;
    }
    const lo = clampIndex(rowStart, 0, this.rowCount - 1);
    const hi = clampIndex(rowEnd, lo, this.rowCount - 1);
    const len = hi - lo + 1;
    if (len <= 1) {
      return;
    }
    const want = targetArgb === null ? null : targetArgb.toUpperCase();
    const dir = direction === "asc" ? 1 : -1;
    const sortedOld = Array.from({ length: len }, (_, i) => lo + i);
    sortedOld.sort((a, b) => {
      const ma = this.getMergeAnchorCell(a, sortCol);
      const mb = this.getMergeAnchorCell(b, sortCol);
      const fa = (this.getCell(ma.row, ma.col).style?.fgArgb ?? "").toUpperCase();
      const fb = (this.getCell(mb.row, mb.col).style?.fgArgb ?? "").toUpperCase();
      const groupA = want === null ? (fa === "" ? 0 : 1) : fa === want ? 0 : 1;
      const groupB = want === null ? (fb === "" ? 0 : 1) : fb === want ? 0 : 1;
      if (groupA !== groupB) {
        return dir * (groupA - groupB);
      }
      return this.compareRowsForColumnSort(a, b, sortCol, dir);
    });
    const inv = this.buildIdentityRowMap();
    for (let i = 0; i < len; i++) {
      inv[sortedOld[i]!] = lo + i;
    }
    const stCol = this.autoFilterByCol.get(sortCol);
    if (stCol !== undefined) {
      stCol.lastSortDirection = direction;
    }
    this.applyRowPermutation(inv);
  }

  /**
   * 在闭区间行范围内按单元格填充色排序；`targetArgb === null` 表示「无填充」优先。
   */
  sortRowsInRangeByColumnFillColor(
    rowStart: number,
    rowEnd: number,
    sortCol: number,
    targetArgb: string | null,
    direction: "asc" | "desc",
  ): void {
    if (!Number.isInteger(sortCol) || sortCol < 0 || sortCol >= this.colCount) {
      return;
    }
    const lo = clampIndex(rowStart, 0, this.rowCount - 1);
    const hi = clampIndex(rowEnd, lo, this.rowCount - 1);
    const len = hi - lo + 1;
    if (len <= 1) {
      return;
    }
    const want = targetArgb === null ? null : targetArgb.toUpperCase();
    const dir = direction === "asc" ? 1 : -1;
    const sortedOld = Array.from({ length: len }, (_, i) => lo + i);
    sortedOld.sort((a, b) => {
      const ma = this.getMergeAnchorCell(a, sortCol);
      const mb = this.getMergeAnchorCell(b, sortCol);
      const fa = (this.getCell(ma.row, ma.col).style?.fillArgb ?? "").toUpperCase();
      const fb = (this.getCell(mb.row, mb.col).style?.fillArgb ?? "").toUpperCase();
      const groupA = want === null ? (fa === "" ? 0 : 1) : fa === want ? 0 : 1;
      const groupB = want === null ? (fb === "" ? 0 : 1) : fb === want ? 0 : 1;
      if (groupA !== groupB) {
        return dir * (groupA - groupB);
      }
      return this.compareRowsForColumnSort(a, b, sortCol, dir);
    });
    const inv = this.buildIdentityRowMap();
    for (let i = 0; i < len; i++) {
      inv[sortedOld[i]!] = lo + i;
    }
    const stCol = this.autoFilterByCol.get(sortCol);
    if (stCol !== undefined) {
      stCol.lastSortDirection = direction;
    }
    this.applyRowPermutation(inv);
  }

  /** 全表按列排序（区间内实现与 `sortRowsInRangeByColumn` 一致）。 */
  sortRowsByColumn(sortCol: number, direction: "asc" | "desc"): void {
    if (this.rowCount <= 1) {
      return;
    }
    this.sortRowsInRangeByColumn(0, this.rowCount - 1, sortCol, direction);
  }

  /** 全表按字体颜色排序。 */
  sortRowsByColumnFontColor(
    sortCol: number,
    targetArgb: string | null,
    direction: "asc" | "desc",
  ): void {
    if (this.rowCount <= 1) {
      return;
    }
    this.sortRowsInRangeByColumnFontColor(0, this.rowCount - 1, sortCol, targetArgb, direction);
  }

  insertRows(atRow: number, count: number): void {
    const start = clampIndex(atRow, 0, this.rowCount);
    const n = Math.max(1, Math.trunc(count));
    this.batch(() => {
      this.shiftAutoFilterRowIndicesOnInsert(start, n);
      this.reindexCells((row, col) => (row >= start ? { row: row + n, col } : { row, col }));
      this.reindexMergeRegions((row, col) => (row >= start ? { row: row + n, col } : { row, col }));
      this.reindexRowMetadata((row) => (row >= start ? row + n : row));
      this.rowCount += n;
      this.refreshAutoFilterConcealment();
      this.notifyDataChanged();
    });
  }

  deleteRows(atRow: number, count: number): void {
    if (this.rowCount <= 1) {
      return;
    }
    const start = clampIndex(atRow, 0, this.rowCount - 1);
    const n = Math.max(1, Math.trunc(count));
    const end = Math.min(this.rowCount - 1, start + n - 1);
    const removed = end - start + 1;
    const nextRowCount = Math.max(1, this.rowCount - removed);
    this.batch(() => {
      this.shiftAutoFilterRowIndicesOnDelete(start, end, removed);
      this.reindexCells((row, col) => {
        if (row < start) {
          return { row, col };
        }
        if (row > end) {
          return { row: row - removed, col };
        }
        return null;
      });
      this.reindexMergeRegions((row, col) => {
        if (row < start) {
          return { row, col };
        }
        if (row > end) {
          return { row: row - removed, col };
        }
        return null;
      });
      this.reindexRowMetadata((row) => {
        if (row < start) {
          return row;
        }
        if (row > end) {
          return row - removed;
        }
        return null;
      });
      this.rowCount = nextRowCount;
      this.refreshAutoFilterConcealment();
      this.notifyDataChanged();
    });
  }

  insertCols(atCol: number, count: number): void {
    const start = clampIndex(atCol, 0, this.colCount);
    const n = Math.max(1, Math.trunc(count));
    this.batch(() => {
      this.reindexCells((row, col) => (col >= start ? { row, col: col + n } : { row, col }));
      this.reindexMergeRegions((row, col) => (col >= start ? { row, col: col + n } : { row, col }));
      this.reindexColMetadata((col) => (col >= start ? col + n : col));
      this.reindexAutoFilterColKeys((col) => (col >= start ? col + n : col));
      this.colCount += n;
      this.refreshAutoFilterConcealment();
      this.notifyDataChanged();
    });
  }

  deleteCols(atCol: number, count: number): void {
    if (this.colCount <= 1) {
      return;
    }
    const start = clampIndex(atCol, 0, this.colCount - 1);
    const n = Math.max(1, Math.trunc(count));
    const end = Math.min(this.colCount - 1, start + n - 1);
    const removed = end - start + 1;
    const nextColCount = Math.max(1, this.colCount - removed);
    this.batch(() => {
      this.reindexCells((row, col) => {
        if (col < start) {
          return { row, col };
        }
        if (col > end) {
          return { row, col: col - removed };
        }
        return null;
      });
      this.reindexMergeRegions((row, col) => {
        if (col < start) {
          return { row, col };
        }
        if (col > end) {
          return { row, col: col - removed };
        }
        return null;
      });
      this.reindexColMetadata((col) => {
        if (col < start) {
          return col;
        }
        if (col > end) {
          return col - removed;
        }
        return null;
      });
      this.reindexAutoFilterColKeys((col) => {
        if (col < start) {
          return col;
        }
        if (col > end) {
          return col - removed;
        }
        return null;
      });
      this.colCount = nextColCount;
      this.refreshAutoFilterConcealment();
      this.notifyDataChanged();
    });
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

  getConditionalFormatRules(): readonly ConditionalFormatRule[] {
    return this.conditionalFormatRules;
  }

  /** 替换整张表的条件格式规则列表。 */
  setConditionalFormatRules(rules: readonly ConditionalFormatRule[]): void {
    this.conditionalFormatRules = rules.map((r) => ({ ...r, range: normalizeSelectionRange(r.range) }));
    this.touchData();
  }

  /**
   * 追加一条条件格式规则。
   * 若已有规则的适用选区与新区规范化后完全一致，则先移除这些规则再追加（同一选区重复设置视为用新规则覆盖）。
   */
  addConditionalFormatRule(rule: ConditionalFormatRule): void {
    const n = normalizeSelectionRange(rule.range);
    const filtered = this.conditionalFormatRules.filter(
      (r) => !selectionRangesEqualNormalized(r.range, n),
    );
    this.conditionalFormatRules = [...filtered, { ...rule, range: n }];
    this.touchData();
  }

  removeConditionalFormatRuleById(ruleId: string): void {
    const next = this.conditionalFormatRules.filter((r) => r.id !== ruleId);
    if (next.length === this.conditionalFormatRules.length) {
      return;
    }
    this.conditionalFormatRules = next;
    this.touchData();
  }

  /** 删除与给定选区相交的所有规则。 */
  clearConditionalFormatRulesIntersecting(range: SelectionRange): void {
    const n = normalizeSelectionRange(range);
    const next = this.conditionalFormatRules.filter((r) => !selectionRangesIntersect(r.range, n));
    if (next.length === this.conditionalFormatRules.length) {
      return;
    }
    this.conditionalFormatRules = next;
    this.touchData();
  }

  clearAllConditionalFormatRules(): void {
    if (this.conditionalFormatRules.length === 0) {
      return;
    }
    this.conditionalFormatRules = [];
    this.touchData();
  }

  /** 供 Canvas 绘制：首条匹配规则的预设/自定义样式叠加（不改变单元格持久样式）。 */
  resolveConditionalFormattingCellOverlay(
    row: number,
    col: number,
  ): ConditionalFormattingOverlay | null {
    return resolveConditionalFormattingOverlay(this, row, col);
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

  /**
   * 若位于合并区域则返回主格；否则返回 `row,col`。
   */
  getMergeAnchorCell(row: number, col: number): { readonly row: number; readonly col: number } {
    const m = this.mergeCellToMaster.get(Cell.key(row, col));
    if (m === undefined) {
      return { row, col };
    }
    const [mr, mc] = parseCellKey(m);
    return { row: mr, col: mc };
  }

  /** 是否合并区域内的非主格（被覆盖格）。 */
  isMergeCoveredCell(row: number, col: number): boolean {
    const m = this.mergeCellToMaster.get(Cell.key(row, col));
    return m !== undefined && m !== Cell.key(row, col);
  }

  /**
   * 合并区域信息：主格为 `anchorRow/anchorCol`；未合并时为 1×1。
   */
  getMergedRectInfo(
    row: number,
    col: number,
  ): {
    readonly anchorRow: number;
    readonly anchorCol: number;
    readonly rowSpan: number;
    readonly colSpan: number;
  } {
    const anchor = this.getMergeAnchorCell(row, col);
    const span = this.mergeRegionsByMaster.get(Cell.key(anchor.row, anchor.col));
    if (span === undefined) {
      return { anchorRow: anchor.row, anchorCol: anchor.col, rowSpan: 1, colSpan: 1 };
    }
    return {
      anchorRow: anchor.row,
      anchorCol: anchor.col,
      rowSpan: span.rowSpan,
      colSpan: span.colSpan,
    };
  }

  /**
   * 键盘方向键移动一步（不含 Shift 扩展）：从合并区域穿出时按 Excel 习惯从主格边缘离开。
   */
  moveFocusFrom(
    row: number,
    col: number,
    deltaRow: number,
    deltaCol: number,
  ): {
    readonly row: number;
    readonly col: number;
  } {
    const dr = Math.sign(deltaRow);
    const dc = Math.sign(deltaCol);
    if (dr === 0 && dc === 0) {
      return { row, col };
    }
    const info = this.getMergedRectInfo(row, col);
    const anchor = this.getMergeAnchorCell(row, col);
    const mr = anchor.row;
    const mc = anchor.col;
    if (info.rowSpan > 1 || info.colSpan > 1) {
      const endR = mr + info.rowSpan - 1;
      const endC = mc + info.colSpan - 1;
      if (dr === 1) {
        return { row: Math.min(this.rowCount - 1, endR + 1), col: mc };
      }
      if (dr === -1) {
        return { row: Math.max(0, mr - 1), col: mc };
      }
      if (dc === 1) {
        return { row: mr, col: Math.min(this.colCount - 1, endC + 1) };
      }
      if (dc === -1) {
        return { row: mr, col: Math.max(0, mc - 1) };
      }
    }
    return {
      row: clampIndex(row + dr, 0, this.rowCount - 1),
      col: clampIndex(col + dc, 0, this.colCount - 1),
    };
  }

  /** 供撤销/重做：序列化当前合并区域。 */
  getMergeRegionsSnapshot(): readonly {
    readonly masterRow: number;
    readonly masterCol: number;
    readonly rowSpan: number;
    readonly colSpan: number;
  }[] {
    const out: {
      masterRow: number;
      masterCol: number;
      rowSpan: number;
      colSpan: number;
    }[] = [];
    for (const [k, span] of this.mergeRegionsByMaster) {
      const [mr, mc] = parseCellKey(k);
      out.push({ masterRow: mr, masterCol: mc, rowSpan: span.rowSpan, colSpan: span.colSpan });
    }
    return out;
  }

  /** 供撤销/重做：恢复合并区域（会覆盖现有合并映射）。 */
  restoreMergeRegionsFromSnapshot(
    regions: readonly {
      readonly masterRow: number;
      readonly masterCol: number;
      readonly rowSpan: number;
      readonly colSpan: number;
    }[],
  ): void {
    this.mergeRegionsByMaster.clear();
    this.mergeCellToMaster.clear();
    for (const r of regions) {
      if (r.rowSpan <= 1 && r.colSpan <= 1) {
        continue;
      }
      this.mergeRegionsByMaster.set(Cell.key(r.masterRow, r.masterCol), {
        rowSpan: r.rowSpan,
        colSpan: r.colSpan,
      });
    }
    this.rebuildMergeCellIndex();
    this.touchData();
  }

  /**
   * 合并选区：与 `mergeCells` / `mergeAcross` / `mergeCenter` 对应；单格无操作。
   */
  applyMergeForSelection(
    range: SelectionRange,
    kind: "mergeCells" | "mergeAcross" | "mergeCenter",
  ): void {
    const n = normalizeSelectionRange(range);
    if (n.startRow === n.endRow && n.startCol === n.endCol) {
      return;
    }
    this.batch(() => {
      this.clearMergeRegionsIntersecting(n);
      if (kind === "mergeAcross") {
        for (let r = n.startRow; r <= n.endRow; r++) {
          if (n.startCol === n.endCol) {
            continue;
          }
          this.setMergeRegionCore(r, n.startCol, 1, n.endCol - n.startCol + 1);
        }
      } else {
        this.setMergeRegionCore(
          n.startRow,
          n.startCol,
          n.endRow - n.startRow + 1,
          n.endCol - n.startCol + 1,
        );
        if (kind === "mergeCenter") {
          const cell = this.getCell(n.startRow, n.startCol);
          this.setCellStyle(
            n.startRow,
            n.startCol,
            applyCellStylePatch(cell.style, { hAlign: "center", indentLevel: null }),
          );
        }
      }
    });
  }

  /** 取消与选区相交的合并。 */
  applyUnmergeForSelection(range: SelectionRange): void {
    const n = normalizeSelectionRange(range);
    this.batch(() => {
      const toRemove: string[] = [];
      for (const [mKey, span] of this.mergeRegionsByMaster) {
        const [mr, mc] = parseCellKey(mKey);
        const endR = mr + span.rowSpan - 1;
        const endC = mc + span.colSpan - 1;
        if (rangesIntersect(n, mr, endR, mc, endC)) {
          toRemove.push(mKey);
        }
      }
      if (toRemove.length === 0) {
        return;
      }
      for (const k of toRemove) {
        this.mergeRegionsByMaster.delete(k);
      }
      this.rebuildMergeCellIndex();
      this.touchData();
    });
  }

  hasCell(row: number, col: number): boolean {
    return this.cells.has(Cell.key(row, col));
  }

  /** 重命名工作表并通知监听者（用于标签栏等 UI）。 */
  setName(next: string): void {
    const t = next.trim();
    if (t.length === 0 || t === this._name) {
      return;
    }
    this._name = t;
    this.flushNotify();
  }

  /** 该列筛选作用行范围内出现的去重筛选显示键（已排序，含空白键）。 */
  collectUniqueAutoFilterKeysInColumn(col: number): string[] {
    if (!Number.isInteger(col) || col < 0 || col >= this.colCount) {
      return [];
    }
    const st = this.autoFilterByCol.get(col);
    const r0 = st !== undefined ? st.rowStart : 0;
    const r1 = st !== undefined ? st.rowEnd : this.rowCount - 1;
    const seen = new Set<string>();
    const raw: string[] = [];
    for (let r = r0; r <= r1; r++) {
      if (r < 0 || r >= this.rowCount) {
        continue;
      }
      const k = cellToAutoFilterDisplayKey(this.getCell(r, col));
      if (!seen.has(k)) {
        seen.add(k);
        raw.push(k);
      }
    }
    return raw.sort(compareAutoFilterDisplayKeys);
  }

  private compareRowsForColumnSort(
    rowA: number,
    rowB: number,
    sortCol: number,
    dir: 1 | -1,
  ): number {
    const ca = this.getCell(rowA, sortCol);
    const cb = this.getCell(rowB, sortCol);
    const va = ca.value;
    const vb = cb.value;
    if (
      typeof va === "number" &&
      typeof vb === "number" &&
      !Number.isNaN(va) &&
      !Number.isNaN(vb)
    ) {
      if (va !== vb) {
        return dir * (va < vb ? -1 : 1);
      }
    } else {
      const sa = formatCellDisplayWithStyle(va, ca.style);
      const sb = formatCellDisplayWithStyle(vb, cb.style);
      if (sa !== sb) {
        return dir * sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
      }
    }
    return rowA - rowB;
  }

  private buildIdentityRowMap(): number[] {
    return Array.from({ length: this.rowCount }, (_, i) => i);
  }

  private applyRowPermutation(inv: readonly number[]): void {
    this.batch(() => {
      this.reindexCells((r, c) => ({ row: inv[r]!, col: c }));
      this.reindexMergeRegions((r, c) => ({ row: inv[r]!, col: c }));
      this.reindexRowMetadata((r) => inv[r]!);
      this.refreshAutoFilterConcealment();
      this.notifyDataChanged();
    });
  }

  private shiftAutoFilterRowIndicesOnInsert(atRow: number, n: number): void {
    for (const st of this.autoFilterByCol.values()) {
      if (st.rowStart >= atRow) {
        st.rowStart += n;
      }
      if (st.rowEnd >= atRow) {
        st.rowEnd += n;
      }
      if (st.bodyAnchorRow >= atRow) {
        st.bodyAnchorRow += n;
      }
    }
  }

  private shiftAutoFilterRowIndicesOnDelete(start: number, end: number, removed: number): void {
    const mapR = (r: number): number | null => {
      if (r < start) {
        return r;
      }
      if (r > end) {
        return r - removed;
      }
      return null;
    };
    const toRemove: number[] = [];
    for (const [c, st] of this.autoFilterByCol) {
      const ns = mapR(st.rowStart);
      const ne = mapR(st.rowEnd);
      const nb = mapR(st.bodyAnchorRow);
      if (ns === null || ne === null || nb === null || ns > ne) {
        toRemove.push(c);
        continue;
      }
      st.rowStart = ns;
      st.rowEnd = ne;
      st.bodyAnchorRow = nb;
    }
    for (const c of toRemove) {
      this.autoFilterByCol.delete(c);
    }
  }

  private refreshAutoFilterConcealment(): void {
    this.autoFilterConcealedRows.clear();
    if (this.autoFilterByCol.size === 0) {
      this.touchData();
      return;
    }
    for (let r = 0; r < this.rowCount; r++) {
      if (!this.rowPassesAllAutoFilters(r)) {
        this.autoFilterConcealedRows.add(r);
      }
    }
    this.touchData();
  }

  private rowPassesAllAutoFilters(row: number): boolean {
    for (const [col, st] of this.autoFilterByCol) {
      if (row < st.rowStart || row > st.rowEnd) {
        continue;
      }
      if (!this.rowPassesColumnAutoFilter(row, col, st)) {
        return false;
      }
    }
    return true;
  }

  private rowPassesColumnAutoFilter(
    row: number,
    col: number,
    st: ColumnAutoFilterMutable,
  ): boolean {
    const anchor = this.getMergeAnchorCell(row, col);
    const cell = this.getCell(anchor.row, anchor.col);
    const displayKey = cellToAutoFilterDisplayKey(cell);
    if (displayKey === AUTO_FILTER_BLANK_KEY) {
      if (!st.includeBlank) {
        return false;
      }
    } else if (!st.checkedKeys.has(displayKey)) {
      return false;
    }
    if (st.fontColorArgb !== null) {
      const fg = (cell.style?.fgArgb ?? "").toUpperCase();
      if (fg !== st.fontColorArgb.toUpperCase()) {
        return false;
      }
    }
    return true;
  }

  private reindexAutoFilterColKeys(mapper: (col: number) => number | null): void {
    const next = new Map<number, ColumnAutoFilterMutable>();
    for (const [c, st] of this.autoFilterByCol) {
      const nc = mapper(c);
      if (nc === null || nc < 0) {
        continue;
      }
      next.set(nc, cloneColumnAutoFilterState(st));
    }
    this.autoFilterByCol.clear();
    for (const [c, st] of next) {
      this.autoFilterByCol.set(c, st);
    }
  }

  private clampAutoFiltersToGridSize(): void {
    const next = new Map<number, ColumnAutoFilterMutable>();
    for (const [c, st] of this.autoFilterByCol) {
      if (c >= 0 && c < this.colCount) {
        next.set(c, st);
      }
    }
    this.autoFilterByCol.clear();
    for (const [c, st] of next) {
      this.autoFilterByCol.set(c, st);
    }
    this.refreshAutoFilterConcealment();
  }

  private reindexCells(
    mapper: (row: number, col: number) => { row: number; col: number } | null,
  ): void {
    const next = new Map<string, Cell>();
    for (const oldCell of this.cells.values()) {
      const target = mapper(oldCell.row, oldCell.col);
      if (target === null) {
        continue;
      }
      if (target.row < 0 || target.col < 0) {
        continue;
      }
      const cell = new Cell(target.row, target.col, oldCell.value);
      cell.formula = oldCell.formula;
      cell.style = oldCell.style;
      next.set(Cell.key(target.row, target.col), cell);
    }
    this.cells.clear();
    for (const [k, v] of next) {
      this.cells.set(k, v);
    }
  }

  private reindexRowMetadata(mapper: (row: number) => number | null): void {
    this.reindexNumberMap(this.rowHeights, mapper);
    this.reindexNumberSet(this.hiddenRows, mapper);
    this.reindexNumberSet(this.autoFilterConcealedRows, mapper);
  }

  private reindexColMetadata(mapper: (col: number) => number | null): void {
    this.reindexNumberMap(this.colWidths, mapper);
    this.reindexNumberSet(this.hiddenCols, mapper);
  }

  private reindexNumberMap(
    map: Map<number, number>,
    mapper: (index: number) => number | null,
  ): void {
    const next = new Map<number, number>();
    for (const [idx, value] of map) {
      const target = mapper(idx);
      if (target === null || target < 0) {
        continue;
      }
      next.set(target, value);
    }
    map.clear();
    for (const [idx, value] of next) {
      map.set(idx, value);
    }
  }

  private reindexNumberSet(set: Set<number>, mapper: (index: number) => number | null): void {
    const next = new Set<number>();
    for (const idx of set) {
      const target = mapper(idx);
      if (target === null || target < 0) {
        continue;
      }
      next.add(target);
    }
    set.clear();
    for (const idx of next) {
      set.add(idx);
    }
  }

  private reindexMetadataOnGridClamp(): void {
    this.reindexRowMetadata((row) => (row >= this.rowCount ? null : row));
    this.reindexColMetadata((col) => (col >= this.colCount ? null : col));
    this.reindexMergeRegions((row, col) => {
      if (row < 0 || col < 0 || row >= this.rowCount || col >= this.colCount) {
        return null;
      }
      return { row, col };
    });
  }

  private rebuildMergeCellIndex(): void {
    this.mergeCellToMaster.clear();
    for (const [mKey, span] of this.mergeRegionsByMaster) {
      const [mr, mc] = parseCellKey(mKey);
      for (let r = mr; r < mr + span.rowSpan; r++) {
        for (let c = mc; c < mc + span.colSpan; c++) {
          this.mergeCellToMaster.set(Cell.key(r, c), mKey);
        }
      }
    }
  }

  private reindexMergeRegions(
    mapper: (row: number, col: number) => { row: number; col: number } | null,
  ): void {
    const next = new Map<string, { rowSpan: number; colSpan: number }>();
    for (const [mKey, span] of this.mergeRegionsByMaster) {
      const [mr, mc] = parseCellKey(mKey);
      let ok = true;
      for (let r = mr; r < mr + span.rowSpan && ok; r++) {
        for (let c = mc; c < mc + span.colSpan && ok; c++) {
          if (mapper(r, c) === null) {
            ok = false;
          }
        }
      }
      if (!ok) {
        continue;
      }
      const nm = mapper(mr, mc)!;
      const br = mapper(mr + span.rowSpan - 1, mc + span.colSpan - 1)!;
      next.set(Cell.key(nm.row, nm.col), {
        rowSpan: br.row - nm.row + 1,
        colSpan: br.col - nm.col + 1,
      });
    }
    this.mergeRegionsByMaster.clear();
    for (const [k, v] of next) {
      this.mergeRegionsByMaster.set(k, v);
    }
    this.rebuildMergeCellIndex();
  }

  private clearMergeRegionsIntersecting(n: SelectionRange): void {
    const toRemove: string[] = [];
    for (const [mKey, span] of this.mergeRegionsByMaster) {
      const [mr, mc] = parseCellKey(mKey);
      const endR = mr + span.rowSpan - 1;
      const endC = mc + span.colSpan - 1;
      if (rangesIntersect(n, mr, endR, mc, endC)) {
        toRemove.push(mKey);
      }
    }
    if (toRemove.length === 0) {
      return;
    }
    for (const k of toRemove) {
      this.mergeRegionsByMaster.delete(k);
    }
    this.rebuildMergeCellIndex();
    this.touchData();
  }

  private setMergeRegionCore(mr: number, mc: number, rowSpan: number, colSpan: number): void {
    if (rowSpan <= 0 || colSpan <= 0) {
      return;
    }
    if (rowSpan === 1 && colSpan === 1) {
      return;
    }
    const endR = mr + rowSpan - 1;
    const endC = mc + colSpan - 1;
    const clip = normalizeSelectionRange({
      startRow: mr,
      endRow: endR,
      startCol: mc,
      endCol: endC,
    });
    this.clearMergeRegionsIntersecting(clip);
    for (let r = mr; r <= endR; r++) {
      for (let c = mc; c <= endC; c++) {
        if (r === mr && c === mc) {
          continue;
        }
        this.setCellLiteral(r, c, null);
      }
    }
    this.mergeRegionsByMaster.set(Cell.key(mr, mc), { rowSpan, colSpan });
    this.rebuildMergeCellIndex();
  }
}

function parseCellKey(key: string): [number, number] {
  const i = key.indexOf(",");
  return [Number(key.slice(0, i)), Number(key.slice(i + 1))];
}

function rangesIntersect(
  sel: SelectionRange,
  r0: number,
  r1: number,
  c0: number,
  c1: number,
): boolean {
  const n = normalizeSelectionRange(sel);
  return !(n.endRow < r0 || n.startRow > r1 || n.endCol < c0 || n.startCol > c1);
}

function clampIndex(v: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, Math.trunc(v)));
}

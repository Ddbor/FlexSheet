import { Cell, applyCellStylePatch, type CellScalar, type CellStyle } from "./cell.js";
import { normalizeSelectionRange, type SelectionRange } from "./selection-range.js";

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

  /** 合并区域：主格键 `row,col` → 跨度（≥2 格）。 */
  private mergeRegionsByMaster = new Map<string, { readonly rowSpan: number; readonly colSpan: number }>();
  /** 格键 → 主格键（含主格自身）。 */
  private mergeCellToMaster = new Map<string, string>();

  private readonly changeListeners = new Set<WorksheetChangeListener>();
  private batchDepth = 0;
  private pendingNotify = false;
  private _revision = 0;

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
    return this.hiddenRows.has(row);
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

  insertRows(atRow: number, count: number): void {
    const start = clampIndex(atRow, 0, this.rowCount);
    const n = Math.max(1, Math.trunc(count));
    this.batch(() => {
      this.reindexCells((row, col) => (row >= start ? { row: row + n, col } : { row, col }));
      this.reindexMergeRegions((row, col) => (row >= start ? { row: row + n, col } : { row, col }));
      this.reindexRowMetadata((row) => (row >= start ? row + n : row));
      this.rowCount += n;
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
      this.colCount += n;
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
      this.colCount = nextColCount;
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
  getMergedRectInfo(row: number, col: number): {
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
    return { anchorRow: anchor.row, anchorCol: anchor.col, rowSpan: span.rowSpan, colSpan: span.colSpan };
  }

  /**
   * 键盘方向键移动一步（不含 Shift 扩展）：从合并区域穿出时按 Excel 习惯从主格边缘离开。
   */
  moveFocusFrom(row: number, col: number, deltaRow: number, deltaCol: number): {
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
            applyCellStylePatch(cell.style, { hAlign: "center" }),
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
  }

  private reindexColMetadata(mapper: (col: number) => number | null): void {
    this.reindexNumberMap(this.colWidths, mapper);
    this.reindexNumberSet(this.hiddenCols, mapper);
  }

  private reindexNumberMap(map: Map<number, number>, mapper: (index: number) => number | null): void {
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

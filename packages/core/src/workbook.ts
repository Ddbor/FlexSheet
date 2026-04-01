import type { Worksheet } from "./worksheet.js";

/** 工作簿级变更（活动表、表集合、任一工作表数据）。 */
export type WorkbookChangeListener = () => void;

/**
 * 工作簿：多工作表容器与活动表索引。
 * `subscribe` 用于在任意表数据或活动表变化时刷新视图（数据驱动渲染）。
 */
export class Workbook {
  private readonly sheets: Worksheet[] = [];
  private readonly sheetUnsubs = new Map<Worksheet, () => void>();
  private readonly listeners = new Set<WorkbookChangeListener>();
  private _activeSheetIndex = 0;

  subscribe(listener: WorkbookChangeListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  private emit(): void {
    for (const fn of this.listeners) {
      fn();
    }
  }

  get activeSheetIndex(): number {
    return this._activeSheetIndex;
  }

  set activeSheetIndex(index: number) {
    const next = Math.trunc(index);
    const clamped =
      this.sheets.length === 0 ? 0 : Math.max(0, Math.min(next, this.sheets.length - 1));
    if (clamped === this._activeSheetIndex) {
      return;
    }
    this._activeSheetIndex = clamped;
    this.emit();
  }

  addSheet(sheet: Worksheet): void {
    this.sheets.push(sheet);
    const unsub = sheet.subscribe(() => {
      this.emit();
    });
    this.sheetUnsubs.set(sheet, unsub);
    this.emit();
  }

  /**
   * 按索引移除工作表；至少保留一张表时返回 true。
   * 移除后会收紧 `activeSheetIndex` 并通知监听者。
   */
  removeSheetAt(index: number): boolean {
    const i = Math.trunc(index);
    if (i < 0 || i >= this.sheets.length || this.sheets.length <= 1) {
      return false;
    }
    const [removed] = this.sheets.splice(i, 1);
    if (removed !== undefined) {
      const unsub = this.sheetUnsubs.get(removed);
      unsub?.();
      this.sheetUnsubs.delete(removed);
    }
    if (this._activeSheetIndex >= this.sheets.length) {
      this._activeSheetIndex = this.sheets.length - 1;
    }
    this.emit();
    return true;
  }

  getSheet(index: number): Worksheet | undefined {
    return this.sheets[index];
  }

  getActiveSheet(): Worksheet | undefined {
    return this.sheets[this._activeSheetIndex];
  }

  get sheetCount(): number {
    return this.sheets.length;
  }

  getSheets(): readonly Worksheet[] {
    return this.sheets;
  }
}

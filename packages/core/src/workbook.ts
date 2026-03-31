import type { Worksheet } from "./worksheet.js";

/** 工作簿级变更（活动表、表集合、任一工作表数据）。 */
export type WorkbookChangeListener = () => void;

/**
 * 工作簿：多工作表容器与活动表索引。
 * `subscribe` 用于在任意表数据或活动表变化时刷新视图（数据驱动渲染）。
 */
export class Workbook {
  private readonly sheets: Worksheet[] = [];
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
    if (next === this._activeSheetIndex) {
      return;
    }
    this._activeSheetIndex = next;
    this.emit();
  }

  addSheet(sheet: Worksheet): void {
    this.sheets.push(sheet);
    sheet.subscribe(() => {
      this.emit();
    });
    this.emit();
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

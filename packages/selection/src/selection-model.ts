import type { Worksheet } from "@flexsheet/core";
import { normalizeSelectionRange, type SelectionRange } from "@flexsheet/core";

export type SheetGetter = () => Worksheet | undefined;

/**
 * 单一数据源：锚点 + 焦点构成矩形选区；焦点即活动单元格。
 */
export class SelectionModel {
  private readonly getSheet: SheetGetter;

  private anchorRow = 0;
  private anchorCol = 0;
  private focusRow = 0;
  private focusCol = 0;

  constructor(getSheet: SheetGetter) {
    this.getSheet = getSheet;
    this.clampToSheet();
  }

  getActiveCell(): { readonly row: number; readonly col: number } {
    return { row: this.focusRow, col: this.focusCol };
  }

  getAnchor(): { readonly row: number; readonly col: number } {
    return { row: this.anchorRow, col: this.anchorCol };
  }

  getNormalizedRange(): SelectionRange {
    return normalizeSelectionRange({
      startRow: this.anchorRow,
      startCol: this.anchorCol,
      endRow: this.focusRow,
      endCol: this.focusCol,
    });
  }

  /** 单击：单格选区，锚点与焦点合一。 */
  selectCell(row: number, col: number): void {
    const sheet = this.getSheet();
    if (sheet === undefined) {
      return;
    }
    const r = clamp(row, 0, sheet.rowCount - 1);
    const c = clamp(col, 0, sheet.colCount - 1);
    this.anchorRow = r;
    this.anchorCol = c;
    this.focusRow = r;
    this.focusCol = c;
  }

  /** 拖拽或 Shift+方向键：锚点不动，扩展焦点。 */
  extendFocusTo(row: number, col: number): void {
    const sheet = this.getSheet();
    if (sheet === undefined) {
      return;
    }
    this.focusRow = clamp(row, 0, sheet.rowCount - 1);
    this.focusCol = clamp(col, 0, sheet.colCount - 1);
  }

  /** 方向键导航：无 Shift 时移动活动格并重置锚点。 */
  moveFocus(deltaRow: number, deltaCol: number, extend: boolean): void {
    const sheet = this.getSheet();
    if (sheet === undefined) {
      return;
    }
    if (!extend) {
      const nr = clamp(this.focusRow + deltaRow, 0, sheet.rowCount - 1);
      const nc = clamp(this.focusCol + deltaCol, 0, sheet.colCount - 1);
      this.anchorRow = nr;
      this.anchorCol = nc;
      this.focusRow = nr;
      this.focusCol = nc;
      return;
    }
    this.focusRow = clamp(this.focusRow + deltaRow, 0, sheet.rowCount - 1);
    this.focusCol = clamp(this.focusCol + deltaCol, 0, sheet.colCount - 1);
  }

  /** 活动表变化时夹紧缩引。 */
  syncWithSheet(): void {
    this.clampToSheet();
  }

  private clampToSheet(): void {
    const sheet = this.getSheet();
    if (sheet === undefined) {
      return;
    }
    this.anchorRow = clamp(this.anchorRow, 0, sheet.rowCount - 1);
    this.anchorCol = clamp(this.anchorCol, 0, sheet.colCount - 1);
    this.focusRow = clamp(this.focusRow, 0, sheet.rowCount - 1);
    this.focusCol = clamp(this.focusCol, 0, sheet.colCount - 1);
  }
}

function clamp(v: number, lo: number, hi: number): number {
  if (lo > hi) {
    return lo;
  }
  return Math.max(lo, Math.min(hi, v));
}

import {
  Cell,
  normalizeSelectionRange,
  type ICommand,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";

import { recalcWorksheet, type CellContentSnapshot } from "../recalc.js";

function cellHasClearableContent(cell: Cell): boolean {
  if (cell.isFormulaCell()) {
    return true;
  }
  return cell.value !== null;
}

interface RegionEntry {
  readonly row: number;
  readonly col: number;
  readonly before: CellContentSnapshot;
}

/**
 * 清空矩形选区内单元格的公式/值，保留 `Cell.style` 与表结构；一次 `execute` / `undo` 各触发一次整表重算。
 */
export class ClearRegionContentsCommand implements ICommand {
  readonly id = "cell.clearRegionContents";
  readonly label = "清除内容";

  private readonly sheet: Worksheet;
  private readonly entries: readonly RegionEntry[];

  constructor(sheet: Worksheet, range: SelectionRange) {
    this.sheet = sheet;
    const n = normalizeSelectionRange(range);
    const list: RegionEntry[] = [];
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        const cell = sheet.getCell(r, c);
        if (!cellHasClearableContent(cell)) {
          continue;
        }
        list.push({
          row: r,
          col: c,
          before: { formula: cell.formula, value: cell.value },
        });
      }
    }
    this.entries = list;
  }

  /** 选区内无可清空内容时不应 `execute` 入栈。 */
  get hasChanges(): boolean {
    return this.entries.length > 0;
  }

  execute(): void {
    if (this.entries.length === 0) {
      return;
    }
    const sheet = this.sheet;
    sheet.batch(() => {
      for (const e of this.entries) {
        sheet.setCellLiteral(e.row, e.col, null);
      }
    });
    recalcWorksheet(sheet);
  }

  undo(): void {
    if (this.entries.length === 0) {
      return;
    }
    const sheet = this.sheet;
    sheet.batch(() => {
      for (const e of this.entries) {
        const { row, col, before } = e;
        if (before.formula !== null) {
          sheet.setCellFormula(row, col, before.formula);
        } else {
          sheet.setCellLiteral(row, col, before.value);
        }
      }
    });
    recalcWorksheet(sheet);
  }
}

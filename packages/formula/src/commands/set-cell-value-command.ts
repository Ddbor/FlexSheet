import type { CellScalar, ICommand } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import { columnIndexToLabel } from "@flexsheet/shared";

import { applyCellSnapshotAndRecalc, setCellValueAndRecalc, type CellContentSnapshot } from "../recalc.js";

/**
 * 将一次单元格编辑封装为可逆命令（execute = 新值 + 重算，undo = 恢复编辑前快照）。
 */
export class SetCellValueCommand implements ICommand {
  readonly id = "cell.setValue";
  readonly label: string;

  private readonly before: CellContentSnapshot;
  private readonly sheet: Worksheet;
  private readonly row: number;
  private readonly col: number;
  private readonly newValue: CellScalar;

  constructor(sheet: Worksheet, row: number, col: number, newValue: CellScalar, label?: string) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.newValue = newValue;
    const cell = sheet.getCell(row, col);
    this.before = { formula: cell.formula, value: cell.value };
    this.label =
      label ?? `编辑 ${columnIndexToLabel(col)}${row + 1}`;
  }

  execute(): void {
    setCellValueAndRecalc(this.sheet, this.row, this.col, this.newValue);
  }

  undo(): void {
    applyCellSnapshotAndRecalc(this.sheet, this.row, this.col, this.before);
  }
}

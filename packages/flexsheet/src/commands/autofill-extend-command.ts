import type { CellScalar, CellStyle, ICommand, SelectionRange, Worksheet } from "@flexsheet/core";
import { normalizeSelectionRange } from "@flexsheet/core";
import { parseEditString } from "@flexsheet/editor";
import { recalcWorksheet } from "@flexsheet/formula";

import { serializeSelection } from "../clipboard/serialize-selection.js";

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

/**
 * 从选区右下角填充柄扩展：将 `fillRange` 相对 `sourceRange` 多出的单元格按源区内容平铺写入（含样式），一次重算。
 */
export class AutofillExtendCommand implements ICommand {
  readonly id = "sheet.autofillExtend";
  readonly label = "填充";

  private readonly sheet: Worksheet;
  private readonly sourceRange: SelectionRange;
  private readonly fillRange: SelectionRange;
  private readonly textRows: readonly (readonly string[])[];
  private readonly styles: readonly (readonly (CellStyle | null)[])[];
  private readonly before: CellFullSnapshot[];

  constructor(sheet: Worksheet, sourceRange: SelectionRange, fillRange: SelectionRange) {
    this.sheet = sheet;
    this.sourceRange = normalizeSelectionRange(sourceRange);
    this.fillRange = normalizeSelectionRange(fillRange);
    const { textRows, styles } = serializeSelection(sheet, this.sourceRange);
    this.textRows = textRows;
    this.styles = styles;

    const S = this.sourceRange;
    const F = this.fillRange;
    const before: CellFullSnapshot[] = [];
    for (let r = F.startRow; r <= F.endRow; r++) {
      for (let c = F.startCol; c <= F.endCol; c++) {
        if (r >= S.startRow && r <= S.endRow && c >= S.startCol && c <= S.endCol) {
          continue;
        }
        if (sheet.isMergeCoveredCell(r, c)) {
          continue;
        }
        before.push(takeSnapshot(sheet, r, c));
      }
    }
    this.before = before;
  }

  execute(): void {
    const { sheet, sourceRange: S, fillRange: F, textRows, styles } = this;
    const sh = S.endRow - S.startRow + 1;
    const sw = S.endCol - S.startCol + 1;
    if (sh <= 0 || sw <= 0 || textRows.length !== sh || textRows[0]!.length !== sw) {
      return;
    }
    sheet.batch(() => {
      for (let r = F.startRow; r <= F.endRow; r++) {
        for (let c = F.startCol; c <= F.endCol; c++) {
          if (r >= S.startRow && r <= S.endRow && c >= S.startCol && c <= S.endCol) {
            continue;
          }
          if (sheet.isMergeCoveredCell(r, c)) {
            continue;
          }
          const pr = ((r - S.startRow) % sh + sh) % sh;
          const pc = ((c - S.startCol) % sw + sw) % sw;
          const raw = textRows[pr]![pc] ?? "";
          sheet.setCellValue(r, c, parseEditString(raw));
          const st = styles[pr]![pc] ?? null;
          sheet.setCellStyle(r, c, st === null ? null : { ...st });
        }
      }
    });
    recalcWorksheet(sheet);
  }

  undo(): void {
    let k = 0;
    const { sheet, fillRange: F, sourceRange: S } = this;
    sheet.batch(() => {
      for (let r = F.startRow; r <= F.endRow; r++) {
        for (let c = F.startCol; c <= F.endCol; c++) {
          if (r >= S.startRow && r <= S.endRow && c >= S.startCol && c <= S.endCol) {
            continue;
          }
          if (sheet.isMergeCoveredCell(r, c)) {
            continue;
          }
          applySnapshot(sheet, this.before[k]!);
          k++;
        }
      }
    });
    recalcWorksheet(sheet);
  }
}

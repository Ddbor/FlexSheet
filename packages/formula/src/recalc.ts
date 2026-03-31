import type { CellScalar } from "@flexsheet/core";
import { Cell } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";

/** 单元格在「一次编辑前」的内容快照（用于撤销）。 */
export interface CellContentSnapshot {
  readonly formula: string | null;
  readonly value: CellScalar;
}

/** 将快照写回并整表重算（与 `setCellValueAndRecalc` 成对用于撤销）。 */
export function applyCellSnapshotAndRecalc(
  sheet: Worksheet,
  row: number,
  col: number,
  snap: CellContentSnapshot,
): void {
  applyWriteAndRecalc(sheet, (s) => {
    if (snap.formula !== null) {
      s.setCellFormula(row, col, snap.formula);
    } else {
      s.setCellLiteral(row, col, snap.value);
    }
  });
}
import { evaluateAst } from "./evaluator.js";
import type { EvalContext } from "./evaluator.js";
import { ParseError, parseFormula } from "./parser.js";

function applyWriteAndRecalc(sheet: Worksheet, write: (s: Worksheet) => void): void {
  sheet.batch(() => {
    write(sheet);
    runRecalcWithoutNotify(sheet);
  });
}

/** 写入单元格（字面量或公式）后整表重算；合并为一次数据变更通知。 */
export function setCellValueAndRecalc(
  sheet: Worksheet,
  row: number,
  col: number,
  value: CellScalar,
): void {
  applyWriteAndRecalc(sheet, (s) => {
    s.setCellValue(row, col, value);
  });
}

/**
 * 仅写入字面量（不解析 `=` 为公式）后整表重算。
 * 用于已知依赖公式的单元格改数值、且不想走 `setCellValue` 字符串分支的场景。
 */
export function setCellLiteralAndRecalc(
  sheet: Worksheet,
  row: number,
  col: number,
  value: CellScalar,
): void {
  applyWriteAndRecalc(sheet, (s) => {
    s.setCellLiteral(row, col, value);
  });
}

/**
 * 对整张表中含公式的单元格求值（递归引用 + 环检测）。
 * 重算过程中直接写入 `Cell.value`，结束时统一 `notifyDataChanged`。
 */
export function recalcWorksheet(sheet: Worksheet): void {
  runRecalcWithoutNotify(sheet);
  sheet.notifyDataChanged();
}

function runRecalcWithoutNotify(sheet: Worksheet): void {
  const jobs: Array<{ row: number; col: number }> = [];
  sheet.iterateCells((c) => {
    if (c.formula !== null) {
      jobs.push({ row: c.row, col: c.col });
    }
  });
  const visiting = new Set<string>();
  for (const { row, col } of jobs) {
    try {
      evaluateFormulaCell(sheet, row, col, visiting);
    } catch {
      sheet.getCell(row, col).value = null;
    }
  }
}

function evaluateFormulaCell(
  sheet: Worksheet,
  row: number,
  col: number,
  visiting: Set<string>,
): CellScalar {
  const key = Cell.key(row, col);
  const cell = sheet.getCell(row, col);
  if (cell.formula === null) {
    return cell.value;
  }
  if (visiting.has(key)) {
    return null;
  }
  visiting.add(key);
  try {
    let ast;
    try {
      ast = parseFormula(cell.formula);
    } catch (e) {
      if (e instanceof ParseError) {
        cell.value = null;
        return null;
      }
      throw e;
    }
    const ctx: EvalContext = {
      getScalar: (r, c) => {
        if (r < 0 || c < 0 || r >= sheet.rowCount || c >= sheet.colCount) {
          return null;
        }
        const dep = sheet.getCell(r, c);
        if (dep.formula !== null) {
          return evaluateFormulaCell(sheet, r, c, visiting);
        }
        return dep.value;
      },
    };
    const v = evaluateAst(ast, ctx);
    cell.value = v;
    return v;
  } finally {
    visiting.delete(key);
  }
}

export type { AstBinary, AstCall, AstNode, AstNumber, AstRange, AstRef, AstUnary } from "./ast.js";

export { evaluateAst, toNumber, type EvalContext } from "./evaluator.js";

export { ParseError, parseFormula, stripFormulaEquals } from "./parser.js";

export {
  applyCellSnapshotAndRecalc,
  evaluateFormulaExpressionOnSheet,
  recalcWorksheet,
  setCellLiteralAndRecalc,
  setCellValueAndRecalc,
  type CellContentSnapshot,
} from "./recalc.js";

export { SetCellValueCommand } from "./commands/set-cell-value-command.js";
export { ClearRegionContentsCommand } from "./commands/clear-region-contents-command.js";

export { FormulaEnginePlugin } from "./formula-plugin.js";

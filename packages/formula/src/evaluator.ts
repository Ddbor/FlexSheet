import type { CellScalar } from "@flexsheet/core";
import type { AstNode } from "./ast.js";

export interface EvalContext {
  getScalar(row: number, col: number): CellScalar;
}

export function evaluateAst(node: AstNode, ctx: EvalContext): CellScalar {
  switch (node.type) {
    case "number":
      return node.value;
    case "ref":
      return ctx.getScalar(node.row, node.col);
    case "range":
      return null;
    case "unary": {
      if (node.op !== "-") {
        return null;
      }
      const v = evaluateAst(node.expr, ctx);
      return -toNumber(v);
    }
    case "binary": {
      const a = toNumber(evaluateAst(node.left, ctx));
      const b = toNumber(evaluateAst(node.right, ctx));
      switch (node.op) {
        case "+":
          return a + b;
        case "-":
          return a - b;
        case "*":
          return a * b;
        case "/":
          if (b === 0) {
            return null;
          }
          return a / b;
        default:
          return null;
      }
    }
    case "call": {
      const name = node.name.toUpperCase();
      if (name === "SUM") {
        let sum = 0;
        for (const arg of node.args) {
          sum += evalSumArgument(arg, ctx);
        }
        return sum;
      }
      if (name === "AVERAGE" || name === "AVG") {
        const nums: number[] = [];
        for (const arg of node.args) {
          forEachNumericInCallArgument(arg, ctx, (n) => {
            nums.push(n);
          });
        }
        if (nums.length === 0) {
          return null;
        }
        return nums.reduce((a, b) => a + b, 0) / nums.length;
      }
      if (name === "COUNT") {
        let n = 0;
        for (const arg of node.args) {
          forEachNumericInCallArgument(arg, ctx, () => {
            n += 1;
          });
        }
        return n;
      }
      if (name === "MAX") {
        let m: number | null = null;
        for (const arg of node.args) {
          forEachNumericInCallArgument(arg, ctx, (v) => {
            m = m === null ? v : Math.max(m, v);
          });
        }
        return m;
      }
      if (name === "MIN") {
        let m: number | null = null;
        for (const arg of node.args) {
          forEachNumericInCallArgument(arg, ctx, (v) => {
            m = m === null ? v : Math.min(m, v);
          });
        }
        return m;
      }
      return null;
    }
    default:
      return null;
  }
}

function evalSumArgument(node: AstNode, ctx: EvalContext): number {
  if (node.type === "range") {
    let s = 0;
    for (let r = node.startRow; r <= node.endRow; r++) {
      for (let c = node.startCol; c <= node.endCol; c++) {
        s += toNumber(ctx.getScalar(r, c));
      }
    }
    return s;
  }
  return toNumber(evaluateAst(node, ctx));
}

/** 与 Excel 接近：仅对可解析为数字的标量计值（`COUNT` / `AVERAGE` / `MAX` / `MIN`）。 */
function tryNumberFromScalar(v: CellScalar): number | null {
  if (v === null || v === "") {
    return null;
  }
  if (typeof v === "number") {
    return Number.isFinite(v) ? v : null;
  }
  if (typeof v === "boolean") {
    return v ? 1 : 0;
  }
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function forEachNumericInCallArgument(
  node: AstNode,
  ctx: EvalContext,
  fn: (n: number) => void,
): void {
  if (node.type === "range") {
    for (let r = node.startRow; r <= node.endRow; r++) {
      for (let c = node.startCol; c <= node.endCol; c++) {
        const t = tryNumberFromScalar(ctx.getScalar(r, c));
        if (t !== null) {
          fn(t);
        }
      }
    }
    return;
  }
  const t = tryNumberFromScalar(evaluateAst(node, ctx));
  if (t !== null) {
    fn(t);
  }
}

export function toNumber(v: CellScalar): number {
  if (v === null || v === "") {
    return 0;
  }
  if (typeof v === "number") {
    return Number.isFinite(v) ? v : 0;
  }
  if (typeof v === "boolean") {
    return v ? 1 : 0;
  }
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

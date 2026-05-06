import type { AstCall, AstNode, AstRange, AstRef } from "./ast.js";

export class ParseError extends Error {
  constructor(
    message: string,
    public readonly offset: number,
  ) {
    super(message);
    this.name = "ParseError";
  }
}

type Token =
  | { kind: "eof"; pos: number }
  | { kind: "num"; value: number; pos: number }
  | { kind: "ref"; row: number; col: number; pos: number }
  | { kind: "ident"; name: string; pos: number }
  | { kind: "op"; op: "+" | "-" | "*" | "/"; pos: number }
  | { kind: "lparen"; pos: number }
  | { kind: "rparen"; pos: number }
  | { kind: "comma"; pos: number }
  | { kind: "colon"; pos: number };

/** 去掉前导 `=` 与首尾空白，供词法使用。 */
export function stripFormulaEquals(raw: string): string {
  const t = raw.trim();
  if (t.startsWith("=")) {
    return t.slice(1).trimStart();
  }
  return t;
}

export function parseFormula(formula: string): AstNode {
  const inner = stripFormulaEquals(formula);
  if (inner === "") {
    throw new ParseError("空公式", 0);
  }
  const tokens = tokenize(inner);
  const p = new Parser(tokens, inner);
  return p.parseDocument();
}

function tokenize(s: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const n = s.length;

  const skipWs = (): void => {
    while (i < n && (s[i] === " " || s[i] === "\t" || s[i] === "\n" || s[i] === "\r")) {
      i++;
    }
  };

  while (i < n) {
    skipWs();
    if (i >= n) {
      break;
    }
    const pos = i;
    const ch = s[i];

    if (ch === "(") {
      i++;
      out.push({ kind: "lparen", pos });
      continue;
    }
    if (ch === ")") {
      i++;
      out.push({ kind: "rparen", pos });
      continue;
    }
    if (ch === ",") {
      i++;
      out.push({ kind: "comma", pos });
      continue;
    }
    if (ch === ":") {
      i++;
      out.push({ kind: "colon", pos });
      continue;
    }
    if (ch === "+" || ch === "*" || ch === "/") {
      i++;
      out.push({ kind: "op", op: ch, pos });
      continue;
    }
    if (ch === "-") {
      i++;
      out.push({ kind: "op", op: "-", pos });
      continue;
    }

    if ((ch >= "0" && ch <= "9") || ch === ".") {
      const start = i;
      if (ch === "-") {
        /* handled above */
      }
      while (i < n && ((s[i] >= "0" && s[i] <= "9") || s[i] === ".")) {
        i++;
      }
      const slice = s.slice(start, i);
      const value = Number(slice);
      if (Number.isNaN(value)) {
        throw new ParseError(`无效数字「${slice}」`, start);
      }
      out.push({ kind: "num", value, pos: start });
      continue;
    }

    /** A1 / $A$1 / A$1 / $A1 等 Excel 引用（`$` 仅表示绝对，解析为相同行列下标）。 */
    if (ch === "$" || (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z")) {
      const start = i;
      if (s[i] === "$") {
        i++;
        if (i >= n || !((s[i] >= "A" && s[i] <= "Z") || (s[i] >= "a" && s[i] <= "z"))) {
          throw new ParseError(`「$」后应为列字母`, pos);
        }
      }
      const colStart = i;
      while (i < n && ((s[i] >= "A" && s[i] <= "Z") || (s[i] >= "a" && s[i] <= "z"))) {
        i++;
      }
      const letters = s.slice(colStart, i);
      if (i < n && s[i] === "$") {
        i++;
      }
      if (i < n && s[i] >= "0" && s[i] <= "9") {
        const d0 = i;
        while (i < n && s[i] >= "0" && s[i] <= "9") {
          i++;
        }
        const row1 = Number(s.slice(d0, i));
        if (!Number.isInteger(row1) || row1 < 1) {
          throw new ParseError(`无效行号「${s.slice(d0, i)}」`, d0);
        }
        const col = colLettersToIndex(letters, colStart);
        const row = row1 - 1;
        out.push({ kind: "ref", row, col, pos: start });
        continue;
      }
      const name = letters.toUpperCase();
      out.push({ kind: "ident", name, pos: start });
      continue;
    }

    throw new ParseError(`无法识别的字符「${ch}」`, pos);
  }

  out.push({ kind: "eof", pos: n });
  return out;
}

function colLettersToIndex(letters: string, pos: number): number {
  if (letters.length === 0) {
    throw new ParseError("缺少列字母", pos);
  }
  const u = letters.toUpperCase();
  let v = 0;
  for (let k = 0; k < u.length; k++) {
    const c = u.charCodeAt(k);
    if (c < 65 || c > 90) {
      throw new ParseError(`无效列「${letters}」`, pos);
    }
    v = v * 26 + (c - 64);
  }
  return v - 1;
}

class Parser {
  private readonly tokens: Token[];
  private readonly source: string;
  private index = 0;

  constructor(tokens: Token[], source: string) {
    this.tokens = tokens;
    this.source = source;
  }

  private peek(): Token {
    return this.tokens[this.index] ?? { kind: "eof", pos: this.source.length };
  }

  private advance(): Token {
    const t = this.peek();
    if (t.kind !== "eof") {
      this.index++;
    }
    return t;
  }

  parseDocument(): AstNode {
    const ast = this.parseExpression();
    this.expectEof();
    return ast;
  }

  parseExpression(): AstNode {
    return this.parseAddSub();
  }

  private parseAddSub(): AstNode {
    let left = this.parseMulDiv();
    for (;;) {
      const t = this.peek();
      if (t.kind === "op" && (t.op === "+" || t.op === "-")) {
        this.advance();
        const right = this.parseMulDiv();
        left = { type: "binary", op: t.op, left, right };
        continue;
      }
      break;
    }
    return left;
  }

  private parseMulDiv(): AstNode {
    let left = this.parseUnary();
    for (;;) {
      const t = this.peek();
      if (t.kind === "op" && (t.op === "*" || t.op === "/")) {
        this.advance();
        const right = this.parseUnary();
        left = { type: "binary", op: t.op, left, right };
        continue;
      }
      break;
    }
    return left;
  }

  private parseUnary(): AstNode {
    const t = this.peek();
    if (t.kind === "op" && t.op === "-") {
      this.advance();
      const expr = this.parseUnary();
      return { type: "unary", op: "-", expr };
    }
    return this.parsePrimary();
  }

  private parsePrimary(): AstNode {
    const t = this.peek();
    if (t.kind === "num") {
      this.advance();
      return { type: "number", value: t.value };
    }
    if (t.kind === "ref") {
      this.advance();
      const first: AstRef = { type: "ref", row: t.row, col: t.col };
      const c = this.peek();
      if (c.kind === "colon") {
        this.advance();
        const t2 = this.peek();
        if (t2.kind !== "ref") {
          throw new ParseError("区域引用应为「A1:B2」形式", t2.pos);
        }
        this.advance();
        const sr = Math.min(first.row, t2.row);
        const er = Math.max(first.row, t2.row);
        const sc = Math.min(first.col, t2.col);
        const ec = Math.max(first.col, t2.col);
        const range: AstRange = {
          type: "range",
          startRow: sr,
          startCol: sc,
          endRow: er,
          endCol: ec,
        };
        return range;
      }
      return first;
    }
    if (t.kind === "ident") {
      this.advance();
      const lp = this.peek();
      if (lp.kind !== "lparen") {
        throw new ParseError(`期望「(」以调用函数`, lp.pos);
      }
      this.advance();
      const args = this.parseArgumentList();
      const rp = this.peek();
      if (rp.kind !== "rparen") {
        throw new ParseError("缺少「)」", rp.pos);
      }
      this.advance();
      const call: AstCall = { type: "call", name: t.name, args };
      return call;
    }
    if (t.kind === "lparen") {
      this.advance();
      const inner = this.parseExpression();
      const rp = this.peek();
      if (rp.kind !== "rparen") {
        throw new ParseError("缺少「)」", rp.pos);
      }
      this.advance();
      return inner;
    }
    throw new ParseError("意外的记号", t.pos);
  }

  private parseArgumentList(): AstNode[] {
    const args: AstNode[] = [];
    if (this.peek().kind === "rparen") {
      return args;
    }
    args.push(this.parseExpression());
    while (this.peek().kind === "comma") {
      this.advance();
      args.push(this.parseExpression());
    }
    return args;
  }

  private expectEof(): void {
    const t = this.peek();
    if (t.kind !== "eof") {
      throw new ParseError("公式末尾有多余内容", t.pos);
    }
  }
}

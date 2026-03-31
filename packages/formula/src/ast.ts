/**
 * 公式 AST（Parser 产出，Calculator 消费）。
 */

export interface AstNumber {
  readonly type: "number";
  readonly value: number;
}

export interface AstRef {
  readonly type: "ref";
  readonly row: number;
  readonly col: number;
}

export interface AstRange {
  readonly type: "range";
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

export interface AstBinary {
  readonly type: "binary";
  readonly op: "+" | "-" | "*" | "/";
  readonly left: AstNode;
  readonly right: AstNode;
}

export interface AstUnary {
  readonly type: "unary";
  readonly op: "-";
  readonly expr: AstNode;
}

export interface AstCall {
  readonly type: "call";
  readonly name: string;
  readonly args: readonly AstNode[];
}

export type AstNode = AstNumber | AstRef | AstRange | AstBinary | AstUnary | AstCall;

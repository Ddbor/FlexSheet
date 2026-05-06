import {
  Cell,
  normalizeSelectionRange,
  selectionRangeContains,
  type CellScalar,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";
import type { AstNode } from "@flexsheet/formula";
import { ParseError, parseFormula } from "@flexsheet/formula";

export type GotoSpecialKind =
  | "comments"
  | "constants"
  | "formulas"
  | "blanks"
  | "currentRegion"
  | "currentArray"
  | "rowDiff"
  | "colDiff"
  | "precedents"
  | "dependents"
  | "lastCell"
  | "visibleOnly"
  | "objects"
  | "conditionalFormats"
  | "dataValidation";

export interface GotoSpecialSubtypeFilters {
  readonly numbers: boolean;
  readonly text: boolean;
  readonly logicals: boolean;
  readonly errors: boolean;
}

export interface GotoSpecialComputeInput {
  readonly sheet: Worksheet;
  readonly selectionRange: SelectionRange;
  readonly activeRow: number;
  readonly activeCol: number;
  readonly kind: GotoSpecialKind;
  /** 常量 / 公式子类型；缺省视为全开。 */
  readonly subtypeFilters?: GotoSpecialSubtypeFilters;
  /** 引用 / 从属：true 表示仅直接一级。 */
  readonly directLinksOnly?: boolean;
}

export type GotoSpecialComputeResult =
  | { readonly ok: true; readonly range: SelectionRange }
  | { readonly ok: false; readonly code: "none" | "unsupported" | "singleCell" };

function defaultFilters(): GotoSpecialSubtypeFilters {
  return { numbers: true, text: true, logicals: true, errors: true };
}

function isBlankScalar(v: CellScalar): boolean {
  return v === null || v === "";
}

function isErrorLikeValue(v: CellScalar): boolean {
  if (typeof v !== "string") {
    return false;
  }
  const t = v.trim();
  return /^#[A-Z]{3,5}[!/]?$/i.test(t) || t.includes("#N/A");
}

function matchesSubtype(v: CellScalar, f: GotoSpecialSubtypeFilters): boolean {
  if (isBlankScalar(v)) {
    return false;
  }
  if (typeof v === "number") {
    return f.numbers;
  }
  if (typeof v === "boolean") {
    return f.logicals;
  }
  if (typeof v === "string") {
    if (isErrorLikeValue(v)) {
      return f.errors;
    }
    return f.text;
  }
  return false;
}

function cellHasDataSparse(sheet: Worksheet, row: number, col: number): boolean {
  if (row < 0 || col < 0 || row >= sheet.rowCount || col >= sheet.colCount) {
    return false;
  }
  const a = sheet.getMergeAnchorCell(row, col);
  if (!sheet.hasCell(a.row, a.col)) {
    return false;
  }
  const cell = sheet.getCell(a.row, a.col);
  if (cell.formula !== null && cell.formula.length > 0) {
    return true;
  }
  if (!isBlankScalar(cell.value)) {
    return true;
  }
  return false;
}

function readAnchorContent(
  sheet: Worksheet,
  row: number,
  col: number,
): { readonly formula: string | null; readonly value: CellScalar } {
  const a = sheet.getMergeAnchorCell(row, col);
  if (!sheet.hasCell(a.row, a.col)) {
    return { formula: null, value: null };
  }
  const cell = sheet.getCell(a.row, a.col);
  return { formula: cell.formula, value: cell.value };
}

function scalarEqual(a: CellScalar, b: CellScalar): boolean {
  return a === b;
}

function bumpRangeBounds(
  r: number,
  c: number,
  bounds: { minR: number; maxR: number; minC: number; maxC: number },
): void {
  bounds.minR = Math.min(bounds.minR, r);
  bounds.maxR = Math.max(bounds.maxR, r);
  bounds.minC = Math.min(bounds.minC, c);
  bounds.maxC = Math.max(bounds.maxC, c);
}

function selectionFromKeys(sheet: Worksheet, keys: Set<string>): SelectionRange | null {
  if (keys.size === 0) {
    return null;
  }
  let minR = sheet.rowCount;
  let maxR = -1;
  let minC = sheet.colCount;
  let maxC = -1;
  for (const k of keys) {
    const [rs, cs] = k.split(",");
    const r = Number(rs);
    const c = Number(cs);
    if (!Number.isInteger(r) || !Number.isInteger(c)) {
      continue;
    }
    minR = Math.min(minR, r);
    maxR = Math.max(maxR, r);
    minC = Math.min(minC, c);
    maxC = Math.max(maxC, c);
  }
  if (maxR < 0) {
    return null;
  }
  return normalizeSelectionRange({ startRow: minR, endRow: maxR, startCol: minC, endCol: maxC });
}

function computeCurrentRegionRange(sheet: Worksheet, ar: number, ac: number): SelectionRange {
  if (!cellHasDataSparse(sheet, ar, ac)) {
    return normalizeSelectionRange({ startRow: ar, endRow: ar, startCol: ac, endCol: ac });
  }
  const visited = new Set<string>();
  const q: Array<{ readonly r: number; readonly c: number }> = [];
  const seed = sheet.getMergeAnchorCell(ar, ac);
  q.push({ r: seed.row, c: seed.col });
  visited.add(Cell.key(seed.row, seed.col));
  let minR = seed.row;
  let maxR = seed.row;
  let minC = seed.col;
  let maxC = seed.col;
  const dirs: ReadonlyArray<readonly [number, number]> = [
    [1, 0],
    [-1, 0],
    [0, 1],
    [0, -1],
  ];
  for (let qi = 0; qi < q.length; qi++) {
    const { r, c } = q[qi]!;
    minR = Math.min(minR, r);
    maxR = Math.max(maxR, r);
    minC = Math.min(minC, c);
    maxC = Math.max(maxC, c);
    for (const [dr, dc] of dirs) {
      const nr = r + dr;
      const nc = c + dc;
      if (nr < 0 || nc < 0 || nr >= sheet.rowCount || nc >= sheet.colCount) {
        continue;
      }
      if (!cellHasDataSparse(sheet, nr, nc)) {
        continue;
      }
      const a = sheet.getMergeAnchorCell(nr, nc);
      const k = Cell.key(a.row, a.col);
      if (visited.has(k)) {
        continue;
      }
      visited.add(k);
      q.push({ r: a.row, c: a.col });
    }
  }
  return normalizeSelectionRange({ startRow: minR, endRow: maxR, startCol: minC, endCol: maxC });
}

function computeLastUsedCellRange(sheet: Worksheet): SelectionRange | null {
  const bounds = { minR: sheet.rowCount, maxR: -1, minC: sheet.colCount, maxC: -1 };
  sheet.iterateCells((cell) => {
    const has =
      (cell.formula !== null && cell.formula.length > 0) ||
      !isBlankScalar(cell.value) ||
      cell.style !== null;
    if (!has) {
      return;
    }
    bumpRangeBounds(cell.row, cell.col, bounds);
  });
  for (const { masterRow, masterCol, rowSpan, colSpan } of sheet.getMergeRegionsSnapshot()) {
    bumpRangeBounds(masterRow, masterCol, bounds);
    bumpRangeBounds(masterRow + rowSpan - 1, masterCol + colSpan - 1, bounds);
  }
  if (bounds.maxR < 0) {
    return null;
  }
  return normalizeSelectionRange({
    startRow: bounds.minR,
    endRow: bounds.maxR,
    startCol: bounds.minC,
    endCol: bounds.maxC,
  });
}

function astTouchesCell(ast: AstNode, row: number, col: number): boolean {
  switch (ast.type) {
    case "ref":
      return ast.row === row && ast.col === col;
    case "range": {
      const r0 = Math.min(ast.startRow, ast.endRow);
      const r1 = Math.max(ast.startRow, ast.endRow);
      const c0 = Math.min(ast.startCol, ast.endCol);
      const c1 = Math.max(ast.startCol, ast.endCol);
      return row >= r0 && row <= r1 && col >= c0 && col <= c1;
    }
    case "binary":
      return astTouchesCell(ast.left, row, col) || astTouchesCell(ast.right, row, col);
    case "unary":
      return astTouchesCell(ast.expr, row, col);
    case "call":
      return ast.args.some((a) => astTouchesCell(a, row, col));
    case "number":
      return false;
    default:
      return false;
  }
}

function tryParseFormulaAst(formula: string): AstNode | null {
  try {
    return parseFormula(formula);
  } catch (e) {
    if (e instanceof ParseError) {
      return null;
    }
    throw e;
  }
}

const MAX_RANGE_EXPAND = 8000;

function collectDirectPrecedentKeys(sheet: Worksheet, row: number, col: number): Set<string> {
  const out = new Set<string>();
  const a = sheet.getMergeAnchorCell(row, col);
  const { formula } = readAnchorContent(sheet, a.row, a.col);
  if (formula === null || formula.length === 0) {
    return out;
  }
  const ast = tryParseFormulaAst(formula);
  if (ast === null) {
    return out;
  }
  const walk = (n: AstNode): void => {
    switch (n.type) {
      case "ref":
        out.add(Cell.key(n.row, n.col));
        return;
      case "range": {
        const r0 = Math.min(n.startRow, n.endRow);
        const r1 = Math.max(n.startRow, n.endRow);
        const c0 = Math.min(n.startCol, n.endCol);
        const c1 = Math.max(n.startCol, n.endCol);
        const area = (r1 - r0 + 1) * (c1 - c0 + 1);
        if (area <= MAX_RANGE_EXPAND) {
          for (let r = r0; r <= r1; r++) {
            for (let c = c0; c <= c1; c++) {
              out.add(Cell.key(r, c));
            }
          }
        } else {
          out.add(Cell.key(r0, c0));
          out.add(Cell.key(r1, c1));
        }
        return;
      }
      case "binary":
        walk(n.left);
        walk(n.right);
        return;
      case "unary":
        walk(n.expr);
        return;
      case "call":
        for (const x of n.args) {
          walk(x);
        }
        return;
      default:
        return;
    }
  };
  walk(ast);
  return out;
}

function collectAllPrecedentKeys(sheet: Worksheet, startRow: number, startCol: number): Set<string> {
  const seen = new Set<string>();
  const queue: Array<{ readonly r: number; readonly c: number }> = [];
  const seed = sheet.getMergeAnchorCell(startRow, startCol);
  const first = collectDirectPrecedentKeys(sheet, seed.row, seed.col);
  for (const k of first) {
    if (!seen.has(k)) {
      seen.add(k);
      const [rs, cs] = k.split(",");
      queue.push({ r: Number(rs), c: Number(cs) });
    }
  }
  const MAX_NODES = 4000;
  for (let qi = 0; qi < queue.length; qi++) {
    if (seen.size > MAX_NODES) {
      break;
    }
    const { r, c } = queue[qi]!;
    const more = collectDirectPrecedentKeys(sheet, r, c);
    for (const k of more) {
      if (seen.has(k)) {
        continue;
      }
      seen.add(k);
      const [rs, cs] = k.split(",");
      queue.push({ r: Number(rs), c: Number(cs) });
    }
  }
  return seen;
}

function astTouchesAnyKey(ast: AstNode, keys: ReadonlySet<string>): boolean {
  for (const k of keys) {
    const p = k.indexOf(",");
    if (p <= 0) {
      continue;
    }
    const r = Number(k.slice(0, p));
    const c = Number(k.slice(p + 1));
    if (!Number.isInteger(r) || !Number.isInteger(c)) {
      continue;
    }
    if (astTouchesCell(ast, r, c)) {
      return true;
    }
  }
  return false;
}

function collectDependentsKeys(
  sheet: Worksheet,
  targetRow: number,
  targetCol: number,
  directOnly: boolean,
): Set<string> {
  const target = sheet.getMergeAnchorCell(targetRow, targetCol);
  const tKey = Cell.key(target.row, target.col);
  const result = new Set<string>();

  const formulaCells: Array<{ readonly r: number; readonly c: number; readonly ast: AstNode }> = [];
  sheet.iterateCells((cell) => {
    if (cell.formula === null || cell.formula.length === 0) {
      return;
    }
    const a = sheet.getMergeAnchorCell(cell.row, cell.col);
    if (a.row !== cell.row || a.col !== cell.col) {
      return;
    }
    const ast = tryParseFormulaAst(cell.formula);
    if (ast === null) {
      return;
    }
    formulaCells.push({ r: cell.row, c: cell.col, ast });
  });

  if (directOnly) {
    for (const it of formulaCells) {
      if (astTouchesCell(it.ast, target.row, target.col)) {
        result.add(Cell.key(it.r, it.c));
      }
    }
    return result;
  }

  const basisKeys = new Set<string>([tKey]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const it of formulaCells) {
      const k = Cell.key(it.r, it.c);
      if (result.has(k)) {
        continue;
      }
      if (astTouchesAnyKey(it.ast, basisKeys)) {
        result.add(k);
        basisKeys.add(k);
        changed = true;
      }
    }
  }
  return result;
}

function forEachDistinctMergeAnchorInSelection(
  sheet: Worksheet,
  range: SelectionRange,
  fn: (anchorRow: number, anchorCol: number) => void,
): void {
  const n = normalizeSelectionRange(range);
  const seen = new Set<string>();
  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      const a = sheet.getMergeAnchorCell(r, c);
      const k = Cell.key(a.row, a.col);
      if (seen.has(k)) {
        continue;
      }
      seen.add(k);
      fn(a.row, a.col);
    }
  }
}

function cellHasConditionalFormat(sheet: Worksheet, row: number, col: number): boolean {
  for (const rule of sheet.getConditionalFormatRules()) {
    if (selectionRangeContains(rule.range, row, col)) {
      return true;
    }
  }
  return false;
}

export function computeGotoSpecialRange(input: GotoSpecialComputeInput): GotoSpecialComputeResult {
  const { sheet, kind } = input;
  const sel = normalizeSelectionRange(input.selectionRange);
  const ar = Math.max(0, Math.min(sheet.rowCount - 1, input.activeRow));
  const ac = Math.max(0, Math.min(sheet.colCount - 1, input.activeCol));
  const filters = input.subtypeFilters ?? defaultFilters();
  const directOnly = input.directLinksOnly === true;

  switch (kind) {
    case "comments":
    case "objects":
    case "dataValidation":
      return { ok: false, code: "unsupported" };
    case "currentRegion": {
      const range = computeCurrentRegionRange(sheet, ar, ac);
      return { ok: true, range };
    }
    case "currentArray": {
      const info = sheet.getMergedRectInfo(ar, ac);
      if (info.rowSpan > 1 || info.colSpan > 1) {
        return {
          ok: true,
          range: normalizeSelectionRange({
            startRow: info.anchorRow,
            endRow: info.anchorRow + info.rowSpan - 1,
            startCol: info.anchorCol,
            endCol: info.anchorCol + info.colSpan - 1,
          }),
        };
      }
      return { ok: false, code: "none" };
    }
    case "lastCell": {
      const r = computeLastUsedCellRange(sheet);
      if (r === null) {
        return { ok: false, code: "none" };
      }
      return { ok: true, range: normalizeSelectionRange({ startRow: r.endRow, endRow: r.endRow, startCol: r.endCol, endCol: r.endCol }) };
    }
    case "precedents": {
      const keys = directOnly
        ? collectDirectPrecedentKeys(sheet, ar, ac)
        : collectAllPrecedentKeys(sheet, ar, ac);
      const range = selectionFromKeys(sheet, keys);
      if (range === null) {
        return { ok: false, code: "none" };
      }
      return { ok: true, range };
    }
    case "dependents": {
      const keys = collectDependentsKeys(sheet, ar, ac, directOnly);
      const range = selectionFromKeys(sheet, keys);
      if (range === null) {
        return { ok: false, code: "none" };
      }
      return { ok: true, range };
    }
      case "rowDiff": {
        if (sel.startRow === sel.endRow && sel.startCol === sel.endCol) {
          return { ok: false, code: "singleCell" };
        }
        const refCol = Math.max(sel.startCol, Math.min(sel.endCol, ac));
        const hits = new Set<string>();
        for (let r = sel.startRow; r <= sel.endRow; r++) {
          const ar0 = sheet.getMergeAnchorCell(r, refCol);
          const refContent = readAnchorContent(sheet, ar0.row, ar0.col);
          const refScalar = refContent.value;
          for (let c = sel.startCol; c <= sel.endCol; c++) {
          if (c === refCol) {
            continue;
          }
          const a = sheet.getMergeAnchorCell(r, c);
          const cur = readAnchorContent(sheet, a.row, a.col);
          const curScalar = cur.value;
          if (!scalarEqual(refScalar, curScalar)) {
            hits.add(Cell.key(a.row, a.col));
          }
        }
      }
      const range = selectionFromKeys(sheet, hits);
      if (range === null) {
        return { ok: false, code: "none" };
      }
      return { ok: true, range };
    }
    case "colDiff": {
      if (sel.startRow === sel.endRow && sel.startCol === sel.endCol) {
        return { ok: false, code: "singleCell" };
      }
      const refRow = Math.max(sel.startRow, Math.min(sel.endRow, ar));
      const hits = new Set<string>();
      for (let c = sel.startCol; c <= sel.endCol; c++) {
        const ar0 = sheet.getMergeAnchorCell(refRow, c);
        const refContent = readAnchorContent(sheet, ar0.row, ar0.col);
        const refScalar = refContent.value;
        for (let r = sel.startRow; r <= sel.endRow; r++) {
          if (r === refRow) {
            continue;
          }
          const a = sheet.getMergeAnchorCell(r, c);
          const cur = readAnchorContent(sheet, a.row, a.col);
          if (!scalarEqual(refScalar, cur.value)) {
            hits.add(Cell.key(a.row, a.col));
          }
        }
      }
      const range = selectionFromKeys(sheet, hits);
      if (range === null) {
        return { ok: false, code: "none" };
      }
      return { ok: true, range };
    }
    default:
      break;
  }

  const hits = new Set<string>();

  const consider = (anchorRow: number, anchorCol: number): boolean => {
    const cell = readAnchorContent(sheet, anchorRow, anchorCol);
    const isFormula = cell.formula !== null && cell.formula.length > 0;
    const blank = !isFormula && isBlankScalar(cell.value);

    switch (kind) {
      case "blanks":
        return blank;
      case "constants": {
        if (isFormula || blank) {
          return false;
        }
        return matchesSubtype(cell.value, filters);
      }
      case "formulas": {
        if (!isFormula) {
          return false;
        }
        return matchesSubtype(cell.value, filters);
      }
      case "visibleOnly": {
        return !sheet.isRowHidden(anchorRow) && !sheet.isColHidden(anchorCol);
      }
      case "conditionalFormats":
        return cellHasConditionalFormat(sheet, anchorRow, anchorCol);
      default:
        return false;
    }
  };

  switch (kind) {
    case "blanks":
    case "constants":
    case "formulas":
    case "visibleOnly":
    case "conditionalFormats":
      forEachDistinctMergeAnchorInSelection(sheet, sel, (r, c) => {
        if (consider(r, c)) {
          hits.add(Cell.key(r, c));
        }
      });
      break;
    default:
      return { ok: false, code: "unsupported" };
  }

  const range = selectionFromKeys(sheet, hits);
  if (range === null) {
    return { ok: false, code: "none" };
  }
  return { ok: true, range };
}

import { formatCellDisplayWithStyle, type Workbook, type Worksheet } from "@flexsheet/core";
import { columnIndexToLabel } from "@flexsheet/shared";

/**
 * 查找/替换。批注在数据层暂无，不产生匹配、不参与替换。
 */
export type FindWithinScope = "sheet" | "workbook";
export type FindSearchOrder = "row" | "column";
export type FindLookIn = "formulas" | "values" | "comments";

export interface FindReplaceScanOptions {
  readonly find: string;
  readonly within: FindWithinScope;
  readonly search: FindSearchOrder;
  readonly lookIn: FindLookIn;
  readonly matchCase: boolean;
  readonly matchEntireCell: boolean;
  /** true：严格区分全半角；false：对两侧做 NFKC 再比。 */
  readonly distinguishWidth: boolean;
}

export interface FindHit {
  readonly sheet: Worksheet;
  readonly sheetIndex: number;
  readonly row: number;
  readonly col: number;
  readonly start: number;
  readonly end: number;
}

function foldCase(s: string, matchCase: boolean): string {
  if (matchCase) {
    return s;
  }
  return s.toLowerCase();
}

function normalizeWidth(s: string, distinguish: boolean): string {
  if (distinguish) {
    return s;
  }
  return s.normalize("NFKC");
}

function sliceMatchesAt(
  hay: string,
  i: number,
  find: string,
  opt: Readonly<FindReplaceScanOptions>,
): boolean {
  if (find.length === 0) {
    return false;
  }
  if (i < 0 || i > hay.length) {
    return false;
  }
  if (opt.matchEntireCell) {
    if (i !== 0 || hay.length !== find.length) {
      return false;
    }
  } else {
    if (i + find.length > hay.length) {
      return false;
    }
  }
  const sub = hay.slice(i, i + find.length);
  const a0 = foldCase(normalizeWidth(sub, opt.distinguishWidth), opt.matchCase);
  const b0 = foldCase(normalizeWidth(find, opt.distinguishWidth), opt.matchCase);
  return a0 === b0;
}

function listOccurrencesInString(hay: string, find: string, opt: Readonly<FindReplaceScanOptions>): { start: number; end: number }[] {
  if (find === "" || hay === "") {
    return [];
  }
  if (opt.matchEntireCell) {
    if (sliceMatchesAt(hay, 0, find, opt)) {
      return [{ start: 0, end: hay.length }];
    }
    return [];
  }
  const out: { start: number; end: number }[] = [];
  for (let i = 0; i < hay.length; i++) {
    if (sliceMatchesAt(hay, i, find, opt)) {
      out.push({ start: i, end: i + find.length });
    }
  }
  return out;
}

export function getHayStringForCell(sheet: Worksheet, row: number, col: number, lookIn: FindLookIn): string | null {
  if (lookIn === "comments") {
    return null;
  }
  const cell = sheet.getCell(row, col);
  if (lookIn === "formulas") {
    if (cell.formula === null || cell.formula === "") {
      return null;
    }
    return cell.formula;
  }
  return formatCellDisplayWithStyle(cell.value, cell.style);
}

function compareHit(a: FindHit, b: FindHit): number {
  if (a.sheetIndex !== b.sheetIndex) {
    return a.sheetIndex - b.sheetIndex;
  }
  if (a.row !== b.row) {
    return a.row - b.row;
  }
  if (a.col !== b.col) {
    return a.col - b.col;
  }
  return a.start - b.start;
}

function* iterCellsInSheet(sh: Worksheet, order: FindSearchOrder): Generator<{ r: number; c: number }> {
  const { rowCount, colCount } = sh;
  if (order === "row") {
    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        yield { r, c };
      }
    }
  } else {
    for (let c = 0; c < colCount; c++) {
      for (let r = 0; r < rowCount; r++) {
        yield { r, c };
      }
    }
  }
}

function forEachCellInOrder(
  wb: Workbook,
  onlyActiveSheet: boolean,
  order: FindSearchOrder,
  visit: (sheet: Worksheet, sheetIndex: number, r: number, c: number) => void,
): void {
  if (onlyActiveSheet) {
    const sh = wb.getActiveSheet();
    if (sh === undefined) {
      return;
    }
    const si = wb.getSheets().indexOf(sh);
    for (const { r, c } of iterCellsInSheet(sh, order)) {
      visit(sh, si, r, c);
    }
    return;
  }
  const sheets = wb.getSheets();
  for (let si = 0; si < sheets.length; si++) {
    const sh = sheets[si]!;
    for (const { r, c } of iterCellsInSheet(sh, order)) {
      visit(sh, si, r, c);
    }
  }
}

/**
 * 按表顺序、按行/列、按格内 start 升序 收集所有命中。
 */
export function collectFindHits(wb: Workbook, opt: Readonly<FindReplaceScanOptions>): FindHit[] {
  if (opt.find === "") {
    return [];
  }
  const out: FindHit[] = [];
  const only = opt.within === "sheet";
  forEachCellInOrder(wb, only, opt.search, (sh, si, r, c) => {
    const hay = getHayStringForCell(sh, r, c, opt.lookIn);
    if (hay === null) {
      return;
    }
    const occ = listOccurrencesInString(hay, opt.find, opt);
    for (const o of occ) {
      out.push({ sheet: sh, sheetIndex: si, row: r, col: c, start: o.start, end: o.end });
    }
  });
  out.sort(compareHit);
  return out;
}

export function hitToA1Ref(hit: Readonly<FindHit>): string {
  return `${columnIndexToLabel(hit.col)}${hit.row + 1}`;
}

export function buildReplacedString(
  hay: string,
  hit: Readonly<Pick<FindHit, "start" | "end">>,
  replaceWith: string,
): string {
  return `${hay.slice(0, hit.start)}${replaceWith}${hay.slice(hit.end)}`;
}

export function isHitReplaceable(lookIn: FindLookIn): boolean {
  return lookIn === "formulas" || lookIn === "values";
}

/**
 * 第一处满足「位于 (end 位置) 或之后」的命中。用于单格替换后跳到「下一处」。
 * `at` 使用上一处 hit 的 `end`（即下一可能匹配可从此下标起）。
 */
export function firstHitIndexAtOrAfter(
  hits: readonly FindHit[],
  pos: { readonly si: number; readonly row: number; readonly col: number; readonly at: number },
): number {
  for (let i = 0; i < hits.length; i++) {
    const h = hits[i]!;
    if (h.sheetIndex < pos.si) {
      continue;
    }
    if (h.sheetIndex > pos.si) {
      return i;
    }
    if (h.row < pos.row) {
      continue;
    }
    if (h.row > pos.row) {
      return i;
    }
    if (h.col < pos.col) {
      continue;
    }
    if (h.col > pos.col) {
      return i;
    }
    if (h.start < pos.at) {
      continue;
    }
    return i;
  }
  return -1;
}

/**
 * 非重叠全局替换单格字符串（从左到右，每次从当前位置尝试匹配，命中则整段 `find.length` 前跳）。
 */
export function replaceAllInString(s: string, find: string, replaceWith: string, opt: Readonly<FindReplaceScanOptions>): string {
  if (find === "" || s === "") {
    return s;
  }
  if (opt.matchEntireCell) {
    if (listOccurrencesInString(s, find, opt).length > 0) {
      return replaceWith;
    }
    return s;
  }
  let out = "";
  let i = 0;
  const scan: FindReplaceScanOptions = { ...opt, matchEntireCell: false };
  while (i < s.length) {
    if (sliceMatchesAt(s, i, find, scan)) {
      out += replaceWith;
      i += find.length;
    } else {
      out += s[i] ?? "";
      i += 1;
    }
  }
  return out;
}

/**
 * 全部替换：扫描写回。`write` 可接 `SetCellValueCommand` 以支持撤销栈。
 */
export function applyReplaceAllWithWriter(
  wb: Workbook,
  find: string,
  replaceWith: string,
  opt: Readonly<FindReplaceScanOptions>,
  write: (sheet: Worksheet, row: number, col: number, newVal: string) => void,
): void {
  if (find === "" || !isHitReplaceable(opt.lookIn)) {
    return;
  }
  const only = opt.within === "sheet";
  const ro: FindReplaceScanOptions = { ...opt, find, matchEntireCell: false };
  forEachCellInOrder(wb, only, opt.search, (sh, _si, r, c) => {
    const hay = getHayStringForCell(sh, r, c, opt.lookIn);
    if (hay === null) {
      return;
    }
    let next: string;
    if (opt.matchEntireCell) {
      if (listOccurrencesInString(hay, find, opt).length === 0) {
        return;
      }
      next = replaceWith;
    } else {
      next = replaceAllInString(hay, find, replaceWith, ro);
    }
    if (next === hay) {
      return;
    }
    write(sh, r, c, next);
  });
}

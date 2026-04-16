import { columnLabelToIndex } from "@flexsheet/shared";
import { normalizeSelectionRange, type SelectionRange } from "@flexsheet/core";

/** 解析 `A1:O28`、`$A$1:$O$28`、`=$A$1:$O$28` 等选区引用；单格 `A1` 视为 `A1:A1`。 */
export function parseFormatAsTableRangeRef(input: string): SelectionRange | null {
  let s = input.trim();
  if (s.startsWith("=")) {
    s = s.slice(1).trim();
  }
  const colon = s.indexOf(":");
  if (colon < 0) {
    const one = parseSingleCellRef(s);
    if (one === null) {
      return null;
    }
    return normalizeSelectionRange({
      startRow: one.row,
      endRow: one.row,
      startCol: one.col,
      endCol: one.col,
    });
  }
  if (colon <= 0) {
    return null;
  }
  const a = parseSingleCellRef(s.slice(0, colon));
  const b = parseSingleCellRef(s.slice(colon + 1));
  if (a === null || b === null) {
    return null;
  }
  return normalizeSelectionRange({
    startRow: Math.min(a.row, b.row),
    endRow: Math.max(a.row, b.row),
    startCol: Math.min(a.col, b.col),
    endCol: Math.max(a.col, b.col),
  });
}

function parseSingleCellRef(raw: string): { readonly row: number; readonly col: number } | null {
  const s = raw.trim().replace(/\$/g, "");
  const m = /^([A-Za-z]+)(\d+)$/.exec(s);
  if (m === null) {
    return null;
  }
  const col = columnLabelToIndex(m[1]!);
  if (col === null) {
    return null;
  }
  const row1 = Number(m[2]);
  if (!Number.isInteger(row1) || row1 < 1) {
    return null;
  }
  return { row: row1 - 1, col };
}

import type { SelectionRange, Worksheet } from "@flexsheet/core";
import { normalizeSelectionRange } from "@flexsheet/core";

function cellHasData(sheet: Worksheet, row: number, col: number): boolean {
  if (row < 0 || col < 0 || row >= sheet.rowCount || col >= sheet.colCount) {
    return false;
  }
  const a = sheet.getMergeAnchorCell(row, col);
  const cell = sheet.getCell(a.row, a.col);
  if (cell.formula !== null && cell.formula.length > 0) {
    return true;
  }
  if (cell.value !== null && cell.value !== "") {
    return true;
  }
  return false;
}

/**
 * 与 Excel 接近的自动求和区域推测：优先同列连续数据（上方）→ 同行左侧 → 首块下方连续区域（BFS）。
 */
export function computeAutoSumRange(
  sheet: Worksheet,
  activeRow: number,
  activeCol: number,
): SelectionRange | null {
  // 1) 同列向上连续块
  if (activeRow > 0 && cellHasData(sheet, activeRow - 1, activeCol)) {
    let r1 = activeRow - 1;
    let r0 = r1;
    while (r0 - 1 >= 0 && cellHasData(sheet, r0 - 1, activeCol)) {
      r0 -= 1;
    }
    return normalizeSelectionRange({
      startRow: r0,
      endRow: r1,
      startCol: activeCol,
      endCol: activeCol,
    });
  }

  // 2) 同行向左连续块
  if (activeCol > 0 && cellHasData(sheet, activeRow, activeCol - 1)) {
    let c1 = activeCol - 1;
    let c0 = c1;
    while (c0 - 1 >= 0 && cellHasData(sheet, activeRow, c0 - 1)) {
      c0 -= 1;
    }
    return normalizeSelectionRange({
      startRow: activeRow,
      endRow: activeRow,
      startCol: c0,
      endCol: c1,
    });
  }

  // 3) 首块下方 BFS 连通区（可覆盖「公式格上方为空、数据在更远处」的表格块）
  const startRow = activeRow + 1;
  if (startRow >= sheet.rowCount) {
    return null;
  }

  let seedR = -1;
  let seedC = -1;
  outer: for (let r = startRow; r < sheet.rowCount; r++) {
    for (let c = 0; c < sheet.colCount; c++) {
      if (cellHasData(sheet, r, c)) {
        seedR = r;
        seedC = c;
        break outer;
      }
    }
  }
  if (seedR < 0) {
    return null;
  }

  const vis = new Set<string>();
  const q: Array<{ r: number; c: number }> = [{ r: seedR, c: seedC }];
  vis.add(`${seedR},${seedC}`);

  let br0 = seedR;
  let br1 = seedR;
  let bc0 = seedC;
  let bc1 = seedC;

  const nbs = (r: number, c: number): Array<{ r: number; c: number }> => {
    return [
      { r: r - 1, c },
      { r: r + 1, c },
      { r, c: c - 1 },
      { r, c: c + 1 },
    ];
  };

  for (let qi = 0; qi < q.length; qi++) {
    const { r, c } = q[qi]!;
    if (r < 0 || c < 0 || r >= sheet.rowCount || c >= sheet.colCount) {
      continue;
    }
    if (!cellHasData(sheet, r, c)) {
      continue;
    }
    br0 = Math.min(br0, r);
    br1 = Math.max(br1, r);
    bc0 = Math.min(bc0, c);
    bc1 = Math.max(bc1, c);
    if (vis.size > 8000) {
      break;
    }
    for (const p of nbs(r, c)) {
      const k = `${p.r},${p.c}`;
      if (vis.has(k)) {
        continue;
      }
      if (p.r < 0 || p.c < 0 || p.r >= sheet.rowCount || p.c >= sheet.colCount) {
        continue;
      }
      if (!cellHasData(sheet, p.r, p.c)) {
        continue;
      }
      vis.add(k);
      q.push(p);
    }
  }

  return normalizeSelectionRange({ startRow: br0, endRow: br1, startCol: bc0, endCol: bc1 });
}

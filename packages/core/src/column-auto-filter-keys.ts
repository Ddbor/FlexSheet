import type { Cell } from "./cell.js";
import { formatCellDisplayWithStyle } from "./excel-number-format.js";

/** 列表中的「空白」项与 `formatCellDisplayWithStyle` 得到空串的单元格对应。 */
export const AUTO_FILTER_BLANK_KEY = "__FS_BLANK__";

export function cellToAutoFilterDisplayKey(cell: Cell): string {
  const t = formatCellDisplayWithStyle(cell.value, cell.style);
  return t === "" ? AUTO_FILTER_BLANK_KEY : t;
}

export function compareAutoFilterDisplayKeys(a: string, b: string): number {
  if (a === AUTO_FILTER_BLANK_KEY) {
    return b === AUTO_FILTER_BLANK_KEY ? 0 : 1;
  }
  if (b === AUTO_FILTER_BLANK_KEY) {
    return -1;
  }
  const na = Number(a);
  const nb = Number(b);
  if (Number.isFinite(na) && Number.isFinite(nb) && a.trim() !== "" && b.trim() !== "") {
    if (na !== nb) {
      return na - nb;
    }
  }
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}

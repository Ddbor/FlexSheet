import type { CellStyle, SelectionRange, Worksheet } from "@flexsheet/core";
import { normalizeSelectionRange } from "@flexsheet/core";
import { cellScalarToEditString } from "@flexsheet/editor";

import { serializeTsv } from "./tsv-io.js";

function cellCopyText(sheet: Worksheet, row: number, col: number): string {
  const cell = sheet.getCell(row, col);
  if (cell.formula !== null && cell.formula.length > 0) {
    return cell.formula;
  }
  return cellScalarToEditString(cell.value);
}

function cloneStyle(style: CellStyle | null): CellStyle | null {
  if (style === null) {
    return null;
  }
  return { ...style };
}

export interface SerializedSelection {
  readonly tsv: string;
  readonly textRows: string[][];
  readonly styles: (CellStyle | null)[][];
}

export function serializeSelection(sheet: Worksheet, range: SelectionRange): SerializedSelection {
  const n = normalizeSelectionRange(range);
  const textRows: string[][] = [];
  const styles: (CellStyle | null)[][] = [];
  for (let r = n.startRow; r <= n.endRow; r++) {
    const tr: string[] = [];
    const sr: (CellStyle | null)[] = [];
    for (let c = n.startCol; c <= n.endCol; c++) {
      tr.push(cellCopyText(sheet, r, c));
      sr.push(cloneStyle(sheet.getCell(r, c).style));
    }
    textRows.push(tr);
    styles.push(sr);
  }
  return { tsv: serializeTsv(textRows), textRows, styles };
}

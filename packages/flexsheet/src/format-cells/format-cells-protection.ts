import {
  normalizeSelectionRange,
  type CellStyle,
  type CellStylePatch,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";

/** 可变的样式补丁对象（`CellStylePatch` 字段为 readonly）。 */
type CellStylePatchMutable = { -readonly [K in keyof CellStylePatch]: CellStylePatch[K] };

/**
 * 与 Excel 一致：未在样式中写 `locked` 时视为「锁定」；`formulaHidden` 仅在为 true 时生效。
 */
export function effectiveCellLocked(style: CellStyle | null): boolean {
  return style?.locked !== false;
}

export function effectiveFormulaHidden(style: CellStyle | null): boolean {
  return style?.formulaHidden === true;
}

/** 选区内主格（跳过合并覆盖格）的锁定/隐藏公式聚合状态，用于对话框初值。 */
export interface FormatCellsProtectionState {
  locked: boolean;
  lockedMixed: boolean;
  hidden: boolean;
  hiddenMixed: boolean;
}

export function inferFormatCellsProtectionState(
  sheet: Worksheet,
  range: SelectionRange,
): FormatCellsProtectionState {
  const n = normalizeSelectionRange(range);
  let firstL: boolean | undefined;
  let lockedMixed = false;
  let firstH: boolean | undefined;
  let hiddenMixed = false;
  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      if (sheet.isMergeCoveredCell(r, c)) {
        continue;
      }
      const st = sheet.getCell(r, c).style;
      const l = effectiveCellLocked(st);
      const h = effectiveFormulaHidden(st);
      if (firstL === undefined) {
        firstL = l;
      } else if (firstL !== l) {
        lockedMixed = true;
      }
      if (firstH === undefined) {
        firstH = h;
      } else if (firstH !== h) {
        hiddenMixed = true;
      }
    }
  }
  return {
    locked: firstL ?? true,
    lockedMixed,
    hidden: firstH ?? false,
    hiddenMixed,
  };
}

/** 对话框内勾选框状态（含不确定态）。 */
export interface FormatCellsProtectionUiState {
  locked: boolean;
  lockedIndeterminate: boolean;
  hidden: boolean;
  hiddenIndeterminate: boolean;
}

export function inferFormatCellsProtectionUiState(
  sheet: Worksheet,
  range: SelectionRange,
): FormatCellsProtectionUiState {
  const s = inferFormatCellsProtectionState(sheet, range);
  return {
    locked: s.locked,
    lockedIndeterminate: s.lockedMixed,
    hidden: s.hidden,
    hiddenIndeterminate: s.hiddenMixed,
  };
}

/** 与初值比较后生成补丁；不确定态且用户未改为确定时，不写入对应字段。 */
export function protectionPatchIfChanged(
  ui: FormatCellsProtectionUiState,
  initial: FormatCellsProtectionState,
): CellStylePatch {
  const p: CellStylePatchMutable = {};
  if (!ui.lockedIndeterminate) {
    if (initial.lockedMixed) {
      p.locked = ui.locked;
    } else if (ui.locked !== initial.locked) {
      p.locked = ui.locked;
    }
  }
  if (!ui.hiddenIndeterminate) {
    const wantHidden = ui.hidden;
    if (initial.hiddenMixed) {
      p.formulaHidden = wantHidden;
    } else if (wantHidden !== initial.hidden) {
      p.formulaHidden = wantHidden;
    }
  }
  return p;
}

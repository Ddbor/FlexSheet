import {
  normalizeSelectionRange,
  type ICommand,
  type SelectionRange,
  type Workbook,
  type Worksheet,
} from "@flexsheet/core";
import { columnIndexToLabel } from "@flexsheet/shared";
import { CreatePivotTableCommand, type PivotAggregateKind } from "./pivot-table-command.js";
import { showPivotTableFieldsPane } from "./pivot-table-fields-pane.js";
import { ensureFsSheetPromptStyles } from "./fs-dialog-styles.js";
import { parseFormatAsTableRangeRef } from "./format-as-table-range.js";

interface PivotFieldOption {
  readonly col: number;
  readonly label: string;
}

export interface PivotTableDialogHost {
  readonly workbook: Workbook | undefined;
  readonly selection: {
    getNormalizedRange(): SelectionRange;
    getActiveCell(): { readonly row: number; readonly col: number };
  };
  readonly workspace: { readonly commands: { execute(cmd: ICommand): void } };
  refresh(): void;
}

let pivotDialogStylesInjected = false;

function ensurePivotDialogStyles(): void {
  ensureFsSheetPromptStyles();
  if (pivotDialogStylesInjected) {
    return;
  }
  pivotDialogStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-pivot-dialog", "1");
  style.textContent = `
.fs-pivot-dialog.fs-sheet-prompt {
  width: min(540px, calc(100vw - 32px));
}
.fs-pivot-dialog__body.fs-sheet-prompt__body {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px 16px;
  padding-top: 12px;
}
.fs-pivot-dialog__full {
  grid-column: 1 / -1;
}
.fs-pivot-dialog__group {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.fs-pivot-dialog__label {
  font-size: 12px;
  color: #605e5c;
}
.fs-pivot-dialog__row {
  display: flex;
  align-items: center;
  gap: 8px;
}
.fs-pivot-dialog__row .fs-sheet-prompt__input,
.fs-pivot-dialog__row .fs-sheet-prompt__select {
  flex: 1;
}
.fs-pivot-dialog__field-hint {
  font-size: 12px;
  color: #605e5c;
  margin: 0;
}
.fs-pivot-dialog__target {
  border: 1px solid #edebe9;
  border-radius: 6px;
  padding: 10px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.fs-pivot-dialog__radio {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #323130;
}
.fs-pivot-dialog__warn {
  margin: 0;
  min-height: 18px;
  font-size: 12px;
  color: #a4262c;
}
`;
  document.head.appendChild(style);
}

function formatRangeAsAbsolute(range: SelectionRange): string {
  const n = normalizeSelectionRange(range);
  const c0 = columnIndexToLabel(n.startCol);
  const c1 = columnIndexToLabel(n.endCol);
  const r0 = n.startRow + 1;
  const r1 = n.endRow + 1;
  return `=$${c0}$${r0}:$${c1}$${r1}`;
}

function formatCellAsAbsolute(row: number, col: number): string {
  return `=$${columnIndexToLabel(col)}$${row + 1}`;
}

function toFieldLabel(value: unknown, fallback: string): string {
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed.length > 0) {
      return trimmed;
    }
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(value);
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return fallback;
}

function collectFieldOptions(
  sheet: Worksheet,
  range: SelectionRange,
  hasHeaders: boolean,
): readonly PivotFieldOption[] {
  const n = normalizeSelectionRange(range);
  const out: PivotFieldOption[] = [];
  for (let c = n.startCol; c <= n.endCol; c++) {
    const fallback = `列${c - n.startCol + 1}`;
    const value = hasHeaders ? sheet.getCell(n.startRow, c).value : null;
    out.push({
      col: c,
      label: toFieldLabel(value, fallback),
    });
  }
  return out;
}

function isRangeInSheet(sheet: Worksheet, range: SelectionRange): boolean {
  const n = normalizeSelectionRange(range);
  return (
    n.startRow >= 0 && n.startCol >= 0 && n.endRow < sheet.rowCount && n.endCol < sheet.colCount
  );
}

function parseDestinationCell(
  input: string,
): { readonly row: number; readonly col: number } | null {
  const parsed = parseFormatAsTableRangeRef(input);
  if (parsed === null) {
    return null;
  }
  return {
    row: parsed.startRow,
    col: parsed.startCol,
  };
}

function aggregateOptions(): readonly {
  readonly value: PivotAggregateKind;
  readonly label: string;
}[] {
  return [
    { value: "sum", label: "求和" },
    { value: "count", label: "计数" },
    { value: "average", label: "平均值" },
    { value: "max", label: "最大值" },
    { value: "min", label: "最小值" },
  ];
}

/**
 * 插入 -> 数据透视表：按选定源区域和字段生成透视结果（支持新建工作表或当前表指定起始格）。
 */
export function showPivotTableDialog(host: PivotTableDialogHost): void {
  const workbook = host.workbook;
  if (workbook === undefined) {
    return;
  }
  const sheet = workbook?.getActiveSheet();
  if (sheet === undefined) {
    return;
  }
  ensurePivotDialogStyles();

  const overlay = document.createElement("div");
  overlay.className = "fs-sheet-prompt-overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-sheet-prompt fs-pivot-dialog";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-pivot-dialog-title");

  const header = document.createElement("div");
  header.className = "fs-sheet-prompt__header";
  const titleEl = document.createElement("div");
  titleEl.id = "fs-pivot-dialog-title";
  titleEl.className = "fs-sheet-prompt__title";
  titleEl.textContent = "创建数据透视表";
  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "fs-sheet-prompt__close";
  closeBtn.setAttribute("aria-label", "关闭");
  closeBtn.textContent = "×";
  header.appendChild(titleEl);
  header.appendChild(closeBtn);

  const body = document.createElement("div");
  body.className = "fs-sheet-prompt__body fs-pivot-dialog__body";

  const rangeGroup = document.createElement("div");
  rangeGroup.className = "fs-pivot-dialog__group fs-pivot-dialog__full";
  const rangeLab = document.createElement("label");
  rangeLab.className = "fs-pivot-dialog__label";
  rangeLab.textContent = "请选择要分析的数据源（Excel 范围引用）";
  const rangeRow = document.createElement("div");
  rangeRow.className = "fs-pivot-dialog__row";
  const rangeInput = document.createElement("input");
  rangeInput.type = "text";
  rangeInput.className = "fs-sheet-prompt__input";
  rangeInput.value = formatRangeAsAbsolute(host.selection.getNormalizedRange());
  rangeInput.setAttribute("autocomplete", "off");
  rangeInput.setAttribute("spellcheck", "false");
  rangeRow.appendChild(rangeInput);
  const hasHeaderLabel = document.createElement("label");
  hasHeaderLabel.className = "fs-pivot-dialog__radio";
  const hasHeaderCheckbox = document.createElement("input");
  hasHeaderCheckbox.type = "checkbox";
  hasHeaderCheckbox.checked = true;
  const hasHeaderText = document.createElement("span");
  hasHeaderText.textContent = "首行作为字段名";
  hasHeaderLabel.appendChild(hasHeaderCheckbox);
  hasHeaderLabel.appendChild(hasHeaderText);
  rangeGroup.appendChild(rangeLab);
  rangeGroup.appendChild(rangeRow);
  rangeGroup.appendChild(hasHeaderLabel);

  const rowFieldGroup = document.createElement("div");
  rowFieldGroup.className = "fs-pivot-dialog__group";
  const rowFieldLab = document.createElement("label");
  rowFieldLab.className = "fs-pivot-dialog__label";
  rowFieldLab.textContent = "行";
  const rowFieldSel = document.createElement("select");
  rowFieldSel.className = "fs-sheet-prompt__select";
  rowFieldGroup.appendChild(rowFieldLab);
  rowFieldGroup.appendChild(rowFieldSel);

  const colFieldGroup = document.createElement("div");
  colFieldGroup.className = "fs-pivot-dialog__group";
  const colFieldLab = document.createElement("label");
  colFieldLab.className = "fs-pivot-dialog__label";
  colFieldLab.textContent = "列";
  const colFieldSel = document.createElement("select");
  colFieldSel.className = "fs-sheet-prompt__select";
  colFieldGroup.appendChild(colFieldLab);
  colFieldGroup.appendChild(colFieldSel);

  const valueFieldGroup = document.createElement("div");
  valueFieldGroup.className = "fs-pivot-dialog__group";
  const valueFieldLab = document.createElement("label");
  valueFieldLab.className = "fs-pivot-dialog__label";
  valueFieldLab.textContent = "值";
  const valueFieldSel = document.createElement("select");
  valueFieldSel.className = "fs-sheet-prompt__select";
  valueFieldGroup.appendChild(valueFieldLab);
  valueFieldGroup.appendChild(valueFieldSel);

  const aggGroup = document.createElement("div");
  aggGroup.className = "fs-pivot-dialog__group";
  const aggLab = document.createElement("label");
  aggLab.className = "fs-pivot-dialog__label";
  aggLab.textContent = "汇总方式";
  const aggSel = document.createElement("select");
  aggSel.className = "fs-sheet-prompt__select";
  for (const opt of aggregateOptions()) {
    const op = document.createElement("option");
    op.value = opt.value;
    op.textContent = opt.label;
    aggSel.appendChild(op);
  }
  aggSel.value = "sum";
  aggGroup.appendChild(aggLab);
  aggGroup.appendChild(aggSel);

  const targetGroup = document.createElement("div");
  targetGroup.className = "fs-pivot-dialog__group fs-pivot-dialog__full";
  const targetLab = document.createElement("label");
  targetLab.className = "fs-pivot-dialog__label";
  targetLab.textContent = "将数据透视表放置到";
  const targetPanel = document.createElement("div");
  targetPanel.className = "fs-pivot-dialog__target";
  const radioNewLabel = document.createElement("label");
  radioNewLabel.className = "fs-pivot-dialog__radio";
  const radioNew = document.createElement("input");
  radioNew.type = "radio";
  radioNew.name = "fs-pivot-target";
  radioNew.checked = true;
  const radioNewText = document.createElement("span");
  radioNewText.textContent = "新工作表";
  radioNewLabel.appendChild(radioNew);
  radioNewLabel.appendChild(radioNewText);
  const radioCurLabel = document.createElement("label");
  radioCurLabel.className = "fs-pivot-dialog__radio";
  const radioCur = document.createElement("input");
  radioCur.type = "radio";
  radioCur.name = "fs-pivot-target";
  const radioCurText = document.createElement("span");
  radioCurText.textContent = "现有工作表（起始单元格）";
  radioCurLabel.appendChild(radioCur);
  radioCurLabel.appendChild(radioCurText);
  const targetCellInput = document.createElement("input");
  targetCellInput.type = "text";
  targetCellInput.className = "fs-sheet-prompt__input";
  targetCellInput.value = formatCellAsAbsolute(
    host.selection.getActiveCell().row,
    host.selection.getActiveCell().col,
  );
  targetCellInput.disabled = true;
  targetPanel.appendChild(radioNewLabel);
  targetPanel.appendChild(radioCurLabel);
  targetPanel.appendChild(targetCellInput);
  targetGroup.appendChild(targetLab);
  targetGroup.appendChild(targetPanel);

  const warn = document.createElement("p");
  warn.className = "fs-pivot-dialog__warn fs-pivot-dialog__full";
  warn.textContent = "";

  const hint = document.createElement("p");
  hint.className = "fs-pivot-dialog__field-hint fs-pivot-dialog__full";
  hint.textContent = "字段来自数据源首行。建议源区域避免空行、空字段名。";

  body.appendChild(rangeGroup);
  body.appendChild(rowFieldGroup);
  body.appendChild(colFieldGroup);
  body.appendChild(valueFieldGroup);
  body.appendChild(aggGroup);
  body.appendChild(targetGroup);
  body.appendChild(warn);
  body.appendChild(hint);

  const footer = document.createElement("div");
  footer.className = "fs-sheet-prompt__footer";
  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.className = "fs-sheet-prompt__btn fs-sheet-prompt__btn--primary";
  okBtn.textContent = "确定";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "fs-sheet-prompt__btn fs-sheet-prompt__btn--secondary";
  cancelBtn.textContent = "取消";
  footer.appendChild(cancelBtn);
  footer.appendChild(okBtn);

  panel.appendChild(header);
  panel.appendChild(body);
  panel.appendChild(footer);
  overlay.appendChild(panel);
  document.body.appendChild(overlay);

  let currentFields: readonly PivotFieldOption[] = [];

  const setWarn = (text: string): void => {
    warn.textContent = text;
  };

  const setSelectOptions = (
    sel: HTMLSelectElement,
    items: readonly PivotFieldOption[],
    includeNone: boolean,
  ): void => {
    const prev = sel.value;
    sel.replaceChildren();
    if (includeNone) {
      const none = document.createElement("option");
      none.value = "";
      none.textContent = "(无)";
      sel.appendChild(none);
    }
    for (const it of items) {
      const op = document.createElement("option");
      op.value = String(it.col);
      op.textContent = it.label;
      sel.appendChild(op);
    }
    if (sel.querySelector(`option[value="${prev}"]`) !== null) {
      sel.value = prev;
      return;
    }
    if (!includeNone && items.length > 0) {
      sel.value = String(items[0].col);
    }
  };

  const refreshFieldOptions = (): void => {
    const parsed = parseFormatAsTableRangeRef(rangeInput.value);
    if (parsed === null || !isRangeInSheet(sheet, parsed)) {
      setWarn("数据源范围无效。");
      return;
    }
    const n = normalizeSelectionRange(parsed);
    const dataRows = hasHeaderCheckbox.checked ? n.endRow - n.startRow : n.endRow - n.startRow + 1;
    if (dataRows <= 0) {
      setWarn("数据源中没有可用于透视的数据行。");
      return;
    }
    currentFields = collectFieldOptions(sheet, n, hasHeaderCheckbox.checked);
    setSelectOptions(rowFieldSel, currentFields, false);
    setSelectOptions(colFieldSel, currentFields, true);
    setSelectOptions(valueFieldSel, currentFields, false);
    if (currentFields.length === 0) {
      setWarn("数据源范围至少需要一列。");
      return;
    }
    if (valueFieldSel.value === "" && currentFields.length > 0) {
      valueFieldSel.value = String(currentFields[0].col);
    }
    setWarn("");
  };

  const remove = (): void => {
    document.removeEventListener("keydown", onKeydown, true);
    overlay.remove();
  };

  const onKeydown = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      remove();
      return;
    }
    if (ev.key === "Enter" && !ev.isComposing) {
      const target = ev.target as HTMLElement | null;
      if (target !== null && target.tagName === "TEXTAREA") {
        return;
      }
      ev.preventDefault();
      void confirm();
    }
  };

  const confirm = async (): Promise<void> => {
    const parsedRange = parseFormatAsTableRangeRef(rangeInput.value);
    if (parsedRange === null || !isRangeInSheet(sheet, parsedRange)) {
      setWarn("请输入有效的数据源范围。");
      rangeInput.focus();
      rangeInput.select();
      return;
    }
    const sourceRange = normalizeSelectionRange(parsedRange);
    const rowFieldCol = Number.parseInt(rowFieldSel.value, 10);
    const colFieldColRaw = colFieldSel.value.trim();
    const colFieldCol = colFieldColRaw === "" ? null : Number.parseInt(colFieldColRaw, 10);
    const valueFieldCol = Number.parseInt(valueFieldSel.value, 10);
    if (!Number.isInteger(rowFieldCol) || !Number.isInteger(valueFieldCol)) {
      setWarn("请选择有效的行字段和值字段。");
      return;
    }
    const aggregate = aggSel.value as PivotAggregateKind;
    if (!["sum", "count", "average", "max", "min"].includes(aggregate)) {
      setWarn("请选择有效的汇总方式。");
      return;
    }
    if (!radioNew.checked && !radioCur.checked) {
      setWarn("请选择数据透视表放置位置。");
      return;
    }
    const columnFieldCols =
      colFieldCol !== null && Number.isInteger(colFieldCol) ? [colFieldCol] : [];
    const valueFields = [{ col: valueFieldCol, aggregate }];

    let cmd: CreatePivotTableCommand;
    if (radioNew.checked) {
      cmd = new CreatePivotTableCommand(workbook, sheet, {
        sourceRange,
        hasHeaders: hasHeaderCheckbox.checked,
        rowFieldCols: [rowFieldCol],
        columnFieldCols,
        filterFieldCols: [],
        valueFields,
        destination: { kind: "newSheet" },
      });
    } else {
      const targetCell = parseDestinationCell(targetCellInput.value);
      if (targetCell === null) {
        setWarn("请输入有效的目标单元格引用。");
        targetCellInput.focus();
        targetCellInput.select();
        return;
      }
      cmd = new CreatePivotTableCommand(workbook, sheet, {
        sourceRange,
        hasHeaders: hasHeaderCheckbox.checked,
        rowFieldCols: [rowFieldCol],
        columnFieldCols,
        filterFieldCols: [],
        valueFields,
        destination: {
          kind: "existingSheet",
          startRow: targetCell.row,
          startCol: targetCell.col,
        },
      });
    }
    const pivotId = cmd.getPivotDefinitionId();
    host.workspace.commands.execute(cmd);
    host.refresh();
    remove();
    const destSheet = radioNew.checked ? (workbook.getActiveSheet() ?? sheet) : sheet;
    showPivotTableFieldsPane(host, destSheet, pivotId);
  };

  closeBtn.addEventListener("click", remove);
  cancelBtn.addEventListener("click", remove);
  okBtn.addEventListener("click", () => {
    void confirm();
  });
  overlay.addEventListener("click", (ev) => {
    if (ev.target === overlay) {
      remove();
    }
  });
  rangeInput.addEventListener("input", refreshFieldOptions);
  hasHeaderCheckbox.addEventListener("change", refreshFieldOptions);
  radioNew.addEventListener("change", () => {
    targetCellInput.disabled = true;
  });
  radioCur.addEventListener("change", () => {
    targetCellInput.disabled = false;
    targetCellInput.focus();
    targetCellInput.select();
  });
  document.addEventListener("keydown", onKeydown, true);

  refreshFieldOptions();
  queueMicrotask(() => {
    rangeInput.focus();
    rangeInput.select();
  });
}

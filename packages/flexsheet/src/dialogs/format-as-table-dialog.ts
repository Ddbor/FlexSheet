import {
  normalizeSelectionRange,
  parseTableStyleRibbonCommand,
  type ICommand,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";
import { columnIndexToLabel } from "@flexsheet/shared";
import { ApplyFormatAsTableCommand } from "../commands/cell-style-commands.js";
import { ensureFsSheetPromptStyles } from "./fs-dialog-styles.js";
import { parseFormatAsTableRangeRef } from "./format-as-table-range.js";
import { createRangePickerIconSvg } from "../pivot/range-picker-icon.js";

let formatAsTableStylesInjected = false;
const LARGE_TABLE_STYLE_CONFIRM_THRESHOLD = 2000;

function ensureFormatAsTableDialogStyles(): void {
  ensureFsSheetPromptStyles();
  if (formatAsTableStylesInjected) {
    return;
  }
  formatAsTableStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-format-as-table", "1");
  style.textContent = `
.fs-fs-table__prompt.fs-sheet-prompt {
  width: min(420px, calc(100vw - 32px));
}
.fs-fs-table__body.fs-sheet-prompt__body {
  padding-top: 10px;
}
.fs-fs-table__hint {
  font-size: 13px;
  color: #323130;
  margin: 0 0 10px 0;
}
.fs-fs-table__range-row {
  display: flex;
  align-items: stretch;
  gap: 6px;
  margin-bottom: 12px;
}
.fs-fs-table__range-row .fs-sheet-prompt__input {
  flex: 1;
  min-width: 0;
}
.fs-fs-table__cb {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #323130;
  cursor: pointer;
  user-select: none;
}
.fs-fs-table__cb input {
  width: 16px;
  height: 16px;
  accent-color: #217346;
}
.fs-fs-large-confirm.fs-sheet-prompt {
  width: min(420px, calc(100vw - 32px));
}
.fs-fs-large-confirm__body.fs-sheet-prompt__body {
  text-align: center;
  line-height: 1.45;
  color: #323130;
}
.fs-fs-large-confirm__main {
  margin: 0;
  font-size: 16px;
  font-weight: 600;
}
.fs-fs-large-confirm__note {
  margin: 12px 0 0 0;
  font-size: 14px;
  color: #605e5c;
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

function countRangeCells(range: SelectionRange): number {
  const n = normalizeSelectionRange(range);
  const rows = n.endRow - n.startRow + 1;
  const cols = n.endCol - n.startCol + 1;
  return rows * cols;
}

function showLargeTableStyleConfirmDialog(): Promise<boolean> {
  return new Promise<boolean>((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "fs-sheet-prompt-overlay";
    overlay.setAttribute("role", "presentation");

    const panel = document.createElement("div");
    panel.className = "fs-sheet-prompt fs-fs-large-confirm";
    panel.setAttribute("role", "dialog");
    panel.setAttribute("aria-modal", "true");
    panel.setAttribute("aria-labelledby", "fs-large-op-confirm-title");

    const body = document.createElement("div");
    body.className = "fs-sheet-prompt__body fs-fs-large-confirm__body";
    const main = document.createElement("p");
    main.className = "fs-fs-large-confirm__main";
    main.id = "fs-large-op-confirm-title";
    main.textContent = "将要执行的操作会影响大量单元格，而且可能需要很长的时间才能完成。是否继续？";
    const note = document.createElement("p");
    note.className = "fs-fs-large-confirm__note";
    body.appendChild(main);
    body.appendChild(note);

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

    panel.appendChild(body);
    panel.appendChild(footer);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    let done = false;
    note.textContent = "注意：请确认是否继续执行该操作。";

    function onKey(e: KeyboardEvent): void {
      if (e.key === "Escape") {
        e.preventDefault();
        cleanup(false);
      }
    }

    function cleanup(result: boolean): void {
      if (done) {
        return;
      }
      done = true;
      document.removeEventListener("keydown", onKey, true);
      overlay.remove();
      resolve(result);
    }

    okBtn.addEventListener("click", () => cleanup(true));
    cancelBtn.addEventListener("click", () => cleanup(false));
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) {
        cleanup(false);
      }
    });
    document.addEventListener("keydown", onKey, true);
    queueMicrotask(() => okBtn.focus());
  });
}

export interface FormatAsTableDialogHost {
  readonly workbook: { getActiveSheet(): Worksheet | undefined } | undefined;
  readonly selection: { getNormalizedRange(): SelectionRange };
  readonly workspace: { readonly commands: { execute(cmd: ICommand): void } };
  refresh(): void;
  pickRangeReferenceFromSheet?(options?: {
    readonly mode?: "range" | "singleCell";
    readonly onRangePreview?: (displayRef: string) => void;
  }): Promise<string | null>;
}

/**
 * 套用表格格式：表数据来源、是否包含标题，确定后写入样式并（可选）为各列启用自动筛选。
 */
export function showFormatAsTableDialog(
  host: FormatAsTableDialogHost,
  ribbonCommandId: string,
): void {
  const parsed = parseTableStyleRibbonCommand(ribbonCommandId);
  if (parsed === null) {
    return;
  }
  const sheet = host.workbook?.getActiveSheet();
  if (sheet === undefined) {
    return;
  }

  ensureFormatAsTableDialogStyles();

  const overlay = document.createElement("div");
  overlay.className = "fs-sheet-prompt-overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-sheet-prompt fs-fs-table__prompt";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-format-as-table-title");

  const header = document.createElement("div");
  header.className = "fs-sheet-prompt__header";
  const titleEl = document.createElement("div");
  titleEl.id = "fs-format-as-table-title";
  titleEl.className = "fs-sheet-prompt__title";
  titleEl.textContent = "套用表格格式";
  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "fs-sheet-prompt__close";
  closeBtn.setAttribute("aria-label", "关闭");
  closeBtn.textContent = "×";
  header.appendChild(titleEl);
  header.appendChild(closeBtn);

  const body = document.createElement("div");
  body.className = "fs-sheet-prompt__body fs-fs-table__body";

  const hint = document.createElement("p");
  hint.className = "fs-fs-table__hint";
  hint.textContent = "表数据的来源?";

  const rangeRow = document.createElement("div");
  rangeRow.className = "fs-fs-table__range-row";
  const rangeInput = document.createElement("input");
  rangeInput.type = "text";
  rangeInput.className = "fs-sheet-prompt__input";
  rangeInput.setAttribute("autocomplete", "off");
  rangeInput.setAttribute("spellcheck", "false");
  rangeInput.value = formatRangeAsAbsolute(host.selection.getNormalizedRange());

  const pickBtn = document.createElement("button");
  pickBtn.type = "button";
  pickBtn.className = "fs-sheet-prompt__range-pick";
  pickBtn.title = "在工作表中选定区域";
  pickBtn.setAttribute("aria-label", pickBtn.title);
  pickBtn.appendChild(createRangePickerIconSvg());
  const pick = host.pickRangeReferenceFromSheet;
  if (pick === undefined) {
    pickBtn.disabled = true;
    pickBtn.title = "当前宿主不支持从工作表选定区域";
  }

  rangeRow.appendChild(rangeInput);
  rangeRow.appendChild(pickBtn);

  const cbLabel = document.createElement("label");
  cbLabel.className = "fs-fs-table__cb";
  const cb = document.createElement("input");
  cb.type = "checkbox";
  // 与 Excel 一致：默认勾选「表包含标题」，表头行自动显示列筛选下拉。
  cb.checked = true;
  const cbText = document.createElement("span");
  cbText.textContent = "表包含标题";
  cbLabel.appendChild(cb);
  cbLabel.appendChild(cbText);

  body.appendChild(hint);
  body.appendChild(rangeRow);
  body.appendChild(cbLabel);

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

  const rangePickBar = document.createElement("div");
  rangePickBar.className = "fs-sheet-prompt-range-pick-bar";
  rangePickBar.setAttribute("aria-hidden", "true");
  const rangePickBarTitle = document.createElement("div");
  rangePickBarTitle.className = "fs-sheet-prompt-range-pick-bar__title";
  rangePickBarTitle.textContent = "套用表格格式";
  const rangePickBarHint = document.createElement("p");
  rangePickBarHint.className = "fs-sheet-prompt-range-pick-bar__hint";
  rangePickBarHint.textContent = "正在选择区域，在表格中拖拽；完成后松开鼠标。按 Esc 取消。";
  const rangePickBarRow = document.createElement("div");
  rangePickBarRow.className = "fs-sheet-prompt-range-pick-bar__row";
  const rangePickBarInput = document.createElement("input");
  rangePickBarInput.type = "text";
  rangePickBarInput.readOnly = true;
  rangePickBarInput.className = "fs-sheet-prompt-range-pick-bar__input";
  rangePickBarInput.setAttribute("aria-label", "当前选定区域预览");
  rangePickBarInput.setAttribute("autocomplete", "off");
  const rangePickBarIconWrap = document.createElement("span");
  rangePickBarIconWrap.className = "fs-sheet-prompt-range-pick-bar__icon-wrap";
  rangePickBarIconWrap.appendChild(createRangePickerIconSvg());
  rangePickBarIconWrap.setAttribute("aria-hidden", "true");
  rangePickBarRow.appendChild(rangePickBarInput);
  rangePickBarRow.appendChild(rangePickBarIconWrap);
  rangePickBar.appendChild(rangePickBarTitle);
  rangePickBar.appendChild(rangePickBarHint);
  rangePickBar.appendChild(rangePickBarRow);
  overlay.appendChild(rangePickBar);

  document.body.appendChild(overlay);

  let rangePicking = false;

  const runRangePick = async (): Promise<void> => {
    if (pick === undefined) {
      return;
    }
    rangePicking = true;
    rangePickBar.setAttribute("aria-hidden", "false");
    overlay.classList.add("fs-sheet-prompt-overlay--range-pick");
    try {
      const ref = await pick.call(host, {
        mode: "range",
        onRangePreview: (display) => {
          rangePickBarInput.value = display;
        },
      });
      if (ref !== null) {
        rangeInput.value = ref;
      }
    } finally {
      rangePicking = false;
      rangePickBar.setAttribute("aria-hidden", "true");
      rangePickBarInput.value = "";
      overlay.classList.remove("fs-sheet-prompt-overlay--range-pick");
    }
    queueMicrotask(() => {
      rangeInput.focus();
      rangeInput.select();
    });
  };

  function onKey(e: KeyboardEvent): void {
    if (e.key === "Escape") {
      if (rangePicking) {
        return;
      }
      e.preventDefault();
      remove();
    }
  }

  const remove = (): void => {
    document.removeEventListener("keydown", onKey, true);
    overlay.remove();
  };

  let confirming = false;
  const confirm = async (): Promise<void> => {
    if (confirming) {
      return;
    }
    confirming = true;
    const parsedRange = parseFormatAsTableRangeRef(rangeInput.value);
    if (parsedRange === null) {
      rangeInput.focus();
      rangeInput.setAttribute("aria-invalid", "true");
      confirming = false;
      return;
    }
    rangeInput.removeAttribute("aria-invalid");
    const n = normalizeSelectionRange(parsedRange);
    if (
      n.startRow < 0 ||
      n.endRow >= sheet.rowCount ||
      n.startCol < 0 ||
      n.endCol >= sheet.colCount
    ) {
      rangeInput.focus();
      rangeInput.setAttribute("aria-invalid", "true");
      confirming = false;
      return;
    }
    const needLargeConfirm = countRangeCells(n) >= LARGE_TABLE_STYLE_CONFIRM_THRESHOLD;
    if (needLargeConfirm) {
      const shouldContinue = await showLargeTableStyleConfirmDialog();
      if (!shouldContinue) {
        confirming = false;
        return;
      }
    }
    const hasHeaders = cb.checked;
    host.workspace.commands.execute(new ApplyFormatAsTableCommand(sheet, n, parsed, hasHeaders));
    host.refresh();
    remove();
    confirming = false;
  };

  closeBtn.addEventListener("click", remove);
  cancelBtn.addEventListener("click", remove);
  okBtn.addEventListener("click", () => {
    void confirm();
  });
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) {
      remove();
    }
  });
  document.addEventListener("keydown", onKey, true);
  if (pick !== undefined) {
    pickBtn.addEventListener("click", () => {
      void runRangePick();
    });
  }
  queueMicrotask(() => {
    rangeInput.focus();
    rangeInput.select();
  });
}

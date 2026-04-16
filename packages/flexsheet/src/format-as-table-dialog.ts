import {
  normalizeSelectionRange,
  parseTableStyleRibbonCommand,
  type ICommand,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";
import { columnIndexToLabel } from "@flexsheet/shared";
import { ApplyFormatAsTableCommand } from "./cell-style-commands.js";
import { ensureFsSheetPromptStyles } from "./fs-dialog-styles.js";
import { parseFormatAsTableRangeRef } from "./format-as-table-range.js";

let formatAsTableStylesInjected = false;

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
.fs-fs-table__range-pick {
  flex-shrink: 0;
  width: 36px;
  padding: 0;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  cursor: not-allowed;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #605e5c;
}
.fs-fs-table__range-pick svg {
  width: 18px;
  height: 18px;
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

function rangePickerIconSvg(): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.6");
  const g = document.createElementNS(ns, "g");
  const r = document.createElementNS(ns, "rect");
  r.setAttribute("x", "3");
  r.setAttribute("y", "4");
  r.setAttribute("width", "14");
  r.setAttribute("height", "14");
  r.setAttribute("rx", "1.5");
  const p = document.createElementNS(ns, "path");
  p.setAttribute(
    "d",
    "M17 8l4-2v12l-4-2M7 9h5M7 12h5M7 15h4",
  );
  g.appendChild(r);
  g.appendChild(p);
  svg.appendChild(g);
  return svg;
}

export interface FormatAsTableDialogHost {
  readonly workbook: { getActiveSheet(): Worksheet | undefined } | undefined;
  readonly selection: { getNormalizedRange(): SelectionRange };
  readonly workspace: { readonly commands: { execute(cmd: ICommand): void } };
  refresh(): void;
}

/**
 * 套用表格格式：表数据来源、是否包含标题，确定后写入样式并（可选）为各列启用自动筛选。
 */
export function showFormatAsTableDialog(host: FormatAsTableDialogHost, ribbonCommandId: string): void {
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
  pickBtn.className = "fs-fs-table__range-pick";
  pickBtn.title = "在工作表中选定区域（暂不支持）";
  pickBtn.setAttribute("aria-label", pickBtn.title);
  pickBtn.disabled = true;
  pickBtn.appendChild(rangePickerIconSvg());

  rangeRow.appendChild(rangeInput);
  rangeRow.appendChild(pickBtn);

  const cbLabel = document.createElement("label");
  cbLabel.className = "fs-fs-table__cb";
  const cb = document.createElement("input");
  cb.type = "checkbox";
  const sel0 = normalizeSelectionRange(host.selection.getNormalizedRange());
  cb.checked = sel0.endRow > sel0.startRow;
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
  document.body.appendChild(overlay);

  function onKey(e: KeyboardEvent): void {
    if (e.key === "Escape") {
      e.preventDefault();
      remove();
    }
  }

  const remove = (): void => {
    document.removeEventListener("keydown", onKey, true);
    overlay.remove();
  };

  const confirm = (): void => {
    const parsedRange = parseFormatAsTableRangeRef(rangeInput.value);
    if (parsedRange === null) {
      rangeInput.focus();
      rangeInput.setAttribute("aria-invalid", "true");
      return;
    }
    rangeInput.removeAttribute("aria-invalid");
    const n = normalizeSelectionRange(parsedRange);
    if (n.startRow < 0 || n.endRow >= sheet.rowCount || n.startCol < 0 || n.endCol >= sheet.colCount) {
      rangeInput.focus();
      rangeInput.setAttribute("aria-invalid", "true");
      return;
    }
    const hasHeaders = cb.checked;
    host.workspace.commands.execute(new ApplyFormatAsTableCommand(sheet, n, parsed, hasHeaders));
    if (hasHeaders) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        sheet.enableColumnAutoFilterFromSelection(n.startRow, c, n);
      }
    }
    host.refresh();
    remove();
  };

  closeBtn.addEventListener("click", remove);
  cancelBtn.addEventListener("click", remove);
  okBtn.addEventListener("click", confirm);
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) {
      remove();
    }
  });
  document.addEventListener("keydown", onKey, true);
  queueMicrotask(() => {
    rangeInput.focus();
    rangeInput.select();
  });
}

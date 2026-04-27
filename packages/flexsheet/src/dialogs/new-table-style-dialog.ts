import type { CellStylePatch } from "@flexsheet/core";
import type { FlexSheet } from "../flex-sheet.js";
import { attachDraggableDialogPanel } from "@flexsheet/shared";
import { ensureFsSheetPromptStyles } from "./fs-dialog-styles.js";
import { mountFormatCellsDialog } from "../format-cells/format-cells-dialog.js";
import type { FormatCellsBorderState } from "../format-cells/format-cells-border.js";

let newTableStyleDialogCssInjected = false;

function ensureNewTableStyleDialogStyles(): void {
  ensureFsSheetPromptStyles();
  if (newTableStyleDialogCssInjected) {
    return;
  }
  newTableStyleDialogCssInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-new-table-style-dialog", "1");
  style.textContent = `
.fs-new-table-style__overlay {
  position: fixed;
  inset: 0;
  z-index: 4000;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(0, 0, 0, 0.24);
}
.fs-new-table-style__panel {
  width: min(620px, calc(100vw - 24px));
  min-height: 420px;
  background: #ffffff;
  border: 1px solid #d1d1d1;
  border-radius: 8px;
  box-shadow: 0 12px 36px rgba(0, 0, 0, 0.18);
  display: flex;
  flex-direction: column;
  overflow: hidden;
}
.fs-new-table-style__header {
  height: 36px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 20px;
  line-height: 1;
  font-weight: 500;
  color: #323130;
}
.fs-new-table-style__body {
  padding: 14px 18px 10px;
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.fs-new-table-style__name-row {
  display: grid;
  grid-template-columns: auto 1fr;
  align-items: center;
  gap: 10px;
}
.fs-new-table-style__name {
  font-size: 14px;
  color: #323130;
}
.fs-new-table-style__input {
  height: 30px;
  border: 1px solid #8a8886;
  border-radius: 2px;
  font-size: 14px;
  padding: 0 8px;
}
.fs-new-table-style__main {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 190px;
  gap: 14px;
  min-height: 230px;
}
.fs-new-table-style__list-wrap {
  display: flex;
  flex-direction: column;
}
.fs-new-table-style__label {
  font-size: 14px;
  color: #323130;
  margin-bottom: 6px;
}
.fs-new-table-style__list {
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  padding: 3px;
  overflow: auto;
  min-height: 188px;
  max-height: 230px;
}
.fs-new-table-style__item {
  width: 100%;
  text-align: left;
  border: 0;
  background: transparent;
  height: 24px;
  padding: 0 6px;
  font-size: 13px;
  color: #323130;
  cursor: pointer;
  border-radius: 2px;
}
.fs-new-table-style__item:hover {
  background: #e8f3ec;
  color: #1f1f1f;
}
.fs-new-table-style__item--active {
  background: #107c41;
  color: #ffffff;
  font-weight: 600;
}
.fs-new-table-style__preview-wrap {
  border-left: 1px solid #e1dfdd;
  padding-left: 14px;
  display: flex;
  flex-direction: column;
}
.fs-new-table-style__preview {
  margin-top: 6px;
  width: 150px;
  height: 82px;
  border: 1px solid #d1d1d1;
  background:
    repeating-linear-gradient(to right, transparent 0, transparent 28px, #7f7f7f 28px, #7f7f7f 30px),
    repeating-linear-gradient(to bottom, transparent 0, transparent 14px, #7f7f7f 14px, #7f7f7f 16px);
  background-color: #ffffff;
}
.fs-new-table-style__format-row {
  display: flex;
  gap: 8px;
  margin-top: 10px;
}
.fs-new-table-style__btn {
  height: 28px;
  border-radius: 14px;
  border: 1px solid #8a8886;
  background: #ffffff;
  color: #323130;
  padding: 0 14px;
  font-size: 13px;
  cursor: pointer;
}
.fs-new-table-style__btn:hover {
  border-color: #605e5c;
}
.fs-new-table-style__btn:disabled {
  opacity: 0.45;
  cursor: default;
}
.fs-new-table-style__checkbox {
  margin-top: 2px;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #323130;
}
.fs-new-table-style__checkbox input {
  width: 14px;
  height: 14px;
  accent-color: #107c41;
}
.fs-new-table-style__footer {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
  padding: 12px 18px 14px;
  border-top: 1px solid #edebe9;
}
.fs-new-table-style__footer .fs-new-table-style__btn--primary {
  background: #107c41;
  border-color: #107c41;
  color: #ffffff;
}
`;
  document.head.appendChild(style);
}

const TABLE_STYLE_ELEMENTS = [
  "整个表",
  "第一列条纹",
  "第二列条纹",
  "第一行条纹",
  "第二行条纹",
  "最后一列",
  "第一列",
  "标题行",
  "汇总行",
  "第一个标题单元格",
  "最后一个标题单元格",
  "第一个汇总单元格",
  "最后一个汇总单元格",
] as const;
type TableStyleElementId = (typeof TABLE_STYLE_ELEMENTS)[number];

interface TableStyleElementFormat {
  readonly basePatch: CellStylePatch;
  readonly border: { readonly apply: boolean; readonly state: FormatCellsBorderState };
}

/**
 * Ribbon「套用表格格式 → 新建表样式」对话框。
 * 当前实现用于交互与样式编辑入口，后续可将 `styleByElement` 接入真正的表样式持久化。
 */
export function showNewTableStyleDialog(
  flex: FlexSheet & {
    createCustomTableStyle?: (name: string) => string;
  },
): void {
  ensureNewTableStyleDialogStyles();

  const overlay = document.createElement("div");
  overlay.className = "fs-new-table-style__overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-new-table-style__panel";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-new-table-style-title");
  overlay.appendChild(panel);

  const header = document.createElement("div");
  header.className = "fs-new-table-style__header";
  const title = document.createElement("div");
  title.id = "fs-new-table-style-title";
  title.textContent = "新建表样式";
  header.appendChild(title);
  panel.appendChild(header);

  const body = document.createElement("div");
  body.className = "fs-new-table-style__body";
  panel.appendChild(body);

  const nameRow = document.createElement("div");
  nameRow.className = "fs-new-table-style__name-row";
  const nameLabel = document.createElement("label");
  nameLabel.className = "fs-new-table-style__name";
  nameLabel.textContent = "名称:";
  const nameInput = document.createElement("input");
  nameInput.className = "fs-new-table-style__input";
  nameInput.type = "text";
  nameInput.value = "表样式 1";
  nameLabel.htmlFor = "fs-new-table-style-name";
  nameInput.id = "fs-new-table-style-name";
  nameRow.appendChild(nameLabel);
  nameRow.appendChild(nameInput);
  body.appendChild(nameRow);

  const main = document.createElement("div");
  main.className = "fs-new-table-style__main";
  body.appendChild(main);

  const listWrap = document.createElement("div");
  listWrap.className = "fs-new-table-style__list-wrap";
  const listLabel = document.createElement("div");
  listLabel.className = "fs-new-table-style__label";
  listLabel.textContent = "表元素";
  const list = document.createElement("div");
  list.className = "fs-new-table-style__list";
  list.setAttribute("role", "listbox");
  listWrap.appendChild(listLabel);
  listWrap.appendChild(list);
  main.appendChild(listWrap);

  const previewWrap = document.createElement("div");
  previewWrap.className = "fs-new-table-style__preview-wrap";
  const previewLabel = document.createElement("div");
  previewLabel.className = "fs-new-table-style__label";
  previewLabel.textContent = "预览";
  const preview = document.createElement("div");
  preview.className = "fs-new-table-style__preview";
  previewWrap.appendChild(previewLabel);
  previewWrap.appendChild(preview);
  main.appendChild(previewWrap);

  const formatRow = document.createElement("div");
  formatRow.className = "fs-new-table-style__format-row";
  const formatBtn = document.createElement("button");
  formatBtn.type = "button";
  formatBtn.className = "fs-new-table-style__btn";
  formatBtn.textContent = "格式";
  const clearBtn = document.createElement("button");
  clearBtn.type = "button";
  clearBtn.className = "fs-new-table-style__btn";
  clearBtn.textContent = "清除";
  formatRow.appendChild(formatBtn);
  formatRow.appendChild(clearBtn);
  listWrap.appendChild(formatRow);

  const defaultCbLabel = document.createElement("label");
  defaultCbLabel.className = "fs-new-table-style__checkbox";
  const defaultCb = document.createElement("input");
  defaultCb.type = "checkbox";
  const defaultText = document.createElement("span");
  defaultText.textContent = "设置为此工作簿的默认表样式";
  defaultCbLabel.appendChild(defaultCb);
  defaultCbLabel.appendChild(defaultText);
  body.appendChild(defaultCbLabel);

  const footer = document.createElement("div");
  footer.className = "fs-new-table-style__footer";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "fs-new-table-style__btn";
  cancelBtn.textContent = "取消";
  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.className = "fs-new-table-style__btn fs-new-table-style__btn--primary";
  okBtn.textContent = "确定";
  footer.appendChild(cancelBtn);
  footer.appendChild(okBtn);
  panel.appendChild(footer);

  const styleByElement = new Map<TableStyleElementId, TableStyleElementFormat>();
  let activeElement: TableStyleElementId = TABLE_STYLE_ELEMENTS[0];
  const itemBtns = new Map<TableStyleElementId, HTMLButtonElement>();

  const syncSelection = (): void => {
    for (const [id, btn] of itemBtns) {
      btn.classList.toggle("fs-new-table-style__item--active", id === activeElement);
    }
    clearBtn.disabled = !styleByElement.has(activeElement);
  };

  for (const item of TABLE_STYLE_ELEMENTS) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "fs-new-table-style__item";
    btn.setAttribute("role", "option");
    btn.textContent = item;
    btn.addEventListener("click", () => {
      activeElement = item;
      syncSelection();
    });
    list.appendChild(btn);
    itemBtns.set(item, btn);
  }
  syncSelection();

  const onKeyDown = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      close();
    }
  };

  const close = (): void => {
    document.removeEventListener("keydown", onKeyDown, true);
    overlay.remove();
    flex.refresh();
  };

  formatBtn.addEventListener("click", () => {
    mountFormatCellsDialog({
      flex,
      visibleTabs: ["font", "border", "fill"],
      hideDescription: true,
      onApply: (payload) => {
        styleByElement.set(activeElement, {
          basePatch: payload.basePatch,
          border: payload.border,
        });
        syncSelection();
      },
    });
  });

  clearBtn.addEventListener("click", () => {
    styleByElement.delete(activeElement);
    syncSelection();
  });

  cancelBtn.addEventListener("click", close);
  okBtn.addEventListener("click", () => {
    const name = nameInput.value.trim();
    if (name.length === 0) {
      nameInput.focus();
      nameInput.select();
      return;
    }
    flex.createCustomTableStyle?.(name);
    close();
  });
  overlay.addEventListener("pointerdown", (ev) => {
    if (ev.target === overlay) {
      close();
    }
  });
  document.addEventListener("keydown", onKeyDown, true);
  document.body.appendChild(overlay);
  attachDraggableDialogPanel(panel, header);
  queueMicrotask(() => {
    nameInput.focus();
    nameInput.select();
  });
}

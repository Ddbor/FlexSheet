import { normalizeSelectionRange, type Worksheet } from "@flexsheet/core";

import type { FlexSheet } from "../flex-sheet.js";
import {
  collectPivotFilterDistinctKeys,
  normalizeFilterSelectedKeys,
  UpdatePivotTableFiltersCommand,
} from "./pivot-table-command.js";

let pivotFilterStylesInjected = false;

function ensurePivotFilterStyles(): void {
  if (pivotFilterStylesInjected) {
    return;
  }
  pivotFilterStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-pivot-filter", "1");
  style.textContent = `
.fs-pivot-filter {
  position: fixed;
  z-index: 10005;
  min-width: 220px;
  max-width: min(280px, calc(100vw - 24px));
  max-height: min(380px, calc(100vh - 24px));
  box-sizing: border-box;
  background: #fff;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  box-shadow: 0 4px 16px rgba(0,0,0,0.15);
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  font-size: 12px;
  color: #323130;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}
.fs-pivot-filter__title {
  padding: 8px 10px;
  font-weight: 600;
  border-bottom: 1px solid #edebe9;
  flex-shrink: 0;
}
.fs-pivot-filter__list {
  flex: 1;
  min-height: 100px;
  max-height: 240px;
  overflow: auto;
  margin: 6px 8px;
  border: 1px solid #edebe9;
  border-radius: 2px;
  padding: 4px 0;
}
.fs-pivot-filter__item {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 8px;
  cursor: pointer;
  user-select: none;
}
.fs-pivot-filter__item:hover {
  background: #f3f2f1;
}
.fs-pivot-filter__footer {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding: 8px 10px;
  border-top: 1px solid #edebe9;
  flex-shrink: 0;
}
.fs-pivot-filter__footer button {
  min-width: 64px;
  padding: 4px 10px;
  border-radius: 2px;
  border: 1px solid #8a8886;
  background: #fff;
  cursor: pointer;
  font: inherit;
}
.fs-pivot-filter__footer button.fs-pivot-filter__ok {
  background: #107c10;
  border-color: #107c10;
  color: #fff;
}
.fs-pivot-filter__footer button:hover {
  filter: brightness(0.97);
}
`;
  document.head.appendChild(style);
}

export interface OpenPivotFilterPanelOptions {
  readonly flex: FlexSheet;
  /** 透视输出所在工作表（与命中单元格一致）。 */
  readonly pivotSheet: Worksheet;
  readonly pivotDefId: string;
  readonly filterFieldIndex: number;
  readonly clientX: number;
  readonly clientY: number;
}

/**
 * 在透视表筛选行上打开字段值多选（与列筛选类似：空数组=全部）。
 */
export function openPivotFilterPanel(options: OpenPivotFilterPanelOptions): void {
  const { flex, pivotSheet, pivotDefId, filterFieldIndex, clientX, clientY } = options;
  const wb = flex.workbook;
  const def = pivotSheet.getPivotTableDefinitionsSnapshot().find((d) => d.id === pivotDefId);
  if (def === undefined) {
    return;
  }
  if (filterFieldIndex < 0 || filterFieldIndex >= def.filterFieldCols.length) {
    return;
  }
  const sourceSheet = wb.getSheet(def.sourceSheetIndex);
  if (sourceSheet === undefined) {
    return;
  }

  ensurePivotFilterStyles();
  const n = normalizeSelectionRange(def.sourceRange);
  const rStart = def.hasHeaders ? n.startRow + 1 : n.startRow;
  const rEnd = n.endRow;
  const filterCol = def.filterFieldCols[filterFieldIndex]!;
  const distinct = collectPivotFilterDistinctKeys(sourceSheet, rStart, rEnd, filterCol);
  const normalized = normalizeFilterSelectedKeys(def.filterFieldCols, def.filterSelectedKeys);
  const active = normalized[filterFieldIndex] ?? [];
  const activeSet = new Set(active);
  const allSelectedInitially =
    active.length === 0 || active.length === distinct.length && distinct.every((k) => activeSet.has(k));

  const root = document.createElement("div");
  root.className = "fs-pivot-filter";
  root.setAttribute("role", "dialog");
  root.setAttribute("aria-modal", "true");

  const fieldName =
    typeof sourceSheet.getCell(n.startRow, filterCol).value === "string"
      ? String(sourceSheet.getCell(n.startRow, filterCol).value).trim() || `列${filterCol + 1}`
      : `列${filterCol + 1}`;

  const title = document.createElement("div");
  title.className = "fs-pivot-filter__title";
  title.textContent = fieldName;

  const list = document.createElement("div");
  list.className = "fs-pivot-filter__list";

  const checked = new Set<string>();
  if (allSelectedInitially) {
    for (const k of distinct) {
      checked.add(k);
    }
  } else {
    for (const k of active) {
      checked.add(k);
    }
  }

  for (const key of distinct) {
    const row = document.createElement("label");
    row.className = "fs-pivot-filter__item";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.dataset.key = key;
    cb.checked = checked.has(key);
    cb.addEventListener("change", () => {
      if (cb.checked) {
        checked.add(key);
      } else {
        checked.delete(key);
      }
    });
    const span = document.createElement("span");
    span.textContent = key;
    row.appendChild(cb);
    row.appendChild(span);
    list.appendChild(row);
  }

  const footer = document.createElement("div");
  footer.className = "fs-pivot-filter__footer";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.textContent = "取消";
  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.className = "fs-pivot-filter__ok";
  okBtn.textContent = "确定";
  footer.appendChild(cancelBtn);
  footer.appendChild(okBtn);

  root.appendChild(title);
  root.appendChild(list);
  root.appendChild(footer);
  document.body.appendChild(root);

  const place = (): void => {
    root.style.position = "fixed";
    const w = root.offsetWidth;
    const h = root.offsetHeight;
    let left = clientX;
    let top = clientY + 4;
    if (left + w > window.innerWidth - 8) {
      left = Math.max(8, window.innerWidth - w - 8);
    }
    if (top + h > window.innerHeight - 8) {
      top = Math.max(8, clientY - h - 4);
    }
    root.style.left = `${Math.round(left)}px`;
    root.style.top = `${Math.round(top)}px`;
  };
  place();

  const remove = (): void => {
    document.removeEventListener("keydown", onKey, true);
    document.removeEventListener("pointerdown", onDocPointer, true);
    root.remove();
  };

  const onKey = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      remove();
    }
  };

  const onDocPointer = (ev: PointerEvent): void => {
    if (!root.contains(ev.target as Node)) {
      remove();
    }
  };

  setTimeout(() => {
    document.addEventListener("keydown", onKey, true);
    document.addEventListener("pointerdown", onDocPointer, true);
  }, 0);

  cancelBtn.addEventListener("click", remove);
  okBtn.addEventListener("click", () => {
    let nextForField: string[];
    if (checked.size === 0 || checked.size === distinct.length) {
      nextForField = [];
    } else {
      nextForField = distinct.filter((k) => checked.has(k));
    }
    const base = normalizeFilterSelectedKeys(def.filterFieldCols, def.filterSelectedKeys);
    const nextAll = base.map((k, i) => (i === filterFieldIndex ? nextForField : [...k]));
    flex.workspace.commands.execute(
      new UpdatePivotTableFiltersCommand(wb, pivotSheet, pivotDefId, nextAll),
    );
    flex.refresh();
    remove();
  });
}

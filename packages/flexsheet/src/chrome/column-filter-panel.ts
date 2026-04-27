import { AUTO_FILTER_BLANK_KEY } from "@flexsheet/core";
import { attachDraggableDialogPanel, columnIndexToLabel } from "@flexsheet/shared";

import type { FlexSheet } from "../flex-sheet.js";

let columnFilterStylesInjected = false;

function ensureColumnFilterStyles(): void {
  if (columnFilterStylesInjected) {
    return;
  }
  columnFilterStylesInjected = true;
  const style = document.createElement("style");
  style.textContent = `
.fs-col-filter {
  position: fixed;
  z-index: 10004;
  min-width: 220px;
  max-width: min(280px, calc(100vw - 24px));
  max-height: min(420px, calc(100vh - 24px));
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
.fs-col-filter__head-drag {
  flex-shrink: 0;
  padding: 6px 10px;
  font-size: 12px;
  font-weight: 600;
  color: #323130;
  background: #f3f2f1;
  border-bottom: 1px solid #edebe9;
  text-align: center;
}
.fs-col-filter__row {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 10px;
  cursor: pointer;
  user-select: none;
  border: none;
  background: #fff;
  width: 100%;
  text-align: left;
  font: inherit;
  color: inherit;
}
.fs-col-filter__row:disabled {
  opacity: 0.45;
  cursor: default;
}
.fs-col-filter__row:not(:disabled):hover,
.fs-col-filter__row:not(:disabled):focus-visible {
  background: #e8f5e9;
  outline: none;
}
.fs-col-filter__row--sub::after {
  content: "";
  margin-left: auto;
  width: 0;
  height: 0;
  border-top: 4px solid transparent;
  border-bottom: 4px solid transparent;
  border-left: 5px solid #605e5c;
}
.fs-col-filter__sep {
  height: 1px;
  background: #edebe9;
  margin: 2px 0;
}
.fs-col-filter__search {
  margin: 6px 10px;
  padding: 5px 8px;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  font: inherit;
  box-sizing: border-box;
}
.fs-col-filter__search:focus {
  outline: none;
  border-color: #217346;
}
.fs-col-filter__bulk {
  display: flex;
  gap: 12px;
  padding: 4px 10px 6px 10px;
}
.fs-col-filter__bulk button {
  border: none;
  background: none;
  padding: 0;
  font: inherit;
  color: #217346;
  cursor: pointer;
  text-decoration: none;
}
.fs-col-filter__bulk button:hover {
  text-decoration: underline;
}
.fs-col-filter__list {
  flex: 1;
  min-height: 120px;
  max-height: 200px;
  overflow: auto;
  margin: 0 8px 6px 8px;
  border: 1px solid #edebe9;
  border-radius: 2px;
  padding: 4px 0;
}
.fs-col-filter__item {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 8px;
  cursor: pointer;
  user-select: none;
}
.fs-col-filter__item:hover {
  background: #f3f2f1;
}
.fs-col-filter__footer {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding: 8px 10px 10px 10px;
  border-top: 1px solid #edebe9;
}
.fs-col-filter__btn {
  min-width: 64px;
  padding: 5px 12px;
  font: inherit;
  border-radius: 2px;
  cursor: pointer;
}
.fs-col-filter__btn--primary {
  border: none;
  background: #217346;
  color: #fff;
}
.fs-col-filter__btn--secondary {
  border: 1px solid #c8c6c4;
  background: #fff;
  color: #323130;
}
.fs-col-filter__submenu {
  position: fixed;
  z-index: 10005;
  min-width: 160px;
  background: #fff;
  border: 1px solid #c8c6c4;
  box-shadow: 0 4px 12px rgba(0,0,0,0.12);
  padding: 4px 0;
}
.fs-col-filter__submenu-title {
  padding: 6px 10px;
  font-weight: 600;
  font-size: 11px;
  color: #605e5c;
}
.fs-col-filter__swatch-row {
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  padding: 6px 10px;
  border: none;
  background: #fff;
  font: inherit;
  cursor: pointer;
  text-align: left;
}
.fs-col-filter__swatch-row:hover {
  background: #e8f5e9;
}
.fs-col-filter__swatch-row--auto {
  font-size: 10.5px;
  color: #605e5c;
  padding-top: 4px;
  padding-bottom: 4px;
}
.fs-col-filter__swatch {
  width: 28px;
  height: 18px;
  border: 1px solid #c8c6c4;
  flex-shrink: 0;
}
`;
  document.head.appendChild(style);
}

function argbToCssColor(argb: string): string {
  const t = argb.trim().toUpperCase();
  if (t.length === 8) {
    return `#${t.slice(2)}`;
  }
  if (t.length === 6) {
    return `#${t}`;
  }
  return "#000000";
}

function labelForFilterKey(key: string): string {
  return key === AUTO_FILTER_BLANK_KEY ? "(空)" : key;
}

export interface OpenColumnFilterPanelOptions {
  readonly flex: FlexSheet;
  readonly col: number;
  readonly clientX: number;
  readonly clientY: number;
}

export function openColumnFilterPanel(options: OpenColumnFilterPanelOptions): void {
  ensureColumnFilterStyles();
  const { flex, col } = options;
  const sheet = flex.workbook.getActiveSheet();
  if (sheet === undefined || !sheet.hasColumnAutoFilter(col)) {
    return;
  }
  const st0 = sheet.getColumnAutoFilterState(col);
  if (st0 === undefined) {
    return;
  }

  let draftChecked = new Set<string>(st0.checkedKeys);
  let draftIncludeBlank = st0.includeBlank;
  let draftFontArgb = st0.fontColorArgb;

  const root = document.createElement("div");
  root.className = "fs-col-filter";
  root.setAttribute("role", "dialog");
  root.setAttribute("aria-label", "列筛选");

  const colLetter = columnIndexToLabel(col);

  const headDrag = document.createElement("div");
  headDrag.className = "fs-col-filter__head-drag";
  headDrag.textContent = `列 ${colLetter} · 筛选`;
  headDrag.title = "按住可拖动";
  root.appendChild(headDrag);

  const addRow = (
    label: string,
    onClick: () => void,
    opts?: { readonly disabled?: boolean; readonly submenu?: boolean },
  ): HTMLButtonElement => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "fs-col-filter__row" + (opts?.submenu === true ? " fs-col-filter__row--sub" : "");
    btn.textContent = label;
    btn.disabled = opts?.disabled === true;
    btn.addEventListener("click", () => {
      if (btn.disabled) {
        return;
      }
      onClick();
    });
    root.appendChild(btn);
    return btn;
  };

  addRow("升序", () => {
    flex.sortActiveSheetByColumn(col, "asc");
    close();
  });
  addRow("降序", () => {
    flex.sortActiveSheetByColumn(col, "desc");
    close();
  });

  const sortColorBtn = addRow("按颜色排序", () => {}, { submenu: true });

  const clearLabel = `从 \"${colLetter}\" 中清除筛选`;
  addRow(clearLabel, () => {
    sheet.removeColumnAutoFilter(col);
    flex.refresh();
    close();
  });

  const filterColorBtn = addRow("按颜色筛选", () => {}, { submenu: true });

  const sep1 = document.createElement("div");
  sep1.className = "fs-col-filter__sep";
  root.appendChild(sep1);

  const search = document.createElement("input");
  search.type = "search";
  search.className = "fs-col-filter__search";
  search.placeholder = "搜索";
  search.autocomplete = "off";
  root.appendChild(search);

  const bulk = document.createElement("div");
  bulk.className = "fs-col-filter__bulk";
  const selAll = document.createElement("button");
  selAll.type = "button";
  selAll.textContent = "全选";
  const selNone = document.createElement("button");
  selNone.type = "button";
  selNone.textContent = "取消全选";
  bulk.appendChild(selAll);
  bulk.appendChild(selNone);
  root.appendChild(bulk);

  const listEl = document.createElement("div");
  listEl.className = "fs-col-filter__list";
  root.appendChild(listEl);

  const footer = document.createElement("div");
  footer.className = "fs-col-filter__footer";
  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.className = "fs-col-filter__btn fs-col-filter__btn--primary";
  okBtn.textContent = "确定";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "fs-col-filter__btn fs-col-filter__btn--secondary";
  cancelBtn.textContent = "取消";
  footer.appendChild(okBtn);
  footer.appendChild(cancelBtn);
  root.appendChild(footer);

  let submenuEl: HTMLDivElement | null = null;

  const closeSubmenu = (): void => {
    if (submenuEl !== null) {
      submenuEl.remove();
      submenuEl = null;
    }
  };

  const positionPanel = (): void => {
    const pad = 8;
    let left = options.clientX;
    let top = options.clientY;
    document.body.appendChild(root);
    const rw = root.offsetWidth;
    const rh = root.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    if (left + rw + pad > vw) {
      left = Math.max(pad, vw - rw - pad);
    }
    if (top + rh + pad > vh) {
      top = Math.max(pad, vh - rh - pad);
    }
    root.style.left = `${left}px`;
    root.style.top = `${top}px`;
  };

  const openSubmenu = (
    anchor: HTMLElement,
    title: string,
    build: (panel: HTMLDivElement) => void,
  ): void => {
    closeSubmenu();
    const sub = document.createElement("div");
    sub.className = "fs-col-filter__submenu";
    const t = document.createElement("div");
    t.className = "fs-col-filter__submenu-title";
    t.textContent = title;
    sub.appendChild(t);
    build(sub);
    document.body.appendChild(sub);
    submenuEl = sub;
    const ar = anchor.getBoundingClientRect();
    let sl = ar.right + 2;
    let st = ar.top;
    requestAnimationFrame(() => {
      if (submenuEl !== sub) {
        return;
      }
      const sw = sub.offsetWidth;
      const sh = sub.offsetHeight;
      const vw = window.innerWidth;
      const vh = window.innerHeight;
      if (sl + sw + 8 > vw) {
        sl = Math.max(8, ar.left - sw - 2);
      }
      if (st + sh + 8 > vh) {
        st = Math.max(8, vh - sh - 8);
      }
      sub.style.left = `${sl}px`;
      sub.style.top = `${st}px`;
    });
  };

  sortColorBtn.addEventListener("mouseenter", () => {
    const colors = sheet.collectUniqueFontColorArgbsInColumn(col);
    openSubmenu(sortColorBtn, "按字体颜色排序", (sub) => {
      for (const argb of colors) {
        const row = document.createElement("button");
        row.type = "button";
        row.className = "fs-col-filter__swatch-row";
        const sw = document.createElement("span");
        sw.className = "fs-col-filter__swatch";
        sw.style.background = argbToCssColor(argb);
        row.appendChild(sw);
        const lab = document.createElement("span");
        lab.textContent = argb;
        row.appendChild(lab);
        row.addEventListener("click", () => {
          flex.sortActiveSheetByColumnFontColor(col, argb, "asc");
          close();
        });
        sub.appendChild(row);
      }
      const autoRow = document.createElement("button");
      autoRow.type = "button";
      autoRow.className = "fs-col-filter__swatch-row fs-col-filter__swatch-row--auto";
      autoRow.textContent = "自动";
      autoRow.addEventListener("click", () => {
        flex.sortActiveSheetByColumnFontColor(col, null, "asc");
        close();
      });
      sub.appendChild(autoRow);
    });
  });

  filterColorBtn.addEventListener("mouseenter", () => {
    const colors = sheet.collectUniqueFontColorArgbsInColumn(col);
    openSubmenu(filterColorBtn, "按字体颜色筛选", (sub) => {
      for (const argb of colors) {
        const row = document.createElement("button");
        row.type = "button";
        row.className = "fs-col-filter__swatch-row";
        const sw = document.createElement("span");
        sw.className = "fs-col-filter__swatch";
        sw.style.background = argbToCssColor(argb);
        row.appendChild(sw);
        const lab = document.createElement("span");
        lab.textContent = argb;
        row.appendChild(lab);
        row.addEventListener("click", () => {
          draftFontArgb = argb;
          closeSubmenu();
        });
        sub.appendChild(row);
      }
      const autoRow = document.createElement("button");
      autoRow.type = "button";
      autoRow.className = "fs-col-filter__swatch-row fs-col-filter__swatch-row--auto";
      autoRow.textContent = "自动";
      autoRow.addEventListener("click", () => {
        draftFontArgb = null;
        closeSubmenu();
      });
      sub.appendChild(autoRow);
    });
  });

  root.addEventListener("mouseleave", (ev) => {
    const rel = ev.relatedTarget as Node | null;
    if (rel !== null && submenuEl !== null && submenuEl.contains(rel)) {
      return;
    }
    if (rel === null || (submenuEl !== null && !submenuEl.contains(rel))) {
      closeSubmenu();
    }
  });

  const uniqueKeys = sheet.collectUniqueAutoFilterKeysInColumn(col);

  const visibleKeys = (): string[] => {
    const q = search.value.trim().toLowerCase();
    if (q === "") {
      return uniqueKeys;
    }
    return uniqueKeys.filter((k) => labelForFilterKey(k).toLowerCase().includes(q));
  };

  const renderList = (): void => {
    listEl.replaceChildren();
    for (const key of visibleKeys()) {
      const row = document.createElement("label");
      row.className = "fs-col-filter__item";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      if (key === AUTO_FILTER_BLANK_KEY) {
        cb.checked = draftIncludeBlank;
      } else {
        cb.checked = draftChecked.has(key);
      }
      cb.addEventListener("change", () => {
        if (key === AUTO_FILTER_BLANK_KEY) {
          draftIncludeBlank = cb.checked;
        } else if (cb.checked) {
          draftChecked.add(key);
        } else {
          draftChecked.delete(key);
        }
      });
      const span = document.createElement("span");
      span.textContent = labelForFilterKey(key);
      row.appendChild(cb);
      row.appendChild(span);
      listEl.appendChild(row);
    }
  };

  search.addEventListener("input", () => {
    renderList();
  });

  selAll.addEventListener("click", () => {
    for (const key of visibleKeys()) {
      if (key === AUTO_FILTER_BLANK_KEY) {
        draftIncludeBlank = true;
      } else {
        draftChecked.add(key);
      }
    }
    renderList();
  });

  selNone.addEventListener("click", () => {
    for (const key of visibleKeys()) {
      if (key === AUTO_FILTER_BLANK_KEY) {
        draftIncludeBlank = false;
      } else {
        draftChecked.delete(key);
      }
    }
    renderList();
  });

  okBtn.addEventListener("click", () => {
    sheet.updateColumnAutoFilterValueSelection(col, draftChecked, draftIncludeBlank);
    sheet.setColumnAutoFilterFontColorArgb(col, draftFontArgb);
    flex.refresh();
    close();
  });

  cancelBtn.addEventListener("click", () => {
    close();
  });

  const onDocDown = (ev: PointerEvent): void => {
    if (root.contains(ev.target as Node)) {
      return;
    }
    if (submenuEl !== null && submenuEl.contains(ev.target as Node)) {
      return;
    }
    close();
  };

  const onKey = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      close();
    }
  };

  const close = (): void => {
    document.removeEventListener("pointerdown", onDocDown, true);
    document.removeEventListener("keydown", onKey, true);
    closeSubmenu();
    root.remove();
  };

  renderList();
  positionPanel();
  attachDraggableDialogPanel(root, headDrag);
  document.addEventListener("pointerdown", onDocDown, true);
  document.addEventListener("keydown", onKey, true);
  requestAnimationFrame(() => {
    search.focus();
  });
}

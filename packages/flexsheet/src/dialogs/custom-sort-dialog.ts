import {
  formatCellDisplayWithStyle,
  normalizeSelectionRange,
  type SelectionRange,
  type Worksheet,
  type WorksheetCustomSortLevel,
} from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";
import { attachDraggableDialogPanel, columnIndexToLabel, showMessageAlert } from "@flexsheet/shared";

import { ensureCustomSortDialogStyles } from "./fs-dialog-styles.js";

export interface OpenCustomSortDialogOptions {
  readonly sheet: Worksheet;
  readonly selectionRange: SelectionRange;
  /** 打开时第一行的默认「列」索引。 */
  readonly defaultSortCol: number;
  /** 确定并成功写入排序后调用（在 `recalcWorksheet` 之后）。 */
  readonly onAfterApply?: () => void;
  readonly onClose: () => void;
}

type SortOn = WorksheetCustomSortLevel["sortOn"];
type OrderOpt = "asc" | "desc";

function columnOptionLabel(
  sheet: Worksheet,
  col: number,
  headerRow: number,
  hasHeaders: boolean,
): string {
  const letter = columnIndexToLabel(col);
  if (!hasHeaders) {
    return `列 ${letter}`;
  }
  const ap = sheet.getMergeAnchorCell(headerRow, col);
  const cell = sheet.getCell(ap.row, ap.col);
  const t = formatCellDisplayWithStyle(cell.value, cell.style).trim();
  if (t.length > 0) {
    const short = t.length > 14 ? `${t.slice(0, 14)}…` : t;
    return `列 ${letter}（${short}）`;
  }
  return `列 ${letter}`;
}

interface RowBinding {
  readonly tr: HTMLTableRowElement;
  readonly colSel: HTMLSelectElement;
  readonly sortOnSel: HTMLSelectElement;
  readonly orderSel: HTMLSelectElement;
  readonly colorSel: HTMLSelectElement;
}

export function openCustomSortDialogWithOverlay(options: OpenCustomSortDialogOptions): HTMLDivElement {
  ensureCustomSortDialogStyles();
  const { sheet, onClose, onAfterApply, defaultSortCol, selectionRange: selRaw } = options;
  const n = normalizeSelectionRange(selRaw);
  const colCount = sheet.colCount;
  /** 参与排序关键字可选的列：仅当前选区内的列，而非整张表。 */
  const selectionColMin = colCount > 0 ? Math.max(0, Math.min(colCount - 1, Math.min(n.startCol, n.endCol))) : 0;
  const selectionColMax = colCount > 0 ? Math.max(0, Math.min(colCount - 1, Math.max(n.startCol, n.endCol))) : 0;
  const clampColToSelection = (c: number): number => {
    if (colCount <= 0) {
      return 0;
    }
    return Math.max(selectionColMin, Math.min(selectionColMax, Math.trunc(c)));
  };
  const safeDefaultCol = clampColToSelection(defaultSortCol);
  const headerRow = n.startRow;

  const overlay = document.createElement("div");
  overlay.className = "fs-custom-sort-overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-custom-sort";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-custom-sort-title");
  panel.tabIndex = -1;

  const titlebar = document.createElement("div");
  titlebar.className = "fs-custom-sort__titlebar";
  const headTitle = document.createElement("div");
  headTitle.id = "fs-custom-sort-title";
  headTitle.className = "fs-custom-sort__head-title";
  headTitle.textContent = "排序";
  titlebar.appendChild(headTitle);

  const top = document.createElement("div");
  top.className = "fs-custom-sort__top";
  const hint = document.createElement("div");
  hint.className = "fs-custom-sort__top-hint";
  hint.textContent = "添加按以下方式排序的级别：";
  const hasHeaderLabel = document.createElement("label");
  hasHeaderLabel.className = "fs-custom-sort__head-check";
  const hasHeaderCb = document.createElement("input");
  hasHeaderCb.type = "checkbox";
  hasHeaderCb.checked = true;
  const hasHeaderText = document.createElement("span");
  hasHeaderText.textContent = "列表包含标题";
  hasHeaderLabel.appendChild(hasHeaderCb);
  hasHeaderLabel.appendChild(hasHeaderText);
  top.appendChild(hint);
  top.appendChild(hasHeaderLabel);

  const tableWrap = document.createElement("div");
  tableWrap.className = "fs-custom-sort__table-wrap";
  const table = document.createElement("table");
  table.className = "fs-custom-sort__table";
  const thead = document.createElement("thead");
  const thr = document.createElement("tr");
  for (const t of [
    { w: "", c: "fs-custom-sort__th" },
    { w: "列", c: "fs-custom-sort__th" },
    { w: "排序依据", c: "fs-custom-sort__th" },
    { w: "顺序", c: "fs-custom-sort__th" },
    { w: "颜色 / 图标", c: "fs-custom-sort__th" },
  ]) {
    const th = document.createElement("th");
    th.className = t.c;
    th.textContent = t.w;
    thr.appendChild(th);
  }
  thead.appendChild(thr);
  const tbody = document.createElement("tbody");
  table.appendChild(thead);
  table.appendChild(tbody);
  tableWrap.appendChild(table);

  const foot = document.createElement("div");
  foot.className = "fs-custom-sort__foot";
  const levelBtns = document.createElement("div");
  levelBtns.className = "fs-custom-sort__level-btns";
  const addBtn = document.createElement("button");
  addBtn.type = "button";
  addBtn.className = "fs-custom-sort__level-btn";
  addBtn.textContent = "+";
  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "fs-custom-sort__level-btn";
  removeBtn.textContent = "—";
  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "fs-custom-sort__level-btn";
  copyBtn.textContent = "复制";
  levelBtns.appendChild(addBtn);
  levelBtns.appendChild(removeBtn);
  levelBtns.appendChild(copyBtn);
  const actions = document.createElement("div");
  actions.className = "fs-custom-sort__actions";
  const optBtn = document.createElement("button");
  optBtn.type = "button";
  optBtn.className = "fs-custom-sort__btn";
  optBtn.textContent = "选项…";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "fs-custom-sort__btn";
  cancelBtn.textContent = "取消";
  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.className = "fs-custom-sort__btn fs-custom-sort__btn--primary";
  okBtn.textContent = "确定";
  actions.appendChild(optBtn);
  actions.appendChild(cancelBtn);
  actions.appendChild(okBtn);
  foot.appendChild(levelBtns);
  foot.appendChild(actions);

  /** 主对话框「确定」时读取；选项子弹窗中编辑，确定后写回。 */
  const sortOptionState: {
    direction: "column" | "row";
    method: "alphabet" | "stroke";
    caseSensitive: boolean;
  } = {
    direction: "column",
    method: "alphabet",
    caseSensitive: false,
  };

  const optionsScrim = document.createElement("div");
  optionsScrim.className = "fs-custom-sort__options-scrim";
  optionsScrim.hidden = true;
  optionsScrim.setAttribute("role", "dialog");
  optionsScrim.setAttribute("aria-modal", "true");
  optionsScrim.setAttribute("aria-label", "排序选项");

  const optionsPanel = document.createElement("div");
  optionsPanel.className = "fs-custom-sort__options-panel";
  const optionsDragTitle = document.createElement("div");
  optionsDragTitle.className = "fs-custom-sort__options-drag-title";
  optionsDragTitle.textContent = "排序选项";
  optionsDragTitle.title = "按住可拖动";
  const optionsGrid = document.createElement("div");
  optionsGrid.className = "fs-custom-sort__options-grid";

  const gDir = document.createElement("div");
  gDir.setAttribute("role", "group");
  gDir.setAttribute("aria-label", "方向");
  const tDir = document.createElement("div");
  tDir.className = "fs-custom-sort__options-group-title";
  tDir.textContent = "方向";
  gDir.appendChild(tDir);
  const rByCol = document.createElement("input");
  rByCol.type = "radio";
  rByCol.name = "fs-csdirection";
  rByCol.value = "column";
  const lByCol = document.createElement("label");
  lByCol.className = "fs-custom-sort__options-radio";
  lByCol.appendChild(rByCol);
  lByCol.appendChild(document.createTextNode("按列排序"));
  const rByRow = document.createElement("input");
  rByRow.type = "radio";
  rByRow.name = "fs-csdirection";
  rByRow.value = "row";
  const lByRow = document.createElement("label");
  lByRow.className = "fs-custom-sort__options-radio";
  lByRow.appendChild(rByRow);
  lByRow.appendChild(document.createTextNode("按行排序"));
  gDir.appendChild(lByCol);
  gDir.appendChild(lByRow);

  const gMethod = document.createElement("div");
  gMethod.setAttribute("role", "group");
  gMethod.setAttribute("aria-label", "方法");
  const tMethod = document.createElement("div");
  tMethod.className = "fs-custom-sort__options-group-title";
  tMethod.textContent = "方法";
  gMethod.appendChild(tMethod);
  const rAlpha = document.createElement("input");
  rAlpha.type = "radio";
  rAlpha.name = "fs-csmethod";
  rAlpha.value = "alphabet";
  const lAlpha = document.createElement("label");
  lAlpha.className = "fs-custom-sort__options-radio";
  lAlpha.appendChild(rAlpha);
  lAlpha.appendChild(document.createTextNode("字母排序"));
  const rStroke = document.createElement("input");
  rStroke.type = "radio";
  rStroke.name = "fs-csmethod";
  rStroke.value = "stroke";
  const lStroke = document.createElement("label");
  lStroke.className = "fs-custom-sort__options-radio";
  lStroke.appendChild(rStroke);
  lStroke.appendChild(document.createTextNode("笔划排序"));
  gMethod.appendChild(lAlpha);
  gMethod.appendChild(lStroke);
  optionsGrid.appendChild(gDir);
  optionsGrid.appendChild(gMethod);
  const caseRow = document.createElement("div");
  const caseCb = document.createElement("input");
  caseCb.type = "checkbox";
  caseCb.id = "fs-cs-case";
  const caseLabel = document.createElement("label");
  caseLabel.htmlFor = "fs-cs-case";
  caseLabel.className = "fs-custom-sort__options-check";
  caseLabel.appendChild(caseCb);
  caseLabel.appendChild(document.createTextNode("区分大小写"));
  caseRow.appendChild(caseLabel);
  const optActions = document.createElement("div");
  optActions.className = "fs-custom-sort__options-actions";
  const optCancel = document.createElement("button");
  optCancel.type = "button";
  optCancel.className = "fs-custom-sort__btn";
  optCancel.textContent = "取消";
  const optOk = document.createElement("button");
  optOk.type = "button";
  optOk.className = "fs-custom-sort__btn fs-custom-sort__btn--primary";
  optOk.textContent = "确定";
  optActions.appendChild(optCancel);
  optActions.appendChild(optOk);
  optionsPanel.appendChild(optionsDragTitle);
  optionsPanel.appendChild(optionsGrid);
  optionsPanel.appendChild(caseRow);
  optionsPanel.appendChild(optActions);
  optionsScrim.appendChild(optionsPanel);

  const syncOptionsUiFromState = (): void => {
    rByCol.checked = sortOptionState.direction === "column";
    rByRow.checked = sortOptionState.direction === "row";
    rAlpha.checked = sortOptionState.method === "alphabet";
    rStroke.checked = sortOptionState.method === "stroke";
    caseCb.checked = sortOptionState.caseSensitive;
  };
  const applyOptionsFromUi = (): void => {
    sortOptionState.direction = rByRow.checked ? "row" : "column";
    sortOptionState.method = rStroke.checked ? "stroke" : "alphabet";
    sortOptionState.caseSensitive = caseCb.checked;
  };
  const closeOptionsPanel = (): void => {
    optionsScrim.hidden = true;
    actions.hidden = false;
  };
  const openOptionsPanel = (): void => {
    syncOptionsUiFromState();
    optionsScrim.hidden = false;
    actions.hidden = true;
  };
  optCancel.addEventListener("click", () => {
    syncOptionsUiFromState();
    closeOptionsPanel();
  });
  optOk.addEventListener("click", () => {
    applyOptionsFromUi();
    closeOptionsPanel();
  });
  optionsScrim.addEventListener("pointerdown", (ev) => {
    if (ev.target === optionsScrim) {
      ev.preventDefault();
      syncOptionsUiFromState();
      closeOptionsPanel();
    }
  });
  optionsPanel.addEventListener("pointerdown", (ev) => {
    ev.stopPropagation();
  });
  syncOptionsUiFromState();

  panel.appendChild(titlebar);
  panel.appendChild(top);
  panel.appendChild(tableWrap);
  panel.appendChild(foot);
  overlay.appendChild(panel);
  overlay.appendChild(optionsScrim);
  attachDraggableDialogPanel(panel, titlebar);
  attachDraggableDialogPanel(optionsPanel, optionsDragTitle);

  let rowBindings: RowBinding[] = [];
  let selectedIndex = 0;

  const getSortDataBounds = (): { start: number; end: number } | null => {
    const hasH = hasHeaderCb.checked;
    const sortStart = hasH ? n.startRow + 1 : n.startRow;
    const sortEnd = n.endRow;
    if (sortStart > sortEnd) {
      return null;
    }
    return { start: sortStart, end: sortEnd };
  };

  const fillOrderSelect = (sel: HTMLSelectElement, current: OrderOpt): void => {
    sel.innerHTML = "";
    const o1 = document.createElement("option");
    o1.value = "asc";
    o1.textContent = "升序";
    const o2 = document.createElement("option");
    o2.value = "desc";
    o2.textContent = "降序";
    sel.appendChild(o1);
    sel.appendChild(o2);
    sel.value = current;
  };

  const fillColSelect = (sel: HTMLSelectElement, current: number): void => {
    const hasH = hasHeaderCb.checked;
    sel.innerHTML = "";
    for (let c = selectionColMin; c <= selectionColMax; c++) {
      const o = document.createElement("option");
      o.value = String(c);
      o.textContent = columnOptionLabel(sheet, c, headerRow, hasH);
      sel.appendChild(o);
    }
    sel.value = String(clampColToSelection(current));
  };

  const readSortOn = (s: string): SortOn => {
    if (s === "fontColor" || s === "fillColor" || s === "cellIcon" || s === "value") {
      return s;
    }
    return "value";
  };

  const fillColorSelect = (row: RowBinding): void => {
    const b = getSortDataBounds();
    if (b === null) {
      row.colorSel.innerHTML = "";
      row.colorSel.disabled = true;
      return;
    }
    const col = Number(row.colSel.value);
    if (!Number.isInteger(col) || col < 0) {
      row.colorSel.innerHTML = "";
      row.colorSel.disabled = true;
      return;
    }
    const on = readSortOn(row.sortOnSel.value);
    if (on === "value" || on === "cellIcon") {
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "—";
      row.colorSel.innerHTML = "";
      row.colorSel.appendChild(opt);
      row.colorSel.value = "";
      row.colorSel.disabled = true;
      return;
    }
    const arbs =
      on === "fontColor"
        ? sheet.collectUniqueFontColorArgbsInRowRange(col, b.start, b.end)
        : sheet.collectUniqueFillColorArgbsInRowRange(col, b.start, b.end);
    const prev = row.colorSel.value;
    row.colorSel.innerHTML = "";
    const auto = document.createElement("option");
    auto.value = "";
    auto.textContent = "（自动）";
    row.colorSel.appendChild(auto);
    for (const a of arbs) {
      const o = document.createElement("option");
      o.value = a;
      o.textContent = a;
      row.colorSel.appendChild(o);
    }
    row.colorSel.disabled = false;
    if (arbs.includes(prev) || prev === "") {
      row.colorSel.value = arbs.includes(prev) ? prev : "";
    } else {
      row.colorSel.value = "";
    }
  };

  const setRowLabelText = (row: RowBinding, levelIndex: number): void => {
    const first = row.tr.querySelector(".fs-custom-sort__td--label");
    if (first !== null) {
      first.textContent = levelIndex === 0 ? "排序依据" : "次要依据";
    }
  };

  const applySelectedStyle = (): void => {
    for (let i = 0; i < rowBindings.length; i++) {
      if (i === selectedIndex) {
        rowBindings[i]!.tr.classList.add("fs-custom-sort__tr--selected");
      } else {
        rowBindings[i]!.tr.classList.remove("fs-custom-sort__tr--selected");
      }
    }
  };

  const refreshAllColorSelects = (): void => {
    for (const r of rowBindings) {
      fillColorSelect(r);
    }
  };

  const rebuildColLabels = (): void => {
    for (const r of rowBindings) {
      const cur = Number(r.colSel.value);
      fillColSelect(r.colSel, Number.isInteger(cur) ? cur : selectionColMin);
    }
  };

  const createRow = (initial: {
    col: number;
    sortOn: SortOn;
    order: OrderOpt;
  }): RowBinding => {
    const tr = document.createElement("tr");
    tr.className = "fs-custom-sort__tr";
    const td0 = document.createElement("td");
    td0.className = "fs-custom-sort__td fs-custom-sort__td--label";
    const td1 = document.createElement("td");
    td1.className = "fs-custom-sort__td";
    const td2 = document.createElement("td");
    td2.className = "fs-custom-sort__td";
    const td3 = document.createElement("td");
    td3.className = "fs-custom-sort__td";
    const td4 = document.createElement("td");
    td4.className = "fs-custom-sort__td";
    const colSel = document.createElement("select");
    colSel.className = "fs-custom-sort__select";
    const sortOnSel = document.createElement("select");
    sortOnSel.className = "fs-custom-sort__select";
    for (const o of [
      { v: "value", t: "值" },
      { v: "fillColor", t: "单元格颜色" },
      { v: "fontColor", t: "字体颜色" },
      { v: "cellIcon", t: "单元格图标" },
    ]) {
      const op = document.createElement("option");
      op.value = o.v;
      op.textContent = o.t;
      sortOnSel.appendChild(op);
    }
    sortOnSel.value = initial.sortOn;
    const orderSel = document.createElement("select");
    orderSel.className = "fs-custom-sort__select";
    fillOrderSelect(orderSel, initial.order);
    const colorSel = document.createElement("select");
    colorSel.className = "fs-custom-sort__select";
    fillColSelect(colSel, initial.col);
    td1.appendChild(colSel);
    td2.appendChild(sortOnSel);
    td3.appendChild(orderSel);
    td4.appendChild(colorSel);
    tr.appendChild(td0);
    tr.appendChild(td1);
    tr.appendChild(td2);
    tr.appendChild(td3);
    tr.appendChild(td4);
    const binding: RowBinding = { tr, colSel, sortOnSel, orderSel, colorSel };
    const rowIndex = (): number => rowBindings.indexOf(binding);

    tr.addEventListener("mousedown", (e) => {
      if (e.target instanceof HTMLSelectElement) {
        return;
      }
      const idx = rowIndex();
      if (idx >= 0) {
        selectedIndex = idx;
        applySelectedStyle();
      }
    });
    const onField = (): void => {
      fillColorSelect(binding);
    };
    colSel.addEventListener("change", onField);
    sortOnSel.addEventListener("change", onField);
    return binding;
  };

  const syncRowLabels = (): void => {
    for (let i = 0; i < rowBindings.length; i++) {
      setRowLabelText(rowBindings[i]!, i);
    }
  };

  const firstRow = createRow({ col: safeDefaultCol, sortOn: "value", order: "desc" });
  rowBindings.push(firstRow);
  tbody.appendChild(firstRow.tr);
  fillColorSelect(firstRow);
  selectedIndex = 0;
  applySelectedStyle();
  syncRowLabels();

  hasHeaderCb.addEventListener("change", () => {
    rebuildColLabels();
    refreshAllColorSelects();
  });

  const updateLevelButtonState = (): void => {
    removeBtn.disabled = rowBindings.length <= 1;
  };
  updateLevelButtonState();

  addBtn.addEventListener("click", () => {
    const last = rowBindings[rowBindings.length - 1]!;
    const nextCol = clampColToSelection(Number(last.colSel.value));
    const nr = createRow({
      col: nextCol,
      sortOn: readSortOn(last.sortOnSel.value),
      order: "desc",
    });
    rowBindings.push(nr);
    tbody.appendChild(nr.tr);
    fillColorSelect(nr);
    selectedIndex = rowBindings.length - 1;
    applySelectedStyle();
    syncRowLabels();
    updateLevelButtonState();
  });

  removeBtn.addEventListener("click", () => {
    if (rowBindings.length <= 1) {
      return;
    }
    const i = Math.max(0, Math.min(selectedIndex, rowBindings.length - 1));
    const rm = rowBindings[i]!;
    rm.tr.remove();
    rowBindings = rowBindings.filter((_, j) => j !== i);
    if (selectedIndex >= rowBindings.length) {
      selectedIndex = rowBindings.length - 1;
    }
    applySelectedStyle();
    syncRowLabels();
    updateLevelButtonState();
  });

  copyBtn.addEventListener("click", () => {
    const src = rowBindings[selectedIndex] ?? rowBindings[0]!;
    const srcIndex = rowBindings.indexOf(src);
    const c = clampColToSelection(Number(src.colSel.value));
    const so = readSortOn(src.sortOnSel.value);
    const ord: OrderOpt = src.orderSel.value === "desc" ? "desc" : "asc";
    const colorVal = src.colorSel.value;
    const nr = createRow({ col: c, sortOn: so, order: ord });
    rowBindings.splice(srcIndex + 1, 0, nr);
    src.tr.insertAdjacentElement("afterend", nr.tr);
    nr.colSel.value = String(c);
    nr.sortOnSel.value = so;
    fillOrderSelect(nr.orderSel, ord);
    fillColorSelect(nr);
    if (colorVal !== "" && (so === "fontColor" || so === "fillColor")) {
      const has = Array.from(nr.colorSel.options).some((o) => o.value === colorVal);
      if (has) {
        nr.colorSel.value = colorVal;
      }
    }
    selectedIndex = srcIndex + 1;
    applySelectedStyle();
    syncRowLabels();
    updateLevelButtonState();
  });

  const dismiss = (): void => {
    onClose();
  };

  const tryApply = (): void => {
    if (sortOptionState.direction === "row") {
      showMessageAlert("按行排序功能尚未实现。");
      return;
    }
    if (sortOptionState.method === "stroke") {
      showMessageAlert("按笔划排序功能尚未实现。");
      return;
    }
    const b = getSortDataBounds();
    if (b === null) {
      showMessageAlert("数据行不足，无法排序。请扩大选区，或取消勾选「列表包含标题」。");
      return;
    }
    const levels: WorksheetCustomSortLevel[] = [];
    for (const r of rowBindings) {
      const c = Math.trunc(Number(r.colSel.value));
      if (!Number.isInteger(c) || c < selectionColMin || c > selectionColMax || c < 0 || c >= colCount) {
        showMessageAlert("请选择有效的列。");
        return;
      }
      const so = readSortOn(r.sortOnSel.value);
      const dir: "asc" | "desc" = r.orderSel.value === "desc" ? "desc" : "asc";
      let colorTarget: string | null = null;
      if (so === "fontColor" || so === "fillColor") {
        const raw = r.colorSel.value.trim();
        colorTarget = raw === "" ? null : raw.toUpperCase();
      }
      levels.push({ col: c, sortOn: so, direction: dir, colorTargetArgb: colorTarget });
    }
    sheet.sortRowsInRangeByCustomSortLevels(b.start, b.end, levels);
    recalcWorksheet(sheet);
    onAfterApply?.();
    dismiss();
  };

  const onOverlayDown = (ev: PointerEvent): void => {
    if (!optionsScrim.hidden) {
      return;
    }
    if (ev.target === overlay) {
      dismiss();
    }
  };
  overlay.addEventListener("pointerdown", onOverlayDown);
  cancelBtn.addEventListener("click", dismiss);
  okBtn.addEventListener("click", tryApply);
  optBtn.addEventListener("click", () => {
    openOptionsPanel();
  });

  const onKey = (kev: KeyboardEvent): void => {
    if (kev.key === "Escape") {
      kev.preventDefault();
      kev.stopPropagation();
      if (!optionsScrim.hidden) {
        syncOptionsUiFromState();
        closeOptionsPanel();
        return;
      }
      dismiss();
    }
  };
  overlay.addEventListener("keydown", onKey);
  panel.addEventListener("keydown", onKey);

  requestAnimationFrame(() => {
    firstRow.colSel.focus();
  });

  return overlay;
}

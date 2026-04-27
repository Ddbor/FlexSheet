import { formatCellDisplayWithStyle, type Workbook, type Worksheet } from "@flexsheet/core";
import { attachDraggableDialogPanel, columnIndexToLabel, showMessageAlert } from "@flexsheet/shared";
import { ensureFindReplaceDialogStyles } from "./fs-dialog-styles.js";
import {
  type FindHit,
  type FindLookIn,
  type FindReplaceScanOptions,
  type FindSearchOrder,
  type FindWithinScope,
  applyReplaceAllWithWriter,
  buildReplacedString,
  collectFindHits,
  firstHitIndexAtOrAfter,
  getHayStringForCell,
  isHitReplaceable,
} from "./find-replace-engine.js";

export interface OpenFindReplaceDialogOptions {
  readonly workbook: Workbook;
  readonly initialTab: "find" | "replace";
  readonly onNavigateToHit: (hit: Readonly<FindHit>) => void;
  readonly setCellValueCommand: (sheet: Worksheet, row: number, col: number, value: string) => void;
  readonly onDataChanged: () => void;
  readonly onClose: () => void;
}

const MAX_HISTORY = 10;

function readScanOptions(get: {
  find: string;
  within: FindWithinScope;
  search: FindSearchOrder;
  lookIn: FindLookIn;
  matchCase: boolean;
  matchEntire: boolean;
  distWidth: boolean;
}): FindReplaceScanOptions {
  return {
    find: get.find,
    within: get.within,
    search: get.search,
    lookIn: get.lookIn,
    matchCase: get.matchCase,
    matchEntireCell: get.matchEntire,
    distinguishWidth: get.distWidth,
  };
}

function hitToAbsDollarRef(hit: Readonly<FindHit>): string {
  return `$${columnIndexToLabel(hit.col)}$${hit.row + 1}`;
}

function updateHistory(arr: string[], value: string): void {
  const t = value.trim();
  if (t === "") {
    return;
  }
  const i = arr.indexOf(t);
  if (i >= 0) {
    arr.splice(i, 1);
  }
  arr.unshift(t);
  while (arr.length > MAX_HISTORY) {
    arr.pop();
  }
}

function syncDatalist(dl: HTMLDataListElement, arr: readonly string[]): void {
  while (dl.firstChild !== null) {
    dl.removeChild(dl.firstChild);
  }
  for (const x of arr) {
    const o = document.createElement("option");
    o.value = x;
    dl.appendChild(o);
  }
}

function createInputRow(labelText: string, placeholder: string, className: string): {
  readonly row: HTMLDivElement;
  readonly input: HTMLInputElement;
  readonly list: HTMLDataListElement;
} {
  const row = document.createElement("div");
  row.className = "fs-fr__row";
  const lab = document.createElement("label");
  lab.textContent = labelText;
  const wrap = document.createElement("div");
  wrap.className = "fs-fr__input-wrap";
  const input = document.createElement("input");
  input.type = "text";
  input.className = className;
  input.autocomplete = "off";
  input.placeholder = placeholder;
  const list = document.createElement("datalist");
  list.id = `fs-fr-dl-${Math.random().toString(36).slice(2, 9)}`;
  input.setAttribute("list", list.id);
  /* 不另加右侧按钮：与 <datalist> 的原生下拉提示重复，会出现双箭头。 */

  wrap.appendChild(input);
  row.appendChild(lab);
  row.appendChild(wrap);
  row.appendChild(list);
  return { row, input, list };
}

function createCheck(label: string): { readonly wrap: HTMLLabelElement; readonly input: HTMLInputElement } {
  const wrap = document.createElement("label");
  const input = document.createElement("input");
  input.type = "checkbox";
  const span = document.createElement("span");
  span.textContent = label;
  wrap.appendChild(input);
  wrap.appendChild(span);
  return { wrap, input };
}

function createSelectRow(label: string, select: HTMLSelectElement): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "fs-fr__sel-row";
  const span = document.createElement("span");
  span.textContent = label;
  row.appendChild(span);
  row.appendChild(select);
  return row;
}

function createButton(text: string, kind: "go" | "secondary"): HTMLButtonElement {
  const b = document.createElement("button");
  b.type = "button";
  b.className = kind === "go" ? "fs-fr__btn fs-fr__btn--go" : "fs-fr__btn";
  b.textContent = text;
  return b;
}

export function openFindReplaceDialogWithOverlay(options: OpenFindReplaceDialogOptions): HTMLDivElement {
  ensureFindReplaceDialogStyles();
  const { workbook, onClose, onDataChanged, onNavigateToHit, setCellValueCommand, initialTab } = options;

  const overlay = document.createElement("div");
  overlay.className = "fs-fr-overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-fr";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-fr-title");
  panel.tabIndex = -1;
  overlay.appendChild(panel);

  const header = document.createElement("div");
  header.className = "fs-fr__header";
  const title = document.createElement("div");
  title.id = "fs-fr-title";
  title.className = "fs-fr__title";
  title.textContent = "查找和替换";
  title.title = "按住可拖动";
  header.appendChild(title);
  panel.appendChild(header);

  const tabsRow = document.createElement("div");
  tabsRow.className = "fs-fr__tabs";
  const tabFind = document.createElement("button");
  tabFind.type = "button";
  tabFind.className = "fs-fr__tab";
  tabFind.setAttribute("role", "tab");
  tabFind.textContent = "查找";
  const tabReplace = document.createElement("button");
  tabReplace.type = "button";
  tabReplace.className = "fs-fr__tab";
  tabReplace.setAttribute("role", "tab");
  tabReplace.textContent = "替换";
  tabsRow.appendChild(tabFind);
  tabsRow.appendChild(tabReplace);
  panel.appendChild(tabsRow);

  const body = document.createElement("div");
  body.className = "fs-fr__body";
  panel.appendChild(body);
  const panelInner = document.createElement("div");
  panelInner.className = "fs-fr__panel";
  body.appendChild(panelInner);

  const findRow = createInputRow("查找内容(K):", "查找内容", "fs-fr__input");
  const replaceRow = createInputRow("替换为(E):", "替换为", "fs-fr__input fs-fr__input--secondary");
  panelInner.appendChild(findRow.row);
  panelInner.appendChild(replaceRow.row);

  const optionToggle = document.createElement("button");
  optionToggle.type = "button";
  optionToggle.className = "fs-fr__opt-toggle";
  optionToggle.textContent = "选项";
  panelInner.appendChild(optionToggle);

  const optionsWrap = document.createElement("div");
  optionsWrap.hidden = true;
  panelInner.appendChild(optionsWrap);

  const withinSel = document.createElement("select");
  withinSel.className = "fs-fr__select fs-fr__select--accent";
  withinSel.innerHTML = "<option value='sheet'>工作表</option><option value='workbook'>工作簿</option>";
  const searchSel = document.createElement("select");
  searchSel.className = "fs-fr__select";
  searchSel.innerHTML = "<option value='row'>按行</option><option value='column'>按列</option>";
  const lookInSel = document.createElement("select");
  lookInSel.className = "fs-fr__select fs-fr__select--accent";
  lookInSel.innerHTML =
    "<option value='values' selected>值</option><option value='formulas'>公式</option><option value='comments'>批注</option>";

  const checkCase = createCheck("区分大小写(H)");
  const checkEntire = createCheck("单元格匹配(C)");
  const checkWidth = createCheck("区分全/半角(S)");

  const left = document.createElement("div");
  left.className = "fs-fr__left-col";
  left.appendChild(createSelectRow("范围(W):", withinSel));
  left.appendChild(createSelectRow("搜索(S):", searchSel));
  left.appendChild(createSelectRow("查找范围(L):", lookInSel));

  const right = document.createElement("div");
  right.className = "fs-fr__checks";
  right.appendChild(checkCase.wrap);
  right.appendChild(checkEntire.wrap);
  right.appendChild(checkWidth.wrap);

  const optGrid = document.createElement("div");
  optGrid.className = "fs-fr__opt-grid";
  optGrid.appendChild(left);
  optGrid.appendChild(right);
  optionsWrap.appendChild(optGrid);

  const WORKBOOK_LIST_LABEL = "工作簿1";
  const resultBox = document.createElement("div");
  resultBox.className = "fs-fr__result-box";
  resultBox.setAttribute("data-visible", "0");
  const resultScroll = document.createElement("div");
  resultScroll.className = "fs-fr__result-scroll";
  const resultTable = document.createElement("table");
  resultTable.className = "fs-fr__result-table";
  const resultThead = document.createElement("thead");
  const headTr = document.createElement("tr");
  for (const t of ["工作簿", "工作表", "名称", "单元格", "值", "公式"] as const) {
    const th = document.createElement("th");
    th.textContent = t;
    headTr.appendChild(th);
  }
  resultThead.appendChild(headTr);
  const resultTbody = document.createElement("tbody");
  resultTable.appendChild(resultThead);
  resultTable.appendChild(resultTbody);
  resultScroll.appendChild(resultTable);
  const resultStatus = document.createElement("div");
  resultStatus.className = "fs-fr__result-status";
  resultStatus.textContent = "";
  resultBox.appendChild(resultScroll);
  resultBox.appendChild(resultStatus);
  panelInner.appendChild(resultBox);

  const footer = document.createElement("div");
  footer.className = "fs-fr__foot";
  panel.appendChild(footer);
  const footerFind = document.createElement("div");
  footerFind.style.cssText =
    "display:flex;flex-wrap:wrap;align-items:center;justify-content:center;gap:6px";
  const footerReplace = document.createElement("div");
  footerReplace.style.cssText =
    "display:none;flex-wrap:wrap;align-items:center;justify-content:center;gap:6px";
  footer.appendChild(footerFind);
  footer.appendChild(footerReplace);

  const findAllFindBtn = createButton("查找全部", "secondary");
  const prevFindBtn = createButton("上一步", "go");
  const nextFindBtn = createButton("下一步", "go");
  const closeFindBtn = createButton("关闭", "secondary");
  footerFind.append(findAllFindBtn, prevFindBtn, nextFindBtn, closeFindBtn);

  const replaceBtn = createButton("替换", "secondary");
  const replaceAllBtn = createButton("全部替换", "secondary");
  const findAllReplaceBtn = createButton("查找全部", "secondary");
  const prevReplaceBtn = createButton("上一步", "go");
  const nextReplaceBtn = createButton("下一步", "go");
  const closeReplaceBtn = createButton("关闭", "secondary");
  footerReplace.append(
    replaceBtn,
    replaceAllBtn,
    findAllReplaceBtn,
    prevReplaceBtn,
    nextReplaceBtn,
    closeReplaceBtn,
  );

  let currentTab: "find" | "replace" = initialTab;
  let hits: FindHit[] = [];
  let cursor = -1;
  const findHistory: string[] = [];
  const replaceHistory: string[] = [];

  const dismiss = (): void => {
    overlay.remove();
    onClose();
  };

  const getScanOptions = (): FindReplaceScanOptions =>
    readScanOptions({
      find: findRow.input.value,
      within: (withinSel.value as FindWithinScope) || "sheet",
      search: (searchSel.value as FindSearchOrder) || "row",
      lookIn: (lookInSel.value as FindLookIn) || "values",
      matchCase: checkCase.input.checked,
      matchEntire: checkEntire.input.checked,
      distWidth: checkWidth.input.checked,
    });

  const rebuildHits = (): void => {
    hits = collectFindHits(workbook, getScanOptions());
  };

  const clearFindResultTable = (): void => {
    resultTbody.innerHTML = "";
    resultStatus.textContent = "";
    resultBox.setAttribute("data-visible", "0");
  };

  const syncFindResultSelection = (): void => {
    const trs = resultTbody.querySelectorAll("tr");
    trs.forEach((tr, i) => {
      tr.classList.toggle("fs-fr__result-row--sel", i === cursor && cursor >= 0);
    });
  };

  const navigateToHit = (idx: number): void => {
    if (idx < 0 || idx >= hits.length) {
      return;
    }
    cursor = idx;
    onNavigateToHit(hits[idx]!);
    if (resultBox.getAttribute("data-visible") === "1") {
      syncFindResultSelection();
    }
  };

  const onFindAll = (): void => {
    const opt = getScanOptions();
    if (opt.find.trim() === "") {
      showMessageAlert("请输入要查找的文本。");
      return;
    }
    updateHistory(findHistory, findRow.input.value);
    syncDatalist(findRow.list, findHistory);
    rebuildHits();
    if (hits.length === 0) {
      cursor = -1;
      clearFindResultTable();
      showMessageAlert("未找到所查询的数据。");
      return;
    }
    resultTbody.innerHTML = "";
    for (let i = 0; i < hits.length; i++) {
      const h = hits[i]!;
      const c = h.sheet.getCell(h.row, h.col);
      const vText = formatCellDisplayWithStyle(c.value, c.style);
      const fText = c.formula !== null && c.formula.length > 0 ? c.formula : "";
      const tr = document.createElement("tr");
      tr.addEventListener("click", (ev) => {
        ev.stopPropagation();
        navigateToHit(i);
      });
      const colTexts = [WORKBOOK_LIST_LABEL, h.sheet.name, "", hitToAbsDollarRef(h), vText, fText] as const;
      for (const t of colTexts) {
        const td = document.createElement("td");
        td.textContent = t;
        tr.appendChild(td);
      }
      resultTbody.appendChild(tr);
    }
    resultStatus.textContent = `${hits.length} 个单元格被找到`;
    resultBox.setAttribute("data-visible", "1");
    cursor = 0;
    syncFindResultSelection();
  };

  const stepFind = (delta: 1 | -1): void => {
    const opt = getScanOptions();
    if (opt.find.trim() === "") {
      showMessageAlert("请输入要查找的文本。");
      return;
    }
    updateHistory(findHistory, findRow.input.value);
    syncDatalist(findRow.list, findHistory);
    rebuildHits();
    if (hits.length === 0) {
      cursor = -1;
      clearFindResultTable();
      showMessageAlert("未找到所查询的数据。");
      return;
    }
    if (resultBox.getAttribute("data-visible") === "1" && resultTbody.querySelectorAll("tr").length !== hits.length) {
      clearFindResultTable();
    }
    if (cursor < 0 || cursor >= hits.length) {
      cursor = 0;
    } else {
      cursor = (cursor + (delta as number) + hits.length) % hits.length;
    }
    navigateToHit(cursor);
  };

  const doReplaceOne = (): void => {
    const opt = getScanOptions();
    if (opt.find.trim() === "") {
      showMessageAlert("请输入要查找的文本。");
      return;
    }
    if (!isHitReplaceable(opt.lookIn)) {
      showMessageAlert("当前「查找范围」下无法执行替换。请选择「公式」或「值」。");
      return;
    }
    updateHistory(findHistory, findRow.input.value);
    updateHistory(replaceHistory, replaceRow.input.value);
    syncDatalist(findRow.list, findHistory);
    syncDatalist(replaceRow.list, replaceHistory);

    rebuildHits();
    if (hits.length === 0) {
      cursor = -1;
      showMessageAlert("未找到所查询的数据。");
      return;
    }
    if (cursor < 0 || cursor >= hits.length) {
      cursor = 0;
    }
    const hit = hits[cursor]!;
    const hay = getHayStringForCell(hit.sheet, hit.row, hit.col, opt.lookIn);
    if (hay === null) {
      return;
    }
    const replaceWith = replaceRow.input.value;
    const replaced = buildReplacedString(hay, hit, replaceWith);
    setCellValueCommand(hit.sheet, hit.row, hit.col, replaced);
    onDataChanged();

    hits = collectFindHits(workbook, getScanOptions());
    const nextIdx = firstHitIndexAtOrAfter(hits, {
      si: hit.sheetIndex,
      row: hit.row,
      col: hit.col,
      at: hit.start + replaceWith.length,
    });
    if (nextIdx >= 0) {
      navigateToHit(nextIdx);
    } else {
      cursor = -1;
    }
  };

  const doReplaceAll = (): void => {
    const opt = getScanOptions();
    if (opt.find.trim() === "") {
      showMessageAlert("请输入要查找的文本。");
      return;
    }
    if (!isHitReplaceable(opt.lookIn)) {
      showMessageAlert("当前「查找范围」下无法执行全部替换。请选择「公式」或「值」。");
      return;
    }
    updateHistory(findHistory, findRow.input.value);
    updateHistory(replaceHistory, replaceRow.input.value);
    syncDatalist(findRow.list, findHistory);
    syncDatalist(replaceRow.list, replaceHistory);
    applyReplaceAllWithWriter(workbook, opt.find, replaceRow.input.value, opt, (sheet, row, col, value) => {
      setCellValueCommand(sheet, row, col, value);
    });
    onDataChanged();
    cursor = -1;
    rebuildHits();
    clearFindResultTable();
  };

  const setTab = (tab: "find" | "replace"): void => {
    currentTab = tab;
    const isFind = tab === "find";
    tabFind.setAttribute("aria-selected", isFind ? "true" : "false");
    tabReplace.setAttribute("aria-selected", isFind ? "false" : "true");
    replaceRow.row.style.display = isFind ? "none" : "flex";
    footerFind.style.display = isFind ? "flex" : "none";
    footerReplace.style.display = isFind ? "none" : "flex";
    if (isFind) {
      findRow.input.focus();
    } else {
      replaceRow.input.focus();
    }
  };

  optionToggle.addEventListener("click", () => {
    optionsWrap.hidden = !optionsWrap.hidden;
  });

  [findRow.input, replaceRow.input, withinSel, searchSel, lookInSel, checkCase.input, checkEntire.input, checkWidth.input].forEach(
    (el) => {
      const reset = (): void => {
        cursor = -1;
        clearFindResultTable();
      };
      el.addEventListener("input", reset);
      el.addEventListener("change", reset);
    },
  );

  tabFind.addEventListener("click", () => setTab("find"));
  tabReplace.addEventListener("click", () => setTab("replace"));

  findAllFindBtn.addEventListener("click", onFindAll);
  findAllReplaceBtn.addEventListener("click", onFindAll);
  nextFindBtn.addEventListener("click", () => stepFind(1));
  nextReplaceBtn.addEventListener("click", () => stepFind(1));
  prevFindBtn.addEventListener("click", () => stepFind(-1));
  prevReplaceBtn.addEventListener("click", () => stepFind(-1));
  replaceBtn.addEventListener("click", doReplaceOne);
  replaceAllBtn.addEventListener("click", doReplaceAll);
  closeFindBtn.addEventListener("click", dismiss);
  closeReplaceBtn.addEventListener("click", dismiss);

  overlay.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      dismiss();
    }
  });
  overlay.addEventListener("click", (ev) => {
    if (ev.target === overlay) {
      dismiss();
    }
  });

  setTab(initialTab);
  attachDraggableDialogPanel(panel, header);
  requestAnimationFrame(() => {
    if (currentTab === "find") {
      findRow.input.focus();
    } else {
      replaceRow.input.focus();
    }
  });
  return overlay;
}

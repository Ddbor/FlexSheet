/**
 * Ribbon「插入函数」右侧「公式生成器」面板：列表数据由后续接入；此处负责布局与交互骨架。
 */

export interface FormulaBuilderFunctionEntry {
  readonly id: string;
  /** 展示名，通常为大写函数名 */
  readonly name: string;
  readonly category: "common" | "all";
  /** 概要 */
  readonly description: string;
  /** 如 `XIRR(values, dates, guess)` */
  readonly syntax: string;
  /** 参数说明：名称 + 说明 */
  readonly parameters: readonly { readonly name: string; readonly text: string }[];
}

export interface FormulaBuilderPanelController {
  readonly root: HTMLElement;
  isOpen(): boolean;
  show(): void;
  hide(): void;
  /** 替换函数列表（如从引擎加载后）；会清空选中并重绘 */
  setEntries(entries: readonly FormulaBuilderFunctionEntry[]): void;
  destroy(): void;
}

export interface CreateFormulaBuilderPanelOptions {
  readonly parent: HTMLElement;
  readonly onClose: () => void;
  /** 当前有选中项且用户点击「插入函数」时触发 */
  readonly onInsert?: (entry: FormulaBuilderFunctionEntry) => void;
}

let stylesInjected = false;

function ensureFormulaBuilderPanelStyles(): void {
  if (stylesInjected) {
    return;
  }
  stylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-formula-builder", "1");
  style.textContent = `
.fs-formula-builder {
  box-sizing: border-box;
  display: none;
  flex-direction: column;
  width: 320px;
  flex-shrink: 0;
  min-width: 0;
  min-height: 0;
  height: 100%;
  background: #fff;
  border-left: 1px solid #c8c6c4;
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  font-size: 13px;
  color: #323130;
  z-index: 2;
}
.fs-formula-builder[data-open="1"] {
  display: flex;
}
.fs-formula-builder__header {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 12px 8px 14px;
  gap: 8px;
}
.fs-formula-builder__title {
  margin: 0;
  font-size: 15px;
  font-weight: 600;
  color: #201f1e;
}
.fs-formula-builder__close {
  flex-shrink: 0;
  width: 32px;
  height: 32px;
  margin: 0;
  padding: 0;
  border: none;
  border-radius: 4px;
  background: transparent;
  color: #605e5c;
  font-size: 18px;
  line-height: 1;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}
.fs-formula-builder__close:hover {
  background: #f3f2f1;
  color: #323130;
}
.fs-formula-builder__search-wrap {
  flex-shrink: 0;
  padding: 0 12px 10px 12px;
}
.fs-formula-builder__search-inner {
  position: relative;
  display: flex;
  align-items: center;
}
.fs-formula-builder__search-icon {
  position: absolute;
  left: 10px;
  width: 16px;
  height: 16px;
  color: #8a8886;
  pointer-events: none;
  display: flex;
  align-items: center;
  justify-content: center;
}
.fs-formula-builder__search-icon svg {
  width: 14px;
  height: 14px;
  display: block;
}
.fs-formula-builder__search {
  width: 100%;
  box-sizing: border-box;
  padding: 7px 10px 7px 34px;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  font: inherit;
  color: inherit;
  outline: none;
}
.fs-formula-builder__search:focus {
  border-color: #217346;
  box-shadow: 0 0 0 1px #217346 inset;
}
.fs-formula-builder__list-outer {
  flex: 1;
  min-height: 0;
  display: flex;
  flex-direction: column;
  margin: 0 12px;
  border: 1px solid #d2d0ce;
  border-radius: 2px;
  background: #fff;
  overflow: hidden;
}
.fs-formula-builder__list-scroll {
  flex: 1;
  min-height: 120px;
  overflow-y: auto;
  overflow-x: hidden;
}
.fs-formula-builder__cat {
  padding: 0 0 6px 0;
}
.fs-formula-builder__cat-title {
  padding: 8px 10px 4px 10px;
  font-size: 11px;
  font-weight: 600;
  color: #605e5c;
  text-transform: none;
  letter-spacing: 0;
}
.fs-formula-builder__fn {
  display: block;
  width: 100%;
  box-sizing: border-box;
  margin: 0;
  padding: 6px 10px;
  border: none;
  background: transparent;
  font: inherit;
  font-size: 13px;
  color: #201f1e;
  text-align: left;
  cursor: pointer;
}
.fs-formula-builder__fn:hover {
  background: #f3f2f1;
}
.fs-formula-builder__fn[data-selected="1"] {
  background: #edebe9;
}
.fs-formula-builder__fn:focus-visible {
  outline: 2px solid #217346;
  outline-offset: -2px;
}
.fs-formula-builder__muted {
  padding: 4px 10px 8px 10px;
  font-size: 12px;
  color: #8a8886;
}
.fs-formula-builder__empty-banner {
  padding: 16px 12px;
  font-size: 13px;
  color: #605e5c;
  text-align: center;
}
.fs-formula-builder__actions {
  flex-shrink: 0;
  padding: 8px 12px 10px 12px;
  text-align: right;
}
.fs-formula-builder__insert {
  padding: 6px 16px;
  font: inherit;
  font-size: 13px;
  border: 1px solid #8a8886;
  border-radius: 2px;
  background: #fff;
  color: #323130;
  cursor: pointer;
}
.fs-formula-builder__insert:hover:not(:disabled) {
  background: #f3f2f1;
}
.fs-formula-builder__insert:disabled {
  opacity: 0.45;
  cursor: default;
}
.fs-formula-builder__detail {
  flex-shrink: 0;
  max-height: 42%;
  min-height: 140px;
  overflow-y: auto;
  border-top: 1px solid #edebe9;
  padding: 12px 14px 16px 14px;
  background: #faf9f8;
}
.fs-formula-builder__detail-title {
  margin: 0 0 8px 0;
  font-size: 15px;
  font-weight: 600;
  color: #201f1e;
}
.fs-formula-builder__detail-title-fx {
  font-weight: 500;
  color: #217346;
  margin-right: 6px;
}
.fs-formula-builder__detail-desc {
  margin: 0 0 12px 0;
  font-size: 12px;
  line-height: 1.5;
  color: #323130;
}
.fs-formula-builder__detail-h {
  margin: 0 0 4px 0;
  font-size: 11px;
  font-weight: 600;
  color: #605e5c;
}
.fs-formula-builder__detail-syntax {
  margin: 0 0 12px 0;
  padding: 6px 8px;
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 12px;
  background: #fff;
  border: 1px solid #edebe9;
  border-radius: 2px;
  color: #201f1e;
  word-break: break-all;
}
.fs-formula-builder__detail-ul {
  margin: 0;
  padding-left: 18px;
  font-size: 12px;
  line-height: 1.55;
  color: #323130;
}
.fs-formula-builder__detail-li {
  margin-bottom: 6px;
}
.fs-formula-builder__detail-li strong {
  font-weight: 600;
}
.fs-formula-builder__placeholder {
  margin: 0;
  font-size: 12px;
  line-height: 1.5;
  color: #8a8886;
}
`;
  document.head.appendChild(style);
}

function iconSearchSvg(): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  svg.style.fill = "none";
  svg.style.stroke = "currentColor";
  svg.style.strokeWidth = "2";
  svg.style.strokeLinecap = "round";
  const c = document.createElementNS("http://www.w3.org/2000/svg", "circle");
  c.setAttribute("cx", "11");
  c.setAttribute("cy", "11");
  c.setAttribute("r", "7");
  const l = document.createElementNS("http://www.w3.org/2000/svg", "line");
  l.setAttribute("x1", "16.5");
  l.setAttribute("y1", "16.5");
  l.setAttribute("x2", "21");
  l.setAttribute("y2", "21");
  svg.appendChild(c);
  svg.appendChild(l);
  return svg;
}

const CATEGORY_ORDER: readonly { readonly id: FormulaBuilderFunctionEntry["category"]; readonly label: string }[] = [
  { id: "common", label: "常用函数" },
  { id: "all", label: "全部" },
];

export function createFormulaBuilderPanel(options: CreateFormulaBuilderPanelOptions): FormulaBuilderPanelController {
  ensureFormulaBuilderPanelStyles();

  let entries: readonly FormulaBuilderFunctionEntry[] = [];
  let searchQuery = "";
  let selectedId: string | null = null;
  let open = false;

  const root = document.createElement("aside");
  root.className = "fs-formula-builder";
  root.setAttribute("aria-label", "公式生成器");
  root.tabIndex = -1;

  const header = document.createElement("div");
  header.className = "fs-formula-builder__header";

  const title = document.createElement("h2");
  title.className = "fs-formula-builder__title";
  title.id = "fs-formula-builder-title";
  title.textContent = "公式生成器";
  root.setAttribute("aria-labelledby", "fs-formula-builder-title");

  const btnClose = document.createElement("button");
  btnClose.type = "button";
  btnClose.className = "fs-formula-builder__close";
  btnClose.setAttribute("aria-label", "关闭");
  btnClose.textContent = "\u00d7";

  header.appendChild(title);
  header.appendChild(btnClose);

  const searchWrap = document.createElement("div");
  searchWrap.className = "fs-formula-builder__search-wrap";
  const searchInner = document.createElement("div");
  searchInner.className = "fs-formula-builder__search-inner";
  const searchIcon = document.createElement("span");
  searchIcon.className = "fs-formula-builder__search-icon";
  searchIcon.appendChild(iconSearchSvg());
  const searchInput = document.createElement("input");
  searchInput.type = "search";
  searchInput.className = "fs-formula-builder__search";
  searchInput.placeholder = "搜索";
  searchInput.setAttribute("aria-label", "搜索函数");
  searchInner.appendChild(searchIcon);
  searchInner.appendChild(searchInput);
  searchWrap.appendChild(searchInner);

  const listOuter = document.createElement("div");
  listOuter.className = "fs-formula-builder__list-outer";
  const listScroll = document.createElement("div");
  listScroll.className = "fs-formula-builder__list-scroll";
  listScroll.setAttribute("role", "listbox");
  listScroll.setAttribute("aria-labelledby", "fs-formula-builder-title");
  listOuter.appendChild(listScroll);

  const actions = document.createElement("div");
  actions.className = "fs-formula-builder__actions";
  const btnInsert = document.createElement("button");
  btnInsert.type = "button";
  btnInsert.className = "fs-formula-builder__insert";
  btnInsert.textContent = "插入函数";
  actions.appendChild(btnInsert);

  const detail = document.createElement("div");
  detail.className = "fs-formula-builder__detail";

  root.appendChild(header);
  root.appendChild(searchWrap);
  root.appendChild(listOuter);
  root.appendChild(actions);
  root.appendChild(detail);

  options.parent.appendChild(root);

  const escapeOnDoc = (ev: KeyboardEvent): void => {
    if (!open) {
      return;
    }
    if (ev.key !== "Escape") {
      return;
    }
    if (ev.defaultPrevented) {
      return;
    }
    ev.preventDefault();
    ev.stopPropagation();
    hide();
    options.onClose();
  };

  function filteredEntries(): readonly FormulaBuilderFunctionEntry[] {
    const q = searchQuery.trim().toLowerCase();
    if (q.length === 0) {
      return entries;
    }
    return entries.filter((e) => e.name.toLowerCase().includes(q));
  }

  function entryById(id: string): FormulaBuilderFunctionEntry | undefined {
    return entries.find((e) => e.id === id);
  }

  function syncInsertEnabled(): void {
    const sel = selectedId !== null ? entryById(selectedId) : undefined;
    const filtered = filteredEntries();
    const ok =
      sel !== undefined && filtered.some((e) => e.id === sel.id);
    btnInsert.disabled = !ok;
  }

  function renderDetail(): void {
    detail.innerHTML = "";
    const sel = selectedId !== null ? entryById(selectedId) : undefined;
    const filtered = filteredEntries();
    const active =
      sel !== undefined && filtered.some((e) => e.id === sel.id) ? sel : undefined;

    if (active === undefined) {
      const p = document.createElement("p");
      p.className = "fs-formula-builder__placeholder";
      p.textContent = "请从上方列表中选择函数";
      detail.appendChild(p);
      syncInsertEnabled();
      return;
    }

    const h = document.createElement("h3");
    h.className = "fs-formula-builder__detail-title";
    const fx = document.createElement("span");
    fx.className = "fs-formula-builder__detail-title-fx";
    fx.textContent = "fx";
    h.appendChild(fx);
    h.appendChild(document.createTextNode(` ${active.name}`));
    detail.appendChild(h);

    const desc = document.createElement("p");
    desc.className = "fs-formula-builder__detail-desc";
    desc.textContent = active.description;
    detail.appendChild(desc);

    const hSyn = document.createElement("p");
    hSyn.className = "fs-formula-builder__detail-h";
    hSyn.textContent = "语法";
    detail.appendChild(hSyn);

    const syn = document.createElement("pre");
    syn.className = "fs-formula-builder__detail-syntax";
    syn.textContent = active.syntax;
    detail.appendChild(syn);

    if (active.parameters.length > 0) {
      const ul = document.createElement("ul");
      ul.className = "fs-formula-builder__detail-ul";
      for (const pr of active.parameters) {
        const li = document.createElement("li");
        li.className = "fs-formula-builder__detail-li";
        const strong = document.createElement("strong");
        strong.textContent = pr.name;
        li.appendChild(strong);
        li.appendChild(document.createTextNode(`：${pr.text}`));
        ul.appendChild(li);
      }
      detail.appendChild(ul);
    }
    syncInsertEnabled();
  }

  function renderList(): void {
    listScroll.innerHTML = "";
    const filtered = filteredEntries();
    const q = searchQuery.trim();

    if (entries.length > 0 && filtered.length === 0 && q.length > 0) {
      const banner = document.createElement("div");
      banner.className = "fs-formula-builder__empty-banner";
      banner.textContent = "无匹配函数";
      listScroll.appendChild(banner);
      if (selectedId !== null && !filtered.some((e) => e.id === selectedId)) {
        selectedId = null;
      }
      renderDetail();
      return;
    }

    if (entries.length === 0) {
      for (const cat of CATEGORY_ORDER) {
        const block = document.createElement("div");
        block.className = "fs-formula-builder__cat";
        const ct = document.createElement("div");
        ct.className = "fs-formula-builder__cat-title";
        ct.textContent = cat.label;
        block.appendChild(ct);
        const muted = document.createElement("div");
        muted.className = "fs-formula-builder__muted";
        muted.textContent = "暂无函数";
        block.appendChild(muted);
        listScroll.appendChild(block);
      }
      selectedId = null;
      renderDetail();
      return;
    }

    for (const cat of CATEGORY_ORDER) {
      const inCat = filtered.filter((e) => e.category === cat.id);
      const block = document.createElement("div");
      block.className = "fs-formula-builder__cat";
      const ct = document.createElement("div");
      ct.className = "fs-formula-builder__cat-title";
      ct.textContent = cat.label;
      block.appendChild(ct);
      if (inCat.length === 0) {
        const muted = document.createElement("div");
        muted.className = "fs-formula-builder__muted";
        muted.textContent = q.length > 0 ? "无匹配" : "（空）";
        block.appendChild(muted);
      } else {
        for (const e of inCat) {
          const row = document.createElement("button");
          row.type = "button";
          row.className = "fs-formula-builder__fn";
          row.setAttribute("role", "option");
          row.dataset.id = e.id;
          row.textContent = e.name;
          if (selectedId === e.id) {
            row.dataset.selected = "1";
            row.setAttribute("aria-selected", "true");
          } else {
            row.setAttribute("aria-selected", "false");
          }
          row.addEventListener("click", () => {
            selectedId = e.id;
            renderList();
            renderDetail();
          });
          block.appendChild(row);
        }
      }
      listScroll.appendChild(block);
    }

    if (selectedId !== null && !filtered.some((e) => e.id === selectedId)) {
      selectedId = null;
    }
    renderDetail();
  }

  function attachEscape(): void {
    document.addEventListener("keydown", escapeOnDoc, true);
  }

  function detachEscape(): void {
    document.removeEventListener("keydown", escapeOnDoc, true);
  }

  function show(): void {
    open = true;
    root.dataset.open = "1";
    attachEscape();
    renderList();
    queueMicrotask(() => {
      searchInput.focus();
    });
  }

  function hide(): void {
    open = false;
    root.removeAttribute("data-open");
    detachEscape();
    searchInput.blur();
  }

  btnClose.addEventListener("click", () => {
    hide();
    options.onClose();
  });

  searchInput.addEventListener("input", () => {
    searchQuery = searchInput.value;
    renderList();
  });

  btnInsert.addEventListener("click", () => {
    if (btnInsert.disabled) {
      return;
    }
    const sel = selectedId !== null ? entryById(selectedId) : undefined;
    if (sel === undefined) {
      return;
    }
    options.onInsert?.(sel);
  });

  return {
    root,
    isOpen(): boolean {
      return open;
    },
    show(): void {
      show();
    },
    hide(): void {
      hide();
    },
    setEntries(next: readonly FormulaBuilderFunctionEntry[]): void {
      entries = next;
      selectedId = null;
      searchQuery = "";
      searchInput.value = "";
      if (open) {
        renderList();
      }
    },
    destroy(): void {
      detachEscape();
      open = false;
      root.removeAttribute("data-open");
      searchInput.blur();
      root.remove();
    },
  };
}

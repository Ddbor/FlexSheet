/**
 * Ribbon「插入函数」右侧「公式生成器」：函数列表、详情、参数编辑（插入函数后）、选区引用与完成插入。
 */

import { FORMULA_BUILDER_COMMON_FUNCTION_ORDER } from "./formula-builder-common-order.js";

export interface FormulaBuilderFunctionEntry {
  readonly id: string;
  /** 展示名，通常为大写函数名 */
  readonly name: string;
  /** `"common"` 表示常用函数；否则为 Microsoft 支持索引中的类型标题（如「财务」「统计」）。 */
  readonly category: string;
  /** 概要 */
  readonly description: string;
  /** 如 `XIRR(values, dates, guess)` */
  readonly syntax: string;
  /** 参数说明：名称 + 说明 */
  readonly parameters: readonly { readonly name: string; readonly text: string }[];
  /** 与 Excel 类似可追加同类参数（如 number1、number2…） */
  readonly variadic?: boolean;
  /** Microsoft 支持文章路径（`https://support.microsoft.com` + 此路径） */
  readonly supportPath?: string;
}

export interface FormulaBuilderPanelController {
  readonly root: HTMLElement;
  isOpen(): boolean;
  /** 是否处于「图 2」参数编辑视图 */
  isParamMode(): boolean;
  /** 参数视图中是否应将当前选区写入活动参数（表体框选/点选结束后由 FlexSheet 调用） */
  isPickingArgRefFromSheet(): boolean;
  /** 将当前表选区对应的绝对引用写入活动参数框（不含前导 `=`） */
  applyArgRefFromSheet(absoluteRefWithoutLeadingEquals: string): void;
  show(): void;
  hide(): void;
  /** 替换函数列表；会清空选中并重绘 */
  setEntries(entries: readonly FormulaBuilderFunctionEntry[]): void;
  destroy(): void;
}

export interface CreateFormulaBuilderPanelOptions {
  readonly parent: HTMLElement;
  readonly onClose: () => void;
  /** 用户点击「插入函数」进入参数编辑时调用（用于锁定要写入的目标单元格）。 */
  readonly onBeginParamEdit?: () => void;
  /** 实时预览 `=…` 公式的标量结果展示文案 */
  readonly onPreviewFormula?: (formulaWithLeadingEquals: string) => string | null;
  /** 用户点击「完成」：写入 `=函数(…)` */
  readonly onApplyFormula?: (formulaWithLeadingEquals: string) => void;
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
.fs-formula-builder__browse {
  flex: 1;
  min-height: 0;
  display: flex;
  flex-direction: column;
}
.fs-formula-builder__browse[data-hidden="1"] {
  display: none;
}
.fs-formula-builder__param-stage {
  flex: 1;
  min-height: 0;
  display: none;
  flex-direction: column;
  overflow: hidden;
}
.fs-formula-builder__param-stage[data-visible="1"] {
  display: flex;
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
.fs-formula-builder__back-row {
  flex-shrink: 0;
  padding: 4px 12px 8px 12px;
}
.fs-formula-builder__back {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 4px 8px;
  margin: 0;
  border: none;
  border-radius: 4px;
  background: transparent;
  font: inherit;
  font-size: 13px;
  color: #217346;
  cursor: pointer;
}
.fs-formula-builder__back:hover {
  background: #f3f2f1;
}
.fs-formula-builder__param-fn {
  flex-shrink: 0;
  margin: 0 14px 8px 14px;
  font-size: 18px;
  font-weight: 600;
  color: #201f1e;
}
.fs-formula-builder__param-scroll {
  flex: 1;
  min-height: 80px;
  overflow-y: auto;
  padding: 0 12px 8px 12px;
}
.fs-formula-builder__param-row {
  display: flex;
  align-items: center;
  gap: 6px;
  margin-bottom: 8px;
}
.fs-formula-builder__param-row-remove {
  flex-shrink: 0;
  width: 28px;
  height: 28px;
  padding: 0;
  border: none;
  border-radius: 4px;
  background: transparent;
  color: #605e5c;
  font-size: 18px;
  line-height: 1;
  cursor: pointer;
  visibility: hidden;
}
.fs-formula-builder__param-row[data-show-remove="1"] .fs-formula-builder__param-row-remove {
  visibility: visible;
}
.fs-formula-builder__param-row-remove:hover {
  background: #f3f2f1;
}
.fs-formula-builder__param-field {
  flex: 1;
  min-width: 0;
  display: flex;
  align-items: stretch;
}
.fs-formula-builder__param-input {
  flex: 1;
  min-width: 0;
  box-sizing: border-box;
  padding: 7px 8px;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  font: inherit;
  color: inherit;
  outline: none;
}
.fs-formula-builder__param-input:focus {
  border-color: #217346;
  box-shadow: 0 0 0 1px #217346 inset;
}
.fs-formula-builder__param-plus {
  margin: 0 12px 8px 40px;
  width: 28px;
  height: 28px;
  padding: 0;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  background: #fff;
  font-size: 18px;
  line-height: 1;
  cursor: pointer;
  color: #323130;
}
.fs-formula-builder__param-plus:hover {
  background: #f3f2f1;
}
.fs-formula-builder__param-plus[data-hidden="1"] {
  display: none;
}
.fs-formula-builder__result-row {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  padding: 8px 12px 10px 12px;
  border-top: 1px solid #edebe9;
  background: #faf9f8;
}
.fs-formula-builder__result-text {
  flex: 1;
  min-width: 0;
  font-size: 12px;
  color: #323130;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.fs-formula-builder__done {
  flex-shrink: 0;
  padding: 6px 14px;
  font: inherit;
  font-size: 13px;
  border: 1px solid #217346;
  border-radius: 2px;
  background: #217346;
  color: #fff;
  cursor: pointer;
}
.fs-formula-builder__done:hover {
  filter: brightness(1.05);
}
.fs-formula-builder__detail {
  flex-shrink: 0;
  max-height: 38%;
  min-height: 120px;
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

/** 文档分组标题的大致展示顺序（未列入的类别按中文排序排在后面）。 */
const MS_CATEGORY_SORT_ORDER: readonly string[] = [
  "财务",
  "日期与时间",
  "数学与三角函数",
  "数学和三角",
  "统计",
  "查找与引用",
  "Database",
  "文本",
  "逻辑",
  "信息",
  "工程",
  "多维数据集",
  "兼容性",
  "Web",
  "加载项和自动化",
  "其他",
];

function sortCommonEntries(list: readonly FormulaBuilderFunctionEntry[]): FormulaBuilderFunctionEntry[] {
  const rank = new Map(FORMULA_BUILDER_COMMON_FUNCTION_ORDER.map((n, i) => [n, i]));
  return [...list].sort((a, b) => (rank.get(a.name) ?? 999) - (rank.get(b.name) ?? 999));
}

function orderedDocCategoriesForFilter(filtered: readonly FormulaBuilderFunctionEntry[]): string[] {
  const set = new Set<string>();
  for (const e of filtered) {
    if (e.category !== "common") {
      set.add(e.category);
    }
  }
  return [...set].sort((a, b) => {
    const ia = MS_CATEGORY_SORT_ORDER.indexOf(a);
    const ib = MS_CATEGORY_SORT_ORDER.indexOf(b);
    const sa = ia === -1 ? 1000 : ia;
    const sb = ib === -1 ? 1000 : ib;
    if (sa !== sb) {
      return sa - sb;
    }
    return a.localeCompare(b, "zh-CN");
  });
}

interface ParamSlot {
  readonly label: string;
  readonly hint: string;
}

function buildParamSlots(entry: FormulaBuilderFunctionEntry): { slots: ParamSlot[]; variadic: boolean; minRows: number } {
  const params = entry.parameters;
  const variadicFlag = entry.variadic === true || /…|\.\.\./.test(entry.syntax);
  if (params.length === 0) {
    return {
      slots: [{ label: "参数1", hint: entry.syntax }],
      variadic: variadicFlag || true,
      minRows: 1,
    };
  }
  if (variadicFlag && params.length >= 2) {
    const a = params[0]!;
    const b = params[1]!;
    return {
      slots: [
        { label: a.name, hint: a.text },
        { label: b.name.includes("…") || b.name.includes("...") ? "Number2" : b.name, hint: b.text },
      ],
      variadic: true,
      minRows: 1,
    };
  }
  if (variadicFlag && params.length === 1) {
    const a = params[0]!;
    return {
      slots: [
        { label: a.name, hint: a.text },
        { label: `${a.name}（2）`, hint: a.text },
      ],
      variadic: true,
      minRows: 1,
    };
  }
  return {
    slots: params.map((p) => ({ label: p.name, hint: p.text })),
    variadic: false,
    minRows: params.length,
  };
}

export function createFormulaBuilderPanel(options: CreateFormulaBuilderPanelOptions): FormulaBuilderPanelController {
  ensureFormulaBuilderPanelStyles();

  let entries: readonly FormulaBuilderFunctionEntry[] = [];
  let searchQuery = "";
  let selectedId: string | null = null;
  let open = false;

  let paramMode = false;
  let paramEntry: FormulaBuilderFunctionEntry | null = null;
  let paramValues: string[] = [];
  let paramSlotTemplates: ParamSlot[] = [];
  let paramVariadic = false;
  let paramMinRows = 1;
  let activeArgIndex = 0;
  let previewRaf = 0;

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

  const browse = document.createElement("div");
  browse.className = "fs-formula-builder__browse";

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

  browse.appendChild(searchWrap);
  browse.appendChild(listOuter);
  browse.appendChild(actions);

  const paramStage = document.createElement("div");
  paramStage.className = "fs-formula-builder__param-stage";

  const backRow = document.createElement("div");
  backRow.className = "fs-formula-builder__back-row";
  const btnShowAll = document.createElement("button");
  btnShowAll.type = "button";
  btnShowAll.className = "fs-formula-builder__back";
  btnShowAll.innerHTML = "<span aria-hidden=\"true\">\u2190</span> 显示所有函数";
  backRow.appendChild(btnShowAll);

  const paramFnTitle = document.createElement("div");
  paramFnTitle.className = "fs-formula-builder__param-fn";

  const paramScroll = document.createElement("div");
  paramScroll.className = "fs-formula-builder__param-scroll";
  const paramRowsHost = document.createElement("div");
  paramScroll.appendChild(paramRowsHost);

  const btnPlus = document.createElement("button");
  btnPlus.type = "button";
  btnPlus.className = "fs-formula-builder__param-plus";
  btnPlus.textContent = "+";
  btnPlus.title = "添加参数";

  const resultRow = document.createElement("div");
  resultRow.className = "fs-formula-builder__result-row";
  const resultText = document.createElement("div");
  resultText.className = "fs-formula-builder__result-text";
  resultText.textContent = "结果: …";
  const btnDone = document.createElement("button");
  btnDone.type = "button";
  btnDone.className = "fs-formula-builder__done";
  btnDone.textContent = "完成";
  resultRow.appendChild(resultText);
  resultRow.appendChild(btnDone);

  paramStage.appendChild(backRow);
  paramStage.appendChild(paramFnTitle);
  paramStage.appendChild(paramScroll);
  paramStage.appendChild(btnPlus);
  paramStage.appendChild(resultRow);

  const detail = document.createElement("div");
  detail.className = "fs-formula-builder__detail";

  root.appendChild(header);
  root.appendChild(browse);
  root.appendChild(paramStage);
  root.appendChild(detail);

  options.parent.appendChild(root);

  function exitParamMode(): void {
    paramMode = false;
    paramEntry = null;
    paramValues = [];
    paramSlotTemplates = [];
    paramVariadic = false;
    paramMinRows = 1;
    activeArgIndex = 0;
    browse.removeAttribute("data-hidden");
    paramStage.removeAttribute("data-visible");
    cancelPreviewRaf();
    resultText.textContent = "结果: …";
    paramRowsHost.innerHTML = "";
  }

  function cancelPreviewRaf(): void {
    if (previewRaf !== 0) {
      cancelAnimationFrame(previewRaf);
      previewRaf = 0;
    }
  }

  function schedulePreview(): void {
    cancelPreviewRaf();
    previewRaf = requestAnimationFrame(() => {
      previewRaf = 0;
      const built = buildFormulaString();
      if (built === null) {
        resultText.textContent = "结果: …";
        return;
      }
      const s = options.onPreviewFormula?.(built);
      resultText.textContent = s !== null && s !== undefined && s.length > 0 ? `结果: ${s}` : `结果: ${built}`;
    });
  }

  function buildFormulaString(): string | null {
    if (paramEntry === null) {
      return null;
    }
    const parts: string[] = [];
    for (const v of paramValues) {
      const t = v.trim();
      if (t.length > 0) {
        parts.push(t);
      }
    }
    return `=${paramEntry.name}(${parts.join(",")})`;
  }

  function slotLabelForIndex(i: number): string {
    if (paramSlotTemplates.length === 0) {
      return `参数${i + 1}`;
    }
    if (i < paramSlotTemplates.length) {
      return paramSlotTemplates[i]!.label;
    }
    if (!paramVariadic) {
      return `参数${i + 1}`;
    }
    const first = paramSlotTemplates[0]!.label;
    if (/^number\s*1$/i.test(first.trim())) {
      return `Number${i + 1}`;
    }
    const last = paramSlotTemplates[paramSlotTemplates.length - 1]!.label;
    const m = /^(.+?)(\d+)$/.exec(last.trim());
    if (m !== null) {
      const prefix = m[1];
      const baseNum = Number(m[2]);
      const delta = i - paramSlotTemplates.length + 1;
      return `${prefix}${baseNum + delta}`;
    }
    return `${last} (${i + 1})`;
  }

  function slotHintForIndex(i: number): string {
    if (paramSlotTemplates.length === 0) {
      return paramEntry?.syntax ?? "";
    }
    if (i < paramSlotTemplates.length) {
      return paramSlotTemplates[i]!.hint;
    }
    return paramSlotTemplates[paramSlotTemplates.length - 1]!.hint;
  }

  function enterParamMode(entry: FormulaBuilderFunctionEntry): void {
    options.onBeginParamEdit?.();
    paramMode = true;
    paramEntry = entry;
    const { slots, variadic, minRows } = buildParamSlots(entry);
    paramSlotTemplates = slots;
    paramVariadic = variadic;
    paramMinRows = minRows;
    paramValues = slots.map(() => "");
    activeArgIndex = 0;
    browse.setAttribute("data-hidden", "1");
    paramStage.setAttribute("data-visible", "1");
    paramFnTitle.textContent = entry.name;
    btnPlus.dataset.hidden = variadic ? "0" : "1";
    renderParamRows();
    schedulePreview();
    renderDetailFor(entry);
  }

  function renderParamRows(): void {
    if (paramEntry === null) {
      return;
    }
    paramRowsHost.innerHTML = "";
    for (let i = 0; i < paramValues.length; i++) {
      const row = document.createElement("div");
      row.className = "fs-formula-builder__param-row";
      const showRemove = paramVariadic && paramValues.length > paramMinRows;
      if (showRemove) {
        row.dataset.showRemove = "1";
      }

      const btnRm = document.createElement("button");
      btnRm.type = "button";
      btnRm.className = "fs-formula-builder__param-row-remove";
      btnRm.setAttribute("aria-label", "移除此参数");
      btnRm.textContent = "\u2212";
      btnRm.addEventListener("click", () => {
        if (!paramVariadic || paramValues.length <= paramMinRows) {
          return;
        }
        paramValues.splice(i, 1);
        activeArgIndex = Math.min(activeArgIndex, paramValues.length - 1);
        renderParamRows();
        schedulePreview();
      });

      const labWrap = document.createElement("label");
      labWrap.style.flex = "1";
      labWrap.style.minWidth = "0";
      labWrap.style.display = "flex";
      labWrap.style.flexDirection = "column";
      labWrap.style.gap = "4px";
      const lab = document.createElement("span");
      lab.style.fontSize = "11px";
      lab.style.color = "#605e5c";
      const hint = slotHintForIndex(i);
      const labelText = slotLabelForIndex(i);
      lab.textContent = hint.length > 0 ? `${labelText} = ${hint.slice(0, 40)}${hint.length > 40 ? "…" : ""}` : labelText;

      const field = document.createElement("div");
      field.className = "fs-formula-builder__param-field";
      const input = document.createElement("input");
      input.type = "text";
      input.className = "fs-formula-builder__param-input";
      input.autocomplete = "off";
      input.value = paramValues[i] ?? "";
      input.addEventListener("focus", () => {
        activeArgIndex = i;
      });
      input.addEventListener("input", () => {
        paramValues[i] = input.value;
        schedulePreview();
      });

      field.appendChild(input);
      labWrap.appendChild(lab);
      labWrap.appendChild(field);
      row.appendChild(btnRm);
      row.appendChild(labWrap);
      paramRowsHost.appendChild(row);
    }
    const inputs = paramRowsHost.querySelectorAll<HTMLInputElement>(".fs-formula-builder__param-input");
    const target = inputs[Math.min(activeArgIndex, inputs.length - 1)];
    if (target !== undefined) {
      queueMicrotask(() => target.focus());
    }
  }

  btnPlus.addEventListener("click", () => {
    if (!paramVariadic) {
      return;
    }
    paramValues.push("");
    activeArgIndex = paramValues.length - 1;
    renderParamRows();
    schedulePreview();
  });

  btnShowAll.addEventListener("click", () => {
    exitParamMode();
    renderList();
    renderDetail();
  });

  btnDone.addEventListener("click", () => {
    const built = buildFormulaString();
    if (built === null || paramEntry === null) {
      return;
    }
    options.onApplyFormula?.(built);
    exitParamMode();
    renderList();
    renderDetail();
  });

  const escapeOnDoc = (ev: KeyboardEvent): void => {
    if (!open) {
      return;
    }
    if (ev.key !== "Escape" || ev.isComposing) {
      return;
    }
    if (ev.defaultPrevented) {
      return;
    }
    if (paramMode) {
      ev.preventDefault();
      ev.stopPropagation();
      exitParamMode();
      renderList();
      renderDetail();
      return;
    }
    ev.preventDefault();
    ev.stopPropagation();
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

  function renderDetailFor(active: FormulaBuilderFunctionEntry | undefined): void {
    detail.innerHTML = "";
    if (active === undefined) {
      if (!paramMode) {
        const p = document.createElement("p");
        p.className = "fs-formula-builder__placeholder";
        p.textContent = "请从上方列表中选择函数";
        detail.appendChild(p);
      }
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

  function renderDetail(): void {
    if (paramMode && paramEntry !== null) {
      renderDetailFor(paramEntry);
      return;
    }
    const sel = selectedId !== null ? entryById(selectedId) : undefined;
    const filtered = filteredEntries();
    const active =
      sel !== undefined && filtered.some((e) => e.id === sel.id) ? sel : undefined;
    renderDetailFor(active);
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
      const banner = document.createElement("div");
      banner.className = "fs-formula-builder__empty-banner";
      banner.textContent = "暂无函数";
      listScroll.appendChild(banner);
      selectedId = null;
      renderDetail();
      return;
    }

    const docCategories = orderedDocCategoriesForFilter(filtered);
    const sectionMetas: { readonly id: string; readonly label: string }[] = [
      { id: "common", label: "常用函数" },
      ...docCategories.map((c) => ({ id: c, label: c })),
    ];

    for (const cat of sectionMetas) {
      const inCat =
        cat.id === "common"
          ? sortCommonEntries(filtered.filter((e) => e.category === "common"))
          : filtered.filter((e) => e.category === cat.id);
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
      if (!paramMode) {
        searchInput.focus();
      }
    });
  }

  function hide(): void {
    open = false;
    root.removeAttribute("data-open");
    detachEscape();
    searchInput.blur();
    exitParamMode();
  }

  btnClose.addEventListener("click", () => {
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
    enterParamMode(sel);
  });

  return {
    root,
    isOpen(): boolean {
      return open;
    },
    isParamMode(): boolean {
      return paramMode;
    },
    isPickingArgRefFromSheet(): boolean {
      return open && paramMode && paramEntry !== null;
    },
    applyArgRefFromSheet(absoluteRefWithoutLeadingEquals: string): void {
      if (!paramMode || paramValues.length === 0) {
        return;
      }
      const i = Math.min(Math.max(0, activeArgIndex), paramValues.length - 1);
      paramValues[i] = absoluteRefWithoutLeadingEquals.trim();
      renderParamRows();
      schedulePreview();
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
      exitParamMode();
      if (open) {
        renderList();
      }
    },
    destroy(): void {
      cancelPreviewRaf();
      detachEscape();
      open = false;
      root.removeAttribute("data-open");
      searchInput.blur();
      exitParamMode();
      root.remove();
    },
  };
}

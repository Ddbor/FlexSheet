/**
 * Office 风格「文件」Backstage：左侧导航 + 右侧内容；保存/导入 FlexSheet JSON。
 */

import {
  DEFAULT_FLEXSHEET_JSON_EXPORT_OPTIONS,
  decodeTextFileBytes,
  downloadJsonText,
  downloadXlsxBlob,
  exportWorkbookToXlsxBlob,
  importXlsxToWorkbook,
  parseFlexSheetJson,
  serializeWorkbookToJsonDocument,
  workbookFromFlexSheetJsonDocument,
  type FlexSheetJsonExportOptions,
  type FlexSheetJsonImportOptions,
  type FlexSheetJsonViewState,
  type XlsxExportOptions,
} from "@flexsheet/import-export";

import type { FlexSheetLike } from "./ribbon-types.js";

import "./ribbon-backstage.css";

export type RibbonBackstageSectionId = "open" | "save" | "import" | "export";

export interface RibbonBackstageHandles {
  readonly root: HTMLElement;
  showSection(id: RibbonBackstageSectionId): void;
  applyThemeMode(mode: "light" | "dark"): void;
}

export interface RibbonBackstageOptions {
  readonly flexSheet?: FlexSheetLike;
}

const SECTION_ORDER: RibbonBackstageSectionId[] = ["open", "save", "import", "export"];

const SECTION_LABEL: Record<RibbonBackstageSectionId, string> = {
  open: "打开",
  save: "保存",
  import: "导入",
  export: "导出",
};

/** 侧栏在哪些项之后插入分隔线（与参考图一致） */
const DIVIDER_AFTER: Partial<Record<RibbonBackstageSectionId, boolean>> = {
  save: true,
  export: true,
};

const JSON_GENERATOR_VERSION = "0.0.1";

function showImportError(msg: HTMLElement, text: string): void {
  msg.textContent = text;
  msg.classList.add("fs-backstage__msg--error");
  msg.hidden = false;
}

function iconBack(): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "20");
  svg.setAttribute("height", "20");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("aria-hidden", "true");
  const p = document.createElementNS("http://www.w3.org/2000/svg", "path");
  p.setAttribute(
    "d",
    "M10.5 3.5L6 8l4.5 4.5-1 1L4 8 9.5 2.5l1 1z",
  );
  p.setAttribute("fill", "currentColor");
  svg.appendChild(p);
  return svg;
}

function iconDocument(): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "48");
  svg.setAttribute("height", "48");
  svg.setAttribute("viewBox", "0 0 48 48");
  svg.setAttribute("aria-hidden", "true");
  const r = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  r.setAttribute("x", "10");
  r.setAttribute("y", "6");
  r.setAttribute("width", "28");
  r.setAttribute("height", "36");
  r.setAttribute("rx", "2");
  r.setAttribute("fill", "none");
  r.setAttribute("stroke", "currentColor");
  r.setAttribute("stroke-width", "2");
  svg.appendChild(r);
  const line = document.createElementNS("http://www.w3.org/2000/svg", "path");
  line.setAttribute("d", "M14 16h20M14 22h16M14 28h20");
  line.setAttribute("stroke", "currentColor");
  line.setAttribute("stroke-width", "1.5");
  svg.appendChild(line);
  return svg;
}

function createCheckboxRow(label: string, checked: boolean): { row: HTMLElement; input: HTMLInputElement } {
  const row = document.createElement("label");
  row.className = "fs-backstage__check-row";
  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = checked;
  const span = document.createElement("span");
  span.textContent = label;
  row.appendChild(input);
  row.appendChild(span);
  return { row, input };
}

function captureViewState(fs: FlexSheetLike): FlexSheetJsonViewState {
  const r = fs.getRenderer();
  const s = r.getScroll();
  return {
    frozenRows: r.frozenRows,
    frozenCols: r.frozenCols,
    scrollX: s.scrollX,
    scrollY: s.scrollY,
  };
}

function applyViewState(fs: FlexSheetLike, v: FlexSheetJsonViewState): void {
  fs.setFrozenPanes(v.frozenRows, v.frozenCols);
  fs.getRenderer().setScroll(v.scrollX, v.scrollY);
  fs.getRenderer().ensureScrollClamped();
  fs.refresh();
}

function readTextFile(file: File, msg: HTMLElement, onOk: (text: string) => void): void {
  const reader = new FileReader();
  reader.onload = (): void => {
    const r = reader.result;
    if (r instanceof ArrayBuffer) {
      onOk(decodeTextFileBytes(new Uint8Array(r)));
      return;
    }
    onOk("");
  };
  reader.onerror = (): void => {
    msg.textContent = "读取文件失败。";
    msg.classList.add("fs-backstage__msg--error");
    msg.hidden = false;
  };
  reader.readAsArrayBuffer(file);
}

function loadFlexSheetJsonFromText(
  text: string,
  fs: FlexSheetLike,
  importOpts: FlexSheetJsonImportOptions,
  msg: HTMLElement,
  successText: string,
  onSuccessClose?: () => void,
): void {
  msg.hidden = true;
  const parsed = parseFlexSheetJson(text);
  if (!parsed.ok) {
    msg.textContent = parsed.error;
    msg.classList.add("fs-backstage__msg--error");
    msg.hidden = false;
    return;
  }
  const lw = fs.loadWorkbook;
  if (lw === undefined) {
    msg.textContent = "当前环境不支持加载工作簿。";
    msg.classList.add("fs-backstage__msg--error");
    msg.hidden = false;
    return;
  }
  const wb = workbookFromFlexSheetJsonDocument(parsed.doc, importOpts);
  lw.call(fs, wb);
  const v = parsed.doc.workbook.view;
  if (v !== undefined) {
    applyViewState(fs, v);
  }
  if (onSuccessClose !== undefined) {
    onSuccessClose();
  } else {
    msg.textContent = successText;
    msg.classList.remove("fs-backstage__msg--error");
    msg.hidden = false;
  }
}

function mountOpenPanel(
  container: HTMLElement,
  flexSheet: FlexSheetLike | undefined,
  onSuccessClose: () => void,
): void {
  container.innerHTML = "";
  const title = document.createElement("h1");
  title.className = "fs-backstage__page-title";
  title.textContent = "打开";

  const sub = document.createElement("p");
  sub.className = "fs-backstage__page-sub";
  sub.textContent = "FlexSheet 工作簿 (.json)";

  const hint = document.createElement("p");
  hint.className = "fs-backstage__import-hint";
  hint.textContent =
    "请选择 .json 文件或拖拽到下方区域。仅加载由 FlexSheet 导出的工作簿 JSON，校验失败将不会替换当前表格。";

  const msg = document.createElement("p");
  msg.className = "fs-backstage__msg";
  msg.setAttribute("role", "status");
  msg.hidden = true;

  if (flexSheet === undefined) {
    const ph = document.createElement("p");
    ph.className = "fs-backstage__placeholder";
    ph.textContent = "当前未连接工作簿，无法打开文件。";
    container.appendChild(title);
    container.appendChild(sub);
    container.appendChild(ph);
    return;
  }
  if (flexSheet.workbook === undefined || flexSheet.loadWorkbook === undefined) {
    const ph = document.createElement("p");
    ph.className = "fs-backstage__placeholder";
    ph.textContent = "当前未连接工作簿，无法打开文件。";
    container.appendChild(title);
    container.appendChild(sub);
    container.appendChild(ph);
    return;
  }

  const fs = flexSheet;

  const groups = document.createElement("div");
  groups.className = "fs-backstage__open-groups";

  const g1 = document.createElement("div");
  g1.className = "fs-backstage__open-group";
  g1.appendChild(createCheckboxRow("增量加载", false).row);
  g1.appendChild(createCheckboxRow("懒加载", false).row);

  const g2 = document.createElement("div");
  g2.className = "fs-backstage__open-group";
  const openStyles = createCheckboxRow("包含样式", true);
  const openFormulas = createCheckboxRow("包含公式", true);
  const openRecalc = createCheckboxRow("导入后自动计算", false);
  g2.appendChild(openStyles.row);
  g2.appendChild(openFormulas.row);
  g2.appendChild(openRecalc.row);

  const g3 = document.createElement("div");
  g3.className = "fs-backstage__open-group";
  g3.appendChild(createCheckboxRow("导入未使用的命名样式", true).row);
  g3.appendChild(createCheckboxRow("增量计算", true).row);
  g3.appendChild(createCheckboxRow("按需计算", true).row);
  g3.appendChild(createCheckboxRow("动态引用", false).row);

  groups.appendChild(g1);
  groups.appendChild(g2);
  groups.appendChild(g3);

  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = "application/json,.json";
  fileInput.className = "fs-backstage__file-input";
  fileInput.hidden = true;

  const card = document.createElement("button");
  card.type = "button";
  card.className = "fs-backstage__open-card fs-backstage__open-card--drop";
  const iconWrap = document.createElement("div");
  iconWrap.className = "fs-backstage__open-card-icon";
  iconWrap.appendChild(iconDocument());
  const cardLabel = document.createElement("span");
  cardLabel.textContent = "打开 JSON 文件";
  card.appendChild(iconWrap);
  card.appendChild(cardLabel);

  const runLoad = (text: string): void => {
    const importOpts: FlexSheetJsonImportOptions = {
      includeStyles: openStyles.input.checked,
      includeFormulas: openFormulas.input.checked,
      recalcAfterImport: openRecalc.input.checked,
    };
    loadFlexSheetJsonFromText(text, fs, importOpts, msg, "已打开工作簿。", onSuccessClose);
  };

  const pickFile = (): void => {
    msg.hidden = true;
    fileInput.click();
  };

  card.addEventListener("click", pickFile);

  card.addEventListener("dragover", (e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    card.classList.add("fs-backstage__open-card--dragover");
  });
  card.addEventListener("dragleave", () => {
    card.classList.remove("fs-backstage__open-card--dragover");
  });
  card.addEventListener("drop", (e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    card.classList.remove("fs-backstage__open-card--dragover");
    const f = e.dataTransfer?.files[0];
    if (f === undefined) {
      return;
    }
    const nameOk = f.name.toLowerCase().endsWith(".json");
    const typeOk = f.type === "application/json" || f.type === "";
    if (!nameOk && !typeOk) {
      msg.textContent = "请选择 .json 文件。";
      msg.classList.add("fs-backstage__msg--error");
      msg.hidden = false;
      return;
    }
    readTextFile(f, msg, runLoad);
  });

  fileInput.addEventListener("change", () => {
    const file = fileInput.files?.[0];
    fileInput.value = "";
    if (file === undefined) {
      return;
    }
    readTextFile(file, msg, runLoad);
  });

  container.appendChild(title);
  container.appendChild(sub);
  container.appendChild(hint);
  container.appendChild(groups);
  container.appendChild(fileInput);
  container.appendChild(card);
  container.appendChild(msg);
}

function mountSavePanel(container: HTMLElement, flexSheet: FlexSheetLike | undefined): void {
  container.innerHTML = "";
  const title = document.createElement("h1");
  title.className = "fs-backstage__page-title";
  title.textContent = "保存";

  const sub = document.createElement("p");
  sub.className = "fs-backstage__page-sub";
  sub.textContent = "FlexSheet 工作簿 (.json)";

  if (flexSheet === undefined) {
    const hint = document.createElement("p");
    hint.className = "fs-backstage__placeholder";
    hint.textContent = "当前未连接工作簿，无法保存。";
    container.appendChild(title);
    container.appendChild(sub);
    container.appendChild(hint);
    return;
  }
  const workbook = flexSheet.workbook;
  if (workbook === undefined) {
    const hint = document.createElement("p");
    hint.className = "fs-backstage__placeholder";
    hint.textContent = "当前未连接工作簿，无法保存。";
    container.appendChild(title);
    container.appendChild(sub);
    container.appendChild(hint);
    return;
  }

  const fs = flexSheet;

  const list = document.createElement("div");
  list.className = "fs-backstage__save-options";

  const o = DEFAULT_FLEXSHEET_JSON_EXPORT_OPTIONS;
  const c1 = createCheckboxRow("包含样式", o.includeStyles);
  const c2 = createCheckboxRow("包含公式", o.includeFormulas);
  const c3 = createCheckboxRow("包含未使用名称", o.includeUnusedNames);
  const c4 = createCheckboxRow("按视图保存", o.saveByView);
  const c5 = createCheckboxRow("包含自动合并的单元格", o.includeAutoMergedCells);
  const c6 = createCheckboxRow("包含计算缓存", o.includeCalculationCache);
  const c7 = createCheckboxRow("保存公式为R1C1格式", o.saveFormulasAsR1C1);
  const c8 = createCheckboxRow("包含绑定数据源", o.includeBoundDataSources);
  const c9 = createCheckboxRow("包含有样式但无数据的最小范围单元格", o.includeSparseStyledEmpty);

  list.appendChild(c1.row);
  list.appendChild(c2.row);
  list.appendChild(c3.row);
  list.appendChild(c4.row);
  list.appendChild(c5.row);
  list.appendChild(c6.row);
  list.appendChild(c7.row);
  list.appendChild(c8.row);
  list.appendChild(c9.row);

  const actions = document.createElement("div");
  actions.className = "fs-backstage__save-actions";

  const nameRow = document.createElement("div");
  nameRow.className = "fs-backstage__filename-row";

  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.className = "fs-backstage__filename-input";
  nameInput.setAttribute("aria-label", "文件名");
  const active = workbook.getActiveSheet();
  nameInput.value = active !== undefined ? active.name.replace(/[^\w\u4e00-\u9fa5.-]+/g, "_") : "workbook";
  nameInput.placeholder = "文件名";

  const ext = document.createElement("span");
  ext.className = "fs-backstage__filename-ext";
  ext.textContent = ".json";

  nameRow.appendChild(nameInput);
  nameRow.appendChild(ext);

  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "fs-backstage__btn-primary";
  saveBtn.textContent = "保存";

  saveBtn.addEventListener("click", () => {
    const opts: FlexSheetJsonExportOptions = {
      includeStyles: c1.input.checked,
      includeFormulas: c2.input.checked,
      includeUnusedNames: c3.input.checked,
      saveByView: c4.input.checked,
      includeAutoMergedCells: c5.input.checked,
      includeCalculationCache: c6.input.checked,
      saveFormulasAsR1C1: c7.input.checked,
      includeBoundDataSources: c8.input.checked,
      includeSparseStyledEmpty: c9.input.checked,
    };
    const viewState = c4.input.checked ? captureViewState(fs) : undefined;
    const doc = serializeWorkbookToJsonDocument(workbook, opts, {
      generatorVersion: JSON_GENERATOR_VERSION,
      ...(viewState !== undefined ? { viewState } : {}),
    });
    const text = JSON.stringify(doc, null, 2);
    const base = nameInput.value.trim() || "workbook";
    downloadJsonText(text, base);
  });

  actions.appendChild(nameRow);
  actions.appendChild(saveBtn);

  container.appendChild(title);
  container.appendChild(sub);
  container.appendChild(list);
  container.appendChild(actions);
}

function mountImportPanel(
  container: HTMLElement,
  flexSheet: FlexSheetLike | undefined,
  onSuccessClose: () => void,
): void {
  container.innerHTML = "";
  container.classList.add("fs-ribbon-backstage__panel--import");

  const msg = document.createElement("p");
  msg.className = "fs-backstage__msg";
  msg.setAttribute("role", "status");
  msg.hidden = true;

  if (flexSheet === undefined) {
    const title = document.createElement("h1");
    title.className = "fs-backstage__page-title";
    title.textContent = "导入";
    const ph = document.createElement("p");
    ph.className = "fs-backstage__placeholder";
    ph.textContent = "当前未连接工作簿，无法导入。";
    container.appendChild(title);
    container.appendChild(ph);
    return;
  }
  if (flexSheet.workbook === undefined || flexSheet.loadWorkbook === undefined) {
    const title = document.createElement("h1");
    title.className = "fs-backstage__page-title";
    title.textContent = "导入";
    const ph = document.createElement("p");
    ph.className = "fs-backstage__placeholder";
    ph.textContent = "当前未连接工作簿，无法导入。";
    container.appendChild(title);
    container.appendChild(ph);
    return;
  }

  const fs = flexSheet;

  const shell = document.createElement("div");
  shell.className = "fs-backstage__import-shell";

  const pageTitle = document.createElement("h1");
  pageTitle.className = "fs-backstage__page-title";
  pageTitle.textContent = "导入";

  const body = document.createElement("div");
  body.className = "fs-backstage__import-body";

  const detailCol = document.createElement("div");
  detailCol.className = "fs-backstage__import-detail";

  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.className = "fs-backstage__file-input";
  fileInput.hidden = true;
  let pickHandler: ((f: File) => void) | null = null;
  fileInput.addEventListener("change", () => {
    const f = fileInput.files?.[0];
    fileInput.value = "";
    if (f === undefined || pickHandler === null) {
      return;
    }
    const h = pickHandler;
    pickHandler = null;
    h(f);
  });

  function openFilePicker(accept: string, handler: (f: File) => void): void {
    msg.hidden = true;
    pickHandler = handler;
    fileInput.accept = accept;
    fileInput.click();
  }

  function mountExcelDetail(): void {
    detailCol.innerHTML = "";
    const sub = document.createElement("p");
    sub.className = "fs-backstage__page-sub";
    sub.textContent = "Excel 文件";

    const openOpts = document.createElement("div");
    openOpts.className = "fs-backstage__import-section-title";
    openOpts.textContent = "打开选项";
    const gOpen = document.createElement("div");
    gOpen.className = "fs-backstage__import-check-grid";
    gOpen.appendChild(createCheckboxRow("增量加载", false).row);
    gOpen.appendChild(createCheckboxRow("懒加载", false).row);

    const gStyle = document.createElement("div");
    gStyle.className = "fs-backstage__import-check-grid";
    gStyle.appendChild(createCheckboxRow("包含样式", true).row);
    gStyle.appendChild(createCheckboxRow("包含公式", true).row);
    gStyle.appendChild(createCheckboxRow("导入后自动计算", false).row);

    const gAdv = document.createElement("div");
    gAdv.className = "fs-backstage__import-check-grid";
    gAdv.appendChild(createCheckboxRow("导入未使用的命名样式", true).row);
    gAdv.appendChild(createCheckboxRow("增量计算", true).row);
    gAdv.appendChild(createCheckboxRow("按需计算", true).row);
    gAdv.appendChild(createCheckboxRow("动态引用", false).row);
    gAdv.appendChild(createCheckboxRow("将工作表中的表格转换为数据表", false).row);

    const gHead = document.createElement("div");
    gHead.className = "fs-backstage__import-check-grid";
    gHead.appendChild(createCheckboxRow("导入作为列标题的冻结行", false).row);
    gHead.appendChild(createCheckboxRow("导入作为行标题的冻结列", false).row);

    const pwdRow = document.createElement("div");
    pwdRow.className = "fs-backstage__import-field-row";
    const pwdLabel = document.createElement("label");
    pwdLabel.className = "fs-backstage__import-field-label";
    pwdLabel.textContent = "密码";
    const pwdInput = document.createElement("input");
    pwdInput.type = "password";
    pwdInput.className = "fs-backstage__filename-input";
    pwdInput.disabled = true;
    pwdInput.title = "当前版本不支持加密 Excel 文件";
    pwdInput.setAttribute("aria-label", "密码");
    pwdRow.appendChild(pwdLabel);
    pwdRow.appendChild(pwdInput);

    const card = document.createElement("button");
    card.type = "button";
    card.className = "fs-backstage__open-card";
    const iconWrap = document.createElement("div");
    iconWrap.className = "fs-backstage__open-card-icon";
    iconWrap.appendChild(iconDocument());
    const cardLabel = document.createElement("span");
    cardLabel.textContent = "导入 Excel 文件";
    card.appendChild(iconWrap);
    card.appendChild(cardLabel);
    card.addEventListener("click", () => {
      openFilePicker(
        ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        (f) => {
          msg.hidden = true;
          if (pwdInput.value.trim().length > 0) {
            showImportError(msg, "当前版本不支持加密 Excel，请移除密码后重试。");
            return;
          }
          void f
            .arrayBuffer()
            .then((buf) =>
              importXlsxToWorkbook(
                new Blob([buf], {
                  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                }),
              ),
            )
            .then((wb) => {
              const lw = fs.loadWorkbook;
              if (lw === undefined) {
                return;
              }
              lw.call(fs, wb);
              onSuccessClose();
            })
            .catch((e: unknown) => {
              showImportError(msg, e instanceof Error ? e.message : String(e));
            });
        },
      );
    });

    detailCol.appendChild(sub);
    detailCol.appendChild(openOpts);
    detailCol.appendChild(gOpen);
    detailCol.appendChild(gStyle);
    detailCol.appendChild(gAdv);
    detailCol.appendChild(gHead);
    detailCol.appendChild(pwdRow);
    detailCol.appendChild(card);
  }

  body.appendChild(detailCol);
  shell.appendChild(pageTitle);
  shell.appendChild(body);

  container.appendChild(shell);
  container.appendChild(fileInput);
  container.appendChild(msg);

  mountExcelDetail();
}

function mountExportPanel(container: HTMLElement, flexSheet: FlexSheetLike | undefined): void {
  container.innerHTML = "";
  container.classList.add("fs-ribbon-backstage__panel--import");

  const msg = document.createElement("p");
  msg.className = "fs-backstage__msg";
  msg.setAttribute("role", "status");
  msg.hidden = true;

  if (flexSheet === undefined || flexSheet.workbook === undefined) {
    const title = document.createElement("h1");
    title.className = "fs-backstage__page-title";
    title.textContent = "导出";
    const ph = document.createElement("p");
    ph.className = "fs-backstage__placeholder";
    ph.textContent = "当前未连接工作簿，无法导出。";
    container.appendChild(title);
    container.appendChild(ph);
    return;
  }

  const workbook = flexSheet.workbook;

  const shell = document.createElement("div");
  shell.className = "fs-backstage__import-shell";

  const pageTitle = document.createElement("h1");
  pageTitle.className = "fs-backstage__page-title";
  pageTitle.textContent = "导出";

  const body = document.createElement("div");
  body.className = "fs-backstage__import-body";

  const detailCol = document.createElement("div");
  detailCol.className = "fs-backstage__import-detail";

  function mountExportExcelDetail(): void {
    detailCol.innerHTML = "";
    const sub = document.createElement("p");
    sub.className = "fs-backstage__page-sub";
    sub.textContent = "Excel 文件";

    const saveTitle = document.createElement("div");
    saveTitle.className = "fs-backstage__import-section-title";
    saveTitle.textContent = "保存选项";

    const g = document.createElement("div");
    g.className = "fs-backstage__import-check-grid";
    const cStyles = createCheckboxRow("包含样式", true);
    const cFormulas = createCheckboxRow("包含公式", true);
    const cSparse = createCheckboxRow("包含有样式但无数据的最小范围单元格", true);
    g.appendChild(cStyles.row);
    g.appendChild(cFormulas.row);
    g.appendChild(createCheckboxRow("包含未使用名称", true).row);
    g.appendChild(createCheckboxRow("作为预览导出", false).row);
    g.appendChild(createCheckboxRow("包含绑定数据源", false).row);
    g.appendChild(cSparse.row);
    g.appendChild(createCheckboxRow("包含自动合并的单元格", false).row);
    g.appendChild(createCheckboxRow("包含无损编辑内容", true).row);

    const fzTitle = document.createElement("div");
    fzTitle.className = "fs-backstage__import-section-title";
    fzTitle.textContent = "冻结窗格";
    const gFz = document.createElement("div");
    gFz.className = "fs-backstage__import-check-grid";
    gFz.appendChild(createCheckboxRow("将行标题作为冻结列导出", false).row);
    gFz.appendChild(createCheckboxRow("将列标题作为冻结行导出", false).row);

    const pwdRow = document.createElement("div");
    pwdRow.className = "fs-backstage__import-field-row";
    const pwdLabel = document.createElement("label");
    pwdLabel.className = "fs-backstage__import-field-label";
    pwdLabel.textContent = "密码";
    const pwdInput = document.createElement("input");
    pwdInput.type = "password";
    pwdInput.className = "fs-backstage__filename-input";
    pwdInput.title = "填写密码后不会生成加密文件，仅阻止导出并提示";
    pwdInput.setAttribute("aria-label", "密码");
    pwdRow.appendChild(pwdLabel);
    pwdRow.appendChild(pwdInput);

    const card = document.createElement("button");
    card.type = "button";
    card.className = "fs-backstage__open-card";
    const iconWrap = document.createElement("div");
    iconWrap.className = "fs-backstage__open-card-icon";
    iconWrap.appendChild(iconDocument());
    const cardLabel = document.createElement("span");
    cardLabel.textContent = "导出 Excel 文件";
    card.appendChild(iconWrap);
    card.appendChild(cardLabel);
    card.addEventListener("click", () => {
      msg.hidden = true;
      if (pwdInput.value.trim().length > 0) {
        showImportError(msg, "当前版本不支持密码保护导出，请清空密码。");
        return;
      }
      try {
        const xlsxOpts: XlsxExportOptions = {
          includeStyles: cStyles.input.checked,
          includeFormulas: cFormulas.input.checked,
          includeSparseStyledEmpty: cSparse.input.checked,
        };
        const blob = exportWorkbookToXlsxBlob(workbook, xlsxOpts);
        const active = workbook.getActiveSheet();
        const base =
          active !== undefined ? active.name.replace(/[^\w\u4e00-\u9fa5.-]+/g, "_") : "export";
        downloadXlsxBlob(blob, `${base || "export"}.xlsx`);
      } catch (e: unknown) {
        showImportError(msg, e instanceof Error ? e.message : String(e));
      }
    });

    detailCol.appendChild(sub);
    detailCol.appendChild(saveTitle);
    detailCol.appendChild(g);
    detailCol.appendChild(fzTitle);
    detailCol.appendChild(gFz);
    detailCol.appendChild(pwdRow);
    detailCol.appendChild(card);
  }

  body.appendChild(detailCol);
  shell.appendChild(pageTitle);
  shell.appendChild(body);

  container.appendChild(shell);
  container.appendChild(msg);

  mountExportExcelDetail();
}

export function createRibbonBackstage(
  onClose: () => void,
  options?: RibbonBackstageOptions,
): RibbonBackstageHandles {
  const flexSheet = options?.flexSheet;
  const root = document.createElement("div");
  root.className = "fs-ribbon-backstage";
  root.setAttribute("role", "dialog");
  root.setAttribute("aria-modal", "true");
  root.setAttribute("aria-label", "文件");

  const sidebar = document.createElement("aside");
  sidebar.className = "fs-ribbon-backstage__sidebar";

  const backWrap = document.createElement("div");
  backWrap.className = "fs-ribbon-backstage__back-wrap";
  const backBtn = document.createElement("button");
  backBtn.type = "button";
  backBtn.className = "fs-ribbon-backstage__back";
  backBtn.setAttribute("aria-label", "返回");
  backBtn.title = "返回";
  backBtn.appendChild(iconBack());
  backBtn.addEventListener("click", () => {
    onClose();
  });
  backWrap.appendChild(backBtn);
  sidebar.appendChild(backWrap);

  const nav = document.createElement("nav");
  nav.className = "fs-ribbon-backstage__nav";
  nav.setAttribute("aria-label", "文件菜单");

  const navButtons = new Map<RibbonBackstageSectionId, HTMLButtonElement>();
  const contentArea = document.createElement("div");
  contentArea.className = "fs-ribbon-backstage__main";

  const panels = new Map<RibbonBackstageSectionId, HTMLElement>();

  for (const id of SECTION_ORDER) {
    const itemWrap = document.createElement("div");
    itemWrap.className = "fs-ribbon-backstage__nav-item-wrap";

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "fs-ribbon-backstage__nav-item";
    btn.textContent = SECTION_LABEL[id];
    btn.dataset.section = id;
    navButtons.set(id, btn);
    itemWrap.appendChild(btn);
    nav.appendChild(itemWrap);

    if (DIVIDER_AFTER[id] === true) {
      const hr = document.createElement("div");
      hr.className = "fs-ribbon-backstage__nav-divider";
      hr.setAttribute("role", "separator");
      nav.appendChild(hr);
    }

    const panel = document.createElement("div");
    panel.className = "fs-ribbon-backstage__panel";
    panel.hidden = id !== "open";
    panel.dataset.section = id;
    panels.set(id, panel);
    contentArea.appendChild(panel);
  }

  function showSection(id: RibbonBackstageSectionId): void {
    for (const [sid, panel] of panels) {
      panel.hidden = sid !== id;
    }
    for (const [sid, b] of navButtons) {
      b.classList.toggle("fs-ribbon-backstage__nav-item--active", sid === id);
    }
    if (id === "open") {
      mountOpenPanel(panels.get("open")!, flexSheet, onClose);
    } else if (id === "save") {
      mountSavePanel(panels.get("save")!, flexSheet);
    } else if (id === "import") {
      mountImportPanel(panels.get("import")!, flexSheet, onClose);
    } else if (id === "export") {
      mountExportPanel(panels.get("export")!, flexSheet);
    }
  }

  for (const id of SECTION_ORDER) {
    const b = navButtons.get(id)!;
    b.addEventListener("click", () => {
      showSection(id);
    });
  }

  mountOpenPanel(panels.get("open")!, flexSheet, onClose);
  navButtons.get("open")!.classList.add("fs-ribbon-backstage__nav-item--active");

  sidebar.appendChild(nav);
  root.appendChild(sidebar);
  root.appendChild(contentArea);

  root.addEventListener("keydown", (ev: KeyboardEvent) => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      onClose();
    }
  });

  return {
    root,
    showSection,
    applyThemeMode(mode: "light" | "dark"): void {
      root.dataset.theme = mode;
    },
  };
}

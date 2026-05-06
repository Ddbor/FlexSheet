import type { SelectionRange, Worksheet } from "@flexsheet/core";
import { attachDraggableDialogPanel, showMessageAlert } from "@flexsheet/shared";

import {
  computeGotoSpecialRange,
  type GotoSpecialKind,
  type GotoSpecialSubtypeFilters,
} from "./goto-special-engine.js";
import { ensureGotoSpecialDialogStyles } from "./fs-dialog-styles.js";

export interface OpenGotoSpecialDialogOptions {
  readonly sheet: Worksheet;
  readonly selectionRange: SelectionRange;
  readonly activeRow: number;
  readonly activeCol: number;
  readonly onApplyRange: (range: SelectionRange) => void;
  readonly onClose: () => void;
}

type KindRow = { readonly kind: GotoSpecialKind; readonly label: string };

/** 左列、右列与 Excel「定位条件」一致。 */
const KIND_ROWS: readonly (readonly [KindRow, KindRow | null])[] = [
  [{ kind: "comments", label: "批注" }, { kind: "precedents", label: "引用单元格" }],
  [{ kind: "constants", label: "常量" }, { kind: "dependents", label: "从属单元格" }],
  [{ kind: "formulas", label: "公式" }, { kind: "lastCell", label: "最后一个单元格" }],
  [{ kind: "blanks", label: "空值" }, { kind: "visibleOnly", label: "仅可见单元格" }],
  [{ kind: "currentRegion", label: "当前区域" }, { kind: "objects", label: "对象" }],
  [{ kind: "currentArray", label: "当前数组" }, { kind: "conditionalFormats", label: "条件格式" }],
  [{ kind: "rowDiff", label: "行内容差异单元格" }, { kind: "dataValidation", label: "数据验证" }],
  [{ kind: "colDiff", label: "列内容差异单元格" }, null],
];

function createRadio(name: string, value: string, label: string, checked: boolean): HTMLLabelElement {
  const lab = document.createElement("label");
  lab.className = "fs-goto-special__radio";
  const input = document.createElement("input");
  input.type = "radio";
  input.name = name;
  input.value = value;
  input.checked = checked;
  const span = document.createElement("span");
  span.textContent = label;
  lab.appendChild(input);
  lab.appendChild(span);
  return lab;
}

function createCheck(label: string, checked: boolean): { wrap: HTMLLabelElement; input: HTMLInputElement } {
  const lab = document.createElement("label");
  lab.className = "fs-goto-special__check";
  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = checked;
  const span = document.createElement("span");
  span.textContent = label;
  lab.appendChild(input);
  lab.appendChild(span);
  return { wrap: lab, input };
}

function readSubtypeFilters(
  n: HTMLInputElement,
  t: HTMLInputElement,
  l: HTMLInputElement,
  e: HTMLInputElement,
): GotoSpecialSubtypeFilters {
  return {
    numbers: n.checked,
    text: t.checked,
    logicals: l.checked,
    errors: e.checked,
  };
}

export function openGotoSpecialDialogWithOverlay(options: OpenGotoSpecialDialogOptions): HTMLDivElement {
  ensureGotoSpecialDialogStyles();
  const { sheet, selectionRange, activeRow, activeCol, onApplyRange, onClose } = options;

  let selectedKind: GotoSpecialKind = "comments";
  const subtypeNumbers = createCheck("数字", true);
  const subtypeText = createCheck("文本", true);
  const subtypeLogical = createCheck("逻辑值", true);
  const subtypeErrors = createCheck("错误", true);
  const linkAll = createRadio("fs-goto-special-link", "all", "所有级别", false);
  const linkDirect = createRadio("fs-goto-special-link", "direct", "仅直接引用", true);

  const overlay = document.createElement("div");
  overlay.className = "fs-goto-special-overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-goto-special";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-goto-special-title");
  panel.tabIndex = -1;
  overlay.appendChild(panel);

  const titlebar = document.createElement("div");
  titlebar.className = "fs-goto-special__titlebar";
  const title = document.createElement("div");
  title.id = "fs-goto-special-title";
  title.className = "fs-goto-special__title";
  title.textContent = "定位条件";
  title.title = "按住标题可拖动";
  titlebar.appendChild(title);
  panel.appendChild(titlebar);

  const body = document.createElement("div");
  body.className = "fs-goto-special__body";
  panel.appendChild(body);

  const chooseLabel = document.createElement("div");
  chooseLabel.className = "fs-goto-special__section-label";
  chooseLabel.textContent = "选择";
  body.appendChild(chooseLabel);

  const kindGrid = document.createElement("div");
  kindGrid.className = "fs-goto-special__kind-grid";
  const kindInputs = new Map<GotoSpecialKind, HTMLInputElement>();

  for (const pair of KIND_ROWS) {
    const lab0 = createRadio("fs-goto-special-kind", pair[0]!.kind, pair[0]!.label, pair[0]!.kind === selectedKind);
    kindInputs.set(pair[0]!.kind, lab0.querySelector("input") as HTMLInputElement);
    kindGrid.appendChild(lab0);
    if (pair[1] !== null) {
      const lab1 = createRadio("fs-goto-special-kind", pair[1]!.kind, pair[1]!.label, pair[1]!.kind === selectedKind);
      kindInputs.set(pair[1]!.kind, lab1.querySelector("input") as HTMLInputElement);
      kindGrid.appendChild(lab1);
    } else {
      const ph = document.createElement("div");
      ph.className = "fs-goto-special__kind-placeholder";
      kindGrid.appendChild(ph);
    }
  }
  body.appendChild(kindGrid);

  const hr = document.createElement("div");
  hr.className = "fs-goto-special__hr";
  body.appendChild(hr);

  const optTitle = document.createElement("div");
  optTitle.className = "fs-goto-special__section-label";
  optTitle.textContent = "选项";
  body.appendChild(optTitle);

  const optionsSlot = document.createElement("div");
  optionsSlot.className = "fs-goto-special__options";
  body.appendChild(optionsSlot);

  const foot = document.createElement("div");
  foot.className = "fs-goto-special__footer";
  const btnCancel = document.createElement("button");
  btnCancel.type = "button";
  btnCancel.className = "fs-goto-special__btn";
  btnCancel.textContent = "取消";
  const btnOk = document.createElement("button");
  btnOk.type = "button";
  btnOk.className = "fs-goto-special__btn fs-goto-special__btn--primary";
  btnOk.textContent = "确定";
  foot.appendChild(btnCancel);
  foot.appendChild(btnOk);
  panel.appendChild(foot);

  const noopHint = document.createElement("div");
  noopHint.className = "fs-goto-special__noop";
  noopHint.textContent = "没有适合此选择的选项。";

  const subtypeWrap = document.createElement("div");
  subtypeWrap.className = "fs-goto-special__subtype-grid";
  subtypeWrap.appendChild(subtypeNumbers.wrap);
  subtypeWrap.appendChild(subtypeText.wrap);
  subtypeWrap.appendChild(subtypeLogical.wrap);
  subtypeWrap.appendChild(subtypeErrors.wrap);

  const linkWrap = document.createElement("div");
  linkWrap.className = "fs-goto-special__link-stack";
  linkWrap.appendChild(linkAll);
  linkWrap.appendChild(linkDirect);

  function syncKindRadios(): void {
    for (const [k, inp] of kindInputs) {
      inp.checked = k === selectedKind;
    }
  }

  function renderOptions(): void {
    while (optionsSlot.firstChild !== null) {
      optionsSlot.removeChild(optionsSlot.firstChild);
    }
    switch (selectedKind) {
      case "comments":
      case "objects":
      case "dataValidation":
        optionsSlot.appendChild(noopHint);
        return;
      case "constants":
      case "formulas":
        optionsSlot.appendChild(subtypeWrap);
        return;
      case "precedents":
      case "dependents":
        optionsSlot.appendChild(linkWrap);
        return;
      default:
        optionsSlot.appendChild(noopHint);
    }
  }

  function onKindChange(kind: GotoSpecialKind): void {
    selectedKind = kind;
    syncKindRadios();
    renderOptions();
  }

  for (const [k, inp] of kindInputs) {
    inp.addEventListener("change", () => {
      if (inp.checked) {
        onKindChange(k);
      }
    });
  }

  renderOptions();

  function readDirectLinksOnly(): boolean {
    const v = (linkDirect.querySelector("input") as HTMLInputElement).checked;
    return v;
  }

  const onKey = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      detachKey();
      overlay.remove();
      onClose();
    }
  };
  const detachKey = (): void => {
    document.removeEventListener("keydown", onKey, true);
  };

  function applyOk(): void {
    const filters = readSubtypeFilters(
      subtypeNumbers.input,
      subtypeText.input,
      subtypeLogical.input,
      subtypeErrors.input,
    );
    if (
      (selectedKind === "constants" || selectedKind === "formulas") &&
      !filters.numbers &&
      !filters.text &&
      !filters.logicals &&
      !filters.errors
    ) {
      showMessageAlert("请至少选择一种类型（数字、文本、逻辑值或错误）。");
      return;
    }

    const res = computeGotoSpecialRange({
      sheet,
      selectionRange,
      activeRow,
      activeCol,
      kind: selectedKind,
      subtypeFilters: filters,
      directLinksOnly: readDirectLinksOnly(),
    });

    if (!res.ok) {
      if (res.code === "unsupported") {
        showMessageAlert("当前版本暂不支持对「批注」「对象」或「数据验证」的定位。");
      } else if (res.code === "singleCell") {
        showMessageAlert("请先选定一块矩形区域（至少两个单元格）。");
      } else {
        showMessageAlert("未找到符合条件的单元格。");
      }
      detachKey();
      overlay.remove();
      onClose();
      return;
    }
    detachKey();
    onApplyRange(res.range);
    overlay.remove();
    onClose();
  }

  btnCancel.addEventListener("click", () => {
    detachKey();
    overlay.remove();
    onClose();
  });
  btnOk.addEventListener("click", () => {
    applyOk();
  });

  overlay.addEventListener("click", (ev) => {
    if (ev.target === overlay) {
      detachKey();
      overlay.remove();
      onClose();
    }
  });

  attachDraggableDialogPanel(panel, titlebar);
  document.addEventListener("keydown", onKey, true);

  requestAnimationFrame(() => {
    panel.focus();
  });

  return overlay;
}

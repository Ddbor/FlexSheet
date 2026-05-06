/**
 * Canvas 右侧「设置图片格式」面板：图片校正 / 颜色 / 透明度 / 裁剪占位数值。
 */

import type { FloatingPictureAdjustments } from "./floating-picture-layer.js";
import { DEFAULT_FLOATING_PICTURE_ADJUSTMENTS } from "./floating-picture-layer.js";

export interface FormatPicturePaneLayout {
  readonly widthPx: number;
  readonly heightPx: number;
  readonly offsetXPx: number;
  readonly offsetYPx: number;
}

export interface FormatPicturePaneController {
  readonly root: HTMLElement;
  show(): void;
  hide(): void;
  isOpen(): boolean;
  /** 从当前选中图同步控件（选中变化或外部改动画后调用） */
  syncFromModel(): void;
  destroy(): void;
}

export interface CreateFormatPicturePaneOptions {
  readonly parent: HTMLElement;
  readonly getAdjustments: () => FloatingPictureAdjustments | null;
  readonly getLayout: () => FormatPicturePaneLayout | null;
  readonly setAdjustments: (patch: Partial<FloatingPictureAdjustments>) => void;
  readonly onClose: () => void;
}

let stylesInjected = false;

function ensureFormatPictureStyles(): void {
  if (stylesInjected) {
    return;
  }
  stylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-format-picture", "1");
  style.textContent = `
.fs-format-picture {
  box-sizing: border-box;
  display: none;
  flex-direction: column;
  width: 300px;
  flex-shrink: 0;
  max-height: 100%;
  min-height: 0;
  border-left: 1px solid var(--fs-ribbon-border, #e1dfdd);
  background: var(--fs-ribbon-panel-bg, #f3f2f1);
  color: var(--fs-ribbon-chrome-text, #323130);
  font-size: 12px;
}
.fs-format-picture[data-open="1"] {
  display: flex;
}
.fs-format-picture__head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 8px 10px;
  border-bottom: 1px solid var(--fs-ribbon-border, #e1dfdd);
  flex-shrink: 0;
}
.fs-format-picture__title {
  font-weight: 600;
  font-size: 13px;
}
.fs-format-picture__close {
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 18px;
  line-height: 1;
  padding: 2px 6px;
  border-radius: 2px;
  color: inherit;
}
.fs-format-picture__close:hover {
  background: var(--fs-ribbon-hover, #edebe9);
}
.fs-format-picture__tabs {
  display: flex;
  flex-direction: row;
  border-bottom: 1px solid var(--fs-ribbon-border, #e1dfdd);
  flex-shrink: 0;
}
.fs-format-picture__tab {
  flex: 1;
  border: none;
  background: transparent;
  padding: 8px 4px;
  cursor: pointer;
  font-size: 11px;
  color: inherit;
  opacity: 0.65;
  border-bottom: 2px solid transparent;
}
.fs-format-picture__tab:hover {
  background: var(--fs-ribbon-hover, #edebe9);
}
.fs-format-picture__tab--active {
  opacity: 1;
  border-bottom-color: #217346;
  font-weight: 600;
}
.fs-format-picture__scroll {
  flex: 1;
  min-height: 0;
  overflow-y: auto;
  padding: 6px 8px 12px;
}
.fs-format-picture details.fs-format-picture__sec {
  border: 1px solid var(--fs-ribbon-border, #e1dfdd);
  border-radius: 2px;
  margin-bottom: 8px;
  background: #fff;
  background: var(--fs-sheet-surface, #fff);
}
.fs-format-picture details.fs-format-picture__sec > summary {
  list-style: none;
  padding: 8px 10px;
  cursor: pointer;
  font-weight: 600;
  user-select: none;
}
.fs-format-picture details.fs-format-picture__sec > summary::-webkit-details-marker {
  display: none;
}
.fs-format-picture__sec-body {
  padding: 0 10px 10px;
}
.fs-format-picture__row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 8px;
}
.fs-format-picture__row label {
  flex: 0 0 72px;
  min-width: 0;
}
.fs-format-picture__row input[type="range"] {
  flex: 1;
  min-width: 0;
}
.fs-format-picture__num {
  width: 52px;
  flex-shrink: 0;
  box-sizing: border-box;
  padding: 2px 4px;
  font-size: 11px;
  border: 1px solid var(--fs-ribbon-border, #c8c6c4);
  border-radius: 2px;
}
.fs-format-picture__subhead {
  font-size: 11px;
  font-weight: 600;
  margin: 10px 0 6px;
  color: inherit;
  opacity: 0.85;
}
.fs-format-picture__reset {
  margin-top: 6px;
  padding: 4px 10px;
  font-size: 11px;
  border: 1px solid var(--fs-ribbon-border, #c8c6c4);
  border-radius: 2px;
  background: var(--fs-ribbon-panel-bg, #faf9f8);
  cursor: pointer;
}
.fs-format-picture__reset:hover {
  background: var(--fs-ribbon-hover, #edebe9);
}
.fs-format-picture__grid2 {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 6px 8px;
  align-items: center;
  margin-top: 6px;
}
.fs-format-picture__grid2 label {
  font-size: 11px;
}
.fs-format-picture__grid2 input.fs-format-picture__num {
  width: 100%;
}
.fs-format-picture__placeholder {
  padding: 24px 12px;
  text-align: center;
  opacity: 0.6;
  font-size: 12px;
}
`;
  document.head.appendChild(style);
}

function pxToIn(px: number): string {
  return (px / 96).toFixed(2);
}

export function createFormatPicturePane(
  options: CreateFormatPicturePaneOptions,
): FormatPicturePaneController {
  ensureFormatPictureStyles();
  let open = false;
  let syncing = false;

  const root = document.createElement("aside");
  root.className = "fs-format-picture";
  root.setAttribute("aria-label", "设置图片格式");

  const head = document.createElement("div");
  head.className = "fs-format-picture__head";
  const title = document.createElement("div");
  title.className = "fs-format-picture__title";
  title.textContent = "设置图片格式";
  const btnClose = document.createElement("button");
  btnClose.type = "button";
  btnClose.className = "fs-format-picture__close";
  btnClose.setAttribute("aria-label", "关闭");
  btnClose.textContent = "×";
  head.appendChild(title);
  head.appendChild(btnClose);

  const tabs = document.createElement("div");
  tabs.className = "fs-format-picture__tabs";
  const tabIds = ["填充", "效果", "大小", "图片"] as const;
  const tabBtns: HTMLButtonElement[] = [];
  for (let i = 0; i < tabIds.length; i++) {
    const t = document.createElement("button");
    t.type = "button";
    t.className = "fs-format-picture__tab" + (i === 3 ? " fs-format-picture__tab--active" : "");
    t.textContent = tabIds[i];
    tabBtns.push(t);
    tabs.appendChild(t);
  }

  const scroll = document.createElement("div");
  scroll.className = "fs-format-picture__scroll";

  const panelPicture = document.createElement("div");
  panelPicture.className = "fs-format-picture__panel-picture";

  const panelOther = document.createElement("div");
  panelOther.className = "fs-format-picture__panel-other";
  panelOther.hidden = true;
  const ph = document.createElement("div");
  ph.className = "fs-format-picture__placeholder";
  ph.textContent = "此分类即将支持";
  panelOther.appendChild(ph);

  function showPanel(idx: number): void {
    const isPic = idx === 3;
    panelPicture.hidden = !isPic;
    panelOther.hidden = isPic;
    for (let i = 0; i < tabBtns.length; i++) {
      tabBtns[i].classList.toggle("fs-format-picture__tab--active", i === idx);
    }
  }

  for (let i = 0; i < tabBtns.length; i++) {
    const idx = i;
    tabBtns[i].addEventListener("click", () => showPanel(idx));
  }

  function mkSec(summaryText: string): { details: HTMLDetailsElement; body: HTMLDivElement } {
    const details = document.createElement("details");
    details.className = "fs-format-picture__sec";
    details.open = true;
    const sum = document.createElement("summary");
    sum.textContent = summaryText;
    const body = document.createElement("div");
    body.className = "fs-format-picture__sec-body";
    details.appendChild(sum);
    details.appendChild(body);
    return { details, body };
  }

  function mkSliderRow(
    body: HTMLDivElement,
    label: string,
    min: number,
    max: number,
    step: number,
    key: keyof FloatingPictureAdjustments,
    format: (v: number) => string,
    parse: (s: string) => number | null,
  ): { range: HTMLInputElement; num: HTMLInputElement } {
    const row = document.createElement("div");
    row.className = "fs-format-picture__row";
    const lab = document.createElement("label");
    lab.textContent = label;
    const range = document.createElement("input");
    range.type = "range";
    range.min = String(min);
    range.max = String(max);
    range.step = String(step);
    const num = document.createElement("input");
    num.type = "text";
    num.className = "fs-format-picture__num";
    const applyFromRange = (): void => {
      if (syncing) {
        return;
      }
      const v = Number(range.value);
      options.setAdjustments({ [key]: v } as Partial<FloatingPictureAdjustments>);
      num.value = format(v);
    };
    const applyFromNum = (): void => {
      if (syncing) {
        return;
      }
      const v = parse(num.value);
      if (v === null || !Number.isFinite(v)) {
        return;
      }
      const cl = Math.min(max, Math.max(min, v));
      options.setAdjustments({ [key]: cl } as Partial<FloatingPictureAdjustments>);
      range.value = String(cl);
      num.value = format(cl);
    };
    range.addEventListener("input", applyFromRange);
    num.addEventListener("change", applyFromNum);
    num.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        applyFromNum();
      }
    });
    row.appendChild(lab);
    row.appendChild(range);
    row.appendChild(num);
    body.appendChild(row);
    return { range, num };
  }

  /* —— 图片校正 —— */
  const secCorr = mkSec("图片校正");
  panelPicture.appendChild(secCorr.details);
  const sh = document.createElement("div");
  sh.className = "fs-format-picture__subhead";
  sh.textContent = "锐化 / 柔化";
  secCorr.body.appendChild(sh);
  const sharpSl = mkSliderRow(
    secCorr.body,
    "清晰度",
    -100,
    100,
    1,
    "sharpnessPct",
    (v) => `${v}%`,
    (s) => {
      const n = Number(s.replace(/%/g, "").trim());
      return Number.isFinite(n) ? n : null;
    },
  );
  const sh2 = document.createElement("div");
  sh2.className = "fs-format-picture__subhead";
  sh2.textContent = "亮度 / 对比度";
  secCorr.body.appendChild(sh2);
  const brSl = mkSliderRow(
    secCorr.body,
    "亮度",
    -100,
    100,
    1,
    "brightnessPct",
    (v) => `${v}%`,
    (s) => {
      const n = Number(s.replace(/%/g, "").trim());
      return Number.isFinite(n) ? n : null;
    },
  );
  const ctSl = mkSliderRow(
    secCorr.body,
    "对比度",
    -100,
    100,
    1,
    "contrastPct",
    (v) => `${v}%`,
    (s) => {
      const n = Number(s.replace(/%/g, "").trim());
      return Number.isFinite(n) ? n : null;
    },
  );
  const resetCorr = document.createElement("button");
  resetCorr.type = "button";
  resetCorr.className = "fs-format-picture__reset";
  resetCorr.textContent = "重置";
  resetCorr.addEventListener("click", () => {
    options.setAdjustments({
      sharpnessPct: DEFAULT_FLOATING_PICTURE_ADJUSTMENTS.sharpnessPct,
      brightnessPct: DEFAULT_FLOATING_PICTURE_ADJUSTMENTS.brightnessPct,
      contrastPct: DEFAULT_FLOATING_PICTURE_ADJUSTMENTS.contrastPct,
    });
    syncFromModel();
  });
  secCorr.body.appendChild(resetCorr);

  /* —— 图片颜色 —— */
  const secCol = mkSec("图片颜色");
  panelPicture.appendChild(secCol.details);
  const satSl = mkSliderRow(
    secCol.body,
    "饱和度",
    0,
    400,
    1,
    "saturationPct",
    (v) => `${v}%`,
    (s) => {
      const n = Number(s.replace(/%/g, "").trim());
      return Number.isFinite(n) ? n : null;
    },
  );
  const tempSl = mkSliderRow(
    secCol.body,
    "色温",
    2000,
    11000,
    100,
    "colorTemperatureK",
    (v) => `${Math.round(v)}`,
    (s) => {
      const n = Number(s.replace(/,/g, "").trim());
      return Number.isFinite(n) ? n : null;
    },
  );
  const recolorNote = document.createElement("div");
  recolorNote.className = "fs-format-picture__subhead";
  recolorNote.textContent = "重新着色";
  secCol.body.appendChild(recolorNote);
  const recolorSel = document.createElement("select");
  recolorSel.style.width = "100%";
  recolorSel.style.padding = "4px";
  recolorSel.style.fontSize = "11px";
  recolorSel.innerHTML =
    '<option value="">（无）</option><option value="gray" disabled>灰度（即将支持）</option>';
  secCol.body.appendChild(recolorSel);
  const resetCol = document.createElement("button");
  resetCol.type = "button";
  resetCol.className = "fs-format-picture__reset";
  resetCol.textContent = "重置";
  resetCol.addEventListener("click", () => {
    options.setAdjustments({
      saturationPct: DEFAULT_FLOATING_PICTURE_ADJUSTMENTS.saturationPct,
      colorTemperatureK: DEFAULT_FLOATING_PICTURE_ADJUSTMENTS.colorTemperatureK,
    });
    syncFromModel();
  });
  secCol.body.appendChild(resetCol);

  /* —— 透明度 —— */
  const secTr = mkSec("图片透明度");
  panelPicture.appendChild(secTr.details);
  const trSl = mkSliderRow(
    secTr.body,
    "透明度",
    0,
    100,
    1,
    "transparencyPct",
    (v) => `${v}%`,
    (s) => {
      const n = Number(s.replace(/%/g, "").trim());
      return Number.isFinite(n) ? n : null;
    },
  );
  const resetTr = document.createElement("button");
  resetTr.type = "button";
  resetTr.className = "fs-format-picture__reset";
  resetTr.textContent = "重置";
  resetTr.addEventListener("click", () => {
    options.setAdjustments({ transparencyPct: 0 });
    syncFromModel();
  });
  secTr.body.appendChild(resetTr);

  /* —— 裁剪（数值展示） —— */
  const secCrop = mkSec("裁剪");
  panelPicture.appendChild(secCrop.details);
  const cropGrid = document.createElement("div");
  cropGrid.className = "fs-format-picture__grid2";
  const cropLabels = ["宽度", "高度", "偏移 X", "偏移 Y"] as const;
  const cropInputs: HTMLInputElement[] = [];
  for (let i = 0; i < 4; i++) {
    const lab = document.createElement("label");
    lab.textContent = cropLabels[i];
    const inp = document.createElement("input");
    inp.type = "text";
    inp.className = "fs-format-picture__num";
    inp.readOnly = true;
    cropInputs.push(inp);
    cropGrid.appendChild(lab);
    cropGrid.appendChild(inp);
  }
  secCrop.body.appendChild(cropGrid);
  const cropHint = document.createElement("div");
  cropHint.style.fontSize = "10px";
  cropHint.style.opacity = "0.65";
  cropHint.style.marginTop = "6px";
  cropHint.textContent = "裁剪框数值为只读预览（英寸，96dpi 换算）。";
  secCrop.body.appendChild(cropHint);

  scroll.appendChild(panelPicture);
  scroll.appendChild(panelOther);

  root.appendChild(head);
  root.appendChild(tabs);
  root.appendChild(scroll);
  options.parent.appendChild(root);

  function syncFromModel(): void {
    const a = options.getAdjustments();
    const layout = options.getLayout();
    syncing = true;
    try {
      if (a === null) {
        return;
      }
      sharpSl.range.value = String(a.sharpnessPct);
      sharpSl.num.value = `${a.sharpnessPct}%`;
      brSl.range.value = String(a.brightnessPct);
      brSl.num.value = `${a.brightnessPct}%`;
      ctSl.range.value = String(a.contrastPct);
      ctSl.num.value = `${a.contrastPct}%`;
      satSl.range.value = String(a.saturationPct);
      satSl.num.value = `${a.saturationPct}%`;
      tempSl.range.value = String(a.colorTemperatureK);
      tempSl.num.value = `${Math.round(a.colorTemperatureK)}`;
      trSl.range.value = String(a.transparencyPct);
      trSl.num.value = `${a.transparencyPct}%`;
      if (layout !== null) {
        cropInputs[0].value = `${pxToIn(layout.widthPx)}"`;
        cropInputs[1].value = `${pxToIn(layout.heightPx)}"`;
        cropInputs[2].value = `${pxToIn(layout.offsetXPx)}"`;
        cropInputs[3].value = `${pxToIn(layout.offsetYPx)}"`;
      }
    } finally {
      syncing = false;
    }
  }

  function show(): void {
    open = true;
    root.dataset.open = "1";
    syncFromModel();
  }

  function hide(): void {
    open = false;
    root.removeAttribute("data-open");
  }

  btnClose.addEventListener("click", () => {
    hide();
    options.onClose();
  });

  return {
    root,
    show,
    hide,
    isOpen(): boolean {
      return open;
    },
    syncFromModel,
    destroy(): void {
      hide();
      root.remove();
    },
  };
}

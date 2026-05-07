/**
 * Canvas 右侧「设置图片格式」面板：图片校正 / 颜色 / 透明度 / 裁剪占位数值。
 */

import { appendRibbonColorPaletteContent, showRibbonColorDialog } from "@flexsheet/toolbar";
import type {
  FloatingPictureAdjustments,
  FloatingPictureFillKind,
  FloatingPictureFrameFill,
} from "./floating-picture-layer.js";
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
  readonly getFrameFill: () => FloatingPictureFrameFill | null;
  readonly setFrameFill: (patch: Partial<FloatingPictureFrameFill>) => void;
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
/* 「填充与线条」：第一节可折叠；「线条」区仅占位 */
.fs-format-picture__sec--interactive > summary {
  pointer-events: auto;
  cursor: pointer;
}
.fs-format-picture__panel-fill-line .fs-format-picture__sec--lines-static > summary {
  pointer-events: none;
  cursor: default;
}
.fs-format-picture__fill-radio-row {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 0;
  font-size: 12px;
  color: var(--fs-ribbon-chrome-text, #323130);
  cursor: pointer;
  user-select: none;
}
.fs-format-picture__fill-radio-row:focus-visible {
  outline: 1px solid #217346;
  outline-offset: 1px;
}
.fs-format-picture__solid-extra {
  border-top: 1px solid var(--fs-ribbon-border, #e1dfdd);
  margin-top: 6px;
  padding-top: 10px;
}
/* 与「设置单元格格式」填充页下拉触发器一致，复用 Ribbon 色板 DOM（appendRibbonColorPaletteContent） */
.fs-format-picture__color-dd-row {
  margin-bottom: 10px;
}
.fs-format-picture__color-dd-row > label {
  flex: 0 0 72px;
  min-width: 0;
  font-size: 12px;
}
.fs-format-picture__fill-dd {
  flex: 1;
  min-width: 0;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 10px;
  border: 1px solid var(--fs-ribbon-border, #c8c6c4);
  border-radius: 4px;
  background: var(--fs-sheet-surface, #fff);
  cursor: pointer;
  font-family: inherit;
  font-size: 12px;
  text-align: left;
  color: inherit;
}
.fs-format-picture__fill-dd:hover {
  background: var(--fs-ribbon-hover, #f3f2f1);
}
.fs-format-picture__fill-swatch {
  width: 22px;
  height: 16px;
  border: 1px solid #a19f9d;
  border-radius: 2px;
  flex-shrink: 0;
  box-sizing: border-box;
}
.fs-format-picture__fill-dd-label {
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.fs-format-picture__fill-dd-arrow {
  width: 0;
  height: 0;
  border-left: 5px solid transparent;
  border-right: 5px solid transparent;
  border-top: 6px solid #217346;
  flex-shrink: 0;
}
.fs-format-picture__color-popover.fs-color-menu {
  position: fixed;
  z-index: 10005;
  min-width: 200px;
  max-width: min(320px, calc(100vw - 24px));
  max-height: min(420px, calc(100vh - 48px));
  overflow: auto;
  background: var(--fs-sheet-surface, #fff);
  border: 1px solid var(--fs-ribbon-border, #c8c6c4);
  border-radius: 4px;
  box-shadow: 0 6px 24px rgba(0, 0, 0, 0.18);
  padding: 10px 14px;
  box-sizing: border-box;
}
.fs-format-picture__color-popover .fs-color-menu__heading {
  padding: 6px 2px 4px;
  font-size: 11px;
  color: #605e5c;
}
.fs-format-picture__sec--chrome > summary {
  background: var(--fs-format-picture-sec-head, #edebe9);
  padding: 8px 10px 8px 8px;
  display: flex;
  align-items: center;
  gap: 6px;
}
.fs-format-picture__sec--chrome > summary::before {
  content: "";
  width: 0;
  height: 0;
  border-left: 4px solid transparent;
  border-right: 4px solid transparent;
  border-top: 5px solid var(--fs-ribbon-chrome-text, #323130);
  opacity: 0.75;
  flex-shrink: 0;
  margin-top: 2px;
}
.fs-format-picture__sec--chrome .fs-format-picture__sec-body {
  background: var(--fs-format-picture-sec-body, #faf9f8);
  padding: 10px 12px 12px;
}
.fs-format-picture__radio-static-row {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 0;
  font-size: 12px;
  color: var(--fs-ribbon-chrome-text, #323130);
  user-select: none;
  pointer-events: none;
  cursor: default;
}
.fs-format-picture__radio-disk {
  width: 14px;
  height: 14px;
  border-radius: 50%;
  border: 1px solid #c8c6c4;
  background: #fff;
  flex-shrink: 0;
  box-sizing: border-box;
  position: relative;
}
.fs-format-picture__radio-disk--on {
  border-color: #217346;
  background: #217346;
}
.fs-format-picture__radio-disk--on::after {
  content: "";
  position: absolute;
  left: 50%;
  top: 50%;
  transform: translate(-50%, -50%);
  width: 5px;
  height: 5px;
  border-radius: 50%;
  background: #fff;
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
  let fillColorPopoverCleanup: (() => void) | null = null;

  function closeFillColorPopover(): void {
    if (fillColorPopoverCleanup !== null) {
      fillColorPopoverCleanup();
      fillColorPopoverCleanup = null;
    }
  }

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
  const tabIds = ["填充与线条", "效果", "大小", "图片"] as const;
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

  const panelFillLine = document.createElement("div");
  panelFillLine.className = "fs-format-picture__panel-fill-line";
  panelFillLine.hidden = true;

  function mkStaticRadioRow(label: string, on: boolean): HTMLDivElement {
    const row = document.createElement("div");
    row.className = "fs-format-picture__radio-static-row";
    const disk = document.createElement("span");
    disk.className =
      "fs-format-picture__radio-disk" + (on ? " fs-format-picture__radio-disk--on" : "");
    disk.setAttribute("aria-hidden", "true");
    const lab = document.createElement("span");
    lab.textContent = label;
    row.appendChild(disk);
    row.appendChild(lab);
    return row;
  }

  const fillKindRows = new Map<FloatingPictureFillKind, HTMLDivElement>();
  const fillDetails = document.createElement("details");
  fillDetails.className =
    "fs-format-picture__sec fs-format-picture__sec--chrome fs-format-picture__sec--interactive";
  fillDetails.open = true;
  const fillSum = document.createElement("summary");
  fillSum.textContent = "填充与线条";
  const fillBody = document.createElement("div");
  fillBody.className = "fs-format-picture__sec-body";

  const FILL_OPTIONS: readonly { readonly kind: FloatingPictureFillKind; readonly label: string }[] =
    [
      { kind: "none", label: "无填充" },
      { kind: "solid", label: "纯色填充" },
      { kind: "gradient", label: "渐变填充" },
      { kind: "picture", label: "图片或纹理填充" },
      { kind: "pattern", label: "图案填充" },
    ];

  for (const { kind, label } of FILL_OPTIONS) {
    const row = document.createElement("div");
    row.className = "fs-format-picture__fill-radio-row";
    row.setAttribute("role", "radio");
    row.tabIndex = 0;
    const disk = document.createElement("span");
    disk.className = "fs-format-picture__radio-disk";
    disk.setAttribute("aria-hidden", "true");
    const lab = document.createElement("span");
    lab.textContent = label;
    row.appendChild(disk);
    row.appendChild(lab);
    row.addEventListener("click", () => {
      if (syncing) {
        return;
      }
      options.setFrameFill({ kind });
    });
    row.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        if (!syncing) {
          options.setFrameFill({ kind });
        }
      }
    });
    fillKindRows.set(kind, row);
    fillBody.appendChild(row);
  }

  const solidExtra = document.createElement("div");
  solidExtra.className = "fs-format-picture__solid-extra";
  solidExtra.hidden = true;

  const colorRow = document.createElement("div");
  colorRow.className = "fs-format-picture__row fs-format-picture__color-dd-row";
  const colorLab = document.createElement("label");
  colorLab.textContent = "颜色";
  const colorBtn = document.createElement("button");
  colorBtn.type = "button";
  colorBtn.className = "fs-format-picture__fill-dd";
  colorBtn.title = "填充颜色";
  const colorSwatch = document.createElement("span");
  colorSwatch.className = "fs-format-picture__fill-swatch";
  colorSwatch.setAttribute("aria-hidden", "true");
  const colorBtnLabel = document.createElement("span");
  colorBtnLabel.className = "fs-format-picture__fill-dd-label";
  const colorBtnArrow = document.createElement("span");
  colorBtnArrow.className = "fs-format-picture__fill-dd-arrow";
  colorBtnArrow.setAttribute("aria-hidden", "true");
  colorBtn.appendChild(colorSwatch);
  colorBtn.appendChild(colorBtnLabel);
  colorBtn.appendChild(colorBtnArrow);

  function mountSolidFillColorPalette(anchor: HTMLButtonElement): void {
    closeFillColorPopover();
    const pop = document.createElement("div");
    pop.className = "fs-format-picture__color-popover fs-color-menu";
    appendRibbonColorPaletteContent(pop, {
      themeHeading: "主题颜色",
      standardHeading: "标准色",
      includeNoneRow: true,
      onNone: () => {
        options.setFrameFill({ kind: "none" });
        closeFillColorPopover();
      },
      onPickHex: (hex: string) => {
        options.setFrameFill({ kind: "solid", solidColor: hex });
        closeFillColorPopover();
      },
      onMoreColors: () => {
        closeFillColorPopover();
        void (async () => {
          const ff = options.getFrameFill();
          const cur =
            ff !== null &&
            ff.kind === "solid" &&
            typeof ff.solidColor === "string" &&
            /^#[0-9a-fA-F]{6}$/.test(ff.solidColor)
              ? ff.solidColor.toLowerCase()
              : "#000000";
          const picked = await showRibbonColorDialog(cur);
          if (picked !== null) {
            options.setFrameFill({ kind: "solid", solidColor: picked });
          }
        })();
      },
    });
    document.body.appendChild(pop);
    const position = (): void => {
      const r = anchor.getBoundingClientRect();
      const pw = pop.offsetWidth;
      const left = Math.max(8, Math.min(r.left, window.innerWidth - pw - 8));
      pop.style.position = "fixed";
      pop.style.left = `${left}px`;
      pop.style.top = `${r.bottom + 4}px`;
    };
    requestAnimationFrame(position);
    const onDoc = (ev: PointerEvent): void => {
      const t = ev.target as Node | null;
      if (t !== null && (pop.contains(t) || anchor.contains(t))) {
        return;
      }
      closeFillColorPopover();
    };
    fillColorPopoverCleanup = (): void => {
      pop.remove();
      document.removeEventListener("pointerdown", onDoc, true);
    };
    setTimeout(() => document.addEventListener("pointerdown", onDoc, true), 0);
  }

  colorBtn.addEventListener("click", () => {
    if (syncing) {
      return;
    }
    mountSolidFillColorPalette(colorBtn);
  });
  colorRow.appendChild(colorLab);
  colorRow.appendChild(colorBtn);

  const fillTransRow = document.createElement("div");
  fillTransRow.className = "fs-format-picture__row";
  const fillTransLab = document.createElement("label");
  fillTransLab.textContent = "透明度";
  const fillTransRange = document.createElement("input");
  fillTransRange.type = "range";
  fillTransRange.min = "0";
  fillTransRange.max = "100";
  fillTransRange.step = "1";
  const fillTransNum = document.createElement("input");
  fillTransNum.type = "text";
  fillTransNum.className = "fs-format-picture__num";
  const applyFillTransFromRange = (): void => {
    if (syncing) {
      return;
    }
    const v = Number(fillTransRange.value);
    options.setFrameFill({ solidTransparencyPct: v });
    fillTransNum.value = `${v}%`;
  };
  const applyFillTransFromNum = (): void => {
    if (syncing) {
      return;
    }
    const n = Number(fillTransNum.value.replace(/%/g, "").trim());
    if (!Number.isFinite(n)) {
      return;
    }
    const cl = Math.min(100, Math.max(0, n));
    options.setFrameFill({ solidTransparencyPct: cl });
    fillTransRange.value = String(cl);
    fillTransNum.value = `${cl}%`;
  };
  fillTransRange.addEventListener("input", applyFillTransFromRange);
  fillTransNum.addEventListener("change", applyFillTransFromNum);
  fillTransNum.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      applyFillTransFromNum();
    }
  });
  fillTransRow.appendChild(fillTransLab);
  fillTransRow.appendChild(fillTransRange);
  fillTransRow.appendChild(fillTransNum);

  solidExtra.appendChild(colorRow);
  solidExtra.appendChild(fillTransRow);
  fillBody.appendChild(solidExtra);
  fillDetails.appendChild(fillSum);
  fillDetails.appendChild(fillBody);
  panelFillLine.appendChild(fillDetails);

  function applyFillUiFromModel(ff: FloatingPictureFrameFill): void {
    for (const [k, row] of fillKindRows) {
      const disk = row.querySelector(".fs-format-picture__radio-disk");
      if (disk !== null) {
        disk.classList.toggle("fs-format-picture__radio-disk--on", ff.kind === k);
      }
      row.setAttribute("aria-checked", ff.kind === k ? "true" : "false");
    }
    solidExtra.hidden = ff.kind !== "solid";
    const hex =
      typeof ff.solidColor === "string" && /^#[0-9a-fA-F]{6}$/.test(ff.solidColor)
        ? ff.solidColor.toLowerCase()
        : "#000000";
    colorSwatch.style.backgroundColor = hex;
    colorBtnLabel.textContent = "自定义颜色";
    fillTransRange.value = String(ff.solidTransparencyPct);
    fillTransNum.value = `${ff.solidTransparencyPct}%`;
  }

  const lineDetails = document.createElement("details");
  lineDetails.className =
    "fs-format-picture__sec fs-format-picture__sec--chrome fs-format-picture__sec--lines-static";
  lineDetails.open = true;
  const lineSum = document.createElement("summary");
  lineSum.textContent = "线条";
  const lineBody = document.createElement("div");
  lineBody.className = "fs-format-picture__sec-body";
  for (const [text, sel] of [
    ["无线条", true],
    ["实线", false],
    ["渐变线", false],
  ] as const) {
    lineBody.appendChild(mkStaticRadioRow(text, sel));
  }
  lineDetails.appendChild(lineSum);
  lineDetails.appendChild(lineBody);
  panelFillLine.appendChild(lineDetails);

  const panelOther = document.createElement("div");
  panelOther.className = "fs-format-picture__panel-other";
  panelOther.hidden = true;
  const ph = document.createElement("div");
  ph.className = "fs-format-picture__placeholder";
  ph.textContent = "此分类即将支持";
  panelOther.appendChild(ph);

  function showPanel(idx: number): void {
    closeFillColorPopover();
    const isPic = idx === 3;
    const isFill = idx === 0;
    panelPicture.hidden = !isPic;
    panelFillLine.hidden = !isFill;
    panelOther.hidden = isPic || isFill;
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

  scroll.appendChild(panelFillLine);
  scroll.appendChild(panelPicture);
  scroll.appendChild(panelOther);

  root.appendChild(head);
  root.appendChild(tabs);
  root.appendChild(scroll);
  options.parent.appendChild(root);

  function syncFromModel(): void {
    const a = options.getAdjustments();
    const layout = options.getLayout();
    const ff = options.getFrameFill();
    syncing = true;
    try {
      if (a === null) {
        return;
      }
      if (ff !== null) {
        applyFillUiFromModel(ff);
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
    closeFillColorPopover();
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

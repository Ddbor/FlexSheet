/**
 * Canvas 右侧「设置图片格式」面板：图片校正 / 颜色 / 透明度 / 裁剪占位数值。
 */

import { appendRibbonColorPaletteContent, showRibbonColorDialog } from "@flexsheet/toolbar";
import type {
  FloatingPictureAdjustments,
  FloatingPictureFillKind,
  FloatingPictureFrameFill,
  FloatingPictureGradientStop,
  FloatingPictureGradientType,
} from "./floating-picture-layer.js";
import {
  DEFAULT_FLOATING_PICTURE_ADJUSTMENTS,
  frameFillToFillLayerStyles,
  gradientStopToRgbaCss,
  gradientStopsToHorizontalBarBackground,
} from "./floating-picture-layer.js";
import {
  FORMAT_PICTURE_GRADIENT_PRESETS,
  LINEAR_DIRECTION_USER_ANGLES,
  presetThumbnailBackground,
  presetToFrameFillPatch,
} from "./format-picture-gradient-presets.js";

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
.fs-format-picture__grad-extra {
  border-top: 1px solid var(--fs-ribbon-border, #e1dfdd);
  margin-top: 6px;
  padding-top: 10px;
}
.fs-format-picture__grad-pop {
  position: fixed;
  z-index: 10006;
  background: var(--fs-sheet-surface, #fff);
  border: 1px solid var(--fs-ribbon-border, #c8c6c4);
  border-radius: 4px;
  box-shadow: 0 6px 24px rgba(0, 0, 0, 0.18);
  padding: 8px;
  box-sizing: border-box;
  max-width: min(280px, calc(100vw - 24px));
}
.fs-format-picture__grad-preset-grid {
  display: grid;
  grid-template-columns: repeat(6, 1fr);
  gap: 4px;
}
.fs-format-picture__grad-preset-cell {
  width: 32px;
  height: 24px;
  padding: 0;
  border: 1px solid #a19f9d;
  border-radius: 2px;
  cursor: pointer;
  box-sizing: border-box;
  background-size: cover;
}
.fs-format-picture__grad-preset-cell:hover {
  outline: 1px solid #217346;
}
.fs-format-picture__grad-dir-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 4px;
}
.fs-format-picture__grad-dir-cell {
  width: 36px;
  height: 28px;
  padding: 0;
  border: 1px solid #a19f9d;
  border-radius: 2px;
  cursor: pointer;
  box-sizing: border-box;
}
.fs-format-picture__grad-dir-cell:hover {
  outline: 1px solid #217346;
}
.fs-format-picture__grad-dir-cell--on {
  outline: 2px solid #217346;
}
.fs-format-picture__grad-select {
  flex: 1;
  min-width: 0;
  padding: 4px 6px;
  font-size: 11px;
  border: 1px solid var(--fs-ribbon-border, #c8c6c4);
  border-radius: 2px;
  font-family: inherit;
  background: var(--fs-sheet-surface, #fff);
  color: inherit;
}
.fs-format-picture__grad-stops-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
  margin-top: 4px;
  margin-bottom: 4px;
}
.fs-format-picture__grad-stops-head > span {
  font-size: 12px;
  font-weight: 600;
}
.fs-format-picture__grad-stops-tools {
  display: flex;
  gap: 4px;
}
.fs-format-picture__grad-stops-tools button {
  width: 22px;
  height: 22px;
  padding: 0;
  border: 1px solid var(--fs-ribbon-border, #c8c6c4);
  border-radius: 2px;
  background: var(--fs-sheet-surface, #fff);
  cursor: pointer;
  font-size: 14px;
  line-height: 1;
}
.fs-format-picture__grad-stops-tools button:hover {
  background: var(--fs-ribbon-hover, #edebe9);
}
.fs-format-picture__grad-track-wrap {
  position: relative;
  height: 28px;
  margin-bottom: 6px;
}
.fs-format-picture__grad-track {
  position: absolute;
  left: 0;
  right: 0;
  top: 4px;
  height: 14px;
  border: 1px solid #a19f9d;
  border-radius: 2px;
  box-sizing: border-box;
  cursor: crosshair;
}
.fs-format-picture__grad-markers {
  position: absolute;
  left: 0;
  right: 0;
  top: 0;
  height: 28px;
  pointer-events: none;
}
.fs-format-picture__grad-marker {
  position: absolute;
  top: 18px;
  width: 0;
  height: 0;
  border-left: 6px solid transparent;
  border-right: 6px solid transparent;
  border-bottom: 9px solid #888;
  filter: drop-shadow(0 0 1px #333);
  padding: 0;
  margin: 0;
  pointer-events: auto;
  cursor: grab;
  transform: translateX(-50%);
  background: none;
  border-top: none;
  border-left-color: transparent;
  border-right-color: transparent;
}
.fs-format-picture__grad-marker--sel {
  outline: 2px solid #c45911;
  outline-offset: 1px;
  border-radius: 1px;
  z-index: 1;
}
.fs-format-picture__grad-rotate-row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-top: 8px;
  font-size: 12px;
  cursor: pointer;
  user-select: none;
}
.fs-format-picture__grad-rotate-row input {
  accent-color: #217346;
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
  let gradPopoverCleanup: (() => void) | null = null;
  let selectedGradientStopIdx = 0;
  /** 拖动光圈时与 `gradientStops` 数组下标一致（已按位置排序） */
  let gradStopDragSortedIdx: number | null = null;

  function closeFillColorPopover(): void {
    if (fillColorPopoverCleanup !== null) {
      fillColorPopoverCleanup();
      fillColorPopoverCleanup = null;
    }
  }

  function closeGradPopovers(): void {
    if (gradPopoverCleanup !== null) {
      gradPopoverCleanup();
      gradPopoverCleanup = null;
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

  const FILL_OPTIONS: readonly {
    readonly kind: FloatingPictureFillKind;
    readonly label: string;
  }[] = [
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

  const gradientExtra = document.createElement("div");
  gradientExtra.className = "fs-format-picture__grad-extra";
  gradientExtra.hidden = true;

  function userAngleToCssDeg(u: number): number {
    const x = ((u % 360) + 360) % 360;
    return (90 + x) % 360;
  }

  const presetRow = document.createElement("div");
  presetRow.className = "fs-format-picture__row fs-format-picture__color-dd-row";
  const presetLab = document.createElement("label");
  presetLab.textContent = "预设渐变";
  const presetBtn = document.createElement("button");
  presetBtn.type = "button";
  presetBtn.className = "fs-format-picture__fill-dd";
  const presetSwatch = document.createElement("span");
  presetSwatch.className = "fs-format-picture__fill-swatch";
  const presetBtnLbl = document.createElement("span");
  presetBtnLbl.className = "fs-format-picture__fill-dd-label";
  presetBtnLbl.textContent = "预设渐变";
  const presetArr = document.createElement("span");
  presetArr.className = "fs-format-picture__fill-dd-arrow";
  presetBtn.appendChild(presetSwatch);
  presetBtn.appendChild(presetBtnLbl);
  presetBtn.appendChild(presetArr);
  presetRow.appendChild(presetLab);
  presetRow.appendChild(presetBtn);

  const typeRow = document.createElement("div");
  typeRow.className = "fs-format-picture__row";
  const typeLab = document.createElement("label");
  typeLab.textContent = "类型";
  const typeSel = document.createElement("select");
  typeSel.className = "fs-format-picture__grad-select";
  for (const [v, lab] of [
    ["linear", "线性"],
    ["radial", "射线"],
  ] as const) {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = lab;
    typeSel.appendChild(o);
  }
  typeRow.appendChild(typeLab);
  typeRow.appendChild(typeSel);

  const dirRow = document.createElement("div");
  dirRow.className = "fs-format-picture__row fs-format-picture__color-dd-row";
  const dirLab = document.createElement("label");
  dirLab.textContent = "方向";
  const dirBtn = document.createElement("button");
  dirBtn.type = "button";
  dirBtn.className = "fs-format-picture__fill-dd";
  const dirSwatch = document.createElement("span");
  dirSwatch.className = "fs-format-picture__fill-swatch";
  const dirBtnLbl = document.createElement("span");
  dirBtnLbl.className = "fs-format-picture__fill-dd-label";
  dirBtnLbl.textContent = "方向";
  const dirArr = document.createElement("span");
  dirArr.className = "fs-format-picture__fill-dd-arrow";
  dirBtn.appendChild(dirSwatch);
  dirBtn.appendChild(dirBtnLbl);
  dirBtn.appendChild(dirArr);
  dirRow.appendChild(dirLab);
  dirRow.appendChild(dirBtn);

  const angRow = document.createElement("div");
  angRow.className = "fs-format-picture__row";
  const angLab = document.createElement("label");
  angLab.textContent = "角度";
  const angNum = document.createElement("input");
  angNum.type = "number";
  angNum.className = "fs-format-picture__num";
  angNum.min = "0";
  angNum.max = "359";
  angNum.step = "1";
  angRow.appendChild(angLab);
  angRow.appendChild(angNum);

  const stopsHead = document.createElement("div");
  stopsHead.className = "fs-format-picture__grad-stops-head";
  const stopsTitle = document.createElement("span");
  stopsTitle.textContent = "渐变光圈";
  const stopsTools = document.createElement("div");
  stopsTools.className = "fs-format-picture__grad-stops-tools";
  const stopAdd = document.createElement("button");
  stopAdd.type = "button";
  stopAdd.textContent = "+";
  stopAdd.title = "添加渐变光圈";
  const stopRem = document.createElement("button");
  stopRem.type = "button";
  stopRem.textContent = "−";
  stopRem.title = "删除渐变光圈";
  stopsTools.appendChild(stopAdd);
  stopsTools.appendChild(stopRem);
  stopsHead.appendChild(stopsTitle);
  stopsHead.appendChild(stopsTools);

  const trackWrap = document.createElement("div");
  trackWrap.className = "fs-format-picture__grad-track-wrap";
  const gradTrack = document.createElement("div");
  gradTrack.className = "fs-format-picture__grad-track";
  const gradMarkers = document.createElement("div");
  gradMarkers.className = "fs-format-picture__grad-markers";
  trackWrap.appendChild(gradTrack);
  trackWrap.appendChild(gradMarkers);

  const gColorRow = document.createElement("div");
  gColorRow.className = "fs-format-picture__row fs-format-picture__color-dd-row";
  const gColorLab = document.createElement("label");
  gColorLab.textContent = "颜色";
  const gColorBtn = document.createElement("button");
  gColorBtn.type = "button";
  gColorBtn.className = "fs-format-picture__fill-dd";
  const gColorSwatch = document.createElement("span");
  gColorSwatch.className = "fs-format-picture__fill-swatch";
  const gColorLbl = document.createElement("span");
  gColorLbl.className = "fs-format-picture__fill-dd-label";
  const gColorArr = document.createElement("span");
  gColorArr.className = "fs-format-picture__fill-dd-arrow";
  gColorBtn.appendChild(gColorSwatch);
  gColorBtn.appendChild(gColorLbl);
  gColorBtn.appendChild(gColorArr);
  gColorRow.appendChild(gColorLab);
  gColorRow.appendChild(gColorBtn);

  const gPosRow = document.createElement("div");
  gPosRow.className = "fs-format-picture__row";
  const gPosLab = document.createElement("label");
  gPosLab.textContent = "位置";
  const gPosNum = document.createElement("input");
  gPosNum.type = "number";
  gPosNum.className = "fs-format-picture__num";
  gPosNum.min = "0";
  gPosNum.max = "100";
  gPosRow.appendChild(gPosLab);
  gPosRow.appendChild(gPosNum);

  const gTrRow = document.createElement("div");
  gTrRow.className = "fs-format-picture__row";
  const gTrLab = document.createElement("label");
  gTrLab.textContent = "透明度";
  const gTrRng = document.createElement("input");
  gTrRng.type = "range";
  gTrRng.min = "0";
  gTrRng.max = "100";
  gTrRng.step = "1";
  const gTrNum = document.createElement("input");
  gTrNum.type = "text";
  gTrNum.className = "fs-format-picture__num";
  gTrRow.appendChild(gTrLab);
  gTrRow.appendChild(gTrRng);
  gTrRow.appendChild(gTrNum);

  const gBrRow = document.createElement("div");
  gBrRow.className = "fs-format-picture__row";
  const gBrLab = document.createElement("label");
  gBrLab.textContent = "亮度";
  const gBrRng = document.createElement("input");
  gBrRng.type = "range";
  gBrRng.min = "-100";
  gBrRng.max = "100";
  gBrRng.step = "1";
  const gBrNum = document.createElement("input");
  gBrNum.type = "text";
  gBrNum.className = "fs-format-picture__num";
  gBrRow.appendChild(gBrLab);
  gBrRow.appendChild(gBrRng);
  gBrRow.appendChild(gBrNum);

  const rotRow = document.createElement("label");
  rotRow.className = "fs-format-picture__grad-rotate-row";
  const rotChk = document.createElement("input");
  rotChk.type = "checkbox";
  rotChk.checked = true;
  rotRow.appendChild(rotChk);
  rotRow.appendChild(document.createTextNode("与形状一起旋转"));

  gradientExtra.appendChild(presetRow);
  gradientExtra.appendChild(typeRow);
  gradientExtra.appendChild(dirRow);
  gradientExtra.appendChild(angRow);
  gradientExtra.appendChild(stopsHead);
  gradientExtra.appendChild(trackWrap);
  gradientExtra.appendChild(gColorRow);
  gradientExtra.appendChild(gPosRow);
  gradientExtra.appendChild(gTrRow);
  gradientExtra.appendChild(gBrRow);
  gradientExtra.appendChild(rotRow);

  function sortedStops(ff: FloatingPictureFrameFill): FloatingPictureGradientStop[] {
    if (ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return [];
    }
    return [...ff.gradientStops].sort((a, b) => a.positionPct - b.positionPct);
  }

  function paintGradientTrack(ff: FloatingPictureFrameFill): void {
    if (ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    gradTrack.style.backgroundImage = gradientStopsToHorizontalBarBackground(ff.gradientStops);
    gradTrack.style.backgroundColor = "";
  }

  function highlightGradientMarkerSelectionOnly(): void {
    const nodes = gradMarkers.querySelectorAll(".fs-format-picture__grad-marker");
    nodes.forEach((el, i) => {
      el.classList.toggle("fs-format-picture__grad-marker--sel", i === selectedGradientStopIdx);
    });
  }

  function updateGradientMarkerPositionsOnly(ff: FloatingPictureFrameFill): void {
    if (ff.kind !== "gradient") {
      return;
    }
    const stops = sortedStops(ff);
    const nodes = gradMarkers.querySelectorAll(".fs-format-picture__grad-marker");
    nodes.forEach((el, i) => {
      const s = stops[i];
      if (s !== undefined && el instanceof HTMLElement) {
        el.style.left = `${s.positionPct}%`;
        el.style.borderBottomColor = gradientStopToRgbaCss(s);
      }
    });
  }

  function paintDirSwatch(userDeg: number): void {
    const css = userAngleToCssDeg(userDeg);
    dirSwatch.style.backgroundImage = `linear-gradient(${css}deg, #5b9bd5, #ffffff)`;
    dirSwatch.style.backgroundColor = "";
  }

  function paintPresetSwatch(ff: FloatingPictureFrameFill): void {
    if (ff.kind !== "gradient") {
      return;
    }
    const st = frameFillToFillLayerStyles(ff);
    presetSwatch.style.backgroundImage = st.backgroundImage;
    presetSwatch.style.backgroundColor =
      st.backgroundColor === "transparent" ? "" : st.backgroundColor;
  }

  function refreshGradientMarkers(ff: FloatingPictureFrameFill): void {
    gradMarkers.replaceChildren();
    if (ff.kind !== "gradient") {
      return;
    }
    const stops = sortedStops(ff);
    if (stops.length === 0) {
      return;
    }
    selectedGradientStopIdx = Math.min(Math.max(0, selectedGradientStopIdx), stops.length - 1);
    for (let si = 0; si < stops.length; si++) {
      const s = stops[si]!;
      const b = document.createElement("button");
      b.type = "button";
      b.className =
        "fs-format-picture__grad-marker" +
        (si === selectedGradientStopIdx ? " fs-format-picture__grad-marker--sel" : "");
      b.style.left = `${s.positionPct}%`;
      b.style.borderBottomColor = gradientStopToRgbaCss(s);
      b.style.borderLeftColor = "transparent";
      b.style.borderRightColor = "transparent";
      b.title = `光圈 ${si + 1}`;
      b.addEventListener("pointerdown", (ev) => {
        if (syncing || ev.button !== 0) {
          return;
        }
        ev.preventDefault();
        ev.stopPropagation();
        selectedGradientStopIdx = si;
        const cur = options.getFrameFill();
        if (cur !== null && cur.kind === "gradient") {
          syncGradientDetailEditors(cur);
          highlightGradientMarkerSelectionOnly();
        }
        gradStopDragSortedIdx = si;
        const onMove = (e: PointerEvent): void => {
          if (gradStopDragSortedIdx === null || syncing) {
            return;
          }
          const r = gradTrack.getBoundingClientRect();
          if (r.width <= 0) {
            return;
          }
          let pct = Math.min(100, Math.max(0, ((e.clientX - r.left) / r.width) * 100));
          const g0 = options.getFrameFill();
          if (g0 === null || g0.kind !== "gradient" || g0.gradientStops === undefined) {
            return;
          }
          const idx = gradStopDragSortedIdx;
          if (idx < 0 || idx >= g0.gradientStops.length) {
            return;
          }
          const eps = 0.08;
          const below = idx > 0 ? g0.gradientStops[idx - 1] : undefined;
          const above = idx < g0.gradientStops.length - 1 ? g0.gradientStops[idx + 1] : undefined;
          const minP = below !== undefined ? below.positionPct + eps : 0;
          const maxP = above !== undefined ? above.positionPct - eps : 100;
          pct = Math.min(maxP, Math.max(minP, pct));
          const next = g0.gradientStops.map((st, i) =>
            i === idx ? { ...st, positionPct: pct } : { ...st },
          );
          options.setFrameFill({ gradientStops: next, gradientPresetId: null });
        };
        const onUp = (): void => {
          document.removeEventListener("pointermove", onMove, true);
          document.removeEventListener("pointerup", onUp, true);
          document.removeEventListener("pointercancel", onUp, true);
          gradStopDragSortedIdx = null;
          const fin = options.getFrameFill();
          if (fin !== null && fin.kind === "gradient") {
            refreshGradientMarkers(fin);
            syncGradientDetailEditors(fin);
          }
        };
        document.addEventListener("pointermove", onMove, true);
        document.addEventListener("pointerup", onUp, true);
        document.addEventListener("pointercancel", onUp, true);
      });
      gradMarkers.appendChild(b);
    }
  }

  function syncGradientDetailEditors(ff: FloatingPictureFrameFill): void {
    if (ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const stops = sortedStops(ff);
    if (stops.length === 0) {
      return;
    }
    selectedGradientStopIdx = Math.min(Math.max(0, selectedGradientStopIdx), stops.length - 1);
    const sel = stops[selectedGradientStopIdx];
    if (sel === undefined) {
      return;
    }
    /* 与预览条、把手一致：含亮度、透明度（非仅十六进制基色） */
    gColorSwatch.style.backgroundColor = gradientStopToRgbaCss(sel);
    gColorSwatch.style.backgroundImage = "";
    gColorLbl.textContent = "颜色";
    gPosNum.value = String(Math.round(sel.positionPct));
    gTrRng.value = String(sel.transparencyPct);
    gTrNum.value = `${sel.transparencyPct}%`;
    gBrRng.value = String(sel.brightnessPct);
    gBrNum.value = `${sel.brightnessPct}%`;
  }

  function syncGradientFromModel(ff: FloatingPictureFrameFill): void {
    if (ff.kind !== "gradient") {
      return;
    }
    const rawGt = ff.gradientType ?? "linear";
    const gt = rawGt === "linear" || rawGt === "radial" ? rawGt : "radial";
    typeSel.value = gt;
    const ang = ff.gradientAngleDeg ?? 90;
    angNum.value = String(Math.round(ang));
    paintDirSwatch(ang);
    paintPresetSwatch(ff);
    dirRow.hidden = gt !== "linear";
    angRow.hidden = gt !== "linear";
    rotChk.checked = ff.gradientRotateWithShape !== false;
    paintGradientTrack(ff);
    refreshGradientMarkers(ff);
    syncGradientDetailEditors(ff);
    const idx = ff.linearDirectionIndex;
    if (typeof idx === "number" && idx >= 0 && idx < 8) {
      dirBtnLbl.textContent = `方向 ${idx + 1}`;
    } else {
      dirBtnLbl.textContent = "方向";
    }
  }

  function mountGradientPresetPopover(anchor: HTMLElement): void {
    closeGradPopovers();
    const pop = document.createElement("div");
    pop.className = "fs-format-picture__grad-pop";
    const grid = document.createElement("div");
    grid.className = "fs-format-picture__grad-preset-grid";
    for (const p of FORMAT_PICTURE_GRADIENT_PRESETS) {
      const cell = document.createElement("button");
      cell.type = "button";
      cell.className = "fs-format-picture__grad-preset-cell";
      cell.style.backgroundImage = presetThumbnailBackground(p);
      cell.title = `预设 ${p.id + 1}`;
      cell.addEventListener("click", () => {
        if (!syncing) {
          options.setFrameFill(presetToFrameFillPatch(p));
        }
        closeGradPopovers();
      });
      grid.appendChild(cell);
    }
    pop.appendChild(grid);
    document.body.appendChild(pop);
    const position = (): void => {
      const r = anchor.getBoundingClientRect();
      const pw = pop.offsetWidth;
      const left = Math.max(8, Math.min(r.left, window.innerWidth - pw - 8));
      pop.style.left = `${left}px`;
      pop.style.top = `${r.bottom + 4}px`;
    };
    requestAnimationFrame(position);
    const onDoc = (ev: PointerEvent): void => {
      const t = ev.target as Node | null;
      if (t !== null && (pop.contains(t) || anchor.contains(t))) {
        return;
      }
      closeGradPopovers();
    };
    gradPopoverCleanup = (): void => {
      pop.remove();
      document.removeEventListener("pointerdown", onDoc, true);
    };
    setTimeout(() => document.addEventListener("pointerdown", onDoc, true), 0);
  }

  function mountDirectionPopover(anchor: HTMLElement): void {
    closeGradPopovers();
    const pop = document.createElement("div");
    pop.className = "fs-format-picture__grad-pop";
    const grid = document.createElement("div");
    grid.className = "fs-format-picture__grad-dir-grid";
    const ff0 = options.getFrameFill();
    const curIdx =
      ff0 !== null && ff0.kind === "gradient" && typeof ff0.linearDirectionIndex === "number"
        ? ff0.linearDirectionIndex
        : -1;
    for (let i = 0; i < LINEAR_DIRECTION_USER_ANGLES.length; i++) {
      const u = LINEAR_DIRECTION_USER_ANGLES[i]!;
      const cell = document.createElement("button");
      cell.type = "button";
      cell.className =
        "fs-format-picture__grad-dir-cell" +
        (i === curIdx ? " fs-format-picture__grad-dir-cell--on" : "");
      const css = userAngleToCssDeg(u);
      cell.style.backgroundImage = `linear-gradient(${css}deg, #5b9bd5, #ffffff)`;
      cell.title = `${u}°`;
      cell.addEventListener("click", () => {
        if (!syncing) {
          options.setFrameFill({
            kind: "gradient",
            linearDirectionIndex: i,
            gradientAngleDeg: u,
            gradientPresetId: null,
          });
        }
        closeGradPopovers();
      });
      grid.appendChild(cell);
    }
    pop.appendChild(grid);
    document.body.appendChild(pop);
    const position = (): void => {
      const r = anchor.getBoundingClientRect();
      const pw = pop.offsetWidth;
      const left = Math.max(8, Math.min(r.left, window.innerWidth - pw - 8));
      pop.style.left = `${left}px`;
      pop.style.top = `${r.bottom + 4}px`;
    };
    requestAnimationFrame(position);
    const onDoc = (ev: PointerEvent): void => {
      const t = ev.target as Node | null;
      if (t !== null && (pop.contains(t) || anchor.contains(t))) {
        return;
      }
      closeGradPopovers();
    };
    gradPopoverCleanup = (): void => {
      pop.remove();
      document.removeEventListener("pointerdown", onDoc, true);
    };
    setTimeout(() => document.addEventListener("pointerdown", onDoc, true), 0);
  }

  presetBtn.addEventListener("click", () => {
    if (!syncing) {
      mountGradientPresetPopover(presetBtn);
    }
  });
  dirBtn.addEventListener("click", () => {
    if (syncing) {
      return;
    }
    const ff = options.getFrameFill();
    if (ff !== null && ff.kind === "gradient" && (ff.gradientType ?? "linear") === "linear") {
      mountDirectionPopover(dirBtn);
    }
  });
  typeSel.addEventListener("change", () => {
    if (syncing) {
      return;
    }
    const v = typeSel.value as FloatingPictureGradientType;
    if (v === "radial") {
      options.setFrameFill({
        kind: "gradient",
        gradientType: "radial",
        gradientPresetId: null,
        radialFillLtrb: { l: 0, t: 0, r: 100000, b: 100000 },
        radialTileLtrb: { l: -100000, t: -100000, r: 0, b: 0 },
      });
    } else {
      options.setFrameFill({ kind: "gradient", gradientType: "linear", gradientPresetId: null });
    }
    const ff = options.getFrameFill();
    if (ff !== null) {
      syncGradientFromModel(ff);
    }
  });
  angNum.addEventListener("change", () => {
    if (syncing) {
      return;
    }
    const n = Number(angNum.value);
    if (!Number.isFinite(n)) {
      return;
    }
    const a = Math.min(359, Math.max(0, Math.round(n)));
    options.setFrameFill({ kind: "gradient", gradientAngleDeg: a, gradientPresetId: null });
  });

  gradTrack.addEventListener("pointerdown", (ev) => {
    if (syncing || ev.button !== 0) {
      return;
    }
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const r = gradTrack.getBoundingClientRect();
    if (r.width <= 0) {
      return;
    }
    const pct = Math.min(100, Math.max(0, ((ev.clientX - r.left) / r.width) * 100));
    const stops = sortedStops(ff);
    let lo: FloatingPictureGradientStop | null = null;
    let hi: FloatingPictureGradientStop | null = null;
    for (let i = 0; i < stops.length; i++) {
      const s = stops[i]!;
      if (s.positionPct <= pct) {
        lo = s;
      }
      if (s.positionPct >= pct && hi === null) {
        hi = s;
      }
    }
    const baseColor =
      lo !== null && hi !== null ? lo.color : (ff.gradientStops[0]?.color ?? "#5b9bd5");
    const newStop: FloatingPictureGradientStop = {
      positionPct: pct,
      color: baseColor,
      transparencyPct: 0,
      brightnessPct: 100,
    };
    const merged = [...ff.gradientStops, newStop].sort((a, b) => a.positionPct - b.positionPct);
    const newIdx = merged.indexOf(newStop);
    selectedGradientStopIdx = newIdx >= 0 ? newIdx : merged.length - 1;
    options.setFrameFill({ gradientStops: merged, gradientPresetId: null });
  });

  stopAdd.addEventListener("click", () => {
    if (syncing) {
      return;
    }
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const stops = [...ff.gradientStops].sort((a, b) => a.positionPct - b.positionPct);
    if (stops.length >= 12) {
      return;
    }
    const mid =
      stops.length >= 2 ? (stops[0]!.positionPct + stops[stops.length - 1]!.positionPct) / 2 : 50;
    const last = stops[stops.length - 1] ?? {
      positionPct: 100,
      color: "#ffffff",
      transparencyPct: 0,
      brightnessPct: 100,
    };
    const merged = [
      ...stops,
      {
        positionPct: Math.min(99, Math.max(1, mid)),
        color: last.color,
        transparencyPct: 0,
        brightnessPct: 100,
      },
    ].sort((a, b) => a.positionPct - b.positionPct);
    selectedGradientStopIdx = merged.length - 1;
    options.setFrameFill({ gradientStops: merged, gradientPresetId: null });
  });
  stopRem.addEventListener("click", () => {
    if (syncing) {
      return;
    }
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    if (ff.gradientStops.length <= 2) {
      return;
    }
    const ordered = [...ff.gradientStops].sort((a, b) => a.positionPct - b.positionPct);
    const si = Math.min(Math.max(0, selectedGradientStopIdx), ordered.length - 1);
    const victim = ordered[si] ?? ordered[ordered.length - 1];
    const merged = ff.gradientStops.filter((s) => s !== victim);
    selectedGradientStopIdx = Math.min(selectedGradientStopIdx, merged.length - 1);
    options.setFrameFill({ gradientStops: merged, gradientPresetId: null });
  });

  function mountGradientStopColorPalette(anchor: HTMLButtonElement): void {
    closeFillColorPopover();
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const stopsSnap = ff.gradientStops;
    const pickIx = Math.min(selectedGradientStopIdx, stopsSnap.length - 1);
    const pop = document.createElement("div");
    pop.className = "fs-format-picture__color-popover fs-color-menu";
    appendRibbonColorPaletteContent(pop, {
      themeHeading: "主题颜色",
      standardHeading: "标准色",
      includeNoneRow: false,
      onPickHex: (hex: string) => {
        const g = options.getFrameFill();
        if (g === null || g.kind !== "gradient" || g.gradientStops === undefined) {
          closeFillColorPopover();
          return;
        }
        const ix = Math.min(selectedGradientStopIdx, g.gradientStops.length - 1);
        const next = g.gradientStops.map((st, i) =>
          i === ix ? { ...st, color: hex.toLowerCase() } : { ...st },
        );
        options.setFrameFill({ gradientStops: next, gradientPresetId: null });
        closeFillColorPopover();
      },
      onMoreColors: () => {
        closeFillColorPopover();
        void (async () => {
          const cur0 = stopsSnap[pickIx]?.color ?? "#000000";
          const cur = /^#[0-9a-fA-F]{6}$/.test(cur0) ? cur0.toLowerCase() : "#000000";
          const picked = await showRibbonColorDialog(cur);
          if (picked !== null) {
            const g2 = options.getFrameFill();
            if (g2 === null || g2.kind !== "gradient" || g2.gradientStops === undefined) {
              return;
            }
            const ix = Math.min(selectedGradientStopIdx, g2.gradientStops.length - 1);
            const next = g2.gradientStops.map((st, i) =>
              i === ix ? { ...st, color: picked } : { ...st },
            );
            options.setFrameFill({ gradientStops: next, gradientPresetId: null });
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

  gColorBtn.addEventListener("click", () => {
    if (!syncing) {
      mountGradientStopColorPalette(gColorBtn);
    }
  });
  const applyGPos = (): void => {
    if (syncing) {
      return;
    }
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const n = Number(gPosNum.value);
    if (!Number.isFinite(n)) {
      return;
    }
    const pct = Math.min(100, Math.max(0, n));
    const ix = Math.min(selectedGradientStopIdx, ff.gradientStops.length - 1);
    const next = ff.gradientStops.map((st, i) =>
      i === ix ? { ...st, positionPct: pct } : { ...st },
    );
    options.setFrameFill({ gradientStops: next, gradientPresetId: null });
  };
  gPosNum.addEventListener("change", applyGPos);
  gPosNum.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      applyGPos();
    }
  });
  const applyGTr = (): void => {
    if (syncing) {
      return;
    }
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const v = Number(gTrRng.value);
    const ix = Math.min(selectedGradientStopIdx, ff.gradientStops.length - 1);
    const next = ff.gradientStops.map((st, i) =>
      i === ix ? { ...st, transparencyPct: v } : { ...st },
    );
    options.setFrameFill({ gradientStops: next, gradientPresetId: null });
    gTrNum.value = `${v}%`;
  };
  gTrRng.addEventListener("input", applyGTr);
  gTrNum.addEventListener("change", () => {
    if (syncing) {
      return;
    }
    const n = Number(gTrNum.value.replace(/%/g, "").trim());
    if (!Number.isFinite(n)) {
      return;
    }
    const v = Math.min(100, Math.max(0, n));
    gTrRng.value = String(v);
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const ix = Math.min(selectedGradientStopIdx, ff.gradientStops.length - 1);
    const next = ff.gradientStops.map((st, i) =>
      i === ix ? { ...st, transparencyPct: v } : { ...st },
    );
    options.setFrameFill({ gradientStops: next, gradientPresetId: null });
    gTrNum.value = `${v}%`;
  });
  const applyGBr = (): void => {
    if (syncing) {
      return;
    }
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const v = Number(gBrRng.value);
    const ix = Math.min(selectedGradientStopIdx, ff.gradientStops.length - 1);
    const next = ff.gradientStops.map((st, i) =>
      i === ix ? { ...st, brightnessPct: v } : { ...st },
    );
    options.setFrameFill({ gradientStops: next, gradientPresetId: null });
    gBrNum.value = `${v}%`;
  };
  gBrRng.addEventListener("input", applyGBr);
  gBrNum.addEventListener("change", () => {
    if (syncing) {
      return;
    }
    const n = Number(gBrNum.value.replace(/%/g, "").trim());
    if (!Number.isFinite(n)) {
      return;
    }
    const v = Math.min(100, Math.max(-100, n));
    gBrRng.value = String(v);
    const ff = options.getFrameFill();
    if (ff === null || ff.kind !== "gradient" || ff.gradientStops === undefined) {
      return;
    }
    const ix = Math.min(selectedGradientStopIdx, ff.gradientStops.length - 1);
    const next = ff.gradientStops.map((st, i) =>
      i === ix ? { ...st, brightnessPct: v } : { ...st },
    );
    options.setFrameFill({ gradientStops: next, gradientPresetId: null });
    gBrNum.value = `${v}%`;
  });
  rotChk.addEventListener("change", () => {
    if (syncing) {
      return;
    }
    options.setFrameFill({ kind: "gradient", gradientRotateWithShape: rotChk.checked });
  });

  rotRow.addEventListener("click", (ev) => {
    if (ev.target === rotChk) {
      return;
    }
    if (syncing) {
      return;
    }
    rotChk.checked = !rotChk.checked;
    options.setFrameFill({ kind: "gradient", gradientRotateWithShape: rotChk.checked });
  });

  fillBody.appendChild(solidExtra);
  fillBody.appendChild(gradientExtra);
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
    gradientExtra.hidden = ff.kind !== "gradient";
    const hex =
      typeof ff.solidColor === "string" && /^#[0-9a-fA-F]{6}$/.test(ff.solidColor)
        ? ff.solidColor.toLowerCase()
        : "#000000";
    colorSwatch.style.backgroundColor = hex;
    colorBtnLabel.textContent = "自定义颜色";
    fillTransRange.value = String(ff.solidTransparencyPct);
    fillTransNum.value = `${ff.solidTransparencyPct}%`;
    if (ff.kind === "gradient") {
      if (gradStopDragSortedIdx !== null) {
        paintGradientTrack(ff);
        updateGradientMarkerPositionsOnly(ff);
        syncGradientDetailEditors(ff);
        highlightGradientMarkerSelectionOnly();
      } else {
        syncGradientFromModel(ff);
      }
    }
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
    closeGradPopovers();
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
    closeGradPopovers();
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

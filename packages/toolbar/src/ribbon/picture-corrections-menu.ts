import { iconFormatPainter } from "../toolbar/icons.js";
import {
  clearToolbarDropdownMenuPosition,
  closeAllRibbonPopups,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import type { FlexSheetLike, FloatingPictureAdjustmentsState } from "./ribbon-types.js";

const DEF: FloatingPictureAdjustmentsState = {
  brightnessPct: 0,
  contrastPct: 0,
  sharpnessPct: 0,
  saturationPct: 100,
  colorTemperatureK: 6500,
  transparencyPct: 0,
};

const SHARP_PRESETS = [-50, -25, 0, 25, 50] as const;
/** 自上而下：暗 → 亮 */
const BC_BRIGHT = [-40, -20, 0, 20, 40] as const;
/** 自左而右：低对比 → 高对比 */
const BC_CONT = [-40, -20, 0, 20, 40] as const;

function clampN(n: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, n));
}

/** 与 flexsheet `buildFloatingPictureCssFilter` 保持一致，供缩略预览。 */
function buildPreviewFilter(a: FloatingPictureAdjustmentsState): string {
  const b = clampN(1 + a.brightnessPct / 100, 0.05, 3);
  const c0 = clampN(1 + a.contrastPct / 100, 0.05, 3);
  const sat = clampN(a.saturationPct / 100, 0, 4);
  const op = clampN(1 - a.transparencyPct / 100, 0, 1);
  let blurPx = 0;
  let sharpenBoost = 1;
  if (a.sharpnessPct < 0) {
    blurPx = clampN((-a.sharpnessPct / 100) * 3, 0, 3.5);
  } else if (a.sharpnessPct > 0) {
    sharpenBoost = 1 + (a.sharpnessPct / 100) * 0.55;
  }
  const c = c0 * sharpenBoost;
  const parts: string[] = [];
  if (blurPx > 0.05) {
    parts.push(`blur(${blurPx}px)`);
  }
  parts.push(`brightness(${b})`);
  parts.push(`contrast(${c})`);
  parts.push(`saturate(${sat})`);
  const tk = a.colorTemperatureK;
  if (Number.isFinite(tk) && tk !== 6500) {
    const deg = clampN(((tk - 6500) / 4500) * 28, -32, 32);
    if (Math.abs(deg) > 0.3) {
      parts.push(`hue-rotate(${deg}deg)`);
    }
  }
  parts.push(`opacity(${op})`);
  return parts.join(" ");
}

function mergeAdj(partial: Partial<FloatingPictureAdjustmentsState>): FloatingPictureAdjustmentsState {
  return { ...DEF, ...partial };
}

/**
 * 「图片格式 → 更正」：锐化/柔化条 + 亮度/对比度 5×5 + 打开右侧格式窗格。
 */
export function mountPictureCorrectionsMenu(
  anchor: HTMLButtonElement,
  getFlexSheet: () => FlexSheetLike | undefined,
): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-picture-format-correct";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-pic-corr-menu";
  menu.hidden = true;
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);
  menu.setAttribute("role", "menu");

  const previewUrl = (): string => {
    const fs = getFlexSheet();
    const u = fs?.getSelectedFloatingPictureDataUrl?.() ?? null;
    return u !== null && u !== "" ? u : PLACEHOLDER_IMG;
  };

  const sharpenRow = document.createElement("div");
  sharpenRow.className = "fs-pic-corr-menu__sh-row";
  const sharpenHead = document.createElement("div");
  sharpenHead.className = "fs-pic-corr-menu__section-head";
  sharpenHead.textContent = "锐化/柔化";
  menu.appendChild(sharpenHead);
  const sharpenBtns: HTMLButtonElement[] = [];
  for (let i = 0; i < SHARP_PRESETS.length; i++) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "fs-pic-corr-menu__thumb";
    b.setAttribute("role", "menuitem");
    const img = document.createElement("img");
    img.className = "fs-pic-corr-menu__thumb-img";
    img.alt = "";
    img.draggable = false;
    b.appendChild(img);
    const idx = i;
    b.addEventListener("click", (ev) => {
      ev.stopPropagation();
      applyPatch({ sharpnessPct: SHARP_PRESETS[idx] });
    });
    sharpenBtns.push(b);
    sharpenRow.appendChild(b);
  }
  menu.appendChild(sharpenRow);

  const bcHead = document.createElement("div");
  bcHead.className = "fs-pic-corr-menu__section-head";
  bcHead.textContent = "亮度/对比度";
  menu.appendChild(bcHead);
  const grid = document.createElement("div");
  grid.className = "fs-pic-corr-menu__grid";
  const gridBtns: HTMLButtonElement[] = [];
  for (let r = 0; r < 5; r++) {
    for (let c = 0; c < 5; c++) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fs-pic-corr-menu__thumb fs-pic-corr-menu__thumb--cell";
      b.setAttribute("role", "menuitem");
      const img = document.createElement("img");
      img.className = "fs-pic-corr-menu__thumb-img";
      img.alt = "";
      img.draggable = false;
      b.appendChild(img);
      const rr = r;
      const cc = c;
      b.addEventListener("click", (ev) => {
        ev.stopPropagation();
        applyPatch({
          brightnessPct: BC_BRIGHT[rr],
          contrastPct: BC_CONT[cc],
        });
      });
      gridBtns.push(b);
      grid.appendChild(b);
    }
  }
  menu.appendChild(grid);

  const sep = document.createElement("div");
  sep.className = "fs-pic-corr-menu__sep";
  menu.appendChild(sep);

  const footer = document.createElement("button");
  footer.type = "button";
  footer.className = "fs-pic-corr-menu__footer";
  footer.setAttribute("role", "menuitem");
  const ic = document.createElement("span");
  ic.className = "fs-pic-corr-menu__footer-icon";
  ic.appendChild(iconFormatPainter());
  const ft = document.createElement("span");
  ft.className = "fs-pic-corr-menu__footer-text";
  ft.textContent = "图片更正选项...";
  footer.appendChild(ic);
  footer.appendChild(ft);
  footer.addEventListener("click", (ev) => {
    ev.stopPropagation();
    const fs = getFlexSheet();
    fs?.openFloatingPictureFormatPane?.();
    menu.hidden = true;
    clearToolbarDropdownMenuPosition(menu);
    anchor.setAttribute("aria-expanded", "false");
  });
  menu.appendChild(footer);

  function applyPatch(patch: Partial<FloatingPictureAdjustmentsState>): void {
    const fs = getFlexSheet();
    fs?.setFloatingPictureAdjustmentsState?.(patch);
    refreshSelection();
  }

  function refreshThumbnails(): void {
    const src = previewUrl();
    for (let i = 0; i < sharpenBtns.length; i++) {
      const img = sharpenBtns[i].querySelector("img");
      if (img instanceof HTMLImageElement) {
        img.src = src;
        img.style.filter = buildPreviewFilter(mergeAdj({ sharpnessPct: SHARP_PRESETS[i] }));
      }
    }
    let k = 0;
    for (let r = 0; r < 5; r++) {
      for (let c = 0; c < 5; c++) {
        const img = gridBtns[k]?.querySelector("img");
        if (img instanceof HTMLImageElement) {
          img.src = src;
          img.style.filter = buildPreviewFilter(
            mergeAdj({ brightnessPct: BC_BRIGHT[r], contrastPct: BC_CONT[c] }),
          );
        }
        k += 1;
      }
    }
  }

  function refreshSelection(): void {
    const fs = getFlexSheet();
    const adj = fs?.getFloatingPictureAdjustmentsState?.() ?? DEF;
    for (let i = 0; i < sharpenBtns.length; i++) {
      sharpenBtns[i].classList.toggle(
        "fs-pic-corr-menu__thumb--selected",
        adj.sharpnessPct === SHARP_PRESETS[i],
      );
    }
    let k = 0;
    for (let r = 0; r < 5; r++) {
      for (let c = 0; c < 5; c++) {
        const sel =
          adj.brightnessPct === BC_BRIGHT[r] && adj.contrastPct === BC_CONT[c];
        gridBtns[k]?.classList.toggle("fs-pic-corr-menu__thumb--selected", sel);
        k += 1;
      }
    }
  }

  const ribbonRoot = anchor.closest(".fs-ribbon");
  (ribbonRoot ?? document.body).appendChild(menu);
  anchor.setAttribute("aria-haspopup", "menu");
  anchor.setAttribute("aria-expanded", "false");

  anchor.addEventListener("fs-dropdown-toggle", (ev) => {
    ev.stopPropagation();
    if (!menu.hidden) {
      menu.hidden = true;
      clearToolbarDropdownMenuPosition(menu);
      anchor.setAttribute("aria-expanded", "false");
      return;
    }
    closeAllRibbonPopups();
    refreshThumbnails();
    refreshSelection();
    menu.hidden = false;
    syncToolbarDropdownMenuPosition(anchor, menu);
    anchor.setAttribute("aria-expanded", "true");
  });
}

/** 无图时的极简占位（灰块），避免空 src。 */
const PLACEHOLDER_IMG =
  "data:image/svg+xml," +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="48" height="36" viewBox="0 0 48 36"><rect fill="#c8c6c4" width="48" height="36"/><path fill="#a19f9d" d="M4 28l12-14 8 10 10-12 10 16H4z"/></svg>`,
  );

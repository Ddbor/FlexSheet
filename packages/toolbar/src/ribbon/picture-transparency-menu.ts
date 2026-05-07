import { iconFormatPainter } from "../toolbar/icons.js";
import {
  clearToolbarDropdownMenuPosition,
  closeAllRibbonPopups,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import { buildFloatingPictureCssFilterPreview } from "./picture-adjust-preview.js";
import type { FlexSheetLike, FloatingPictureAdjustmentsState } from "./ribbon-types.js";

const DEF: FloatingPictureAdjustmentsState = {
  brightnessPct: 0,
  contrastPct: 0,
  sharpnessPct: 0,
  saturationPct: 100,
  colorTemperatureK: 6500,
  transparencyPct: 0,
  recolorPreset: "none",
};

/**
 * 与 Excel「图片格式 → 透明度」预设一致：左不透明 → 右高透明（共 7 档）。
 */
const TRANSPARENCY_PRESETS = [0, 15, 30, 50, 65, 80, 95] as const;

function mergeAdj(partial: Partial<FloatingPictureAdjustmentsState>): FloatingPictureAdjustmentsState {
  return { ...DEF, ...partial };
}

/**
 * 「图片格式 → 透明度」：预设透明度一行 + 打开右侧格式窗格。
 */
export function mountPictureTransparencyMenu(
  anchor: HTMLButtonElement,
  getFlexSheet: () => FlexSheetLike | undefined,
): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-picture-format-transparency";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-pic-transparency-menu";
  menu.hidden = true;
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);
  menu.setAttribute("role", "menu");

  const previewUrl = (): string => {
    const fs = getFlexSheet();
    const u = fs?.getSelectedFloatingPictureDataUrl?.() ?? null;
    return u !== null && u !== "" ? u : PLACEHOLDER_IMG;
  };

  const head = document.createElement("div");
  head.className = "fs-pic-corr-menu__section-head";
  head.textContent = "预设透明度";
  menu.appendChild(head);

  const row = document.createElement("div");
  row.className = "fs-pic-corr-menu__sh-row";
  const btns: HTMLButtonElement[] = [];
  for (let i = 0; i < TRANSPARENCY_PRESETS.length; i++) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "fs-pic-corr-menu__thumb";
    b.setAttribute("role", "menuitem");
    b.title = `${TRANSPARENCY_PRESETS[i]}%`;
    const img = document.createElement("img");
    img.className = "fs-pic-corr-menu__thumb-img";
    img.alt = "";
    img.draggable = false;
    b.appendChild(img);
    const idx = i;
    b.addEventListener("click", (ev) => {
      ev.stopPropagation();
      applyPatch({ transparencyPct: TRANSPARENCY_PRESETS[idx] });
    });
    btns.push(b);
    row.appendChild(b);
  }
  menu.appendChild(row);

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
  ft.textContent = "图片透明度选项...";
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
    for (let i = 0; i < btns.length; i++) {
      const img = btns[i]!.querySelector("img");
      if (img instanceof HTMLImageElement) {
        img.src = src;
        img.style.filter = buildFloatingPictureCssFilterPreview(
          mergeAdj({ transparencyPct: TRANSPARENCY_PRESETS[i] }),
        );
      }
    }
  }

  function refreshSelection(): void {
    const fs = getFlexSheet();
    const adj = fs?.getFloatingPictureAdjustmentsState?.() ?? DEF;
    for (let i = 0; i < btns.length; i++) {
      btns[i]!.classList.toggle(
        "fs-pic-corr-menu__thumb--selected",
        adj.transparencyPct === TRANSPARENCY_PRESETS[i],
      );
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

const PLACEHOLDER_IMG =
  "data:image/svg+xml," +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="48" height="36" viewBox="0 0 48 36"><rect fill="#c8c6c4" width="48" height="36"/><path fill="#a19f9d" d="M4 28l12-14 8 10 10-12 10 16H4z"/></svg>`,
  );

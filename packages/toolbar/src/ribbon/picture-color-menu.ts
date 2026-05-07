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

/** 与 flexsheet `FLOATING_PICTURE_RECOLOR_PRESET_GRID` 一致 */
const RECOLOR_GRID: readonly (readonly string[])[] = [
  ["none", "r_gray", "r_sepia", "r_washout", "r_bwsoft", "r_bwmid", "r_bwhard"],
  ["r_dkgray", "r_dkblue", "r_dkorange", "r_dksilver", "r_dkgold", "r_dklblue", "r_dkgreen"],
  ["r_plgray", "r_plblue", "r_plorange", "r_plsilver", "r_plgold", "r_pllblue", "r_plgreen"],
];

/** 左 → 右：低饱和 → 高饱和 */
const SATURATION_PRESETS = [0, 40, 70, 100, 130, 160, 200] as const;
/** 左 → 右：偏冷 → 偏暖（K 值） */
const TEMPERATURE_PRESETS = [11000, 10000, 8500, 6500, 4500, 3200, 2000] as const;

function mergeAdj(partial: Partial<FloatingPictureAdjustmentsState>): FloatingPictureAdjustmentsState {
  return { ...DEF, ...partial };
}

function recolorSelectionPatch(preset: string): Partial<FloatingPictureAdjustmentsState> {
  if (preset === "none") {
    return { recolorPreset: "none" };
  }
  return {
    recolorPreset: preset,
    brightnessPct: 0,
    contrastPct: 0,
    saturationPct: 100,
    colorTemperatureK: 6500,
  };
}

/**
 * 「图片格式 → 颜色」：饱和度 / 色温行 + 重新着色 3×7 + 打开右侧格式窗格。
 */
export function mountPictureColorMenu(
  anchor: HTMLButtonElement,
  getFlexSheet: () => FlexSheetLike | undefined,
): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-picture-format-color";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-pic-color-menu";
  menu.hidden = true;
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);
  menu.setAttribute("role", "menu");

  const previewUrl = (): string => {
    const fs = getFlexSheet();
    const u = fs?.getSelectedFloatingPictureDataUrl?.() ?? null;
    return u !== null && u !== "" ? u : PLACEHOLDER_IMG;
  };

  const headSat = document.createElement("div");
  headSat.className = "fs-pic-corr-menu__section-head";
  headSat.textContent = "颜色饱和度";
  menu.appendChild(headSat);
  const satRow = document.createElement("div");
  satRow.className = "fs-pic-corr-menu__sh-row";
  const satBtns: HTMLButtonElement[] = [];
  for (let i = 0; i < SATURATION_PRESETS.length; i++) {
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
      applyPatch({
        saturationPct: SATURATION_PRESETS[idx],
        recolorPreset: "none",
      });
    });
    satBtns.push(b);
    satRow.appendChild(b);
  }
  menu.appendChild(satRow);

  const headTone = document.createElement("div");
  headTone.className = "fs-pic-corr-menu__section-head";
  headTone.textContent = "色调";
  menu.appendChild(headTone);
  const toneRow = document.createElement("div");
  toneRow.className = "fs-pic-corr-menu__sh-row";
  const toneBtns: HTMLButtonElement[] = [];
  for (let i = 0; i < TEMPERATURE_PRESETS.length; i++) {
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
      applyPatch({
        colorTemperatureK: TEMPERATURE_PRESETS[idx],
        recolorPreset: "none",
      });
    });
    toneBtns.push(b);
    toneRow.appendChild(b);
  }
  menu.appendChild(toneRow);

  const headRec = document.createElement("div");
  headRec.className = "fs-pic-corr-menu__section-head";
  headRec.textContent = "重新着色";
  menu.appendChild(headRec);
  const recGrid = document.createElement("div");
  recGrid.className = "fs-pic-color-menu__rec-grid";
  const recBtns: HTMLButtonElement[] = [];
  for (let r = 0; r < RECOLOR_GRID.length; r++) {
    const row = RECOLOR_GRID[r]!;
    for (let c = 0; c < row.length; c++) {
      const preset = row[c]!;
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fs-pic-corr-menu__thumb fs-pic-corr-menu__thumb--cell";
      b.setAttribute("role", "menuitem");
      b.dataset.fsRecolorPreset = preset;
      const img = document.createElement("img");
      img.className = "fs-pic-corr-menu__thumb-img";
      img.alt = "";
      img.draggable = false;
      b.appendChild(img);
      b.addEventListener("click", (ev) => {
        ev.stopPropagation();
        applyPatch(recolorSelectionPatch(preset));
      });
      recBtns.push(b);
      recGrid.appendChild(b);
    }
  }
  menu.appendChild(recGrid);

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
  ft.textContent = "图片颜色选项...";
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
    for (let i = 0; i < satBtns.length; i++) {
      const img = satBtns[i]!.querySelector("img");
      if (img instanceof HTMLImageElement) {
        img.src = src;
        img.style.filter = buildFloatingPictureCssFilterPreview(
          mergeAdj({ saturationPct: SATURATION_PRESETS[i], recolorPreset: "none" }),
        );
      }
    }
    for (let i = 0; i < toneBtns.length; i++) {
      const img = toneBtns[i]!.querySelector("img");
      if (img instanceof HTMLImageElement) {
        img.src = src;
        img.style.filter = buildFloatingPictureCssFilterPreview(
          mergeAdj({ colorTemperatureK: TEMPERATURE_PRESETS[i], recolorPreset: "none" }),
        );
      }
    }
    for (const b of recBtns) {
      const preset = b.dataset.fsRecolorPreset ?? "none";
      const img = b.querySelector("img");
      if (img instanceof HTMLImageElement) {
        img.src = src;
        img.style.filter = buildFloatingPictureCssFilterPreview(mergeAdj(recolorSelectionPatch(preset)));
      }
    }
  }

  function refreshSelection(): void {
    const fs = getFlexSheet();
    const adj = fs?.getFloatingPictureAdjustmentsState?.() ?? DEF;
    for (let i = 0; i < satBtns.length; i++) {
      satBtns[i]!.classList.toggle(
        "fs-pic-corr-menu__thumb--selected",
        adj.recolorPreset === "none" && adj.saturationPct === SATURATION_PRESETS[i],
      );
    }
    for (let i = 0; i < toneBtns.length; i++) {
      toneBtns[i]!.classList.toggle(
        "fs-pic-corr-menu__thumb--selected",
        adj.recolorPreset === "none" && adj.colorTemperatureK === TEMPERATURE_PRESETS[i],
      );
    }
    for (const b of recBtns) {
      const preset = b.dataset.fsRecolorPreset ?? "none";
      b.classList.toggle("fs-pic-corr-menu__thumb--selected", adj.recolorPreset === preset);
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

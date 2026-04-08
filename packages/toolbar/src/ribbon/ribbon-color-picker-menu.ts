import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";

/** 自定义区顶部 9 色（黑、浅灰、蓝灰、蓝、橙、灰、黄、天蓝、绿） */
const THEME_TOP_ROW: readonly string[] = [
  "#000000",
  "#d9d9d9",
  "#44546a",
  "#4472c4",
  "#ed7d31",
  "#7f7f7f",
  "#ffc000",
  "#5b9bd5",
  "#70ad47",
];

/**
 * 自定义区 10 列 × 6 行：每列为同色系由浅到深（Excel 风格近似）。
 * 行优先展开为 grid。
 */
const CUSTOM_COLUMNS: readonly (readonly string[])[] = [
  ["#ffffff", "#e7e6e6", "#d0cece", "#aeaaaa", "#767171", "#3b3838"],
  ["#d6dce4", "#adb9ca", "#8596b2", "#667e96", "#4a657d", "#323b4c"],
  ["#d9e2f3", "#b4c6e7", "#8faadb", "#6993d5", "#4472c4", "#2f528f"],
  ["#fce4d6", "#f8cbad", "#f4b084", "#f29565", "#ed7d31", "#c55a11"],
  ["#fff9e6", "#fff2cc", "#ffe699", "#ffd34d", "#ffc000", "#bf9000"],
  ["#e2efda", "#c5e0b4", "#a9d08e", "#8fab6a", "#70ad47", "#507e32"],
  ["#d9fcf0", "#b7f0dc", "#8fe3c8", "#5dd4b0", "#00b050", "#00754a"],
  ["#d9f3fc", "#b7e8f7", "#8fdbf2", "#5fcbed", "#00b0f0", "#0069a8"],
  ["#edeaf5", "#d4cbe9", "#bba9dd", "#a084c8", "#7030a0", "#4a216b"],
  ["#fce8ec", "#f8cdd8", "#f4b1c3", "#e895ab", "#c75071", "#8b2942"],
];

/** 标准区 10 色（高饱和） */
const STANDARD_ROW: readonly string[] = [
  "#c00000",
  "#ff0000",
  "#ffc000",
  "#ffff00",
  "#92d050",
  "#00b050",
  "#00b0f0",
  "#0070c0",
  "#002060",
  "#7030a0",
];

export type RibbonColorPickerKind = "fill" | "font";

function commandIds(kind: RibbonColorPickerKind): { pick: string; none: string; more: string } {
  if (kind === "fill") {
    return { pick: "home.font.fill.pick", none: "home.font.fill.none", more: "home.font.fill.more" };
  }
  return { pick: "home.font.color.pick", none: "home.font.color.none", more: "home.font.color.more" };
}

function normalizeHex(hex: string): string {
  const t = hex.trim().toLowerCase();
  return t.startsWith("#") ? t : `#${t}`;
}

function buildCustomGridCells(): string[][] {
  const rows = 6;
  const cols = 10;
  const grid: string[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: string[] = [];
    for (let c = 0; c < cols; c++) {
      const col = CUSTOM_COLUMNS[c];
      row.push(col !== undefined && col[r] !== undefined ? col[r]! : "#cccccc");
    }
    grid.push(row);
  }
  return grid;
}

function paletteIconSvg(): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("width", "16");
  svg.setAttribute("height", "16");
  svg.setAttribute("aria-hidden", "true");
  const g = document.createElementNS(ns, "g");
  g.setAttribute("fill", "none");
  const path = document.createElementNS(ns, "path");
  path.setAttribute("fill", "#ffc83d");
  path.setAttribute(
    "d",
    "M8.2 1.5c.35 0 .65.2.8.5l1.2 2.4 2.6.4c.45.07.63.62.3.94l-1.9 1.85.45 2.59c.08.45-.4.8-.8.58L8 11.9l-2.35 1.23c-.4.21-.88-.13-.8-.58l.45-2.6-1.9-1.84c-.32-.32-.14-.87.3-.94l2.6-.4 1.2-2.4c.15-.3.45-.5.8-.5z",
  );
  g.appendChild(path);
  const d1 = document.createElementNS(ns, "circle");
  d1.setAttribute("cx", "5.2");
  d1.setAttribute("cy", "5.2");
  d1.setAttribute("r", "1");
  d1.setAttribute("fill", "#e74c3c");
  g.appendChild(d1);
  const d2 = document.createElementNS(ns, "circle");
  d2.setAttribute("cx", "11");
  d2.setAttribute("cy", "5.5");
  d2.setAttribute("r", "0.9");
  d2.setAttribute("fill", "#2980b9");
  g.appendChild(d2);
  const d3 = document.createElementNS(ns, "circle");
  d3.setAttribute("cx", "6.5");
  d3.setAttribute("cy", "9.5");
  d3.setAttribute("r", "0.85");
  d3.setAttribute("fill", "#27ae60");
  g.appendChild(d3);
  const d4 = document.createElementNS(ns, "circle");
  d4.setAttribute("cx", "10");
  d4.setAttribute("cy", "9.8");
  d4.setAttribute("r", "0.8");
  d4.setAttribute("fill", "#9b59b6");
  g.appendChild(d4);
  svg.appendChild(g);
  return svg;
}

function createSwatch(hex: string, onPick: (h: string) => void): HTMLButtonElement {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "fs-color-menu__swatch";
  const n = normalizeHex(hex);
  b.style.backgroundColor = n;
  b.title = n;
  b.setAttribute("aria-label", n);
  b.dataset.hex = n;
  b.addEventListener("click", (ev) => {
    ev.stopPropagation();
    onPick(n);
  });
  return b;
}

/**
 * 填充色 / 字体颜色面板：自定义区、标准色、无颜色、其他颜色。
 * 挂在 `.fs-ribbon` 下以继承主题变量。
 */
export function mountRibbonColorPickerMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
  kind: RibbonColorPickerKind,
): void {
  if (anchor.id === "") {
    anchor.id = kind === "fill" ? "fs-ribbon-home-font-fill" : "fs-ribbon-home-font-color";
  }

  const ids = commandIds(kind);
  const menu = document.createElement("div");
  menu.className = "fs-color-menu";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  const closeMenu = (): void => {
    menu.hidden = true;
    clearToolbarDropdownMenuPosition(menu);
    anchor.setAttribute("aria-expanded", "false");
  };

  const pick = (hex: string): void => {
    emit(ids.pick, tab, { hex });
    closeMenu();
  };

  const heading = (text: string): HTMLDivElement => {
    const h = document.createElement("div");
    h.className = "fs-color-menu__heading";
    h.textContent = text;
    return h;
  };

  menu.appendChild(heading("自定义"));

  const topRow = document.createElement("div");
  topRow.className = "fs-color-menu__row fs-color-menu__row--top";
  for (const hex of THEME_TOP_ROW) {
    topRow.appendChild(createSwatch(hex, pick));
  }
  menu.appendChild(topRow);

  const gridWrap = document.createElement("div");
  gridWrap.className = "fs-color-menu__grid";
  const grid = buildCustomGridCells();
  for (const row of grid) {
    for (const hex of row) {
      gridWrap.appendChild(createSwatch(hex, pick));
    }
  }
  menu.appendChild(gridWrap);

  menu.appendChild(heading("标准"));

  const stdRow = document.createElement("div");
  stdRow.className = "fs-color-menu__row fs-color-menu__row--standard";
  for (const hex of STANDARD_ROW) {
    stdRow.appendChild(createSwatch(hex, pick));
  }
  menu.appendChild(stdRow);

  const noneBtn = document.createElement("button");
  noneBtn.type = "button";
  noneBtn.className = "fs-color-menu__row-btn";
  noneBtn.setAttribute("role", "menuitem");
  const noneIcon = document.createElement("span");
  noneIcon.className = "fs-color-menu__none-icon";
  noneIcon.setAttribute("aria-hidden", "true");
  const noneLab = document.createElement("span");
  noneLab.className = "fs-color-menu__row-btn-label";
  noneLab.textContent = "无颜色";
  noneBtn.appendChild(noneIcon);
  noneBtn.appendChild(noneLab);
  noneBtn.addEventListener("click", (ev) => {
    ev.stopPropagation();
    emit(ids.none, tab);
    closeMenu();
  });
  menu.appendChild(noneBtn);

  const sep = document.createElement("div");
  sep.className = "fs-color-menu__sep";
  sep.setAttribute("role", "separator");
  menu.appendChild(sep);

  const moreBtn = document.createElement("button");
  moreBtn.type = "button";
  moreBtn.className = "fs-color-menu__row-btn";
  moreBtn.setAttribute("role", "menuitem");
  const palWrap = document.createElement("span");
  palWrap.className = "fs-color-menu__palette-icon";
  palWrap.appendChild(paletteIconSvg());
  const moreLab = document.createElement("span");
  moreLab.className = "fs-color-menu__row-btn-label";
  moreLab.textContent = "其他颜色...";
  moreBtn.appendChild(palWrap);
  moreBtn.appendChild(moreLab);
  moreBtn.addEventListener("click", (ev) => {
    ev.stopPropagation();
    emit(ids.more, tab);
    closeMenu();
  });
  menu.appendChild(moreBtn);

  const ribbonRoot = anchor.closest(".fs-ribbon");
  (ribbonRoot ?? document.body).appendChild(menu);
  anchor.setAttribute("aria-haspopup", "menu");
  anchor.setAttribute("aria-expanded", "false");

  anchor.addEventListener("fs-color-picker-toggle", (ev) => {
    ev.stopPropagation();
    if (!menu.hidden) {
      closeMenu();
      return;
    }
    closeAllRibbonPopups();
    menu.hidden = false;
    syncToolbarDropdownMenuPosition(anchor, menu);
    anchor.setAttribute("aria-expanded", "true");
  });
}

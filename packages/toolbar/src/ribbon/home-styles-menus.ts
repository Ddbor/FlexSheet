import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import { createTableStyleThumbnailSvg, tableStyleCommandId } from "./table-style-gallery.js";
import { createColorScaleFlyoutThumbnail } from "./color-scale-flyout-thumbnail.js";
import { createDataBarFlyoutThumbnail } from "./data-bar-flyout-thumbnail.js";

type ConditionalMenuRow =
  | { readonly kind: "sep" }
  | {
      readonly kind: "item";
      readonly id: string;
      readonly label: string;
      readonly submenu: boolean;
    };

/** 与 Excel「条件格式」下拉结构一致：分段、无左侧图标、有子菜单项右侧三角 */
const CONDITIONAL_MENU_ROWS: readonly ConditionalMenuRow[] = [
  {
    kind: "item",
    id: "home.style.conditional.highlightCells",
    label: "突出显示单元格规则",
    submenu: true,
  },
  { kind: "item", id: "home.style.conditional.topBottom", label: "最前 / 最后规则", submenu: true },
  { kind: "sep" },
  { kind: "item", id: "home.style.conditional.dataBars", label: "数据条", submenu: true },
  { kind: "item", id: "home.style.conditional.colorScales", label: "色阶", submenu: true },
  // { kind: "item", id: "home.style.conditional.iconSets", label: "图标集", submenu: true },
  { kind: "sep" },
  { kind: "item", id: "home.style.conditional.newRule", label: "新建规则...", submenu: false },
  {
    kind: "item",
    id: "home.style.conditional.clearRulesFromSelection",
    label: "清除所选单元格的规则",
    submenu: false,
  },
  {
    kind: "item",
    id: "home.style.conditional.clearRulesFromSheet",
    label: "清除整个工作表的规则",
    submenu: false,
  },
  { kind: "item", id: "home.style.conditional.manageRules", label: "管理规则...", submenu: false },
];

type ConditionalSubmenuRow =
  | { readonly kind: "sep" }
  | { readonly kind: "item"; readonly id: string; readonly label: string };

/** 「突出显示单元格规则」右侧二级菜单（纯文字、无图标） */
const HIGHLIGHT_CELLS_FLYOUT_ROWS: readonly ConditionalSubmenuRow[] = [
  { kind: "item", id: "home.style.conditional.highlightCells.greaterThan", label: "大于..." },
  { kind: "item", id: "home.style.conditional.highlightCells.lessThan", label: "小于..." },
  { kind: "item", id: "home.style.conditional.highlightCells.between", label: "介于..." },
  { kind: "item", id: "home.style.conditional.highlightCells.equalTo", label: "等于..." },
  { kind: "item", id: "home.style.conditional.highlightCells.textContains", label: "文本包含..." },
  { kind: "item", id: "home.style.conditional.highlightCells.dateOccurring", label: "发生日期..." },
  { kind: "item", id: "home.style.conditional.highlightCells.duplicateValues", label: "重复值..." },
  { kind: "sep" },
  { kind: "item", id: "home.style.conditional.highlightCells.moreRules", label: "其他规则..." },
];

/** 「最前 / 最后规则」右侧二级菜单（纯文字、无图标） */
const TOP_BOTTOM_FLYOUT_ROWS: readonly ConditionalSubmenuRow[] = [
  { kind: "item", id: "home.style.conditional.topBottom.top10Items", label: "前 10 项..." },
  { kind: "item", id: "home.style.conditional.topBottom.top10Percent", label: "前 10%..." },
  { kind: "item", id: "home.style.conditional.topBottom.bottom10Items", label: "最后 10 项..." },
  { kind: "item", id: "home.style.conditional.topBottom.bottom10Percent", label: "最后 10%..." },
  { kind: "item", id: "home.style.conditional.topBottom.aboveAverage", label: "高于平均值..." },
  { kind: "item", id: "home.style.conditional.topBottom.belowAverage", label: "低于平均值..." },
  { kind: "sep" },
  { kind: "item", id: "home.style.conditional.topBottom.moreRules", label: "其他规则..." },
];

const DATA_BAR_GRADIENT_PRESETS: readonly {
  readonly id: string;
  readonly color: string;
  readonly name: string;
}[] = [
  { id: "home.style.conditional.dataBars.gradient.blue", color: "#638ec6", name: "蓝色" },
  { id: "home.style.conditional.dataBars.gradient.green", color: "#5cb85c", name: "绿色" },
  { id: "home.style.conditional.dataBars.gradient.red", color: "#e74c3c", name: "红色" },
  { id: "home.style.conditional.dataBars.gradient.yellow", color: "#f1c40f", name: "黄色" },
  { id: "home.style.conditional.dataBars.gradient.cyan", color: "#17c0d8", name: "青色" },
  { id: "home.style.conditional.dataBars.gradient.pink", color: "#e91e8c", name: "粉红" },
];

const DATA_BAR_SOLID_PRESETS: readonly {
  readonly id: string;
  readonly color: string;
  readonly name: string;
}[] = [
  { id: "home.style.conditional.dataBars.solid.blue", color: "#638ec6", name: "蓝色" },
  { id: "home.style.conditional.dataBars.solid.green", color: "#5cb85c", name: "绿色" },
  { id: "home.style.conditional.dataBars.solid.red", color: "#e74c3c", name: "红色" },
  { id: "home.style.conditional.dataBars.solid.yellow", color: "#f1c40f", name: "黄色" },
  { id: "home.style.conditional.dataBars.solid.darkBlue", color: "#2f5597", name: "深蓝" },
  { id: "home.style.conditional.dataBars.solid.pink", color: "#e91e8c", name: "粉红" },
];

/** 「色阶」快捷项：与 Excel 常见 12 种一致（3×4 + 其他规则） */
const COLOR_SCALE_FLYOUT_PRESETS: readonly {
  readonly id: string;
  readonly name: string;
  readonly kind: "two" | "three";
  readonly min: string;
  readonly mid?: string;
  readonly max: string;
}[] = [
  {
    id: "home.style.conditional.colorScales.gyr",
    name: "绿黄红",
    kind: "three",
    min: "#63be7b",
    mid: "#ffeb84",
    max: "#f8696b",
  },
  {
    id: "home.style.conditional.colorScales.ryg",
    name: "红黄绿",
    kind: "three",
    min: "#f8696b",
    mid: "#ffeb84",
    max: "#63be7b",
  },
  {
    id: "home.style.conditional.colorScales.gwr",
    name: "绿白红",
    kind: "three",
    min: "#63be7b",
    mid: "#ffffff",
    max: "#f8696b",
  },
  {
    id: "home.style.conditional.colorScales.rwg",
    name: "红白绿",
    kind: "three",
    min: "#f8696b",
    mid: "#ffffff",
    max: "#63be7b",
  },
  {
    id: "home.style.conditional.colorScales.bwr",
    name: "蓝白红",
    kind: "three",
    min: "#638ec6",
    mid: "#ffffff",
    max: "#f8696b",
  },
  {
    id: "home.style.conditional.colorScales.rwb",
    name: "红白蓝",
    kind: "three",
    min: "#f8696b",
    mid: "#ffffff",
    max: "#638ec6",
  },
  {
    id: "home.style.conditional.colorScales.whiteRed",
    name: "白红",
    kind: "two",
    min: "#ffffff",
    max: "#f8696b",
  },
  {
    id: "home.style.conditional.colorScales.redWhite",
    name: "红白",
    kind: "two",
    min: "#f8696b",
    max: "#ffffff",
  },
  {
    id: "home.style.conditional.colorScales.greenWhite",
    name: "绿白",
    kind: "two",
    min: "#63be7b",
    max: "#ffffff",
  },
  {
    id: "home.style.conditional.colorScales.whiteGreen",
    name: "白绿",
    kind: "two",
    min: "#ffffff",
    max: "#63be7b",
  },
  {
    id: "home.style.conditional.colorScales.greenYellow",
    name: "绿黄",
    kind: "two",
    min: "#63be7b",
    max: "#ffeb84",
  },
  {
    id: "home.style.conditional.colorScales.yellowGreen",
    name: "黄绿",
    kind: "two",
    min: "#ffeb84",
    max: "#63be7b",
  },
];

const HIGHLIGHT_CELLS_PARENT_ID = "home.style.conditional.highlightCells";
const TOP_BOTTOM_PARENT_ID = "home.style.conditional.topBottom";
const DATA_BARS_PARENT_ID = "home.style.conditional.dataBars";
const COLOR_SCALES_PARENT_ID = "home.style.conditional.colorScales";

/** 「色阶」二级菜单：3×4 彩色 SVG +「其他规则」 */
function mountColorScalesConditionalFlyout(flyout: HTMLElement): void {
  flyout.classList.add("fs-bd-menu--conditional-flyout--data-bars");

  const grid = document.createElement("div");
  grid.className = "fs-bd-menu__db-grid";
  for (const p of COLOR_SCALE_FLYOUT_PRESETS) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "fs-bd-menu__db-thumb-btn";
    b.dataset.commandId = p.id;
    b.setAttribute("aria-label", `色阶：${p.name}`);
    const thumb =
      p.kind === "three"
        ? createColorScaleFlyoutThumbnail("three", p.min, p.max, p.mid)
        : createColorScaleFlyoutThumbnail("two", p.min, p.max);
    b.appendChild(thumb);
    grid.appendChild(b);
  }
  flyout.appendChild(grid);

  const sep = document.createElement("div");
  sep.className = "fs-bd-menu__sep";
  sep.setAttribute("role", "separator");
  sep.setAttribute("aria-hidden", "true");
  flyout.appendChild(sep);

  const more = document.createElement("button");
  more.type = "button";
  more.className = "fs-bd-menu__item fs-bd-menu__item--no-icon";
  more.setAttribute("role", "menuitem");
  more.dataset.commandId = "home.style.conditional.colorScales.moreRules";
  const ml = document.createElement("span");
  ml.className = "fs-bd-menu__label";
  ml.textContent = "其他规则...";
  more.appendChild(ml);
  flyout.appendChild(more);
}

/** 「数据条」二级菜单：分组标题 + 彩色 SVG 缩略图网格 +「其他规则」 */
function mountDataBarsConditionalFlyout(flyout: HTMLElement): void {
  flyout.classList.add("fs-bd-menu--conditional-flyout--data-bars");

  const mkHead = (text: string): HTMLDivElement => {
    const h = document.createElement("div");
    h.className = "fs-bd-menu__db-sec-head";
    h.textContent = text;
    return h;
  };

  const mkGrid = (
    fill: "gradient" | "solid",
    presets: readonly { readonly id: string; readonly color: string; readonly name: string }[],
    sectionLabel: string,
  ): void => {
    flyout.appendChild(mkHead(sectionLabel));
    const grid = document.createElement("div");
    grid.className = "fs-bd-menu__db-grid";
    for (const p of presets) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fs-bd-menu__db-thumb-btn";
      b.dataset.commandId = p.id;
      b.setAttribute("aria-label", `${sectionLabel}：${p.name}`);
      b.appendChild(createDataBarFlyoutThumbnail(fill, p.color));
      grid.appendChild(b);
    }
    flyout.appendChild(grid);
  };

  mkGrid("gradient", DATA_BAR_GRADIENT_PRESETS, "渐变填充");
  mkGrid("solid", DATA_BAR_SOLID_PRESETS, "实心填充");

  const sep = document.createElement("div");
  sep.className = "fs-bd-menu__sep";
  sep.setAttribute("role", "separator");
  sep.setAttribute("aria-hidden", "true");
  flyout.appendChild(sep);

  const more = document.createElement("button");
  more.type = "button";
  more.className = "fs-bd-menu__item fs-bd-menu__item--no-icon";
  more.setAttribute("role", "menuitem");
  more.dataset.commandId = "home.style.conditional.dataBars.moreRules";
  const ml = document.createElement("span");
  ml.className = "fs-bd-menu__label";
  ml.textContent = "其他规则...";
  more.appendChild(ml);
  flyout.appendChild(more);
}

function syncConditionalFlyoutToRow(flyout: HTMLElement, row: HTMLElement): void {
  const r = row.getBoundingClientRect();
  flyout.style.position = "fixed";
  const w = flyout.offsetWidth || 200;
  const left = Math.min(r.right + 2, Math.max(4, window.innerWidth - w - 4));
  flyout.style.left = `${left}px`;
  flyout.style.top = `${r.top}px`;
  flyout.style.zIndex = "5001";
}

function appendConditionalFlyoutRows(
  flyout: HTMLElement,
  rows: readonly ConditionalSubmenuRow[],
): void {
  for (const fr of rows) {
    if (fr.kind === "sep") {
      const sep = document.createElement("div");
      sep.className = "fs-bd-menu__sep";
      sep.setAttribute("role", "separator");
      sep.setAttribute("aria-hidden", "true");
      flyout.appendChild(sep);
      continue;
    }
    const sub = document.createElement("button");
    sub.type = "button";
    sub.className = "fs-bd-menu__item fs-bd-menu__item--no-icon";
    sub.setAttribute("role", "menuitem");
    sub.dataset.commandId = fr.id;
    const lab = document.createElement("span");
    lab.className = "fs-bd-menu__label";
    lab.textContent = fr.label;
    sub.appendChild(lab);
    flyout.appendChild(sub);
  }
}

/** 「条件格式」整钮展开浮动菜单（纯文字 + 分段 + 子菜单三角 + 二级菜单） */
export function mountConditionalFormatMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
): void {
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--conditional";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  const flyoutHighlight = document.createElement("div");
  flyoutHighlight.className = "fs-bd-menu fs-bd-menu--conditional-flyout";
  flyoutHighlight.hidden = true;
  flyoutHighlight.setAttribute("role", "menu");
  flyoutHighlight.setAttribute("aria-label", "突出显示单元格规则");
  flyoutHighlight.setAttribute("data-fs-conditional-flyout", "");
  appendConditionalFlyoutRows(flyoutHighlight, HIGHLIGHT_CELLS_FLYOUT_ROWS);

  const flyoutTopBottom = document.createElement("div");
  flyoutTopBottom.className = "fs-bd-menu fs-bd-menu--conditional-flyout";
  flyoutTopBottom.hidden = true;
  flyoutTopBottom.setAttribute("role", "menu");
  flyoutTopBottom.setAttribute("aria-label", "最前 / 最后规则");
  flyoutTopBottom.setAttribute("data-fs-conditional-flyout", "");
  appendConditionalFlyoutRows(flyoutTopBottom, TOP_BOTTOM_FLYOUT_ROWS);

  const flyoutDataBars = document.createElement("div");
  flyoutDataBars.className = "fs-bd-menu fs-bd-menu--conditional-flyout";
  flyoutDataBars.hidden = true;
  flyoutDataBars.setAttribute("role", "menu");
  flyoutDataBars.setAttribute("aria-label", "数据条");
  flyoutDataBars.setAttribute("data-fs-conditional-flyout", "");
  mountDataBarsConditionalFlyout(flyoutDataBars);

  const flyoutColorScales = document.createElement("div");
  flyoutColorScales.className = "fs-bd-menu fs-bd-menu--conditional-flyout";
  flyoutColorScales.hidden = true;
  flyoutColorScales.setAttribute("role", "menu");
  flyoutColorScales.setAttribute("aria-label", "色阶");
  flyoutColorScales.setAttribute("data-fs-conditional-flyout", "");
  mountColorScalesConditionalFlyout(flyoutColorScales);

  let hideFlyoutTimer: ReturnType<typeof setTimeout> | null = null;
  let highlightRowEl: HTMLButtonElement | null = null;
  let topBottomRowEl: HTMLButtonElement | null = null;
  let dataBarsRowEl: HTMLButtonElement | null = null;
  let colorScalesRowEl: HTMLButtonElement | null = null;

  const cancelHideFlyout = (): void => {
    if (hideFlyoutTimer !== null) {
      clearTimeout(hideFlyoutTimer);
      hideFlyoutTimer = null;
    }
  };

  const hideAllFlyouts = (): void => {
    cancelHideFlyout();
    flyoutHighlight.hidden = true;
    flyoutTopBottom.hidden = true;
    flyoutDataBars.hidden = true;
    flyoutColorScales.hidden = true;
    clearToolbarDropdownMenuPosition(flyoutHighlight);
    clearToolbarDropdownMenuPosition(flyoutTopBottom);
    clearToolbarDropdownMenuPosition(flyoutDataBars);
    clearToolbarDropdownMenuPosition(flyoutColorScales);
  };

  const scheduleHideFlyout = (): void => {
    cancelHideFlyout();
    hideFlyoutTimer = setTimeout(() => {
      hideFlyoutTimer = null;
      hideAllFlyouts();
    }, 220);
  };

  const showFlyoutForRow = (rowEl: HTMLButtonElement, flyout: HTMLElement): void => {
    cancelHideFlyout();
    if (flyout !== flyoutHighlight) {
      flyoutHighlight.hidden = true;
      clearToolbarDropdownMenuPosition(flyoutHighlight);
    }
    if (flyout !== flyoutTopBottom) {
      flyoutTopBottom.hidden = true;
      clearToolbarDropdownMenuPosition(flyoutTopBottom);
    }
    if (flyout !== flyoutDataBars) {
      flyoutDataBars.hidden = true;
      clearToolbarDropdownMenuPosition(flyoutDataBars);
    }
    if (flyout !== flyoutColorScales) {
      flyoutColorScales.hidden = true;
      clearToolbarDropdownMenuPosition(flyoutColorScales);
    }
    flyout.hidden = false;
    requestAnimationFrame(() => {
      syncConditionalFlyoutToRow(flyout, rowEl);
    });
  };

  const closeMain = (): void => {
    hideAllFlyouts();
    menu.hidden = true;
    clearToolbarDropdownMenuPosition(menu);
    anchor.setAttribute("aria-expanded", "false");
  };

  const wireFlyoutClicks = (flyout: HTMLElement): void => {
    for (const fr of flyout.querySelectorAll("button[data-command-id]")) {
      if (!(fr instanceof HTMLButtonElement)) {
        continue;
      }
      fr.addEventListener("click", (ev) => {
        ev.stopPropagation();
        const id = fr.dataset.commandId;
        if (id !== undefined) {
          emit(id, tab);
        }
        closeMain();
      });
    }
    flyout.addEventListener("mouseenter", cancelHideFlyout);
    flyout.addEventListener("mouseleave", scheduleHideFlyout);
  };
  wireFlyoutClicks(flyoutHighlight);
  wireFlyoutClicks(flyoutTopBottom);
  wireFlyoutClicks(flyoutDataBars);
  wireFlyoutClicks(flyoutColorScales);

  const onScrollOrResize = (): void => {
    if (!menu.isConnected) {
      return;
    }
    if (!menu.hidden) {
      syncToolbarDropdownMenuPosition(anchor, menu);
    }
    if (!flyoutHighlight.hidden && highlightRowEl !== null) {
      syncConditionalFlyoutToRow(flyoutHighlight, highlightRowEl);
    }
    if (!flyoutTopBottom.hidden && topBottomRowEl !== null) {
      syncConditionalFlyoutToRow(flyoutTopBottom, topBottomRowEl);
    }
    if (!flyoutDataBars.hidden && dataBarsRowEl !== null) {
      syncConditionalFlyoutToRow(flyoutDataBars, dataBarsRowEl);
    }
    if (!flyoutColorScales.hidden && colorScalesRowEl !== null) {
      syncConditionalFlyoutToRow(flyoutColorScales, colorScalesRowEl);
    }
  };
  window.addEventListener("scroll", onScrollOrResize, true);
  window.addEventListener("resize", onScrollOrResize);

  for (const row of CONDITIONAL_MENU_ROWS) {
    if (row.kind === "sep") {
      const sep = document.createElement("div");
      sep.className = "fs-bd-menu__sep";
      sep.setAttribute("role", "separator");
      sep.setAttribute("aria-hidden", "true");
      menu.appendChild(sep);
      continue;
    }
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "fs-bd-menu__item fs-bd-menu__item--no-icon";
    btn.setAttribute("role", "menuitem");
    btn.dataset.commandId = row.id;
    const lab = document.createElement("span");
    lab.className = "fs-bd-menu__label";
    lab.textContent = row.label;
    btn.appendChild(lab);
    if (row.submenu) {
      const chev = document.createElement("span");
      chev.className = "fs-bd-menu__subchev";
      chev.setAttribute("aria-hidden", "true");
      chev.textContent = "\u25B8";
      btn.appendChild(chev);
    }

    if (row.id === HIGHLIGHT_CELLS_PARENT_ID) {
      highlightRowEl = btn;
      btn.addEventListener("mouseenter", () => {
        showFlyoutForRow(btn, flyoutHighlight);
      });
      btn.addEventListener("mouseleave", scheduleHideFlyout);
      btn.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        if (flyoutHighlight.hidden) {
          showFlyoutForRow(btn, flyoutHighlight);
        } else {
          hideAllFlyouts();
        }
      });
    } else if (row.id === TOP_BOTTOM_PARENT_ID) {
      topBottomRowEl = btn;
      btn.addEventListener("mouseenter", () => {
        showFlyoutForRow(btn, flyoutTopBottom);
      });
      btn.addEventListener("mouseleave", scheduleHideFlyout);
      btn.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        if (flyoutTopBottom.hidden) {
          showFlyoutForRow(btn, flyoutTopBottom);
        } else {
          hideAllFlyouts();
        }
      });
    } else if (row.id === DATA_BARS_PARENT_ID) {
      dataBarsRowEl = btn;
      btn.addEventListener("mouseenter", () => {
        showFlyoutForRow(btn, flyoutDataBars);
      });
      btn.addEventListener("mouseleave", scheduleHideFlyout);
      btn.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        if (flyoutDataBars.hidden) {
          showFlyoutForRow(btn, flyoutDataBars);
        } else {
          hideAllFlyouts();
        }
      });
    } else if (row.id === COLOR_SCALES_PARENT_ID) {
      colorScalesRowEl = btn;
      btn.addEventListener("mouseenter", () => {
        showFlyoutForRow(btn, flyoutColorScales);
      });
      btn.addEventListener("mouseleave", scheduleHideFlyout);
      btn.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        if (flyoutColorScales.hidden) {
          showFlyoutForRow(btn, flyoutColorScales);
        } else {
          hideAllFlyouts();
        }
      });
    } else {
      btn.addEventListener("click", (ev) => {
        ev.stopPropagation();
        emit(row.id, tab);
        closeMain();
      });
    }
    menu.appendChild(btn);
  }

  const ribbonRoot = anchor.closest(".fs-ribbon");
  const host = ribbonRoot ?? document.body;
  host.appendChild(menu);
  host.appendChild(flyoutHighlight);
  host.appendChild(flyoutTopBottom);
  host.appendChild(flyoutDataBars);
  host.appendChild(flyoutColorScales);
  anchor.setAttribute("aria-haspopup", "menu");
  anchor.setAttribute("aria-expanded", "false");

  anchor.addEventListener("fs-dropdown-toggle", (ev) => {
    ev.stopPropagation();
    if (!menu.hidden) {
      closeMain();
      return;
    }
    closeAllRibbonPopups();
    menu.hidden = false;
    hideAllFlyouts();
    syncToolbarDropdownMenuPosition(anchor, menu);
    anchor.setAttribute("aria-expanded", "true");
  });
}

/** 「套用表格格式」整钮展开：浅色 / 中等色 / 深色 分节 + 7 列缩略图网格 */
export function mountTableFormatStyleMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
  options?: {
    readonly getCustomTableStyles?: () => readonly {
      readonly id: string;
      readonly name: string;
      readonly commandId: string;
    }[];
  },
): void {
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--table-styles";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  const mkSection = (
    title: string,
    section: "light" | "medium" | "dark",
    rowCount: number,
  ): void => {
    const head = document.createElement("div");
    head.className = "fs-ts-sec-head";
    head.textContent = title;
    menu.appendChild(head);
    const grid = document.createElement("div");
    grid.className = "fs-ts-grid";
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < 7; col++) {
        const id = tableStyleCommandId(section, row, col);
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "fs-ts-thumb-btn";
        btn.dataset.commandId = id;
        btn.setAttribute("aria-label", `${title} 第 ${row + 1} 行，第 ${col + 1} 列主题`);
        btn.appendChild(
          createTableStyleThumbnailSvg({
            section,
            row,
            col,
          }),
        );
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          emit(id, tab);
          menu.hidden = true;
          clearToolbarDropdownMenuPosition(menu);
          anchor.setAttribute("aria-expanded", "false");
        });
        grid.appendChild(btn);
      }
    }
    menu.appendChild(grid);
  };

  const closeMenu = (): void => {
    menu.hidden = true;
    clearToolbarDropdownMenuPosition(menu);
    anchor.setAttribute("aria-expanded", "false");
  };

  const renderMenuContent = (): void => {
    menu.replaceChildren();
    mkSection("浅色", "light", 3);
    mkSection("中等色", "medium", 4);
    mkSection("深色", "dark", 2);

    const customItems = options?.getCustomTableStyles?.() ?? [];
    if (customItems.length > 0) {
      const customHead = document.createElement("div");
      customHead.className = "fs-ts-sec-head";
      customHead.textContent = "自定义";
      menu.appendChild(customHead);
      const customGrid = document.createElement("div");
      customGrid.className = "fs-ts-grid";
      for (const item of customItems) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "fs-ts-thumb-btn";
        btn.dataset.commandId = item.commandId;
        btn.setAttribute("aria-label", `自定义表样式：${item.name}`);
        const parsed = parseBuiltInTableStyleCommand(item.commandId);
        btn.appendChild(
          createTableStyleThumbnailSvg(
            parsed ?? {
              section: "medium",
              row: 2,
              col: 0,
            },
          ),
        );
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          emit(item.commandId, tab);
          closeMenu();
        });
        customGrid.appendChild(btn);
      }
      menu.appendChild(customGrid);
    }

    const sep = document.createElement("div");
    sep.className = "fs-bd-menu__sep";
    sep.setAttribute("role", "separator");
    sep.setAttribute("aria-hidden", "true");
    menu.appendChild(sep);

    const newStyleBtn = document.createElement("button");
    newStyleBtn.type = "button";
    newStyleBtn.className = "fs-bd-menu__item fs-bd-menu__item--no-icon";
    newStyleBtn.setAttribute("role", "menuitem");
    newStyleBtn.dataset.commandId = "home.style.table.newStyle";
    const newStyleLabel = document.createElement("span");
    newStyleLabel.className = "fs-bd-menu__label";
    newStyleLabel.textContent = "新建表样式...";
    newStyleBtn.appendChild(newStyleLabel);
    newStyleBtn.addEventListener("click", (ev) => {
      ev.stopPropagation();
      emit("home.style.table.newStyle", tab);
      closeMenu();
    });
    menu.appendChild(newStyleBtn);
  };

  renderMenuContent();

  const ribbonRoot = anchor.closest(".fs-ribbon");
  (ribbonRoot ?? document.body).appendChild(menu);
  anchor.setAttribute("aria-haspopup", "menu");
  anchor.setAttribute("aria-expanded", "false");

  anchor.addEventListener("fs-dropdown-toggle", (ev) => {
    ev.stopPropagation();
    if (!menu.hidden) {
      closeMenu();
      return;
    }
    renderMenuContent();
    closeAllRibbonPopups();
    menu.hidden = false;
    syncToolbarDropdownMenuPosition(anchor, menu);
    anchor.setAttribute("aria-expanded", "true");
  });
}

function parseBuiltInTableStyleCommand(
  commandId: string,
): { section: "light" | "medium" | "dark"; row: number; col: number } | null {
  const m = /^home\.style\.table\.(light|medium|dark)\.r(\d+)c(\d+)$/.exec(commandId);
  if (m === null) {
    return null;
  }
  const section = m[1];
  const row = Number(m[2]);
  const col = Number(m[3]);
  if (
    (section !== "light" && section !== "medium" && section !== "dark") ||
    !Number.isInteger(row) ||
    !Number.isInteger(col) ||
    row < 0 ||
    col < 0 ||
    col > 6
  ) {
    return null;
  }
  return { section, row, col };
}

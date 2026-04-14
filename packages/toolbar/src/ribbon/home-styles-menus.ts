import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import { iconFormatGeneral } from "../toolbar/icons.js";

type StyleMenuRow = {
  readonly id: string;
  readonly label: string;
  readonly icon: () => SVGSVGElement;
};

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
  { kind: "item", id: "home.style.conditional.iconSets", label: "图标集", submenu: true },
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

type HighlightFlyoutRow =
  | { readonly kind: "sep" }
  | { readonly kind: "item"; readonly id: string; readonly label: string };

/** 「突出显示单元格规则」右侧二级菜单（纯文字） */
const HIGHLIGHT_CELLS_FLYOUT_ROWS: readonly HighlightFlyoutRow[] = [
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

const TABLE_STYLE_MENU_ITEMS: readonly StyleMenuRow[] = [
  { id: "home.style.table.light1", label: "表样式浅色 1", icon: iconFormatGeneral },
  { id: "home.style.table.light2", label: "表样式浅色 2", icon: iconFormatGeneral },
  { id: "home.style.table.medium1", label: "表样式中等深浅 1", icon: iconFormatGeneral },
  { id: "home.style.table.medium2", label: "表样式中等深浅 2", icon: iconFormatGeneral },
  { id: "home.style.table.dark1", label: "表样式深色 1", icon: iconFormatGeneral },
  { id: "home.style.table.dark2", label: "表样式深色 2", icon: iconFormatGeneral },
];

const CELL_STYLE_MENU_ITEMS: readonly StyleMenuRow[] = [
  { id: "home.style.cell.good", label: "好", icon: iconFormatGeneral },
  { id: "home.style.cell.bad", label: "差", icon: iconFormatGeneral },
  { id: "home.style.cell.neutral", label: "中性", icon: iconFormatGeneral },
  { id: "home.style.cell.normal", label: "常规", icon: iconFormatGeneral },
  { id: "home.style.cell.title", label: "标题", icon: iconFormatGeneral },
];

function mountStyleFloatingMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
  items: readonly StyleMenuRow[],
): void {
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const it of items) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fs-bd-menu__item";
    row.setAttribute("role", "menuitem");
    row.dataset.commandId = it.id;
    const iconWrap = document.createElement("span");
    iconWrap.className = "fs-bd-menu__icon";
    iconWrap.appendChild(it.icon());
    const lab = document.createElement("span");
    lab.className = "fs-bd-menu__label";
    lab.textContent = it.label;
    row.appendChild(iconWrap);
    row.appendChild(lab);
    row.addEventListener("click", (ev) => {
      ev.stopPropagation();
      emit(it.id, tab);
      menu.hidden = true;
      clearToolbarDropdownMenuPosition(menu);
      anchor.setAttribute("aria-expanded", "false");
    });
    menu.appendChild(row);
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
    menu.hidden = false;
    syncToolbarDropdownMenuPosition(anchor, menu);
    anchor.setAttribute("aria-expanded", "true");
  });
}

const HIGHLIGHT_CELLS_PARENT_ID = "home.style.conditional.highlightCells";

function syncHighlightFlyoutToRow(flyout: HTMLElement, row: HTMLElement): void {
  const r = row.getBoundingClientRect();
  flyout.style.position = "fixed";
  const w = flyout.offsetWidth || 200;
  const left = Math.min(r.right + 2, Math.max(4, window.innerWidth - w - 4));
  flyout.style.left = `${left}px`;
  flyout.style.top = `${r.top}px`;
  flyout.style.zIndex = "5001";
}

/** 「条件格式」整钮展开浮动菜单（纯文字 + 分段 + 子菜单三角 +「突出显示」二级菜单） */
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

  const flyout = document.createElement("div");
  flyout.className = "fs-bd-menu fs-bd-menu--conditional-flyout";
  flyout.hidden = true;
  flyout.setAttribute("role", "menu");
  flyout.setAttribute("aria-label", "突出显示单元格规则");
  flyout.setAttribute("data-fs-conditional-flyout", "");

  for (const fr of HIGHLIGHT_CELLS_FLYOUT_ROWS) {
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

  let hideFlyoutTimer: ReturnType<typeof setTimeout> | null = null;
  let highlightRowEl: HTMLButtonElement | null = null;

  const cancelHideFlyout = (): void => {
    if (hideFlyoutTimer !== null) {
      clearTimeout(hideFlyoutTimer);
      hideFlyoutTimer = null;
    }
  };

  const hideFlyout = (): void => {
    cancelHideFlyout();
    flyout.hidden = true;
    clearToolbarDropdownMenuPosition(flyout);
  };

  const scheduleHideFlyout = (): void => {
    cancelHideFlyout();
    hideFlyoutTimer = setTimeout(() => {
      hideFlyoutTimer = null;
      hideFlyout();
    }, 220);
  };

  const showFlyout = (): void => {
    cancelHideFlyout();
    if (highlightRowEl === null) {
      return;
    }
    flyout.hidden = false;
    requestAnimationFrame(() => {
      syncHighlightFlyoutToRow(flyout, highlightRowEl!);
    });
  };

  const closeMain = (): void => {
    cancelHideFlyout();
    hideFlyout();
    menu.hidden = true;
    clearToolbarDropdownMenuPosition(menu);
    anchor.setAttribute("aria-expanded", "false");
  };

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

  const onScrollOrResize = (): void => {
    if (!menu.isConnected) {
      return;
    }
    if (!menu.hidden) {
      syncToolbarDropdownMenuPosition(anchor, menu);
    }
    if (!flyout.hidden && highlightRowEl !== null) {
      syncHighlightFlyoutToRow(flyout, highlightRowEl);
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
        showFlyout();
      });
      btn.addEventListener("mouseleave", scheduleHideFlyout);
      btn.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        if (flyout.hidden) {
          showFlyout();
        } else {
          hideFlyout();
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
  host.appendChild(flyout);
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
    hideFlyout();
    syncToolbarDropdownMenuPosition(anchor, menu);
    anchor.setAttribute("aria-expanded", "true");
  });
}

/** 「套用表格格式」整钮展开浮动菜单 */
export function mountTableFormatStyleMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
): void {
  mountStyleFloatingMenu(anchor, emit, tab, TABLE_STYLE_MENU_ITEMS);
}

/** 「单元格样式」整钮展开浮动菜单 */
export function mountCellStyleRibbonMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
): void {
  mountStyleFloatingMenu(anchor, emit, tab, CELL_STYLE_MENU_ITEMS);
}

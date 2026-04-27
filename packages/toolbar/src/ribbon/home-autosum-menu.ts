import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";

type AutoSumItem = {
  readonly id: string;
  readonly label: string;
  readonly separatorTop?: boolean;
};

const AUTOSUM_MENU_ITEMS: readonly AutoSumItem[] = [
  { id: "autoSum.sub.sum", label: "求和" },
  { id: "autoSum.sub.average", label: "平均值" },
  { id: "autoSum.sub.count", label: "计数" },
  { id: "autoSum.sub.max", label: "最大值" },
  { id: "autoSum.sub.min", label: "最小值" },
  /** 与「公式」下拉里 `formula.fn.more` 相同，由宿主打开插入函数 / 更多函数 UI。 */
  { id: "formula.fn.more", label: "其他函数…", separatorTop: true },
];

/** 挂载 Ribbon「自动求和」下拉面板的交互（`split` 的箭头应触发展开）。 */
export function mountAutoSumSubmenu(anchor: HTMLButtonElement, emit: RibbonEmit, tab: RibbonTabId): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-cells-autosum";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--autosum";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const it of AUTOSUM_MENU_ITEMS) {
    if (it.separatorTop === true) {
      const sep = document.createElement("hr");
      sep.className = "fs-bd-menu__sep";
      sep.setAttribute("aria-hidden", "true");
      menu.appendChild(sep);
    }
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fs-bd-menu__item";
    row.setAttribute("role", "menuitem");
    row.dataset.commandId = it.id;
    const label = document.createElement("span");
    label.className = "fs-bd-menu__label";
    label.textContent = it.label;
    row.appendChild(label);
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

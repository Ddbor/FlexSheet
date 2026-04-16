import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import {
  iconFillDown,
  iconFillLeft,
  iconFillRight,
  iconFillUp,
} from "../toolbar/icons.js";

type FillMenuRow = {
  readonly id: string;
  readonly label: string;
  readonly icon?: () => SVGSVGElement;
  readonly separatorTop?: boolean;
};

const FILL_MENU_ITEMS: readonly FillMenuRow[] = [
  { id: "home.cells.fill.down", label: "向下", icon: iconFillDown },
  { id: "home.cells.fill.right", label: "向右", icon: iconFillRight },
  { id: "home.cells.fill.up", label: "向上", icon: iconFillUp },
  { id: "home.cells.fill.left", label: "向左", icon: iconFillLeft },
  { id: "home.cells.fill.series", label: "系列", separatorTop: true },
];

/** Ribbon「填充」下拉菜单。 */
export function mountHomeFillMenu(anchor: HTMLButtonElement, emit: RibbonEmit, tab: RibbonTabId): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-home-cells-fill";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--home-fill";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const it of FILL_MENU_ITEMS) {
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
    if (it.icon !== undefined) {
      const iconWrap = document.createElement("span");
      iconWrap.className = "fs-bd-menu__icon";
      iconWrap.appendChild(it.icon());
      row.appendChild(iconWrap);
    } else {
      row.classList.add("fs-bd-menu__item--no-icon");
    }
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

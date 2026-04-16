import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  clearToolbarDropdownMenuPosition,
  closeAllRibbonPopups,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";

type ClearMenuRow = {
  readonly id: string;
  readonly label: string;
};

const CLEAR_MENU_ITEMS: readonly ClearMenuRow[] = [
  { id: "home.cells.clear.all", label: "清除全部" },
  { id: "home.cells.clear.formats", label: "清除格式" },
  { id: "home.cells.clear.contents", label: "清除内容" },
];

/** Ribbon「清除」下拉菜单。 */
export function mountHomeClearMenu(anchor: HTMLButtonElement, emit: RibbonEmit, tab: RibbonTabId): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-home-cells-clear";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--home-clear";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const it of CLEAR_MENU_ITEMS) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fs-bd-menu__item fs-bd-menu__item--no-icon";
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

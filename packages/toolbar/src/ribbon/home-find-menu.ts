import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  clearToolbarDropdownMenuPosition,
  closeAllRibbonPopups,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";

type FindMenuRow = {
  readonly id: string;
  readonly label: string;
};

const ROWS: readonly FindMenuRow[] = [
  { id: "home.find.open", label: "查找…" },
  { id: "home.find.replace", label: "替换…" },
  { id: "home.find.goto", label: "转到…" },
  { id: "home.find.gotoSpecial", label: "定位条件…" },
];

/** 「开始 → 查找」浮动菜单：先 UI，项点击发对应 commandId。 */
export function mountHomeFindMenu(anchor: HTMLButtonElement, emit: RibbonEmit, tab: RibbonTabId): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-home-find";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--find";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const it of ROWS) {
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

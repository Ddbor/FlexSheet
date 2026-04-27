import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import {
  iconMenuFilterClear,
  iconMenuFilterOnly,
  iconMenuFilterReapply,
  iconMenuSortAsc,
  iconMenuSortCustom,
  iconMenuSortDesc,
} from "./home-sort-filter-icons.js";

type SortFilterRow = {
  readonly id: string;
  readonly label: string;
  readonly icon: () => SVGSVGElement;
};

const ROWS: readonly SortFilterRow[] = [
  { id: "home.sortFilter.asc", label: "升序", icon: iconMenuSortAsc },
  { id: "home.sortFilter.desc", label: "降序", icon: iconMenuSortDesc },
  { id: "home.sortFilter.custom", label: "自定义排序…", icon: iconMenuSortCustom },
  { id: "home.sortFilter.filter", label: "筛选", icon: iconMenuFilterOnly },
  { id: "home.sortFilter.clear", label: "清除筛选", icon: iconMenuFilterClear },
  { id: "home.sortFilter.reapply", label: "重新应用", icon: iconMenuFilterReapply },
];

/** 「排序和筛选」浮动菜单：仅 UI，项点击发出对应 commandId。 */
export function mountHomeSortFilterMenu(anchor: HTMLButtonElement, emit: RibbonEmit, tab: RibbonTabId): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-home-sort-filter";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--sort-filter";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const it of ROWS) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fs-bd-menu__item";
    row.setAttribute("role", "menuitem");
    row.dataset.commandId = it.id;
    const iconWrap = document.createElement("span");
    iconWrap.className = "fs-bd-menu__icon";
    iconWrap.appendChild(it.icon());
    const label = document.createElement("span");
    label.className = "fs-bd-menu__label";
    label.textContent = it.label;
    row.appendChild(iconWrap);
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

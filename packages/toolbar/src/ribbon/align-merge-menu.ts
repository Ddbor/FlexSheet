import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import { iconMerge, iconMergeAcross, iconMergeCells, iconUnmergeCells } from "../toolbar/icons.js";

type MergeMenuRow = {
  readonly id: string;
  readonly label: string;
  readonly icon: () => SVGSVGElement;
};

const MERGE_MENU_ITEMS: readonly MergeMenuRow[] = [
  { id: "home.align.merge", label: "合并后居中", icon: iconMerge },
  { id: "home.align.mergeAcross", label: "跨越合并", icon: iconMergeAcross },
  { id: "home.align.mergeCells", label: "合并单元格", icon: iconMergeCells },
  { id: "home.align.unmerge", label: "取消单元格合并", icon: iconUnmergeCells },
];

/**
 * 「合并后居中」分割按钮下拉：与 `.fs-bd-menu` 样式一致，挂在 `.fs-ribbon` 下。
 */
export function mountAlignMergeMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-home-align-merge";
  }

  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--merge";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const it of MERGE_MENU_ITEMS) {
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

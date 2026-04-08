import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import {
  iconRotateTextClockwise,
  iconRotateTextCounterClockwise,
  iconRotateTextDown,
  iconRotateTextUp,
  iconTextVertical,
} from "../toolbar/icons.js";

type Row = { readonly id: string; readonly label: string; readonly icon: () => SVGSVGElement };

const ITEMS: readonly Row[] = [
  { id: "home.align.textDirection.counterClockwise", label: "逆时针角度", icon: iconRotateTextCounterClockwise },
  { id: "home.align.textDirection.clockwise", label: "顺时针角度", icon: iconRotateTextClockwise },
  { id: "home.align.textDirection.vertical", label: "竖排文字", icon: iconTextVertical },
  { id: "home.align.textDirection.rotateUp", label: "向上旋转文字", icon: iconRotateTextUp },
  { id: "home.align.textDirection.rotateDown", label: "向下旋转文字", icon: iconRotateTextDown },
];

/** 「方向」分割按钮下拉 */
export function mountAlignOrientationMenu(anchor: HTMLButtonElement, emit: RibbonEmit, tab: RibbonTabId): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-home-align-orientation";
  }

  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--orientation";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const it of ITEMS) {
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

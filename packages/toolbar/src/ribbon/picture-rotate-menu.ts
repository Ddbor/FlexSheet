import {
  clearToolbarDropdownMenuPosition,
  closeAllRibbonPopups,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import type { FlexSheetLike } from "./ribbon-types.js";

/**
 * 「图片格式 → 旋转」：向右/向左 90°。
 */
export function mountPictureRotateMenu(
  anchor: HTMLButtonElement,
  getFlexSheet: () => FlexSheetLike | undefined,
): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-picture-format-rotate";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--picture-format";
  menu.hidden = true;
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);
  menu.setAttribute("role", "menu");

  const mkItem = (label: string, clockwise: boolean): void => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fs-bd-menu__item fs-bd-menu__item--no-icon";
    row.setAttribute("role", "menuitem");
    const span = document.createElement("span");
    span.className = "fs-bd-menu__label";
    span.textContent = label;
    row.appendChild(span);
    row.addEventListener("click", (ev) => {
      ev.stopPropagation();
      getFlexSheet()?.rotateSelectedFloatingPicture90Degrees?.(clockwise);
      menu.hidden = true;
      clearToolbarDropdownMenuPosition(menu);
      anchor.setAttribute("aria-expanded", "false");
    });
    menu.appendChild(row);
  };

  mkItem("向右旋转 90°", true);
  mkItem("向左旋转 90°", false);

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

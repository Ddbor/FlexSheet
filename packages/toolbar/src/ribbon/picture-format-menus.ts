import {
  clearToolbarDropdownMenuPosition,
  closeAllRibbonPopups,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";

/**
 * 图片格式下拉占位菜单：仅关闭菜单，不派发命令（功能后续再接）。
 */
export function mountPictureFormatPlaceholderMenu(
  anchor: HTMLButtonElement,
  items: readonly string[],
): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-picture-format-menu-anchor";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--picture-format";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (const text of items) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fs-bd-menu__item fs-bd-menu__item--no-icon";
    row.setAttribute("role", "menuitem");
    const label = document.createElement("span");
    label.className = "fs-bd-menu__label";
    label.textContent = text;
    row.appendChild(label);
    row.addEventListener("click", (ev) => {
      ev.stopPropagation();
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

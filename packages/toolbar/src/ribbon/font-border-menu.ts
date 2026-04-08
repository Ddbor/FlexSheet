import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";
import {
  iconBorderAll,
  iconBorderBottom,
  iconBorderDoubleBottom,
  iconBorderLeft,
  iconBorderNone,
  iconBorderOutside,
  iconBorderRight,
  iconBorderThickBottom,
  iconBorderThickBox,
  iconBorderTop,
  iconBorderTopBottom,
  iconBorderTopDoubleBottom,
  iconBorderTopThickBottom,
} from "../toolbar/icons.js";

type BorderMenuRow = {
  readonly id: string;
  readonly label: string;
  readonly icon: () => SVGSVGElement;
};

const BORDER_MENU_SECTIONS: readonly BorderMenuRow[][] = [
  [
    { id: "home.font.border.bottom", label: "下边框", icon: iconBorderBottom },
    { id: "home.font.border.top", label: "上边框", icon: iconBorderTop },
    { id: "home.font.border.left", label: "左边框", icon: iconBorderLeft },
    { id: "home.font.border.right", label: "右边框", icon: iconBorderRight },
  ],
  [
    { id: "home.font.border.none", label: "无框线", icon: iconBorderNone },
    { id: "home.font.border.all", label: "所有框线", icon: iconBorderAll },
    { id: "home.font.border.outside", label: "外侧框线", icon: iconBorderOutside },
    { id: "home.font.border.thickBox", label: "粗匣框线", icon: iconBorderThickBox },
  ],
  [
    { id: "home.font.border.doubleBottom", label: "双底框线", icon: iconBorderDoubleBottom },
    { id: "home.font.border.thickBottom", label: "粗底框线", icon: iconBorderThickBottom },
    { id: "home.font.border.topBottom", label: "上下框线", icon: iconBorderTopBottom },
    {
      id: "home.font.border.topThickBottom",
      label: "上框线和粗下框线",
      icon: iconBorderTopThickBottom,
    },
    {
      id: "home.font.border.topDoubleBottom",
      label: "上框线和双下框线",
      icon: iconBorderTopDoubleBottom,
    },
  ],
];

/**
 * 将边框下拉箭头与浮动菜单关联；菜单挂在 `.fs-ribbon` 根下以继承主题变量（背景色等），与 `.fs-dd` 共用定位与全局关闭逻辑。
 */
export function mountFontBorderMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-home-font-border";
  }

  const menu = document.createElement("div");
  menu.className = "fs-bd-menu";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  for (let si = 0; si < BORDER_MENU_SECTIONS.length; si++) {
    if (si > 0) {
      const sep = document.createElement("div");
      sep.className = "fs-bd-menu__sep";
      sep.setAttribute("role", "separator");
      menu.appendChild(sep);
    }
    const section = BORDER_MENU_SECTIONS[si];
    if (section === undefined) {
      continue;
    }
    for (const it of section) {
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

import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";

export interface FormulaRibbonFnMenuItem {
  readonly commandId: string;
  readonly label: string;
  readonly separatorTop?: boolean;
  readonly payload?: Readonly<Record<string, unknown>>;
}

/** 公式选项卡分类下拉里动态填充项（每次展开时调用 `getItems`）。 */
export function mountFormulaRibbonFnMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
  getItems: () => readonly FormulaRibbonFnMenuItem[],
): void {
  if (anchor.id === "") {
    anchor.id = "fs-ribbon-formula-fn-menu-anchor";
  }
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--formula-fn";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  const rebuild = (): void => {
    menu.replaceChildren();
    for (const it of getItems()) {
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
      row.dataset.commandId = it.commandId;
      const label = document.createElement("span");
      label.className = "fs-bd-menu__label";
      label.textContent = it.label;
      row.appendChild(label);
      row.addEventListener("click", (ev) => {
        ev.stopPropagation();
        if (it.payload !== undefined) {
          emit(it.commandId, tab, it.payload);
        } else {
          emit(it.commandId, tab);
        }
        menu.hidden = true;
        clearToolbarDropdownMenuPosition(menu);
        anchor.setAttribute("aria-expanded", "false");
      });
      menu.appendChild(row);
    }
  };

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
    rebuild();
    menu.hidden = false;
    syncToolbarDropdownMenuPosition(anchor, menu);
    anchor.setAttribute("aria-expanded", "true");
  });
}

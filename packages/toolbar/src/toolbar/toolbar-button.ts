import type { RibbonTabId } from "../ribbon/ribbon-types.js";
import { iconChevronDown } from "./icons.js";

export interface ToolbarButtonOptions {
  readonly id: string;
  readonly tab: RibbonTabId;
  readonly label: string;
  /** 大号：上图下文（剪贴板粘贴等） */
  readonly variant?: "default" | "large";
  readonly icon?: SVGSVGElement;
  /** 带下拉箭头，点击箭头展开菜单（需配合 onOpenDropdown） */
  readonly splitDropdown?: boolean;
  /**
   * 大号 Ribbon 专用：整钮点击只触发展开菜单（`fs-dropdown-toggle`），不发送 `id`；
   * 布局为图标、标签、下拉箭头自上而下，无分割条（与 `splitDropdown` 互斥，优先本项）。
   */
  readonly menuTrigger?: boolean;
  /** 点击仅触发展开颜色面板（`fs-color-picker-toggle`），不发送 commandId */
  readonly colorPickerToggle?: boolean;
  readonly title?: string;
  readonly disabled?: boolean;
}

export interface RibbonEmit {
  (id: string, tab: RibbonTabId, payload?: Readonly<Record<string, unknown>>): void;
}

/**
 * Ribbon 内标准按钮（小：横排图标+文字；大：竖排）。
 */
export function createToolbarButton(
  options: ToolbarButtonOptions,
  emit: RibbonEmit,
): {
  readonly element: HTMLButtonElement;
  setDisabled(v: boolean): void;
  setPressed(pressed: boolean): void;
} {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className =
    options.variant === "large" ? "fs-tb-btn fs-tb-btn--large" : "fs-tb-btn fs-tb-btn--default";
  btn.dataset.commandId = options.id;
  if (options.title !== undefined) {
    btn.title = options.title;
  } else {
    btn.title = options.label;
  }
  if (options.disabled === true) {
    btn.disabled = true;
  }

  const row = document.createElement("span");
  row.className = "fs-tb-btn__row";

  if (options.menuTrigger === true) {
    btn.classList.add("fs-tb-btn--ribbon-menu");
    if (options.icon !== undefined) {
      const wrap = document.createElement("span");
      wrap.className = "fs-tb-btn__icon";
      wrap.appendChild(options.icon);
      row.appendChild(wrap);
    }
    const text = document.createElement("span");
    text.className = "fs-tb-btn__label";
    text.textContent = options.label;
    const chevWrap = document.createElement("span");
    chevWrap.className = "fs-tb-btn__menu-chev-wrap";
    chevWrap.setAttribute("aria-hidden", "true");
    const chev = iconChevronDown();
    chev.classList.add("fs-tb-btn__menu-chev");
    chevWrap.appendChild(chev);
    row.appendChild(text);
    row.appendChild(chevWrap);
    btn.appendChild(row);
  } else {
    if (options.icon !== undefined) {
      const wrap = document.createElement("span");
      wrap.className = "fs-tb-btn__icon";
      wrap.appendChild(options.icon);
      row.appendChild(wrap);
    }

    const text = document.createElement("span");
    text.className = "fs-tb-btn__label";
    text.textContent = options.label;
    row.appendChild(text);

    if (options.splitDropdown === true) {
      const split = document.createElement("span");
      split.className = "fs-tb-btn__split";
      const chev = iconChevronDown();
      chev.classList.add("fs-tb-btn__chev");
      split.appendChild(chev);
      btn.appendChild(row);
      btn.appendChild(split);
      btn.classList.add("fs-tb-btn--split");
    } else {
      btn.appendChild(row);
    }
  }

  btn.addEventListener("click", (ev) => {
    if (btn.disabled) {
      return;
    }
    if (options.colorPickerToggle === true) {
      ev.stopPropagation();
      btn.dispatchEvent(new CustomEvent("fs-color-picker-toggle", { bubbles: true }));
      return;
    }
    if (options.menuTrigger === true) {
      ev.stopPropagation();
      btn.dispatchEvent(new CustomEvent("fs-dropdown-toggle", { bubbles: true }));
      return;
    }
    if (options.splitDropdown === true) {
      const target = ev.target as Node;
      if (btn.querySelector(".fs-tb-btn__split")?.contains(target)) {
        ev.stopPropagation();
        btn.dispatchEvent(new CustomEvent("fs-dropdown-toggle", { bubbles: true }));
        return;
      }
    }
    emit(options.id, options.tab);
  });

  return {
    element: btn,
    setDisabled(v: boolean): void {
      btn.disabled = v;
    },
    /** 用于加粗/斜体等切换按钮，与 `.fs-tb-btn[aria-pressed="true"]` 样式一致。 */
    setPressed(pressed: boolean): void {
      if (pressed) {
        btn.setAttribute("aria-pressed", "true");
      } else {
        btn.removeAttribute("aria-pressed");
      }
    },
  };
}

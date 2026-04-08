import type { RibbonTabId } from "../ribbon/ribbon-types.js";
import type { RibbonEmit } from "./toolbar-button.js";
import { iconChevronDown } from "./icons.js";

export interface DropdownItem {
  readonly id: string;
  readonly label: string;
  readonly disabled?: boolean;
  /** 菜单项以该 CSS font-family 展示名称（字体预览） */
  readonly previewFontFamily?: string;
  /** 左侧图标（与 Ribbon 命令图标一致，单色 currentColor） */
  readonly icon?: () => SVGSVGElement;
}

export interface ToolbarDropdownOptions {
  readonly id: string;
  readonly tab: RibbonTabId;
  readonly label: string;
  readonly items: readonly DropdownItem[];
  readonly title?: string;
  readonly wide?: boolean;
  /** 追加到 `.fs-dd__menu` 的类名（空格分隔） */
  readonly menuClassName?: string;
  /** 触发器标签初始 font-family（与默认选中字体一致） */
  readonly initialLabelFontFamily?: string;
}

/**
 * 下拉选择器：主区域点击展开，选项点击后发出命令并关闭。
 */
export function createToolbarDropdown(
  options: ToolbarDropdownOptions,
  emit: RibbonEmit,
): {
  readonly element: HTMLElement;
  close(): void;
  setLabel(text: string, previewFontFamily?: string): void;
} {
  const wrap = document.createElement("div");
  wrap.className = options.wide === true ? "fs-dd fs-dd--wide" : "fs-dd";
  wrap.dataset.dropdownId = options.id;

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "fs-dd__trigger";
  btn.title = options.title ?? options.label;
  const labelSpan = document.createElement("span");
  labelSpan.className = "fs-dd__label";
  labelSpan.textContent = options.label;
  if (options.initialLabelFontFamily !== undefined && options.initialLabelFontFamily !== "") {
    labelSpan.style.fontFamily = options.initialLabelFontFamily;
  }
  const chevSpan = document.createElement("span");
  chevSpan.className = "fs-dd__chev";
  chevSpan.setAttribute("aria-hidden", "true");
  chevSpan.appendChild(iconChevronDown());
  btn.appendChild(labelSpan);
  btn.appendChild(chevSpan);

  const menu = document.createElement("div");
  menu.className = "fs-dd__menu";
  if (options.menuClassName !== undefined && options.menuClassName !== "") {
    for (const c of options.menuClassName.trim().split(/\s+/)) {
      if (c !== "") {
        menu.classList.add(c);
      }
    }
  }
  menu.setAttribute("role", "menu");
  menu.hidden = true;

  const applyLabel = (text: string, previewFontFamily?: string): void => {
    labelSpan.textContent = text;
    if (previewFontFamily !== undefined && previewFontFamily !== "") {
      labelSpan.style.fontFamily = previewFontFamily;
    } else {
      labelSpan.style.removeProperty("font-family");
    }
  };

  for (const it of options.items) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fs-dd__item";
    row.setAttribute("role", "menuitem");
    row.dataset.commandId = it.id;
    if (it.previewFontFamily !== undefined && it.previewFontFamily !== "") {
      row.style.fontFamily = it.previewFontFamily;
    }
    if (it.disabled === true) {
      row.disabled = true;
    }
    if (it.icon !== undefined) {
      row.classList.add("fs-dd__item--with-icon");
      const icWrap = document.createElement("span");
      icWrap.className = "fs-dd__item-icon";
      icWrap.setAttribute("aria-hidden", "true");
      icWrap.appendChild(it.icon());
      const lab = document.createElement("span");
      lab.className = "fs-dd__item-label";
      lab.textContent = it.label;
      row.appendChild(icWrap);
      row.appendChild(lab);
    } else {
      row.textContent = it.label;
    }
    row.addEventListener("click", (ev) => {
      ev.stopPropagation();
      if (row.disabled) {
        return;
      }
      applyLabel(it.label, it.previewFontFamily);
      emit(it.id, options.tab);
      menu.hidden = true;
      wrap.classList.remove("fs-dd--open");
      clearToolbarDropdownMenuPosition(menu);
    });
    menu.appendChild(row);
  }

  btn.addEventListener("click", (ev) => {
    ev.stopPropagation();
    const open = menu.hidden;
    closeAllRibbonPopups();
    if (open) {
      menu.hidden = false;
      wrap.classList.add("fs-dd--open");
      syncToolbarDropdownMenuPosition(btn, menu);
    }
  });

  wrap.appendChild(btn);
  wrap.appendChild(menu);

  const close = (): void => {
    menu.hidden = true;
    wrap.classList.remove("fs-dd--open");
    clearToolbarDropdownMenuPosition(menu);
  };

  const setLabel = (text: string, previewFontFamily?: string): void => {
    applyLabel(text, previewFontFamily);
  };

  return { element: wrap, close, setLabel };
}

/** 关闭所有 `.fs-dd` 下拉与带 `data-fs-floating-menu` 的浮动菜单（如边框面板） */
export function closeAllRibbonPopups(): void {
  document.querySelectorAll(".fs-dd.fs-dd--open").forEach((el) => {
    el.classList.remove("fs-dd--open");
    const m = el.querySelector(".fs-dd__menu");
    if (m instanceof HTMLElement) {
      m.hidden = true;
      clearToolbarDropdownMenuPosition(m);
    }
  });
  document.querySelectorAll("[data-fs-floating-menu]").forEach((el) => {
    if (el instanceof HTMLElement && !el.hidden) {
      el.hidden = true;
      clearToolbarDropdownMenuPosition(el);
      const anchorId = el.dataset.fsMenuAnchorId;
      if (anchorId !== undefined && anchorId !== "") {
        document.getElementById(anchorId)?.setAttribute("aria-expanded", "false");
      }
    }
  });
}

/** 展开时脱离 Ribbon overflow，用视口坐标贴触发器下沿；宽度由 CSS（内容自适应）决定 */
export function syncToolbarDropdownMenuPosition(
  trigger: HTMLButtonElement,
  menu: HTMLElement,
): void {
  const r = trigger.getBoundingClientRect();
  menu.style.position = "fixed";
  menu.style.left = `${r.left}px`;
  menu.style.top = `${r.bottom + 2}px`;
  menu.style.zIndex = "5000";
}

export function clearToolbarDropdownMenuPosition(menu: HTMLElement): void {
  menu.style.removeProperty("position");
  menu.style.removeProperty("left");
  menu.style.removeProperty("top");
  menu.style.removeProperty("min-width");
  menu.style.removeProperty("width");
  menu.style.removeProperty("max-width");
  menu.style.removeProperty("z-index");
}

function repositionOpenToolbarDropdowns(): void {
  document.querySelectorAll(".fs-dd.fs-dd--open").forEach((el) => {
    const m = el.querySelector(".fs-dd__menu");
    const b = el.querySelector(".fs-dd__trigger");
    if (m instanceof HTMLElement && !m.hidden && b instanceof HTMLButtonElement) {
      syncToolbarDropdownMenuPosition(b, m);
    }
  });
  document.querySelectorAll("[data-fs-floating-menu]").forEach((el) => {
    if (!(el instanceof HTMLElement) || el.hidden) {
      return;
    }
    const anchorId = el.dataset.fsMenuAnchorId;
    if (anchorId === undefined || anchorId === "") {
      return;
    }
    const b = document.getElementById(anchorId);
    if (b instanceof HTMLButtonElement) {
      syncToolbarDropdownMenuPosition(b, el);
    }
  });
}

if (typeof window !== "undefined") {
  window.addEventListener("scroll", () => repositionOpenToolbarDropdowns(), true);
  window.addEventListener("resize", () => repositionOpenToolbarDropdowns());
}

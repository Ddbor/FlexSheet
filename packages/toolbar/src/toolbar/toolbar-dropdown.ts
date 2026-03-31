import type { RibbonTabId } from "../ribbon/ribbon-types.js";
import type { RibbonEmit } from "./toolbar-button.js";

export interface DropdownItem {
  readonly id: string;
  readonly label: string;
  readonly disabled?: boolean;
}

export interface ToolbarDropdownOptions {
  readonly id: string;
  readonly tab: RibbonTabId;
  readonly label: string;
  readonly items: readonly DropdownItem[];
  readonly title?: string;
  readonly wide?: boolean;
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
  setLabel(text: string): void;
} {
  const wrap = document.createElement("div");
  wrap.className = options.wide === true ? "fs-dd fs-dd--wide" : "fs-dd";
  wrap.dataset.dropdownId = options.id;

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "fs-dd__trigger";
  btn.title = options.title ?? options.label;
  btn.innerHTML = `<span class="fs-dd__label">${escapeHtml(options.label)}</span><span class="fs-dd__chev" aria-hidden="true">▾</span>`;

  const menu = document.createElement("div");
  menu.className = "fs-dd__menu";
  menu.setAttribute("role", "menu");
  menu.hidden = true;

  for (const it of options.items) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fs-dd__item";
    row.setAttribute("role", "menuitem");
    row.textContent = it.label;
    row.dataset.commandId = it.id;
    if (it.disabled === true) {
      row.disabled = true;
    }
    row.addEventListener("click", (ev) => {
      ev.stopPropagation();
      if (row.disabled) {
        return;
      }
      emit(it.id, options.tab);
      menu.hidden = true;
      wrap.classList.remove("fs-dd--open");
    });
    menu.appendChild(row);
  }

  btn.addEventListener("click", (ev) => {
    ev.stopPropagation();
    const open = menu.hidden;
    closeAllDropdownsInDocument();
    if (open) {
      menu.hidden = false;
      wrap.classList.add("fs-dd--open");
    }
  });

  wrap.appendChild(btn);
  wrap.appendChild(menu);

  const close = (): void => {
    menu.hidden = true;
    wrap.classList.remove("fs-dd--open");
  };

  const setLabel = (text: string): void => {
    const span = btn.querySelector(".fs-dd__label");
    if (span instanceof HTMLElement) {
      span.textContent = text;
    }
  };

  return { element: wrap, close, setLabel };
}

function closeAllDropdownsInDocument(): void {
  document.querySelectorAll(".fs-dd.fs-dd--open").forEach((el) => {
    el.classList.remove("fs-dd--open");
    const m = el.querySelector(".fs-dd__menu");
    if (m instanceof HTMLElement) {
      m.hidden = true;
    }
  });
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

import { PluginBase, type ContextMenuItem, type PluginContext } from "@flexsheet/core";

import type { FlexSheet, FlexSheetSurfaceHit } from "./flex-sheet.js";

/** 右键菜单扩展 scope：内置项与 `UIRegistry.getContextMenuItems(scope)` 合并。 */
export const CONTEXT_MENU_SCOPE = {
  cell: "sheet.context.cell",
  rowHeader: "sheet.context.rowHeader",
  columnHeader: "sheet.context.columnHeader",
  sheetCorner: "sheet.context.sheetCorner",
} as const;

function buildBuiltinItems(
  scope: string,
  hit: FlexSheetSurfaceHit,
  flex: FlexSheet,
  openCellInsertSubmenu: () => void,
): readonly ContextMenuItem[] {
  if (scope === CONTEXT_MENU_SCOPE.rowHeader && hit.kind === "rowHeader") {
    return [
      {
        id: "insert",
        label: "插入",
        order: 0,
        onSelect: () => {
          // 行标题右键插入：在当前行上方插入
          flex.insertRows(hit.row, 1);
        },
      },
    ];
  }
  if (scope === CONTEXT_MENU_SCOPE.columnHeader && hit.kind === "columnHeader") {
    return [
      {
        id: "insert",
        label: "插入",
        order: 0,
        onSelect: () => {
          // 列标题右键插入：在当前列左侧插入
          flex.insertCols(hit.col, 1);
        },
      },
    ];
  }
  if (scope === CONTEXT_MENU_SCOPE.cell && hit.kind === "cell") {
    return [{ id: "insert", label: "插入", order: 0, onSelect: openCellInsertSubmenu }];
  }
  return [{ id: "insert", label: "插入", order: 0, disabled: true }];
}

function mergeContextMenuItems(
  scope: string,
  ctx: PluginContext,
  hit: FlexSheetSurfaceHit,
  flex: FlexSheet,
  openCellInsertSubmenu: () => void,
): ContextMenuItem[] {
  const base = buildBuiltinItems(scope, hit, flex, openCellInsertSubmenu);
  const extra = ctx.ui.getContextMenuItems(scope);
  return [...base, ...extra].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
}

export interface SheetContextMenuPluginOptions {
  readonly canvas: HTMLCanvasElement;
  readonly getFlexSheet: () => FlexSheet;
}

export function useSheetContextMenu(options: SheetContextMenuPluginOptions): SheetContextMenuPlugin {
  return new SheetContextMenuPlugin(options);
}

/**
 * 画布区域自定义右键菜单（与行列头、单元格区分 scope，便于后续分别扩展）。
 */
export class SheetContextMenuPlugin extends PluginBase {
  readonly name = "flexsheet.sheetContextMenu";
  private static styleInjected = false;

  private ctx: PluginContext | null = null;
  private readonly canvas: HTMLCanvasElement;
  private readonly getFlexSheet: () => FlexSheet;
  private menuEl: HTMLDivElement | null = null;
  private readonly onDocPointerDown = (ev: PointerEvent): void => {
    if (this.menuEl !== null && !this.menuEl.contains(ev.target as Node)) {
      this.hideMenu();
    }
  };
  private readonly onDocKeyDown = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      this.hideMenu();
    }
  };

  constructor(options: SheetContextMenuPluginOptions) {
    super();
    this.canvas = options.canvas;
    this.getFlexSheet = options.getFlexSheet;
  }

  override install(ctx: PluginContext): void {
    this.ctx = ctx;
  }

  override activate(): void {
    this.ensureMenuStyles();
    this.canvas.addEventListener("contextmenu", this.onCanvasContextMenu, true);
    document.addEventListener("pointerdown", this.onDocPointerDown, true);
    document.addEventListener("keydown", this.onDocKeyDown, true);
  }

  override deactivate(): void {
    this.canvas.removeEventListener("contextmenu", this.onCanvasContextMenu, true);
    document.removeEventListener("pointerdown", this.onDocPointerDown, true);
    document.removeEventListener("keydown", this.onDocKeyDown, true);
    this.hideMenu();
  }

  override destroy(): void {
    this.ctx = null;
    this.hideMenu();
  }

  private readonly onCanvasContextMenu = (ev: MouseEvent): void => {
    const flex = this.getFlexSheet();
    if (flex.isCellEditing()) {
      return;
    }
    ev.preventDefault();
    ev.stopPropagation();
    const hit = flex.hitTestSurface(ev.clientX, ev.clientY);
    if (hit === null) {
      this.hideMenu();
      return;
    }
    const ctx = this.ctx;
    if (ctx === null) {
      return;
    }
    let scope: string;
    if (hit.kind === "corner") {
      scope = CONTEXT_MENU_SCOPE.sheetCorner;
    } else if (hit.kind === "columnHeader") {
      scope = CONTEXT_MENU_SCOPE.columnHeader;
    } else if (hit.kind === "rowHeader") {
      scope = CONTEXT_MENU_SCOPE.rowHeader;
    } else {
      scope = CONTEXT_MENU_SCOPE.cell;
    }
    const items = mergeContextMenuItems(scope, ctx, hit, flex, () => {
      this.openCellInsertSubmenu(flex, ev.clientX, ev.clientY);
    });
    this.showMenu(ev.clientX, ev.clientY, items);
  };

  private showMenu(clientX: number, clientY: number, items: readonly ContextMenuItem[]): void {
    this.hideMenu();
    if (items.length === 0) {
      return;
    }
    const root = document.createElement("div");
    root.className = "fs-sheet-context-menu";
    root.setAttribute("role", "menu");
    root.style.position = "fixed";
    root.style.zIndex = "10001";
    root.style.minWidth = "140px";
    root.style.padding = "4px 0";
    root.style.margin = "0";
    root.style.border = "1px solid #c8c6c4";
    root.style.borderRadius = "2px";
    root.style.background = "#fff";
    root.style.boxShadow = "0 2px 6px rgba(0, 0, 0, 0.12)";
    root.style.fontSize = "12px";
    root.style.fontFamily = "system-ui, -apple-system, sans-serif";

    for (const item of items) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "fs-sheet-context-menu__item";
      btn.setAttribute("role", "menuitem");
      btn.textContent = item.label;
      btn.disabled = item.disabled === true;
      btn.style.display = "block";
      btn.style.width = "100%";
      btn.style.padding = "6px 12px";
      btn.style.border = "none";
      btn.style.background = "#fff";
      btn.style.textAlign = "left";
      btn.style.cursor = item.disabled === true ? "default" : "pointer";
      btn.style.font = "inherit";
      btn.addEventListener("click", () => {
        if (item.disabled === true) {
          return;
        }
        const clickedMenu = this.menuEl;
        item.onSelect?.();
        if (this.menuEl === clickedMenu) {
          this.hideMenu();
        }
      });
      root.appendChild(btn);
    }

    document.body.appendChild(root);
    this.menuEl = root;

    const pad = 4;
    let left = clientX;
    let top = clientY;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    requestAnimationFrame(() => {
      if (this.menuEl === null) {
        return;
      }
      const rw = this.menuEl.offsetWidth;
      const rh = this.menuEl.offsetHeight;
      if (left + rw + pad > vw) {
        left = Math.max(pad, vw - rw - pad);
      }
      if (top + rh + pad > vh) {
        top = Math.max(pad, vh - rh - pad);
      }
      this.menuEl.style.left = `${left}px`;
      this.menuEl.style.top = `${top}px`;
    });
  }

  private hideMenu(): void {
    if (this.menuEl !== null) {
      this.menuEl.remove();
      this.menuEl = null;
    }
  }

  private openCellInsertSubmenu(flex: FlexSheet, clientX: number, clientY: number): void {
    const { row, col } = flex.selection.getActiveCell();
    const subItems: readonly ContextMenuItem[] = [
      {
        id: "insert.cell.shiftRight",
        label: "活动单元格右移",
        order: 0,
        onSelect: () => {
          flex.insertCellsShiftRight(row, col, 1);
        },
      },
      {
        id: "insert.cell.shiftDown",
        label: "活动单元格下移",
        order: 1,
        onSelect: () => {
          flex.insertCellsShiftDown(row, col, 1);
        },
      },
      {
        id: "insert.row.whole",
        label: "整行",
        order: 2,
        onSelect: () => {
          flex.insertRows(row, 1);
        },
      },
      {
        id: "insert.col.whole",
        label: "整列",
        order: 3,
        onSelect: () => {
          flex.insertCols(col, 1);
        },
      },
    ];
    this.showMenu(clientX + 8, clientY, subItems);
  }

  private ensureMenuStyles(): void {
    if (SheetContextMenuPlugin.styleInjected) {
      return;
    }
    const style = document.createElement("style");
    style.textContent = `
.fs-sheet-context-menu__item:not(:disabled):hover,
.fs-sheet-context-menu__item:not(:disabled):focus-visible {
  background: #e8f5e9 !important;
  outline: none;
}
`;
    document.head.appendChild(style);
    SheetContextMenuPlugin.styleInjected = true;
  }
}

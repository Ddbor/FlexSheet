import {
  PluginBase,
  selectionRangeContains,
  type ContextMenuBuiltinIconId,
  type ContextMenuEntry,
  type ContextMenuItem,
  type ContextMenuSeparator,
  type ContextMenuSubItem,
  type PluginContext,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";
import { iconCopy, iconCut, iconPaste } from "@flexsheet/toolbar";

import type { FlexSheet, FlexSheetSurfaceHit, SelectionCellDeleteMode } from "./flex-sheet.js";
import { ensureFsSheetPromptStyles } from "./fs-dialog-styles.js";

/** 选区为整表宽的连续行块，且命中行落在该块内（行标题拖动多选后右键应保留选区）。 */
function isRowHeaderHitInsideEntireRowBlockSelection(
  sheet: Worksheet,
  range: SelectionRange,
  hitRow: number,
): boolean {
  const lastC = Math.max(0, sheet.colCount - 1);
  return (
    range.startCol === 0 &&
    range.endCol === lastC &&
    hitRow >= range.startRow &&
    hitRow <= range.endRow
  );
}

/** 选区为整表高的连续列块，且命中列落在该块内（列标题拖动多选后右键应保留选区）。 */
function isColumnHeaderHitInsideEntireColBlockSelection(
  sheet: Worksheet,
  range: SelectionRange,
  hitCol: number,
): boolean {
  const lastR = Math.max(0, sheet.rowCount - 1);
  return (
    range.startRow === 0 &&
    range.endRow === lastR &&
    hitCol >= range.startCol &&
    hitCol <= range.endCol
  );
}

function cloneBuiltinMenuIcon(id: ContextMenuBuiltinIconId): SVGSVGElement {
  const el =
    id === "cut" ? iconCut() : id === "copy" ? iconCopy() : iconPaste();
  const node = el.cloneNode(true);
  return node as SVGSVGElement;
}

function isContextMenuSeparator(e: ContextMenuEntry): e is ContextMenuSeparator {
  return "kind" in e && e.kind === "separator";
}

function contextMenuItemHasSubmenu(
  e: ContextMenuEntry,
): e is ContextMenuItem & { readonly submenu: readonly ContextMenuSubItem[] } {
  return (
    !isContextMenuSeparator(e) &&
    Array.isArray((e as ContextMenuItem).submenu) &&
    (e as ContextMenuItem).submenu!.length > 0
  );
}

/** 排序后去掉首尾与连续重复的分割线。 */
function normalizeContextMenuEntries(entries: readonly ContextMenuEntry[]): ContextMenuEntry[] {
  if (entries.length === 0) {
    return [];
  }
  const sorted = [...entries].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  const out: ContextMenuEntry[] = [];
  let prevSep = false;
  for (const e of sorted) {
    if (isContextMenuSeparator(e)) {
      if (out.length === 0) {
        continue;
      }
      if (prevSep) {
        continue;
      }
      prevSep = true;
      out.push(e);
    } else {
      prevSep = false;
      out.push(e);
    }
  }
  while (out.length > 0 && isContextMenuSeparator(out[out.length - 1]!)) {
    out.pop();
  }
  return out;
}

/** 右键菜单扩展 scope：内置项与 `UIRegistry.getContextMenuItems(scope)` 合并。 */
export const CONTEXT_MENU_SCOPE = {
  cell: "sheet.context.cell",
  rowHeader: "sheet.context.rowHeader",
  columnHeader: "sheet.context.columnHeader",
  sheetCorner: "sheet.context.sheetCorner",
} as const;

function buildClipboardGroupEntries(flex: FlexSheet): readonly ContextMenuEntry[] {
  return [
    {
      id: "clipboard.cut",
      label: "剪切",
      icon: "cut",
      order: -30,
      onSelect: () => {
        void flex.clipboardCut();
      },
    },
    {
      id: "clipboard.copy",
      label: "复制",
      icon: "copy",
      order: -20,
      onSelect: () => {
        void flex.clipboardCopy();
      },
    },
    {
      id: "clipboard.paste",
      label: "粘贴",
      icon: "paste",
      order: -10,
      onSelect: () => {
        void flex.clipboardPaste();
      },
    },
    { kind: "separator", id: "sep.afterClipboard", order: -5 },
    {
      id: "selection.clearContents",
      label: "清空内容",
      order: -4,
      onSelect: () => {
        flex.clearSelectionContents();
      },
    },
  ];
}

interface BuiltinMenuActions {
  readonly openCellInsertSubmenu: () => void;
  readonly openCellDeleteDialog: () => void;
  readonly openRowHeightPrompt: () => void;
  readonly openColWidthPrompt: () => void;
  readonly openColumnFilterFromCell?: () => void;
}

function buildCellSortSubmenu(
  flex: FlexSheet,
  hit: { readonly row: number; readonly col: number },
): readonly ContextMenuSubItem[] {
  return [
    {
      id: "cell.sort.asc",
      label: "升序",
      onSelect: () => {
        flex.sortSelectionRowsByKeyColumn(hit.col, { type: "value", direction: "asc" });
      },
    },
    {
      id: "cell.sort.desc",
      label: "降序",
      onSelect: () => {
        flex.sortSelectionRowsByKeyColumn(hit.col, { type: "value", direction: "desc" });
      },
    },
    {
      id: "cell.sort.fillOnTop",
      label: "将所选单元格颜色放在最前面",
      onSelect: () => {
        flex.sortSelectionRowsByKeyColumn(hit.col, {
          type: "fillColorOnTop",
          styleAnchorRow: hit.row,
          styleAnchorCol: hit.col,
        });
      },
    },
    {
      id: "cell.sort.fontOnTop",
      label: "将所选字体颜色放在最前面",
      onSelect: () => {
        flex.sortSelectionRowsByKeyColumn(hit.col, {
          type: "fontColorOnTop",
          styleAnchorRow: hit.row,
          styleAnchorCol: hit.col,
        });
      },
    },
  ];
}

function buildBuiltinItems(
  scope: string,
  hit: FlexSheetSurfaceHit,
  flex: FlexSheet,
  actions: BuiltinMenuActions,
): readonly ContextMenuEntry[] {
  if (scope === CONTEXT_MENU_SCOPE.rowHeader && hit.kind === "rowHeader") {
    const sheet = flex.workbook.getActiveSheet();
    const cannotDeleteRow = sheet === undefined || sheet.rowCount <= 1;
    return [
      ...buildClipboardGroupEntries(flex),
      {
        id: "insert",
        label: "插入",
        order: 0,
        onSelect: () => {
          // 行标题右键插入：在当前行上方插入
          flex.insertRows(hit.row, 1);
        },
      },
      {
        id: "structure.deleteRows",
        label: "删除",
        order: 2,
        disabled: cannotDeleteRow,
        onSelect: () => {
          flex.deleteSelectedRows();
        },
      },
      { kind: "separator", id: "sep.beforeRowHeight", order: 5 },
      {
        id: "rowHeight",
        label: "行高",
        order: 10,
        onSelect: actions.openRowHeightPrompt,
      },
    ];
  }
  if (scope === CONTEXT_MENU_SCOPE.columnHeader && hit.kind === "columnHeader") {
    const sheet = flex.workbook.getActiveSheet();
    const cannotDeleteCol = sheet === undefined || sheet.colCount <= 1;
    return [
      ...buildClipboardGroupEntries(flex),
      {
        id: "insert",
        label: "插入",
        order: 0,
        onSelect: () => {
          // 列标题右键插入：在当前列左侧插入
          flex.insertCols(hit.col, 1);
        },
      },
      {
        id: "structure.deleteCols",
        label: "删除",
        order: 2,
        disabled: cannotDeleteCol,
        onSelect: () => {
          flex.deleteSelectedCols();
        },
      },
      { kind: "separator", id: "sep.beforeColWidth", order: 5 },
      {
        id: "colWidth",
        label: "列宽",
        order: 10,
        onSelect: actions.openColWidthPrompt,
      },
    ];
  }
  if (scope === CONTEXT_MENU_SCOPE.cell && hit.kind === "cell") {
    const cellExtras: ContextMenuEntry[] =
      actions.openColumnFilterFromCell !== undefined
        ? [
            {
              id: "data.filter",
              label: "筛选",
              order: 3,
              onSelect: actions.openColumnFilterFromCell,
            },
          ]
        : [];
    return [
      ...buildClipboardGroupEntries(flex),
      { id: "insert", label: "插入", order: 0, onSelect: actions.openCellInsertSubmenu },
      {
        id: "cell.delete",
        label: "删除",
        order: 2,
        onSelect: actions.openCellDeleteDialog,
      },
      { kind: "separator", id: "sep.beforeSortFilter", order: 2.35 },
      {
        id: "cell.sort",
        label: "排序",
        order: 2.5,
        submenu: buildCellSortSubmenu(flex, hit),
      },
      ...cellExtras,
    ];
  }
  return [{ id: "insert", label: "插入", order: 0, disabled: true }];
}

function mergeContextMenuItems(
  scope: string,
  ctx: PluginContext,
  hit: FlexSheetSurfaceHit,
  flex: FlexSheet,
  actions: BuiltinMenuActions,
): ContextMenuEntry[] {
  const base = buildBuiltinItems(scope, hit, flex, actions);
  const extra = ctx.ui.getContextMenuItems(scope);
  return normalizeContextMenuEntries([...base, ...extra]);
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
  private contextSubmenuFlyoutEl: HTMLDivElement | null = null;
  private contextSubmenuHideTimer: number | undefined;
  private promptRoot: HTMLDivElement | null = null;
  private readonly onDocPointerDown = (ev: PointerEvent): void => {
    if (this.promptRoot !== null && this.promptRoot.contains(ev.target as Node)) {
      return;
    }
    const t = ev.target as Node;
    const inRoot = this.menuEl !== null && this.menuEl.contains(t);
    const inFlyout =
      this.contextSubmenuFlyoutEl !== null && this.contextSubmenuFlyoutEl.contains(t);
    if (this.menuEl !== null && !inRoot && !inFlyout) {
      this.hideMenu();
    }
  };
  private readonly onDocKeyDown = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      if (this.promptRoot !== null) {
        this.closePrompt();
        ev.preventDefault();
        ev.stopPropagation();
        return;
      }
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
    this.closePrompt();
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
    if (hit.kind === "cell") {
      const r = flex.selection.getNormalizedRange();
      if (!selectionRangeContains(r, hit.row, hit.col)) {
        flex.focusCellAt(hit.row, hit.col);
      }
    } else if (hit.kind === "rowHeader") {
      const sheet = flex.workbook.getActiveSheet();
      const r = flex.selection.getNormalizedRange();
      if (
        sheet !== undefined &&
        isRowHeaderHitInsideEntireRowBlockSelection(sheet, r, hit.row)
      ) {
        // 已在行标题多选范围内右键：不收缩为单行
      } else {
        flex.focusEntireRowForContextMenu(hit.row);
      }
    } else if (hit.kind === "columnHeader") {
      const sheet = flex.workbook.getActiveSheet();
      const r = flex.selection.getNormalizedRange();
      if (
        sheet !== undefined &&
        isColumnHeaderHitInsideEntireColBlockSelection(sheet, r, hit.col)
      ) {
        // 已在列标题多选范围内右键：不收缩为单列
      } else {
        flex.focusEntireColumnForContextMenu(hit.col);
      }
    }
    const items = mergeContextMenuItems(scope, ctx, hit, flex, {
      openCellInsertSubmenu: () => {
        this.openCellInsertSubmenu(flex, ev.clientX, ev.clientY);
      },
      openCellDeleteDialog: () => {
        this.openCellDeleteDialog(flex);
      },
      openRowHeightPrompt: () => {
        this.openRowHeightPrompt(flex);
      },
      openColWidthPrompt: () => {
        this.openColWidthPrompt(flex);
      },
      openColumnFilterFromCell:
        hit.kind === "cell"
          ? () => {
              const sh = flex.workbook.getActiveSheet();
              if (sh === undefined) {
                return;
              }
              sh.enableColumnAutoFilterFromSelection(
                hit.row,
                hit.col,
                flex.selection.getNormalizedRange(),
              );
              flex.refresh();
            }
          : undefined,
    });
    this.showMenu(ev.clientX, ev.clientY, items);
  };

  private showMenu(clientX: number, clientY: number, items: readonly ContextMenuEntry[]): void {
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

    const pad = 4;
    const vw = window.innerWidth;
    const vh = window.innerHeight;

    const openSubmenuFlyout = (hostBtn: HTMLButtonElement, subs: readonly ContextMenuSubItem[]): void => {
      this.cancelContextSubmenuHide();
      if (this.contextSubmenuFlyoutEl !== null) {
        this.contextSubmenuFlyoutEl.remove();
        this.contextSubmenuFlyoutEl = null;
      }
      const fly = document.createElement("div");
      fly.className = "fs-sheet-context-menu fs-sheet-context-menu--flyout";
      fly.setAttribute("role", "menu");
      for (const sub of subs) {
        const subBtn = document.createElement("button");
        subBtn.type = "button";
        subBtn.className = "fs-sheet-context-menu__item";
        subBtn.setAttribute("role", "menuitem");
        subBtn.disabled = sub.disabled === true;
        subBtn.style.display = "flex";
        subBtn.style.alignItems = "center";
        subBtn.style.gap = "8px";
        subBtn.style.width = "100%";
        subBtn.style.padding = "6px 12px";
        subBtn.style.border = "none";
        subBtn.style.background = "#fff";
        subBtn.style.textAlign = "left";
        subBtn.style.cursor = sub.disabled === true ? "default" : "pointer";
        subBtn.style.font = "inherit";
        const subLabel = document.createElement("span");
        subLabel.className = "fs-sheet-context-menu__label";
        subLabel.textContent = sub.label;
        subBtn.appendChild(subLabel);
        if (sub.shortcutHint !== undefined) {
          const sc = document.createElement("span");
          sc.className = "fs-sheet-context-menu__shortcut";
          sc.textContent = sub.shortcutHint;
          subBtn.appendChild(sc);
        }
        subBtn.addEventListener("pointerdown", (e) => {
          e.stopPropagation();
        });
        subBtn.addEventListener("click", (e) => {
          e.stopPropagation();
          if (sub.disabled === true) {
            return;
          }
          sub.onSelect?.();
          this.hideMenu();
        });
        fly.appendChild(subBtn);
      }
      document.body.appendChild(fly);
      this.contextSubmenuFlyoutEl = fly;
      const br = hostBtn.getBoundingClientRect();
      let fl = br.right - 6;
      let ft = br.top;
      fly.style.left = `${fl}px`;
      fly.style.top = `${ft}px`;
      requestAnimationFrame(() => {
        if (this.contextSubmenuFlyoutEl !== fly) {
          return;
        }
        const fw = fly.offsetWidth;
        const fh = fly.offsetHeight;
        if (fl + fw + pad > vw) {
          fl = Math.max(pad, br.left - fw + 6);
          fly.style.left = `${fl}px`;
        }
        if (ft + fh + pad > vh) {
          ft = Math.max(pad, vh - fh - pad);
          fly.style.top = `${ft}px`;
        }
      });
      fly.addEventListener("mouseenter", () => {
        this.cancelContextSubmenuHide();
      });
      fly.addEventListener("mouseleave", () => {
        this.scheduleContextSubmenuHide();
      });
    };

    for (const item of items) {
      if (isContextMenuSeparator(item)) {
        const sep = document.createElement("div");
        sep.className = "fs-sheet-context-menu__sep";
        sep.setAttribute("role", "separator");
        sep.setAttribute("aria-orientation", "horizontal");
        root.appendChild(sep);
        continue;
      }
      if (contextMenuItemHasSubmenu(item)) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "fs-sheet-context-menu__item fs-sheet-context-menu__item--submenu";
        btn.setAttribute("role", "menuitem");
        btn.setAttribute("aria-haspopup", "menu");
        btn.setAttribute("aria-expanded", "false");
        btn.disabled = item.disabled === true;
        btn.style.display = "flex";
        btn.style.alignItems = "center";
        btn.style.gap = "8px";
        btn.style.width = "100%";
        btn.style.padding = "6px 12px";
        btn.style.border = "none";
        btn.style.background = "#fff";
        btn.style.textAlign = "left";
        btn.style.cursor = item.disabled === true ? "default" : "pointer";
        btn.style.font = "inherit";
        const iconWrap = document.createElement("span");
        iconWrap.className = "fs-sheet-context-menu__icon";
        iconWrap.setAttribute("aria-hidden", "true");
        if (item.icon !== undefined) {
          const svg = cloneBuiltinMenuIcon(item.icon);
          svg.style.width = "16px";
          svg.style.height = "16px";
          svg.style.flexShrink = "0";
          svg.style.display = "block";
          iconWrap.appendChild(svg);
        }
        btn.appendChild(iconWrap);
        const label = document.createElement("span");
        label.className = "fs-sheet-context-menu__label";
        label.textContent = item.label;
        btn.appendChild(label);
        const chevron = document.createElement("span");
        chevron.className = "fs-sheet-context-menu__submenu-chevron";
        chevron.textContent = "▸";
        chevron.setAttribute("aria-hidden", "true");
        btn.appendChild(chevron);
        btn.addEventListener("mouseenter", () => {
          if (btn.disabled) {
            return;
          }
          openSubmenuFlyout(btn, item.submenu);
          btn.setAttribute("aria-expanded", "true");
        });
        btn.addEventListener("mouseleave", () => {
          this.scheduleContextSubmenuHide();
        });
        btn.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          if (btn.disabled) {
            return;
          }
          openSubmenuFlyout(btn, item.submenu);
          btn.setAttribute("aria-expanded", "true");
        });
        root.appendChild(btn);
        continue;
      }
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "fs-sheet-context-menu__item";
      btn.setAttribute("role", "menuitem");
      btn.disabled = item.disabled === true;
      btn.style.display = "flex";
      btn.style.alignItems = "center";
      btn.style.gap = "8px";
      btn.style.width = "100%";
      btn.style.padding = "6px 12px";
      btn.style.border = "none";
      btn.style.background = "#fff";
      btn.style.textAlign = "left";
      btn.style.cursor = item.disabled === true ? "default" : "pointer";
      btn.style.font = "inherit";
      const iconWrap = document.createElement("span");
      iconWrap.className = "fs-sheet-context-menu__icon";
      iconWrap.setAttribute("aria-hidden", "true");
      if (item.icon !== undefined) {
        const svg = cloneBuiltinMenuIcon(item.icon);
        svg.style.width = "16px";
        svg.style.height = "16px";
        svg.style.flexShrink = "0";
        svg.style.display = "block";
        iconWrap.appendChild(svg);
      }
      btn.appendChild(iconWrap);
      const label = document.createElement("span");
      label.className = "fs-sheet-context-menu__label";
      label.textContent = item.label;
      btn.appendChild(label);
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

    let left = clientX;
    let top = clientY;
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
    this.clearContextSubmenuFlyout();
    if (this.menuEl !== null) {
      this.menuEl.remove();
      this.menuEl = null;
    }
  }

  private clearContextSubmenuFlyout(): void {
    if (this.contextSubmenuHideTimer !== undefined) {
      window.clearTimeout(this.contextSubmenuHideTimer);
      this.contextSubmenuHideTimer = undefined;
    }
    if (this.contextSubmenuFlyoutEl !== null) {
      this.contextSubmenuFlyoutEl.remove();
      this.contextSubmenuFlyoutEl = null;
    }
  }

  private scheduleContextSubmenuHide(): void {
    if (this.contextSubmenuHideTimer !== undefined) {
      window.clearTimeout(this.contextSubmenuHideTimer);
    }
    this.contextSubmenuHideTimer = window.setTimeout(() => {
      this.contextSubmenuHideTimer = undefined;
      if (this.contextSubmenuFlyoutEl !== null) {
        this.contextSubmenuFlyoutEl.remove();
        this.contextSubmenuFlyoutEl = null;
      }
    }, 220);
  }

  private cancelContextSubmenuHide(): void {
    if (this.contextSubmenuHideTimer !== undefined) {
      window.clearTimeout(this.contextSubmenuHideTimer);
      this.contextSubmenuHideTimer = undefined;
    }
  }

  private openRowHeightPrompt(flex: FlexSheet): void {
    const initial = String(flex.getRowHeightSampleForSelection());
    this.openNumericPrompt({
      title: "行高",
      fieldLabel: "行高:",
      initialValue: initial,
      onConfirm: (value) => {
        flex.applyRowHeightToSelection(value);
      },
    });
  }

  private openColWidthPrompt(flex: FlexSheet): void {
    const initial = String(flex.getColWidthSampleForSelection());
    this.openNumericPrompt({
      title: "列宽",
      fieldLabel: "列宽:",
      initialValue: initial,
      onConfirm: (value) => {
        flex.applyColWidthToSelection(value);
      },
    });
  }

  private openNumericPrompt(options: {
    readonly title: string;
    readonly fieldLabel: string;
    readonly initialValue: string;
    readonly onConfirm: (value: number) => void;
  }): void {
    this.closePrompt();
    this.ensureMenuStyles();

    const overlay = document.createElement("div");
    overlay.className = "fs-sheet-prompt-overlay";
    overlay.setAttribute("role", "presentation");

    const panel = document.createElement("div");
    panel.className = "fs-sheet-prompt";
    panel.setAttribute("role", "dialog");
    panel.setAttribute("aria-modal", "true");
    panel.setAttribute("aria-labelledby", "fs-sheet-prompt-title");

    const header = document.createElement("div");
    header.className = "fs-sheet-prompt__header";
    const titleEl = document.createElement("div");
    titleEl.id = "fs-sheet-prompt-title";
    titleEl.className = "fs-sheet-prompt__title";
    titleEl.textContent = options.title;
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = "fs-sheet-prompt__close";
    closeBtn.setAttribute("aria-label", "关闭");
    closeBtn.textContent = "×";
    header.appendChild(titleEl);
    header.appendChild(closeBtn);

    const body = document.createElement("div");
    body.className = "fs-sheet-prompt__body";
    const label = document.createElement("label");
    label.className = "fs-sheet-prompt__label";
    const labelText = document.createElement("span");
    labelText.textContent = options.fieldLabel;
    const input = document.createElement("input");
    input.type = "text";
    input.className = "fs-sheet-prompt__input";
    input.value = options.initialValue;
    input.setAttribute("inputmode", "decimal");
    label.appendChild(labelText);
    label.appendChild(input);
    body.appendChild(label);

    const footer = document.createElement("div");
    footer.className = "fs-sheet-prompt__footer";
    const okBtn = document.createElement("button");
    okBtn.type = "button";
    okBtn.className = "fs-sheet-prompt__btn fs-sheet-prompt__btn--primary";
    okBtn.textContent = "确定";
    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "fs-sheet-prompt__btn fs-sheet-prompt__btn--secondary";
    cancelBtn.textContent = "取消";
    footer.appendChild(okBtn);
    footer.appendChild(cancelBtn);

    panel.appendChild(header);
    panel.appendChild(body);
    panel.appendChild(footer);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);
    this.promptRoot = overlay;

    const tryConfirm = (): void => {
      const raw = input.value.trim().replace(/,/g, "");
      const n = Number(raw);
      if (!Number.isFinite(n) || raw === "") {
        input.focus();
        input.select();
        return;
      }
      options.onConfirm(n);
      this.closePrompt();
    };

    const onOverlayPointerDown = (ev: PointerEvent): void => {
      if (ev.target === overlay) {
        this.closePrompt();
      }
    };
    overlay.addEventListener("pointerdown", onOverlayPointerDown);

    closeBtn.addEventListener("click", () => {
      this.closePrompt();
    });
    cancelBtn.addEventListener("click", () => {
      this.closePrompt();
    });
    okBtn.addEventListener("click", () => {
      tryConfirm();
    });
    input.addEventListener("keydown", (kev) => {
      if (kev.key === "Enter") {
        kev.preventDefault();
        tryConfirm();
      }
    });

    requestAnimationFrame(() => {
      input.focus();
      input.select();
    });
  }

  private closePrompt(): void {
    if (this.promptRoot !== null) {
      this.promptRoot.remove();
      this.promptRoot = null;
    }
  }

  private openCellInsertSubmenu(flex: FlexSheet, clientX: number, clientY: number): void {
    const { row, col } = flex.selection.getActiveCell();
    const subItems: readonly ContextMenuEntry[] = [
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

  private openCellDeleteDialog(flex: FlexSheet): void {
    this.closePrompt();
    this.ensureMenuStyles();

    const sheet = flex.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const cannotDeleteEntireRow = sheet.rowCount <= 1;
    const cannotDeleteEntireCol = sheet.colCount <= 1;

    const overlay = document.createElement("div");
    overlay.className = "fs-sheet-prompt-overlay";
    overlay.setAttribute("role", "presentation");

    const panel = document.createElement("div");
    panel.className = "fs-sheet-prompt fs-sheet-prompt--delete";
    panel.setAttribute("role", "dialog");
    panel.setAttribute("aria-modal", "true");
    panel.setAttribute("aria-labelledby", "fs-sheet-delete-dialog-title");

    const header = document.createElement("div");
    header.className = "fs-sheet-prompt__header";
    const titleEl = document.createElement("div");
    titleEl.id = "fs-sheet-delete-dialog-title";
    titleEl.className = "fs-sheet-prompt__title";
    titleEl.textContent = "删除";
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = "fs-sheet-prompt__close";
    closeBtn.setAttribute("aria-label", "关闭");
    closeBtn.textContent = "×";
    header.appendChild(titleEl);
    header.appendChild(closeBtn);

    const body = document.createElement("div");
    body.className = "fs-sheet-prompt__body";
    const fieldset = document.createElement("fieldset");
    fieldset.className = "fs-sheet-prompt__fieldset";
    const legend = document.createElement("legend");
    legend.className = "fs-sheet-prompt__fieldset-legend";
    legend.textContent = "删除";
    fieldset.appendChild(legend);

    const optionDefs: readonly {
      readonly value: SelectionCellDeleteMode;
      readonly label: string;
      readonly disabled: boolean;
    }[] = [
      { value: "shiftLeft", label: "右侧单元格左移", disabled: false },
      { value: "shiftUp", label: "下方单元格上移", disabled: false },
      { value: "entireRow", label: "整行", disabled: cannotDeleteEntireRow },
      { value: "entireCol", label: "整列", disabled: cannotDeleteEntireCol },
    ];

    let selected: SelectionCellDeleteMode = "shiftLeft";
    const radios: HTMLInputElement[] = [];

    for (const def of optionDefs) {
      const row = document.createElement("label");
      row.className = "fs-sheet-prompt__radio-row";
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = "fs-cell-delete-mode";
      radio.value = def.value;
      radio.disabled = def.disabled;
      radios.push(radio);
      const text = document.createElement("span");
      text.textContent = def.label;
      row.appendChild(radio);
      row.appendChild(text);
      fieldset.appendChild(row);
      radio.addEventListener("change", () => {
        if (radio.checked) {
          selected = def.value;
        }
      });
    }

    const firstEnabled = optionDefs.find((d) => !d.disabled);
    if (firstEnabled !== undefined) {
      selected = firstEnabled.value;
    }
    for (let i = 0; i < optionDefs.length; i++) {
      const r = radios[i];
      const d = optionDefs[i];
      if (r !== undefined && d !== undefined) {
        r.checked = !d.disabled && d.value === selected;
      }
    }

    body.appendChild(fieldset);

    const footer = document.createElement("div");
    footer.className = "fs-sheet-prompt__footer";
    const okBtn = document.createElement("button");
    okBtn.type = "button";
    okBtn.className = "fs-sheet-prompt__btn fs-sheet-prompt__btn--primary";
    okBtn.textContent = "确定";
    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "fs-sheet-prompt__btn fs-sheet-prompt__btn--secondary";
    cancelBtn.textContent = "取消";
    footer.appendChild(okBtn);
    footer.appendChild(cancelBtn);

    panel.appendChild(header);
    panel.appendChild(body);
    panel.appendChild(footer);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);
    this.promptRoot = overlay;

    const tryConfirm = (): void => {
      const choice =
        (radios.find((r) => r.checked)?.value as SelectionCellDeleteMode | undefined) ?? selected;
      flex.executeSelectionCellDelete(choice);
      this.closePrompt();
    };

    const onOverlayPointerDown = (ev: PointerEvent): void => {
      if (ev.target === overlay) {
        this.closePrompt();
      }
    };
    overlay.addEventListener("pointerdown", onOverlayPointerDown);

    closeBtn.addEventListener("click", () => {
      this.closePrompt();
    });
    cancelBtn.addEventListener("click", () => {
      this.closePrompt();
    });
    okBtn.addEventListener("click", () => {
      tryConfirm();
    });

    requestAnimationFrame(() => {
      const focusTarget = radios.find((r) => r.checked && !r.disabled) ?? radios.find((r) => !r.disabled);
      focusTarget?.focus();
    });
  }

  private ensureMenuStyles(): void {
    ensureFsSheetPromptStyles();
    if (SheetContextMenuPlugin.styleInjected) {
      return;
    }
    const style = document.createElement("style");
    style.textContent = `
.fs-sheet-context-menu__icon {
  box-sizing: border-box;
  flex-shrink: 0;
  width: 16px;
  min-width: 16px;
  height: 16px;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #323130;
}
.fs-sheet-context-menu__label {
  flex: 1;
  min-width: 0;
  text-align: left;
}
.fs-sheet-context-menu__sep {
  height: 1px;
  margin: 4px 8px;
  background: #c8c6c4;
  pointer-events: none;
}
.fs-sheet-context-menu__item:not(:disabled):hover,
.fs-sheet-context-menu__item:not(:disabled):focus-visible {
  background: #e8f5e9 !important;
  outline: none;
}
.fs-sheet-context-menu__item--submenu {
  justify-content: flex-start;
}
.fs-sheet-context-menu__submenu-chevron {
  margin-left: auto;
  padding-left: 8px;
  flex-shrink: 0;
  opacity: 0.55;
  font-size: 11px;
}
.fs-sheet-context-menu--flyout {
  position: fixed;
  z-index: 10003;
  min-width: 200px;
  padding: 4px 0;
  margin: 0;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  background: #fff;
  box-shadow: 0 2px 6px rgba(0, 0, 0, 0.12);
  font-size: 12px;
  font-family: system-ui, -apple-system, sans-serif;
}
.fs-sheet-context-menu--flyout .fs-sheet-context-menu__shortcut {
  margin-left: auto;
  padding-left: 16px;
  flex-shrink: 0;
  opacity: 0.55;
  font-size: 11px;
  white-space: nowrap;
}
.fs-sheet-prompt--delete {
  width: min(340px, calc(100vw - 32px));
}
.fs-sheet-prompt__fieldset {
  margin: 0;
  padding: 10px 12px 12px 12px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  font-size: 13px;
  color: #323130;
}
.fs-sheet-prompt__fieldset-legend {
  padding: 0 6px;
  font-size: 13px;
  font-weight: 600;
  color: #323130;
}
.fs-sheet-prompt__radio-row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin: 6px 0 0 0;
  cursor: pointer;
  user-select: none;
}
.fs-sheet-prompt__radio-row input[type="radio"] {
  margin: 0;
  flex-shrink: 0;
  accent-color: #217346;
}
.fs-sheet-prompt__radio-row:has(input:disabled) {
  cursor: default;
  color: #a19f9d;
}
`;
    document.head.appendChild(style);
    SheetContextMenuPlugin.styleInjected = true;
  }
}

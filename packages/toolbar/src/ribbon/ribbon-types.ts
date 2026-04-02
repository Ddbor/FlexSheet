/**
 * Ribbon 标签与对外命令事件（宿主可据此驱动 FlexSheet / CanvasRenderer）。
 */

import type { SelectionRange } from "@flexsheet/core";
import type { CanvasRenderer } from "@flexsheet/renderer";
import type { SheetTheme } from "@flexsheet/theme";
export type RibbonTabId = "home" | "insert" | "pageLayout" | "formula" | "data" | "view";

export interface RibbonCommandEvent {
  /** 点状命令 id，如 `home.font.bold`、`view.zoom.100` */
  readonly id: string;
  readonly tab: RibbonTabId;
  readonly payload?: Readonly<Record<string, unknown>>;
}

export interface FlexSheetRibbonOptions {
  /** 挂载容器（通常为 #toolbar） */
  readonly container: HTMLElement;
  /**
   * 「文件」Backstage 全屏覆盖的祖先节点（通常为含编辑栏与表格的 #fs-sheet-chrome）。
   * 未传时使用 `container.parentElement`，需为该节点设置 `position: relative`。
   */
  readonly backstageCoverRoot?: HTMLElement;
  /** 可选：传入后可通过 `syncTheme` 与 `getRenderer()` 联动 */
  readonly flexSheet?: FlexSheetLike;
  /** 按钮、下拉项等交互回调 */
  readonly onCommand?: (ev: RibbonCommandEvent) => void;
  /** 「开始」选项卡挂载完成，用于撤销/重做按钮与命令栈同步 */
  readonly onHomeTabMounted?: (handles: import("./tabs/home-tab.js").HomeTabHandles) => void;
  /** 视图选项卡挂载完成，用于 `ViewRibbonController` 绑定句柄 */
  readonly onViewTabMounted?: (handles: import("./tabs/view-tab.js").ViewTabHandles) => void;
}

/** 与 `FlexSheet` 对齐的最小门面，便于 Ribbon 联动主题与冻结等 */
export interface FlexSheetLike {
  getTheme(): SheetTheme;
  setTheme(theme: SheetTheme): void;
  toggleColorMode(): void;
  getRenderer(): CanvasRenderer;
  refresh(): void;
  setFrozenPanes(frozenRows: number, frozenCols: number): void;
  readonly selection: {
    getNormalizedRange(): SelectionRange;
    getActiveCell(): { readonly row: number; readonly col: number };
  };
  setFormulaBarVisible?(visible: boolean): void;
  isFormulaBarVisible?(): boolean;
  undo(): boolean;
  redo(): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
  subscribeUndoRedo(listener: () => void): () => void;
}

export type { ViewTabHandles } from "./tabs/view-tab.js";

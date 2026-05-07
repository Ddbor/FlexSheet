/**
 * Ribbon 标签与对外命令事件（宿主可据此驱动 FlexSheet / CanvasRenderer）。
 */

import type {
  CellStyle,
  CellStylePatch,
  ConditionalFormatRule,
  SelectionRange,
  Workbook,
} from "@flexsheet/core";
import type { FlexSheetLoadWorkbookOptions, XlsxFloatingPictureExport } from "@flexsheet/import-export";
import type { CanvasRenderer } from "@flexsheet/renderer";
import type { SheetTheme } from "@flexsheet/theme";
export type RibbonTabId =
  | "home"
  | "insert"
  | "pageLayout"
  | "formula"
  | "data"
  | "view"
  /** 浮动图片选中时动态插入的上下文选项卡 */
  | "pictureFormat";

/** 与 flexsheet 浮动图 `FloatingPictureAdjustments` 字段一致（工具栏侧不依赖 flexsheet 包）。 */
export interface FloatingPictureAdjustmentsState {
  brightnessPct: number;
  contrastPct: number;
  sharpnessPct: number;
  saturationPct: number;
  colorTemperatureK: number;
  transparencyPct: number;
}

export interface RibbonCommandEvent {
  /** 点状命令 id，如 `home.font.bold`、`view.zoom.100` */
  readonly id: string;
  readonly tab: RibbonTabId;
  readonly payload?: Readonly<Record<string, unknown>>;
}

export interface RibbonCustomTableStyleEntry {
  /** 自定义样式稳定 id。 */
  readonly id: string;
  /** 在样式库中展示的名称。 */
  readonly name: string;
  /**
   * 套用时要触发的命令 id（当前可复用内置 `home.style.table.*` 命令）。
   * 后续若接入完整自定义样式持久化，可换为专用命令 id。
   */
  readonly commandId: string;
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
  /** 与 Ribbon 剪贴板、快捷键共用实现 */
  clipboardCopy(): Promise<void>;
  clipboardCut(): Promise<void>;
  clipboardPaste(): Promise<void>;
  /** 选区字体/填充等样式（与 Ribbon「开始」字体组共用，可撤销）。 */
  applySelectionStylePatch(patch: CellStylePatch): void;
  applySelectionFontSizeStep(dir: 1 | -1): void;
  applySelectionIndentStep(dir: 1 | -1): void;
  /** Ribbon「开始 -> 填充」：按方向扩展并填充当前选区。 */
  applySelectionFillDirection?(dir: "down" | "right" | "up" | "left"): void;
  /** Ribbon「开始」对齐组：合并/取消合并（与 `home.align.merge*` 命令对应）。 */
  applySelectionMerge?(kind: "mergeCells" | "mergeAcross" | "mergeCenter" | "unmerge"): void;
  /** Ribbon「开始」字体组边框按钮及下拉（`home.font.border*`）。 */
  applyRibbonBorderCommand?(commandId: string): void;
  /** 活动单元格样式（Ribbon 字体/颜色条同步）。 */
  getActiveCellStyle(): CellStyle | null;
  /** 打开「设置单元格格式」对话框（单元格样式库「新建单元格样式」等）。 */
  openFormatCellsDialog?: () => void;
  /** Ribbon「开始 → 查找… / 替换…」：打开查找和替换。 */
  openFindReplaceFromRibbon?(tab: "find" | "replace"): void;
  /** Ribbon「开始 → 查找 → 定位条件…」：打开定位条件对话框。 */
  openGotoSpecialDialogFromRibbon?(): void;
  /** 存在时 Backstage 可提供 JSON 保存/导入。 */
  readonly workbook?: Workbook;
  loadWorkbook?(wb: Workbook, options?: FlexSheetLoadWorkbookOptions): void;
  /** 数据选项卡：按活动列对选区行排序。 */
  sortSelectionRowsByKeyColumn?(
    sortCol: number,
    kind:
      | { readonly type: "value"; readonly direction: "asc" | "desc" }
      | {
          readonly type: "fontColorOnTop";
          readonly styleAnchorRow: number;
          readonly styleAnchorCol: number;
        }
      | {
          readonly type: "fillColorOnTop";
          readonly styleAnchorRow: number;
          readonly styleAnchorCol: number;
        },
  ): void;
  /** 打开自定义排序对话框（多关键字、标题行等）。 */
  openCustomSortDialog?(): void;
  /** 按选区与活动单元格启用列自动筛选（与右键「筛选」一致）。 */
  enableColumnAutoFilterFromSelection?(): void;
  /** 清除工作表上全部列自动筛选。 */
  clearAllColumnAutoFilters?(): void;
  /** 按当前筛选条件重新计算隐藏行。 */
  reapplyAutoFilterConcealment?(): void;
  /** 条件格式：追加一条规则（Ribbon 对话框确定后调用，可撤销）。 */
  addConditionalFormatRuleFromUi?(rule: ConditionalFormatRule): void;
  /** 条件格式：替换当前表全部规则（管理规则对话框，可撤销）。 */
  replaceConditionalFormatRulesFromUi?(rules: readonly ConditionalFormatRule[]): void;
  /** 条件格式：清除与选区相交的规则（可撤销）。 */
  clearConditionalFormatRulesInSelection?(): void;
  /** 条件格式：清除整张表规则（可撤销）。 */
  clearAllConditionalFormatRulesFromUi?(): void;
  /** Ribbon「套用表格格式」：弹出表数据来源对话框并应用所选样式。 */
  openFormatAsTableFromRibbon?(ribbonCommandId: string): void;
  /** Ribbon「套用表格格式」：打开“新建表样式”对话框。 */
  openNewTableStyleDialog?(): void;
  /** Ribbon「填充 -> 系列」：打开系列填充对话框。 */
  openFillSeriesDialog?(): void;
  /** Ribbon「插入 -> 图片」：打开本地文件选择并插入浮动图片。 */
  openInsertPictureFromRibbon?(): void;
  /** 导出 Excel 时附带浮动图片（与 `viewZoom` 配合写入 DrawingML）。 */
  getFloatingPicturesForXlsxExport?(): readonly XlsxFloatingPictureExport[];
  /** 导出前异步合成大图框留白（优先于 `getFloatingPicturesForXlsxExport`）。 */
  prepareFloatingPicturesForXlsxExport?(): Promise<readonly XlsxFloatingPictureExport[]>;
  /**
   * 浮动插入图片获得/失去选中时通知（用于 Ribbon 动态「图片格式」选项卡）。
   * 订阅时应立即用当前状态调用一次 listener。
   */
  subscribeFloatingPictureFocus?(listener: (active: boolean) => void): () => void;
  /** 当前选中浮动图 dataUrl（用于「更正」缩略预览）；无选中时为 null。 */
  getSelectedFloatingPictureDataUrl?(): string | null;
  getFloatingPictureAdjustmentsState?(): FloatingPictureAdjustmentsState | null;
  setFloatingPictureAdjustmentsState?(patch: Partial<FloatingPictureAdjustmentsState>): void;
  /** 打开 Canvas 右侧「设置图片格式」面板（需已选中浮动图）。 */
  openFloatingPictureFormatPane?(): void;
  /** 将当前浮动图显示调整恢复为默认（亮度、对比度、透明度等）。 */
  resetFloatingPictureFormatting?(): void;
  /** Ribbon「插入 -> 数据透视表」：打开数据透视表创建对话框。 */
  openPivotTableDialog?(): void;
  /** Ribbon「数据 -> 字段列表」：在透视区域内打开字段窗格。 */
  openPivotTableFieldsPane?(): void;
  /** Ribbon「清除」：清除当前选区内容（保留格式）。 */
  clearSelectionContents?(): void;
  /** Ribbon「清除」：清除当前选区格式（保留内容）。 */
  clearSelectionFormats?(): void;
  /** Ribbon「清除」：清除当前选区内容与格式。 */
  clearSelectionAll?(): void;
  /**
   * Ribbon「自动求和」主钮 / 下拉里「求和、平均值…」：在活跃格插入 `=SUM` 等并推测区域、显示参数提示。
   */
  applyAutoSumFromRibbon?(
    aggregate: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN",
  ): void;
  /** Ribbon「插入函数」（`formula.insertFunction`）、「其他函数…」（`formula.fn.more`、自动求和下拉）打开插入函数相关 UI。 */
  openInsertFunctionDialogFromRibbon?(): void;
  /** Ribbon「公式」分类下拉：按文档类列出函数名（与函数目录一致）。`categories` 含 `__other__` 时表示非主分类。 */
  listFormulaNamesForRibbonCategories?(
    categories: readonly string[],
    maxNames?: number,
  ): readonly string[];
  /** Ribbon「公式」：打开公式生成器并按类筛选 / 预选函数（`formula.fn.pick`）。 */
  openFormulaBuilderFromRibbon?(options?: {
    readonly categories?: readonly string[];
    readonly sectionLabel?: string;
    readonly selectFunctionName?: string;
  }): void;
  /** Ribbon「套用表格格式」：读取“自定义”分组样式项。 */
  getCustomTableStyleEntries?(): readonly RibbonCustomTableStyleEntry[];
}

export type { ViewTabHandles } from "./tabs/view-tab.js";

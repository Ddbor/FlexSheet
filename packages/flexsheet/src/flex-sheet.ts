import {
  isUnconfiguredPivotDefinition,
  normalizeSelectionRange,
  PLUGIN_SERVICE_KEYS,
  selectionRangesEqualNormalized,
  Workbook,
  Workspace,
  Worksheet,
  writeUnconfiguredPivotPlaceholderToSheet,
  type CellStyle,
  type CellStylePatch,
  type ConditionalFormatRule,
  type SelectionRange,
} from "@flexsheet/core";
import {
  ClearRegionContentsCommand,
  recalcWorksheet,
  SetCellValueCommand,
} from "@flexsheet/formula";
import {
  CellEditor,
  EditorPlugin,
  cellScalarToEditString,
  parseEditString,
  type BeginEditOptions,
} from "@flexsheet/editor";
import { SelectionModel } from "@flexsheet/selection";
import {
  RendererPlugin,
  scrollToRevealCell,
  hitTestCell,
  hitTestBodyCellAutoFilterButton,
  hitTestColumnHeaderFilterButton,
  hitTestHeadingPointer,
  buildFrozenLayout,
  computeColumnAutoWidth,
  computeRowAutoHeight,
  computeScrollLimits,
  expandSelectionRangeForMergePaint,
  SELECTION_OUTLINE_VISUAL_SCALE,
  type CanvasRenderer,
  type HeadingHit,
} from "@flexsheet/renderer";
import { ScrollPlugin } from "@flexsheet/scroll";
import { SelectionRegistryPlugin } from "@flexsheet/selection";
import { columnIndexToLabel, columnLabelToIndex } from "@flexsheet/shared";
import { createDefaultDarkTheme, createDefaultLightTheme, type SheetTheme } from "@flexsheet/theme";

import { runClipboardCopy, runClipboardCut, runClipboardPaste } from "./clipboard/clipboard-run.js";
import { useClipboard } from "./plugins/clipboard-plugin.js";
import {
  DeleteCellsShiftLeftCommand,
  DeleteCellsShiftUpCommand,
  DeleteColsCommand,
  DeleteRowsCommand,
  InsertColsCommand,
  InsertCellsShiftDownCommand,
  InsertCellsShiftRightCommand,
  InsertRowsCommand,
  SetColHiddenCommand,
  SetColWidthsInRangeCommand,
  SetColWidthCommand,
  SetRowHeightsInRangeCommand,
  SetRowHeightCommand,
  SetRowHiddenCommand,
} from "./commands/sheet-structure-commands.js";
import {
  ApplySelectionBorderRibbonCommand,
  ApplySelectionCellStylePatchCommand,
  ApplySelectionFormatCellsDialogCommand,
  ApplySelectionFontSizeStepCommand,
  ApplySelectionIndentStepCommand,
  ClearSelectionFormatsCommand,
  isRibbonBorderCommandId,
} from "./commands/cell-style-commands.js";
import type { FormatCellsBorderState } from "./format-cells/format-cells-border.js";
import {
  AddConditionalFormatRuleCommand,
  ClearAllConditionalFormatRulesCommand,
  ClearConditionalFormatRulesIntersectingCommand,
  SetConditionalFormatRulesCommand,
} from "./commands/conditional-format-commands.js";
import { SelectionMergeCommand } from "./commands/merge-commands.js";
import { useSheetChromeGuard } from "./plugins/sheet-chrome-guard-plugin.js";
import { openColumnFilterPanel } from "./chrome/column-filter-panel.js";
import { openPivotFilterPanel } from "./pivot/pivot-filter-panel.js";
import { showFormatAsTableDialog } from "./dialogs/format-as-table-dialog.js";
import { ensureFsSheetPromptStyles } from "./dialogs/fs-dialog-styles.js";
import { showFillSeriesDialog } from "./dialogs/fill-series-dialog.js";
import { showNewTableStyleDialog } from "./dialogs/new-table-style-dialog.js";
import { parseFormatAsTableRangeRef } from "./dialogs/format-as-table-range.js";
import { showPivotTableDialog } from "./pivot/pivot-table-dialog.js";
import {
  syncPivotTableFieldsPaneWithSelection,
  tryOpenPivotFieldsPaneForSelection,
} from "./pivot/pivot-table-fields-pane.js";
import { refreshPivotTableDefinition } from "./pivot/pivot-table-command.js";
import { useSheetContextMenu } from "./plugins/sheet-context-menu-plugin.js";
import { mountFormatCellsDialog } from "./format-cells/format-cells-dialog.js";
import { useUndoRedo } from "./plugins/undo-redo-plugin.js";
import { AutofillExtendCommand } from "./commands/autofill-extend-command.js";
import { computeAutoSumRange } from "./util/compute-auto-sum-range.js";

/** 指针命中画布表面时的区域类型（供右键菜单等扩展使用）。 */
export type FlexSheetSurfaceHit =
  | { readonly kind: "corner" }
  | { readonly kind: "columnHeader"; readonly col: number }
  | { readonly kind: "rowHeader"; readonly row: number }
  | { readonly kind: "cell"; readonly row: number; readonly col: number };

/** 单元格右键「删除」弹窗中的删除方式。 */
export type SelectionCellDeleteMode = "shiftLeft" | "shiftUp" | "entireRow" | "entireCol";

export interface FlexSheetOptions {
  /** 挂载容器，将插入全尺寸 canvas。 */
  readonly container: HTMLElement;
  /** 可选自定义工作簿；默认带一张示例表。 */
  readonly workbook?: Workbook;
  /** 可选主题；默认浅色 SpreadJS 风格。 */
  readonly theme?: SheetTheme;
  /** 冻结行数（从顶部起）。 */
  readonly frozenRows?: number;
  /** 冻结列数（从左侧起）。 */
  readonly frozenCols?: number;
  /** 可选：与「视图」选项卡联动的编辑栏容器（显示/隐藏）。 */
  readonly formulaBar?: HTMLElement;
  /**
   * 可选：Ribbon + 编辑栏 + 表格 + 底部栏等共同祖先。
   * 传入后在此区域内统一拦截浏览器默认右键与部分快捷键，并在其上监听撤销/剪贴板快捷键。
   */
  readonly chromeRoot?: HTMLElement;
}

/** 距表体边缘（CSS 像素）进入自动滚动的条带宽度。 */
const DRAG_AUTOSCROLL_MARGIN_PX = 36;

/** 指针贴在边缘时的最大滚动速度（文档像素/秒，与 scrollX/scrollY 同单位）。 */
const DRAG_AUTOSCROLL_MAX_SPEED = 880;

/**
 * 对外门面：基于 Workspace + 插件挂载 Canvas，持有 Workbook 与 CanvasRenderer。
 */
export class FlexSheet {
  private _workbook: Workbook;

  /** 插件根容器（命令 / 事件 / UI 扩展点）。 */
  readonly workspace: Workspace;

  get workbook(): Workbook {
    return this._workbook;
  }

  readonly selection: SelectionModel;
  private readonly host: HTMLElement;
  private readonly canvas: HTMLCanvasElement;
  private readonly renderer: CanvasRenderer;
  private readonly cellEditor: CellEditor;
  private theme: SheetTheme;
  private readonly formulaBarEl: HTMLElement | null;
  private readonly formulaBarNameEl: HTMLElement | null;
  private readonly formulaBarInputEl: HTMLTextAreaElement | null;
  private formulaBarSkipBlurCommit = false;
  private formulaBarVisible = true;
  private resizeObserver: ResizeObserver | null = null;
  private dragSelecting = false;
  private resizing = false;
  private resizingKind: "row" | "col" | null = null;
  private resizingIndex = -1;
  private resizingStartClientX = 0;
  private resizingStartClientY = 0;
  private resizingOriginSize = 0;
  private resizingCurrentSize = 0;
  private hoverResizeKind: "row" | "col" | null = null;
  private resizePreviewEl: HTMLDivElement | null = null;
  private dragPointerId: number | null = null;
  /** 框选拖拽时用于边沿自动滚动的上一帧 client 坐标。 */
  private lastDragClientX = 0;
  private lastDragClientY = 0;
  private dragAutoscrollRafId: number | null = null;
  private dragAutoscrollPrevNow = 0;
  /** 在列标题或行标题上按下后拖拽，扩展多列/多行选区（与表体框选互斥）。 */
  private headingDrag:
    | { readonly kind: "column"; readonly originCol: number }
    | { readonly kind: "row"; readonly originRow: number }
    | null = null;
  /** 填充柄拖拽：源矩形（含合并外扩）与当前预览角格。 */
  private fillDrag: {
    readonly sourceRange: SelectionRange;
    previewRow: number;
    previewCol: number;
  } | null = null;
  private workbookUnsub: (() => void) | null = null;
  /** 防止透视自动刷新回写时再次触发递归刷新。 */
  private pivotAutoRefreshRunning = false;
  /** 透视定义键 -> 上次已同步的数据源修订号。 */
  private readonly pivotSourceRevisionByDefKey = new Map<string, number>();
  private lastWorkbookActiveIndex = 0;
  private activeSheetFormattingUnsub: (() => void) | null = null;
  private readonly formattingChromeListeners = new Set<() => void>();
  /** `loadWorkbook` 替换工作簿实例后通知（Chrome 等需重新 `Workbook.subscribe`）。 */
  private readonly workbookReplacedListeners = new Set<() => void>();
  /** 复制/剪切后的走马灯虚线框范围（与当前选区独立）。 */
  private clipboardMarqueeRange: SelectionRange | null = null;
  /** 自动求和等公式编辑中，在参照区域上绘制的虚线预览（与 `SelectionPaintSnapshot.formulaReferencePreviewRange` 一致）。 */
  private formulaRefPreviewRange: SelectionRange | null = null;
  /**
   * 在仅插入 `=函数()` 时显示参数提示；插入完整区域引用后不显示。
   * 与 `cellEditor` 的 `onEditTextChange` 联动。
   */
  private functionHintMode: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN" | null = null;
  private functionHintEl: HTMLDivElement | null = null;
  /**
   * 内联 `=函数()` 在表体上点选/拖拽时，用锚点+对角格构成引用并写回编辑框（与 Excel 选区线一致，不结束编辑）。
   */
  private inlineFormulaRefDrag: {
    readonly fn: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN";
    readonly editRow: number;
    readonly editCol: number;
    readonly anchorR: number;
    readonly anchorC: number;
    focusR: number;
    focusC: number;
  } | null = null;
  /**
   * 自动求和/空参或单区引用时的「表体选区」模式：在 Enter、Esc 或结束编辑前可反复点/拖重选，不因第二次按下而退格为普通点格。
   */
  private formulaArgRangeSession = false;
  /** 延迟剪切：已写入剪贴板但源格尚未清空，粘贴匹配内部载荷后再清源区（与 Excel 一致）。 */
  private pendingClipboardCut: { sheet: Worksheet; range: SelectionRange } | null = null;
  /** 「自定义排序」对话框根节点（打开时独占，关闭时移除）。 */
  private customSortOverlay: HTMLDivElement | null = null;
  /**
   * 对话框「从工作表选定区域」：框选结束写入引用；ESC 取消并恢复进入前的选区。
   */
  private rangeReferencePick:
    | {
        readonly resolve: (value: string | null) => void;
        readonly mode: "range" | "singleCell";
        readonly savedSelection: SelectionRange;
        readonly escHandler: (ev: KeyboardEvent) => void;
        readonly onRangePreview?: (displayRef: string) => void;
      }
    | null = null;
  /** 卸载 `chromeRoot` 上 ⇧⌘R 捕获监听。 */
  private chromeRootSortShortcutCleanup: (() => void) | null = null;
  /** Ribbon「套用表格格式 -> 自定义」样式条目（当前以内置样式命令作为应用后端）。 */
  private readonly customTableStyles: Array<{
    readonly id: string;
    name: string;
    commandId: string;
  }> = [];

  constructor(options: FlexSheetOptions) {
    this.formulaBarEl = options.formulaBar ?? null;
    this.formulaBarNameEl = this.formulaBarEl?.querySelector("#formula-name-box") ?? null;
    this.formulaBarInputEl =
      (this.formulaBarEl?.querySelector("#formula-input") as HTMLTextAreaElement | null) ?? null;
    this.host = options.container;
    this._workbook = options.workbook ?? createDefaultWorkbook();
    this.theme = options.theme ?? createDefaultLightTheme();

    this.selection = new SelectionModel(() => this.workbook.getActiveSheet());
    this.workspace = new Workspace(this._workbook);
    this.workspace.use(new SelectionRegistryPlugin(this.selection));

    const rendererPlugin = new RendererPlugin({
      container: this.host,
      workbook: this._workbook,
      theme: this.theme,
      frozenRows: options.frozenRows ?? 0,
      frozenCols: options.frozenCols ?? 0,
      getSelectionSnapshot: () => {
        const sheet = this.workbook.getActiveSheet();
        if (sheet === undefined) {
          return null;
        }
        const cell = this.selection.getActiveCell();
        const base = {
          range: this.selection.getNormalizedRange(),
          activeRow: cell.row,
          activeCol: cell.col,
        };
        if (this.fillDrag !== null) {
          return {
            ...base,
            fillPreviewRange: this.computeFillPreviewRange(
              this.fillDrag.sourceRange,
              this.fillDrag.previewRow,
              this.fillDrag.previewCol,
            ),
            formulaReferencePreviewRange: this.formulaRefPreviewRange,
          };
        }
        return {
          ...base,
          formulaReferencePreviewRange: this.formulaRefPreviewRange,
        };
      },
      getClipboardMarqueeRange: () => this.clipboardMarqueeRange,
    });
    this.workspace.use(rendererPlugin);
    this.renderer = rendererPlugin.getRenderer();
    this.canvas = rendererPlugin.getCanvas();

    const editorPlugin = new EditorPlugin({
      host: this.host,
      getCanvas: () => this.canvas,
      getTheme: () => this.theme,
      getCellRect: (row, col) => this.renderer.getCellRectInCanvasPixels(row, col),
      getCellFontCss: (row, col) => this.renderer.getCellEditorFontCss(row, col),
      onCommit: (row, col, value) => {
        const sheet = this.workbook.getActiveSheet();
        if (sheet === undefined) {
          return;
        }
        const cmd = new SetCellValueCommand(sheet, row, col, value);
        this.workspace.commands.execute(cmd);
        this.autoExpandRowHeightForMultilineValue(sheet, row, col, value);
      },
      onEditTextChange: (text: string) => {
        if (this.formulaBarInputEl !== null && document.activeElement !== this.formulaBarInputEl) {
          this.formulaBarInputEl.value = text;
        }
        this.updateAutoSumFunctionHintFromEditText(text);
      },
      onEditEnd: () => {
        this.clearAutoSumFormulaPreviewUi();
        this.syncFormulaBar();
        queueMicrotask(() => {
          this.canvas.focus();
        });
      },
    });
    this.workspace.use(editorPlugin);
    this.cellEditor = editorPlugin.getCellEditor();

    const keyRoot = options.chromeRoot ?? this.canvas;

    this.workspace.use(useUndoRedo({ canvas: this.canvas, keyTarget: keyRoot }));

    this.workspace.use(
      useClipboard({
        canvas: this.canvas,
        keyTarget: keyRoot,
        getFlexSheet: () => this,
      }),
    );

    this.workspace.use(new ScrollPlugin());

    this.workspace.use(
      useSheetContextMenu({
        canvas: this.canvas,
        getFlexSheet: () => this,
      }),
    );

    if (options.chromeRoot !== undefined) {
      const chromeRoot = options.chromeRoot;
      this.workspace.use(useSheetChromeGuard({ chromeRoot }));
      const onCustomSortShortcut = (ev: KeyboardEvent): void => {
        if (!ev.shiftKey || !(ev.metaKey || ev.ctrlKey) || (ev.key !== "r" && ev.key !== "R")) {
          return;
        }
        if (this.cellEditor.isEditing()) {
          return;
        }
        const rawT = ev.target;
        if (rawT instanceof HTMLElement) {
          const host = rawT.closest("input, textarea, select, [contenteditable='true']");
          if (host !== null) {
            return;
          }
        }
        ev.preventDefault();
        ev.stopPropagation();
        this.openCustomSortDialog();
      };
      chromeRoot.addEventListener("keydown", onCustomSortShortcut, true);
      this.chromeRootSortShortcutCleanup = () => {
        chromeRoot.removeEventListener("keydown", onCustomSortShortcut, true);
      };
    }

    this.workspace.pluginContext.register(PLUGIN_SERVICE_KEYS.flexSheet, this);

    this.bindResize();
    this.bindSelectionPointer();
    this.bindSelectionKeyboard();
    this.attachWorkbookForDataDrive();
    this.rebindActiveSheetFormattingListener();
    this.bindFormulaBar();
    this.syncSizeAndDraw();
  }

  private attachWorkbookForDataDrive(): void {
    this.workbookUnsub?.();
    this.pivotSourceRevisionByDefKey.clear();
    this.syncPivotAutoRefreshBaseline();
    this.lastWorkbookActiveIndex = this._workbook.activeSheetIndex;
    this.workbookUnsub = this._workbook.subscribe(() => {
      this.tryAutoRefreshPivotTables();
      const cur = this._workbook.activeSheetIndex;
      if (cur !== this.lastWorkbookActiveIndex) {
        this.lastWorkbookActiveIndex = cur;
        this.clipboardMarqueeRange = null;
        this.pendingClipboardCut = null;
        if (this.cellEditor.isEditing()) {
          this.cellEditor.cancelWithoutCommit();
        }
        this.selection.syncWithSheet();
        this.renderer.ensureScrollClamped();
        this.cellEditor.syncLayout();
        this.rebindActiveSheetFormattingListener();
        this.syncPivotTableFieldsPaneToSelection();
      }
      this.renderer.requestRedraw();
    });
  }

  loadWorkbook(wb: Workbook): void {
    this.workbookUnsub?.();
    this._workbook = wb;
    this.clipboardMarqueeRange = null;
    this.pendingClipboardCut = null;
    this.workspace.commands.clear();
    this.attachWorkbookForDataDrive();
    if (wb.sheetCount > 0) {
      wb.activeSheetIndex = Math.min(wb.activeSheetIndex, wb.sheetCount - 1);
    }
    this.renderer.setWorkbook(wb);
    this.lastWorkbookActiveIndex = wb.activeSheetIndex;
    this.renderer.ensureScrollClamped();
    this.selection.syncWithSheet();
    for (let i = 0; i < wb.sheetCount; i++) {
      const sh = wb.getSheet(i);
      if (sh !== undefined) {
        recalcWorksheet(sh);
      }
    }
    for (let i = 0; i < wb.sheetCount; i++) {
      const sh = wb.getSheet(i);
      if (sh === undefined) {
        continue;
      }
      for (const d of sh.getPivotTableDefinitionsSnapshot()) {
        if (isUnconfiguredPivotDefinition(d)) {
          writeUnconfiguredPivotPlaceholderToSheet(sh, d);
        }
      }
    }
    this.rebindActiveSheetFormattingListener();
    this.renderer.requestRedraw();
    this.syncPivotTableFieldsPaneToSelection();
    this.emitWorkbookReplaced();
  }

  getTheme(): SheetTheme {
    return this.theme;
  }

  setTheme(theme: SheetTheme): void {
    this.theme = theme;
    this.renderer.setTheme(theme);
    this.cellEditor.applyTheme(theme);
    this.cellEditor.syncLayout();
    this.renderer.requestRedraw();
  }

  toggleColorMode(): void {
    const next = this.theme.mode === "light" ? createDefaultDarkTheme() : createDefaultLightTheme();
    this.setTheme(next);
  }

  setFrozenPanes(frozenRows: number, frozenCols: number): void {
    this.renderer.setFrozenPanes(frozenRows, frozenCols);
    this.renderer.ensureScrollClamped();
    this.cellEditor.syncLayout();
    this.renderer.requestRedraw();
  }

  getCanvas(): HTMLCanvasElement {
    return this.canvas;
  }

  getRenderer(): CanvasRenderer {
    return this.renderer;
  }

  /** 活动单元格当前样式（无表或越界时为 `null`）。 */
  getActiveCellStyle(): CellStyle | null {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return null;
    }
    const { row, col } = this.selection.getActiveCell();
    const a = sheet.getMergeAnchorCell(row, col);
    return sheet.getCell(a.row, a.col).style;
  }

  /**
   * 选区变化、活动表切换、当前表数据变更时会通知（用于 Ribbon 字体组等与活动格样式同步）。
   * 注册后会立即回调一次。
   */
  subscribeFormattingChrome(listener: () => void): () => void {
    this.formattingChromeListeners.add(listener);
    listener();
    return () => {
      this.formattingChromeListeners.delete(listener);
    };
  }

  /**
   * `loadWorkbook` 会替换 `Workbook` 实例；对旧实例的 `subscribe` 不会随新表生效。
   * 需在回调中重新订阅 `workbook` 并刷新依赖表列表的 UI（如底部标签栏）。
   * 注册后会立即回调一次。
   */
  subscribeWorkbookReplaced(listener: () => void): () => void {
    this.workbookReplacedListeners.add(listener);
    listener();
    return () => {
      this.workbookReplacedListeners.delete(listener);
    };
  }

  private emitWorkbookReplaced(): void {
    for (const fn of this.workbookReplacedListeners) {
      fn();
    }
  }

  /** 是否处于单元格内联编辑态（供右键菜单等扩展判断）。 */
  isCellEditing(): boolean {
    return this.cellEditor.isEditing();
  }

  /**
   * Ribbon「自动求和」主钮及下拉：在活跃单元格插入聚合公式，并尽量推测数据区域、显示参照预览与空参时的函数提示。
   * 与 `home.cells.autoSum`、`formula.autoSum`、`autoSum.sub.*` 一致。
   */
  applyAutoSumFromRibbon(aggregate: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN" = "SUM"): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.clearAutoSumFunctionHint();
    this.formulaRefPreviewRange = null;
    this.formulaArgRangeSession = false;
    this.functionHintMode = null;
    this.renderer.requestRedraw();

    const ac = this.selection.getActiveCell();
    const anchor = sheet.getMergeAnchorCell(ac.row, ac.col);
    const dataRange = computeAutoSumRange(sheet, anchor.row, anchor.col);

    if (dataRange !== null) {
      this.formulaRefPreviewRange = dataRange;
      const refText = this.makeA1RefFromRange(dataRange, sheet);
      const text = `=${aggregate}(${refText})`;
      this.beginEditAt(anchor.row, anchor.col, { initialTextOverride: text, selectAll: false });
    } else {
      const paren = `=${aggregate}()`;
      this.functionHintMode = aggregate;
      this.beginEditAt(anchor.row, anchor.col, {
        initialTextOverride: paren,
        selectAll: false,
        selectionStart: 1 + aggregate.length + 1,
        selectionEnd: 1 + aggregate.length + 1,
      });
    }
    this.revealActiveCellInViewport();
    this.formulaArgRangeSession = true;
    this.updateAutoSumFunctionHintFromEditText(this.cellEditor.getEditingText());
    this.refresh();
  }

  private makeA1RefFromRange(range: SelectionRange, _sheet: Worksheet): string {
    const a = normalizeSelectionRange(range);
    const c0 = columnIndexToLabel(a.startCol);
    const c1 = columnIndexToLabel(a.endCol);
    const r0 = a.startRow + 1;
    const r1 = a.endRow + 1;
    if (a.startRow === a.endRow && a.startCol === a.endCol) {
      return `${c0}${r0}`;
    }
    return `${c0}${r0}:${c1}${r1}`;
  }

  /**
   * 解析为「仅一个 A1/区域 参数（或空）」的聚合，用于表体选区；多参、表名、复杂表达式均返回 `null`。
   */
  private parseSingleArgAggregateRef(
    text: string,
  ): {
    fn: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN";
    isEmpty: boolean;
    range: SelectionRange | null;
  } | null {
    const t = text.trim();
    const m = /^=(SUM|AVERAGE|AVG|COUNT|MAX|MIN)\(([^)]*)\)\s*$/i.exec(t);
    if (m === null) {
      return null;
    }
    const inner = (m[2] ?? "").trim();
    if (inner.includes(",")) {
      return null;
    }
    if (inner !== "" && /[!]/.test(inner)) {
      return null;
    }
    let f = m[1].toUpperCase();
    if (f === "AVG") {
      f = "AVERAGE";
    }
    if (f !== "SUM" && f !== "AVERAGE" && f !== "COUNT" && f !== "MAX" && f !== "MIN") {
      return null;
    }
    const fn = f as "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN";
    if (inner === "") {
      return { fn, isEmpty: true, range: null };
    }
    const r = parseFormatAsTableRangeRef(inner);
    if (r === null) {
      return null;
    }
    return { fn, isEmpty: false, range: r };
  }

  private shouldStartInlineRefDragOnBody(): boolean {
    if (this.rangeReferencePick !== null) {
      return false;
    }
    if (!this.cellEditor.isEditing()) {
      return false;
    }
    const p = this.parseSingleArgAggregateRef(this.cellEditor.getEditingText());
    if (p === null) {
      return false;
    }
    if (p.isEmpty) {
      return true;
    }
    return this.formulaArgRangeSession;
  }

  private beginInlineFormulaRefDrag(
    sheet: Worksheet,
    hit: { row: number; col: number },
    ev: PointerEvent,
  ): void {
    const p = this.parseSingleArgAggregateRef(this.cellEditor.getEditingText());
    const ed = this.cellEditor.getEditingCell();
    if (p === null || ed === null) {
      return;
    }
    if (!p.isEmpty && !this.formulaArgRangeSession) {
      return;
    }
    this.formulaArgRangeSession = true;
    const a0 = sheet.getMergeAnchorCell(hit.row, hit.col);
    this.inlineFormulaRefDrag = {
      fn: p.fn,
      editRow: ed.row,
      editCol: ed.col,
      anchorR: a0.row,
      anchorC: a0.col,
      focusR: a0.row,
      focusC: a0.col,
    };
    this.updateInlineFormulaRefTextAndPreview();
    ev.preventDefault();
    this.lastDragClientX = ev.clientX;
    this.lastDragClientY = ev.clientY;
    this.dragPointerId = ev.pointerId;
    this.attachDocumentDragListeners();
    try {
      this.canvas.setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
    this.applyPointerCursor("crosshair");
    this.renderer.requestRedraw();
    queueMicrotask(() => {
      this.cellEditor.refocusInput();
    });
  }

  private updateInlineFormulaRefTextAndPreview(): void {
    if (this.inlineFormulaRefDrag === null) {
      return;
    }
    const d = this.inlineFormulaRefDrag;
    const n = normalizeSelectionRange({
      startRow: d.anchorR,
      startCol: d.anchorC,
      endRow: d.focusR,
      endCol: d.focusC,
    });
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const ref = this.makeA1RefFromRange(n, sheet);
    const text = `=${d.fn}(${ref})`;
    const tlen = text.length;
    this.formulaRefPreviewRange = n;
    this.functionHintMode = null;
    this.hideAutoSumFunctionHintElementOnly();
    this.cellEditor.setEditingText(text, tlen, tlen);
    this.updateAutoSumFunctionHintFromEditText(this.cellEditor.getEditingText());
    this.cellEditor.syncLayout();
    this.renderer.requestRedraw();
  }

  private hideAutoSumFunctionHintElementOnly(): void {
    if (this.functionHintEl !== null) {
      this.functionHintEl.style.display = "none";
    }
  }

  private clearAutoSumFunctionHint(): void {
    this.functionHintMode = null;
    this.hideAutoSumFunctionHintElementOnly();
  }

  private clearAutoSumFormulaPreviewUi(): void {
    this.clearAutoSumFunctionHint();
    this.formulaRefPreviewRange = null;
    this.formulaArgRangeSession = false;
  }

  private updateAutoSumFunctionHintFromEditText(text: string): void {
    if (this.functionHintMode === null) {
      this.hideAutoSumFunctionHintElementOnly();
      return;
    }
    const fn = this.functionHintMode;
    const re = new RegExp(`^=${fn}\\(\\s*\\)$`, "i");
    if (re.test(text.trimEnd())) {
      this.ensureAndShowFunctionHintFor(fn);
    } else {
      this.hideAutoSumFunctionHintElementOnly();
    }
  }

  private ensureAndShowFunctionHintFor(fn: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN"): void {
    if (this.functionHintEl === null) {
      const el = document.createElement("div");
      el.className = "fs-autosum-fx-hint";
      el.setAttribute("role", "tooltip");
      el.style.cssText = [
        "position:absolute",
        "z-index:30",
        "display:none",
        "min-width:220px",
        "max-width:min(400px,92vw)",
        "padding:8px 10px",
        "background:#fff",
        "color:#201f1e",
        "border:1px solid #c8c6c4",
        "box-shadow:0 2px 6px rgba(0,0,0,.12)",
        "font:12px/1.45 system-ui,-apple-system,sans-serif",
        "pointer-events:none",
      ].join(";");
      this.functionHintEl = el;
      this.host.appendChild(el);
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const { row, col } = this.selection.getActiveCell();
    const a = sheet.getMergeAnchorCell(row, col);
    const rect = this.renderer.getCellRectInCanvasPixels(a.row, a.col);
    if (rect === null) {
      return;
    }
    const { param, desc } = this.getLocalFunctionHelpStrings(fn);
    this.functionHintEl.innerHTML = "";
    const sigLine = document.createElement("div");
    sigLine.style.cssText = "word-break:break-all;margin:0 0 6px";
    this.appendFunctionHintSignature(sigLine, fn, param);
    this.functionHintEl.appendChild(sigLine);
    const sumLabel = document.createElement("div");
    sumLabel.style.cssText = "color:#201f1e;font-weight:600;font-size:11px;margin:0 0 2px";
    sumLabel.textContent = "概要";
    this.functionHintEl.appendChild(sumLabel);
    const body = document.createElement("p");
    body.style.cssText = "margin:0;white-space:pre-wrap;word-break:break-word";
    body.textContent = desc;
    this.functionHintEl.appendChild(body);

    this.functionHintEl.style.display = "block";
    this.functionHintEl.style.left = `${rect.x}px`;
    this.functionHintEl.style.top = `${rect.y + rect.height + 4}px`;
  }

  private getLocalFunctionHelpStrings(fn: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN"): {
    readonly param: string;
    readonly desc: string;
  } {
    if (fn === "SUM") {
      return { param: "number1", desc: "此函数返回某一单元格区域中所有数字之和。" };
    }
    if (fn === "AVERAGE") {
      return { param: "number1", desc: "此函数返回其参数的算术平均值。" };
    }
    if (fn === "COUNT") {
      return { param: "value1", desc: "此函数会统计所给参数中数字的个数。" };
    }
    if (fn === "MAX") {
      return { param: "number1", desc: "此函数返回一组数中的最大数值。" };
    }
    return { param: "number1", desc: "此函数返回一组数中的最小数值。" };
  }

  private appendFunctionHintSignature(
    el: HTMLDivElement,
    fn: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN",
    firstParam: string,
  ): void {
    const head = document.createElement("span");
    head.appendChild(document.createTextNode(`${fn}(`));
    const hi = document.createElement("span");
    hi.textContent = firstParam;
    hi.style.background = "#fff4cc";
    const mid = document.createTextNode(", number2, ...)");
    el.appendChild(head);
    el.appendChild(hi);
    el.appendChild(mid);
  }

  private syncAutoSumFunctionHintLayout(): void {
    if (this.functionHintMode === null || this.functionHintEl === null) {
      return;
    }
    this.updateAutoSumFunctionHintFromEditText(this.cellEditor.getEditingText());
  }

  /**
   * 屏幕坐标命中画布：左上角全选角、列标题、行标题或单元格。
   * 坐标不在画布矩形内时返回 null。
   */
  hitTestSurface(clientX: number, clientY: number): FlexSheetSurfaceHit | null {
    const br = this.canvas.getBoundingClientRect();
    if (clientX < br.left || clientX > br.right || clientY < br.top || clientY > br.bottom) {
      return null;
    }
    const heading = this.hitTestHeadingFromClient(clientX, clientY);
    if (heading !== null) {
      if (heading.kind === "selectAllCorner") {
        return { kind: "corner" };
      }
      if (heading.kind === "columnHeader") {
        return { kind: "columnHeader", col: heading.col };
      }
      return { kind: "rowHeader", row: heading.row };
    }
    const cell = this.hitTestClient(clientX, clientY);
    if (cell === null) {
      return null;
    }
    return { kind: "cell", row: cell.row, col: cell.col };
  }

  refresh(): void {
    this.cellEditor.syncLayout();
    this.syncAutoSumFunctionHintLayout();
    this.renderer.requestRedraw();
  }

  /** 表体 Canvas 所在挂载节点；数据透视字段面板与之并排插入同一父级 flex 行。 */
  getSheetContainerElement(): HTMLElement {
    return this.host;
  }

  /**
   * 将选区设为单格并同步视口与编辑栏（与在表体左键点选该格一致；用于右键菜单等，不进入拖拽选区）。
   */
  focusCellAt(row: number, col: number): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.selection.selectCell(row, col);
    this.afterSelectionChanged();
  }

  /**
   * 右键行标题等：将选区设为整行并同步视口与编辑栏（与左键点行号一致）。
   */
  focusEntireRowForContextMenu(row: number): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.selection.selectEntireRow(row);
    this.afterSelectionChanged();
  }

  /**
   * 右键列标题等：将选区设为整列并同步视口与编辑栏（与左键点列标一致）。
   */
  focusEntireColumnForContextMenu(col: number): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.selection.selectEntireColumn(col);
    this.afterSelectionChanged();
  }

  insertRows(atRow: number, count = 1): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    this.workspace.commands.execute(new InsertRowsCommand(sheet, this.selection, atRow, count));
    this.refresh();
  }

  insertCols(atCol: number, count = 1): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    this.workspace.commands.execute(new InsertColsCommand(sheet, this.selection, atCol, count));
    this.refresh();
  }

  insertCellsShiftRight(row: number, col: number, count = 1): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    this.workspace.commands.execute(
      new InsertCellsShiftRightCommand(sheet, this.selection, row, col, count),
    );
    this.refresh();
  }

  insertCellsShiftDown(row: number, col: number, count = 1): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    this.workspace.commands.execute(
      new InsertCellsShiftDownCommand(sheet, this.selection, row, col, count),
    );
    this.refresh();
  }

  deleteRows(atRow: number, count = 1): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    this.workspace.commands.execute(new DeleteRowsCommand(sheet, this.selection, atRow, count));
    this.refresh();
  }

  deleteCols(atCol: number, count = 1): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    this.workspace.commands.execute(new DeleteColsCommand(sheet, this.selection, atCol, count));
    this.refresh();
  }

  /**
   * 删除当前选区内的所有行（选区来自行标题或整行框选时即为这些行；表至少保留一行时不执行）。
   */
  deleteSelectedRows(): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined || sheet.rowCount <= 1) {
      return;
    }
    const r = this.selection.getNormalizedRange();
    const startRow = r.startRow;
    const count = r.endRow - startRow + 1;
    this.deleteRows(startRow, count);
  }

  /**
   * 删除当前选区内的所有列（选区来自列标题或整列框选时即为这些列；表至少保留一列时不执行）。
   */
  deleteSelectedCols(): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined || sheet.colCount <= 1) {
      return;
    }
    const r = this.selection.getNormalizedRange();
    const startCol = r.startCol;
    const count = r.endCol - startCol + 1;
    this.deleteCols(startCol, count);
  }

  /**
   * 按「删除」弹窗选项处理当前选区：左移/上移填补，或删除整行、整列。
   */
  executeSelectionCellDelete(mode: SelectionCellDeleteMode): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const r = normalizeSelectionRange(this.selection.getNormalizedRange());
    if (mode === "shiftLeft") {
      this.workspace.commands.execute(new DeleteCellsShiftLeftCommand(sheet, this.selection, r));
      this.refresh();
      return;
    }
    if (mode === "shiftUp") {
      this.workspace.commands.execute(new DeleteCellsShiftUpCommand(sheet, this.selection, r));
      this.refresh();
      return;
    }
    if (mode === "entireRow") {
      if (sheet.rowCount <= 1) {
        return;
      }
      this.deleteRows(r.startRow, r.endRow - r.startRow + 1);
      return;
    }
    if (mode === "entireCol") {
      if (sheet.colCount <= 1) {
        return;
      }
      this.deleteCols(r.startCol, r.endCol - r.startCol + 1);
      return;
    }
  }

  setRowHidden(row: number, hidden: boolean): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    this.workspace.commands.execute(new SetRowHiddenCommand(sheet, row, hidden));
    this.refresh();
  }

  setColHidden(col: number, hidden: boolean): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    this.workspace.commands.execute(new SetColHiddenCommand(sheet, col, hidden));
    this.refresh();
  }

  /**
   * 行标题右键「行高」：默认展示值（活动格所在行，文档像素，与拖拽调整一致）。
   */
  getRowHeightSampleForSelection(): number {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return 20;
    }
    const { row } = this.selection.getActiveCell();
    return Math.round(sheet.getRowHeight(row));
  }

  /**
   * 列标题右键「列宽」：默认展示值（活动格所在列，文档像素）。
   */
  getColWidthSampleForSelection(): number {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return 64;
    }
    const { col } = this.selection.getActiveCell();
    return Math.round(sheet.getColWidth(col));
  }

  /** 将当前选区所含各行设为同一高度（与行标题右键一致，可一次撤销）。 */
  applyRowHeightToSelection(heightPx: number): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const h = clampRowColSizePx(heightPx);
    const { startRow, endRow } = this.selection.getNormalizedRange();
    this.workspace.commands.execute(new SetRowHeightsInRangeCommand(sheet, startRow, endRow, h));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** 将当前选区所含各列设为同一宽度（与列标题右键一致，可一次撤销）。 */
  applyColWidthToSelection(widthPx: number): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const w = clampRowColSizePx(widthPx);
    const { startCol, endCol } = this.selection.getNormalizedRange();
    this.workspace.commands.execute(new SetColWidthsInRangeCommand(sheet, startCol, endCol, w));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /**
   * 对当前选区合并单元格样式补丁（可撤销）。
   * `CellStylePatch` 中字段为 `null` 时表示清除该项。
   */
  applySelectionStylePatch(patch: CellStylePatch): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const range = this.selection.getNormalizedRange();
    this.workspace.commands.execute(new ApplySelectionCellStylePatchCommand(sheet, range, patch));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /**
   * 「设置单元格格式」确定：一次合并数字/对齐/字体与可选的边框几何（单条撤销）。
   * `border.apply` 为 false 时不改写边框，仅应用 `basePatch` 中非边框字段。
   */
  applyFormatCellsDialogOk(
    basePatch: CellStylePatch,
    border: { readonly apply: boolean; readonly state: FormatCellsBorderState },
  ): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const range = this.selection.getNormalizedRange();
    this.workspace.commands.execute(
      new ApplySelectionFormatCellsDialogCommand(
        sheet,
        range,
        basePatch,
        border.apply,
        border.state,
      ),
    );
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** 打开「设置单元格格式」对话框（与右键菜单一致；供 Ribbon「新建单元格样式」等使用）。 */
  openFormatCellsDialog(): void {
    mountFormatCellsDialog({ flex: this });
  }

  /** 选区内按字号阶梯增大（`1`）或减小（`-1`）字号，每格相对当前字号，可撤销。 */
  applySelectionFontSizeStep(dir: 1 | -1): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const range = this.selection.getNormalizedRange();
    this.workspace.commands.execute(new ApplySelectionFontSizeStepCommand(sheet, range, dir));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** 选区内逐格增加（`1`）或减少（`-1`）缩进等级，可撤销。 */
  applySelectionIndentStep(dir: 1 | -1): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const range = this.selection.getNormalizedRange();
    this.workspace.commands.execute(new ApplySelectionIndentStepCommand(sheet, range, dir));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** Ribbon「填充」：按方向扩展一个选区跨度并执行平铺填充（可撤销）。 */
  applySelectionFillDirection(dir: "down" | "right" | "up" | "left"): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const source = normalizeSelectionRange(this.selection.getNormalizedRange());
    const rowSpan = source.endRow - source.startRow + 1;
    const colSpan = source.endCol - source.startCol + 1;
    const lastRow = sheet.rowCount - 1;
    const lastCol = sheet.colCount - 1;
    let fill = source;

    if (dir === "down") {
      if (source.endRow >= lastRow) {
        return;
      }
      fill = {
        startRow: source.startRow,
        startCol: source.startCol,
        endRow: Math.min(lastRow, source.endRow + rowSpan),
        endCol: source.endCol,
      };
    } else if (dir === "right") {
      if (source.endCol >= lastCol) {
        return;
      }
      fill = {
        startRow: source.startRow,
        startCol: source.startCol,
        endRow: source.endRow,
        endCol: Math.min(lastCol, source.endCol + colSpan),
      };
    } else if (dir === "up") {
      if (source.startRow <= 0) {
        return;
      }
      fill = {
        startRow: Math.max(0, source.startRow - rowSpan),
        startCol: source.startCol,
        endRow: source.endRow,
        endCol: source.endCol,
      };
    } else {
      if (source.startCol <= 0) {
        return;
      }
      fill = {
        startRow: source.startRow,
        startCol: Math.max(0, source.startCol - colSpan),
        endRow: source.endRow,
        endCol: source.endCol,
      };
    }

    if (selectionRangesEqualNormalized(fill, source)) {
      return;
    }
    this.workspace.commands.execute(new AutofillExtendCommand(sheet, source, fill));
    this.selection.setNormalizedRange(fill);
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /**
   * Ribbon「合并后居中 / 跨越合并 / 合并单元格 / 取消合并」：对当前选区生效，可撤销。
   */
  applySelectionMerge(kind: "mergeCells" | "mergeAcross" | "mergeCenter" | "unmerge"): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const range = this.selection.getNormalizedRange();
    this.workspace.commands.execute(new SelectionMergeCommand(sheet, range, kind));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** Ribbon 边框主按钮 / 下拉项（`home.font.border*`），可撤销。 */
  applyRibbonBorderCommand(commandId: string): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined || !isRibbonBorderCommandId(commandId)) {
      return;
    }
    const range = this.selection.getNormalizedRange();
    this.workspace.commands.execute(new ApplySelectionBorderRibbonCommand(sheet, range, commandId));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** Ribbon「条件格式」：追加一条规则（可撤销）。 */
  addConditionalFormatRuleFromUi(rule: ConditionalFormatRule): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.workspace.commands.execute(new AddConditionalFormatRuleCommand(sheet, rule));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** Ribbon「管理规则」：整体替换当前工作表条件格式列表（可撤销）。 */
  replaceConditionalFormatRulesFromUi(rules: readonly ConditionalFormatRule[]): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.workspace.commands.execute(new SetConditionalFormatRulesCommand(sheet, rules));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** 清除与当前选区相交的条件格式规则（可撤销）。 */
  clearConditionalFormatRulesInSelection(): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const range = this.selection.getNormalizedRange();
    this.workspace.commands.execute(
      new ClearConditionalFormatRulesIntersectingCommand(sheet, range),
    );
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** 清除当前工作表全部条件格式规则（可撤销）。 */
  clearAllConditionalFormatRulesFromUi(): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.workspace.commands.execute(new ClearAllConditionalFormatRulesCommand(sheet));
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** 撤销上一命令（若存在）。 */
  undo(): boolean {
    const ok = this.workspace.commands.undo();
    if (ok) {
      this.cellEditor.syncLayout();
      this.renderer.requestRedraw();
    }
    return ok;
  }

  /** 重做上一撤销项（若存在）。 */
  redo(): boolean {
    const ok = this.workspace.commands.redo();
    if (ok) {
      this.cellEditor.syncLayout();
      this.renderer.requestRedraw();
    }
    return ok;
  }

  canUndo(): boolean {
    return this.workspace.commands.canUndo();
  }

  canRedo(): boolean {
    return this.workspace.commands.canRedo();
  }

  /** 撤销栈变化时回调（用于 Ribbon 按钮状态）。 */
  subscribeUndoRedo(listener: () => void): () => void {
    return this.workspace.commands.subscribe(listener);
  }

  setFormulaBarVisible(visible: boolean): void {
    this.formulaBarVisible = visible;
    if (this.formulaBarEl !== null) {
      this.formulaBarEl.style.display = visible ? "" : "none";
    }
  }

  isFormulaBarVisible(): boolean {
    return this.formulaBarVisible;
  }

  /**
   * 设置复制/剪切后的走马灯边框范围；`null` 关闭。
   * 不改变选区模型，仅叠加绘制。
   */
  setClipboardMarquee(range: SelectionRange | null): void {
    if (range === null) {
      this.clipboardMarqueeRange = null;
      this.pendingClipboardCut = null;
      this.renderer.requestRedraw();
      return;
    }
    this.clipboardMarqueeRange = normalizeSelectionRange(range);
    this.renderer.requestRedraw();
  }

  /**
   * 关闭走马灯并取消延迟剪切（若有）；返回是否此前存在走马灯或待完成剪切（便于 Esc 时 `preventDefault`）。
   */
  clearClipboardMarquee(): boolean {
    const had = this.clipboardMarqueeRange !== null || this.pendingClipboardCut !== null;
    if (!had) {
      return false;
    }
    this.clipboardMarqueeRange = null;
    this.pendingClipboardCut = null;
    this.renderer.requestRedraw();
    return true;
  }

  setPendingClipboardCut(sheet: Worksheet, range: SelectionRange): void {
    this.pendingClipboardCut = { sheet, range: normalizeSelectionRange(range) };
  }

  clearPendingClipboardCut(): void {
    this.pendingClipboardCut = null;
  }

  getPendingClipboardCut(): { sheet: Worksheet; range: SelectionRange } | null {
    return this.pendingClipboardCut;
  }

  /** 复制当前选区（与 Ctrl/Cmd+C 一致）；单元格内联编辑时不执行。 */
  async clipboardCopy(): Promise<void> {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    await runClipboardCopy(this, sheet);
  }

  /**
   * 延迟剪切（与 Excel / Ctrl+X 一致）：写入剪贴板并显示走马灯，源单元格在成功「完成移动」粘贴前仍保留；
   * 复制、Esc、换表、进入编辑等会取消待剪切；内联编辑时不执行。
   */
  async clipboardCut(): Promise<void> {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    await runClipboardCut(this, sheet);
  }

  /** 粘贴到以活动单元格为左上角（与 Ctrl/Cmd+V 一致）；内联编辑时不执行。 */
  async clipboardPaste(): Promise<void> {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    await runClipboardPaste(this, sheet);
  }

  /**
   * 清空当前选区内单元格的值与公式，保留格式（与 Delete / Backspace 一致）；
   * 内联编辑或无活动表时不执行。
   */
  clearSelectionContents(): void {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.clearClipboardMarquee();
    const cmd = new ClearRegionContentsCommand(sheet, this.selection.getNormalizedRange());
    if (cmd.hasChanges) {
      this.workspace.commands.execute(cmd);
    }
    this.renderer.requestRedraw();
  }

  /** 清空当前选区内单元格格式，保留值与公式。 */
  clearSelectionFormats(): void {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.clearClipboardMarquee();
    const cmd = new ClearSelectionFormatsCommand(sheet, this.selection.getNormalizedRange());
    if (cmd.hasChanges) {
      this.workspace.commands.execute(cmd);
    }
    this.cellEditor.syncLayout();
    this.refresh();
  }

  /** 清空当前选区内单元格内容与格式。 */
  clearSelectionAll(): void {
    if (this.isCellEditing()) {
      return;
    }
    this.clearSelectionContents();
    this.clearSelectionFormats();
  }

  /**
   * 按当前选区的行范围排序，排序关键字列为 `sortCol`（通常为右键列或活动列）。
   * 与 Excel 右键「排序」子菜单行为一致。
   */
  sortSelectionRowsByKeyColumn(
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
  ): void {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    if (!Number.isInteger(sortCol) || sortCol < 0 || sortCol >= sheet.colCount) {
      return;
    }
    const range = normalizeSelectionRange(this.selection.getNormalizedRange());
    const rowStart = range.startRow;
    const rowEnd = range.endRow;
    if (kind.type === "value") {
      sheet.sortRowsInRangeByColumn(rowStart, rowEnd, sortCol, kind.direction);
    } else if (kind.type === "fontColorOnTop") {
      const ap = sheet.getMergeAnchorCell(kind.styleAnchorRow, kind.styleAnchorCol);
      const fg = sheet.getCell(ap.row, ap.col).style?.fgArgb;
      const target = fg !== undefined && fg !== "" ? fg.toUpperCase() : null;
      sheet.sortRowsInRangeByColumnFontColor(rowStart, rowEnd, sortCol, target, "asc");
    } else {
      const ap = sheet.getMergeAnchorCell(kind.styleAnchorRow, kind.styleAnchorCol);
      const fill = sheet.getCell(ap.row, ap.col).style?.fillArgb;
      const target = fill !== undefined && fill !== "" ? fill.toUpperCase() : null;
      sheet.sortRowsInRangeByColumnFillColor(rowStart, rowEnd, sortCol, target, "asc");
    }
    recalcWorksheet(sheet);
    this.refresh();
  }

  /**
   * 按当前选区启用列自动筛选（与右键「筛选」、数据选项卡「筛选」一致；
   * 活动单元格行/列 + 选区行范围，下拉绘制在表体顶行等逻辑见 `Worksheet.enableColumnAutoFilterFromSelection`）。
   */
  enableColumnAutoFilterFromSelection(): void {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const ac = this.selection.getActiveCell();
    sheet.enableColumnAutoFilterFromSelection(ac.row, ac.col, this.selection.getNormalizedRange());
    this.refresh();
  }

  /** 清除当前工作表上全部列自动筛选。 */
  clearAllColumnAutoFilters(): void {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    sheet.clearAllColumnAutoFilters();
    this.refresh();
  }

  /** 按当前筛选条件重新计算隐藏行（「重新应用」）。 */
  reapplyAutoFilterConcealment(): void {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    sheet.reapplyAutoFilterConcealment();
    this.refresh();
  }

  /** 简单「自定义排序」：指定列标与升/降序，在选区行范围内按值排序。 */
  openCustomSortDialog(): void {
    if (this.isCellEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    ensureFsSheetPromptStyles();
    this.closeCustomSortDialog();

    const ac = this.selection.getActiveCell();
    const defaultLabel = columnIndexToLabel(ac.col);

    const overlay = document.createElement("div");
    overlay.className = "fs-sheet-prompt-overlay";
    overlay.setAttribute("role", "presentation");

    const panel = document.createElement("div");
    panel.className = "fs-sheet-prompt";
    panel.style.width = "min(340px, calc(100vw - 32px))";
    panel.setAttribute("role", "dialog");
    panel.setAttribute("aria-modal", "true");
    panel.setAttribute("aria-labelledby", "fs-custom-sort-title");

    const header = document.createElement("div");
    header.className = "fs-sheet-prompt__header";
    const titleEl = document.createElement("div");
    titleEl.id = "fs-custom-sort-title";
    titleEl.className = "fs-sheet-prompt__title";
    titleEl.textContent = "自定义排序";
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = "fs-sheet-prompt__close";
    closeBtn.setAttribute("aria-label", "关闭");
    closeBtn.textContent = "×";
    header.appendChild(titleEl);
    header.appendChild(closeBtn);

    const body = document.createElement("div");
    body.className = "fs-sheet-prompt__body";

    const rowCol = document.createElement("label");
    rowCol.className = "fs-sheet-prompt__label";
    const colSpan = document.createElement("span");
    colSpan.textContent = "列标";
    const colInput = document.createElement("input");
    colInput.type = "text";
    colInput.className = "fs-sheet-prompt__input";
    colInput.value = defaultLabel;
    colInput.setAttribute("autocomplete", "off");
    colInput.setAttribute("spellcheck", "false");
    rowCol.appendChild(colSpan);
    rowCol.appendChild(colInput);

    const rowOrder = document.createElement("label");
    rowOrder.className = "fs-sheet-prompt__label";
    const orderSpan = document.createElement("span");
    orderSpan.textContent = "次序";
    const orderSel = document.createElement("select");
    orderSel.className = "fs-sheet-prompt__select";
    const optAsc = document.createElement("option");
    optAsc.value = "asc";
    optAsc.textContent = "升序";
    const optDesc = document.createElement("option");
    optDesc.value = "desc";
    optDesc.textContent = "降序";
    orderSel.appendChild(optAsc);
    orderSel.appendChild(optDesc);
    rowOrder.appendChild(orderSpan);
    rowOrder.appendChild(orderSel);

    body.appendChild(rowCol);
    body.appendChild(rowOrder);

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
    footer.appendChild(cancelBtn);
    footer.appendChild(okBtn);

    panel.appendChild(header);
    panel.appendChild(body);
    panel.appendChild(footer);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);
    this.customSortOverlay = overlay;

    const close = (): void => {
      this.closeCustomSortDialog();
    };

    const tryConfirm = (): void => {
      const colIdx = columnLabelToIndex(colInput.value);
      if (colIdx === null || colIdx < 0 || colIdx >= sheet.colCount) {
        colInput.focus();
        colInput.select();
        return;
      }
      const direction = orderSel.value === "desc" ? "desc" : "asc";
      const range = normalizeSelectionRange(this.selection.getNormalizedRange());
      sheet.sortRowsInRangeByColumn(range.startRow, range.endRow, colIdx, direction);
      recalcWorksheet(sheet);
      this.refresh();
      close();
    };

    const onOverlayPointerDown = (ev: PointerEvent): void => {
      if (ev.target === overlay) {
        close();
      }
    };
    overlay.addEventListener("pointerdown", onOverlayPointerDown);
    closeBtn.addEventListener("click", close);
    cancelBtn.addEventListener("click", close);
    okBtn.addEventListener("click", tryConfirm);
    colInput.addEventListener("keydown", (kev) => {
      if (kev.key === "Enter") {
        kev.preventDefault();
        tryConfirm();
      }
    });

    requestAnimationFrame(() => {
      colInput.focus();
      colInput.select();
    });
  }

  private closeCustomSortDialog(): void {
    if (this.customSortOverlay !== null) {
      this.customSortOverlay.remove();
      this.customSortOverlay = null;
    }
  }

  /**
   * Ribbon「套用表格格式」：在样式库中选定预设计后弹出「表数据来源」对话框；
   * 确定后对目标区域应用表样式，若勾选「表包含标题」则为各列启用自动筛选。
   */
  openFormatAsTableFromRibbon(ribbonCommandId: string): void {
    if (this.isCellEditing()) {
      return;
    }
    showFormatAsTableDialog(this, ribbonCommandId);
  }

  /** Ribbon「套用表格格式」菜单项：打开「新建表样式」对话框。 */
  openNewTableStyleDialog(): void {
    if (this.isCellEditing()) {
      return;
    }
    showNewTableStyleDialog(this);
  }

  /** Ribbon「填充 -> 系列」：打开系列填充对话框。 */
  openFillSeriesDialog(): void {
    if (this.isCellEditing()) {
      return;
    }
    showFillSeriesDialog(this);
  }

  /** Ribbon「插入 -> 数据透视表」：打开数据透视表创建对话框。 */
  openPivotTableDialog(): void {
    if (this.isCellEditing()) {
      return;
    }
    showPivotTableDialog(this);
  }

  /**
   * 当活动单元格位于已注册透视输出区域内时，打开右侧「数据透视表字段」窗格。
   * Ribbon「数据 -> 字段列表」。
   */
  openPivotTableFieldsPane(): void {
    if (this.isCellEditing()) {
      return;
    }
    const sh = this.workbook.getActiveSheet();
    if (sh === undefined) {
      return;
    }
    const ac = this.selection.getActiveCell();
    tryOpenPivotFieldsPaneForSelection(this, sh, ac.row, ac.col);
  }

  /**
   * 收起对话框后在表格上拖拽框选，返回与「套用表格格式 / 数据透视表」输入框一致的绝对引用（如 `=$A$1:$C$10`）。
   * `mode: "singleCell"` 时仅取活动单元格（如 `=$A$1`）。按 ESC 取消，返回 `null`。
   * 使用箭头函数属性，避免对话框内 `const pick = host.pickRangeReferenceFromSheet` 解构后丢失 `this`。
   */
  pickRangeReferenceFromSheet = (options?: {
    readonly mode?: "range" | "singleCell";
    readonly onRangePreview?: (displayRef: string) => void;
  }): Promise<string | null> => {
    if (this.rangeReferencePick !== null) {
      return Promise.resolve(null);
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return Promise.resolve(null);
    }
    return new Promise<string | null>((resolve) => {
      const savedSelection = normalizeSelectionRange(this.selection.getNormalizedRange());
      const mode = options?.mode ?? "range";
      const escHandler = (ev: KeyboardEvent): void => {
        if (ev.key !== "Escape" || ev.isComposing) {
          return;
        }
        ev.preventDefault();
        this.cancelRangeReferencePick();
      };
      document.addEventListener("keydown", escHandler, true);
      this.rangeReferencePick = {
        resolve,
        mode,
        savedSelection,
        escHandler,
        onRangePreview: options?.onRangePreview,
      };
      queueMicrotask(() => {
        this.emitRangePickPreviewIfNeeded();
        this.canvas.focus();
      });
    });
  };

  /**
   * 注册一个新建表样式，并返回稳定 id。
   * 当前实现先将自定义样式映射到一套可用的内置预设，保证样式库中可见且可套用。
   */
  createCustomTableStyle(name: string): string {
    const trimmed = name.trim();
    const displayName =
      trimmed.length > 0 ? trimmed : `表样式 ${this.customTableStyles.length + 1}`;
    const id = `custom-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    this.customTableStyles.push({
      id,
      name: displayName,
      commandId: "home.style.table.medium.r2c0",
    });
    return id;
  }

  /** Ribbon「套用表格格式」读取“自定义”分组列表。 */
  getCustomTableStyleEntries(): readonly {
    readonly id: string;
    readonly name: string;
    readonly commandId: string;
  }[] {
    return this.customTableStyles.map((it) => ({
      id: it.id,
      name: it.name,
      commandId: it.commandId,
    }));
  }

  /** 在列筛选作用行范围内按该列升序/降序排序并重算公式。 */
  sortActiveSheetByColumn(col: number, direction: "asc" | "desc"): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const meta = sheet.getColumnAutoFilterMeta(col);
    if (meta === undefined) {
      return;
    }
    sheet.sortRowsInRangeByColumn(meta.rowStart, meta.rowEnd, col, direction);
    recalcWorksheet(sheet);
    this.refresh();
  }

  /** 在列筛选作用行范围内按字体颜色排序；`targetArgb` 为 `null` 时表示「自动」优先。 */
  sortActiveSheetByColumnFontColor(
    col: number,
    targetArgb: string | null,
    direction: "asc" | "desc",
  ): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const meta = sheet.getColumnAutoFilterMeta(col);
    if (meta === undefined) {
      return;
    }
    sheet.sortRowsInRangeByColumnFontColor(meta.rowStart, meta.rowEnd, col, targetArgb, direction);
    recalcWorksheet(sheet);
    this.refresh();
  }

  /** 打开列筛选面板（`col` 须已启用自动筛选）。 */
  openColumnFilterUi(col: number, clientX: number, clientY: number): void {
    openColumnFilterPanel({ flex: this, col, clientX, clientY });
  }

  /** 打开数据透视表某一筛选项的多选面板。 */
  openPivotFilterUi(
    pivotSheet: Worksheet,
    pivotDefId: string,
    filterFieldIndex: number,
    clientX: number,
    clientY: number,
  ): void {
    openPivotFilterPanel({
      flex: this,
      pivotSheet,
      pivotDefId,
      filterFieldIndex,
      clientX,
      clientY,
    });
  }

  private tryOpenPivotFilterFromClient(clientX: number, clientY: number): boolean {
    if (this.rangeReferencePick !== null) {
      return false;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return false;
    }
    const hit = this.hitTestClient(clientX, clientY);
    if (hit === null) {
      return false;
    }
    for (const def of sheet.getPivotTableDefinitionsSnapshot()) {
      const fCount = def.filterFieldCols.length;
      if (fCount === 0) {
        continue;
      }
      const filterBase = def.pageFilterStartRow ?? def.destinationRow;
      const c0 = def.destinationCol;
      for (let i = 0; i < fCount; i++) {
        if (
          hit.row === filterBase + i &&
          (hit.col === c0 || hit.col === c0 + 1)
        ) {
          this.openPivotFilterUi(sheet, def.id, i, clientX, clientY);
          return true;
        }
      }
    }
    return false;
  }

  destroy(): void {
    this.cancelDragAutoscrollRaf();
    this.headingDrag = null;
    this.inlineFormulaRefDrag = null;
    this.formulaArgRangeSession = false;
    this.detachDocumentDragListeners();
    this.renderer.cancelPendingRedraw();
    this.hideResizePreviewLine();
    this.activeSheetFormattingUnsub?.();
    this.activeSheetFormattingUnsub = null;
    this.formattingChromeListeners.clear();
    this.workbookUnsub?.();
    this.workbookUnsub = null;
    this.canvas.removeEventListener("pointerdown", this.onPointerDown);
    this.canvas.removeEventListener("pointermove", this.onCanvasPointerMove);
    this.canvas.removeEventListener("pointerleave", this.onCanvasPointerLeave);
    this.canvas.removeEventListener("dblclick", this.onCanvasDoubleClick);
    this.canvas.removeEventListener("lostpointercapture", this.onLostPointerCapture);
    this.detachDocumentDragListeners();
    this.canvas.removeEventListener("keydown", this.onKeyDown);
    if (this.chromeRootSortShortcutCleanup !== null) {
      this.chromeRootSortShortcutCleanup();
      this.chromeRootSortShortcutCleanup = null;
    }
    if (this.resizeObserver !== null) {
      this.resizeObserver.disconnect();
      this.resizeObserver = null;
    } else {
      window.removeEventListener("resize", this.onWindowResize);
    }
    this.closeCustomSortDialog();
    this.workspace.destroy();
  }

  private bindResize(): void {
    if (typeof ResizeObserver === "undefined") {
      window.addEventListener("resize", this.onWindowResize);
      this.syncSizeAndDraw();
      return;
    }
    this.resizeObserver = new ResizeObserver(() => {
      this.syncSizeAndDraw();
    });
    this.resizeObserver.observe(this.canvas.parentElement ?? this.canvas);
  }

  private readonly onWindowResize = (): void => {
    this.syncSizeAndDraw();
  };

  private bindSelectionPointer(): void {
    this.canvas.addEventListener("pointerdown", this.onPointerDown);
    this.canvas.addEventListener("pointermove", this.onCanvasPointerMove);
    this.canvas.addEventListener("pointerleave", this.onCanvasPointerLeave);
    this.canvas.addEventListener("dblclick", this.onCanvasDoubleClick);
    this.canvas.addEventListener("lostpointercapture", this.onLostPointerCapture);
  }

  private attachDocumentDragListeners(): void {
    this.detachDocumentDragListeners();
    document.addEventListener("pointermove", this.onDocumentPointerMoveCapture, true);
    document.addEventListener("pointerup", this.onDocumentPointerUpCapture, true);
    document.addEventListener("pointercancel", this.onDocumentPointerCancelCapture, true);
  }

  private detachDocumentDragListeners(): void {
    document.removeEventListener("pointermove", this.onDocumentPointerMoveCapture, true);
    document.removeEventListener("pointerup", this.onDocumentPointerUpCapture, true);
    document.removeEventListener("pointercancel", this.onDocumentPointerCancelCapture, true);
  }

  private finishDrag(ev?: PointerEvent): void {
    if (this.resizing) {
      const sheet = this.workbook.getActiveSheet();
      if (sheet !== undefined && this.resizingKind !== null && this.resizingIndex >= 0) {
        const finalSize = this.resizingCurrentSize;
        if (this.resizingKind === "col") {
          if (Math.abs(finalSize - this.resizingOriginSize) > 1e-6) {
            this.workspace.commands.execute(
              new SetColWidthCommand(sheet, this.resizingIndex, finalSize),
            );
          }
        } else {
          if (Math.abs(finalSize - this.resizingOriginSize) > 1e-6) {
            this.workspace.commands.execute(
              new SetRowHeightCommand(sheet, this.resizingIndex, finalSize),
            );
          }
        }
      }
      this.resizing = false;
      this.resizingKind = null;
      this.resizingIndex = -1;
      this.hoverResizeKind = null;
      this.hideResizePreviewLine();
      this.dragPointerId = null;
      this.detachDocumentDragListeners();
      this.clearResizeDragBodyCursor();
      if (ev !== undefined) {
        this.syncHeadingCursorFromClient(ev.clientX, ev.clientY);
        try {
          this.canvas.releasePointerCapture(ev.pointerId);
        } catch {
          /* ignore */
        }
      } else {
        this.applyPointerCursor("default");
      }
      this.selection.syncWithSheet();
      this.afterSelectionChanged();
      return;
    }

    if (this.headingDrag !== null) {
      this.cancelDragAutoscrollRaf();
      this.headingDrag = null;
      this.dragPointerId = null;
      this.detachDocumentDragListeners();
      if (ev !== undefined) {
        try {
          this.canvas.releasePointerCapture(ev.pointerId);
        } catch {
          /* ignore */
        }
      }
      this.afterSelectionChanged();
      this.tryCommitRangeReferencePick();
      return;
    }

    if (this.fillDrag !== null) {
      this.cancelDragAutoscrollRaf();
      const fd = this.fillDrag;
      this.fillDrag = null;
      this.dragPointerId = null;
      this.detachDocumentDragListeners();
      if (ev !== undefined) {
        try {
          this.canvas.releasePointerCapture(ev.pointerId);
        } catch {
          /* ignore */
        }
      }
      const activeSheet = this.workbook.getActiveSheet();
      if (activeSheet !== undefined) {
        const F = this.computeFillPreviewRange(fd.sourceRange, fd.previewRow, fd.previewCol);
        const S = normalizeSelectionRange(fd.sourceRange);
        if (!selectionRangesEqualNormalized(F, S)) {
          this.workspace.commands.execute(new AutofillExtendCommand(activeSheet, S, F));
        }
        this.selection.setNormalizedRange(F);
      }
      this.afterSelectionChanged();
      if (ev !== undefined) {
        this.syncHeadingCursorFromClient(ev.clientX, ev.clientY);
      }
      return;
    }

    if (this.inlineFormulaRefDrag !== null) {
      this.cancelDragAutoscrollRaf();
      this.inlineFormulaRefDrag = null;
      this.dragPointerId = null;
      this.detachDocumentDragListeners();
      if (ev !== undefined) {
        try {
          this.canvas.releasePointerCapture(ev.pointerId);
        } catch {
          /* ignore */
        }
      }
      this.applyPointerCursor("default");
      const t = this.cellEditor.getEditingText();
      const tlen = t.length;
      this.cellEditor.setEditingText(t, tlen, tlen);
      queueMicrotask(() => {
        this.cellEditor.refocusInput();
      });
      this.updateAutoSumFunctionHintFromEditText(t);
      this.refresh();
      return;
    }

    if (!this.dragSelecting) {
      return;
    }
    this.cancelDragAutoscrollRaf();
    this.dragSelecting = false;
    this.dragPointerId = null;
    this.detachDocumentDragListeners();
    if (ev !== undefined) {
      try {
        this.canvas.releasePointerCapture(ev.pointerId);
      } catch {
        /* ignore */
      }
    }
    this.afterSelectionChanged();
    this.tryCommitRangeReferencePick();
  }

  private formatAbsoluteRangeRefString(range: SelectionRange): string {
    const n = normalizeSelectionRange(range);
    const c0 = columnIndexToLabel(n.startCol);
    const c1 = columnIndexToLabel(n.endCol);
    const r0 = n.startRow + 1;
    const r1 = n.endRow + 1;
    return `=$${c0}$${r0}:$${c1}$${r1}`;
  }

  private formatAbsoluteCellRefString(row: number, col: number): string {
    return `=$${columnIndexToLabel(col)}$${row + 1}`;
  }

  private tryCommitRangeReferencePick(): void {
    const session = this.rangeReferencePick;
    if (session === null) {
      return;
    }
    document.removeEventListener("keydown", session.escHandler, true);
    if (this.workbook.getActiveSheet() === undefined) {
      session.resolve(null);
      this.rangeReferencePick = null;
      return;
    }
    const text =
      session.mode === "singleCell"
        ? (() => {
            const ac = this.selection.getActiveCell();
            return this.formatAbsoluteCellRefString(ac.row, ac.col);
          })()
        : this.formatAbsoluteRangeRefString(this.selection.getNormalizedRange());
    session.resolve(text);
    this.rangeReferencePick = null;
  }

  private cancelRangeReferencePick(): void {
    const session = this.rangeReferencePick;
    if (session === null) {
      return;
    }
    document.removeEventListener("keydown", session.escHandler, true);
    this.selection.setNormalizedRange(session.savedSelection);
    this.afterSelectionChanged();
    session.resolve(null);
    this.rangeReferencePick = null;
  }

  private bindSelectionKeyboard(): void {
    this.canvas.addEventListener("keydown", this.onKeyDown);
  }

  private readonly onLostPointerCapture = (ev: PointerEvent): void => {
    if (
      (!this.dragSelecting &&
        !this.resizing &&
        this.headingDrag === null &&
        this.fillDrag === null &&
        this.inlineFormulaRefDrag === null) ||
      ev.pointerId !== this.dragPointerId
    ) {
      return;
    }
    this.finishDrag();
  };

  private readonly onCanvasPointerMove = (ev: PointerEvent): void => {
    if (
      this.resizing ||
      this.dragSelecting ||
      this.headingDrag !== null ||
      this.fillDrag !== null ||
      this.inlineFormulaRefDrag !== null
    ) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet !== undefined && this.tryHitFillHandle(ev.clientX, ev.clientY)) {
      this.applyPointerCursor("crosshair");
      return;
    }
    this.syncHeadingCursorFromClient(ev.clientX, ev.clientY);
  };

  /** 行列标题区：与悬停一致（分割线 col-resize / row-resize，其余 default）。 */
  private syncHeadingCursorFromClient(clientX: number, clientY: number): void {
    const headingHit = this.hitTestHeadingFromClient(clientX, clientY);
    if (headingHit === null) {
      this.hoverResizeKind = null;
      this.applyPointerCursor("default");
      return;
    }
    const resizeHit = this.tryHitResizeHandle(headingHit, clientX, clientY);
    this.hoverResizeKind = resizeHit?.kind ?? null;
    if (this.hoverResizeKind === "col") {
      this.applyPointerCursor("col-resize");
    } else if (this.hoverResizeKind === "row") {
      this.applyPointerCursor("row-resize");
    } else {
      this.applyPointerCursor("default");
    }
  }

  private readonly onCanvasPointerLeave = (): void => {
    if (
      this.resizing ||
      this.dragSelecting ||
      this.headingDrag !== null ||
      this.fillDrag !== null
    ) {
      return;
    }
    this.hoverResizeKind = null;
    this.applyPointerCursor("default");
  };

  private readonly onDocumentPointerMoveCapture = (ev: PointerEvent): void => {
    if (ev.pointerId !== this.dragPointerId) {
      return;
    }
    if (this.resizing) {
      this.handleResizingPointerMove(ev);
      return;
    }
    if (this.fillDrag !== null) {
      this.lastDragClientX = ev.clientX;
      this.lastDragClientY = ev.clientY;
      const hit = this.hitTestClient(ev.clientX, ev.clientY, { clampToBody: true });
      if (hit === null) {
        return;
      }
      this.fillDrag = {
        sourceRange: this.fillDrag.sourceRange,
        previewRow: hit.row,
        previewCol: hit.col,
      };
      this.renderer.requestRedraw();
      return;
    }
    if (this.inlineFormulaRefDrag !== null) {
      this.lastDragClientX = ev.clientX;
      this.lastDragClientY = ev.clientY;
      const sh = this.workbook.getActiveSheet();
      if (sh !== undefined) {
        const hit = this.hitTestClient(ev.clientX, ev.clientY, { clampToBody: true });
        if (hit !== null) {
          const a1 = sh.getMergeAnchorCell(hit.row, hit.col);
          this.inlineFormulaRefDrag.focusR = a1.row;
          this.inlineFormulaRefDrag.focusC = a1.col;
          this.updateInlineFormulaRefTextAndPreview();
        }
      }
      this.renderer.requestRedraw();
      if (this.isDragAutoscrollActive(ev.clientX, ev.clientY)) {
        this.ensureDragAutoscrollRaf();
      }
      return;
    }
    if (this.headingDrag !== null) {
      this.lastDragClientX = ev.clientX;
      this.lastDragClientY = ev.clientY;

      if (this.isDragAutoscrollActive(ev.clientX, ev.clientY)) {
        this.ensureDragAutoscrollRaf();
        return;
      }

      this.cancelDragAutoscrollRaf();
      this.applyHeadingDragSelectionFromClient(ev.clientX, ev.clientY);
      this.afterSelectionChanged();
      return;
    }
    if (!this.dragSelecting) {
      return;
    }
    this.lastDragClientX = ev.clientX;
    this.lastDragClientY = ev.clientY;

    if (this.isDragAutoscrollActive(ev.clientX, ev.clientY)) {
      this.ensureDragAutoscrollRaf();
      return;
    }

    this.cancelDragAutoscrollRaf();
    const hit = this.hitTestClient(ev.clientX, ev.clientY, { clampToBody: true });
    if (hit === null) {
      return;
    }
    this.selection.extendFocusTo(hit.row, hit.col);
    this.afterSelectionChanged();
  };

  private readonly onDocumentPointerUpCapture = (ev: PointerEvent): void => {
    if (ev.pointerId !== this.dragPointerId) {
      return;
    }
    this.finishDrag(ev);
  };

  private readonly onDocumentPointerCancelCapture = (ev: PointerEvent): void => {
    if (ev.pointerId !== this.dragPointerId) {
      return;
    }
    this.finishDrag(ev);
  };

  private beginEditAt(
    row: number,
    col: number,
    options?: BeginEditOptions & { readonly initialTextOverride?: string },
  ): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    this.clearClipboardMarquee();
    const anchor = sheet.getMergeAnchorCell(row, col);
    const rect = this.renderer.getCellRectInCanvasPixels(anchor.row, anchor.col);
    if (rect === null) {
      return;
    }
    const cell = sheet.getCell(anchor.row, anchor.col);
    const text =
      options?.initialTextOverride !== undefined
        ? options.initialTextOverride
        : cell.formula !== null
          ? cell.formula
          : cellScalarToEditString(cell.value);

    let editorOpts: BeginEditOptions | undefined;
    if (options?.cursorClientX !== undefined) {
      editorOpts = { cursorClientX: options.cursorClientX };
    } else if (options?.selectionStart !== undefined) {
      editorOpts = {
        selectAll: false,
        selectionStart: options.selectionStart,
        selectionEnd: options.selectionEnd,
      };
    } else if (options?.selectAll === false) {
      editorOpts = { selectAll: false };
    } else if (options?.initialTextOverride !== undefined) {
      editorOpts = { selectAll: false };
    } else if (options?.selectAll === true) {
      editorOpts = { selectAll: true };
    } else {
      editorOpts = undefined;
    }
    this.cellEditor.beginEdit(anchor.row, anchor.col, text, rect, editorOpts);
  }

  /** 双击按位置进入编辑；单击仅选格（与 Excel 一致，键入 / F2 再进编辑）。 */
  private readonly onCanvasDoubleClick = (ev: MouseEvent): void => {
    ev.preventDefault();
    if (this.cellEditor.isEditing()) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }

    const headingHit = this.hitTestHeadingFromClient(ev.clientX, ev.clientY);
    if (headingHit !== null) {
      const resizeHit = this.tryHitResizeHandle(headingHit, ev.clientX, ev.clientY);
      if (resizeHit !== null) {
        const z = this.renderer.getViewZoom();
        if (resizeHit.kind === "col") {
          const w = computeColumnAutoWidth(sheet, resizeHit.index, z);
          this.workspace.commands.execute(new SetColWidthCommand(sheet, resizeHit.index, w));
        } else {
          const h = computeRowAutoHeight(sheet, resizeHit.index, z);
          this.workspace.commands.execute(new SetRowHeightCommand(sheet, resizeHit.index, h));
        }
        this.refresh();
        this.syncHeadingCursorFromClient(ev.clientX, ev.clientY);
        return;
      }
      this.syncHeadingCursorFromClient(ev.clientX, ev.clientY);
      return;
    }

    const bodyFilterCol = this.tryHitBodyAutoFilterFromClient(ev.clientX, ev.clientY);
    if (bodyFilterCol !== null) {
      this.openColumnFilterUi(bodyFilterCol, ev.clientX, ev.clientY);
      return;
    }

    const hit = this.hitTestClient(ev.clientX, ev.clientY);
    if (hit === null) {
      return;
    }
    this.selection.selectCell(hit.row, hit.col);
    this.afterSelectionChanged();
    this.beginEditAt(hit.row, hit.col, { cursorClientX: ev.clientX });
  };

  private readonly onPointerDown = (ev: PointerEvent): void => {
    if (ev.button !== 0) {
      return;
    }

    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }

    const headingHit = this.hitTestHeadingFromClient(ev.clientX, ev.clientY);
    if (headingHit !== null) {
      this.canvas.focus();
      const resizeHit = this.tryHitResizeHandle(headingHit, ev.clientX, ev.clientY);
      if (resizeHit !== null) {
        ev.preventDefault();
        this.beginResizing(resizeHit.kind, resizeHit.index, ev);
        return;
      }
      if (headingHit.kind === "columnHeader") {
        const w = this.canvas.clientWidth;
        const h = this.canvas.clientHeight;
        const corner = this.renderer.getCornerSize();
        const layout = buildFrozenLayout(
          sheet,
          corner.width,
          corner.height,
          w,
          h,
          this.renderer.frozenRows,
          this.renderer.frozenCols,
          this.renderer.viewZoom,
        );
        const { x, y } = this.clientToCanvasXY(ev.clientX, ev.clientY);
        const filterCol = hitTestColumnHeaderFilterButton(
          x,
          y,
          sheet,
          layout,
          this.renderer.scrollX,
          this.renderer.scrollY,
          this.renderer.viewZoom,
        );
        if (filterCol !== null) {
          if (this.rangeReferencePick === null) {
            ev.preventDefault();
            if (this.cellEditor.isEditing()) {
              this.cellEditor.cancelWithoutCommit();
            }
            this.openColumnFilterUi(filterCol, ev.clientX, ev.clientY);
            return;
          }
        }
      }
      ev.preventDefault();
      if (this.cellEditor.isEditing()) {
        this.cellEditor.cancelWithoutCommit();
      }
      const expand = ev.shiftKey || ev.ctrlKey || ev.metaKey;
      this.applyHeadingHitSelection(headingHit, sheet, expand);
      this.afterSelectionChanged();
      if (!expand && (headingHit.kind === "columnHeader" || headingHit.kind === "rowHeader")) {
        this.lastDragClientX = ev.clientX;
        this.lastDragClientY = ev.clientY;
        this.headingDrag =
          headingHit.kind === "columnHeader"
            ? { kind: "column", originCol: headingHit.col }
            : { kind: "row", originRow: headingHit.row };
        this.dragPointerId = ev.pointerId;
        this.attachDocumentDragListeners();
        try {
          this.canvas.setPointerCapture(ev.pointerId);
        } catch {
          /* ignore */
        }
      }
      this.syncHeadingCursorFromClient(ev.clientX, ev.clientY);
      return;
    }

    const bodyFilterCol = this.tryHitBodyAutoFilterFromClient(ev.clientX, ev.clientY);
    if (bodyFilterCol !== null) {
      if (this.rangeReferencePick === null) {
        this.canvas.focus();
        ev.preventDefault();
        if (this.cellEditor.isEditing()) {
          this.cellEditor.cancelWithoutCommit();
        }
        this.openColumnFilterUi(bodyFilterCol, ev.clientX, ev.clientY);
        return;
      }
    }

    if (this.tryOpenPivotFilterFromClient(ev.clientX, ev.clientY)) {
      this.canvas.focus();
      ev.preventDefault();
      if (this.cellEditor.isEditing()) {
        this.cellEditor.cancelWithoutCommit();
      }
      return;
    }

    if (this.tryHitFillHandle(ev.clientX, ev.clientY)) {
      this.canvas.focus();
      ev.preventDefault();
      if (this.cellEditor.isEditing()) {
        this.cellEditor.cancelWithoutCommit();
      }
      const sourceRange = expandSelectionRangeForMergePaint(
        sheet,
        normalizeSelectionRange(this.selection.getNormalizedRange()),
      );
      this.fillDrag = {
        sourceRange,
        previewRow: sourceRange.endRow,
        previewCol: sourceRange.endCol,
      };
      this.lastDragClientX = ev.clientX;
      this.lastDragClientY = ev.clientY;
      this.dragPointerId = ev.pointerId;
      this.attachDocumentDragListeners();
      try {
        this.canvas.setPointerCapture(ev.pointerId);
      } catch {
        /* ignore */
      }
      this.applyPointerCursor("crosshair");
      this.renderer.requestRedraw();
      return;
    }

    const hit = this.hitTestClient(ev.clientX, ev.clientY);
    if (hit === null) {
      this.canvas.focus();
      return;
    }
    if (this.rangeReferencePick === null && this.shouldStartInlineRefDragOnBody()) {
      this.beginInlineFormulaRefDrag(sheet, hit, ev);
      return;
    }
    this.canvas.focus();
    if (this.cellEditor.isEditing()) {
      this.cellEditor.cancelWithoutCommit();
    }
    this.selection.selectCell(hit.row, hit.col);
    this.revealActiveCellInViewport();
    this.lastDragClientX = ev.clientX;
    this.lastDragClientY = ev.clientY;
    this.dragSelecting = true;
    this.dragPointerId = ev.pointerId;
    this.attachDocumentDragListeners();
    try {
      this.canvas.setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
    this.cellEditor.syncLayout();
    this.renderer.requestRedraw();
  };

  private tryHitResizeHandle(
    headingHit: HeadingHit,
    clientX: number,
    clientY: number,
  ): { kind: "row" | "col"; index: number } | null {
    const edgeTolerance = 4;
    if (headingHit.kind === "columnHeader") {
      const rect = this.renderer.getCellRectInCanvasPixels(0, headingHit.col);
      if (rect === null) return null;
      const { x } = this.clientToCanvasXY(clientX, clientY);
      if (Math.abs(x - (rect.x + rect.width)) <= edgeTolerance) {
        return { kind: "col", index: headingHit.col };
      }
      if (headingHit.col > 0 && Math.abs(x - rect.x) <= edgeTolerance) {
        return { kind: "col", index: headingHit.col - 1 };
      }
      return null;
    }
    if (headingHit.kind === "rowHeader") {
      const rect = this.renderer.getCellRectInCanvasPixels(headingHit.row, 0);
      if (rect === null) return null;
      const { y } = this.clientToCanvasXY(clientX, clientY);
      if (Math.abs(y - (rect.y + rect.height)) <= edgeTolerance) {
        return { kind: "row", index: headingHit.row };
      }
      if (headingHit.row > 0 && Math.abs(y - rect.y) <= edgeTolerance) {
        return { kind: "row", index: headingHit.row - 1 };
      }
    }
    return null;
  }

  private beginResizing(kind: "row" | "col", index: number, ev: PointerEvent): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) return;
    if (this.rangeReferencePick !== null) {
      this.cancelRangeReferencePick();
    }
    this.resizing = true;
    this.resizingKind = kind;
    this.resizingIndex = index;
    this.resizingStartClientX = ev.clientX;
    this.resizingStartClientY = ev.clientY;
    this.resizingOriginSize = kind === "col" ? sheet.getColWidth(index) : sheet.getRowHeight(index);
    this.resizingCurrentSize = this.resizingOriginSize;
    this.updateResizePreviewLine(kind, index, this.resizingCurrentSize);
    const resizeCursor = kind === "col" ? "col-resize" : "row-resize";
    this.applyPointerCursor(resizeCursor);
    this.setResizeDragBodyCursor(kind);
    this.dragPointerId = ev.pointerId;
    this.attachDocumentDragListeners();
    try {
      this.canvas.setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
  }

  private applyPointerCursor(cursor: string): void {
    this.canvas.style.cursor = cursor;
  }

  private setResizeDragBodyCursor(kind: "col" | "row"): void {
    if (typeof document !== "undefined") {
      document.body.style.cursor = kind === "col" ? "col-resize" : "row-resize";
    }
  }

  private clearResizeDragBodyCursor(): void {
    if (typeof document !== "undefined") {
      document.body.style.removeProperty("cursor");
    }
  }

  private ensureResizePreviewLine(): HTMLDivElement {
    if (this.resizePreviewEl !== null) {
      return this.resizePreviewEl;
    }
    const line = document.createElement("div");
    line.style.position = "absolute";
    line.style.pointerEvents = "none";
    line.style.display = "none";
    line.style.zIndex = "20";
    line.style.border = "0";
    if (getComputedStyle(this.host).position === "static") {
      this.host.style.position = "relative";
    }
    this.host.appendChild(line);
    this.resizePreviewEl = line;
    return line;
  }

  private updateResizePreviewLine(kind: "row" | "col", index: number, nextSize: number): void {
    const line = this.ensureResizePreviewLine();
    const canvasRect = this.canvas.getBoundingClientRect();
    const hostRect = this.host.getBoundingClientRect();
    const left0 = canvasRect.left - hostRect.left;
    const top0 = canvasRect.top - hostRect.top;
    if (kind === "col") {
      const rect = this.renderer.getCellRectInCanvasPixels(0, index);
      if (rect === null) {
        return;
      }
      const x = left0 + rect.x + nextSize * this.renderer.getViewZoom();
      line.style.display = "block";
      line.style.left = `${Math.round(x)}px`;
      line.style.top = `${Math.round(top0)}px`;
      line.style.width = "0px";
      line.style.height = `${Math.round(this.canvas.clientHeight)}px`;
      line.style.borderTop = "0";
      line.style.borderLeft = `1px dashed ${this.getResizePreviewColor()}`;
    } else {
      const rect = this.renderer.getCellRectInCanvasPixels(index, 0);
      if (rect === null) {
        return;
      }
      const y = top0 + rect.y + nextSize * this.renderer.getViewZoom();
      line.style.display = "block";
      line.style.left = `${Math.round(left0)}px`;
      line.style.top = `${Math.round(y)}px`;
      line.style.width = `${Math.round(this.canvas.clientWidth)}px`;
      line.style.height = "0px";
      line.style.borderLeft = "0";
      line.style.borderTop = `1px dashed ${this.getResizePreviewColor()}`;
    }
  }

  private hideResizePreviewLine(): void {
    if (this.resizePreviewEl !== null) {
      this.resizePreviewEl.style.display = "none";
    }
  }

  /** 预览线颜色：比网格线稍深，避免与激活态边框混淆。 */
  private getResizePreviewColor(): string {
    return darkenColor(this.theme.gridLineColor, 0.22);
  }

  private handleResizingPointerMove(ev: PointerEvent): void {
    if (this.resizingKind === null || this.resizingIndex < 0) {
      return;
    }
    if (this.resizingKind === "col") {
      const delta = ev.clientX - this.resizingStartClientX;
      const next = Math.max(8, this.resizingOriginSize + delta / this.renderer.getViewZoom());
      this.resizingCurrentSize = next;
    } else {
      const delta = ev.clientY - this.resizingStartClientY;
      const next = Math.max(8, this.resizingOriginSize + delta / this.renderer.getViewZoom());
      this.resizingCurrentSize = next;
    }
    this.updateResizePreviewLine(this.resizingKind, this.resizingIndex, this.resizingCurrentSize);
    const k = this.resizingKind;
    if (k === "col") {
      this.applyPointerCursor("col-resize");
    } else if (k === "row") {
      this.applyPointerCursor("row-resize");
    }
  }

  private readonly onKeyDown = (ev: KeyboardEvent): void => {
    if (this.cellEditor.isEditing()) {
      return;
    }
    if (this.workbook.getActiveSheet() === undefined) {
      return;
    }

    if (ev.shiftKey && (ev.metaKey || ev.ctrlKey) && (ev.key === "r" || ev.key === "R")) {
      ev.preventDefault();
      this.openCustomSortDialog();
      return;
    }

    if (ev.ctrlKey || ev.metaKey || ev.altKey) {
      return;
    }

    const arrow =
      ev.key === "ArrowUp" ||
      ev.key === "ArrowDown" ||
      ev.key === "ArrowLeft" ||
      ev.key === "ArrowRight";
    if (arrow) {
      ev.preventDefault();
      const extend = ev.shiftKey;
      let dr = 0;
      let dc = 0;
      if (ev.key === "ArrowUp") {
        dr = -1;
      } else if (ev.key === "ArrowDown") {
        dr = 1;
      } else if (ev.key === "ArrowLeft") {
        dc = -1;
      } else if (ev.key === "ArrowRight") {
        dc = 1;
      }
      this.selection.moveFocus(dr, dc, extend);
      this.afterSelectionChanged();
      return;
    }

    if (ev.key === "Enter") {
      ev.preventDefault();
      this.selection.moveFocus(1, 0, false);
      this.afterSelectionChanged();
      return;
    }

    if (ev.key === "Tab") {
      ev.preventDefault();
      if (ev.shiftKey) {
        this.selection.moveFocus(0, -1, false);
      } else {
        this.selection.moveFocus(0, 1, false);
      }
      this.afterSelectionChanged();
      return;
    }

    if (ev.key === "F2") {
      ev.preventDefault();
      const ac = this.selection.getActiveCell();
      this.beginEditAt(ac.row, ac.col, { selectAll: false });
      return;
    }

    if (ev.key === "Delete" || ev.key === "Backspace") {
      ev.preventDefault();
      this.clearSelectionContents();
      return;
    }

    if (ev.key.length === 1 && !ev.isComposing) {
      ev.preventDefault();
      const ac = this.selection.getActiveCell();
      this.beginEditAt(ac.row, ac.col, { initialTextOverride: ev.key });
      return;
    }
  };

  private clientToCanvasXY(clientX: number, clientY: number): { x: number; y: number } {
    const rect = this.canvas.getBoundingClientRect();
    const cw = this.canvas.clientWidth;
    const ch = this.canvas.clientHeight;
    const rw = Math.max(rect.width, 1e-6);
    const rh = Math.max(rect.height, 1e-6);
    return {
      x: ((clientX - rect.left) / rw) * cw,
      y: ((clientY - rect.top) / rh) * ch,
    };
  }

  private hitTestHeadingFromClient(clientX: number, clientY: number): HeadingHit | null {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return null;
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const { x, y } = this.clientToCanvasXY(clientX, clientY);
    const corner = this.renderer.getCornerSize();
    const layout = buildFrozenLayout(
      sheet,
      corner.width,
      corner.height,
      w,
      h,
      this.renderer.frozenRows,
      this.renderer.frozenCols,
      this.renderer.viewZoom,
    );
    return hitTestHeadingPointer(
      x,
      y,
      sheet,
      layout,
      this.renderer.scrollX,
      this.renderer.scrollY,
      this.renderer.viewZoom,
    );
  }

  /** 命中表体内列筛选锚点按钮时返回列号，否则 `null`。 */
  private tryHitBodyAutoFilterFromClient(clientX: number, clientY: number): number | null {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return null;
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const { x, y } = this.clientToCanvasXY(clientX, clientY);
    const corner = this.renderer.getCornerSize();
    const layout = buildFrozenLayout(
      sheet,
      corner.width,
      corner.height,
      w,
      h,
      this.renderer.frozenRows,
      this.renderer.frozenCols,
      this.renderer.viewZoom,
    );
    return hitTestBodyCellAutoFilterButton(
      x,
      y,
      sheet,
      layout,
      this.renderer.scrollX,
      this.renderer.scrollY,
      this.renderer.viewZoom,
    );
  }

  private applyHeadingHitSelection(hit: HeadingHit, sheet: Worksheet, expand: boolean): void {
    const lastR = Math.max(0, sheet.rowCount - 1);
    const lastC = Math.max(0, sheet.colCount - 1);
    if (hit.kind === "selectAllCorner") {
      if (expand) {
        this.selection.unionWithRange({
          startRow: 0,
          startCol: 0,
          endRow: lastR,
          endCol: lastC,
        });
      } else {
        this.selection.selectEntireSheet();
      }
      return;
    }
    if (hit.kind === "columnHeader") {
      const c = hit.col;
      if (expand) {
        this.selection.unionWithRange({
          startRow: 0,
          startCol: c,
          endRow: lastR,
          endCol: c,
        });
      } else {
        this.selection.selectEntireColumn(c);
      }
      return;
    }
    if (hit.kind === "rowHeader") {
      const r = hit.row;
      if (expand) {
        this.selection.unionWithRange({
          startRow: r,
          startCol: 0,
          endRow: r,
          endCol: lastC,
        });
      } else {
        this.selection.selectEntireRow(r);
      }
    }
  }

  private hitTestClient(
    clientX: number,
    clientY: number,
    options?: { readonly clampToBody?: boolean },
  ): { row: number; col: number } | null {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return null;
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    let { x, y } = this.clientToCanvasXY(clientX, clientY);
    const corner = this.renderer.getCornerSize();
    if (options?.clampToBody === true) {
      const eps = 0.5;
      x = Math.max(corner.width + eps, Math.min(w - eps, x));
      y = Math.max(corner.height + eps, Math.min(h - eps, y));
    }
    const layout = buildFrozenLayout(
      sheet,
      corner.width,
      corner.height,
      w,
      h,
      this.renderer.frozenRows,
      this.renderer.frozenCols,
      this.renderer.viewZoom,
    );
    return hitTestCell(
      x,
      y,
      sheet,
      layout,
      this.renderer.scrollX,
      this.renderer.scrollY,
      this.renderer.viewZoom,
    );
  }

  private computeFillPreviewRange(
    source: SelectionRange,
    hitRow: number,
    hitCol: number,
  ): SelectionRange {
    const S = normalizeSelectionRange(source);
    return normalizeSelectionRange({
      startRow: Math.min(S.startRow, hitRow),
      startCol: Math.min(S.startCol, hitCol),
      endRow: Math.max(S.endRow, hitRow),
      endCol: Math.max(S.endCol, hitCol),
    });
  }

  private tryHitFillHandle(clientX: number, clientY: number): boolean {
    if (this.rangeReferencePick !== null) {
      return false;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined || this.cellEditor.isEditing()) {
      return false;
    }
    const span = expandSelectionRangeForMergePaint(
      sheet,
      normalizeSelectionRange(this.selection.getNormalizedRange()),
    );
    const rect = this.renderer.getCellRectInCanvasPixels(span.endRow, span.endCol);
    if (rect === null) {
      return false;
    }
    const { x, y } = this.clientToCanvasXY(clientX, clientY);
    const z = this.renderer.viewZoom;
    const handleSize = Math.max(4, 6 * SELECTION_OUTLINE_VISUAL_SCALE * z);
    const half = handleSize / 2;
    const pad = 3;
    const cx = rect.x + rect.width;
    const cy = rect.y + rect.height;
    return Math.abs(x - cx) <= half + pad && Math.abs(y - cy) <= half + pad;
  }

  private applyHeadingDragSelectionFromClient(clientX: number, clientY: number): void {
    const drag = this.headingDrag;
    if (drag === null) {
      return;
    }
    if (drag.kind === "column") {
      const endCol = this.resolveColumnIndexForHeadingDrag(clientX, clientY, drag.originCol);
      this.selection.selectEntireColumnRange(drag.originCol, endCol);
      return;
    }
    const endRow = this.resolveRowIndexForHeadingDrag(clientX, clientY, drag.originRow);
    this.selection.selectEntireRowRange(drag.originRow, endRow);
  }

  private resolveColumnIndexForHeadingDrag(
    clientX: number,
    clientY: number,
    fallbackCol: number,
  ): number {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return fallbackCol;
    }
    const heading = this.hitTestHeadingFromClient(clientX, clientY);
    if (heading?.kind === "columnHeader") {
      return heading.col;
    }
    const pt = this.clientToCanvasXY(clientX, clientY);
    let x = pt.x;
    const y = pt.y;
    const corner = this.renderer.getCornerSize();
    const layout = buildFrozenLayout(
      sheet,
      corner.width,
      corner.height,
      this.canvas.clientWidth,
      this.canvas.clientHeight,
      this.renderer.frozenRows,
      this.renderer.frozenCols,
      this.renderer.viewZoom,
    );
    if (x < layout.headerW && y >= layout.headerH) {
      x = layout.headerW + 0.5;
    }
    const cell = hitTestCell(
      x,
      y,
      sheet,
      layout,
      this.renderer.scrollX,
      this.renderer.scrollY,
      this.renderer.viewZoom,
    );
    if (cell !== null) {
      return cell.col;
    }
    return fallbackCol;
  }

  private resolveRowIndexForHeadingDrag(
    clientX: number,
    clientY: number,
    fallbackRow: number,
  ): number {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return fallbackRow;
    }
    const heading = this.hitTestHeadingFromClient(clientX, clientY);
    if (heading?.kind === "rowHeader") {
      return heading.row;
    }
    const pt = this.clientToCanvasXY(clientX, clientY);
    const x = pt.x;
    let y = pt.y;
    const corner = this.renderer.getCornerSize();
    const layout = buildFrozenLayout(
      sheet,
      corner.width,
      corner.height,
      this.canvas.clientWidth,
      this.canvas.clientHeight,
      this.renderer.frozenRows,
      this.renderer.frozenCols,
      this.renderer.viewZoom,
    );
    if (y < layout.headerH && x >= layout.headerW) {
      y = layout.headerH + 0.5;
    }
    const cell = hitTestCell(
      x,
      y,
      sheet,
      layout,
      this.renderer.scrollX,
      this.renderer.scrollY,
      this.renderer.viewZoom,
    );
    if (cell !== null) {
      return cell.row;
    }
    return fallbackRow;
  }

  /**
   * 将活动单元格滚入视口（键盘/单击等）。
   * 鼠标框选拖拽中不调用，避免与边沿自动滚动争抢 scroll。
   */
  private revealActiveCellInViewport(): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const corner = this.renderer.getCornerSize();
    const layout = buildFrozenLayout(
      sheet,
      corner.width,
      corner.height,
      w,
      h,
      this.renderer.frozenRows,
      this.renderer.frozenCols,
      this.renderer.viewZoom,
    );
    const limits = computeScrollLimits(sheet, layout, this.renderer.viewZoom);
    const active = this.selection.getActiveCell();
    const next = scrollToRevealCell(
      sheet,
      layout,
      limits,
      active.row,
      active.col,
      this.renderer.scrollX,
      this.renderer.scrollY,
      this.renderer.viewZoom,
    );
    this.renderer.setScroll(next.scrollX, next.scrollY);
  }

  private afterSelectionChanged(): void {
    if (!this.dragSelecting) {
      this.revealActiveCellInViewport();
    }
    this.cellEditor.syncLayout();
    this.renderer.requestRedraw();
    this.notifyFormattingChrome();
    this.syncFormulaBar();
    this.syncPivotTableFieldsPaneToSelection();
    this.emitRangePickPreviewIfNeeded();
  }

  /** Excel 风格显示用引用：`表名!$A$1:$B$2`（无 `=`），供选区模式悬浮条实时预览。 */
  private emitRangePickPreviewIfNeeded(): void {
    const session = this.rangeReferencePick;
    if (session === null || session.onRangePreview === undefined) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const qName = this.quoteSheetNameForRef(sheet.name);
    let display: string;
    if (session.mode === "singleCell") {
      const ac = this.selection.getActiveCell();
      const col = columnIndexToLabel(ac.col);
      display = `${qName}!$${col}$${ac.row + 1}`;
    } else {
      const n = normalizeSelectionRange(this.selection.getNormalizedRange());
      const c0 = columnIndexToLabel(n.startCol);
      const c1 = columnIndexToLabel(n.endCol);
      display = `${qName}!$${c0}$${n.startRow + 1}:$${c1}$${n.endRow + 1}`;
    }
    try {
      session.onRangePreview(display);
    } catch {
      /* 宿主预览回调异常不影响选区 */
    }
  }

  /** 工作表名在区域引用中的转义（与 Excel 单引号规则一致）。 */
  private quoteSheetNameForRef(name: string): string {
    if (/^[A-Za-z0-9_.]+$/i.test(name)) {
      return name;
    }
    return `'${name.replace(/'/g, "''")}'`;
  }

  /** 活动单元格在透视输出区域内时打开字段窗格，离开时关闭。 */
  private syncPivotTableFieldsPaneToSelection(): void {
    const sh = this.workbook.getActiveSheet();
    const ac = this.selection.getActiveCell();
    syncPivotTableFieldsPaneWithSelection(this, sh, ac.row, ac.col);
  }

  private pivotDefStorageKey(sheetIndex: number, defId: string): string {
    return `${sheetIndex}::${defId}`;
  }

  private syncPivotAutoRefreshBaseline(): void {
    const alive = new Set<string>();
    for (let si = 0; si < this._workbook.sheetCount; si++) {
      const pivotSheet = this._workbook.getSheet(si);
      if (pivotSheet === undefined) {
        continue;
      }
      for (const def of pivotSheet.getPivotTableDefinitionsSnapshot()) {
        const key = this.pivotDefStorageKey(si, def.id);
        alive.add(key);
        const src = this._workbook.getSheet(def.sourceSheetIndex);
        if (src !== undefined) {
          this.pivotSourceRevisionByDefKey.set(key, src.revision);
        }
      }
    }
    for (const key of this.pivotSourceRevisionByDefKey.keys()) {
      if (!alive.has(key)) {
        this.pivotSourceRevisionByDefKey.delete(key);
      }
    }
  }

  private tryAutoRefreshPivotTables(): void {
    if (this.pivotAutoRefreshRunning) {
      return;
    }
    const toRefresh: Array<{ sheet: Worksheet; defId: string }> = [];
    const alive = new Set<string>();
    for (let si = 0; si < this._workbook.sheetCount; si++) {
      const pivotSheet = this._workbook.getSheet(si);
      if (pivotSheet === undefined) {
        continue;
      }
      for (const def of pivotSheet.getPivotTableDefinitionsSnapshot()) {
        const key = this.pivotDefStorageKey(si, def.id);
        alive.add(key);
        const src = this._workbook.getSheet(def.sourceSheetIndex);
        if (src === undefined) {
          continue;
        }
        const prev = this.pivotSourceRevisionByDefKey.get(key);
        const cur = src.revision;
        if (prev === undefined) {
          this.pivotSourceRevisionByDefKey.set(key, cur);
          continue;
        }
        if (prev !== cur) {
          toRefresh.push({ sheet: pivotSheet, defId: def.id });
        }
      }
    }
    for (const key of this.pivotSourceRevisionByDefKey.keys()) {
      if (!alive.has(key)) {
        this.pivotSourceRevisionByDefKey.delete(key);
      }
    }
    if (toRefresh.length === 0) {
      return;
    }
    this.pivotAutoRefreshRunning = true;
    try {
      for (const it of toRefresh) {
        refreshPivotTableDefinition(this._workbook, it.sheet, it.defId);
      }
    } finally {
      this.pivotAutoRefreshRunning = false;
      this.syncPivotAutoRefreshBaseline();
    }
  }

  private notifyFormattingChrome(): void {
    for (const fn of this.formattingChromeListeners) {
      fn();
    }
  }

  private rebindActiveSheetFormattingListener(): void {
    this.activeSheetFormattingUnsub?.();
    this.activeSheetFormattingUnsub = null;
    const sh = this.workbook.getActiveSheet();
    if (sh === undefined) {
      this.notifyFormattingChrome();
      this.syncFormulaBar();
      return;
    }
    this.activeSheetFormattingUnsub = sh.subscribe(() => {
      this.notifyFormattingChrome();
      this.syncFormulaBar();
    });
    this.notifyFormattingChrome();
    this.syncFormulaBar();
  }

  private bindFormulaBar(): void {
    if (this.formulaBarInputEl === null) {
      return;
    }
    this.formulaBarInputEl.addEventListener("keydown", this.onFormulaBarKeyDown);
    this.formulaBarInputEl.addEventListener("blur", this.onFormulaBarBlur);
  }

  private readonly onFormulaBarKeyDown = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      if (ev.isComposing) {
        return;
      }
      ev.preventDefault();
      this.restoreFormulaBarFromActiveCell();
      queueMicrotask(() => {
        this.canvas.focus();
      });
      return;
    }
    if (ev.key !== "Enter" || ev.isComposing) {
      return;
    }
    if (ev.ctrlKey || ev.metaKey) {
      ev.preventDefault();
      this.insertNewlineAtFormulaBarCaret();
      return;
    }
    if (ev.altKey) {
      return;
    }
    ev.preventDefault();
    this.formulaBarSkipBlurCommit = true;
    this.commitFormulaBarIfChanged();
    this.formulaBarInputEl?.blur();
    queueMicrotask(() => {
      this.formulaBarSkipBlurCommit = false;
      this.canvas.focus();
    });
  };

  private readonly onFormulaBarBlur = (): void => {
    if (this.formulaBarSkipBlurCommit) {
      return;
    }
    this.commitFormulaBarIfChanged();
  };

  private insertNewlineAtFormulaBarCaret(): void {
    const el = this.formulaBarInputEl;
    if (el === null) {
      return;
    }
    const start = el.selectionStart ?? 0;
    const end = el.selectionEnd ?? 0;
    const v = el.value;
    el.value = `${v.slice(0, start)}\n${v.slice(end)}`;
    const pos = start + 1;
    el.setSelectionRange(pos, pos);
  }

  /** 取消编辑栏中的修改，恢复为当前活动格内容（与单元格 Esc 一致）。 */
  private restoreFormulaBarFromActiveCell(): void {
    const el = this.formulaBarInputEl;
    if (el === null) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      el.value = "";
      return;
    }
    const { row, col } = this.selection.getActiveCell();
    const a = sheet.getMergeAnchorCell(row, col);
    const cell = sheet.getCell(a.row, a.col);
    el.value =
      cell.formula !== null && cell.formula.length > 0
        ? cell.formula
        : cellScalarToEditString(cell.value);
    el.blur();
  }

  /**
   * 同步编辑栏名称框与输入框：活动单元格地址（合并格以锚点为准）及公式/显示文本；
   * 内联编辑时输入框与单元格编辑器一致。
   */
  private syncFormulaBar(): void {
    if (this.formulaBarNameEl === null && this.formulaBarInputEl === null) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      if (this.formulaBarNameEl !== null) {
        this.formulaBarNameEl.textContent = "";
      }
      if (this.formulaBarInputEl !== null) {
        this.formulaBarInputEl.value = "";
      }
      return;
    }
    if (document.activeElement === this.formulaBarInputEl) {
      return;
    }
    const { row, col } = this.selection.getActiveCell();
    const a = sheet.getMergeAnchorCell(row, col);
    const addr = `${columnIndexToLabel(a.col)}${a.row + 1}`;
    if (this.formulaBarNameEl !== null) {
      this.formulaBarNameEl.textContent = addr;
    }
    if (this.formulaBarInputEl === null) {
      return;
    }
    if (this.cellEditor.isEditing()) {
      this.formulaBarInputEl.value = this.cellEditor.getEditingText();
      return;
    }
    const cell = sheet.getCell(a.row, a.col);
    const text =
      cell.formula !== null && cell.formula.length > 0
        ? cell.formula
        : cellScalarToEditString(cell.value);
    this.formulaBarInputEl.value = text;
  }

  private commitFormulaBarIfChanged(): void {
    if (this.formulaBarInputEl === null) {
      return;
    }
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    if (this.cellEditor.isEditing()) {
      return;
    }
    const { row, col } = this.selection.getActiveCell();
    const a = sheet.getMergeAnchorCell(row, col);
    const cell = sheet.getCell(a.row, a.col);
    const current =
      cell.formula !== null && cell.formula.length > 0
        ? cell.formula
        : cellScalarToEditString(cell.value);
    const raw = this.formulaBarInputEl.value;
    if (raw === current) {
      return;
    }
    const value = parseEditString(raw);
    const cmd = new SetCellValueCommand(sheet, a.row, a.col, value);
    this.workspace.commands.execute(cmd);
    this.autoExpandRowHeightForMultilineValue(sheet, a.row, a.col, value);
    this.refresh();
  }

  private cancelDragAutoscrollRaf(): void {
    if (this.dragAutoscrollRafId !== null) {
      cancelAnimationFrame(this.dragAutoscrollRafId);
      this.dragAutoscrollRafId = null;
    }
  }

  private ensureDragAutoscrollRaf(): void {
    if (this.dragAutoscrollRafId !== null) {
      return;
    }
    this.dragAutoscrollPrevNow = performance.now();
    this.dragAutoscrollRafId = requestAnimationFrame(() => {
      this.dragAutoscrollLoop();
    });
  }

  /**
   * 边沿自动滚动：速度随进入条带的深度线性增加，抵住滚动极限时速度为 0（停止续帧）。
   */
  private computeDragAutoscrollPixelsPerSec(
    clientX: number,
    clientY: number,
  ): {
    sx: number;
    sy: number;
  } {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return { sx: 0, sy: 0 };
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const corner = this.renderer.getCornerSize();
    const layout = buildFrozenLayout(
      sheet,
      corner.width,
      corner.height,
      w,
      h,
      this.renderer.frozenRows,
      this.renderer.frozenCols,
      this.renderer.viewZoom,
    );
    const limits = computeScrollLimits(sheet, layout, this.renderer.viewZoom);
    const { scrollX, scrollY } = this.renderer.getScroll();
    const { x, y } = this.clientToCanvasXY(clientX, clientY);
    const M = DRAG_AUTOSCROLL_MARGIN_PX;
    const MAX = DRAG_AUTOSCROLL_MAX_SPEED;

    const { headerW, headerH, scrollViewportW, scrollViewportH } = layout;

    /** 纯数据单元格区域（不含行号列、列标题行）：此处永不自动滚动。 */
    const inDataCellArea = x > headerW && y > headerH && x <= w && y <= h;
    if (inDataCellArea) {
      return { sx: 0, sy: 0 };
    }

    const inColHead = y <= headerH && x >= 0 && x <= w;
    const inRowHead = x <= headerW && y >= 0 && y <= h;
    const outside = x < 0 || y < 0 || x > w || y > h;
    const canAutoscroll = inColHead || inRowHead || outside;

    let sx = 0;
    let sy = 0;

    /**
     * 触发区仅限：列标题带、行号列、画布外（与 hitTest 行列头一致）。
     * 条带深度与线性 t、MAX 不变，保证速度与平滑度。
     */
    if (scrollViewportW > 0 && limits.maxScrollX > 0 && canAutoscroll) {
      const wantRight = x >= w - M || x > w;
      const wantLeft = x <= headerW + M || x < 0;
      if (wantRight && (inColHead || x > w || y < 0 || y > h)) {
        const t = x > w ? 1 : Math.min(1, Math.max(0, (x - (w - M)) / M));
        if (scrollX < limits.maxScrollX - 1e-6) {
          sx = t * MAX;
        }
      } else if (
        wantLeft &&
        (inRowHead || (inColHead && x <= headerW + M) || x < 0 || y < 0 || y > h)
      ) {
        const t = x < 0 ? 1 : Math.min(1, Math.max(0, (headerW + M - x) / M));
        if (scrollX > 1e-6) {
          sx = -t * MAX;
        }
      }
    }

    if (scrollViewportH > 0 && limits.maxScrollY > 0 && canAutoscroll) {
      const wantDown = y >= h - M || y > h;
      const wantUp = y <= headerH + M || y < 0;
      if (wantDown && (inRowHead || y > h || x < 0 || x > w)) {
        const t = y > h ? 1 : Math.min(1, Math.max(0, (y - (h - M)) / M));
        if (scrollY < limits.maxScrollY - 1e-6) {
          sy = t * MAX;
        }
      } else if (
        wantUp &&
        (inColHead || (inRowHead && y <= headerH + M) || y < 0 || x < 0 || x > w)
      ) {
        const t = y < 0 ? 1 : Math.min(1, Math.max(0, (headerH + M - y) / M));
        if (scrollY > 1e-6) {
          sy = -t * MAX;
        }
      }
    }

    return { sx, sy };
  }

  private isDragAutoscrollActive(clientX: number, clientY: number): boolean {
    const { sx, sy } = this.computeDragAutoscrollPixelsPerSec(clientX, clientY);
    return Math.abs(sx) > 1e-6 || Math.abs(sy) > 1e-6;
  }

  private readonly dragAutoscrollLoop = (): void => {
    this.dragAutoscrollRafId = null;
    if (!this.dragSelecting && this.headingDrag === null && this.inlineFormulaRefDrag === null) {
      return;
    }

    const now = performance.now();
    const dt = Math.min(0.08, Math.max(0, (now - this.dragAutoscrollPrevNow) / 1000));
    this.dragAutoscrollPrevNow = now;

    const { sx, sy } = this.computeDragAutoscrollPixelsPerSec(
      this.lastDragClientX,
      this.lastDragClientY,
    );

    if (Math.abs(sx) > 1e-6 || Math.abs(sy) > 1e-6) {
      this.renderer.applyScrollDelta(sx * dt, sy * dt);
      if (this.headingDrag !== null) {
        this.applyHeadingDragSelectionFromClient(this.lastDragClientX, this.lastDragClientY);
      } else if (this.inlineFormulaRefDrag !== null) {
        const sh = this.workbook.getActiveSheet();
        if (sh !== undefined) {
          const hit = this.hitTestClient(this.lastDragClientX, this.lastDragClientY, {
            clampToBody: true,
          });
          if (hit !== null) {
            const a1 = sh.getMergeAnchorCell(hit.row, hit.col);
            this.inlineFormulaRefDrag.focusR = a1.row;
            this.inlineFormulaRefDrag.focusC = a1.col;
            this.updateInlineFormulaRefTextAndPreview();
          }
        }
      } else {
        const hit = this.hitTestClient(this.lastDragClientX, this.lastDragClientY, {
          clampToBody: true,
        });
        if (hit !== null) {
          this.selection.extendFocusTo(hit.row, hit.col);
        }
      }
      this.cellEditor.syncLayout();
      this.renderer.requestRedraw();
    }

    if (
      (this.dragSelecting || this.headingDrag !== null || this.inlineFormulaRefDrag !== null) &&
      this.isDragAutoscrollActive(this.lastDragClientX, this.lastDragClientY)
    ) {
      this.dragAutoscrollRafId = requestAnimationFrame(() => {
        this.dragAutoscrollLoop();
      });
    }
  };

  private syncSizeAndDraw(): void {
    const parent = this.canvas.parentElement;
    const w = parent?.clientWidth ?? this.canvas.clientWidth;
    const h = parent?.clientHeight ?? this.canvas.clientHeight;
    this.renderer.resize(w, h);
    this.cellEditor.syncLayout();
    this.renderer.cancelPendingRedraw();
    this.renderer.draw();
  }

  /**
   * 仅在内容包含换行时做行高自适应；列宽保持固定，不做自动扩列。
   */
  private autoExpandRowHeightForMultilineValue(
    sheet: Worksheet,
    row: number,
    col: number,
    value: string | number | boolean | null,
  ): void {
    if (typeof value !== "string" || !value.includes("\n")) {
      return;
    }
    const lines = value.split("\n").length;
    const fontCss = this.renderer.getCellEditorFontCss(row, col);
    const fontSize = parseFontPx(fontCss);
    const lineHeight = Math.max(16, Math.round(fontSize * 1.25));
    const target = Math.max(sheet.getRowHeight(row), lines * lineHeight + 4);
    if (target > sheet.getRowHeight(row) + 1e-6) {
      this.workspace.commands.execute(new SetRowHeightCommand(sheet, row, target));
    }
  }
}

/** 与表头拖拽调整行列尺寸一致的下上限（文档像素）。 */
function clampRowColSizePx(px: number): number {
  if (!Number.isFinite(px)) {
    return 8;
  }
  return Math.max(8, Math.min(4096, Math.round(px)));
}

export function createDefaultWorkbook(): Workbook {
  const wb = new Workbook();
  const sheet = new Worksheet("Sheet1", 200, 32);
  wb.addSheet(sheet);
  recalcWorksheet(sheet);
  return wb;
}

function darkenColor(input: string, ratio: number): string {
  const rgb = parseCssColor(input);
  if (rgb === null) {
    return "#8a8a8a";
  }
  const k = Math.max(0, Math.min(1, ratio));
  const r = Math.round(rgb.r * (1 - k));
  const g = Math.round(rgb.g * (1 - k));
  const b = Math.round(rgb.b * (1 - k));
  return `rgb(${r}, ${g}, ${b})`;
}

function parseCssColor(input: string): { r: number; g: number; b: number } | null {
  const s = input.trim();
  const hex = s.startsWith("#") ? s.slice(1) : s;
  if (/^[\dA-Fa-f]{6}$/.test(hex)) {
    return {
      r: parseInt(hex.slice(0, 2), 16),
      g: parseInt(hex.slice(2, 4), 16),
      b: parseInt(hex.slice(4, 6), 16),
    };
  }
  const m = s.match(/^rgba?\(([^)]+)\)$/i);
  if (m === null) {
    return null;
  }
  const parts = m[1].split(",").map((v) => Number(v.trim()));
  if (parts.length < 3 || parts.some((n) => Number.isNaN(n))) {
    return null;
  }
  return {
    r: clampByte(parts[0] ?? 0),
    g: clampByte(parts[1] ?? 0),
    b: clampByte(parts[2] ?? 0),
  };
}

function clampByte(v: number): number {
  return Math.max(0, Math.min(255, Math.round(v)));
}

function parseFontPx(fontCss: string): number {
  const m = fontCss.match(/(\d+(?:\.\d+)?)px/);
  if (m === null) {
    return 13;
  }
  return Math.max(8, Number(m[1]));
}

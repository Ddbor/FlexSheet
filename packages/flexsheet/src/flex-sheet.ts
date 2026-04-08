import {
  normalizeSelectionRange,
  PLUGIN_SERVICE_KEYS,
  Workbook,
  Workspace,
  Worksheet,
  type CellStyle,
  type CellStylePatch,
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
  hitTestHeadingPointer,
  buildFrozenLayout,
  computeColumnAutoWidth,
  computeRowAutoHeight,
  computeScrollLimits,
  type CanvasRenderer,
  type HeadingHit,
} from "@flexsheet/renderer";
import { ScrollPlugin } from "@flexsheet/scroll";
import { SelectionRegistryPlugin } from "@flexsheet/selection";
import { columnIndexToLabel } from "@flexsheet/shared";
import { createDefaultDarkTheme, createDefaultLightTheme, type SheetTheme } from "@flexsheet/theme";

import { runClipboardCopy, runClipboardCut, runClipboardPaste } from "./clipboard/clipboard-run.js";
import { useClipboard } from "./clipboard-plugin.js";
import {
  DeleteColsCommand,
  DeleteRowsCommand,
  InsertColsCommand,
  InsertCellsShiftDownCommand,
  InsertCellsShiftRightCommand,
  InsertRowsCommand,
  SetColHiddenCommand,
  SetColWidthCommand,
  SetRowHeightCommand,
  SetRowHiddenCommand,
} from "./sheet-structure-commands.js";
import {
  ApplySelectionBorderRibbonCommand,
  ApplySelectionCellStylePatchCommand,
  ApplySelectionFontSizeStepCommand,
  ApplySelectionIndentStepCommand,
  isRibbonBorderCommandId,
} from "./cell-style-commands.js";
import { SelectionMergeCommand } from "./merge-commands.js";
import { useSheetChromeGuard } from "./sheet-chrome-guard-plugin.js";
import { useSheetContextMenu } from "./sheet-context-menu-plugin.js";
import { useUndoRedo } from "./undo-redo-plugin.js";

/** 指针命中画布表面时的区域类型（供右键菜单等扩展使用）。 */
export type FlexSheetSurfaceHit =
  | { readonly kind: "corner" }
  | { readonly kind: "columnHeader"; readonly col: number }
  | { readonly kind: "rowHeader"; readonly row: number }
  | { readonly kind: "cell"; readonly row: number; readonly col: number };

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
  private workbookUnsub: (() => void) | null = null;
  private lastWorkbookActiveIndex = 0;
  private activeSheetFormattingUnsub: (() => void) | null = null;
  private readonly formattingChromeListeners = new Set<() => void>();
  /** 复制/剪切后的走马灯虚线框范围（与当前选区独立）。 */
  private clipboardMarqueeRange: SelectionRange | null = null;
  /** 延迟剪切：已写入剪贴板但源格尚未清空，粘贴匹配内部载荷后再清源区（与 Excel 一致）。 */
  private pendingClipboardCut: { sheet: Worksheet; range: SelectionRange } | null = null;

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
        return {
          range: this.selection.getNormalizedRange(),
          activeRow: cell.row,
          activeCol: cell.col,
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
      },
      onEditEnd: () => {
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
      this.workspace.use(useSheetChromeGuard({ chromeRoot: options.chromeRoot }));
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
    this.lastWorkbookActiveIndex = this._workbook.activeSheetIndex;
    this.workbookUnsub = this._workbook.subscribe(() => {
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
    this.rebindActiveSheetFormattingListener();
    this.renderer.requestRedraw();
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

  /** 是否处于单元格内联编辑态（供右键菜单等扩展判断）。 */
  isCellEditing(): boolean {
    return this.cellEditor.isEditing();
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
    this.renderer.requestRedraw();
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

  destroy(): void {
    this.cancelDragAutoscrollRaf();
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
    if (this.resizeObserver !== null) {
      this.resizeObserver.disconnect();
      this.resizeObserver = null;
    } else {
      window.removeEventListener("resize", this.onWindowResize);
    }
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
    if (!this.dragSelecting) {
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
      }
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
  }

  private bindSelectionKeyboard(): void {
    this.canvas.addEventListener("keydown", this.onKeyDown);
  }

  private readonly onLostPointerCapture = (ev: PointerEvent): void => {
    if ((!this.dragSelecting && !this.resizing) || ev.pointerId !== this.dragPointerId) {
      return;
    }
    this.finishDrag();
  };

  private readonly onCanvasPointerMove = (ev: PointerEvent): void => {
    if (this.resizing || this.dragSelecting) {
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
    if (this.resizing || this.dragSelecting) {
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

    const editorOpts: BeginEditOptions | undefined =
      options?.cursorClientX !== undefined
        ? { cursorClientX: options.cursorClientX }
        : options?.initialTextOverride !== undefined
          ? { selectAll: false }
          : options?.selectAll !== undefined
            ? { selectAll: options.selectAll }
            : undefined;
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
    this.canvas.focus();

    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }

    const headingHit = this.hitTestHeadingFromClient(ev.clientX, ev.clientY);
    if (headingHit !== null) {
      const resizeHit = this.tryHitResizeHandle(headingHit, ev.clientX, ev.clientY);
      if (resizeHit !== null) {
        ev.preventDefault();
        this.beginResizing(resizeHit.kind, resizeHit.index, ev);
        return;
      }
      ev.preventDefault();
      if (this.cellEditor.isEditing()) {
        this.cellEditor.cancelWithoutCommit();
      }
      const expand = ev.shiftKey || ev.ctrlKey || ev.metaKey;
      this.applyHeadingHitSelection(headingHit, sheet, expand);
      this.afterSelectionChanged();
      this.syncHeadingCursorFromClient(ev.clientX, ev.clientY);
      return;
    }

    const hit = this.hitTestClient(ev.clientX, ev.clientY);
    if (hit === null) {
      return;
    }
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
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
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
      this.clearClipboardMarquee();
      const cmd = new ClearRegionContentsCommand(sheet, this.selection.getNormalizedRange());
      if (cmd.hasChanges) {
        this.workspace.commands.execute(cmd);
      }
      this.renderer.requestRedraw();
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
      cell.formula !== null && cell.formula.length > 0 ? cell.formula : cellScalarToEditString(cell.value);
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
      cell.formula !== null && cell.formula.length > 0 ? cell.formula : cellScalarToEditString(cell.value);
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
      cell.formula !== null && cell.formula.length > 0 ? cell.formula : cellScalarToEditString(cell.value);
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
    if (!this.dragSelecting) {
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
      const hit = this.hitTestClient(this.lastDragClientX, this.lastDragClientY, {
        clampToBody: true,
      });
      if (hit !== null) {
        this.selection.extendFocusTo(hit.row, hit.col);
      }
      this.cellEditor.syncLayout();
      this.renderer.requestRedraw();
    }

    if (
      this.dragSelecting &&
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

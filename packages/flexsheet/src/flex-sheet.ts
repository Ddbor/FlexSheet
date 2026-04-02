import {
  PLUGIN_SERVICE_KEYS,
  Workbook,
  Workspace,
  Worksheet,
} from "@flexsheet/core";
import { recalcWorksheet, SetCellValueCommand } from "@flexsheet/formula";
import {
  CellEditor,
  EditorPlugin,
  cellScalarToEditString,
  type BeginEditOptions,
} from "@flexsheet/editor";
import { SelectionModel } from "@flexsheet/selection";
import {
  RendererPlugin,
  scrollToRevealCell,
  hitTestCell,
  hitTestHeadingPointer,
  buildFrozenLayout,
  computeScrollLimits,
  type CanvasRenderer,
  type HeadingHit,
} from "@flexsheet/renderer";
import { ScrollPlugin } from "@flexsheet/scroll";
import { SelectionRegistryPlugin } from "@flexsheet/selection";
import {
  createDefaultDarkTheme,
  createDefaultLightTheme,
  type SheetTheme,
} from "@flexsheet/theme";

import { useUndoRedo } from "./undo-redo-plugin.js";

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
  private formulaBarVisible = true;
  private resizeObserver: ResizeObserver | null = null;
  private dragSelecting = false;
  private dragPointerId: number | null = null;
  /** 框选拖拽时用于边沿自动滚动的上一帧 client 坐标。 */
  private lastDragClientX = 0;
  private lastDragClientY = 0;
  private dragAutoscrollRafId: number | null = null;
  private dragAutoscrollPrevNow = 0;
  private workbookUnsub: (() => void) | null = null;
  private lastWorkbookActiveIndex = 0;

  constructor(options: FlexSheetOptions) {
    this.formulaBarEl = options.formulaBar ?? null;
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
      },
      onEditEnd: () => {
        queueMicrotask(() => {
          this.canvas.focus();
        });
      },
    });
    this.workspace.use(editorPlugin);
    this.cellEditor = editorPlugin.getCellEditor();

    this.workspace.use(useUndoRedo({ canvas: this.canvas }));

    this.workspace.use(new ScrollPlugin());

    this.workspace.pluginContext.register(PLUGIN_SERVICE_KEYS.flexSheet, this);

    this.bindResize();
    this.bindSelectionPointer();
    this.bindSelectionKeyboard();
    this.attachWorkbookForDataDrive();
    this.syncSizeAndDraw();
  }

  private attachWorkbookForDataDrive(): void {
    this.workbookUnsub?.();
    this.lastWorkbookActiveIndex = this._workbook.activeSheetIndex;
    this.workbookUnsub = this._workbook.subscribe(() => {
      const cur = this._workbook.activeSheetIndex;
      if (cur !== this.lastWorkbookActiveIndex) {
        this.lastWorkbookActiveIndex = cur;
        if (this.cellEditor.isEditing()) {
          this.cellEditor.cancelWithoutCommit();
        }
        this.selection.syncWithSheet();
        this.renderer.ensureScrollClamped();
        this.cellEditor.syncLayout();
      }
      this.renderer.requestRedraw();
    });
  }

  loadWorkbook(wb: Workbook): void {
    this.workbookUnsub?.();
    this._workbook = wb;
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

  refresh(): void {
    this.cellEditor.syncLayout();
    this.renderer.requestRedraw();
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

  destroy(): void {
    this.cancelDragAutoscrollRaf();
    this.detachDocumentDragListeners();
    this.renderer.cancelPendingRedraw();
    this.workbookUnsub?.();
    this.workbookUnsub = null;
    this.canvas.removeEventListener("pointerdown", this.onPointerDown);
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
    if (!this.dragSelecting || ev.pointerId !== this.dragPointerId) {
      return;
    }
    this.finishDrag();
  };

  private readonly onDocumentPointerMoveCapture = (ev: PointerEvent): void => {
    if (!this.dragSelecting || ev.pointerId !== this.dragPointerId) {
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
    if (!this.dragSelecting || ev.pointerId !== this.dragPointerId) {
      return;
    }
    this.finishDrag(ev);
  };

  private readonly onDocumentPointerCancelCapture = (ev: PointerEvent): void => {
    if (!this.dragSelecting || ev.pointerId !== this.dragPointerId) {
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
    const rect = this.renderer.getCellRectInCanvasPixels(row, col);
    if (rect === null) {
      return;
    }
    const cell = sheet.getCell(row, col);
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
    this.cellEditor.beginEdit(row, col, text, rect, editorOpts);
  }

  /** 双击按位置进入编辑；单击仅选格（与 Excel 一致，键入 / F2 再进编辑）。 */
  private readonly onCanvasDoubleClick = (ev: MouseEvent): void => {
    ev.preventDefault();
    if (this.cellEditor.isEditing()) {
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
      ev.preventDefault();
      if (this.cellEditor.isEditing()) {
        this.cellEditor.cancelWithoutCommit();
      }
      const expand = ev.shiftKey || ev.ctrlKey || ev.metaKey;
      this.applyHeadingHitSelection(headingHit, sheet, expand);
      this.afterSelectionChanged();
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
      const ac = this.selection.getActiveCell();
      const cmd = new SetCellValueCommand(sheet, ac.row, ac.col, null);
      this.workspace.commands.execute(cmd);
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
  private computeDragAutoscrollPixelsPerSec(clientX: number, clientY: number): {
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
}

export function createDefaultWorkbook(): Workbook {
  const wb = new Workbook();
  const sheet = new Worksheet("Sheet1", 200, 32);
  wb.addSheet(sheet);
  recalcWorksheet(sheet);
  return wb;
}

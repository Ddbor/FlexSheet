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
  buildFrozenLayout,
  computeScrollLimits,
  type CanvasRenderer,
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
  private workbookUnsub: (() => void) | null = null;

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
    this.workbookUnsub = this._workbook.subscribe(() => {
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
    const hit = this.hitTestClient(ev.clientX, ev.clientY);
    if (hit === null) {
      return;
    }
    this.dragSelecting = true;
    this.dragPointerId = ev.pointerId;
    this.attachDocumentDragListeners();
    this.selection.selectCell(hit.row, hit.col);
    try {
      this.canvas.setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
    this.afterSelectionChanged();
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

  private afterSelectionChanged(): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      this.renderer.requestRedraw();
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
    this.cellEditor.syncLayout();
    this.renderer.requestRedraw();
  }

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

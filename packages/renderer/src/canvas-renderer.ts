import type { CellScalar } from "@flexsheet/core";
import type { Workbook } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import {
  normalizeSelectionRange,
  type SelectionPaintSnapshot,
  type SelectionRange,
} from "@flexsheet/core";
import type { SheetTheme } from "@flexsheet/theme";
import { columnIndexToLabel } from "@flexsheet/shared";
import { scrollToRevealCell } from "./grid-hit-test.js";
import {
  buildFrozenLayout,
  clampScroll,
  computeScrollLimits,
  visibleScrollableCellRange,
  type FrozenLayout,
} from "./viewport.js";

const VIEW_ZOOM_MIN = 0.25;
const VIEW_ZOOM_MAX = 4;

/** 与画布横向滚动区一致的度量，供底部栏等宿主同步假滚动条。 */
export interface HorizontalScrollMetrics {
  readonly scrollX: number;
  readonly maxScrollX: number;
  readonly scrollViewportW: number;
  readonly contentScrollWidth: number;
}

/** 与画布纵向滚动区一致的度量，供右侧假滚动条等宿主同步。 */
export interface VerticalScrollMetrics {
  readonly scrollY: number;
  readonly maxScrollY: number;
  readonly scrollViewportH: number;
  readonly contentScrollHeight: number;
}

export interface CanvasRendererOptions {
  canvas: HTMLCanvasElement;
  workbook: Workbook;
  theme: SheetTheme;
  /** 选区绘制数据源；缺省则不绘选区。 */
  getSelectionSnapshot?: () => SelectionPaintSnapshot | null;
}

/** OOXML 风格 8 位 ARGB → CSS 颜色（含 alpha）。 */
function argbToCss(argb: string): string | undefined {
  const s = argb.trim();
  if (!/^[\dA-Fa-f]{8}$/.test(s)) {
    return undefined;
  }
  const a = parseInt(s.slice(0, 2), 16) / 255;
  const R = s.slice(2, 4);
  const G = s.slice(4, 6);
  const B = s.slice(6, 8);
  const r = parseInt(R, 16);
  const g = parseInt(G, 16);
  const b = parseInt(B, 16);
  if (a >= 1 - 1e-6) {
    return `#${R}${G}${B}`;
  }
  return `rgba(${r},${g},${b},${a})`;
}

/**
 * 全表统一 1px 描边坐标：对齐到 CSS 像素中心，与 Excel 类表格一致。
 * 列头竖线、行头横线、表体网格、冻结线、标题分割线均须用此函数，避免缝与双线。
 */
function snapLine(v: number): number {
  return Math.round(v) + 0.5;
}

/**
 * 选区外框、填充柄相对「逻辑 2px 描边」的视觉比例（略细，接近 Excel）。
 */
const SELECTION_OUTLINE_VISUAL_SCALE = 2 / 3;

/** zoom=1 时行号列默认宽度（CSS px），与表体同乘 viewZoom。 */
const HEADER_STRIP_BASE_WIDTH = 40;

/** zoom=1 时列标区域高度（CSS px），与表体同乘 viewZoom。 */
const HEADER_STRIP_BASE_HEIGHT = 24;

/**
 * 激活单元格/选区描边与行列标题强调带共用线宽（CSS 像素），随 viewZoom 线性缩放，保证视觉一致。
 */
function viewScaledSelectionOutlineWidth(viewZoom: number): number {
  const z = Math.max(VIEW_ZOOM_MIN, Math.min(VIEW_ZOOM_MAX, viewZoom));
  return Math.max(1, 2 * SELECTION_OUTLINE_VISUAL_SCALE * z);
}

/**
 * Canvas 2D 主渲染器：行列标题、网格线、冻结窗格与视口虚拟滚动（只读数据）。
 */
export class CanvasRenderer {
  private readonly canvas: HTMLCanvasElement;
  private readonly ctx: CanvasRenderingContext2D;
  private workbook: Workbook;
  private theme: SheetTheme;
  private getSelectionSnapshot: () => SelectionPaintSnapshot | null;

  /** 仅作用于非冻结列/行（文档像素）。 */
  scrollX = 0;
  scrollY = 0;

  frozenRows = 0;
  frozenCols = 0;

  /** 可见区外多绘制的行列数，减少快速滚动白边。 */
  viewportBuffer = 2;

  /** 视图缩放（行列像素 = 默认宽高 × viewZoom）。 */
  viewZoom = 1;

  showGridLines = true;

  showHeadings = true;

  /** 与 `showHeadings` 同时为 true 时，在列标题区底部绘制刻度线（主题 `headerLineColor`）。 */
  showRuler = false;

  /** 宏录制状态（Ribbon 同步；业务录制逻辑由宿主扩展）。 */
  macroRecording = false;

  /** 宏是否使用相对引用。 */
  macroUseRelativeReference = false;

  private rafId: number | null = null;

  private readonly viewZoomListeners = new Set<() => void>();

  private readonly scrollListeners = new Set<() => void>();

  constructor(options: CanvasRendererOptions) {
    this.canvas = options.canvas;
    const ctx = options.canvas.getContext("2d");
    if (ctx === null) {
      throw new Error("CanvasRenderer: 无法获取 2D 上下文");
    }
    this.ctx = ctx;
    this.workbook = options.workbook;
    this.theme = options.theme;
    this.getSelectionSnapshot = options.getSelectionSnapshot ?? (() => null);
  }

  setSelectionSnapshotProvider(fn: (() => SelectionPaintSnapshot | null) | undefined): void {
    this.getSelectionSnapshot = fn ?? (() => null);
  }

  /**
   * 行列标题区尺寸（隐藏标题时为 0，表体从画布左上角起算）。
   * 与单元格 defaultColWidth/defaultRowHeight 一样乘以 viewZoom，缩放时与网格同步。
   */
  getHeaderSize(): Readonly<{ width: number; height: number }> {
    if (!this.showHeadings) {
      return { width: 0, height: 0 };
    }
    const z = this.viewZoom;
    return {
      width: HEADER_STRIP_BASE_WIDTH * z,
      height: HEADER_STRIP_BASE_HEIGHT * z,
    };
  }

  getCornerSize(): Readonly<{ width: number; height: number }> {
    return this.getHeaderSize();
  }

  getViewZoom(): number {
    return this.viewZoom;
  }

  setViewZoom(scale: number): void {
    const z = Math.max(VIEW_ZOOM_MIN, Math.min(VIEW_ZOOM_MAX, scale));
    if (z === this.viewZoom) {
      return;
    }
    this.viewZoom = z;
    this.ensureScrollClamped();
    this.notifyViewZoomChanged();
  }

  /** 缩放比例变化时通知（含 Ribbon、状态栏等宿主 UI）。 */
  subscribeViewZoom(listener: () => void): () => void {
    this.viewZoomListeners.add(listener);
    return () => {
      this.viewZoomListeners.delete(listener);
    };
  }

  private notifyViewZoomChanged(): void {
    for (const fn of this.viewZoomListeners) {
      fn();
    }
  }

  /**
   * 横向滚动或视口夹紧后通知（滚轮、拖拽假滚动条、`setScroll` 等）。
   */
  subscribeScroll(listener: () => void): () => void {
    this.scrollListeners.add(listener);
    return () => {
      this.scrollListeners.delete(listener);
    };
  }

  private notifyScrollChanged(): void {
    for (const fn of this.scrollListeners) {
      fn();
    }
  }

  /**
   * 当前活动表下与 `scrollX` 一致的横向滚动度量；无表时返回 `null`。
   */
  getHorizontalScrollMetrics(): HorizontalScrollMetrics | null {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return null;
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const { width: hw, height: hh } = this.getHeaderSize();
    const layout = buildFrozenLayout(
      sheet,
      hw,
      hh,
      w,
      h,
      this.frozenRows,
      this.frozenCols,
      this.viewZoom,
    );
    const limits = computeScrollLimits(sheet, layout, this.viewZoom);
    const colW = sheet.defaultColWidth * this.viewZoom;
    const scrollColCount = Math.max(0, sheet.colCount - layout.frozenCols);
    const contentScrollWidth = scrollColCount * colW;
    return {
      scrollX: this.scrollX,
      maxScrollX: limits.maxScrollX,
      scrollViewportW: layout.scrollViewportW,
      contentScrollWidth,
    };
  }

  /**
   * 当前活动表下与 `scrollY` 一致的纵向滚动度量；无表时返回 `null`。
   */
  getVerticalScrollMetrics(): VerticalScrollMetrics | null {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return null;
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const { width: hw, height: hh } = this.getHeaderSize();
    const layout = buildFrozenLayout(
      sheet,
      hw,
      hh,
      w,
      h,
      this.frozenRows,
      this.frozenCols,
      this.viewZoom,
    );
    const limits = computeScrollLimits(sheet, layout, this.viewZoom);
    const rowH = sheet.defaultRowHeight * this.viewZoom;
    const scrollRowCount = Math.max(0, sheet.rowCount - layout.frozenRows);
    const contentScrollHeight = scrollRowCount * rowH;
    return {
      scrollY: this.scrollY,
      maxScrollY: limits.maxScrollY,
      scrollViewportH: layout.scrollViewportH,
      contentScrollHeight,
    };
  }

  zoomIn(): void {
    this.setViewZoom(this.viewZoom * 1.1);
  }

  zoomOut(): void {
    this.setViewZoom(this.viewZoom / 1.1);
  }

  resetZoom100(): void {
    this.setViewZoom(1);
  }

  /**
   * 将缩放设为使当前选区矩形在可滚动区内尽可能完整显示（迭代收敛），并滚动到合适位置。
   */
  zoomToFitRange(range: SelectionRange): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const n = normalizeSelectionRange(range);
    const r0 = n.startRow;
    const r1 = n.endRow;
    const c0 = n.startCol;
    const c1 = n.endCol;
    const cols = c1 - c0 + 1;
    const rows = r1 - r0 + 1;
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const { width: hw, height: hh } = this.getHeaderSize();
    const fr = this.frozenRows;
    const fc = this.frozenCols;
    const dcw = sheet.defaultColWidth;
    const drh = sheet.defaultRowHeight;
    let z = this.viewZoom;
    for (let i = 0; i < 6; i++) {
      const prev = z;
      const layoutIter = buildFrozenLayout(sheet, hw, hh, w, h, fr, fc, z);
      const vpW = layoutIter.scrollViewportW;
      const vpH = layoutIter.scrollViewportH;
      const zw = cols > 0 && dcw > 0 ? vpW / (cols * dcw) : z;
      const zh = rows > 0 && drh > 0 ? vpH / (rows * drh) : z;
      const next = Math.min(zw, zh) * 0.92;
      z = Math.max(VIEW_ZOOM_MIN, Math.min(VIEW_ZOOM_MAX, next));
      if (Math.abs(z - prev) < 1e-5) {
        break;
      }
    }
    const prevZ = this.viewZoom;
    this.viewZoom = z;
    this.ensureScrollClamped();
    if (Math.abs(this.viewZoom - prevZ) > 1e-6) {
      this.notifyViewZoomChanged();
    }
    const layout = buildFrozenLayout(sheet, hw, hh, w, h, fr, fc, this.viewZoom);
    const limits = computeScrollLimits(sheet, layout, this.viewZoom);
    const revealed = scrollToRevealCell(
      sheet,
      layout,
      limits,
      r0,
      c0,
      this.scrollX,
      this.scrollY,
      this.viewZoom,
    );
    this.scrollX = revealed.scrollX;
    this.scrollY = revealed.scrollY;
    this.notifyScrollChanged();
  }

  setShowGridLines(show: boolean): void {
    this.showGridLines = show;
  }

  setShowHeadings(show: boolean): void {
    this.showHeadings = show;
    this.ensureScrollClamped();
  }

  setShowRuler(show: boolean): void {
    this.showRuler = show;
  }

  setMacroRecording(on: boolean): void {
    this.macroRecording = on;
  }

  setMacroUseRelativeReference(on: boolean): void {
    this.macroUseRelativeReference = on;
  }

  getFrozenPanes(): { readonly frozenRows: number; readonly frozenCols: number } {
    return { frozenRows: this.frozenRows, frozenCols: this.frozenCols };
  }

  private scaledColW(sheet: Worksheet): number {
    return sheet.defaultColWidth * this.viewZoom;
  }

  private scaledRowH(sheet: Worksheet): number {
    return sheet.defaultRowHeight * this.viewZoom;
  }

  private scaledFontSizePx(base: number): number {
    return Math.max(8, Math.min(36, Math.round(base * this.viewZoom)));
  }

  /**
   * 单元格在 canvas 元素**客户端坐标系**中的矩形（与绘制用的 cellLeftX/cellTopY 一致，CSS 像素）。
   */
  /**
   * 与 `paintBodyCellTexts` 中单元格正文一致，用于浮层编辑框字体与测量（含缩放与加粗）。
   */
  getCellEditorFontCss(row: number, col: number): string {
    const sheet = this.workbook.getActiveSheet();
    let weight = "400";
    if (sheet !== undefined) {
      const c = sheet.getCell(row, col);
      if (c.style?.bold === true) {
        weight = "600";
      }
    }
    return `${weight} ${this.scaledFontSizePx(13)}px system-ui, -apple-system, sans-serif`;
  }

  getCellRectInCanvasPixels(
    row: number,
    col: number,
  ): { x: number; y: number; width: number; height: number } | null {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return null;
    }
    if (row < 0 || col < 0 || row >= sheet.rowCount || col >= sheet.colCount) {
      return null;
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const { width: hw, height: hh } = this.getHeaderSize();
    const layout = buildFrozenLayout(
      sheet,
      hw,
      hh,
      w,
      h,
      this.frozenRows,
      this.frozenCols,
      this.viewZoom,
    );
    return {
      x: this.cellLeftX(sheet, layout, col),
      y: this.cellTopY(sheet, layout, row),
      width: this.scaledColW(sheet),
      height: this.scaledRowH(sheet),
    };
  }

  setWorkbook(workbook: Workbook): void {
    this.workbook = workbook;
  }

  setTheme(theme: SheetTheme): void {
    this.theme = theme;
  }

  setFrozenPanes(frozenRows: number, frozenCols: number): void {
    this.frozenRows = Math.max(0, frozenRows);
    this.frozenCols = Math.max(0, frozenCols);
  }

  getScroll(): { scrollX: number; scrollY: number } {
    return { scrollX: this.scrollX, scrollY: this.scrollY };
  }

  setScroll(scrollX: number, scrollY: number): void {
    this.scrollX = scrollX;
    this.scrollY = scrollY;
    this.notifyScrollChanged();
  }

  /**
   * 按当前活动表与画布尺寸夹紧 scroll（导入工作簿或表尺寸变化后调用，避免下一帧绘制前 hitTest 与 scroll 不一致）。
   */
  ensureScrollClamped(): void {
    this.clampScrollToActiveSheet();
  }

  applyScrollDelta(deltaX: number, deltaY: number): void {
    this.scrollX += deltaX;
    this.scrollY += deltaY;
    this.clampScrollToActiveSheet();
  }

  /** 与 `paint` 使用相同 limits，保证滚轮后、下一帧绘制前 hitTest / 编辑器定位与视口一致。 */
  private clampScrollToActiveSheet(): void {
    const sheet = this.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;
    const { width: hw, height: hh } = this.getHeaderSize();
    const layout = buildFrozenLayout(
      sheet,
      hw,
      hh,
      w,
      h,
      this.frozenRows,
      this.frozenCols,
      this.viewZoom,
    );
    const limits = computeScrollLimits(sheet, layout, this.viewZoom);
    const c = clampScroll(this.scrollX, this.scrollY, limits);
    this.scrollX = c.scrollX;
    this.scrollY = c.scrollY;
    this.notifyScrollChanged();
  }

  resize(cssWidth: number, cssHeight: number): void {
    const dpr = window.devicePixelRatio || 1;
    this.canvas.style.width = `${cssWidth}px`;
    this.canvas.style.height = `${cssHeight}px`;
    this.canvas.width = Math.max(1, Math.floor(cssWidth * dpr));
    this.canvas.height = Math.max(1, Math.floor(cssHeight * dpr));
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    this.clampScrollToActiveSheet();
  }

  /** 合并到下一帧绘制，减少同帧多次改数据导致的闪烁。 */
  requestRedraw(): void {
    if (typeof requestAnimationFrame === "undefined") {
      this.paint();
      return;
    }
    if (this.rafId !== null) {
      return;
    }
    this.rafId = requestAnimationFrame(() => {
      this.rafId = null;
      this.paint();
    });
  }

  /** 取消尚未执行的帧（如组件销毁时）。 */
  cancelPendingRedraw(): void {
    if (this.rafId !== null) {
      cancelAnimationFrame(this.rafId);
      this.rafId = null;
    }
  }

  /** 立即全量绘制（与 `requestRedraw` 相对）。 */
  draw(): void {
    this.paint();
  }

  private paint(): void {
    const sheet = this.workbook.getActiveSheet();
    const { ctx } = this;
    const w = this.canvas.clientWidth;
    const h = this.canvas.clientHeight;

    ctx.save();
    ctx.globalAlpha = 1;
    ctx.globalCompositeOperation = "source-over";
    ctx.imageSmoothingEnabled = true;
    ctx.lineCap = "butt";
    ctx.lineJoin = "miter";
    ctx.lineWidth = 1;

    ctx.fillStyle = this.theme.canvasBg;
    ctx.fillRect(0, 0, w, h);

    if (sheet === undefined) {
      ctx.fillStyle = this.theme.cellColor;
      ctx.font = "14px system-ui, sans-serif";
      ctx.textAlign = "left";
      ctx.textBaseline = "alphabetic";
      ctx.fillText("无工作表", 16, 32);
      ctx.restore();
      return;
    }

    const headerW = this.getHeaderSize().width;
    const headerH = this.getHeaderSize().height;
    const layout = buildFrozenLayout(
      sheet,
      headerW,
      headerH,
      w,
      h,
      this.frozenRows,
      this.frozenCols,
      this.viewZoom,
    );
    const limits = computeScrollLimits(sheet, layout, this.viewZoom);
    const clamped = clampScroll(this.scrollX, this.scrollY, limits);
    this.scrollX = clamped.scrollX;
    this.scrollY = clamped.scrollY;

    const selectionSnap = this.getSelectionSnapshot();

    if (this.showHeadings) {
      this.drawCorner(headerW, headerH);
      if (this.showRuler) {
        this.drawRulerOverlay(sheet, layout, headerW, headerH, w);
      }
      this.drawAllColumnHeaders(sheet, layout, headerW, headerH, w, selectionSnap);
      this.drawAllRowHeaders(sheet, layout, headerW, headerH, h, selectionSnap);
    }
    this.drawBody(sheet, layout, headerW, headerH, w, h);
    if (this.showHeadings) {
      this.drawHeadingBodyDividerLines(headerW, headerH, w, h);
      /** 列底/行右强调线在分割线之上；与表体选区邻接时省略对应边，避免与选区描边叠粗。 */
      this.paintAllHeaderSelectionAccents(sheet, layout, headerW, headerH, w, h, selectionSnap);
    }
    this.drawFreezeLines(layout, w, h);
    this.drawSelectionOverlay(sheet, layout, headerW, headerH, w, h);

    ctx.restore();
  }

  private drawCorner(headerW: number, headerH: number): void {
    const { ctx } = this;
    ctx.fillStyle = this.theme.headerBg;
    ctx.fillRect(0, 0, headerW, headerH);
    ctx.strokeStyle = this.theme.headerLineColor;
    ctx.lineWidth = 1;
    /** 只画与画布外缘相接的左、上两边；右、下与列头/行头网格共用同一条 snapLine，避免错位与加粗。 */
    ctx.beginPath();
    const xt = snapLine(0);
    const yt = snapLine(0);
    const xr = snapLine(headerW);
    const yb = snapLine(headerH);
    ctx.moveTo(xt, yt);
    ctx.lineTo(xr, yt);
    ctx.moveTo(xt, yt);
    ctx.lineTo(xt, yb);
    ctx.stroke();

    /** 右下角直角三角标（全选夹角），与 Excel 类似。 */
    const inset = 4;
    const leg = Math.max(5, Math.min(headerW, headerH) * 0.36);
    const brx = headerW - inset;
    const bry = headerH - inset;
    ctx.beginPath();
    ctx.moveTo(brx - leg, bry);
    ctx.lineTo(brx, bry);
    ctx.lineTo(brx, bry - leg);
    ctx.closePath();
    ctx.fillStyle = this.theme.headerColor;
    ctx.globalAlpha = 0.38;
    ctx.fill();
    ctx.globalAlpha = 1;
    ctx.strokeStyle = this.theme.headerLineColor;
    ctx.lineWidth = 1;
    ctx.stroke();
  }

  /** 列标题区域底部刻度；水平基线与表体网格同色，由 `drawHeadingBodyDividerLines` 在表体绘制之后统一画出，此处只画向上刻度。 */
  private drawRulerOverlay(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
  ): void {
    const { ctx } = this;
    const colW = this.scaledColW(sheet);
    const yRuler = snapLine(headerH);
    ctx.strokeStyle = this.theme.gridLineColor;
    ctx.lineWidth = 1;
    const { frozenCols } = layout;
    const maxC = sheet.colCount - 1;
    const fw = layout.frozenWidthPx;
    const sx0 = headerW + fw;
    const tickH = 4;
    const drawTicks = (c0: number, c1: number): void => {
      for (let c = c0; c <= c1; c++) {
        const x = this.cellLeftX(sheet, layout, c);
        if (x + colW <= headerW || x >= canvasW) {
          continue;
        }
        ctx.beginPath();
        ctx.moveTo(snapLine(x), yRuler);
        ctx.lineTo(snapLine(x), yRuler - tickH);
        ctx.stroke();
      }
    };
    if (frozenCols > 0) {
      drawTicks(0, Math.min(frozenCols - 1, maxC));
    }
    if (layout.scrollViewportW > 0 && frozenCols <= maxC) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(sx0, 0, canvasW - sx0, headerH);
      ctx.clip();
      const first = frozenCols + Math.floor(this.scrollX / colW) - 1;
      const last = frozenCols + Math.ceil((this.scrollX + layout.scrollViewportW) / colW) + 1;
      drawTicks(Math.max(frozenCols, first), Math.min(maxC, last));
      ctx.restore();
    }
  }

  /**
   * 当前选区在表内的行列闭区间（与框选矩形一致）；无快照时返回 null。
   */
  private getClampedSelectionSpan(
    sheet: Worksheet,
    selectionSnap: SelectionPaintSnapshot | null,
  ): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
    if (selectionSnap === null) {
      return null;
    }
    const n = normalizeSelectionRange(selectionSnap.range);
    const maxC = sheet.colCount - 1;
    const maxR = sheet.rowCount - 1;
    if (maxC < 0 || maxR < 0) {
      return null;
    }
    return {
      startCol: Math.max(0, Math.min(n.startCol, maxC)),
      endCol: Math.max(0, Math.min(n.endCol, maxC)),
      startRow: Math.max(0, Math.min(n.startRow, maxR)),
      endRow: Math.max(0, Math.min(n.endRow, maxR)),
    };
  }

  private drawAllColumnHeaders(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
    selectionSnap: SelectionPaintSnapshot | null,
  ): void {
    const { ctx } = this;
    const colW = this.scaledColW(sheet);
    const selSpan = this.getClampedSelectionSpan(sheet, selectionSnap);
    const columnInSelection = (c: number): boolean =>
      selSpan !== null && c >= selSpan.startCol && c <= selSpan.endCol;
    const { frozenCols } = layout;
    const buf = this.viewportBuffer;
    const maxC = sheet.colCount - 1;
    const fw = layout.frozenWidthPx;
    const sx0 = headerW + fw;

    const collectFrozenCols = (): number[] => {
      const out: number[] = [];
      for (let c = 0; c < frozenCols && c <= maxC; c++) {
        const x = this.cellLeftX(sheet, layout, c);
        if (x + colW <= headerW || x >= canvasW) {
          continue;
        }
        out.push(c);
      }
      return out;
    };

    const collectScrollCols = (): number[] => {
      const out: number[] = [];
      if (frozenCols > maxC || layout.scrollViewportW <= 0) {
        return out;
      }
      const colW2 = this.scaledColW(sheet);
      const first = frozenCols + Math.floor(this.scrollX / colW2) - buf;
      const last =
        frozenCols + Math.ceil((this.scrollX + layout.scrollViewportW) / colW2) - 1 + buf;
      const c0 = Math.max(frozenCols, first);
      const c1 = Math.min(maxC, last);
      for (let c = c0; c <= c1; c++) {
        const x = this.cellLeftX(sheet, layout, c);
        if (x + colW <= headerW || x >= canvasW) {
          continue;
        }
        out.push(c);
      }
      return out;
    };

    const strokeColumnHeaderGrid = (visibleCols: number[], clipLeft: number): void => {
      if (visibleCols.length === 0) {
        return;
      }
      const yTop = snapLine(0);
      const yBot = snapLine(headerH);
      const xClip = snapLine(clipLeft);
      ctx.strokeStyle = this.theme.headerLineColor;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(xClip, yTop);
      ctx.lineTo(xClip, yBot);
      ctx.moveTo(xClip, yTop);
      ctx.lineTo(snapLine(canvasW), yTop);
      /** 底边与表体分隔线由 `drawHeadingBodyDividerLines` 用 gridLineColor 绘制，避免双线。 */
      for (const c of visibleCols) {
        const x = this.cellLeftX(sheet, layout, c);
        const xs = snapLine(x);
        ctx.moveTo(xs, yTop);
        ctx.lineTo(xs, yBot);
      }
      const lastC = visibleCols[visibleCols.length - 1];
      const xe = snapLine(this.cellLeftX(sheet, layout, lastC) + colW);
      ctx.moveTo(xe, yTop);
      ctx.lineTo(xe, yBot);
      ctx.stroke();
    };

    /** 先滚动区再冻结区，避免滚动列标题盖住冻结列标题（同索引排序时后者会叠在上层）。 */
    const scrollCols = collectScrollCols();
    if (scrollCols.length > 0 && frozenCols <= maxC && layout.scrollViewportW > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(sx0, 0, canvasW - sx0, headerH);
      ctx.clip();
      for (const c of scrollCols) {
        const x = this.cellLeftX(sheet, layout, c);
        this.paintColumnHeaderCell(ctx, c, x, colW, headerH, columnInSelection(c));
      }
      strokeColumnHeaderGrid(scrollCols, sx0);
      ctx.restore();
    }

    const frozenColsList = collectFrozenCols();
    if (frozenColsList.length > 0 && frozenCols > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(headerW, 0, fw, headerH);
      ctx.clip();
      for (const c of frozenColsList) {
        const x = this.cellLeftX(sheet, layout, c);
        this.paintColumnHeaderCell(ctx, c, x, colW, headerH, columnInSelection(c));
      }
      strokeColumnHeaderGrid(frozenColsList, headerW);
      ctx.restore();
    }
  }

  private paintColumnHeaderCell(
    ctx: CanvasRenderingContext2D,
    c: number,
    x: number,
    colW: number,
    headerH: number,
    isSelectionColumn: boolean,
  ): void {
    ctx.fillStyle = isSelectionColumn ? this.theme.headerActiveBg : this.theme.headerBg;
    ctx.fillRect(x, 0, colW, headerH);
    ctx.fillStyle = isSelectionColumn ? this.theme.activeCellBorderColor : this.theme.headerColor;
    ctx.font = `${isSelectionColumn ? "bold " : ""}${this.scaledFontSizePx(12)}px system-ui, -apple-system, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(columnIndexToLabel(c), x + colW / 2, headerH / 2);
  }

  private drawAllRowHeaders(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasH: number,
    selectionSnap: SelectionPaintSnapshot | null,
  ): void {
    const { ctx } = this;
    const rowH = this.scaledRowH(sheet);
    const { frozenRows } = layout;
    const buf = this.viewportBuffer;
    const maxR = sheet.rowCount - 1;
    const fh = layout.frozenHeightPx;
    const sy0 = headerH + fh;
    const selSpanRows = this.getClampedSelectionSpan(sheet, selectionSnap);
    const rowInSelection = (r: number): boolean =>
      selSpanRows !== null && r >= selSpanRows.startRow && r <= selSpanRows.endRow;

    const collectFrozenRows = (): number[] => {
      const out: number[] = [];
      for (let r = 0; r < frozenRows && r <= maxR; r++) {
        const y = this.cellTopY(sheet, layout, r);
        if (y + rowH <= headerH || y >= canvasH) {
          continue;
        }
        out.push(r);
      }
      return out;
    };

    const collectScrollRows = (): number[] => {
      const out: number[] = [];
      if (frozenRows > maxR || layout.scrollViewportH <= 0) {
        return out;
      }
      const rowH2 = this.scaledRowH(sheet);
      const first = frozenRows + Math.floor(this.scrollY / rowH2) - buf;
      const last =
        frozenRows + Math.ceil((this.scrollY + layout.scrollViewportH) / rowH2) - 1 + buf;
      const r0 = Math.max(frozenRows, first);
      const r1 = Math.min(maxR, last);
      for (let r = r0; r <= r1; r++) {
        const y = this.cellTopY(sheet, layout, r);
        if (y + rowH <= headerH || y >= canvasH) {
          continue;
        }
        out.push(r);
      }
      return out;
    };

    const strokeRowHeaderGrid = (visibleRows: number[], clipTop: number): void => {
      if (visibleRows.length === 0) {
        return;
      }
      const xLeft = snapLine(0);
      const xRight = snapLine(headerW);
      const yClip = snapLine(clipTop);
      const yEnd = snapLine(canvasH);
      ctx.strokeStyle = this.theme.headerLineColor;
      ctx.lineWidth = 1;
      ctx.beginPath();
      if (clipTop !== headerH) {
        ctx.moveTo(xLeft, yClip);
        ctx.lineTo(xRight, yClip);
      }
      ctx.moveTo(xLeft, yClip);
      ctx.lineTo(xLeft, yEnd);
      /** 行号区右侧与表体分隔线由 `drawHeadingBodyDividerLines` 用 gridLineColor 绘制，避免双线。 */
      for (const r of visibleRows) {
        const y = this.cellTopY(sheet, layout, r);
        const ys = snapLine(y);
        ctx.moveTo(xLeft, ys);
        ctx.lineTo(xRight, ys);
      }
      const lastR = visibleRows[visibleRows.length - 1];
      const ye = snapLine(this.cellTopY(sheet, layout, lastR) + rowH);
      ctx.moveTo(xLeft, ye);
      ctx.lineTo(xRight, ye);
      ctx.stroke();
    };

    /** 先滚动区再冻结区，避免滚动行标题盖住冻结行标题。 */
    const scrollRows = collectScrollRows();
    if (scrollRows.length > 0 && frozenRows <= maxR && layout.scrollViewportH > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, sy0, headerW, canvasH - sy0);
      ctx.clip();
      for (const r of scrollRows) {
        const y = this.cellTopY(sheet, layout, r);
        this.paintRowHeaderCell(ctx, r, y, rowH, headerW, rowInSelection(r));
      }
      strokeRowHeaderGrid(scrollRows, sy0);
      ctx.restore();
    }

    const frozenRowsList = collectFrozenRows();
    if (frozenRowsList.length > 0 && frozenRows > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, headerH, headerW, fh);
      ctx.clip();
      for (const r of frozenRowsList) {
        const y = this.cellTopY(sheet, layout, r);
        this.paintRowHeaderCell(ctx, r, y, rowH, headerW, rowInSelection(r));
      }
      strokeRowHeaderGrid(frozenRowsList, headerH);
      ctx.restore();
    }
  }

  private paintRowHeaderCell(
    ctx: CanvasRenderingContext2D,
    r: number,
    y: number,
    rowH: number,
    headerW: number,
    isSelectionRow: boolean,
  ): void {
    ctx.fillStyle = isSelectionRow ? this.theme.headerActiveBg : this.theme.headerBg;
    ctx.fillRect(0, y, headerW, rowH);
    ctx.fillStyle = isSelectionRow ? this.theme.activeCellBorderColor : this.theme.headerColor;
    ctx.font = `${isSelectionRow ? "bold " : ""}${this.scaledFontSizePx(12)}px system-ui, -apple-system, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(String(r + 1), headerW / 2, y + rowH / 2);
  }

  /**
   * 选区对应的列标题底边、行标题右侧强调带（`activeCellBorderColor`）。
   * 须在 `drawHeadingBodyDividerLines` 之后绘制，压在网格分割线上方。
   * 选区包含第 0 行时不再画列底强调（与表体顶边选框邻接）；包含第 0 列时不再画行右强调（与表体左边选框邻接），避免双线叠粗。
   */
  private paintAllHeaderSelectionAccents(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
    canvasH: number,
    selectionSnap: SelectionPaintSnapshot | null,
  ): void {
    if (selectionSnap === null) {
      return;
    }
    const selSpan = this.getClampedSelectionSpan(sheet, selectionSnap);
    if (selSpan === null) {
      return;
    }
    const { ctx } = this;
    const colW = this.scaledColW(sheet);
    const rowH = this.scaledRowH(sheet);
    const t = viewScaledSelectionOutlineWidth(this.viewZoom);
    const omitColBottom = selSpan.startRow === 0;
    const omitRowRight = selSpan.startCol === 0;

    ctx.save();
    ctx.globalAlpha = 1;
    ctx.globalCompositeOperation = "source-over";

    if (!omitColBottom) {
      ctx.beginPath();
      ctx.rect(headerW, 0, canvasW - headerW, headerH);
      ctx.clip();
      ctx.fillStyle = this.theme.activeCellBorderColor;
      for (let c = selSpan.startCol; c <= selSpan.endCol; c++) {
        if (c < 0 || c >= sheet.colCount) {
          continue;
        }
        const x = this.cellLeftX(sheet, layout, c);
        if (x + colW <= headerW || x >= canvasW) {
          continue;
        }
        ctx.fillRect(x, headerH - t, colW, t);
      }
    }

    ctx.restore();

    ctx.save();
    ctx.globalAlpha = 1;
    ctx.globalCompositeOperation = "source-over";

    if (!omitRowRight) {
      ctx.beginPath();
      ctx.rect(0, headerH, headerW, canvasH - headerH);
      ctx.clip();
      ctx.fillStyle = this.theme.activeCellBorderColor;
      for (let r = selSpan.startRow; r <= selSpan.endRow; r++) {
        if (r < 0 || r >= sheet.rowCount) {
          continue;
        }
        const y = this.cellTopY(sheet, layout, r);
        if (y + rowH <= headerH || y >= canvasH) {
          continue;
        }
        ctx.fillRect(headerW - t, y, t, rowH);
      }
    }

    ctx.restore();
  }

  /** 列 `c` 左边缘画布 X（含冻结与滚动）。 */
  private cellLeftX(sheet: Worksheet, layout: FrozenLayout, c: number): number {
    const colW = this.scaledColW(sheet);
    const { headerW, frozenCols, frozenWidthPx } = layout;
    if (c < frozenCols) {
      return headerW + c * colW;
    }
    return headerW + frozenWidthPx + (c - frozenCols) * colW - this.scrollX;
  }

  /** 行 `r` 顶边画布 Y（含冻结与滚动）。 */
  private cellTopY(sheet: Worksheet, layout: FrozenLayout, r: number): number {
    const rowH = this.scaledRowH(sheet);
    const { headerH, frozenRows, frozenHeightPx } = layout;
    if (r < frozenRows) {
      return headerH + r * rowH;
    }
    return headerH + frozenHeightPx + (r - frozenRows) * rowH - this.scrollY;
  }

  private cellIntersectsCanvas(
    x: number,
    y: number,
    colW: number,
    rowH: number,
    headerW: number,
    headerH: number,
    canvasW: number,
    canvasH: number,
  ): boolean {
    return x + colW > headerW && x < canvasW && y + rowH > headerH && y < canvasH;
  }

  /**
   * 冻结窗格表体：
   * 1. 先画滚动象限（clip 为 [sx0,sy0] 起右下角矩形），网格线端点限制在同一 clipBounds 内，不画入冻结带。
   * 2. 再画冻结顶条 / 左条 / 角区，各自独立 clip + clipBounds。
   * 3. 冻结分隔线由 `paint()` 中 `drawFreezeLines` 在表体之后绘制（不与滚动区网格混画）。
   */
  private drawBody(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
    canvasH: number,
  ): void {
    const buf = this.viewportBuffer;
    const colW = this.scaledColW(sheet);
    const rowH = this.scaledRowH(sheet);
    const { ctx } = this;
    const fw = layout.frozenWidthPx;
    const fh = layout.frozenHeightPx;
    const sx0 = headerW + fw;
    const sy0 = headerH + fh;
    const fr = layout.frozenRows;
    const fc = layout.frozenCols;
    const vr = visibleScrollableCellRange(
      sheet,
      layout,
      this.scrollX,
      this.scrollY,
      buf,
      this.viewZoom,
    );

    const drawBodyPass = (
      clipX: number,
      clipY: number,
      clipW: number,
      clipH: number,
      r0: number,
      r1: number,
      c0: number,
      c1: number,
    ): void => {
      if (r0 > r1 || c0 > c1) {
        return;
      }
      /** 网格线与表体 clip 外缘对齐（snapLine），不再内缩，避免与标题/冻结线出现可见缝。 */
      let strokeBounds: { minX: number; minY: number; maxX: number; maxY: number } | null = null;
      if (clipW > 0 && clipH > 0) {
        strokeBounds = {
          minX: clipX,
          minY: clipY,
          maxX: clipX + clipW,
          maxY: clipY + clipH,
        };
      }
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipW, clipH);
      ctx.clip();
      this.paintBodyCellFills(sheet, layout, headerW, headerH, canvasW, canvasH, r0, r1, c0, c1);

      if (strokeBounds !== null && this.showGridLines) {
        this.strokeBodyGrid(sheet, layout, r0, r1, c0, c1, colW, rowH, strokeBounds);
      }

      this.paintBodyCellTexts(sheet, layout, headerW, headerH, canvasW, canvasH, r0, r1, c0, c1);
      ctx.restore();
    };

    if (
      fr < sheet.rowCount &&
      fc < sheet.colCount &&
      vr.startRow <= vr.endRow &&
      vr.startCol <= vr.endCol
    ) {
      drawBodyPass(
        sx0,
        sy0,
        canvasW - sx0,
        canvasH - sy0,
        vr.startRow,
        vr.endRow,
        vr.startCol,
        vr.endCol,
      );
    }

    if (fr > 0 && fc < sheet.colCount && layout.scrollViewportW > 0) {
      const maxC = sheet.colCount - 1;
      const firstCol = fc + Math.floor(this.scrollX / colW) - buf;
      const lastCol = fc + Math.ceil((this.scrollX + layout.scrollViewportW) / colW) - 1 + buf;
      const c0 = Math.max(fc, firstCol);
      const c1 = Math.min(maxC, lastCol);
      const r0 = 0;
      const r1 = Math.min(fr - 1, sheet.rowCount - 1);
      if (c0 <= c1 && r0 <= r1) {
        drawBodyPass(sx0, headerH, canvasW - sx0, fh, r0, r1, c0, c1);
      }
    }

    if (fc > 0 && fr < sheet.rowCount && vr.startRow <= vr.endRow && layout.scrollViewportH > 0) {
      const c0 = 0;
      const c1 = Math.min(fc - 1, sheet.colCount - 1);
      drawBodyPass(headerW, sy0, fw, canvasH - sy0, vr.startRow, vr.endRow, c0, c1);
    }

    if (fr > 0 && fc > 0) {
      const r0 = 0;
      const r1 = Math.min(fr - 1, sheet.rowCount - 1);
      const c0 = 0;
      const c1 = Math.min(fc - 1, sheet.colCount - 1);
      if (r0 <= r1 && c0 <= c1) {
        drawBodyPass(headerW, headerH, fw, fh, r0, r1, c0, c1);
      }
    }
  }

  private paintBodyCellFills(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
    canvasH: number,
    r0: number,
    r1: number,
    c0: number,
    c1: number,
  ): void {
    const { ctx } = this;
    const colW = this.scaledColW(sheet);
    const rowH = this.scaledRowH(sheet);
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const x = this.cellLeftX(sheet, layout, c);
        const y = this.cellTopY(sheet, layout, r);
        if (!this.cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
          continue;
        }
        const cell = sheet.getCell(r, c);
        const fillArgb = cell.style?.fillArgb;
        const fillCss =
          fillArgb !== undefined && fillArgb !== ""
            ? (argbToCss(fillArgb) ?? this.theme.cellBg)
            : this.theme.cellBg;
        ctx.fillStyle = fillCss;
        ctx.fillRect(x, y, colW, rowH);
      }
    }
  }

  private paintBodyCellTexts(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
    canvasH: number,
    r0: number,
    r1: number,
    c0: number,
    c1: number,
  ): void {
    const { ctx } = this;
    const colW = this.scaledColW(sheet);
    const rowH = this.scaledRowH(sheet);
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const x = this.cellLeftX(sheet, layout, c);
        const y = this.cellTopY(sheet, layout, r);
        if (!this.cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
          continue;
        }
        const cell = sheet.getCell(r, c);
        const text = this.formatCellDisplay(cell.value);
        if (text === "") {
          continue;
        }
        const fgArgb = cell.style?.fgArgb;
        ctx.fillStyle =
          fgArgb !== undefined && fgArgb !== ""
            ? (argbToCss(fgArgb) ?? this.theme.cellColor)
            : this.theme.cellColor;
        const weight = cell.style?.bold === true ? "600" : "400";
        ctx.font = `${weight} ${this.scaledFontSizePx(13)}px system-ui, -apple-system, sans-serif`;
        ctx.textAlign = "left";
        ctx.textBaseline = "middle";
        const pad = 4;
        const maxTextW = colW - pad * 2;
        const display = this.truncateText(ctx, text, maxTextW);
        ctx.fillText(display, x + pad, y + rowH / 2);
      }
    }
  }

  private strokeBodyGrid(
    sheet: Worksheet,
    layout: FrozenLayout,
    r0: number,
    r1: number,
    c0: number,
    c1: number,
    colW: number,
    rowH: number,
    clipBounds: { minX: number; minY: number; maxX: number; maxY: number },
  ): void {
    const { ctx } = this;
    if (r0 > r1 || c0 > c1) return;

    const { minX: bx0, minY: by0, maxX: bx1, maxY: by1 } = clipBounds;
    const { frozenCols, frozenRows } = layout;

    ctx.strokeStyle = this.theme.gridLineColor;
    ctx.lineWidth = 1;
    ctx.beginPath();

    const yTop = this.cellTopY(sheet, layout, r0);
    const yBottom = this.cellTopY(sheet, layout, r1) + rowH;
    const xLeft = this.cellLeftX(sheet, layout, c0);
    const xRight = this.cellLeftX(sheet, layout, c1) + colW;

    // 竖线
    for (let c = c0; c <= c1 + 1; c++) {
      const x =
        c < sheet.colCount
          ? this.cellLeftX(sheet, layout, c)
          : this.cellLeftX(sheet, layout, sheet.colCount - 1) + colW;

      // ========== 核心修复：跳过冻结边界那条竖线（C 列左边线）==========
      if (c === frozenCols) continue;

      if (x < bx0 || x > bx1) continue;
      const xs = snapLine(x);
      const y1 = snapLine(Math.max(by0, yTop));
      const y2 = snapLine(Math.min(by1, yBottom));
      if (y2 > y1) {
        ctx.moveTo(xs, y1);
        ctx.lineTo(xs, y2);
      }
    }

    // 横线
    for (let r = r0; r <= r1 + 1; r++) {
      const y =
        r < sheet.rowCount
          ? this.cellTopY(sheet, layout, r)
          : this.cellTopY(sheet, layout, sheet.rowCount - 1) + rowH;

      // ========== 核心修复：跳过冻结边界那条横线（第3行上边线）==========
      if (r === frozenRows) continue;

      if (y < by0 || y > by1) continue;
      const ys = snapLine(y);
      const x1 = snapLine(Math.max(bx0, xLeft));
      const x2 = snapLine(Math.min(bx1, xRight));
      if (x2 > x1) {
        ctx.moveTo(x1, ys);
        ctx.lineTo(x2, ys);
      }
    }

    ctx.stroke();
  }

  private drawSelectionOverlay(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
    canvasH: number,
  ): void {
    const snap = this.getSelectionSnapshot();
    if (snap === null) {
      return;
    }
    const { range, activeRow, activeCol } = snap;
    const { ctx } = this;
    ctx.globalAlpha = 1;
    ctx.globalCompositeOperation = "source-over";
    const colW = this.scaledColW(sheet);
    const rowH = this.scaledRowH(sheet);
    const buf = this.viewportBuffer;
    const fw = layout.frozenWidthPx;
    const fh = layout.frozenHeightPx;
    const sx0 = headerW + fw;
    const sy0 = headerH + fh;
    const fr = layout.frozenRows;
    const fc = layout.frozenCols;
    const vr = visibleScrollableCellRange(
      sheet,
      layout,
      this.scrollX,
      this.scrollY,
      buf,
      this.viewZoom,
    );

    const drawFillPass = (
      clipX: number,
      clipY: number,
      clipW: number,
      clipH: number,
      pr0: number,
      pr1: number,
      pc0: number,
      pc1: number,
    ): void => {
      const r0 = Math.max(pr0, range.startRow);
      const r1 = Math.min(pr1, range.endRow);
      const c0 = Math.max(pc0, range.startCol);
      const c1 = Math.min(pc1, range.endCol);
      if (r0 > r1 || c0 > c1) {
        return;
      }
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipW, clipH);
      ctx.clip();
      for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
          if (r === activeRow && c === activeCol) {
            continue;
          }
          const x = this.cellLeftX(sheet, layout, c);
          const y = this.cellTopY(sheet, layout, r);
          if (!this.cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
            continue;
          }
          ctx.fillStyle = this.theme.selectionFillColor;
          ctx.fillRect(x, y, colW, rowH);
        }
      }
      ctx.restore();
    };

    // 1. 绘制选区填充（分四个象限，严格裁剪）
    if (
      fr < sheet.rowCount &&
      fc < sheet.colCount &&
      vr.startRow <= vr.endRow &&
      vr.startCol <= vr.endCol
    ) {
      drawFillPass(
        sx0,
        sy0,
        canvasW - sx0,
        canvasH - sy0,
        vr.startRow,
        vr.endRow,
        vr.startCol,
        vr.endCol,
      );
    }
    if (fr > 0 && fc < sheet.colCount && layout.scrollViewportW > 0) {
      const maxC = sheet.colCount - 1;
      const firstCol = fc + Math.floor(this.scrollX / colW) - buf;
      const lastCol = fc + Math.ceil((this.scrollX + layout.scrollViewportW) / colW) - 1 + buf;
      const c0 = Math.max(fc, firstCol);
      const c1 = Math.min(maxC, lastCol);
      const r0 = 0;
      const r1 = Math.min(fr - 1, sheet.rowCount - 1);
      if (c0 <= c1 && r0 <= r1) {
        drawFillPass(sx0, headerH, canvasW - sx0, fh, r0, r1, c0, c1);
      }
    }
    if (fc > 0 && fr < sheet.rowCount && vr.startRow <= vr.endRow && layout.scrollViewportH > 0) {
      const c0 = 0;
      const c1 = Math.min(fc - 1, sheet.colCount - 1);
      drawFillPass(headerW, sy0, fw, canvasH - sy0, vr.startRow, vr.endRow, c0, c1);
    }
    if (fr > 0 && fc > 0) {
      const r0 = 0;
      const r1 = Math.min(fr - 1, sheet.rowCount - 1);
      const c0 = 0;
      const c1 = Math.min(fc - 1, sheet.colCount - 1);
      if (r0 <= r1 && c0 <= c1) {
        drawFillPass(headerW, headerH, fw, fh, r0, r1, c0, c1);
      }
    }

    // 2. 绘制选区外框 + 激活单元格（分象限严格裁剪，彻底解决穿透）
    const drawBorderPass = (
      clipX: number,
      clipY: number,
      clipW: number,
      clipH: number,
      pr0: number,
      pr1: number,
      pc0: number,
      pc1: number,
    ) => {
      const r0 = Math.max(pr0, range.startRow);
      const r1 = Math.min(pr1, range.endRow);
      const c0 = Math.max(pc0, range.startCol);
      const c1 = Math.min(pc1, range.endCol);
      if (r0 > r1 || c0 > c1) return;

      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipW, clipH);
      ctx.clip();
      // 选区覆盖层统一关闭阴影，保持 Excel 式干净描边
      ctx.shadowColor = "transparent";
      ctx.shadowBlur = 0;
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 0;

      // 选区外框
      const x0 = this.cellLeftX(sheet, layout, c0);
      const y0 = this.cellTopY(sheet, layout, r0);
      const bw = (c1 - c0 + 1) * colW;
      const bh = (r1 - r0 + 1) * rowH;
      ctx.strokeStyle = this.theme.selectionBorderColor;
      ctx.lineWidth = viewScaledSelectionOutlineWidth(this.viewZoom);
      ctx.strokeRect(x0 + 0.5, y0 + 0.5, bw - 1, bh - 1);

      ctx.restore();
    };

    // 右下角滚动象限
    if (
      fr < sheet.rowCount &&
      fc < sheet.colCount &&
      vr.startRow <= vr.endRow &&
      vr.startCol <= vr.endCol
    ) {
      drawBorderPass(
        sx0,
        sy0,
        canvasW - sx0,
        canvasH - sy0,
        vr.startRow,
        vr.endRow,
        vr.startCol,
        vr.endCol,
      );
    }
    // 顶部冻结+滚动列
    if (fr > 0 && fc < sheet.colCount && layout.scrollViewportW > 0) {
      const maxC = sheet.colCount - 1;
      const firstCol = fc + Math.floor(this.scrollX / colW) - buf;
      const lastCol = fc + Math.ceil((this.scrollX + layout.scrollViewportW) / colW) - 1 + buf;
      const c0 = Math.max(fc, firstCol);
      const c1 = Math.min(maxC, lastCol);
      const r0 = 0;
      const r1 = Math.min(fr - 1, sheet.rowCount - 1);
      if (c0 <= c1 && r0 <= r1) {
        drawBorderPass(sx0, headerH, canvasW - sx0, fh, r0, r1, c0, c1);
      }
    }
    // 左侧冻结+滚动行
    if (fc > 0 && fr < sheet.rowCount && vr.startRow <= vr.endRow && layout.scrollViewportH > 0) {
      const c0 = 0;
      const c1 = Math.min(fc - 1, sheet.colCount - 1);
      drawBorderPass(headerW, sy0, fw, canvasH - sy0, vr.startRow, vr.endRow, c0, c1);
    }
    // 左上角冻结区
    if (fr > 0 && fc > 0) {
      const r0 = 0;
      const r1 = Math.min(fr - 1, sheet.rowCount - 1);
      const c0 = 0;
      const c1 = Math.min(fc - 1, sheet.colCount - 1);
      if (r0 <= r1 && c0 <= c1) {
        drawBorderPass(headerW, headerH, fw, fh, r0, r1, c0, c1);
      }
    }

    // 3. Excel 风格填充柄（选区右下角小方块）
    const handleCenterX = this.cellLeftX(sheet, layout, range.endCol) + colW;
    const handleCenterY = this.cellTopY(sheet, layout, range.endRow) + rowH;
    const handleSize = Math.max(4, 6 * SELECTION_OUTLINE_VISUAL_SCALE * this.viewZoom);
    const handleHalf = handleSize / 2;
    const bodyX = headerW;
    const bodyY = headerH;
    const bodyW = Math.max(0, canvasW - headerW);
    const bodyH = Math.max(0, canvasH - headerH);
    if (bodyW > 0 && bodyH > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(bodyX, bodyY, bodyW, bodyH);
      ctx.clip();
      ctx.shadowColor = "transparent";
      ctx.shadowBlur = 0;
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 0;
      ctx.fillStyle = this.theme.selectionBorderColor;
      ctx.fillRect(handleCenterX - handleHalf, handleCenterY - handleHalf, handleSize, handleSize);
      ctx.restore();
    }
  }

  /**
   * 列标题底边、行标题右侧与表体之间的分隔线：与表体网格同色同线宽，snapLine 对齐。
   * 须在 `drawBody` 之后绘制，以免首行/首列填充盖住线段；与 `strokeBodyGrid` 在 (headerW,headerH) 处衔接，无双线。
   */
  private drawHeadingBodyDividerLines(
    headerW: number,
    headerH: number,
    canvasW: number,
    canvasH: number,
  ): void {
    if (headerW <= 0 || headerH <= 0) {
      return;
    }
    const { ctx } = this;
    const x0 = snapLine(0);
    const xHw = snapLine(headerW);
    const y0 = snapLine(headerH);
    const x1 = snapLine(canvasW);
    const y1 = snapLine(canvasH);
    ctx.save();
    ctx.globalAlpha = 1;
    ctx.globalCompositeOperation = "source-over";
    ctx.strokeStyle = this.theme.gridLineColor;
    ctx.lineWidth = 1;
    ctx.lineCap = "butt";
    ctx.lineJoin = "miter";
    ctx.beginPath();
    ctx.moveTo(x0, y0);
    ctx.lineTo(x1, y0);
    ctx.moveTo(xHw, y0);
    ctx.lineTo(xHw, y1);
    ctx.stroke();
    ctx.restore();
  }

  /**
   * 冻结分隔线：只画在表体区域 [headerW,canvasW)×[headerH,canvasH)，不贯穿行列标题。
   * 若从 (0,0) 或 (0,sy0) 画到全画布，会在 (sx0,sy0) 形成十字并叠在冻结格与 C3 角上，
   * 滚动时像「C3 左上交叉线」穿入冻结区；表体-only 后十字仅出现在可滚动象限内。
   */
  private drawFreezeLines(layout: FrozenLayout, canvasW: number, canvasH: number): void {
    const { ctx } = this;
    const { headerW, headerH, frozenCols, frozenRows, frozenWidthPx, frozenHeightPx } = layout;
    if (frozenCols === 0 && frozenRows === 0) {
      return;
    }
    ctx.save();
    ctx.globalAlpha = 1;
    ctx.globalCompositeOperation = "source-over";
    ctx.strokeStyle = this.theme.freezeLineColor;
    ctx.lineWidth = 1;
    const xBody0 = snapLine(headerW);
    const yBody0 = snapLine(headerH);
    const xBody1 = snapLine(canvasW);
    const yBody1 = snapLine(canvasH);
    if (frozenCols > 0) {
      const x = headerW + frozenWidthPx;
      ctx.beginPath();
      ctx.moveTo(snapLine(x), yBody0);
      ctx.lineTo(snapLine(x), yBody1);
      ctx.stroke();
    }
    if (frozenRows > 0) {
      const y = headerH + frozenHeightPx;
      ctx.beginPath();
      ctx.moveTo(xBody0, snapLine(y));
      ctx.lineTo(xBody1, snapLine(y));
      ctx.stroke();
    }
    ctx.restore();
  }

  private formatCellDisplay(value: CellScalar): string {
    if (value === null || value === "") {
      return "";
    }
    if (typeof value === "boolean") {
      return value ? "TRUE" : "FALSE";
    }
    return String(value);
  }

  private truncateText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string {
    if (ctx.measureText(text).width <= maxWidth) {
      return text;
    }
    let s = text;
    const ell = "…";
    while (s.length > 0 && ctx.measureText(s + ell).width > maxWidth) {
      s = s.slice(0, -1);
    }
    return s + ell;
  }
}

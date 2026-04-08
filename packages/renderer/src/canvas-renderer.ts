import {
  normalizeSelectionRange,
  type SelectionPaintSnapshot,
  type SelectionRange,
  type Workbook,
} from "@flexsheet/core";
import type { SheetTheme } from "@flexsheet/theme";
import { scrollToRevealCell } from "./grid-hit-test.js";
import { tryActiveFrozenContext } from "./canvas-renderer-active-layout.js";
import { drawBody, type BodyPaintEnv } from "./canvas-renderer-body.js";
import {
  HEADER_STRIP_BASE_HEIGHT,
  HEADER_STRIP_BASE_WIDTH,
  VIEW_ZOOM_MAX,
  VIEW_ZOOM_MIN,
} from "./canvas-renderer-constants.js";
import { drawFreezeLines } from "./canvas-renderer-freeze.js";
import {
  drawAllColumnHeaders,
  drawAllRowHeaders,
  drawCorner,
  drawHeadingBodyDividerLines,
  drawRulerOverlay,
  paintAllHeaderSelectionAccents,
  type HeaderPaintEnv,
} from "./canvas-renderer-headers.js";
import { cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import { drawClipboardMarqueeOverlay } from "./canvas-renderer-clipboard-marquee.js";
import {
  drawSelectionOverlay,
  type SelectionOverlayEnv,
} from "./canvas-renderer-selection-overlay.js";
import type {
  CanvasRendererOptions,
  HorizontalScrollMetrics,
  VerticalScrollMetrics,
} from "./canvas-renderer-types.js";
export type {
  CanvasRendererOptions,
  HorizontalScrollMetrics,
  VerticalScrollMetrics,
} from "./canvas-renderer-types.js";
import {
  buildCellCanvasFont,
  scaledColWidthAt,
  scaledRowHeightAt,
} from "./canvas-renderer-utils.js";
import { buildFrozenLayout, clampScroll, computeScrollLimits } from "./viewport.js";

/**
 * Canvas 2D 主渲染器：行列标题、网格线、冻结窗格与视口虚拟滚动（只读数据）。
 */
export class CanvasRenderer {
  private readonly canvas: HTMLCanvasElement;
  private readonly ctx: CanvasRenderingContext2D;
  private workbook: Workbook;
  private theme: SheetTheme;
  private getSelectionSnapshot: () => SelectionPaintSnapshot | null;
  private getClipboardMarqueeRange: () => SelectionRange | null;

  /** 走马灯动画：在 `paint` 末尾调度下一帧重绘。 */
  private marqueeContinuationRaf: number | null = null;

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
    this.getClipboardMarqueeRange = options.getClipboardMarqueeRange ?? (() => null);
  }

  setSelectionSnapshotProvider(fn: (() => SelectionPaintSnapshot | null) | undefined): void {
    this.getSelectionSnapshot = fn ?? (() => null);
  }

  setClipboardMarqueeRangeProvider(fn: (() => SelectionRange | null) | undefined): void {
    this.getClipboardMarqueeRange = fn ?? (() => null);
  }

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

  getHorizontalScrollMetrics(): HorizontalScrollMetrics | null {
    const ctx = tryActiveFrozenContext(
      this.workbook,
      this.canvas,
      this.showHeadings,
      this.viewZoom,
      this.frozenRows,
      this.frozenCols,
    );
    if (ctx === null) {
      return null;
    }
    const { sheet, layout } = ctx;
    let contentScrollWidth = 0;
    for (let c = layout.frozenCols; c < sheet.colCount; c++) {
      contentScrollWidth += scaledColWidthAt(sheet, c, this.viewZoom);
    }
    return {
      scrollX: this.scrollX,
      maxScrollX: ctx.limits.maxScrollX,
      scrollViewportW: layout.scrollViewportW,
      contentScrollWidth,
    };
  }

  getVerticalScrollMetrics(): VerticalScrollMetrics | null {
    const ctx = tryActiveFrozenContext(
      this.workbook,
      this.canvas,
      this.showHeadings,
      this.viewZoom,
      this.frozenRows,
      this.frozenCols,
    );
    if (ctx === null) {
      return null;
    }
    const { sheet, layout } = ctx;
    let contentScrollHeight = 0;
    for (let r = layout.frozenRows; r < sheet.rowCount; r++) {
      contentScrollHeight += scaledRowHeightAt(sheet, r, this.viewZoom);
    }
    return {
      scrollY: this.scrollY,
      maxScrollY: ctx.limits.maxScrollY,
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

  getCellEditorFontCss(row: number, col: number): string {
    const sheet = this.workbook.getActiveSheet();
    const style = sheet !== undefined ? sheet.getCell(row, col).style : null;
    return buildCellCanvasFont(style, this.viewZoom);
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
    const ctx = tryActiveFrozenContext(
      this.workbook,
      this.canvas,
      this.showHeadings,
      this.viewZoom,
      this.frozenRows,
      this.frozenCols,
    );
    if (ctx === null) {
      return null;
    }
    return {
      x: cellLeftX(sheet, ctx.layout, col, this.viewZoom, this.scrollX),
      y: cellTopY(sheet, ctx.layout, row, this.viewZoom, this.scrollY),
      width: scaledColWidthAt(sheet, col, this.viewZoom),
      height: scaledRowHeightAt(sheet, row, this.viewZoom),
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

  ensureScrollClamped(): void {
    this.clampScrollToActiveSheet();
  }

  applyScrollDelta(deltaX: number, deltaY: number): void {
    this.scrollX += deltaX;
    this.scrollY += deltaY;
    this.clampScrollToActiveSheet();
  }

  private clampScrollToActiveSheet(): void {
    const ctx = tryActiveFrozenContext(
      this.workbook,
      this.canvas,
      this.showHeadings,
      this.viewZoom,
      this.frozenRows,
      this.frozenCols,
    );
    if (ctx === null) {
      return;
    }
    const c = clampScroll(this.scrollX, this.scrollY, ctx.limits);
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

  cancelPendingRedraw(): void {
    if (this.rafId !== null) {
      cancelAnimationFrame(this.rafId);
      this.rafId = null;
    }
    this.clearMarqueeContinuation();
  }

  private clearMarqueeContinuation(): void {
    if (this.marqueeContinuationRaf !== null) {
      cancelAnimationFrame(this.marqueeContinuationRaf);
      this.marqueeContinuationRaf = null;
    }
  }

  private scheduleMarqueeContinuation(): void {
    if (typeof requestAnimationFrame === "undefined") {
      return;
    }
    if (this.marqueeContinuationRaf !== null) {
      return;
    }
    this.marqueeContinuationRaf = requestAnimationFrame(() => {
      this.marqueeContinuationRaf = null;
      this.requestRedraw();
    });
  }

  draw(): void {
    this.paint();
  }

  private headerEnv(): HeaderPaintEnv {
    return {
      ctx: this.ctx,
      theme: this.theme,
      viewZoom: this.viewZoom,
      viewportBuffer: this.viewportBuffer,
      scrollX: this.scrollX,
      scrollY: this.scrollY,
    };
  }

  private bodyEnv(): BodyPaintEnv {
    return {
      ctx: this.ctx,
      theme: this.theme,
      viewZoom: this.viewZoom,
      showGridLines: this.showGridLines,
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      viewportBuffer: this.viewportBuffer,
    };
  }

  private selectionEnv(): SelectionOverlayEnv {
    return {
      ctx: this.ctx,
      theme: this.theme,
      viewZoom: this.viewZoom,
      viewportBuffer: this.viewportBuffer,
      scrollX: this.scrollX,
      scrollY: this.scrollY,
    };
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
      this.clearMarqueeContinuation();
      ctx.fillStyle = this.theme.cellColor;
      ctx.font = "14px system-ui, sans-serif";
      ctx.textAlign = "left";
      ctx.textBaseline = "alphabetic";
      ctx.fillText("无工作表", 16, 32);
      ctx.restore();
      return;
    }

    const frozenCtx = tryActiveFrozenContext(
      this.workbook,
      this.canvas,
      this.showHeadings,
      this.viewZoom,
      this.frozenRows,
      this.frozenCols,
    );
    if (frozenCtx === null) {
      this.clearMarqueeContinuation();
      ctx.restore();
      return;
    }
    const { layout, limits } = frozenCtx;
    const headerW = frozenCtx.hw;
    const headerH = frozenCtx.hh;

    const clamped = clampScroll(this.scrollX, this.scrollY, limits);
    this.scrollX = clamped.scrollX;
    this.scrollY = clamped.scrollY;

    const selectionSnap = this.getSelectionSnapshot();
    const hEnv = this.headerEnv();

    if (this.showHeadings) {
      drawCorner(ctx, this.theme, headerW, headerH);
      if (this.showRuler) {
        drawRulerOverlay(hEnv, sheet, layout, headerW, headerH, w);
      }
      drawAllColumnHeaders(hEnv, sheet, layout, headerW, headerH, w, selectionSnap);
      drawAllRowHeaders(hEnv, sheet, layout, headerW, headerH, h, selectionSnap);
    }
    drawBody(this.bodyEnv(), sheet, layout, headerW, headerH, w, h);
    if (this.showHeadings) {
      drawHeadingBodyDividerLines(ctx, this.theme, headerW, headerH, w, h);
      paintAllHeaderSelectionAccents(hEnv, sheet, layout, headerW, headerH, w, h, selectionSnap);
    }
    drawFreezeLines(ctx, this.theme, layout, w, h);
    drawSelectionOverlay(this.selectionEnv(), sheet, layout, headerW, headerH, w, h, selectionSnap);

    const clipboardMarqueeRange = this.getClipboardMarqueeRange();
    if (clipboardMarqueeRange === null) {
      this.clearMarqueeContinuation();
    } else {
      const phasePx = performance.now() * 0.011;
      drawClipboardMarqueeOverlay(
        this.selectionEnv(),
        sheet,
        layout,
        headerW,
        headerH,
        w,
        h,
        clipboardMarqueeRange,
        phasePx,
      );
      this.scheduleMarqueeContinuation();
    }

    ctx.restore();
  }
}

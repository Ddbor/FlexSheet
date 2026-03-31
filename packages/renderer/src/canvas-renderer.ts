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

/** 1px 线段对齐到设备无关像素中心（避免模糊与双线粗细不均）。 */
function snapLine(v: number): number {
  return Math.round(v) + 0.5;
}

/** 网格线 stroke 相对内容 clip 的内缩（CSS 像素），避免 1px 描边在 clip 边界外溢到冻结区。 */
const GRID_STROKE_CLIP_INSET = 0.5;

/**
 * 右下角滚动象限网格专用内缩（≥ 线宽的一半 + 抗锯齿余量）。
 * 仅 0.5 时，stroke 仍可能在冻结边界内侧露出，滚动时像「网格线滚进冻结区」。
 */
const SCROLL_PANE_GRID_CLIP_INSET = 1;

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

  private readonly cornerSize = { width: 48, height: 24 };

  private rafId: number | null = null;

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
   */
  getHeaderSize(): Readonly<{ width: number; height: number }> {
    if (!this.showHeadings) {
      return { width: 0, height: 0 };
    }
    return this.cornerSize;
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
    this.viewZoom = z;
    this.ensureScrollClamped();
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
  }

  resize(cssWidth: number, cssHeight: number): void {
    const dpr = window.devicePixelRatio || 1;
    this.canvas.style.width = `${cssWidth}px`;
    this.canvas.style.height = `${cssHeight}px`;
    this.canvas.width = Math.max(1, Math.floor(cssWidth * dpr));
    this.canvas.height = Math.max(1, Math.floor(cssHeight * dpr));
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
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

    if (this.showHeadings) {
      this.drawCorner(headerW, headerH);
      if (this.showRuler) {
        this.drawRulerOverlay(sheet, layout, headerW, headerH, w);
      }
      this.drawAllColumnHeaders(sheet, layout, headerW, headerH, w);
      this.drawAllRowHeaders(sheet, layout, headerW, headerH, h);
    }
    this.drawBody(sheet, layout, headerW, headerH, w, h);
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
    ctx.strokeRect(0.5, 0.5, headerW - 1, headerH - 1);
  }

  /** 列标题区域底部刻度（与主题 header 线色一致）。 */
  private drawRulerOverlay(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
  ): void {
    const { ctx } = this;
    const colW = this.scaledColW(sheet);
    const yBase = headerH - 1.5;
    ctx.strokeStyle = this.theme.headerLineColor;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(headerW, snapLine(yBase));
    ctx.lineTo(canvasW, snapLine(yBase));
    ctx.stroke();
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
        ctx.moveTo(snapLine(x), yBase);
        ctx.lineTo(snapLine(x), yBase - tickH);
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

  private drawAllColumnHeaders(
    sheet: Worksheet,
    layout: FrozenLayout,
    headerW: number,
    headerH: number,
    canvasW: number,
  ): void {
    const { ctx } = this;
    const colW = this.scaledColW(sheet);
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
      ctx.strokeStyle = this.theme.headerLineColor;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(snapLine(clipLeft), 0);
      ctx.lineTo(snapLine(clipLeft), headerH);
      ctx.moveTo(snapLine(clipLeft), snapLine(0));
      ctx.lineTo(snapLine(canvasW), snapLine(0));
      ctx.moveTo(snapLine(clipLeft), snapLine(headerH));
      ctx.lineTo(snapLine(canvasW), snapLine(headerH));
      for (const c of visibleCols) {
        const x = this.cellLeftX(sheet, layout, c);
        const xs = snapLine(x);
        ctx.moveTo(xs, 0);
        ctx.lineTo(xs, headerH);
      }
      const lastC = visibleCols[visibleCols.length - 1];
      const xe = snapLine(this.cellLeftX(sheet, layout, lastC) + colW);
      ctx.moveTo(xe, 0);
      ctx.lineTo(xe, headerH);
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
        this.paintColumnHeaderCell(ctx, c, x, colW, headerH);
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
        this.paintColumnHeaderCell(ctx, c, x, colW, headerH);
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
  ): void {
    ctx.fillStyle = this.theme.headerBg;
    ctx.fillRect(x, 0, colW, headerH);
    ctx.fillStyle = this.theme.headerColor;
    ctx.font = `${this.scaledFontSizePx(12)}px system-ui, -apple-system, sans-serif`;
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
  ): void {
    const { ctx } = this;
    const rowH = this.scaledRowH(sheet);
    const { frozenRows } = layout;
    const buf = this.viewportBuffer;
    const maxR = sheet.rowCount - 1;
    const fh = layout.frozenHeightPx;
    const sy0 = headerH + fh;

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
      ctx.strokeStyle = this.theme.headerLineColor;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(0, snapLine(clipTop));
      ctx.lineTo(snapLine(headerW), snapLine(clipTop));
      ctx.moveTo(snapLine(0), snapLine(clipTop));
      ctx.lineTo(snapLine(0), snapLine(canvasH));
      ctx.moveTo(snapLine(headerW), snapLine(clipTop));
      ctx.lineTo(snapLine(headerW), snapLine(canvasH));
      for (const r of visibleRows) {
        const y = this.cellTopY(sheet, layout, r);
        const ys = snapLine(y);
        ctx.moveTo(0, ys);
        ctx.lineTo(headerW, ys);
      }
      const lastR = visibleRows[visibleRows.length - 1];
      const ye = snapLine(this.cellTopY(sheet, layout, lastR) + rowH);
      ctx.moveTo(0, ye);
      ctx.lineTo(headerW, ye);
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
        this.paintRowHeaderCell(ctx, r, y, rowH, headerW);
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
        this.paintRowHeaderCell(ctx, r, y, rowH, headerW);
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
  ): void {
    ctx.fillStyle = this.theme.headerBg;
    ctx.fillRect(0, y, headerW, rowH);
    ctx.fillStyle = this.theme.headerColor;
    ctx.font = `${this.scaledFontSizePx(12)}px system-ui, -apple-system, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(String(r + 1), headerW / 2, y + rowH / 2);
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
      const inset = GRID_STROKE_CLIP_INSET;
      /** 右下角滚动象限：仅内缩靠冻结区的一侧（左、上），避免 1px 网格线外溢进冻结带；右、下仍贴画布边 */
      const isScrollPane =
        Math.abs(clipX - sx0) < 1e-6 && Math.abs(clipY - sy0) < 1e-6 && clipW > 0 && clipH > 0;

      let innerX: number;
      let innerY: number;
      let innerW: number;
      let innerH: number;
      let strokeBounds: { minX: number; minY: number; maxX: number; maxY: number } | null = null;

      if (isScrollPane) {
        const sInset = SCROLL_PANE_GRID_CLIP_INSET;
        innerX = clipX + sInset;
        innerY = clipY + sInset;
        innerW = Math.max(0, clipW - sInset);
        innerH = Math.max(0, clipH - sInset);
        if (innerW > 0 && innerH > 0) {
          /** 与嵌套 `rect(innerX,innerY,innerW,innerH)` 完全一致，线段端点不超出实际 clip */
          strokeBounds = {
            minX: innerX,
            minY: innerY,
            maxX: innerX + innerW,
            maxY: innerY + innerH,
          };
        }
      } else {
        innerX = clipX + inset;
        innerY = clipY + inset;
        innerW = Math.max(0, clipW - 2 * inset);
        innerH = Math.max(0, clipH - 2 * inset);
        if (innerW > 0 && innerH > 0) {
          strokeBounds = {
            minX: innerX,
            minY: innerY,
            maxX: clipX + clipW - inset,
            maxY: clipY + clipH - inset,
          };
        }
      }

      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipW, clipH);
      ctx.clip();
      this.paintBodyCellFills(sheet, layout, headerW, headerH, canvasW, canvasH, r0, r1, c0, c1);

      if (strokeBounds !== null && this.showGridLines) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(innerX, innerY, innerW, innerH);
        ctx.clip();
        this.strokeBodyGrid(sheet, layout, r0, r1, c0, c1, colW, rowH, strokeBounds);
        ctx.restore();
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
      ctx.lineWidth = 2;
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
    const handleSize = 6;
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
    ctx.strokeStyle = this.theme.freezeLineColor;
    ctx.lineWidth = 1;
    const bodyTop = headerH;
    const bodyBottom = canvasH;
    const bodyLeft = headerW;
    const bodyRight = canvasW;
    if (frozenCols > 0) {
      const x = headerW + frozenWidthPx;
      ctx.beginPath();
      ctx.moveTo(snapLine(x), bodyTop);
      ctx.lineTo(snapLine(x), bodyBottom);
      ctx.stroke();
    }
    if (frozenRows > 0) {
      const y = headerH + frozenHeightPx;
      ctx.beginPath();
      ctx.moveTo(bodyLeft, snapLine(y));
      ctx.lineTo(bodyRight, snapLine(y));
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

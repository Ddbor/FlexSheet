/**
 * 工作表上的浮动图片层（DOM 覆盖在 Canvas 之上，非单元格内嵌）。
 * 插入后可选中，通过把手缩放、旋转与拖动。
 */

import type { Workbook } from "@flexsheet/core";
import type { XlsxFloatingPictureExport } from "@flexsheet/import-export";
import {
  HEADER_STRIP_BASE_HEIGHT,
  HEADER_STRIP_BASE_WIDTH,
  type CanvasRenderer,
} from "@flexsheet/renderer";

const STYLE_ID = "fs-floating-picture-styles";

const HANDLE_PX = 7;
const MIN_W = 24;
const MIN_H = 24;
const INSERT_MAX_DIM = 280;

function dot2(ax: number, ay: number, bx: number, by: number): number {
  return ax * bx + ay * by;
}

function ensureStyles(): void {
  if (typeof document === "undefined" || document.getElementById(STYLE_ID) !== null) {
    return;
  }
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
.fs-fp-root{position:absolute;inset:0;pointer-events:none;z-index:5;overflow:visible;}
.fs-fp-item{position:absolute;pointer-events:auto;box-sizing:border-box;}
.fs-fp-item__img{display:block;width:100%;height:100%;object-fit:fill;user-select:none;-webkit-user-drag:none;}
.fs-fp-item__focus{position:absolute;inset:0;pointer-events:none;border:1px solid #333;box-sizing:border-box;}
.fs-fp-item--selected .fs-fp-item__focus{display:block;}
.fs-fp-item:not(.fs-fp-item--selected) .fs-fp-item__focus{display:none;}
.fs-fp-handle{position:absolute;width:${HANDLE_PX}px;height:${HANDLE_PX}px;margin:-${Math.floor(HANDLE_PX / 2)}px;background:#fff;border:1px solid #333;box-sizing:border-box;pointer-events:auto;z-index:2;}
.fs-fp-handle--n{left:50%;top:0;cursor:ns-resize;}
.fs-fp-handle--s{left:50%;bottom:0;cursor:ns-resize;}
.fs-fp-handle--e{right:0;top:50%;cursor:ew-resize;}
.fs-fp-handle--w{left:0;top:50%;cursor:ew-resize;}
.fs-fp-handle--nw{left:0;top:0;cursor:nwse-resize;}
.fs-fp-handle--ne{right:0;top:0;cursor:nesw-resize;}
.fs-fp-handle--sw{left:0;bottom:0;cursor:nesw-resize;}
.fs-fp-handle--se{right:0;bottom:0;cursor:nwse-resize;}
.fs-fp-rotate{position:absolute;left:50%;width:18px;height:18px;margin-left:-9px;top:-28px;border-radius:50%;background:#fff;border:1px solid #333;cursor:grab;box-sizing:border-box;pointer-events:auto;z-index:2;display:flex;align-items:center;justify-content:center;}
.fs-fp-rotate::after{content:"";width:10px;height:10px;border-radius:50%;border:2px solid #333;border-right-color:transparent;border-bottom-color:transparent;transform:rotate(45deg);}
.fs-fp-rotate-line{position:absolute;left:50%;width:1px;height:12px;margin-left:-0.5px;top:-12px;background:#333;pointer-events:none;}
`;
  document.head.appendChild(s);
}

export interface FloatingPictureLayerOptions {
  readonly mount: HTMLElement;
  readonly getCanvas: () => HTMLCanvasElement;
  readonly getRenderer: () => CanvasRenderer;
  readonly getWorkbook: () => Workbook;
  /** 插入图片时锚定的单元格（通常为当前活动单元格）。 */
  readonly getAnchorCell: () => { readonly row: number; readonly col: number };
}

/** 浮动图显示调整（CSS filter，与「设置图片格式」面板一致）。 */
export interface FloatingPictureAdjustments {
  /** 亮度偏移 %，约 -100～100，0 为原始 */
  brightnessPct: number;
  /** 对比度偏移 % */
  contrastPct: number;
  /** 锐化/柔化：负值柔化（blur），正值增强边缘（对比度近似），约 -100～100 */
  sharpnessPct: number;
  /** 饱和度 %，100 为原始 */
  saturationPct: number;
  /** 色温 K，约 2000～11000，6500 中性 */
  colorTemperatureK: number;
  /** 不透明度损失 %，0 完全不透明 */
  transparencyPct: number;
}

export const DEFAULT_FLOATING_PICTURE_ADJUSTMENTS: FloatingPictureAdjustments = {
  brightnessPct: 0,
  contrastPct: 0,
  sharpnessPct: 0,
  saturationPct: 100,
  colorTemperatureK: 6500,
  transparencyPct: 0,
};

function clampN(n: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, n));
}

function cloneAdjustments(a: FloatingPictureAdjustments): FloatingPictureAdjustments {
  return { ...a };
}

/** 将调整合成为 `img` 的 `filter` 字符串（Canvas 上 DOM 叠加层使用）。 */
export function buildFloatingPictureCssFilter(a: FloatingPictureAdjustments): string {
  const b = clampN(1 + a.brightnessPct / 100, 0.05, 3);
  const c0 = clampN(1 + a.contrastPct / 100, 0.05, 3);
  const sat = clampN(a.saturationPct / 100, 0, 4);
  const op = clampN(1 - a.transparencyPct / 100, 0, 1);
  let blurPx = 0;
  let sharpenBoost = 1;
  if (a.sharpnessPct < 0) {
    blurPx = clampN((-a.sharpnessPct / 100) * 3, 0, 3.5);
  } else if (a.sharpnessPct > 0) {
    sharpenBoost = 1 + (a.sharpnessPct / 100) * 0.55;
  }
  const c = c0 * sharpenBoost;
  const parts: string[] = [];
  if (blurPx > 0.05) {
    parts.push(`blur(${blurPx}px)`);
  }
  parts.push(`brightness(${b})`);
  parts.push(`contrast(${c})`);
  parts.push(`saturate(${sat})`);
  const tk = a.colorTemperatureK;
  if (Number.isFinite(tk) && tk !== 6500) {
    const deg = clampN(((tk - 6500) / 4500) * 28, -32, 32);
    if (Math.abs(deg) > 0.3) {
      parts.push(`hue-rotate(${deg}deg)`);
    }
  }
  parts.push(`opacity(${op})`);
  return parts.join(" ");
}

/** 浮动图完整状态快照（剪切/粘贴撤销用）。 */
export interface FloatingPictureSnapshot {
  readonly id: string;
  readonly sheetName: string;
  readonly sheetIndex: number;
  readonly anchorRow: number;
  readonly anchorCol: number;
  readonly relCX: number;
  readonly relCY: number;
  readonly width: number;
  readonly height: number;
  readonly rotationRad: number;
  readonly dataUrl: string;
  readonly z: number;
  readonly adjustments: FloatingPictureAdjustments;
}

/** 粘贴前异步算好的几何与像素（不含 id/z，避免未执行命令就占用序号）。 */
export interface FloatingPicturePastePrepared {
  readonly sheetName: string;
  readonly sheetIndex: number;
  readonly anchorRow: number;
  readonly anchorCol: number;
  readonly relCX: number;
  readonly relCY: number;
  readonly width: number;
  readonly height: number;
  readonly rotationRad: number;
  readonly dataUrl: string;
}

interface PictureModel {
  readonly id: string;
  /** 归属工作表名（主键，避免仅依赖索引在插入/删除表时错位） */
  sheetName: string;
  /** 与 `sheetName` 解析结果同步的缓存索引 */
  sheetIndex: number;
  anchorRow: number;
  anchorCol: number;
  relCX: number;
  relCY: number;
  width: number;
  height: number;
  rotationRad: number;
  dataUrl: string;
  z: number;
  adjustments: FloatingPictureAdjustments;
}

type ResizeHandleId = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

type DragMode =
  | { kind: "move"; startCX: number; startCY: number; startRelCX: number; startRelCY: number }
  | {
      kind: "resize";
      handle: ResizeHandleId;
      startCenterX: number;
      startCenterY: number;
      startW: number;
      startH: number;
      startRot: number;
    }
  | { kind: "rotate"; startAngle: number; startRot: number; cx: number; cy: number };

function rotLocalToWorld(
  lx: number,
  ly: number,
  rot: number,
): { readonly x: number; readonly y: number } {
  const c = Math.cos(rot);
  const s = Math.sin(rot);
  return { x: lx * c - ly * s, y: lx * s + ly * c };
}

export class FloatingPictureLayer {
  private readonly mount: HTMLElement;
  private readonly getCanvas: () => HTMLCanvasElement;
  private readonly getRenderer: () => CanvasRenderer;
  private readonly getWorkbook: () => Workbook;
  private readonly getAnchorCell: () => { readonly row: number; readonly col: number };
  private readonly root: HTMLDivElement;
  private readonly byId = new Map<string, { model: PictureModel; el: HTMLDivElement }>();
  private readonly floatingPictureFocusListeners = new Set<(active: boolean) => void>();
  private zCounter = 10;
  private selectedId: string | null = null;
  private drag: (DragMode & { id: string }) | null = null;
  private idSeq = 0;

  constructor(options: FloatingPictureLayerOptions) {
    ensureStyles();
    this.mount = options.mount;
    this.getCanvas = options.getCanvas;
    this.getRenderer = options.getRenderer;
    this.getWorkbook = options.getWorkbook;
    this.getAnchorCell = options.getAnchorCell;
    this.root = document.createElement("div");
    this.root.className = "fs-fp-root";
    this.mount.appendChild(this.root);
    this.syncClipToTableBody();
  }

  /**
   * 裁剪到表体区域（与 Canvas 行列标题带一致），避免拖动时图片盖住行号/列标。
   */
  syncClipToTableBody(): void {
    const renderer = this.getRenderer();
    if (!renderer.showHeadings) {
      this.root.style.clipPath = "none";
      this.root.style.removeProperty("-webkit-clip-path");
      return;
    }
    const z = renderer.getViewZoom();
    const top = HEADER_STRIP_BASE_HEIGHT * z;
    const left = HEADER_STRIP_BASE_WIDTH * z;
    const clip = `inset(${top}px 0 0 ${left}px)`;
    this.root.style.clipPath = clip;
    this.root.style.setProperty("-webkit-clip-path", clip);
  }

  destroy(): void {
    this.endDrag();
    const hadFocus = this.selectedId !== null;
    this.selectedId = null;
    this.root.remove();
    this.byId.clear();
    if (hadFocus) {
      this.emitFloatingPictureFocus(false);
    }
    this.floatingPictureFocusListeners.clear();
  }

  /** Ribbon「图片格式」等：浮动图选中/取消时订阅；订阅后会立即用当前状态回调一次。 */
  subscribeFloatingPictureFocus(listener: (active: boolean) => void): () => void {
    this.floatingPictureFocusListeners.add(listener);
    listener(this.selectedId !== null);
    return () => this.floatingPictureFocusListeners.delete(listener);
  }

  private emitFloatingPictureFocus(active: boolean): void {
    for (const fn of this.floatingPictureFocusListeners) {
      try {
        fn(active);
      } catch {
        /* 宿主回调错误不影响编辑 */
      }
    }
  }

  clearAll(): void {
    this.endDrag();
    const hadFocus = this.selectedId !== null;
    this.selectedId = null;
    for (const { el } of this.byId.values()) {
      el.remove();
    }
    this.byId.clear();
    if (hadFocus) {
      this.emitFloatingPictureFocus(false);
    }
  }

  getDataUrlForPicture(id: string): string | null {
    return this.byId.get(id)?.model.dataUrl ?? null;
  }

  /** 移除指定浮动图（右键剪切等）。 */
  removePictureById(id: string): void {
    const rec = this.byId.get(id);
    if (rec === undefined) {
      return;
    }
    if (this.drag?.id === id) {
      this.endDrag();
    }
    if (this.selectedId === id) {
      this.deselect();
    }
    rec.el.remove();
    this.byId.delete(id);
  }

  /** 浮动层根节点（供右键菜单等挂载监听）。 */
  getRootElement(): HTMLElement {
    return this.root;
  }

  /** 右键菜单等：选中指定浮动图（若存在）。 */
  focusPictureById(id: string): void {
    if (!this.byId.has(id)) {
      return;
    }
    this.selectById(id);
  }

  getSelectedPictureId(): string | null {
    return this.selectedId;
  }

  takeFloatingPictureSnapshot(pictureId: string): FloatingPictureSnapshot | null {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return null;
    }
    const m = rec.model;
    return {
      id: m.id,
      sheetName: m.sheetName,
      sheetIndex: m.sheetIndex,
      anchorRow: m.anchorRow,
      anchorCol: m.anchorCol,
      relCX: m.relCX,
      relCY: m.relCY,
      width: m.width,
      height: m.height,
      rotationRad: m.rotationRad,
      dataUrl: m.dataUrl,
      z: m.z,
      adjustments: cloneAdjustments(m.adjustments),
    };
  }

  /** 按快照恢复浮动图（剪切撤销 / 粘贴重做）；若 id 已存在则忽略。 */
  restoreFloatingPicture(snapshot: FloatingPictureSnapshot): void {
    if (this.byId.has(snapshot.id)) {
      return;
    }
    const model: PictureModel = {
      id: snapshot.id,
      sheetName: snapshot.sheetName,
      sheetIndex: snapshot.sheetIndex,
      anchorRow: snapshot.anchorRow,
      anchorCol: snapshot.anchorCol,
      relCX: snapshot.relCX,
      relCY: snapshot.relCY,
      width: snapshot.width,
      height: snapshot.height,
      rotationRad: snapshot.rotationRad,
      dataUrl: snapshot.dataUrl,
      z: snapshot.z,
      adjustments: cloneAdjustments(snapshot.adjustments),
    };
    this.zCounter = Math.max(this.zCounter, snapshot.z);
    const seqNum = Number.parseInt(snapshot.id.replace(/^fp-/, ""), 10);
    if (Number.isFinite(seqNum)) {
      this.idSeq = Math.max(this.idSeq, seqNum);
    }
    const el = this.createItemElement(model);
    this.byId.set(snapshot.id, { model, el });
    this.root.appendChild(el);
    this.selectById(snapshot.id);
    this.layout();
  }

  /**
   * 为「粘贴」异步解码尺寸并得到可导出 dataUrl；不修改文档、不占用 id/z 序号。
   * 几何与 `addPictureFromDataUrl` 一致。
   */
  buildPasteFloatingPicturePrepared(dataUrl: string): Promise<FloatingPicturePastePrepared | null> {
    const renderer = this.getRenderer();
    const wb = this.getWorkbook();
    const sheet = wb.getActiveSheet();
    if (sheet === undefined) {
      return Promise.resolve(null);
    }
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = (): void => {
        let w = img.naturalWidth || 1;
        let h = img.naturalHeight || 1;
        const scale = Math.min(1, INSERT_MAX_DIM / Math.max(w, h));
        w = Math.max(MIN_W, w * scale);
        h = Math.max(MIN_H, h * scale);

        const ac = this.getAnchorCell();
        const sheet0 = wb.getActiveSheet();
        if (sheet0 === undefined) {
          resolve(null);
          return;
        }
        const cell = {
          row: Math.max(0, Math.min(sheet0.rowCount - 1, ac.row)),
          col: Math.max(0, Math.min(sheet0.colCount - 1, ac.col)),
        };
        const cr = renderer.getCellRectInCanvasPixels(cell.row, cell.col);
        if (cr === null) {
          resolve(null);
          return;
        }
        const exportableUrl = normalizeToExportableDataUrl(img, dataUrl);
        resolve({
          sheetName: sheet0.name,
          sheetIndex: wb.activeSheetIndex,
          anchorRow: cell.row,
          anchorCol: cell.col,
          relCX: 0,
          relCY: 0,
          width: w,
          height: h,
          rotationRad: 0,
          dataUrl: exportableUrl,
        });
      };
      img.onerror = (): void => {
        resolve(null);
      };
      img.src = dataUrl;
    });
  }

  /** 将 prepared 插入为新的浮动图并返回完整快照（粘贴命令首次 execute）。 */
  insertFloatingPictureFromPrepared(p: FloatingPicturePastePrepared): FloatingPictureSnapshot {
    const id = `fp-${++this.idSeq}`;
    const z = ++this.zCounter;
    const model: PictureModel = {
      id,
      sheetName: p.sheetName,
      sheetIndex: p.sheetIndex,
      anchorRow: p.anchorRow,
      anchorCol: p.anchorCol,
      relCX: p.relCX,
      relCY: p.relCY,
      width: p.width,
      height: p.height,
      rotationRad: p.rotationRad,
      dataUrl: p.dataUrl,
      z,
      adjustments: cloneAdjustments(DEFAULT_FLOATING_PICTURE_ADJUSTMENTS),
    };
    const el = this.createItemElement(model);
    this.byId.set(id, { model, el });
    this.root.appendChild(el);
    this.selectById(id);
    this.layout();
    return {
      id: model.id,
      sheetName: model.sheetName,
      sheetIndex: model.sheetIndex,
      anchorRow: model.anchorRow,
      anchorCol: model.anchorCol,
      relCX: model.relCX,
      relCY: model.relCY,
      width: model.width,
      height: model.height,
      rotationRad: model.rotationRad,
      dataUrl: model.dataUrl,
      z: model.z,
      adjustments: cloneAdjustments(model.adjustments),
    };
  }

  /** 供 XLSX 导出：当前浮动图片快照（画布像素 + `viewZoom` 在导出侧换算）。 */
  getPicturesForXlsxExport(): readonly XlsxFloatingPictureExport[] {
    const out: XlsxFloatingPictureExport[] = [];
    for (const { model } of this.byId.values()) {
      out.push({
        sheetName: model.sheetName,
        anchorRow: model.anchorRow,
        anchorCol: model.anchorCol,
        relCX: model.relCX,
        relCY: model.relCY,
        width: model.width,
        height: model.height,
        rotationRad: model.rotationRad,
        dataUrl: model.dataUrl,
      });
    }
    return out;
  }

  deselect(): void {
    if (this.selectedId === null) {
      return;
    }
    const rec = this.byId.get(this.selectedId);
    if (rec !== undefined) {
      rec.el.classList.remove("fs-fp-item--selected");
    }
    this.selectedId = null;
    this.emitFloatingPictureFocus(false);
  }

  /** 取消选中且不通知（用于选中另一张图时避免 ribbon 先拆后装）。 */
  private clearSelectionWithoutNotify(): void {
    if (this.selectedId === null) {
      return;
    }
    const rec = this.byId.get(this.selectedId);
    if (rec !== undefined) {
      rec.el.classList.remove("fs-fp-item--selected");
    }
    this.selectedId = null;
  }

  /** 将图片解析到当前工作簿中的表索引：优先表名；仅当索引处表名仍一致时回退索引（表已删或改名则 -1）。 */
  private resolvePictureSheetIndex(model: PictureModel): number {
    const wb = this.getWorkbook();
    const sheets = wb.getSheets();
    const iName = sheets.findIndex((s) => s.name === model.sheetName);
    if (iName >= 0) {
      return iName;
    }
    const at = wb.getSheet(model.sheetIndex);
    if (at !== undefined && at.name === model.sheetName) {
      return model.sheetIndex;
    }
    return -1;
  }

  private isPictureOnActiveSheet(model: PictureModel): boolean {
    return this.resolvePictureSheetIndex(model) === this.getWorkbook().activeSheetIndex;
  }

  /** 视图缩放变化时按比例缩放几何（保持与网格视觉一致）。 */
  scaleForZoomChange(prevZ: number, nextZ: number): void {
    if (prevZ <= 0 || !Number.isFinite(prevZ) || !Number.isFinite(nextZ)) {
      return;
    }
    const k = nextZ / prevZ;
    for (const { model } of this.byId.values()) {
      model.relCX *= k;
      model.relCY *= k;
      model.width *= k;
      model.height *= k;
    }
  }

  layout(): void {
    this.syncClipToTableBody();
    const orphanIds: string[] = [];
    for (const [id, rec] of this.byId) {
      const idx = this.resolvePictureSheetIndex(rec.model);
      if (idx < 0) {
        orphanIds.push(id);
        continue;
      }
      rec.model.sheetIndex = idx;
    }
    for (const oid of orphanIds) {
      const rec = this.byId.get(oid);
      if (rec !== undefined) {
        if (this.selectedId === oid) {
          this.deselect();
        }
        rec.el.remove();
        this.byId.delete(oid);
      }
    }

    for (const [, rec] of this.byId) {
      if (!this.isPictureOnActiveSheet(rec.model)) {
        rec.el.style.display = "none";
        continue;
      }
      rec.el.style.display = "";
      this.applyGeometry(rec.model, rec.el);
    }
    if (this.selectedId !== null) {
      const rec = this.byId.get(this.selectedId);
      if (rec === undefined || !this.isPictureOnActiveSheet(rec.model)) {
        this.deselect();
      }
    }
  }

  addPictureFromDataUrl(dataUrl: string): void {
    const renderer = this.getRenderer();
    const wb = this.getWorkbook();
    const sheet = wb.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    const img = new Image();
    img.onload = (): void => {
      let w = img.naturalWidth || 1;
      let h = img.naturalHeight || 1;
      const scale = Math.min(1, INSERT_MAX_DIM / Math.max(w, h));
      w = Math.max(MIN_W, w * scale);
      h = Math.max(MIN_H, h * scale);

      const ac = this.getAnchorCell();
      const sheet0 = wb.getActiveSheet();
      if (sheet0 === undefined) {
        return;
      }
      const cell = {
        row: Math.max(0, Math.min(sheet0.rowCount - 1, ac.row)),
        col: Math.max(0, Math.min(sheet0.colCount - 1, ac.col)),
      };
      const cr = renderer.getCellRectInCanvasPixels(cell.row, cell.col);
      if (cr === null) {
        return;
      }
      // Excel only supports PNG/JPEG in drawings; convert other formats (WebP, GIF, etc.) to PNG
      const exportableUrl = normalizeToExportableDataUrl(img, dataUrl);
      const id = `fp-${++this.idSeq}`;
      const model: PictureModel = {
        id,
        sheetName: sheet0.name,
        sheetIndex: wb.activeSheetIndex,
        anchorRow: cell.row,
        anchorCol: cell.col,
        relCX: 0,
        relCY: 0,
        width: w,
        height: h,
        rotationRad: 0,
        dataUrl: exportableUrl,
        z: ++this.zCounter,
        adjustments: cloneAdjustments(DEFAULT_FLOATING_PICTURE_ADJUSTMENTS),
      };
      const el = this.createItemElement(model);
      this.byId.set(id, { model, el });
      this.root.appendChild(el);
      this.selectById(id);
      this.layout();
    };
    img.onerror = (): void => {
      /* 忽略损坏文件 */
    };
    img.src = dataUrl;
  }

  private createItemElement(model: PictureModel): HTMLDivElement {
    const wrap = document.createElement("div");
    wrap.className = "fs-fp-item";
    wrap.dataset.fpId = model.id;
    wrap.dataset.fsFloatingSheet = model.sheetName;
    const im = document.createElement("img");
    im.className = "fs-fp-item__img";
    im.draggable = false;
    im.alt = "";
    im.src = model.dataUrl;
    const focus = document.createElement("div");
    focus.className = "fs-fp-item__focus";
    const handles: ResizeHandleId[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
    for (const h of handles) {
      const d = document.createElement("div");
      d.className = `fs-fp-handle fs-fp-handle--${h}`;
      d.dataset.handle = h;
      d.addEventListener("pointerdown", (ev) => this.onHandlePointerDown(ev, model.id, h));
      focus.appendChild(d);
    }
    const rotLine = document.createElement("div");
    rotLine.className = "fs-fp-rotate-line";
    focus.appendChild(rotLine);
    const rot = document.createElement("div");
    rot.className = "fs-fp-rotate";
    rot.dataset.handle = "rotate";
    rot.addEventListener("pointerdown", (ev) => this.onRotatePointerDown(ev, model.id));
    focus.appendChild(rot);

    wrap.appendChild(im);
    wrap.appendChild(focus);
    wrap.addEventListener("pointerdown", (ev) => this.onItemPointerDown(ev, model.id));
    this.applyImageFilterToElement(wrap, model);
    return wrap;
  }

  private applyImageFilterToElement(el: HTMLDivElement, model: PictureModel): void {
    const img = el.querySelector(".fs-fp-item__img");
    if (img instanceof HTMLImageElement) {
      img.style.filter = buildFloatingPictureCssFilter(model.adjustments);
    }
  }

  getFloatingPictureAdjustments(pictureId: string): FloatingPictureAdjustments | null {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return null;
    }
    return cloneAdjustments(rec.model.adjustments);
  }

  setFloatingPictureAdjustments(
    pictureId: string,
    patch: Partial<FloatingPictureAdjustments>,
  ): void {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return;
    }
    rec.model.adjustments = { ...rec.model.adjustments, ...patch };
    this.applyImageFilterToElement(rec.el, rec.model);
  }

  resetFloatingPictureAdjustments(pictureId: string): void {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return;
    }
    rec.model.adjustments = cloneAdjustments(DEFAULT_FLOATING_PICTURE_ADJUSTMENTS);
    this.applyImageFilterToElement(rec.el, rec.model);
  }

  /** 画布像素下的尺寸与偏移（格式窗格「裁剪」区只读展示）。 */
  getFloatingPictureLayoutPx(pictureId: string): {
    readonly widthPx: number;
    readonly heightPx: number;
    readonly offsetXPx: number;
    readonly offsetYPx: number;
  } | null {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return null;
    }
    const m = rec.model;
    return { widthPx: m.width, heightPx: m.height, offsetXPx: m.relCX, offsetYPx: m.relCY };
  }

  private getCenterCanvas(model: PictureModel): { cx: number; cy: number } | null {
    const renderer = this.getRenderer();
    const wb = this.getWorkbook();
    if (!this.isPictureOnActiveSheet(model)) {
      return null;
    }
    const sheet = wb.getActiveSheet();
    if (sheet === undefined) {
      return null;
    }
    const cr = renderer.getCellRectInCanvasPixels(model.anchorRow, model.anchorCol);
    if (cr === null) {
      return null;
    }
    return {
      cx: cr.x + cr.width / 2 + model.relCX,
      cy: cr.y + cr.height / 2 + model.relCY,
    };
  }

  private applyGeometry(
    model: PictureModel,
    el: HTMLDivElement,
    options?: { readonly snapPixels?: boolean },
  ): void {
    const c = this.getCenterCanvas(model);
    if (c === null) {
      return;
    }
    const left = c.cx - model.width / 2;
    const top = c.cy - model.height / 2;
    const snap = options?.snapPixels ?? true;
    if (snap) {
      el.style.left = `${Math.round(left)}px`;
      el.style.top = `${Math.round(top)}px`;
      el.style.width = `${Math.round(model.width)}px`;
      el.style.height = `${Math.round(model.height)}px`;
    } else {
      el.style.left = `${left}px`;
      el.style.top = `${top}px`;
      el.style.width = `${model.width}px`;
      el.style.height = `${model.height}px`;
    }
    el.style.zIndex = String(model.z);
    el.style.transform = `rotate(${model.rotationRad}rad)`;
    el.style.transformOrigin = "center center";
  }

  private selectById(id: string): void {
    this.clearSelectionWithoutNotify();
    this.selectedId = id;
    const rec = this.byId.get(id);
    if (rec === undefined) {
      this.emitFloatingPictureFocus(false);
      return;
    }
    rec.el.classList.add("fs-fp-item--selected");
    this.zCounter += 1;
    rec.model.z = this.zCounter;
    rec.el.style.zIndex = String(rec.model.z);
    this.emitFloatingPictureFocus(true);
  }

  private onItemPointerDown(ev: PointerEvent, id: string): void {
    if (ev.button !== 0) {
      return;
    }
    const t = ev.target as HTMLElement;
    if (t.closest(".fs-fp-handle, .fs-fp-rotate")) {
      return;
    }
    ev.stopPropagation();
    ev.preventDefault();
    this.selectById(id);
    const rec = this.byId.get(id);
    if (rec === undefined) {
      return;
    }
    const c = this.getCenterCanvas(rec.model);
    if (c === null) {
      return;
    }
    const { x, y } = this.clientToCanvas(ev.clientX, ev.clientY);
    this.drag = {
      kind: "move",
      id,
      startCX: x,
      startCY: y,
      startRelCX: rec.model.relCX,
      startRelCY: rec.model.relCY,
    };
    this.attachDragListeners();
    try {
      (ev.currentTarget as HTMLElement).setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
  }

  private onHandlePointerDown(ev: PointerEvent, id: string, handle: ResizeHandleId): void {
    if (ev.button !== 0) {
      return;
    }
    ev.stopPropagation();
    ev.preventDefault();
    this.selectById(id);
    const rec = this.byId.get(id);
    if (rec === undefined) {
      return;
    }
    const c = this.getCenterCanvas(rec.model);
    if (c === null) {
      return;
    }
    this.drag = {
      kind: "resize",
      id,
      handle,
      startCenterX: c.cx,
      startCenterY: c.cy,
      startW: rec.model.width,
      startH: rec.model.height,
      startRot: rec.model.rotationRad,
    };
    this.attachDragListeners();
    try {
      (ev.currentTarget as HTMLElement).setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
  }

  private onRotatePointerDown(ev: PointerEvent, id: string): void {
    if (ev.button !== 0) {
      return;
    }
    ev.stopPropagation();
    ev.preventDefault();
    this.selectById(id);
    const rec = this.byId.get(id);
    if (rec === undefined) {
      return;
    }
    const c = this.getCenterCanvas(rec.model);
    if (c === null) {
      return;
    }
    const { x, y } = this.clientToCanvas(ev.clientX, ev.clientY);
    const startAngle = Math.atan2(y - c.cy, x - c.cx);
    this.drag = {
      kind: "rotate",
      id,
      startAngle,
      startRot: rec.model.rotationRad,
      cx: c.cx,
      cy: c.cy,
    };
    this.attachDragListeners();
    try {
      (ev.currentTarget as HTMLElement).setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
  }

  private clientToCanvas(clientX: number, clientY: number): { x: number; y: number } {
    const canvas = this.getCanvas();
    const rect = canvas.getBoundingClientRect();
    const cw = canvas.clientWidth;
    const ch = canvas.clientHeight;
    const rw = Math.max(rect.width, 1e-6);
    const rh = Math.max(rect.height, 1e-6);
    return {
      x: ((clientX - rect.left) / rw) * cw,
      y: ((clientY - rect.top) / rh) * ch,
    };
  }

  private attachDragListeners(): void {
    document.addEventListener("pointermove", this.onDocPointerMove);
    document.addEventListener("pointerup", this.onDocPointerUp);
    document.addEventListener("pointercancel", this.onDocPointerUp);
  }

  private detachDragListeners(): void {
    document.removeEventListener("pointermove", this.onDocPointerMove);
    document.removeEventListener("pointerup", this.onDocPointerUp);
    document.removeEventListener("pointercancel", this.onDocPointerUp);
  }

  private readonly onDocPointerMove = (ev: PointerEvent): void => {
    if (this.drag === null) {
      return;
    }
    const rec = this.byId.get(this.drag.id);
    if (rec === undefined) {
      return;
    }
    const m = rec.model;
    const { x, y } = this.clientToCanvas(ev.clientX, ev.clientY);
    if (this.drag.kind === "move") {
      const dx = x - this.drag.startCX;
      const dy = y - this.drag.startCY;
      m.relCX = this.drag.startRelCX + dx;
      m.relCY = this.drag.startRelCY + dy;
      this.applyGeometry(m, rec.el, { snapPixels: false });
      return;
    }
    if (this.drag.kind === "rotate") {
      const ang = Math.atan2(y - this.drag.cy, x - this.drag.cx);
      m.rotationRad = this.drag.startRot + (ang - this.drag.startAngle);
      this.applyGeometry(m, rec.el, { snapPixels: false });
      return;
    }

    /**
     * resize：画布空间几何同步
     * - 边：固定对边中点，宽度/高度 = 指针在边法向上的投影（与鼠标 1:1）
     * - 角：固定对角顶点，沿「宽+高比例」对角射线投影，保持起始宽高比
     */
    const d = this.drag;
    const rot = d.startRot;
    const sx = d.startCenterX;
    const sy = d.startCenterY;
    const startHw = d.startW / 2;
    const startHh = d.startH / 2;
    const r = d.startH / Math.max(1e-9, d.startW);

    const renderer = this.getRenderer();
    const crNow = renderer.getCellRectInCanvasPixels(m.anchorRow, m.anchorCol);
    if (crNow === null) {
      return;
    }
    const cellCx = crNow.x + crNow.width / 2;
    const cellCy = crNow.y + crNow.height / 2;

    const applyCenterSize = (c1x: number, c1y: number, nw: number, nh: number): void => {
      m.width = nw;
      m.height = nh;
      m.relCX = c1x - cellCx;
      m.relCY = c1y - cellCy;
      this.applyGeometry(m, rec.el, { snapPixels: false });
    };

    const nX = rotLocalToWorld(1, 0, rot);
    const nY = rotLocalToWorld(0, 1, rot);

    if (d.handle === "e") {
      const wMid = rotLocalToWorld(-startHw, 0, rot);
      const Wfx = sx + wMid.x;
      const Wfy = sy + wMid.y;
      const wFull = Math.max(MIN_W, dot2(x - Wfx, y - Wfy, nX.x, nX.y));
      const half = wFull / 2;
      applyCenterSize(Wfx + nX.x * half, Wfy + nX.y * half, wFull, d.startH);
      return;
    }
    if (d.handle === "w") {
      const eMid = rotLocalToWorld(startHw, 0, rot);
      const Efx = sx + eMid.x;
      const Efy = sy + eMid.y;
      const wFull = Math.max(MIN_W, dot2(Efx - x, Efy - y, nX.x, nX.y));
      const half = wFull / 2;
      applyCenterSize(Efx - nX.x * half, Efy - nX.y * half, wFull, d.startH);
      return;
    }
    if (d.handle === "s") {
      const nMid = rotLocalToWorld(0, -startHh, rot);
      const Nfx = sx + nMid.x;
      const Nfy = sy + nMid.y;
      const hFull = Math.max(MIN_H, dot2(x - Nfx, y - Nfy, nY.x, nY.y));
      const half = hFull / 2;
      applyCenterSize(Nfx + nY.x * half, Nfy + nY.y * half, d.startW, hFull);
      return;
    }
    if (d.handle === "n") {
      const sMid = rotLocalToWorld(0, startHh, rot);
      const Sfx = sx + sMid.x;
      const Sfy = sy + sMid.y;
      const hFull = Math.max(MIN_H, dot2(Sfx - x, Sfy - y, nY.x, nY.y));
      const half = hFull / 2;
      applyCenterSize(Sfx - nY.x * half, Sfy - nY.y * half, d.startW, hFull);
      return;
    }

    if (d.handle === "se") {
      const cNW = rotLocalToWorld(-startHw, -startHh, rot);
      const Ox = sx + cNW.x;
      const Oy = sy + cNW.y;
      const Dx = nX.x + r * nY.x;
      const Dy = nX.y + r * nY.y;
      const dd = dot2(Dx, Dy, Dx, Dy);
      const wFull = Math.max(MIN_W, dd > 1e-12 ? dot2(x - Ox, y - Oy, Dx, Dy) / dd : MIN_W);
      const hFull = wFull * r;
      const halfW = (wFull * nX.x + hFull * nY.x) / 2;
      const halfH = (wFull * nX.y + hFull * nY.y) / 2;
      applyCenterSize(Ox + halfW, Oy + halfH, wFull, hFull);
      return;
    }
    if (d.handle === "nw") {
      const cSE = rotLocalToWorld(startHw, startHh, rot);
      const Ox = sx + cSE.x;
      const Oy = sy + cSE.y;
      const Dx = nX.x + r * nY.x;
      const Dy = nX.y + r * nY.y;
      const dd = dot2(Dx, Dy, Dx, Dy);
      const wFull = Math.max(MIN_W, dd > 1e-12 ? dot2(Ox - x, Oy - y, Dx, Dy) / dd : MIN_W);
      const hFull = wFull * r;
      const halfW = (wFull * nX.x + hFull * nY.x) / 2;
      const halfH = (wFull * nX.y + hFull * nY.y) / 2;
      applyCenterSize(Ox - halfW, Oy - halfH, wFull, hFull);
      return;
    }
    if (d.handle === "ne") {
      const cSW = rotLocalToWorld(-startHw, startHh, rot);
      const Ox = sx + cSW.x;
      const Oy = sy + cSW.y;
      const Dx = nX.x - r * nY.x;
      const Dy = nX.y - r * nY.y;
      const dd = dot2(Dx, Dy, Dx, Dy);
      const wFull = Math.max(MIN_W, dd > 1e-12 ? dot2(x - Ox, y - Oy, Dx, Dy) / dd : MIN_W);
      const hFull = wFull * r;
      const halfW = (wFull * nX.x - hFull * nY.x) / 2;
      const halfH = (wFull * nX.y - hFull * nY.y) / 2;
      applyCenterSize(Ox + halfW, Oy + halfH, wFull, hFull);
      return;
    }
    if (d.handle === "sw") {
      const cNE = rotLocalToWorld(startHw, -startHh, rot);
      const Ox = sx + cNE.x;
      const Oy = sy + cNE.y;
      const Dx = -nX.x + r * nY.x;
      const Dy = -nX.y + r * nY.y;
      const dd = dot2(Dx, Dy, Dx, Dy);
      const wFull = Math.max(MIN_W, dd > 1e-12 ? dot2(x - Ox, y - Oy, Dx, Dy) / dd : MIN_W);
      const hFull = wFull * r;
      const halfW = (-wFull * nX.x + hFull * nY.x) / 2;
      const halfH = (-wFull * nX.y + hFull * nY.y) / 2;
      applyCenterSize(Ox + halfW, Oy + halfH, wFull, hFull);
    }
  };

  private readonly onDocPointerUp = (): void => {
    this.endDrag();
  };

  private endDrag(): void {
    if (this.drag === null) {
      return;
    }
    const dragId = this.drag.id;
    const kind = this.drag.kind;
    this.detachDragListeners();
    this.drag = null;

    const rec = this.byId.get(dragId);
    if (rec === undefined) {
      return;
    }
    if (kind === "move" || kind === "resize" || kind === "rotate") {
      rec.model.width = Math.max(MIN_W, Math.round(rec.model.width));
      rec.model.height = Math.max(MIN_H, Math.round(rec.model.height));
      rec.model.relCX = Math.round(rec.model.relCX);
      rec.model.relCY = Math.round(rec.model.relCY);
      this.applyGeometry(rec.model, rec.el, { snapPixels: true });
    }
  }
}

function normalizeToExportableDataUrl(img: HTMLImageElement, dataUrl: string): string {
  const m = /^data:([^;,]+)/i.exec(dataUrl);
  const mime = m?.[1]?.toLowerCase() ?? "";
  if (mime === "image/png" || mime === "image/jpeg" || mime === "image/jpg") {
    return dataUrl;
  }
  const canvas = document.createElement("canvas");
  canvas.width = img.naturalWidth || 1;
  canvas.height = img.naturalHeight || 1;
  const ctx = canvas.getContext("2d");
  if (ctx === null) {
    return dataUrl;
  }
  ctx.drawImage(img, 0, 0);
  return canvas.toDataURL("image/png");
}

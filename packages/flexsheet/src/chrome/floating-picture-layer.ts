/**
 * 工作表上的浮动图片层（DOM 覆盖在 Canvas 之上，非单元格内嵌）。
 * 插入后可选中，通过把手缩放、旋转与拖动。
 */

import type { Workbook } from "@flexsheet/core";
import {
  floatingPictureNeedsRasterForXlsxExport,
  type XlsxFloatingPictureExport,
  type XlsxImportedFloatingPicture,
} from "@flexsheet/import-export";
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
.fs-fp-root.fs-fp-root--crop-active{z-index:120;}
.fs-fp-item{position:absolute;pointer-events:auto;box-sizing:border-box;}
.fs-fp-item__body{position:relative;width:100%;height:100%;overflow:hidden;box-sizing:border-box;}
.fs-fp-item__img-wrap{position:absolute;left:0;top:0;box-sizing:border-box;}
.fs-fp-item__img{display:block;width:100%;height:100%;object-fit:fill;user-select:none;-webkit-user-drag:none;}
.fs-fp-item__focus{position:absolute;inset:0;pointer-events:none;border:1px solid #333;box-sizing:border-box;}
.fs-fp-item--selected .fs-fp-item__focus{display:block;}
.fs-fp-item:not(.fs-fp-item--selected) .fs-fp-item__focus{display:none;}
.fs-fp-item--cropping.fs-fp-item--selected .fs-fp-item__focus{visibility:hidden;}
.fs-fp-item--cropping.fs-fp-item--selected .fs-fp-item__focus .fs-fp-handle,
.fs-fp-item--cropping.fs-fp-item--selected .fs-fp-item__focus .fs-fp-rotate,
.fs-fp-item--cropping.fs-fp-item--selected .fs-fp-item__focus .fs-fp-rotate-line{display:none !important;}
.fs-fp-item:focus{outline:none;}
.fs-fp-crop-overlay{position:absolute;inset:0;pointer-events:none;z-index:10;}
.fs-fp-crop-frame-handles{position:absolute;inset:0;pointer-events:none;z-index:2;}
.fs-fp-crop-img-handles{position:absolute;pointer-events:none;z-index:1;}
.fs-fp-crop-h{position:absolute;box-sizing:border-box;pointer-events:auto;}
.fs-fp-crop-h--frame.fs-fp-crop-h--nw{left:-3px;top:-3px;width:14px;height:14px;border-left:3px solid #111;border-top:3px solid #111;cursor:nwse-resize;}
.fs-fp-crop-h--frame.fs-fp-crop-h--n{left:50%;top:-4px;width:28px;height:8px;margin-left:-14px;border-top:3px solid #111;cursor:ns-resize;}
.fs-fp-crop-h--frame.fs-fp-crop-h--ne{right:-3px;top:-3px;width:14px;height:14px;border-right:3px solid #111;border-top:3px solid #111;cursor:nesw-resize;}
.fs-fp-crop-h--frame.fs-fp-crop-h--e{right:-4px;top:50%;width:8px;height:28px;margin-top:-14px;border-right:3px solid #111;cursor:ew-resize;}
.fs-fp-crop-h--frame.fs-fp-crop-h--se{right:-3px;bottom:-3px;width:14px;height:14px;border-right:3px solid #111;border-bottom:3px solid #111;cursor:nwse-resize;}
.fs-fp-crop-h--frame.fs-fp-crop-h--s{left:50%;bottom:-4px;width:28px;height:8px;margin-left:-14px;border-bottom:3px solid #111;cursor:ns-resize;}
.fs-fp-crop-h--frame.fs-fp-crop-h--sw{left:-3px;bottom:-3px;width:14px;height:14px;border-left:3px solid #111;border-bottom:3px solid #111;cursor:nesw-resize;}
.fs-fp-crop-h--frame.fs-fp-crop-h--w{left:-4px;top:50%;width:8px;height:28px;margin-top:-14px;border-left:3px solid #111;cursor:ew-resize;}
.fs-fp-crop-h--img{background:#fff;border:1px solid #333;}
.fs-fp-crop-h--img.fs-fp-crop-h--nw{left:-4px;top:-4px;width:7px;height:7px;margin:0;cursor:nwse-resize;border:none;}
.fs-fp-crop-h--img.fs-fp-crop-h--n{left:50%;top:-4px;width:7px;height:7px;margin-left:-4px;cursor:ns-resize;border:none;}
.fs-fp-crop-h--img.fs-fp-crop-h--ne{right:-4px;top:-4px;width:7px;height:7px;margin:0;cursor:nesw-resize;border:none;}
.fs-fp-crop-h--img.fs-fp-crop-h--e{right:-4px;top:50%;width:7px;height:7px;margin-top:-4px;cursor:ew-resize;border:none;}
.fs-fp-crop-h--img.fs-fp-crop-h--se{right:-4px;bottom:-4px;width:7px;height:7px;margin:0;cursor:nwse-resize;border:none;}
.fs-fp-crop-h--img.fs-fp-crop-h--s{left:50%;bottom:-4px;width:7px;height:7px;margin-left:-4px;cursor:ns-resize;border:none;}
.fs-fp-crop-h--img.fs-fp-crop-h--sw{left:-4px;bottom:-4px;width:7px;height:7px;margin:0;cursor:nesw-resize;border:none;}
.fs-fp-crop-h--img.fs-fp-crop-h--w{left:-4px;top:50%;width:7px;height:7px;margin-top:-4px;cursor:ew-resize;border:none;}
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
  /** 结束裁剪会话时提交 before/after（用于撤销栈）；由 FlexSheet 注入。 */
  readonly onCommitCropSession?: (payload: {
    readonly before: FloatingPictureSnapshot;
    readonly after: FloatingPictureSnapshot;
  }) => void;
  /** Delete / Backspace 删除选中浮动图（由 FlexSheet 入栈撤销）。 */
  readonly onRequestDeleteFloatingPicture?: (pictureId: string) => void;
  /**
   * 裁剪会话期间需抬高整块表体视图（含浮动层）的叠放顺序，避免被宿主上其它控件（如内联编辑器 z-20）压住黑框。
   */
  readonly onCropSessionBoostChange?: (boost: boolean) => void;
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

/** 裁剪框内背景填充类型（与「设置图片格式 → 填充与线条」一致）。 */
export type FloatingPictureFillKind = "none" | "solid" | "gradient" | "picture" | "pattern";

/** 浮动图占位矩形（`fs-fp-item__body`）的填充；当前实现纯色，其它 kind 仅占位无绘制。 */
export interface FloatingPictureFrameFill {
  readonly kind: FloatingPictureFillKind;
  /** `solid` 时使用，CSS `#rrggbb` */
  readonly solidColor: string;
  /** 纯色填充透明度 0～100（0 为完全不透明） */
  readonly solidTransparencyPct: number;
}

export const DEFAULT_FLOATING_PICTURE_FRAME_FILL: FloatingPictureFrameFill = {
  kind: "none",
  solidColor: "#000000",
  solidTransparencyPct: 0,
};

function cloneFrameFill(f: FloatingPictureFrameFill): FloatingPictureFrameFill {
  return { kind: f.kind, solidColor: f.solidColor, solidTransparencyPct: f.solidTransparencyPct };
}

function frameFillFromSnapshot(s: FloatingPictureSnapshot): FloatingPictureFrameFill {
  return s.frameFill !== undefined
    ? cloneFrameFill(s.frameFill)
    : cloneFrameFill(DEFAULT_FLOATING_PICTURE_FRAME_FILL);
}

function normalizeSolidColorHex(input: string): string {
  const t = input.trim();
  if (/^#[0-9a-fA-F]{6}$/.test(t)) {
    return t.toLowerCase();
  }
  if (/^#[0-9a-fA-F]{3}$/.test(t)) {
    const r = t[1]!;
    const g = t[2]!;
    const b = t[3]!;
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }
  if (/^[0-9a-fA-F]{6}$/.test(t)) {
    return `#${t.toLowerCase()}`;
  }
  return "#000000";
}

function parseHexRgb(
  hex: string,
): { readonly r: number; readonly g: number; readonly b: number } | null {
  const h = normalizeSolidColorHex(hex);
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/.exec(h);
  if (m === null) {
    return null;
  }
  return {
    r: Number.parseInt(m[1]!, 16),
    g: Number.parseInt(m[2]!, 16),
    b: Number.parseInt(m[3]!, 16),
  };
}

/** 将填充模型转为 `background-color`（非 solid 为透明）。 */
export function frameFillToCssBackground(f: FloatingPictureFrameFill): string {
  if (f.kind !== "solid") {
    return "transparent";
  }
  const rgb = parseHexRgb(f.solidColor);
  if (rgb === null) {
    return "transparent";
  }
  const a = clampN(1 - f.solidTransparencyPct / 100, 0, 1);
  return `rgba(${rgb.r},${rgb.g},${rgb.b},${a})`;
}

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
  /** 裁剪框内背景填充；旧快照省略则视为无填充 */
  readonly frameFill?: FloatingPictureFrameFill;
  /** 源图像素（解码后；旧快照省略则为 0） */
  readonly naturalWidth?: number;
  readonly naturalHeight?: number;
  /** 相对裁剪框（frame）左上角的图片内容矩形（画布像素） */
  readonly imgBoxX?: number;
  readonly imgBoxY?: number;
  readonly imgBoxW?: number;
  readonly imgBoxH?: number;
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
  readonly naturalWidth?: number;
  readonly naturalHeight?: number;
  readonly imgBoxX?: number;
  readonly imgBoxY?: number;
  readonly imgBoxW?: number;
  readonly imgBoxH?: number;
  readonly frameFill?: FloatingPictureFrameFill;
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
  naturalWidth: number;
  naturalHeight: number;
  /** 图片内容在裁剪框（width×height）内的局部矩形 */
  imgBoxX: number;
  imgBoxY: number;
  imgBoxW: number;
  imgBoxH: number;
  frameFill: FloatingPictureFrameFill;
}

type ResizeHandleId = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

const CROP_HANDLE_IDS: readonly ResizeHandleId[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

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
  | { kind: "rotate"; startAngle: number; startRot: number; cx: number; cy: number }
  | {
      kind: "cropImgMove";
      startCX: number;
      startCY: number;
      startImgBoxX: number;
      startImgBoxY: number;
    }
  | {
      kind: "cropImgResize";
      handle: ResizeHandleId;
      startX: number;
      startY: number;
      startW: number;
      startH: number;
    };

function rotLocalToWorld(
  lx: number,
  ly: number,
  rot: number,
): { readonly x: number; readonly y: number } {
  const c = Math.cos(rot);
  const s = Math.sin(rot);
  return { x: lx * c - ly * s, y: lx * s + ly * c };
}

function canvasDeltaToLocal(dcx: number, dcy: number, rot: number): { dlx: number; dly: number } {
  const c = Math.cos(-rot);
  const s = Math.sin(-rot);
  return { dlx: dcx * c - dcy * s, dly: dcx * s + dcy * c };
}

function normalizeImgBoxInModel(m: PictureModel): void {
  m.imgBoxW = Math.max(MIN_W, m.imgBoxW);
  m.imgBoxH = Math.max(MIN_H, m.imgBoxH);
}

export class FloatingPictureLayer {
  private readonly mount: HTMLElement;
  private readonly getCanvas: () => HTMLCanvasElement;
  private readonly getRenderer: () => CanvasRenderer;
  private readonly getWorkbook: () => Workbook;
  private readonly getAnchorCell: () => { readonly row: number; readonly col: number };
  private readonly onCommitCropSession?: FloatingPictureLayerOptions["onCommitCropSession"];
  private readonly onRequestDeleteFloatingPicture?: FloatingPictureLayerOptions["onRequestDeleteFloatingPicture"];
  private readonly onCropSessionBoostChange?: FloatingPictureLayerOptions["onCropSessionBoostChange"];
  private readonly root: HTMLDivElement;
  private readonly byId = new Map<string, { model: PictureModel; el: HTMLDivElement }>();
  private readonly floatingPictureFocusListeners = new Set<(active: boolean) => void>();
  private zCounter = 10;
  private selectedId: string | null = null;
  private drag: (DragMode & { id: string }) | null = null;
  private idSeq = 0;
  /** 交互裁剪会话：右键「裁剪」进入，失焦 / Esc / Enter 结束并提交。 */
  private croppingPictureId: string | null = null;
  private cropEnterSnapshot: FloatingPictureSnapshot | null = null;
  private cropKeyHandler: ((ev: KeyboardEvent) => void) | null = null;
  private cropPointerHandler: ((ev: PointerEvent) => void) | null = null;

  constructor(options: FloatingPictureLayerOptions) {
    ensureStyles();
    this.mount = options.mount;
    this.getCanvas = options.getCanvas;
    this.getRenderer = options.getRenderer;
    this.getWorkbook = options.getWorkbook;
    this.getAnchorCell = options.getAnchorCell;
    this.onCommitCropSession = options.onCommitCropSession;
    this.onRequestDeleteFloatingPicture = options.onRequestDeleteFloatingPicture;
    this.onCropSessionBoostChange = options.onCropSessionBoostChange;
    this.root = document.createElement("div");
    this.root.className = "fs-fp-root";
    this.mount.appendChild(this.root);
    this.syncClipToTableBody();
  }

  /**
   * 裁剪到表体区域（与 Canvas 行列标题带一致），避免拖动时图片盖住行号/列标。
   */
  syncClipToTableBody(): void {
    /** 裁剪时把手伸出图片外，根层 clip-path 会裁切黑框；会话期间临时关闭裁剪。 */
    if (this.croppingPictureId !== null) {
      this.root.style.clipPath = "none";
      this.root.style.removeProperty("-webkit-clip-path");
      return;
    }
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

  /** 裁剪会话叠放：抬高浮动根与宿主表体布局，避免黑框被内联编辑器等盖住。 */
  private syncCropSessionStacking(): void {
    const active = this.croppingPictureId !== null;
    this.root.classList.toggle("fs-fp-root--crop-active", active);
    try {
      this.onCropSessionBoostChange?.(active);
    } catch {
      /* 宿主回调错误不阻断裁剪 */
    }
  }

  destroy(): void {
    this.endDrag();
    this.detachCropOutsideListeners();
    this.croppingPictureId = null;
    this.cropEnterSnapshot = null;
    this.syncCropSessionStacking();
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
    this.detachCropOutsideListeners();
    this.croppingPictureId = null;
    this.cropEnterSnapshot = null;
    const hadFocus = this.selectedId !== null;
    this.selectedId = null;
    for (const { el } of this.byId.values()) {
      el.remove();
    }
    this.byId.clear();
    this.syncCropSessionStacking();
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
    if (this.croppingPictureId === id) {
      this.detachCropOutsideListeners();
      this.croppingPictureId = null;
      this.cropEnterSnapshot = null;
      rec.el.classList.remove("fs-fp-item--cropping");
      this.setCropOverlayVisible(rec.el, false);
      this.syncClipToTableBody();
      this.syncCropSessionStacking();
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
      naturalWidth: m.naturalWidth,
      naturalHeight: m.naturalHeight,
      imgBoxX: m.imgBoxX,
      imgBoxY: m.imgBoxY,
      imgBoxW: m.imgBoxW,
      imgBoxH: m.imgBoxH,
      frameFill: cloneFrameFill(m.frameFill),
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
      naturalWidth: snapshot.naturalWidth ?? 0,
      naturalHeight: snapshot.naturalHeight ?? 0,
      imgBoxX: snapshot.imgBoxX ?? 0,
      imgBoxY: snapshot.imgBoxY ?? 0,
      imgBoxW: snapshot.imgBoxW ?? snapshot.width,
      imgBoxH: snapshot.imgBoxH ?? snapshot.height,
      frameFill: frameFillFromSnapshot(snapshot),
    };
    normalizeImgBoxInModel(model);
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
          naturalWidth: img.naturalWidth || 0,
          naturalHeight: img.naturalHeight || 0,
          imgBoxX: 0,
          imgBoxY: 0,
          imgBoxW: w,
          imgBoxH: h,
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
      naturalWidth: p.naturalWidth ?? 0,
      naturalHeight: p.naturalHeight ?? 0,
      imgBoxX: p.imgBoxX ?? 0,
      imgBoxY: p.imgBoxY ?? 0,
      imgBoxW: p.imgBoxW ?? p.width,
      imgBoxH: p.imgBoxH ?? p.height,
      frameFill:
        p.frameFill !== undefined
          ? cloneFrameFill(p.frameFill)
          : cloneFrameFill(DEFAULT_FLOATING_PICTURE_FRAME_FILL),
    };
    normalizeImgBoxInModel(model);
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
      naturalWidth: model.naturalWidth,
      naturalHeight: model.naturalHeight,
      imgBoxX: model.imgBoxX,
      imgBoxY: model.imgBoxY,
      imgBoxW: model.imgBoxW,
      imgBoxH: model.imgBoxH,
      frameFill: cloneFrameFill(model.frameFill),
    };
  }

  /**
   * XLSX 导入后批量恢复浮动图（`import-xlsx-drawing` 使用工作表逻辑 px，在此乘 `viewZoom` 与导出侧一致）。
   */
  applyImportedFloatingPicturesFromXlsx(
    pictures: readonly XlsxImportedFloatingPicture[],
    viewZoom: number,
  ): void {
    const z = Math.max(1e-6, viewZoom);
    for (const pic of pictures) {
      this.insertFloatingPictureFromPrepared({
        sheetName: pic.sheetName,
        sheetIndex: pic.sheetIndex,
        anchorRow: pic.anchorRow,
        anchorCol: pic.anchorCol,
        relCX: pic.relCX * z,
        relCY: pic.relCY * z,
        width: pic.width * z,
        height: pic.height * z,
        rotationRad: pic.rotationRad,
        dataUrl: pic.dataUrl,
        naturalWidth: pic.naturalWidth,
        naturalHeight: pic.naturalHeight,
        imgBoxX: pic.imgBoxX * z,
        imgBoxY: pic.imgBoxY * z,
        imgBoxW: pic.imgBoxW * z,
        imgBoxH: pic.imgBoxH * z,
        frameFill: pic.frameFill,
      });
    }
    this.deselect();
  }

  private pictureModelToXlsxExportDto(model: PictureModel): XlsxFloatingPictureExport {
    return {
      sheetName: model.sheetName,
      anchorRow: model.anchorRow,
      anchorCol: model.anchorCol,
      relCX: model.relCX,
      relCY: model.relCY,
      width: model.width,
      height: model.height,
      rotationRad: model.rotationRad,
      dataUrl: model.dataUrl,
      naturalWidth: model.naturalWidth,
      naturalHeight: model.naturalHeight,
      imgBoxX: model.imgBoxX,
      imgBoxY: model.imgBoxY,
      imgBoxW: model.imgBoxW,
      imgBoxH: model.imgBoxH,
      frameFill: cloneFrameFill(model.frameFill),
    };
  }

  /**
   * 将单张浮动图栅格化为与裁剪框同尺寸的 PNG（透明留白 + `imgBox` 内绘图），供 OOXML 占位与格线后内容对齐。
   */
  private async rasterizePictureModelToFramePngDataUrl(
    model: PictureModel,
  ): Promise<string | null> {
    if (typeof document === "undefined") {
      return null;
    }
    const W = Math.max(1, Math.round(model.width));
    const H = Math.max(1, Math.round(model.height));
    let nw = model.naturalWidth;
    let nh = model.naturalHeight;
    const canvas = document.createElement("canvas");
    canvas.width = W;
    canvas.height = H;
    const ctx = canvas.getContext("2d");
    if (ctx === null) {
      return null;
    }
    const img = new Image();
    img.decoding = "async";
    try {
      await new Promise<void>((resolve, reject) => {
        img.onload = (): void => resolve();
        img.onerror = (): void => reject(new Error("floating picture load failed"));
        img.src = model.dataUrl;
      });
    } catch {
      return null;
    }
    if (nw <= 0 || nh <= 0) {
      nw = img.naturalWidth;
      nh = img.naturalHeight;
    }
    if (nw <= 0 || nh <= 0) {
      return null;
    }
    const f = buildFloatingPictureCssFilter(model.adjustments);
    ctx.clearRect(0, 0, W, H);
    if (model.frameFill.kind === "solid") {
      ctx.fillStyle = frameFillToCssBackground(model.frameFill);
      ctx.fillRect(0, 0, W, H);
    }
    ctx.filter = f;
    try {
      ctx.drawImage(img, 0, 0, nw, nh, model.imgBoxX, model.imgBoxY, model.imgBoxW, model.imgBoxH);
    } catch {
      ctx.filter = "none";
      ctx.clearRect(0, 0, W, H);
      try {
        ctx.drawImage(
          img,
          0,
          0,
          nw,
          nh,
          model.imgBoxX,
          model.imgBoxY,
          model.imgBoxW,
          model.imgBoxH,
        );
      } catch {
        return null;
      }
    }
    ctx.filter = "none";
    try {
      return canvas.toDataURL("image/png");
    } catch {
      return null;
    }
  }

  /** 供 XLSX 导出：当前浮动图片快照（画布像素 + `viewZoom` 在导出侧换算）。 */
  getPicturesForXlsxExport(): readonly XlsxFloatingPictureExport[] {
    const out: XlsxFloatingPictureExport[] = [];
    for (const { model } of this.byId.values()) {
      out.push(this.pictureModelToXlsxExportDto(model));
    }
    return out;
  }

  /**
   * 供 XLSX 导出：在裁剪框大于图片留白时，先合成与裁剪框同尺寸的 PNG，保证 Excel 中占位与格线/后续内容一致。
   */
  async preparePicturesForXlsxExport(): Promise<readonly XlsxFloatingPictureExport[]> {
    const out: XlsxFloatingPictureExport[] = [];
    for (const { model } of this.byId.values()) {
      const dto = this.pictureModelToXlsxExportDto(model);
      if (floatingPictureNeedsRasterForXlsxExport(dto)) {
        const url = await this.rasterizePictureModelToFramePngDataUrl(model);
        if (url !== null) {
          const rw = Math.max(1, Math.round(model.width));
          const rh = Math.max(1, Math.round(model.height));
          out.push({
            ...dto,
            dataUrl: url,
            naturalWidth: rw,
            naturalHeight: rh,
            imgBoxX: 0,
            imgBoxY: 0,
            imgBoxW: rw,
            imgBoxH: rh,
            frameFill: undefined,
          });
          continue;
        }
      }
      out.push(dto);
    }
    return out;
  }

  deselect(): void {
    if (this.croppingPictureId !== null) {
      this.finishCropSessionAndNotify();
    }
    if (this.selectedId === null) {
      return;
    }
    const rec = this.byId.get(this.selectedId);
    if (rec !== undefined) {
      rec.el.classList.remove("fs-fp-item--selected");
      if (document.activeElement instanceof Node && rec.el.contains(document.activeElement)) {
        rec.el.blur();
      }
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
      if (document.activeElement instanceof Node && rec.el.contains(document.activeElement)) {
        rec.el.blur();
      }
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
      model.imgBoxX *= k;
      model.imgBoxY *= k;
      model.imgBoxW *= k;
      model.imgBoxH *= k;
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
        naturalWidth: img.naturalWidth || 0,
        naturalHeight: img.naturalHeight || 0,
        imgBoxX: 0,
        imgBoxY: 0,
        imgBoxW: w,
        imgBoxH: h,
        frameFill: cloneFrameFill(DEFAULT_FLOATING_PICTURE_FRAME_FILL),
      };
      normalizeImgBoxInModel(model);
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

  private applyFrameFillToItem(el: HTMLDivElement, fill: FloatingPictureFrameFill): void {
    const body = el.querySelector(".fs-fp-item__body");
    if (body instanceof HTMLElement) {
      body.style.backgroundColor = frameFillToCssBackground(fill);
    }
  }

  private createItemElement(model: PictureModel): HTMLDivElement {
    const wrap = document.createElement("div");
    wrap.className = "fs-fp-item";
    wrap.tabIndex = -1;
    wrap.dataset.fpId = model.id;
    wrap.dataset.fsFloatingSheet = model.sheetName;
    wrap.addEventListener("keydown", (ev: KeyboardEvent) => {
      if (ev.isComposing || ev.ctrlKey || ev.metaKey || ev.altKey) {
        return;
      }
      if (ev.key !== "Delete" && ev.key !== "Backspace") {
        return;
      }
      ev.preventDefault();
      ev.stopPropagation();
      this.onRequestDeleteFloatingPicture?.(model.id);
    });
    const body = document.createElement("div");
    body.className = "fs-fp-item__body";
    const imgWrap = document.createElement("div");
    imgWrap.className = "fs-fp-item__img-wrap";
    const im = document.createElement("img");
    im.className = "fs-fp-item__img";
    im.draggable = false;
    im.alt = "";
    im.src = model.dataUrl;
    im.addEventListener("load", () => {
      const r = this.byId.get(model.id);
      if (r === undefined) {
        return;
      }
      r.model.naturalWidth = im.naturalWidth || 0;
      r.model.naturalHeight = im.naturalHeight || 0;
    });
    imgWrap.appendChild(im);
    body.appendChild(imgWrap);
    const focus = document.createElement("div");
    focus.className = "fs-fp-item__focus";
    for (const h of CROP_HANDLE_IDS) {
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

    wrap.appendChild(body);
    wrap.appendChild(focus);
    wrap.addEventListener("pointerdown", (ev) => this.onItemPointerDown(ev, model.id));
    this.applyImageFilterToElement(wrap, model);
    this.applyFrameFillToItem(wrap, model.frameFill);
    this.ensureCropOverlayMounted(wrap);
    this.syncInnerLayout(model, wrap);
    return wrap;
  }

  private applyImageFilterToElement(el: HTMLDivElement, model: PictureModel): void {
    const img = el.querySelector(".fs-fp-item__img");
    if (img instanceof HTMLImageElement) {
      img.style.filter = buildFloatingPictureCssFilter(model.adjustments);
    }
  }

  private syncInnerLayout(model: PictureModel, el: HTMLDivElement): void {
    const wrap = el.querySelector(".fs-fp-item__img-wrap");
    if (wrap instanceof HTMLElement) {
      wrap.style.left = `${model.imgBoxX}px`;
      wrap.style.top = `${model.imgBoxY}px`;
      wrap.style.width = `${model.imgBoxW}px`;
      wrap.style.height = `${model.imgBoxH}px`;
    }
    const imgHandles = el.querySelector(".fs-fp-crop-img-handles");
    if (imgHandles instanceof HTMLElement) {
      imgHandles.style.left = `${model.imgBoxX}px`;
      imgHandles.style.top = `${model.imgBoxY}px`;
      imgHandles.style.width = `${model.imgBoxW}px`;
      imgHandles.style.height = `${model.imgBoxH}px`;
    }
  }

  private ensureCropOverlayMounted(wrap: HTMLDivElement): void {
    if (wrap.querySelector(":scope > .fs-fp-crop-overlay") !== null) {
      return;
    }
    const overlay = document.createElement("div");
    overlay.className = "fs-fp-crop-overlay";
    overlay.style.display = "none";
    const fh = document.createElement("div");
    fh.className = "fs-fp-crop-frame-handles";
    const ih = document.createElement("div");
    ih.className = "fs-fp-crop-img-handles";
    for (const h of CROP_HANDLE_IDS) {
      const d = document.createElement("div");
      d.className = `fs-fp-crop-h fs-fp-crop-h--frame fs-fp-crop-h--${h}`;
      d.dataset.cropFrameHandle = h;
      d.addEventListener("pointerdown", (ev) => this.onCropFrameHandlePointerDown(ev, wrap, h));
      fh.appendChild(d);
      const di = document.createElement("div");
      di.className = `fs-fp-crop-h fs-fp-crop-h--img fs-fp-crop-h--${h}`;
      di.dataset.cropImgHandle = h;
      di.addEventListener("pointerdown", (ev) => this.onCropImgHandlePointerDown(ev, wrap, h));
      ih.appendChild(di);
    }
    /* 图片白把手在后、裁剪黑框在前，避免角部重叠时白块盖住 L 形黑框 */
    overlay.appendChild(ih);
    overlay.appendChild(fh);
    wrap.appendChild(overlay);
  }

  private setCropOverlayVisible(wrap: HTMLDivElement, vis: boolean): void {
    const ov = wrap.querySelector(":scope > .fs-fp-crop-overlay");
    if (ov instanceof HTMLElement) {
      ov.style.display = vis ? "" : "none";
    }
  }

  private attachCropOutsideListeners(): void {
    if (this.cropKeyHandler !== null) {
      return;
    }
    this.cropKeyHandler = (ev: KeyboardEvent): void => {
      if (this.croppingPictureId === null) {
        return;
      }
      if (ev.key === "Escape" || ev.key === "Enter") {
        ev.preventDefault();
        ev.stopPropagation();
        this.finishCropSessionAndNotify();
      }
    };
    this.cropPointerHandler = (ev: PointerEvent): void => {
      if (this.croppingPictureId === null) {
        return;
      }
      const rec = this.byId.get(this.croppingPictureId);
      if (rec === undefined) {
        this.finishCropSessionAndNotify();
        return;
      }
      if (rec.el.contains(ev.target as Node)) {
        return;
      }
      this.finishCropSessionAndNotify();
    };
    document.addEventListener("keydown", this.cropKeyHandler, true);
    document.addEventListener("pointerdown", this.cropPointerHandler, true);
  }

  private detachCropOutsideListeners(): void {
    if (this.cropKeyHandler !== null) {
      document.removeEventListener("keydown", this.cropKeyHandler, true);
      this.cropKeyHandler = null;
    }
    if (this.cropPointerHandler !== null) {
      document.removeEventListener("pointerdown", this.cropPointerHandler, true);
      this.cropPointerHandler = null;
    }
  }

  /** 右键「裁剪」进入交互编辑；已处于该图裁剪时重复调用为 no-op。 */
  startCropSession(pictureId: string): boolean {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return false;
    }
    if (this.croppingPictureId === pictureId) {
      this.setCropOverlayVisible(rec.el, true);
      rec.el.classList.add("fs-fp-item--cropping");
      this.syncClipToTableBody();
      this.syncCropSessionStacking();
      this.layout();
      return true;
    }
    if (this.croppingPictureId !== null) {
      this.finishCropSessionAndNotify();
    }
    this.selectById(pictureId);
    const snap = this.takeFloatingPictureSnapshot(pictureId);
    if (snap === null) {
      return false;
    }
    this.croppingPictureId = pictureId;
    this.cropEnterSnapshot = snap;
    rec.el.classList.add("fs-fp-item--cropping");
    this.ensureCropOverlayMounted(rec.el);
    this.setCropOverlayVisible(rec.el, true);
    this.syncInnerLayout(rec.model, rec.el);
    this.attachCropOutsideListeners();
    this.syncClipToTableBody();
    this.syncCropSessionStacking();
    this.layout();
    return true;
  }

  isCropSessionActive(): boolean {
    return this.croppingPictureId !== null;
  }

  private finishCropSessionAndNotify(): void {
    if (this.croppingPictureId === null || this.cropEnterSnapshot === null) {
      return;
    }
    const id = this.croppingPictureId;
    const before = this.cropEnterSnapshot;
    this.detachCropOutsideListeners();
    const rec = this.byId.get(id);
    if (rec !== undefined) {
      rec.el.classList.remove("fs-fp-item--cropping");
      this.setCropOverlayVisible(rec.el, false);
    }
    this.croppingPictureId = null;
    this.cropEnterSnapshot = null;
    const after = this.takeFloatingPictureSnapshot(id);
    if (after !== null) {
      this.onCommitCropSession?.({ before, after });
    }
    this.syncClipToTableBody();
    this.syncCropSessionStacking();
  }

  /** 将快照写回已存在的浮动图（撤销/重做）；id 必须已存在。 */
  applySnapshotToPicture(snapshot: FloatingPictureSnapshot): void {
    const rec = this.byId.get(snapshot.id);
    if (rec === undefined) {
      return;
    }
    const m = rec.model;
    m.sheetName = snapshot.sheetName;
    m.sheetIndex = snapshot.sheetIndex;
    m.anchorRow = snapshot.anchorRow;
    m.anchorCol = snapshot.anchorCol;
    m.relCX = snapshot.relCX;
    m.relCY = snapshot.relCY;
    m.width = snapshot.width;
    m.height = snapshot.height;
    m.rotationRad = snapshot.rotationRad;
    m.dataUrl = snapshot.dataUrl;
    m.z = snapshot.z;
    m.adjustments = cloneAdjustments(snapshot.adjustments);
    m.naturalWidth = snapshot.naturalWidth ?? 0;
    m.naturalHeight = snapshot.naturalHeight ?? 0;
    m.imgBoxX = snapshot.imgBoxX ?? 0;
    m.imgBoxY = snapshot.imgBoxY ?? 0;
    m.imgBoxW = snapshot.imgBoxW ?? snapshot.width;
    m.imgBoxH = snapshot.imgBoxH ?? snapshot.height;
    m.frameFill = frameFillFromSnapshot(snapshot);
    normalizeImgBoxInModel(m);
    rec.el.style.zIndex = String(m.z);
    const img = rec.el.querySelector(".fs-fp-item__img");
    if (img instanceof HTMLImageElement && img.src !== m.dataUrl) {
      img.src = m.dataUrl;
    }
    this.applyImageFilterToElement(rec.el, m);
    this.applyFrameFillToItem(rec.el, m.frameFill);
    this.syncInnerLayout(m, rec.el);
    this.applyGeometry(m, rec.el);
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

  getFloatingPictureFrameFill(pictureId: string): FloatingPictureFrameFill | null {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return null;
    }
    return cloneFrameFill(rec.model.frameFill);
  }

  setFloatingPictureFrameFill(pictureId: string, patch: Partial<FloatingPictureFrameFill>): void {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return;
    }
    const cur = rec.model.frameFill;
    const kind = patch.kind ?? cur.kind;
    const solidColor =
      patch.solidColor !== undefined ? normalizeSolidColorHex(patch.solidColor) : cur.solidColor;
    const solidTransparencyPct =
      patch.solidTransparencyPct !== undefined
        ? clampN(patch.solidTransparencyPct, 0, 100)
        : cur.solidTransparencyPct;
    rec.model.frameFill = { kind, solidColor, solidTransparencyPct };
    this.applyFrameFillToItem(rec.el, rec.model.frameFill);
  }

  resetFloatingPictureFrameFill(pictureId: string): void {
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return;
    }
    rec.model.frameFill = cloneFrameFill(DEFAULT_FLOATING_PICTURE_FRAME_FILL);
    this.applyFrameFillToItem(rec.el, rec.model.frameFill);
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
    this.syncInnerLayout(model, el);
  }

  private selectById(id: string): void {
    if (this.croppingPictureId !== null && this.croppingPictureId !== id) {
      this.finishCropSessionAndNotify();
    }
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
    try {
      rec.el.focus({ preventScroll: true });
    } catch {
      /* 部分环境下 focus 可能不可用 */
    }
  }

  private onItemPointerDown(ev: PointerEvent, id: string): void {
    if (ev.button !== 0) {
      return;
    }
    const t = ev.target as HTMLElement;
    if (t.closest(".fs-fp-handle, .fs-fp-rotate, .fs-fp-crop-h")) {
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
    if (this.croppingPictureId === id && t.closest(".fs-fp-item__img-wrap")) {
      this.drag = {
        kind: "cropImgMove",
        id,
        startCX: x,
        startCY: y,
        startImgBoxX: rec.model.imgBoxX,
        startImgBoxY: rec.model.imgBoxY,
      };
    } else {
      this.drag = {
        kind: "move",
        id,
        startCX: x,
        startCY: y,
        startRelCX: rec.model.relCX,
        startRelCY: rec.model.relCY,
      };
    }
    this.attachDragListeners();
    try {
      (ev.currentTarget as HTMLElement).setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
  }

  private onCropFrameHandlePointerDown(
    ev: PointerEvent,
    wrap: HTMLDivElement,
    handle: ResizeHandleId,
  ): void {
    if (ev.button !== 0) {
      return;
    }
    const pictureId = wrap.dataset.fpId;
    if (pictureId === undefined || pictureId === "") {
      return;
    }
    if (this.croppingPictureId !== pictureId) {
      return;
    }
    ev.stopPropagation();
    ev.preventDefault();
    this.selectById(pictureId);
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return;
    }
    const c = this.getCenterCanvas(rec.model);
    if (c === null) {
      return;
    }
    this.drag = {
      kind: "resize",
      id: pictureId,
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

  private onCropImgHandlePointerDown(
    ev: PointerEvent,
    wrap: HTMLDivElement,
    handle: ResizeHandleId,
  ): void {
    if (ev.button !== 0) {
      return;
    }
    const pictureId = wrap.dataset.fpId;
    if (pictureId === undefined || pictureId === "") {
      return;
    }
    if (this.croppingPictureId !== pictureId) {
      return;
    }
    ev.stopPropagation();
    ev.preventDefault();
    this.selectById(pictureId);
    const rec = this.byId.get(pictureId);
    if (rec === undefined) {
      return;
    }
    const m = rec.model;
    this.drag = {
      kind: "cropImgResize",
      id: pictureId,
      handle,
      startX: m.imgBoxX,
      startY: m.imgBoxY,
      startW: m.imgBoxW,
      startH: m.imgBoxH,
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
    if (this.croppingPictureId === id) {
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
    if (this.croppingPictureId === id) {
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

  /** 指针在画布中的位置换算到浮动图局部坐标（原点为裁剪框左上角，未旋转前）。 */
  private clientToFrameLocal(
    clientX: number,
    clientY: number,
    m: PictureModel,
  ): { lx: number; ly: number } | null {
    const c = this.getCenterCanvas(m);
    if (c === null) {
      return null;
    }
    const { x, y } = this.clientToCanvas(clientX, clientY);
    const dx = x - c.cx;
    const dy = y - c.cy;
    const rot = m.rotationRad;
    const co = Math.cos(-rot);
    const si = Math.sin(-rot);
    const lx = dx * co - dy * si + m.width / 2;
    const ly = dx * si + dy * co + m.height / 2;
    return { lx, ly };
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

    if (this.drag.kind === "cropImgMove") {
      const dx = x - this.drag.startCX;
      const dy = y - this.drag.startCY;
      const { dlx, dly } = canvasDeltaToLocal(dx, dy, m.rotationRad);
      m.imgBoxX = this.drag.startImgBoxX + dlx;
      m.imgBoxY = this.drag.startImgBoxY + dly;
      normalizeImgBoxInModel(m);
      this.syncInnerLayout(m, rec.el);
      return;
    }

    if (this.drag.kind === "cropImgResize") {
      const pl = this.clientToFrameLocal(ev.clientX, ev.clientY, m);
      if (pl === null) {
        return;
      }
      const { lx, ly } = pl;
      const d = this.drag;
      const sx = d.startX;
      const sy = d.startY;
      const sw = d.startW;
      const sh = d.startH;
      let nx = sx;
      let ny = sy;
      let nw = sw;
      let nh = sh;
      switch (d.handle) {
        case "e":
          nw = Math.max(MIN_W, lx - sx);
          break;
        case "w": {
          const right = sx + sw;
          nx = Math.min(lx, right - MIN_W);
          nw = Math.max(MIN_W, right - nx);
          break;
        }
        case "s":
          nh = Math.max(MIN_H, ly - sy);
          break;
        case "n": {
          const bottom = sy + sh;
          ny = Math.min(ly, bottom - MIN_H);
          nh = Math.max(MIN_H, bottom - ny);
          break;
        }
        case "se":
          nw = Math.max(MIN_W, lx - sx);
          nh = Math.max(MIN_H, ly - sy);
          break;
        case "nw": {
          const rx = sx + sw;
          const by = sy + sh;
          nx = Math.min(lx, rx - MIN_W);
          ny = Math.min(ly, by - MIN_H);
          nw = Math.max(MIN_W, rx - nx);
          nh = Math.max(MIN_H, by - ny);
          break;
        }
        case "ne": {
          const bottom = sy + sh;
          nw = Math.max(MIN_W, lx - sx);
          nh = Math.max(MIN_H, bottom - ly);
          nx = sx;
          ny = bottom - nh;
          break;
        }
        case "sw": {
          const right = sx + sw;
          nw = Math.max(MIN_W, right - lx);
          nh = Math.max(MIN_H, ly - sy);
          nx = right - nw;
          ny = sy;
          break;
        }
      }
      m.imgBoxX = nx;
      m.imgBoxY = ny;
      m.imgBoxW = nw;
      m.imgBoxH = nh;
      normalizeImgBoxInModel(m);
      this.syncInnerLayout(m, rec.el);
      return;
    }

    /**
     * resize：画布空间几何同步
     * - 边：固定对边中点，宽度/高度 = 指针在边法向上的投影（与鼠标 1:1）
     * - 角：固定对角顶点，沿「宽+高比例」对角射线投影，保持起始宽高比
     */
    if (this.drag.kind !== "resize") {
      return;
    }
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
    if (kind === "cropImgMove" || kind === "cropImgResize") {
      const m = rec.model;
      m.imgBoxX = Math.round(m.imgBoxX);
      m.imgBoxY = Math.round(m.imgBoxY);
      m.imgBoxW = Math.max(MIN_W, Math.round(m.imgBoxW));
      m.imgBoxH = Math.max(MIN_H, Math.round(m.imgBoxH));
      this.syncInnerLayout(m, rec.el);
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

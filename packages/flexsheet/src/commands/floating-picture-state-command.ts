import type { ICommand } from "@flexsheet/core";

import {
  DEFAULT_FLOATING_PICTURE_FRAME_FILL,
  type FloatingPictureAdjustments,
  type FloatingPictureFrameFill,
  type FloatingPictureSnapshot,
} from "../chrome/floating-picture-layer.js";

/** 由 FlexSheet 实现，避免命令模块反向依赖 flex-sheet（循环引用）。 */
export interface FloatingPictureSnapshotApplier {
  applyFloatingPictureSnapshotInPlace(snapshot: FloatingPictureSnapshot): void;
}

function adjustmentsEqual(a: FloatingPictureAdjustments, b: FloatingPictureAdjustments): boolean {
  return (
    a.brightnessPct === b.brightnessPct &&
    a.contrastPct === b.contrastPct &&
    a.sharpnessPct === b.sharpnessPct &&
    a.saturationPct === b.saturationPct &&
    a.colorTemperatureK === b.colorTemperatureK &&
    a.transparencyPct === b.transparencyPct
  );
}

function snapshotFrameFill(s: FloatingPictureSnapshot): FloatingPictureFrameFill {
  return s.frameFill ?? DEFAULT_FLOATING_PICTURE_FRAME_FILL;
}

function frameFillEqual(a: FloatingPictureFrameFill, b: FloatingPictureFrameFill): boolean {
  return (
    a.kind === b.kind &&
    a.solidColor === b.solidColor &&
    a.solidTransparencyPct === b.solidTransparencyPct
  );
}

function snapshotsFloatingGeometryEqual(a: FloatingPictureSnapshot, b: FloatingPictureSnapshot): boolean {
  return (
    a.id === b.id &&
    a.sheetName === b.sheetName &&
    a.sheetIndex === b.sheetIndex &&
    a.anchorRow === b.anchorRow &&
    a.anchorCol === b.anchorCol &&
    a.relCX === b.relCX &&
    a.relCY === b.relCY &&
    a.width === b.width &&
    a.height === b.height &&
    a.rotationRad === b.rotationRad &&
    a.z === b.z &&
    (a.naturalWidth ?? 0) === (b.naturalWidth ?? 0) &&
    (a.naturalHeight ?? 0) === (b.naturalHeight ?? 0) &&
    (a.imgBoxX ?? 0) === (b.imgBoxX ?? 0) &&
    (a.imgBoxY ?? 0) === (b.imgBoxY ?? 0) &&
    (a.imgBoxW ?? a.width) === (b.imgBoxW ?? b.width) &&
    (a.imgBoxH ?? a.height) === (b.imgBoxH ?? b.height) &&
    adjustmentsEqual(a.adjustments, b.adjustments) &&
    frameFillEqual(snapshotFrameFill(a), snapshotFrameFill(b))
  );
}

/** 将浮动图几何/裁剪盒还原为某快照（同一 id，用于裁剪结束等可撤销操作）。 */
export class ReplaceFloatingPictureStateCommand implements ICommand {
  readonly id = "floatingPicture.replaceState";
  readonly label = "编辑图片裁剪";

  constructor(
    private readonly host: FloatingPictureSnapshotApplier,
    private readonly before: FloatingPictureSnapshot,
    private readonly after: FloatingPictureSnapshot,
  ) {}

  execute(): void {
    this.host.applyFloatingPictureSnapshotInPlace(this.after);
  }

  undo(): void {
    this.host.applyFloatingPictureSnapshotInPlace(this.before);
  }

  /** 无变化时不应入栈（由调用方判断）。 */
  static isNoOp(before: FloatingPictureSnapshot, after: FloatingPictureSnapshot): boolean {
    return snapshotsFloatingGeometryEqual(before, after);
  }
}

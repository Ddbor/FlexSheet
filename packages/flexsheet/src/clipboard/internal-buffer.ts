import type { CellStyle } from "@flexsheet/core";

/** 与系统剪贴板 TSV 同步的应用内载荷，用于在纯文本剪贴板之外恢复样式。 */
export interface InternalClipboardPayload {
  readonly tsv: string;
  /** 与 TSV 矩阵同形 */
  readonly styles: ReadonlyArray<ReadonlyArray<CellStyle | null>>;
}

let last: InternalClipboardPayload | null = null;

/** 应用内浮动图片剪贴板（与单元格 TSV 载荷互斥；不依赖系统图像剪贴板 API）。 */
let floatingPictureDataUrl: string | null = null;

export function setInternalClipboardPayload(payload: InternalClipboardPayload): void {
  floatingPictureDataUrl = null;
  last = payload;
}

export function setFloatingPictureClipboard(dataUrl: string): void {
  last = null;
  floatingPictureDataUrl = dataUrl;
}

export function getFloatingPictureClipboard(): string | null {
  return floatingPictureDataUrl;
}

export function getInternalClipboardPayload(): InternalClipboardPayload | null {
  return last;
}

export function clearInternalClipboardPayload(): void {
  last = null;
  floatingPictureDataUrl = null;
}

/** 仅当剪贴板文本与最近一次复制完全一致时返回样式矩阵，否则返回 null（外部粘贴只改值）。 */
export function matchInternalStyles(tsvFromClipboard: string): ReadonlyArray<ReadonlyArray<CellStyle | null>> | null {
  if (last === null || last.tsv !== tsvFromClipboard) {
    return null;
  }
  return last.styles;
}

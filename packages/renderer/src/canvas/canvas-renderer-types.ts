import type { SelectionPaintSnapshot, SelectionRange, Workbook } from "@flexsheet/core";
import type { SheetTheme } from "@flexsheet/theme";

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
  /** 复制/剪切后的走马灯虚线框范围；缺省不绘制。 */
  getClipboardMarqueeRange?: () => SelectionRange | null;
}

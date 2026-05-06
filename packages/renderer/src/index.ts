export {
  CanvasRenderer,
  type CanvasRendererOptions,
  type HorizontalScrollMetrics,
  type VerticalScrollMetrics,
} from "./canvas/canvas-renderer.js";
export { RendererPlugin, type RendererPluginOptions } from "./renderer-plugin.js";
export {
  bodyColumnAutoFilterTextReservePx,
  COLUMN_HEADER_FILTER_BUTTON_CSS_PX,
  hitTestBodyCellAutoFilterButton,
  hitTestCell,
  hitTestColumnHeaderFilterButton,
  hitTestHeadingPointer,
  scrollToRevealCell,
  type HeadingHit,
} from "./hit-test/grid-hit-test.js";
export {
  bodyPaintExtents,
  buildFrozenLayout,
  clampScroll,
  computeScrollLimits,
  visibleScrollableCellRange,
  type FrozenLayout,
  type ViewportScrollLimits,
  type VisibleScrollRange,
} from "./layout/viewport.js";
export { computeColumnAutoWidth, computeRowAutoHeight } from "./cell-auto-fit.js";
export { expandSelectionRangeForMergePaint, getClampedSelectionSpan } from "./canvas/canvas-renderer-selection-span.js";
export { paintCellFillPatternOverlay } from "./canvas/canvas-cell-fill-pattern.js";
export { SELECTION_OUTLINE_VISUAL_SCALE } from "./canvas/canvas-renderer-constants.js";

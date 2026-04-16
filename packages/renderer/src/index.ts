export {
  CanvasRenderer,
  type CanvasRendererOptions,
  type HorizontalScrollMetrics,
  type VerticalScrollMetrics,
} from "./canvas-renderer.js";
export { RendererPlugin, type RendererPluginOptions } from "./renderer-plugin.js";
export {
  COLUMN_HEADER_FILTER_BUTTON_CSS_PX,
  hitTestBodyCellAutoFilterButton,
  hitTestCell,
  hitTestColumnHeaderFilterButton,
  hitTestHeadingPointer,
  scrollToRevealCell,
  type HeadingHit,
} from "./grid-hit-test.js";
export {
  bodyPaintExtents,
  buildFrozenLayout,
  clampScroll,
  computeScrollLimits,
  visibleScrollableCellRange,
  type FrozenLayout,
  type ViewportScrollLimits,
  type VisibleScrollRange,
} from "./viewport.js";
export { computeColumnAutoWidth, computeRowAutoHeight } from "./cell-auto-fit.js";
export { expandSelectionRangeForMergePaint, getClampedSelectionSpan } from "./canvas-renderer-selection-span.js";
export { paintCellFillPatternOverlay } from "./canvas-cell-fill-pattern.js";
export { SELECTION_OUTLINE_VISUAL_SCALE } from "./canvas-renderer-constants.js";

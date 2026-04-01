export {
  CanvasRenderer,
  type CanvasRendererOptions,
  type HorizontalScrollMetrics,
  type VerticalScrollMetrics,
} from "./canvas-renderer.js";
export { RendererPlugin, type RendererPluginOptions } from "./renderer-plugin.js";
export { hitTestCell, scrollToRevealCell } from "./grid-hit-test.js";
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

/** 视图缩放上下限，与 `CanvasRenderer` 一致。 */
export const VIEW_ZOOM_MIN = 0.25;
export const VIEW_ZOOM_MAX = 4;

/**
 * 选区外框、填充柄相对「逻辑 2px 描边」的视觉比例（略细，接近 Excel）。
 */
export const SELECTION_OUTLINE_VISUAL_SCALE = 2 / 3;

/** zoom=1 时行号列默认宽度（CSS px），与表体同乘 viewZoom。 */
export const HEADER_STRIP_BASE_WIDTH = 40;

/** zoom=1 时列标区域高度（CSS px），与表体同乘 viewZoom。 */
export const HEADER_STRIP_BASE_HEIGHT = 24;

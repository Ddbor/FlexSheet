import {
  PluginBase,
  PLUGIN_SERVICE_KEYS,
  type PluginContext,
  type SelectionPaintSnapshot,
  type SelectionRange,
  type Workbook,
} from "@flexsheet/core";
import type { SheetTheme } from "@flexsheet/theme";
import { CanvasRenderer } from "./canvas/canvas-renderer.js";

export interface RendererPluginOptions {
  readonly container: HTMLElement;
  readonly workbook: Workbook;
  readonly theme: SheetTheme;
  readonly frozenRows: number;
  readonly frozenCols: number;
  readonly getSelectionSnapshot: () => SelectionPaintSnapshot | null;
  /** 复制/剪切后的走马灯范围；可选。 */
  readonly getClipboardMarqueeRange?: () => SelectionRange | null;
}

/**
 * Canvas 主渲染器插件：挂载 canvas、注册 `PLUGIN_SERVICE_KEYS.renderer`。
 */
export class RendererPlugin extends PluginBase {
  readonly name = "flexsheet.renderer";

  private canvas: HTMLCanvasElement | null = null;
  private renderer: CanvasRenderer | null = null;

  constructor(private readonly options: RendererPluginOptions) {
    super();
  }

  getRenderer(): CanvasRenderer {
    if (this.renderer === null) {
      throw new Error("RendererPlugin: renderer not initialized");
    }
    return this.renderer;
  }

  getCanvas(): HTMLCanvasElement {
    if (this.canvas === null) {
      throw new Error("RendererPlugin: canvas not initialized");
    }
    return this.canvas;
  }

  override install(ctx: PluginContext): void {
    const canvas = document.createElement("canvas");
    canvas.style.display = "block";
    canvas.style.width = "100%";
    canvas.style.height = "100%";
    canvas.tabIndex = 0;
    canvas.style.outline = "none";
    canvas.style.userSelect = "none";
    canvas.style.touchAction = "none";
    this.options.container.appendChild(canvas);
    this.canvas = canvas;

    this.renderer = new CanvasRenderer({
      canvas,
      workbook: this.options.workbook,
      theme: this.options.theme,
      getSelectionSnapshot: this.options.getSelectionSnapshot,
      getClipboardMarqueeRange: this.options.getClipboardMarqueeRange,
    });
    this.renderer.setFrozenPanes(this.options.frozenRows, this.options.frozenCols);

    ctx.register(PLUGIN_SERVICE_KEYS.canvas, canvas);
    ctx.register(PLUGIN_SERVICE_KEYS.renderer, this.renderer);
  }

  override destroy(): void {
    if (this.renderer !== null) {
      this.renderer.cancelPendingRedraw();
      this.renderer = null;
    }
    if (this.canvas !== null) {
      this.canvas.remove();
      this.canvas = null;
    }
  }
}

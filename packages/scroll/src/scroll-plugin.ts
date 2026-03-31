import { PluginBase, PLUGIN_SERVICE_KEYS, type PluginContext } from "@flexsheet/core";
import type { CanvasRenderer } from "@flexsheet/renderer";
import type { CellEditor } from "@flexsheet/editor";

/**
 * 视口滚轮滚动插件（与冻结窗格、编辑器布局同步）。
 */
export class ScrollPlugin extends PluginBase {
  readonly name = "flexsheet.scroll";

  private canvas: HTMLCanvasElement | null = null;

  private readonly onWheel = (ev: WheelEvent): void => {
    ev.preventDefault();
    const renderer = this.rendererRef;
    const cellEditor = this.cellEditorRef;
    if (renderer === null) {
      return;
    }
    let dx = ev.deltaX;
    let dy = ev.deltaY;
    if (ev.shiftKey && dx === 0) {
      dx = dy;
      dy = 0;
    }
    renderer.applyScrollDelta(dx, dy);
    cellEditor?.syncLayout();
    renderer.requestRedraw();
  };

  private rendererRef: CanvasRenderer | null = null;
  private cellEditorRef: CellEditor | null = null;

  override install(ctx: PluginContext): void {
    const canvas = ctx.get<HTMLCanvasElement>(PLUGIN_SERVICE_KEYS.canvas);
    const renderer = ctx.get<CanvasRenderer>(PLUGIN_SERVICE_KEYS.renderer);
    const cellEditor = ctx.get<CellEditor>(PLUGIN_SERVICE_KEYS.cellEditor);
    if (canvas === undefined || renderer === undefined) {
      return;
    }
    this.canvas = canvas;
    this.rendererRef = renderer;
    this.cellEditorRef = cellEditor ?? null;
  }

  override deactivate(): void {
    if (this.canvas !== null) {
      this.canvas.removeEventListener("wheel", this.onWheel);
    }
  }

  override activate(): void {
    if (this.canvas !== null) {
      this.canvas.addEventListener("wheel", this.onWheel, { passive: false });
    }
  }

  override destroy(): void {
    if (this.canvas !== null) {
      this.canvas.removeEventListener("wheel", this.onWheel);
      this.canvas = null;
    }
    this.rendererRef = null;
    this.cellEditorRef = null;
  }
}

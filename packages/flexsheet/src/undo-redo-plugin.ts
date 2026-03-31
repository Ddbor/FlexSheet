import { PluginBase, PLUGIN_SERVICE_KEYS, type PluginContext } from "@flexsheet/core";
import type { CellEditor } from "@flexsheet/editor";

export interface UndoRedoPluginOptions {
  readonly canvas: HTMLCanvasElement;
}

/**
 * 在画布上绑定 Ctrl+Z / Ctrl+Y（及 Ctrl+Shift+Z 重做），与 `Workspace.commands` 联动。
 * 需在 `EditorPlugin` 之后注册，以便在单元格内联编辑时跳过快捷键。
 */
export class UndoRedoPlugin extends PluginBase {
  readonly name = "flexsheet.undoRedo";

  private ctx: PluginContext | null = null;
  private readonly canvas: HTMLCanvasElement;

  constructor(options: UndoRedoPluginOptions) {
    super();
    this.canvas = options.canvas;
  }

  override install(ctx: PluginContext): void {
    this.ctx = ctx;
  }

  override activate(): void {
    this.canvas.addEventListener("keydown", this.onKeyDown, true);
  }

  override deactivate(): void {
    this.canvas.removeEventListener("keydown", this.onKeyDown, true);
  }

  override destroy(): void {
    this.ctx = null;
  }

  private readonly onKeyDown = (ev: KeyboardEvent): void => {
    const ctx = this.ctx;
    if (ctx === null) {
      return;
    }
    const editor = ctx.get<CellEditor>(PLUGIN_SERVICE_KEYS.cellEditor);
    if (editor?.isEditing() === true) {
      return;
    }
    const mod = ev.ctrlKey || ev.metaKey;
    if (!mod || ev.altKey) {
      return;
    }
    const key = ev.key.toLowerCase();
    if (key === "z" && !ev.shiftKey) {
      if (ctx.workspace.commands.undo()) {
        ev.preventDefault();
        ev.stopPropagation();
      }
      return;
    }
    if (key === "y" || (key === "z" && ev.shiftKey)) {
      if (ctx.workspace.commands.redo()) {
        ev.preventDefault();
        ev.stopPropagation();
      }
    }
  };
}

/** 工厂函数，与名称 `useUndoRedo` 对齐（接入撤销重做快捷键与命令栈）。 */
export function useUndoRedo(options: UndoRedoPluginOptions): UndoRedoPlugin {
  return new UndoRedoPlugin(options);
}

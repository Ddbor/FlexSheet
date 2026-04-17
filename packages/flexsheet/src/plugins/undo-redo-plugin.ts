import { PluginBase, PLUGIN_SERVICE_KEYS, type PluginContext } from "@flexsheet/core";
import type { CellEditor } from "@flexsheet/editor";

import { isEditableKeydownTarget } from "../chrome/keyboard-editable-target.js";

export interface UndoRedoPluginOptions {
  readonly canvas: HTMLCanvasElement;
  /** 快捷键监听目标；通常为 `chromeRoot`（含 Ribbon、画布、底部栏），否则为画布。 */
  readonly keyTarget?: HTMLElement;
}

/**
 * 在 `keyTarget` 上绑定 Ctrl+Z / Ctrl+Y（及 Ctrl+Shift+Z 重做），与 `Workspace.commands` 联动。
 * 在单元格内联编辑、编辑栏或其它输入控件聚焦时跳过，以保留浏览器撤消/重做。
 */
export class UndoRedoPlugin extends PluginBase {
  readonly name = "flexsheet.undoRedo";

  private ctx: PluginContext | null = null;
  private readonly keyTarget: HTMLElement;

  constructor(options: UndoRedoPluginOptions) {
    super();
    this.keyTarget = options.keyTarget ?? options.canvas;
  }

  override install(ctx: PluginContext): void {
    this.ctx = ctx;
  }

  override activate(): void {
    this.keyTarget.addEventListener("keydown", this.onKeyDown, true);
  }

  override deactivate(): void {
    this.keyTarget.removeEventListener("keydown", this.onKeyDown, true);
  }

  override destroy(): void {
    this.ctx = null;
  }

  private readonly onKeyDown = (ev: KeyboardEvent): void => {
    if (isEditableKeydownTarget(ev)) {
      return;
    }
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
